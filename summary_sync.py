"""
summary_sync.py
================

Безопасная синхронизация двух Google Sheets-источников в итоговую таблицу.

ВАЖНО: бизнес-логика подсчётов (analyze_single_sheet) сохранена 1-в-1 из
старой версии. Изменена только обёртка: безопасность, dry-run, backup,
проверка структуры, retry, логирование.

Ничего никогда не пишется за пределы A:M и (опционально) N1.
Никакие листы не удаляются. Никакие колонки/строки не дропаются.
Хвост блока чистится ТОЛЬКО если явно CLEAR_TAIL=true и пройдены все
проверки (drop ratio, не-пустой результат, нашёлся блок, корректный header).
"""

from __future__ import annotations

import base64
import hashlib
import json
import os
import re
import time
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple
from zoneinfo import ZoneInfo

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# =====================================================================
#                       Боевые ID и константы
# =====================================================================

# Эти ID — production. В ENVIRONMENT=test их использовать запрещено.
PRODUCTION_SPREADSHEET_IDS = frozenset({
    "1OgF4xLUqwSHs2S2NPXCJfsVgLM4V9W5c3yQq5lDoS-o",   # OUR_GRID_ID prod
    "18nfkpxiPG6xB7uLcCpwV8Qwmx7VJ3ii7uHUzfUh4L2Y",   # SUMMARY prod
    "1Qf4vPXqfpa83NkCrsTa0OHqv5PTKUB89_lzXIok1NWQ",   # YANDEX prod
})

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Колонки в итоговом блоке: A..M (13 колонок). Никогда не выходим за это.
BLOCK_WIDTH = 13

# Ожидаемые заголовки итогового блока (в порядке A..M).
EXPECTED_BLOCK_HEADERS = [
    "менеджеры",
    "офферты всего",
    "ип",
    "тоо",
    "договор есть",
    "акцепт/оплата",
    "акцепт %",
    "метка nib_sale",
    "метка nib",
    "метка 0",
    "пусто",
    "другое",
    "красные",
]

# Маркеры заголовка блока в итоговой таблице (для row_looks_like_header).
HEADER_MARKERS = [
    "менеджер", "менеджеры", "офферт", "офферты", "ип", "тоо",
    "договор", "акцепт", "акцепт %", "пусто", "другое", "красные",
    "метка",
]

# Ключевые слова для поиска колонок в ИСТОЧНИКАХ (Наша/Яндекс сетка).
SOURCE_COL_KEYWORDS = {
    "MANAGER":  ["менеджер", "сотрудник", "manager"],
    "OPF":      ["опф", "форма"],
    "CONTRACT": ["договор", "контракт"],
    "ACCEPT":   ["акцепт", "платежки", "оплата", "поехали"],
    "TAGS":     ["метки", "наличие метки", "nib"],
}

# Слова, которые в строке-данных трактуются как «красная».
RED_TEXT_MARKERS = ("красн", "red", "красный", "красная")


# =====================================================================
#               Русские названия месяцев (для AUTO_MONTH)
# =====================================================================
#
# В Settings и в названиях листов используется именительный падеж:
# "Январь 2026", "Февраль 2026", ... — поэтому генерируем строки сами,
# а не через locale-зависимый strftime("%B"), который ненадёжен.

RU_MONTH_NAMES = [
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
]
RU_MONTH_TO_IDX: Dict[str, int] = {n.lower(): i + 1 for i, n in enumerate(RU_MONTH_NAMES)}


def compute_auto_month(tz: str, offset: int) -> str:
    """
    Текущий месяц в TZ + offset месяцев.
    offset=0  -> текущий месяц
    offset=-1 -> предыдущий
    offset=1  -> следующий
    Возвращает строку формата "Июнь 2026".
    """
    dt = datetime.now(ZoneInfo(tz))
    total = dt.year * 12 + (dt.month - 1) + offset
    year = total // 12
    month_idx = total % 12 + 1
    return f"{RU_MONTH_NAMES[month_idx - 1]} {year}"


def parse_ru_month(s: str) -> Optional[Tuple[int, int]]:
    """'Июнь 2026' -> (2026, 6). Возвращает None при невалидном формате."""
    if not s:
        return None
    parts = s.strip().split()
    if len(parts) != 2:
        return None
    idx = RU_MONTH_TO_IDX.get(parts[0].lower())
    if idx is None:
        return None
    try:
        year = int(parts[1])
    except ValueError:
        return None
    return (year, idx)


def find_closest_settings_month(
    settings_pairs: List[Tuple[str, str, str]], target_month: str
) -> Optional[str]:
    """
    Используется при AUTO_DISCOVER_MONTH=true, если вычисленного месяца
    нет в Settings. Возвращает самый свежий месяц из Settings, который
    не позже target_month, либо None.
    """
    target = parse_ru_month(target_month)
    if not target:
        return None
    candidates: List[Tuple[Tuple[int, int], str]] = []
    for month_name, _, _ in settings_pairs:
        parsed = parse_ru_month(month_name)
        if parsed and parsed <= target:
            candidates.append((parsed, month_name))
    if not candidates:
        return None
    candidates.sort(reverse=True)
    return candidates[0][1]


# =====================================================================
#                              Config
# =====================================================================

def _env_bool(name: str, default: bool = False) -> bool:
    v = os.getenv(name)
    if v is None:
        return default
    return str(v).strip().lower() in ("1", "true", "yes", "y", "on")


def _env_str(name: str, default: str = "") -> str:
    v = os.getenv(name)
    return str(v).strip() if v is not None else default


def _env_int(name: str, default: int) -> int:
    v = os.getenv(name)
    if v is None or str(v).strip() == "":
        return default
    try:
        return int(v)
    except ValueError:
        return default


def _env_float(name: str, default: float) -> float:
    v = os.getenv(name)
    if v is None or str(v).strip() == "":
        return default
    try:
        return float(v)
    except ValueError:
        return default


def _env_list(name: str, default: Optional[List[str]] = None) -> List[str]:
    """Comma-separated list. Пустые значения отбрасываются."""
    v = os.getenv(name)
    if v is None or str(v).strip() == "":
        return list(default) if default else []
    return [s.strip() for s in v.split(",") if s.strip()]


def _env_json_dict(name: str, default: Optional[Dict[str, str]] = None) -> Dict[str, str]:
    """
    JSON-объект из env. Ключи нормализуются в lowercase (для регистронезависимого
    поиска). Значения остаются как есть — это канонические имена.
    Невалидный JSON => SystemExit (лучше упасть, чем тихо мискатегоризировать).
    """
    v = os.getenv(name)
    if v is None or str(v).strip() == "":
        return dict(default) if default else {}
    try:
        d = json.loads(v)
    except json.JSONDecodeError as e:
        raise SystemExit(f"{name} is not valid JSON: {e}")
    if not isinstance(d, dict):
        raise SystemExit(f"{name} must be a JSON object, got {type(d).__name__}")
    out: Dict[str, str] = {}
    for k, val in d.items():
        if not isinstance(k, str) or not isinstance(val, str):
            raise SystemExit(f"{name}: keys and values must be strings, got {k!r}->{val!r}")
        key = k.strip().lower()
        if key:
            out[key] = val.strip()
    return out


# =====================================================================
#                  Нормализация имён менеджеров
# =====================================================================
#
# Цель: безопасно матчить "Айша", " Айша ", "Айша\u00A0", "АЙША" в одну и
# ту же запись. Раньше мы делали только .strip().lower(), и NBSP/двойные
# пробелы/zero-width символы превращались в разные ключи, из-за чего
# alias не срабатывал и менеджер появлялся в сводной как ноль.
#
# Что делает normalize_text_key:
#   - заменяет NBSP/тонкие пробелы на обычный пробел
#   - удаляет zero-width символы (\u200B, \u200C, \u200D, \uFEFF)
#   - схлопывает множественные пробелы
#   - strip + strip пунктуации на краях (.,;: и т.п.)
#   - lower()
#   - ё -> е (частая опечатка / разные раскладки)
#
# Что НЕ делает (специально):
#   - не транслитерирует кириллицу/латиницу — это могло бы склеить разных
#     людей (рус. "Айша" vs набранное латиницей "Aйшa").
#   - не делает fuzzy-match — может склеить «Наташа» и «Наталья», это
#     разные люди. Для опечаток есть MANAGER_ALIASES.

_NBSP_LIKE = ("\u00A0", "\u202F", "\u2009", "\u200A", "\u2007", "\u2008")
_ZERO_WIDTH = ("\u200B", "\u200C", "\u200D", "\uFEFF")
_STRIP_PUNCT = ".,;:!?\"'«»()[]{}"


def normalize_text_key(value: Any) -> str:
    """
    Канонический ключ для сравнения имён.
    Идемпотентна: normalize_text_key(normalize_text_key(x)) == normalize_text_key(x).
    Возвращает пустую строку, если value None/пустой.
    """
    if value is None:
        return ""
    s = str(value)
    for ws in _NBSP_LIKE:
        s = s.replace(ws, " ")
    for zw in _ZERO_WIDTH:
        s = s.replace(zw, "")
    s = re.sub(r"\s+", " ", s).strip()
    s = s.strip(_STRIP_PUNCT)
    s = s.lower()
    s = s.replace("ё", "е")
    return s


def normalize_manager_display(name: Any) -> Optional[str]:
    """
    Аккуратное display-имя для случая, когда менеджер не нашёлся ни в
    aliases, ни в MANAGER_ORDER. В strict-режиме такие имена потом всё
    равно отфильтруются. Тут просто чистим внешние артефакты.
    """
    if name is None:
        return None
    s = str(name)
    for ws in _NBSP_LIKE:
        s = s.replace(ws, " ")
    for zw in _ZERO_WIDTH:
        s = s.replace(zw, "")
    s = re.sub(r"\s+", " ", s).strip()
    s = s.strip(_STRIP_PUNCT)
    if len(s) < 2:
        return None
    return s[:1].upper() + s[1:].lower()


# =====================================================================
#                              Config
# =====================================================================

@dataclass
class Config:
    # ---- режим работы ----
    environment: str = "test"                # test | production
    run_mode: str = "once"                   # once | loop
    dry_run: bool = True
    safe_mode: bool = True
    allow_production_write: bool = False

    # ---- spreadsheets ----
    our_grid_id: str = ""
    yandex_grid_id: str = ""
    summary_spreadsheet_id: str = ""

    expected_our_title: str = ""
    expected_yandex_title: str = ""
    expected_summary_title: str = ""

    summary_settings_sheet_name: str = "Settings"

    # ---- поведение ----
    require_yandex: bool = False
    allow_create_sheets: bool = False
    create_log_sheet: bool = False

    summary_write_start_col: str = "A"
    summary_write_end_col: str = "M"
    updated_at_cell: str = "N1"

    clear_tail: bool = False
    max_drop_ratio: float = 0.7
    red_gap_rows: int = 5
    max_data_rows: int = 5000          # верхняя граница строк, читаемых из источника

    # ---- backup ----
    backup_before_write: bool = True
    backup_dir: str = "backups"

    # ---- цикл ----
    loop_sleep_sec: int = 15
    hot_month: str = ""
    hot_write_interval_sec: int = 60
    cold_refresh_sec: int = 300

    # ---- креды (имена; сами секреты в env) ----
    creds_source: str = ""             # имя источника, чтобы залогировать

    # ---- runtime ----
    tz: str = "Asia/Almaty"
    target_month: Optional[str] = None  # CLI --month

    # ---- автоматический выбор месяца ----
    auto_month: bool = False
    month_offset: int = 0
    auto_discover_month: bool = False

    # ---- строгая фильтрация менеджеров ----
    manager_strict_order_only: bool = False
    include_zero_managers_from_order: bool = True
    manager_order: List[str] = field(default_factory=list)
    manager_aliases: Dict[str, str] = field(default_factory=dict)  # lower-key -> canonical

    # ---- диагностика менеджеров ----
    debug_managers: bool = False

    # Размеры блоков в итоговой таблице
    our_block_rows: int = 20            # сколько строк отведено под "Наша сетка"
    yandex_block_rows: int = 60         # сколько строк отведено под "Яндекс сетка"

    def __post_init__(self) -> None:
        """
        Пред-вычисляем нормализованные lookup-таблицы для O(1) матчинга.
        Это запускается автоматически при создании Config (в т.ч. в from_env).
        """
        # MANAGER_ALIASES: ключи через normalize_text_key (а не просто .lower()),
        # чтобы NBSP/двойные пробелы/zero-width не ломали поиск.
        self._normalized_aliases: Dict[str, str] = {}
        for k, v in (self.manager_aliases or {}).items():
            nk = normalize_text_key(k)
            nv = (v or "").strip()
            if nk and nv:
                self._normalized_aliases[nk] = nv

        # MANAGER_ORDER: то же самое. Значение — оригинальное имя (как пользователь
        # написал), чтобы оно попало в итог именно в этом виде.
        # Проверяем дубли после нормализации: "Айша" и " айша " — это одно и то же.
        self._normalized_order: Dict[str, str] = {}
        seen_keys: Dict[str, str] = {}
        for n in (self.manager_order or []):
            nk = normalize_text_key(n)
            stripped = (n or "").strip()
            if not (nk and stripped):
                continue
            if nk in seen_keys:
                raise SystemExit(
                    f"MANAGER_ORDER has duplicate entries after normalization: "
                    f"{seen_keys[nk]!r} and {stripped!r} both normalize to {nk!r}. "
                    f"Remove the duplicate."
                )
            seen_keys[nk] = stripped
            self._normalized_order[nk] = stripped

    @classmethod
    def from_env(cls) -> "Config":
        c = cls(
            environment=_env_str("ENVIRONMENT", "test").lower(),
            run_mode=_env_str("RUN_MODE", "once").lower(),
            dry_run=_env_bool("DRY_RUN", True),
            safe_mode=_env_bool("SAFE_MODE", True),
            allow_production_write=_env_bool("ALLOW_PRODUCTION_WRITE", False),

            our_grid_id=_env_str("OUR_GRID_ID"),
            yandex_grid_id=_env_str("YANDEX_GRID_ID"),
            summary_spreadsheet_id=_env_str("SUMMARY_SPREADSHEET_ID"),

            expected_our_title=_env_str("EXPECTED_OUR_TITLE"),
            expected_yandex_title=_env_str("EXPECTED_YANDEX_TITLE"),
            expected_summary_title=_env_str("EXPECTED_SUMMARY_TITLE"),

            summary_settings_sheet_name=_env_str("SUMMARY_SETTINGS_SHEET_NAME", "Settings"),

            require_yandex=_env_bool("REQUIRE_YANDEX", False),
            allow_create_sheets=_env_bool("ALLOW_CREATE_SHEETS", False),
            create_log_sheet=_env_bool("CREATE_LOG_SHEET", False),

            summary_write_start_col=_env_str("SUMMARY_WRITE_START_COL", "A"),
            summary_write_end_col=_env_str("SUMMARY_WRITE_END_COL", "M"),
            updated_at_cell=_env_str("UPDATED_AT_CELL", "N1"),

            clear_tail=_env_bool("CLEAR_TAIL", False),
            max_drop_ratio=_env_float("MAX_DROP_RATIO", 0.7),
            red_gap_rows=_env_int("RED_GAP_ROWS", 5),
            max_data_rows=_env_int("MAX_DATA_ROWS", 5000),

            backup_before_write=_env_bool("BACKUP_BEFORE_WRITE", True),
            backup_dir=_env_str("BACKUP_DIR", "backups"),

            loop_sleep_sec=_env_int("LOOP_SLEEP_SEC", 15),
            hot_month=_env_str("HOT_MONTH"),
            hot_write_interval_sec=_env_int("HOT_WRITE_INTERVAL_SEC", 60),
            cold_refresh_sec=_env_int("COLD_REFRESH_SEC", 300),

            tz=_env_str("TZ", "Asia/Almaty"),
            our_block_rows=_env_int("OUR_BLOCK_ROWS", 20),
            yandex_block_rows=_env_int("YANDEX_BLOCK_ROWS", 60),

            auto_month=_env_bool("AUTO_MONTH", False),
            month_offset=_env_int("MONTH_OFFSET", 0),
            auto_discover_month=_env_bool("AUTO_DISCOVER_MONTH", False),

            manager_strict_order_only=_env_bool("MANAGER_STRICT_ORDER_ONLY", False),
            include_zero_managers_from_order=_env_bool("INCLUDE_ZERO_MANAGERS_FROM_ORDER", True),
            manager_order=_env_list("MANAGER_ORDER"),
            manager_aliases=_env_json_dict("MANAGER_ALIASES"),

            debug_managers=_env_bool("DEBUG_MANAGERS", False),
        )
        return c

    def validate(self) -> None:
        """Жёсткая проверка конфигурации. Бросает SystemExit при опасной комбинации."""
        if self.environment not in ("test", "production"):
            raise SystemExit(f"Invalid ENVIRONMENT={self.environment!r} (test|production)")

        if self.run_mode not in ("once", "loop"):
            raise SystemExit(f"Invalid RUN_MODE={self.run_mode!r} (once|loop)")

        if not self.our_grid_id:
            raise SystemExit("OUR_GRID_ID is empty")
        if not self.summary_spreadsheet_id:
            raise SystemExit("SUMMARY_SPREADSHEET_ID is empty")
        # YANDEX можно оставить пустым только если REQUIRE_YANDEX=false
        if self.require_yandex and not self.yandex_grid_id:
            raise SystemExit("REQUIRE_YANDEX=true but YANDEX_GRID_ID is empty")

        # Запрет боевых ID в test-режиме
        if self.environment == "test":
            for label, sid in (
                ("OUR_GRID_ID", self.our_grid_id),
                ("YANDEX_GRID_ID", self.yandex_grid_id),
                ("SUMMARY_SPREADSHEET_ID", self.summary_spreadsheet_id),
            ):
                if sid and sid in PRODUCTION_SPREADSHEET_IDS:
                    raise SystemExit(
                        f"ENVIRONMENT=test but production spreadsheet ID was provided "
                        f"({label}={sid}). Use a copy."
                    )

        # В production без явного разрешения — только dry-run
        if self.environment == "production" and not self.allow_production_write:
            if not self.dry_run:
                print("[SAFETY] ENVIRONMENT=production but ALLOW_PRODUCTION_WRITE=false "
                      "-> forcing DRY_RUN=true")
                self.dry_run = True

        if self.max_drop_ratio < 0 or self.max_drop_ratio > 1:
            raise SystemExit(f"MAX_DROP_RATIO must be in [0..1], got {self.max_drop_ratio}")

        if self.red_gap_rows < 1:
            raise SystemExit(f"RED_GAP_ROWS must be >= 1, got {self.red_gap_rows}")

        # MONTH_OFFSET — целое в разумных пределах
        if abs(self.month_offset) > 120:
            raise SystemExit(
                f"MONTH_OFFSET={self.month_offset} is unreasonable (max ±120 months)"
            )

        # MANAGER_STRICT_ORDER_ONLY=true без списка — опасно (всё будет отфильтровано)
        if self.manager_strict_order_only and not self.manager_order:
            raise SystemExit(
                "MANAGER_STRICT_ORDER_ONLY=true but MANAGER_ORDER is empty. "
                "Set MANAGER_ORDER=Имя1,Имя2,... or turn off strict mode."
            )

        # Проверка, что aliases не мапят на пустые строки
        for k, v in self.manager_aliases.items():
            if not v:
                raise SystemExit(f"MANAGER_ALIASES has empty value for key {k!r}")

        # Если strict-режим и alias мапит на значение, которого нет в MANAGER_ORDER,
        # такие строки будут тихо отфильтрованы. Это легко не заметить — печатаем
        # WARN один раз при старте (а не в каждом analyze_single_sheet).
        if self.manager_strict_order_only and self._normalized_order:
            outside = sorted({
                v for v in self._normalized_aliases.values()
                if normalize_text_key(v) not in self._normalized_order
            })
            if outside:
                print(f"[WARN] MANAGER_ALIASES has values not in MANAGER_ORDER: "
                      f"{outside}. In strict mode rows mapped to these names "
                      f"will be EXCLUDED. Either add them to MANAGER_ORDER or "
                      f"remove from MANAGER_ALIASES.")

    def banner(self) -> str:
        resolved = self.resolve_target_month()
        if self.target_month:
            month_line = f"target_month     : {self.target_month}  (from CLI --month)"
        elif self.auto_month:
            month_line = (f"target_month     : {resolved}  "
                          f"(AUTO_MONTH=true, MONTH_OFFSET={self.month_offset})")
        else:
            month_line = "target_month     : <all months from Settings>"

        lines = [
            "=" * 64,
            f" Google Sheets Summary Sync ",
            "=" * 64,
            f" timestamp        : {self.now_local().isoformat()}",
            f" ENVIRONMENT      : {self.environment}",
            f" RUN_MODE         : {self.run_mode}",
            f" DRY_RUN          : {self.dry_run}",
            f" SAFE_MODE        : {self.safe_mode}",
            f" ALLOW_PROD_WRITE : {self.allow_production_write}",
            f" REQUIRE_YANDEX   : {self.require_yandex}",
            f" CLEAR_TAIL       : {self.clear_tail}",
            f" BACKUP_BEFORE_WR : {self.backup_before_write}",
            f" RED_GAP_ROWS     : {self.red_gap_rows}",
            f" MAX_DROP_RATIO   : {self.max_drop_ratio}",
            f" MAX_DATA_ROWS    : {self.max_data_rows}",
            f" AUTO_MONTH       : {self.auto_month}  (MONTH_OFFSET={self.month_offset}, "
            f"AUTO_DISCOVER_MONTH={self.auto_discover_month})",
            f" MGR_STRICT       : {self.manager_strict_order_only}  "
            f"(order={len(self.manager_order)}, aliases={len(self.manager_aliases)}, "
            f"include_zero={self.include_zero_managers_from_order})",
            f" DEBUG_MANAGERS   : {self.debug_managers}",
            f" {month_line}",
            f" OUR_GRID_ID      : {_mask_id(self.our_grid_id)}",
            f" YANDEX_GRID_ID   : {_mask_id(self.yandex_grid_id) or '<empty>'}",
            f" SUMMARY_ID       : {_mask_id(self.summary_spreadsheet_id)}",
            f" creds_source     : {self.creds_source or '<auto>'}",
            "=" * 64,
        ]
        return "\n".join(lines)

    def resolve_target_month(self) -> Optional[str]:
        """
        Какой месяц обрабатывать.

        Приоритет:
          1. CLI --month (config.target_month) — самый высокий.
          2. AUTO_MONTH=true → compute_auto_month(tz, month_offset).
          3. Иначе None → run_summary_once обработает все месяцы из Settings.
        """
        if self.target_month:
            return self.target_month
        if self.auto_month:
            return compute_auto_month(self.tz, self.month_offset)
        return None

    def now_local(self) -> datetime:
        return datetime.now(ZoneInfo(self.tz))


def _mask_id(sid: str) -> str:
    if not sid:
        return ""
    if len(sid) <= 8:
        return "***"
    return sid[:4] + "…" + sid[-4:]


# =====================================================================
#                          Авторизация Google
# =====================================================================

def _validate_sa_info(info: Dict[str, Any]) -> None:
    """Проверяем минимальный набор полей у service account JSON."""
    if not isinstance(info, dict):
        raise SystemExit("Service account JSON is not an object")
    missing = [k for k in ("client_email", "token_uri", "private_key") if not info.get(k)]
    if missing:
        raise SystemExit(f"Service account JSON missing fields: {', '.join(missing)}")


def build_sheets_service(config: Config):
    """
    Приоритет источников creds:
      1. GCP_SA_JSON      — JSON service account прямо в env
      2. GOOGLE_CREDS_B64 — base64 от JSON service account
      3. GOOGLE_APPLICATION_CREDENTIALS — путь к JSON-файлу
      4. Application Default Credentials
    """
    sa_json = os.getenv("GCP_SA_JSON")
    if sa_json:
        try:
            info = json.loads(sa_json)
        except json.JSONDecodeError as e:
            raise SystemExit(f"GCP_SA_JSON is not valid JSON: {e}")
        _validate_sa_info(info)
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        config.creds_source = "GCP_SA_JSON"
        return build("sheets", "v4", credentials=creds, cache_discovery=False)

    b64 = os.getenv("GOOGLE_CREDS_B64")
    if b64:
        try:
            info = json.loads(base64.b64decode(b64).decode("utf-8"))
        except Exception as e:
            raise SystemExit(f"GOOGLE_CREDS_B64 is not valid base64 JSON: {e}")
        _validate_sa_info(info)
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        config.creds_source = "GOOGLE_CREDS_B64"
        return build("sheets", "v4", credentials=creds, cache_discovery=False)

    path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    if path:
        if not os.path.isfile(path):
            raise SystemExit(f"GOOGLE_APPLICATION_CREDENTIALS path not found: {path}")
        try:
            with open(path, "r", encoding="utf-8") as f:
                info = json.load(f)
        except Exception as e:
            raise SystemExit(f"Cannot read service account file: {e}")
        _validate_sa_info(info)
        creds = service_account.Credentials.from_service_account_file(path, scopes=SCOPES)
        config.creds_source = f"GOOGLE_APPLICATION_CREDENTIALS={path}"
        return build("sheets", "v4", credentials=creds, cache_discovery=False)

    # ADC
    try:
        import google.auth
        creds, _ = google.auth.default(scopes=SCOPES)
    except Exception as e:
        raise SystemExit(
            "No credentials provided. Set one of: "
            "GCP_SA_JSON / GOOGLE_CREDS_B64 / GOOGLE_APPLICATION_CREDENTIALS. "
            f"ADC fallback failed: {e}"
        )
    config.creds_source = "ADC"
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


# =====================================================================
#                       API helpers + retry
# =====================================================================

# 403/404 — НЕ ретраим, выбрасываем сразу с понятным сообщением.
# 429/500/502/503/504 — ретраим до 3 раз с экспоненциальной паузой.
RETRYABLE_STATUSES = {429, 500, 502, 503, 504}


def _http_status(e: HttpError) -> Optional[int]:
    try:
        return int(e.resp.status)
    except Exception:
        return None


def _explain_http_error(e: HttpError, action: str) -> str:
    status = _http_status(e)
    if status == 403:
        return (f"403 Permission denied while {action}. "
                f"Проверь что у service account есть доступ Editor к таблице.")
    if status == 404:
        return f"404 Not found while {action}. Проверь spreadsheet ID и название листа."
    if status == 401:
        return f"401 Unauthorized while {action}. Проверь creds."
    return f"HTTP {status} while {action}: {e}"


def call_with_retry(fn, action: str, max_retries: int = 3, base_sleep: float = 1.5):
    """
    Запускает fn() с ретраями только на 429/500/502/503/504.
    На 403/404/401 — сразу бросает с понятным сообщением.
    """
    last_exc: Optional[Exception] = None
    for attempt in range(1, max_retries + 1):
        try:
            return fn()
        except HttpError as e:
            status = _http_status(e)
            if status in (403, 404, 401):
                raise SystemExit("[FATAL] " + _explain_http_error(e, action))
            if status in RETRYABLE_STATUSES and attempt < max_retries:
                sleep_s = base_sleep * (2 ** (attempt - 1))
                print(f"[WARN] {action}: HTTP {status}, retry {attempt}/{max_retries} "
                      f"in {sleep_s:.1f}s")
                time.sleep(sleep_s)
                last_exc = e
                continue
            # не ретраимая или закончились попытки
            raise
        except Exception as e:
            last_exc = e
            raise
    if last_exc:
        raise last_exc
    raise RuntimeError(f"Retry loop exhausted for: {action}")


def read_values(service, spreadsheet_id: str, a1_range: str) -> List[List[Any]]:
    def _call():
        return service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=a1_range,
            valueRenderOption="FORMATTED_VALUE",
        ).execute()
    resp = call_with_retry(_call, f"read {a1_range} from {_mask_id(spreadsheet_id)}")
    return resp.get("values", [])


def write_values(service, spreadsheet_id: str, a1_range: str, values: List[List[Any]]) -> None:
    def _call():
        return service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=a1_range,
            valueInputOption="RAW",
            body={"values": values},
        ).execute()
    call_with_retry(_call, f"write {a1_range} to {_mask_id(spreadsheet_id)}")


def get_spreadsheet_meta(service, spreadsheet_id: str) -> Dict[str, Any]:
    def _call():
        return service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    return call_with_retry(_call, f"get meta {_mask_id(spreadsheet_id)}")


def get_sheet_titles(service, spreadsheet_id: str) -> Dict[str, str]:
    """lower_name -> real_name"""
    meta = get_spreadsheet_meta(service, spreadsheet_id)
    return {
        s["properties"]["title"].lower(): s["properties"]["title"]
        for s in meta.get("sheets", [])
    }


def get_spreadsheet_title(meta: Dict[str, Any]) -> str:
    return meta.get("properties", {}).get("title", "")


def find_sheet_smart(titles_lower: Dict[str, str], partial_name: str) -> Optional[str]:
    if not partial_name:
        return None
    search = str(partial_name).strip().lower()
    if search in titles_lower:
        return titles_lower[search]
    search_clean = re.sub(r"\s+", "", search)
    for low, real in titles_lower.items():
        clean = re.sub(r"\s+", "", low)
        if search_clean and (search_clean in clean or clean in search_clean):
            return real
    return None


# =====================================================================
#               Проверка title и структуры (safety gates)
# =====================================================================

def verify_spreadsheet_title(
    service, spreadsheet_id: str, expected_title: str, label: str
) -> str:
    """Возвращает фактический title. Если ожидаемый задан и не совпадает — SystemExit."""
    meta = get_spreadsheet_meta(service, spreadsheet_id)
    actual = get_spreadsheet_title(meta)
    if expected_title:
        if actual.strip() != expected_title.strip():
            raise SystemExit(
                f"[FATAL] {label}: spreadsheet title mismatch. "
                f"Expected={expected_title!r}, actual={actual!r}. "
                f"Refusing to continue."
            )
        print(f"[OK] {label} title verified: {actual!r}")
    else:
        print(f"[INFO] {label} title: {actual!r} (no EXPECTED_*_TITLE set)")
    return actual


def row_looks_like_header(row_vals: List[Any]) -> bool:
    if not row_vals:
        return False
    txt = " ".join(str(x).strip().lower() for x in row_vals if str(x).strip())
    if not txt:
        return False
    hits = sum(1 for m in HEADER_MARKERS if m in txt)
    return hits >= 2 or ("менедж" in txt)


def find_title_row(
    service, summary_id: str, sheet_title: str, label: str, search_rows: int = 200
) -> Optional[int]:
    """1-based номер строки, где встретился label (case-insensitive)."""
    vals = read_values(service, summary_id, f"{sheet_title}!A1:M{search_rows}")
    lab = label.lower()
    for i, row in enumerate(vals, start=1):
        txt = " ".join(str(x).strip().lower() for x in row if str(x).strip())
        if lab in txt:
            return i
    return None


def header_matches_expected(header_row: List[Any]) -> Tuple[bool, str]:
    """Допускаем небольшие отличия регистра/пробелов."""
    if not header_row:
        return False, "header row is empty"
    norm = [re.sub(r"\s+", " ", str(c)).strip().lower() for c in header_row[:BLOCK_WIDTH]]
    expected = EXPECTED_BLOCK_HEADERS
    # Должны совпадать хотя бы 10 из 13 (точное совпадение, регистро-нечувствительное).
    # Это терпит мелкие опечатки. Но "менеджеры" в A — обязательное условие.
    if not norm or "менедж" not in norm[0]:
        return False, f"col A is not 'Менеджеры', got {norm[0]!r}"
    hits = 0
    for i, exp in enumerate(expected):
        if i < len(norm) and norm[i] == exp:
            hits += 1
    if hits < 10:
        return False, (f"only {hits}/13 expected columns match. "
                       f"Got: {norm}. Expected: {expected}")
    return True, "ok"


@dataclass
class BlockLocation:
    label: str
    title_row: int        # строка с "НАША СЕТКА"/"ЯНДЕКС СЕТКА"
    header_row: int       # строка заголовков
    data_start_row: int   # первая строка данных (managers)
    data_max_rows: int    # сколько строк под данные отведено в блоке


def locate_block(
    service, config: Config, real_sheet_title: str, label: str, block_height_hint: int,
    next_block_title_row: Optional[int] = None
) -> Optional[BlockLocation]:
    """
    Ищет блок по тексту title. Проверяет, что строкой ниже идёт header.
    Возвращает None и пишет в лог, если что-то не так (SAFE_MODE => не писать).
    """
    title_row = find_title_row(service, config.summary_spreadsheet_id, real_sheet_title, label)
    if not title_row:
        print(f"[WARN] Block title {label!r} not found in {real_sheet_title!r}")
        return None

    header_row = title_row + 1
    data_start = title_row + 2

    # Проверим, что header_row действительно "пахнет header'ом" и совпадает с expected
    header_vals = read_values(
        service, config.summary_spreadsheet_id,
        f"{real_sheet_title}!A{header_row}:M{header_row}",
    )
    header = header_vals[0] if header_vals else []

    ok, reason = header_matches_expected(header)
    if not ok:
        print(f"[WARN] Block {label!r} header check failed: {reason}")
        return None

    # Защита: если data_start вдруг тоже похожа на header — отказываемся писать.
    ds_vals = read_values(
        service, config.summary_spreadsheet_id,
        f"{real_sheet_title}!A{data_start}:M{data_start}",
    )
    ds_row = ds_vals[0] if ds_vals else []
    if row_looks_like_header(ds_row):
        print(f"[WARN] Block {label!r}: data_start_row={data_start} looks like header. "
              f"Refusing to write.")
        return None

    # Сколько строк отвести под данные
    if next_block_title_row is not None and next_block_title_row > data_start:
        # данные не должны заходить на title следующего блока
        data_max_rows = max(1, next_block_title_row - 1 - data_start + 1 - 1)
    else:
        data_max_rows = block_height_hint

    return BlockLocation(
        label=label,
        title_row=title_row,
        header_row=header_row,
        data_start_row=data_start,
        data_max_rows=data_max_rows,
    )


# =====================================================================
#         Бизнес-логика подсчёта (ПЕРЕНЕСЕНА БЕЗ ИЗМЕНЕНИЙ)
# =====================================================================

def _normalize_manager_name(
    name: Any, aliases: Optional[Dict[str, str]] = None,
) -> Optional[str]:
    """
    LEGACY-функция. Оставлена для обратной совместимости (используется только
    в старых тестах). В рабочем коде используется canonical_manager_name(),
    который применяет normalize_text_key и предвычисленные lookup'ы из Config.
    """
    if not name:
        return None
    s = str(name).strip()
    if len(s) < 2:
        return None
    if aliases:
        low = s.lower()
        if low in aliases:
            return aliases[low]
    return s[:1].upper() + s[1:].lower()


def canonical_manager_name(raw_name: Any, config: "Config") -> Optional[str]:
    """
    Каноническая форма имени менеджера.

    Алгоритм:
      1. raw_key = normalize_text_key(raw_name)
      2. если raw_key пустой → None
      3. если raw_key есть в config._normalized_aliases → значение alias
         (alias переопределяет всё)
      4. если raw_key есть в config._normalized_order → каноническое имя
         из MANAGER_ORDER (с пользовательским регистром / пробелами)
      5. иначе normalize_manager_display(raw_name) — display-вариант,
         который в strict-режиме потом отфильтруется.

    Это единственная точка, через которую сырое имя из источника попадает
    в stats. NBSP, двойные пробелы, ё/е, регистр — всё схлопывается в
    один и тот же ключ, и подсчёты больше не «теряют» менеджера из-за
    невидимого артефакта.
    """
    raw_key = normalize_text_key(raw_name)
    if not raw_key:
        return None
    aliases = getattr(config, "_normalized_aliases", None) or {}
    if raw_key in aliases:
        return aliases[raw_key]
    order = getattr(config, "_normalized_order", None) or {}
    if raw_key in order:
        return order[raw_key]
    return normalize_manager_display(raw_name)


def _find_idx(headers: List[str], keywords: List[str]) -> int:
    for i, h in enumerate(headers):
        for k in keywords:
            if k in h:
                return i
    return -1


def analyze_single_sheet(
    service, source_id: str, source_titles_lower: Dict[str, str],
    sheet_name: str, config: Config,
) -> List[List[Any]]:
    """
    БИЗНЕС-ЛОГИКА СОХРАНЕНА 1-в-1 ИЗ ИСХОДНОГО КОДА.

    Возвращает список строк итоговой таблицы (по 13 значений в каждой):
      [manager, total, ip, too, contract, accept, percent_str,
       nib_sale, nib, zero, empty_tag, other_tag, red]

    Сортировано по имени менеджера.
    """
    if not sheet_name:
        return []
    real_name = find_sheet_smart(source_titles_lower, sheet_name)
    if not real_name:
        print(f"[INFO] Source sheet {sheet_name!r} not found in {_mask_id(source_id)}, skipped.")
        return []

    # Защита от безумных объёмов: ограничиваем диапазон чтения сверху.
    a1 = f"{real_name}!A1:Z{config.max_data_rows + 1}"
    data = read_values(service, source_id, a1)
    if len(data) < 2:
        print(f"[INFO] Source sheet {real_name!r} is empty, skipped.")
        return []

    headers = [str(h).lower().strip() for h in data[0]]

    idx = {
        "man":      _find_idx(headers, SOURCE_COL_KEYWORDS["MANAGER"]),
        "opf":      _find_idx(headers, SOURCE_COL_KEYWORDS["OPF"]),
        "contract": _find_idx(headers, SOURCE_COL_KEYWORDS["CONTRACT"]),
        "accept":   _find_idx(headers, SOURCE_COL_KEYWORDS["ACCEPT"]),
        "tags":     _find_idx(headers, SOURCE_COL_KEYWORDS["TAGS"]),
    }

    # Иногда первая строка не header, а заголовок — пробуем 2-ю.
    if idx["man"] == -1 and len(data) > 2:
        headers2 = [str(h).lower().strip() for h in data[1]]
        idx["man"] = _find_idx(headers2, SOURCE_COL_KEYWORDS["MANAGER"])

    if idx["man"] == -1:
        print(f"[WARN] Manager column not found in {real_name!r}")
        return []

    stats: Dict[str, Dict[str, int]] = {}
    is_red_section = False
    consecutive_empty_rows = 0

    # Диагностические счётчики (используются для DEBUG_MANAGERS-логов и для
    # подсказок про excluded). Лёгкие, считаются всегда — на больших листах
    # ~ничтожный оверхед.
    raw_counter: "Counter[str]" = Counter()
    canonical_counter: "Counter[str]" = Counter()

    for i in range(1, len(data)):
        row = data[i]
        manager_raw = row[idx["man"]] if idx["man"] < len(row) else ""

        if not manager_raw or str(manager_raw).strip() == "":
            consecutive_empty_rows += 1
            if consecutive_empty_rows >= config.red_gap_rows:
                is_red_section = True
            continue
        else:
            consecutive_empty_rows = 0

        manager = canonical_manager_name(manager_raw, config)
        if not manager:
            continue

        # Диагностика: считаем сырое имя как пользователь его видит (после
        # минимального cleanup — replace NBSP, collapse spaces, strip).
        # Lower не делаем, чтобы в логах было видно реальное написание.
        raw_for_log = re.sub(r"\s+", " ",
                             str(manager_raw).replace("\u00A0", " ")).strip() or "(empty)"
        raw_counter[raw_for_log] += 1
        canonical_counter[manager] += 1

        if manager not in stats:
            stats[manager] = {
                "total": 0, "ip": 0, "too": 0, "contract": 0, "accept": 0,
                "nib_sale": 0, "nib": 0, "zero": 0, "empty_tag": 0,
                "other_tag": 0, "red": 0,
            }
        s = stats[manager]

        # В красной секции — НЕ считаем обычные показатели
        if is_red_section:
            s["red"] += 1
            continue

        s["total"] += 1

        # ИП / ТОО — как в оригинале (поиск по тексту строки + колонке ОПФ)
        opf_text = ""
        if idx["opf"] > -1 and idx["opf"] < len(row):
            opf_text += str(row[idx["opf"]]).lower()
        opf_text += " " + " ".join(str(x).lower() for x in row)

        if ("ип " in opf_text) or ('ип"' in opf_text) or ("жк " in opf_text):
            s["ip"] += 1
        if "тоо" in opf_text:
            s["too"] += 1

        # Договор есть: значение НЕ в ("", "нет", "0", "-", "—")
        if idx["contract"] > -1 and idx["contract"] < len(row):
            val = str(row[idx["contract"]]).lower().strip()
            if val not in ("", "нет", "0", "-", "—"):
                s["contract"] += 1

        # Акцепт/Оплата: длина > 1 и нет "нет"/"отказ"/"ошибка"
        if idx["accept"] > -1 and idx["accept"] < len(row):
            val = str(row[idx["accept"]]).lower()
            if len(val) > 1 and ("нет" not in val) and ("отказ" not in val) and ("ошибка" not in val):
                s["accept"] += 1

        # Метки: nib_sale -> nib -> "0"/"0.0" -> пусто -> другое
        tag_val = ""
        if idx["tags"] > -1 and idx["tags"] < len(row):
            tag_val = str(row[idx["tags"]]).lower().strip()
        if "nib_sale" in tag_val:
            s["nib_sale"] += 1
        elif tag_val == "nib" or " nib " in f" {tag_val} ":
            s["nib"] += 1
        elif tag_val in ("0", "0.0"):
            s["zero"] += 1
        elif tag_val == "":
            s["empty_tag"] += 1
        else:
            s["other_tag"] += 1

        # Дополнительная пометка "красная" по тексту строки.
        # В оригинале это ДОБАВКА (не эксклюзивно). Сохраняем поведение.
        row_text = " ".join(str(x).lower() for x in row)
        if any(m in row_text for m in RED_TEXT_MARKERS):
            s["red"] += 1

    # ---- финальное представление строк ----
    def _build_row(name: str, s: Dict[str, int]) -> List[Any]:
        percent = (s["accept"] / s["total"]) if s["total"] > 0 else 0
        percent_str = f"{round(percent * 100)}%"
        return [
            name, s["total"], s["ip"], s["too"], s["contract"], s["accept"], percent_str,
            s["nib_sale"], s["nib"], s["zero"], s["empty_tag"], s["other_tag"], s["red"],
        ]

    # Шаблон "нулевого" менеджера — используется в STRICT при include_zero=true
    def _zero_stats() -> Dict[str, int]:
        return {
            "total": 0, "ip": 0, "too": 0, "contract": 0, "accept": 0,
            "nib_sale": 0, "nib": 0, "zero": 0, "empty_tag": 0,
            "other_tag": 0, "red": 0,
        }

    # ---- DEBUG_MANAGERS: подробная диагностика менеджеров ----
    if config.debug_managers:
        # Топ-N, чтобы лог не вырастал бесконтрольно
        TOPN = 30
        raw_top = raw_counter.most_common(TOPN)
        can_top = canonical_counter.most_common(TOPN)
        rows_read = max(0, len(data) - 1)
        print(f"[DEBUG_MANAGER] sheet={real_name!r} manager_col={idx['man']} "
              f"rows={rows_read} red_section_at_end={is_red_section}")
        print(f"[DEBUG_MANAGER] raw managers (top {TOPN}, "
              f"unique={len(raw_counter)}): {raw_top}")
        print(f"[DEBUG_MANAGER] canonical managers (top {TOPN}, "
              f"unique={len(canonical_counter)}): {can_top}")
        if config.manager_strict_order_only and config.manager_order:
            order_norm = set(config._normalized_order.keys())
            canonical_norm = {normalize_text_key(n) for n in canonical_counter}
            missing = [n for n in config.manager_order
                       if normalize_text_key(n) not in canonical_norm]
            excluded = [n for n in canonical_counter
                        if normalize_text_key(n) not in order_norm]
            print(f"[DEBUG_MANAGER] order missing (no rows in source): {missing}")
            print(f"[DEBUG_MANAGER] excluded (in source but not in order): {excluded}")
            if excluded:
                # Подсказка по тому, как починить:
                hint = excluded[0]
                hint_key = normalize_text_key(hint)
                print(f"[DEBUG_MANAGER] hint: if any of these should be in MANAGER_ORDER, "
                      f"add to MANAGER_ALIASES, e.g. "
                      f"{{ {hint_key!r}: \"CanonicalName\" }}.")

    # ---- STRICT ORDER: только менеджеры из MANAGER_ORDER, строго в его порядке ----
    if config.manager_strict_order_only and config.manager_order:
        # Все имена сравниваем по normalize_text_key (NBSP/двойные пробелы/ё/Е/регистр).
        # Канонический вид берём из MANAGER_ORDER.
        stats_by_key: Dict[str, Tuple[str, Dict[str, int]]] = {
            normalize_text_key(k): (k, v) for k, v in stats.items()
        }
        order_keys: Dict[str, str] = {
            normalize_text_key(n): n for n in config.manager_order
        }
        result: List[List[Any]] = []
        matched_keys: set = set()

        for canonical in config.manager_order:
            key = normalize_text_key(canonical)
            if key in stats_by_key:
                _, s = stats_by_key[key]
                # ВАЖНО: пишем имя из MANAGER_ORDER (как пользователь его задал),
                # а не вариант из источника. Подсчёты при этом не меняются.
                result.append(_build_row(canonical, s))
                matched_keys.add(key)
            else:
                if config.include_zero_managers_from_order:
                    result.append(_build_row(canonical, _zero_stats()))

        # Лог: кого выкинули — диагностика опечаток или пропущенных aliases
        excluded = [
            name for key, (name, _s) in stats_by_key.items()
            if key not in matched_keys
        ]
        if excluded:
            print(f"[INFO] STRICT_ORDER ({real_name}): filtered out "
                  f"{len(excluded)} managers not in MANAGER_ORDER: {excluded}. "
                  f"Add them to MANAGER_ORDER or to MANAGER_ALIASES.")

        return result

    # ---- Старое поведение: всё, что нашли, по алфавиту ----
    result = []
    for m, s in stats.items():
        result.append(_build_row(m, s))
    result.sort(key=lambda x: str(x[0]))
    return result


# =====================================================================
#                Запись/Очистка/Backup итогового блока
# =====================================================================

def _pad_or_trim_row(row: List[Any], width: int = BLOCK_WIDTH) -> List[Any]:
    row = list(row) if row else []
    if len(row) < width:
        row += [""] * (width - len(row))
    elif len(row) > width:
        row = row[:width]
    return row


def _count_filled_rows(block_values: List[List[Any]]) -> int:
    n = 0
    for r in block_values:
        if r and any(str(c).strip() for c in r):
            n += 1
    return n


def read_block_old_values(
    service, config: Config, sheet_title: str, loc: BlockLocation,
) -> List[List[Any]]:
    """Читает текущее содержимое блока A:M — нужно для backup и drop-ratio."""
    end_row = loc.data_start_row + loc.data_max_rows - 1
    a1 = f"{sheet_title}!A{loc.data_start_row}:M{end_row}"
    return read_values(service, config.summary_spreadsheet_id, a1)


def backup_block_to_disk(
    config: Config, sheet_title: str, loc: BlockLocation, values: List[List[Any]],
) -> Optional[str]:
    """Сохраняет JSON-снимок блока. Возвращает путь или None при ошибке."""
    try:
        os.makedirs(config.backup_dir, exist_ok=True)
        ts = config.now_local().strftime("%Y%m%d_%H%M%S")
        safe_sheet = re.sub(r"[^A-Za-z0-9_.-]+", "_", sheet_title)
        safe_label = re.sub(r"[^A-Za-z0-9_.-]+", "_", loc.label)
        fname = f"backup_{ts}_{safe_sheet}_{safe_label}.json"
        path = os.path.join(config.backup_dir, fname)
        end_row = loc.data_start_row + loc.data_max_rows - 1
        payload = {
            "timestamp": config.now_local().isoformat(),
            "spreadsheet_id": config.summary_spreadsheet_id,
            "sheet_name": sheet_title,
            "block_name": loc.label,
            "range": f"{sheet_title}!A{loc.data_start_row}:M{end_row}",
            "values": values,
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        return path
    except Exception as e:
        print(f"[WARN] backup failed for {loc.label}: {e}")
        return None


@dataclass
class WriteDecision:
    will_write: bool
    will_clear: bool
    rows_to_write: int
    rows_to_clear: int
    old_filled: int
    new_filled: int
    skip_reason: str = ""
    range_write: str = ""
    range_clear: str = ""


def decide_write(
    config: Config, sheet_title: str, loc: BlockLocation,
    new_rows: List[List[Any]], old_values: List[List[Any]],
) -> WriteDecision:
    """Решаем, можно ли писать. Никакой записи здесь не делается."""
    new_filled = len(new_rows)
    old_filled = _count_filled_rows(old_values)

    # 1) если новых данных 0 — никогда не пишем и не чистим
    if new_filled == 0:
        return WriteDecision(
            will_write=False, will_clear=False,
            rows_to_write=0, rows_to_clear=0,
            old_filled=old_filled, new_filled=0,
            skip_reason="source returned 0 rows (refusing to touch the block)",
        )

    # 2) drop ratio
    if old_filled > 0:
        drop = 1 - (new_filled / old_filled)
        if drop > config.max_drop_ratio:
            return WriteDecision(
                will_write=False, will_clear=False,
                rows_to_write=new_filled, rows_to_clear=0,
                old_filled=old_filled, new_filled=new_filled,
                skip_reason=(f"suspicious drop {drop:.0%} > MAX_DROP_RATIO "
                             f"{config.max_drop_ratio:.0%} "
                             f"(old={old_filled}, new={new_filled})"),
            )

    # 3) превышение размера блока — обрежем
    rows_to_write = min(new_filled, loc.data_max_rows)
    if new_filled > loc.data_max_rows:
        print(f"[WARN] {loc.label}: {new_filled} rows to write, but block can only "
              f"hold {loc.data_max_rows} rows (bounded by next block / block_height_hint). "
              f"Truncating to first {loc.data_max_rows}.")

    # 4) очистка хвоста — только если разрешено
    rows_to_clear = 0
    range_clear = ""
    if config.clear_tail and rows_to_write < loc.data_max_rows:
        tail_start = loc.data_start_row + rows_to_write
        tail_end = loc.data_start_row + loc.data_max_rows - 1
        if tail_end >= tail_start:
            rows_to_clear = tail_end - tail_start + 1
            range_clear = f"{sheet_title}!A{tail_start}:M{tail_end}"

    end_row = loc.data_start_row + rows_to_write - 1
    range_write = f"{sheet_title}!A{loc.data_start_row}:M{end_row}"

    return WriteDecision(
        will_write=True, will_clear=(rows_to_clear > 0),
        rows_to_write=rows_to_write, rows_to_clear=rows_to_clear,
        old_filled=old_filled, new_filled=new_filled,
        range_write=range_write, range_clear=range_clear,
    )


def apply_write_decision(
    service, config: Config, sheet_title: str, loc: BlockLocation,
    new_rows: List[List[Any]], decision: WriteDecision,
) -> None:
    """Реально пишет (если не dry-run)."""
    if not decision.will_write:
        return

    rows = [_pad_or_trim_row(r, BLOCK_WIDTH) for r in new_rows[:decision.rows_to_write]]
    if config.dry_run:
        print(f"[DRY] would WRITE {decision.rows_to_write} rows to {decision.range_write}")
    else:
        write_values(service, config.summary_spreadsheet_id, decision.range_write, rows)
        print(f"[OK]  WROTE {decision.rows_to_write} rows to {decision.range_write}")

    if decision.will_clear and decision.range_clear:
        blanks = [[""] * BLOCK_WIDTH for _ in range(decision.rows_to_clear)]
        if config.dry_run:
            print(f"[DRY] would CLEAR tail {decision.rows_to_clear} rows "
                  f"at {decision.range_clear}")
        else:
            write_values(service, config.summary_spreadsheet_id,
                         decision.range_clear, blanks)
            print(f"[OK]  CLEARED tail {decision.rows_to_clear} rows "
                  f"at {decision.range_clear}")


# =====================================================================
#                            Sync_Log
# =====================================================================

LOG_HEADERS = [
    "Timestamp", "Environment", "DryRun", "SafeMode", "Month", "Block",
    "SourceSpreadsheet", "SourceSheet", "TargetSheet",
    "RowsRead", "RowsWritten", "RowsCleared", "RangeWritten", "Status", "Message",
]


def maybe_write_log(
    service, config: Config, log_row: Dict[str, Any],
) -> None:
    """
    Пишет строку в Sync_Log в SUMMARY таблице.
    - При DRY_RUN — не пишет.
    - Если Sync_Log нет: при SAFE_MODE только warning, при CREATE_LOG_SHEET=true создаёт.
    """
    if config.dry_run:
        return

    summary_titles = get_sheet_titles(service, config.summary_spreadsheet_id)
    real = summary_titles.get("sync_log")

    if not real:
        if not config.create_log_sheet:
            if config.safe_mode:
                # warning один раз — не спамим
                return
        else:
            # создаём, только если CREATE_LOG_SHEET=true И не dry-run И не SAFE_MODE-блок
            if config.safe_mode and not config.allow_create_sheets:
                print("[WARN] Sync_Log missing and SAFE_MODE=true. Set ALLOW_CREATE_SHEETS=true "
                      "and CREATE_LOG_SHEET=true to create it.")
                return
            try:
                req = {"requests": [{"addSheet": {"properties": {"title": "Sync_Log"}}}]}
                service.spreadsheets().batchUpdate(
                    spreadsheetId=config.summary_spreadsheet_id, body=req
                ).execute()
                # заголовки
                write_values(
                    service, config.summary_spreadsheet_id,
                    "Sync_Log!A1:O1", [LOG_HEADERS],
                )
                real = "Sync_Log"
            except Exception as e:
                print(f"[WARN] cannot create Sync_Log: {e}")
                return

    # append через values.append
    row = [
        log_row.get("Timestamp", ""),
        log_row.get("Environment", ""),
        str(log_row.get("DryRun", "")),
        str(log_row.get("SafeMode", "")),
        log_row.get("Month", ""),
        log_row.get("Block", ""),
        log_row.get("SourceSpreadsheet", ""),
        log_row.get("SourceSheet", ""),
        log_row.get("TargetSheet", ""),
        log_row.get("RowsRead", 0),
        log_row.get("RowsWritten", 0),
        log_row.get("RowsCleared", 0),
        log_row.get("RangeWritten", ""),
        log_row.get("Status", ""),
        log_row.get("Message", ""),
    ]

    def _call():
        return service.spreadsheets().values().append(
            spreadsheetId=config.summary_spreadsheet_id,
            range="Sync_Log!A:O",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": [row]},
        ).execute()
    try:
        call_with_retry(_call, "append Sync_Log row")
    except Exception as e:
        print(f"[WARN] Sync_Log append failed: {e}")


# =====================================================================
#                  Запуск обновления одного месяца
# =====================================================================

def _ensure_summary_sheet(
    service, config: Config, summary_titles_lower: Dict[str, str], target_name: str,
) -> Optional[str]:
    """Возвращает реальное имя листа 'Сводная - {month}', или None если не разрешено создавать."""
    real = find_sheet_smart(summary_titles_lower, target_name)
    if real:
        return real
    # листа нет
    if config.safe_mode:
        print(f"[INFO] Target sheet {target_name!r} not found. "
              f"Waiting until humans create it.")
        return None
    if not config.allow_create_sheets:
        print(f"[INFO] Target sheet {target_name!r} not found and "
              f"ALLOW_CREATE_SHEETS=false. Waiting until humans create it.")
        return None
    if config.environment == "production" and not config.allow_production_write:
        print(f"[WARN] Refusing to create sheet in production without ALLOW_PRODUCTION_WRITE.")
        return None
    if config.dry_run:
        print(f"[DRY] would CREATE sheet {target_name!r}")
        return target_name  # для dry-run сделаем вид, что есть, но фактически не пишем
    # реально создаём
    try:
        req = {"requests": [{"addSheet": {"properties": {"title": target_name}}}]}
        service.spreadsheets().batchUpdate(
            spreadsheetId=config.summary_spreadsheet_id, body=req
        ).execute()
        print(f"[OK] created sheet {target_name!r}")
        return target_name
    except Exception as e:
        print(f"[ERROR] cannot create sheet {target_name!r}: {e}")
        return None


def run_month_update(
    service, config: Config,
    summary_titles_lower: Dict[str, str],
    our_titles_lower: Dict[str, str],
    yandex_titles_lower: Optional[Dict[str, str]],
    month_name: str, our_sheet: str, yandex_sheet: str,
) -> None:
    target_name = f"Сводная - {month_name}"
    print(f"\n--- month: {month_name} -> target sheet: {target_name!r}")

    real_title = _ensure_summary_sheet(service, config, summary_titles_lower, target_name)
    if not real_title:
        print(f"[SKIP] month {month_name}: target sheet unavailable")
        return

    # Локация блоков.
    # КРИТИЧНО: чтобы CLEAR_TAIL для НАША СЕТКА не зашёл в блок ЯНДЕКС
    # (и не затёр его title/header/первые строки), сначала находим title-row
    # ЯНДЕКС-блока и передаём его как next_block_title_row для OUR.
    # locate_block тогда сам обрежет data_max_rows OUR'a так, чтобы хвост
    # остановился до title ЯНДЕКС-блока.
    yandex_title_row_hint = find_title_row(
        service, config.summary_spreadsheet_id, real_title, "ЯНДЕКС СЕТКА"
    )
    if yandex_title_row_hint is None:
        print(f"[INFO] {target_name!r}: ЯНДЕКС СЕТКА title not found "
              f"(or blanked). OUR block will use unbounded block_height_hint.")

    our_loc = locate_block(
        service, config, real_title, "НАША СЕТКА",
        block_height_hint=config.our_block_rows,
        next_block_title_row=yandex_title_row_hint,
    )
    yandex_loc = locate_block(
        service, config, real_title, "ЯНДЕКС СЕТКА",
        block_height_hint=config.yandex_block_rows,
    )

    if not our_loc and not yandex_loc:
        print(f"[SKIP] {target_name!r}: no blocks found, will not write")
        return

    # --- Source: OUR ---
    our_new_rows = analyze_single_sheet(
        service, config.our_grid_id, our_titles_lower, our_sheet, config,
    )

    # --- Source: YANDEX (опционально) ---
    yandex_new_rows: List[List[Any]] = []
    yandex_available = True
    if yandex_titles_lower is None:
        yandex_available = False
        msg = "YANDEX source unavailable, skipped YANDEX block"
        print(f"[INFO] {msg}")
    elif not yandex_sheet:
        yandex_available = False
        print(f"[INFO] YANDEX sheet name is empty for month {month_name} -> skipping YANDEX block")
    else:
        try:
            yandex_new_rows = analyze_single_sheet(
                service, config.yandex_grid_id, yandex_titles_lower, yandex_sheet, config,
            )
        except SystemExit:
            raise
        except Exception as e:
            yandex_available = False
            if config.require_yandex:
                raise SystemExit(f"[FATAL] REQUIRE_YANDEX=true and yandex analyze failed: {e}")
            print(f"[WARN] YANDEX analyze failed, skipping YANDEX block: {e}")

    # --- OUR block: backup -> decide -> write ---
    if our_loc:
        _process_block(
            service=service, config=config, sheet_title=real_title,
            loc=our_loc, new_rows=our_new_rows,
            source_id=config.our_grid_id, source_sheet=our_sheet, month=month_name,
        )

    # --- YANDEX block ---
    if yandex_loc:
        if yandex_available:
            _process_block(
                service=service, config=config, sheet_title=real_title,
                loc=yandex_loc, new_rows=yandex_new_rows,
                source_id=config.yandex_grid_id, source_sheet=yandex_sheet, month=month_name,
            )
        else:
            # ВАЖНО: не трогаем старый яндекс-блок
            print(f"[INFO] {target_name}: keeping existing YANDEX block intact (source unavailable)")
            maybe_write_log(service, config, {
                "Timestamp": config.now_local().isoformat(),
                "Environment": config.environment,
                "DryRun": config.dry_run, "SafeMode": config.safe_mode,
                "Month": month_name, "Block": "ЯНДЕКС СЕТКА",
                "SourceSpreadsheet": _mask_id(config.yandex_grid_id),
                "SourceSheet": yandex_sheet, "TargetSheet": real_title,
                "RowsRead": 0, "RowsWritten": 0, "RowsCleared": 0,
                "RangeWritten": "", "Status": "skipped",
                "Message": "YANDEX source unavailable",
            })

    # --- timestamp в N1 ---
    if not config.dry_run:
        try:
            updated = config.now_local().strftime("%d.%m %H:%M:%S")
            write_values(
                service, config.summary_spreadsheet_id,
                f"{real_title}!{config.updated_at_cell}",
                [[f"Обновлено: {updated}"]],
            )
        except Exception as e:
            print(f"[WARN] cannot write timestamp to {config.updated_at_cell}: {e}")
    else:
        print(f"[DRY] would write timestamp to {real_title}!{config.updated_at_cell}")


def _process_block(
    *, service, config: Config, sheet_title: str, loc: BlockLocation,
    new_rows: List[List[Any]],
    source_id: str, source_sheet: str, month: str,
) -> None:
    """Backup + decide + apply для одного блока."""
    old_values = read_block_old_values(service, config, sheet_title, loc)

    backup_path: Optional[str] = None
    if config.backup_before_write and not config.dry_run:
        backup_path = backup_block_to_disk(config, sheet_title, loc, old_values)
        if backup_path:
            print(f"[OK] backup saved: {backup_path}")
        else:
            if config.safe_mode:
                print(f"[SKIP] backup failed and SAFE_MODE=true -> refusing to write {loc.label}")
                maybe_write_log(service, config, {
                    "Timestamp": config.now_local().isoformat(),
                    "Environment": config.environment,
                    "DryRun": config.dry_run, "SafeMode": config.safe_mode,
                    "Month": month, "Block": loc.label,
                    "SourceSpreadsheet": _mask_id(source_id),
                    "SourceSheet": source_sheet, "TargetSheet": sheet_title,
                    "RowsRead": len(new_rows), "RowsWritten": 0, "RowsCleared": 0,
                    "RangeWritten": "", "Status": "skipped",
                    "Message": "backup failed in SAFE_MODE",
                })
                return
            else:
                print(f"[WARN] backup failed but SAFE_MODE=false -> proceeding")

    decision = decide_write(config, sheet_title, loc, new_rows, old_values)

    if not decision.will_write:
        print(f"[SKIP] {loc.label}: {decision.skip_reason} "
              f"(old_filled={decision.old_filled}, new={decision.new_filled})")
        maybe_write_log(service, config, {
            "Timestamp": config.now_local().isoformat(),
            "Environment": config.environment,
            "DryRun": config.dry_run, "SafeMode": config.safe_mode,
            "Month": month, "Block": loc.label,
            "SourceSpreadsheet": _mask_id(source_id),
            "SourceSheet": source_sheet, "TargetSheet": sheet_title,
            "RowsRead": decision.new_filled, "RowsWritten": 0, "RowsCleared": 0,
            "RangeWritten": "", "Status": "skipped",
            "Message": decision.skip_reason,
        })
        return

    print(f"[PLAN] {loc.label}: write {decision.rows_to_write} rows -> {decision.range_write}"
          + (f"; clear {decision.rows_to_clear} rows -> {decision.range_clear}"
             if decision.will_clear else "; CLEAR_TAIL=false (tail untouched)"))

    apply_write_decision(service, config, sheet_title, loc, new_rows, decision)

    maybe_write_log(service, config, {
        "Timestamp": config.now_local().isoformat(),
        "Environment": config.environment,
        "DryRun": config.dry_run, "SafeMode": config.safe_mode,
        "Month": month, "Block": loc.label,
        "SourceSpreadsheet": _mask_id(source_id),
        "SourceSheet": source_sheet, "TargetSheet": sheet_title,
        "RowsRead": decision.new_filled,
        "RowsWritten": decision.rows_to_write,
        "RowsCleared": decision.rows_to_clear,
        "RangeWritten": decision.range_write,
        "Status": "ok" if not config.dry_run else "dry-run",
        "Message": "",
    })


# =====================================================================
#                Settings / orchestration
# =====================================================================

def read_settings_pairs(service, config: Config) -> List[Tuple[str, str, str]]:
    """Читает Settings!A2:B -> [(month, our_sheet, yandex_sheet), ...]"""
    rng = f"{config.summary_settings_sheet_name}!A2:B"
    rows = read_values(service, config.summary_spreadsheet_id, rng)
    pairs: List[Tuple[str, str, str]] = []
    for row in rows:
        a = (row[0] if len(row) > 0 else "").strip()
        b = (row[1] if len(row) > 1 else "").strip()
        month = a or b
        if not month:
            continue
        pairs.append((month, a, b))
    return pairs


def run_summary_once(config: Config) -> None:
    """
    Один проход по всем месяцам из Settings (или только по config.target_month).
    Подходит для --once и для одной итерации --loop.
    """
    print(config.banner())

    service = build_sheets_service(config)

    # ---- проверка title таблиц ----
    verify_spreadsheet_title(
        service, config.summary_spreadsheet_id,
        config.expected_summary_title, "SUMMARY",
    )
    verify_spreadsheet_title(
        service, config.our_grid_id,
        config.expected_our_title, "OUR",
    )

    yandex_titles_lower: Optional[Dict[str, str]] = None
    if config.yandex_grid_id:
        try:
            verify_spreadsheet_title(
                service, config.yandex_grid_id,
                config.expected_yandex_title, "YANDEX",
            )
            yandex_titles_lower = get_sheet_titles(service, config.yandex_grid_id)
        except SystemExit:
            raise
        except Exception as e:
            if config.require_yandex:
                raise SystemExit(f"[FATAL] REQUIRE_YANDEX=true and YANDEX read failed: {e}")
            print(f"[WARN] YANDEX source unavailable: {e}. Continuing without it.")
            yandex_titles_lower = None
    else:
        print("[INFO] YANDEX_GRID_ID is empty -> skip YANDEX entirely")

    summary_titles_lower = get_sheet_titles(service, config.summary_spreadsheet_id)
    our_titles_lower = get_sheet_titles(service, config.our_grid_id)

    pairs = read_settings_pairs(service, config)
    if not pairs:
        print(f"[WARN] Settings ({config.summary_settings_sheet_name}!A2:B) is empty")
        return

    months_in_settings = [p[0] for p in pairs]
    print(f"[INFO] Settings months: {months_in_settings}")

    # ---- выбор месяца(ев) для обработки ----
    # Приоритет:
    #   1. CLI --month (config.target_month)
    #   2. AUTO_MONTH=true -> compute_auto_month()
    #   3. иначе все месяцы из Settings
    resolved = config.resolve_target_month()
    selection_source = (
        "CLI --month" if config.target_month
        else "AUTO_MONTH" if config.auto_month
        else None
    )

    if resolved:
        wanted = resolved.strip().lower()
        matched = [p for p in pairs if p[0].strip().lower() == wanted]

        if not matched and config.auto_month and not config.target_month \
                and config.auto_discover_month:
            # Fallback: ищем самый свежий месяц из Settings, не позднее resolved
            fallback = find_closest_settings_month(pairs, resolved)
            if fallback:
                print(f"[INFO] AUTO_DISCOVER_MONTH: {resolved!r} not in Settings, "
                      f"falling back to {fallback!r}.")
                wanted = fallback.strip().lower()
                matched = [p for p in pairs if p[0].strip().lower() == wanted]
                resolved = fallback

        if not matched:
            # Стандартный лог из ТЗ — для AUTO_MONTH случая:
            if config.auto_month and not config.target_month:
                print(f"[INFO] AUTO_MONTH selected {resolved}, but it is not "
                      f"found in Settings. Waiting until humans add it.")
            else:
                print(f"[WARN] --month {resolved!r} not found in Settings")
            return

        pairs = matched
        print(f"[INFO] Selected month: {resolved!r} (source: {selection_source})")
    else:
        print(f"[INFO] No target month — processing ALL {len(pairs)} months from Settings")

    success = 0
    failed = 0
    for month, our_sheet, yandex_sheet in pairs:
        try:
            run_month_update(
                service, config,
                summary_titles_lower, our_titles_lower, yandex_titles_lower,
                month, our_sheet, yandex_sheet,
            )
            success += 1
        except SystemExit:
            raise
        except Exception as e:
            failed += 1
            print(f"[ERROR] month {month}: {type(e).__name__}: {e}")

    print(f"\n[SUMMARY] success={success}, failed={failed}, "
          f"total_months={len(pairs)}, dry_run={config.dry_run}")