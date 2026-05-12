"""
main.py — CLI и точка входа.

Использование:
  python main.py --once
  python main.py --loop
  python main.py --once --dry-run
  python main.py --once --month "Май 2026"
  python main.py --once --dry-run --month "Май 2026"

Если аргументов нет, берётся RUN_MODE из env (по умолчанию once).
"""

from __future__ import annotations

import argparse
import os
import sys
import time
import traceback

# .env подгружаем как можно раньше, до Config.from_env()
try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

from summary_sync import Config, run_summary_once


def parse_args(argv=None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="summary-sync",
        description=("Safe Google Sheets summary sync. "
                     "Reads OUR/YANDEX sources, writes per-manager stats "
                     "into the SUMMARY spreadsheet."),
    )
    mode = p.add_mutually_exclusive_group()
    mode.add_argument(
        "--once", dest="run_mode", action="store_const", const="once",
        help="run a single iteration and exit",
    )
    mode.add_argument(
        "--loop", dest="run_mode", action="store_const", const="loop",
        help="run continuously (sleep LOOP_SLEEP_SEC between iterations)",
    )
    p.add_argument(
        "--dry-run", action="store_true",
        help="force DRY_RUN=true (no writes anywhere)",
    )
    p.add_argument(
        "--month", default=None,
        help='process only this month from Settings (e.g. --month "Май 2026")',
    )
    return p.parse_args(argv)


def apply_cli_to_config(args: argparse.Namespace, config: Config) -> Config:
    if args.run_mode:
        config.run_mode = args.run_mode
    if args.dry_run:
        config.dry_run = True
    if args.month:
        config.target_month = args.month.strip()
    return config


def main(argv=None) -> int:
    args = parse_args(argv)

    config = Config.from_env()
    config = apply_cli_to_config(args, config)

    try:
        config.validate()
    except SystemExit as e:
        # выводим явно и выходим с кодом 2 (конфиг ошибка)
        print(f"[CONFIG ERROR] {e}", file=sys.stderr)
        return 2

    if config.run_mode == "once":
        try:
            run_summary_once(config)
            return 0
        except SystemExit as e:
            print(f"{e}", file=sys.stderr)
            return 1
        except Exception:
            traceback.print_exc()
            return 1

    # loop
    print(f"[INFO] entering loop mode, sleep={config.loop_sleep_sec}s between iterations")
    while True:
        started = time.time()
        try:
            run_summary_once(config)
        except SystemExit as e:
            # фатальная ошибка конфигурации/доступа — не имеет смысла спать и повторять
            print(f"[FATAL] {e}", file=sys.stderr)
            return 1
        except KeyboardInterrupt:
            print("[INFO] interrupted by user")
            return 0
        except Exception as e:
            traceback.print_exc()
            print(f"[ERROR] iteration failed: {type(e).__name__}: {e}")
        elapsed = time.time() - started
        sleep_s = max(1.0, config.loop_sleep_sec - elapsed)
        try:
            time.sleep(sleep_s)
        except KeyboardInterrupt:
            print("[INFO] interrupted by user")
            return 0


if __name__ == "__main__":
    sys.exit(main())