"""
worker.py — тонкая безопасная обёртка для совместимости со старыми деплоями.

Раньше worker.py содержал свой собственный бесконечный цикл, что приводило
к двум независимым процессам, которые могли писать в одну и ту же таблицу.

Теперь worker.py НЕ имеет своей логики. Он просто вызывает main(--loop),
поэтому все проверки безопасности (test/production, DRY_RUN, SAFE_MODE,
backup, drop ratio) применяются единообразно.

Если на Railway раньше был Start Command `python worker.py`, его можно
не менять — поведение будет идентично `python main.py --loop`.
Лучше всё-таки сменить на `python main.py --loop`, чтобы был CLI-контроль.
"""

from __future__ import annotations

import sys

from main import main


if __name__ == "__main__":
    # форсируем loop-режим; всё остальное берётся из env
    sys.exit(main(["--loop"]))