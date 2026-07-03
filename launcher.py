from __future__ import annotations

import os
import sys
from pathlib import Path


APP_TITLE = "Kobe強強照片黏貼"
HTML_NAME = "photo35.html"


def resource_path(name: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    direct = base / name
    if direct.exists():
        return direct
    return base / "src" / name


def main() -> int:
    src = resource_path(HTML_NAME)
    if not src.exists():
        os.system(f'msg * "{APP_TITLE} 啟動失敗：找不到 {HTML_NAME}"')
        return 1

    try:
        import webview
    except Exception as exc:
        os.system(f'msg * "{APP_TITLE} 啟動失敗：缺少本機視窗元件 {exc}"')
        return 1

    webview.create_window(
        APP_TITLE,
        html=src.read_text(encoding="utf-8"),
        width=1280,
        height=860,
        min_size=(980, 680),
        text_select=True,
    )
    webview.start()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
