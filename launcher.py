from __future__ import annotations

import os
import json
import subprocess
import sys
import tempfile
import threading
import urllib.request
from pathlib import Path

from template_export import export_word


APP_TITLE = "Kobe強強照片黏貼"
HTML_NAME = "photo35.html"
REPO_API = "https://api.github.com/repos/waynelord0628-beep/kobe_photo_paste2.0/releases/latest"


class ProjectApi:
    def __init__(self) -> None:
        self.window = None

    def save_project(self, content: str) -> dict:
        try:
            import webview

            if self.window is None:
                return {"ok": False, "error": "視窗尚未初始化"}

            selected = self.window.create_file_dialog(
                webview.SAVE_DIALOG,
                save_filename="照片黏貼紀錄表專案.json",
                file_types=("JSON 檔案 (*.json)", "所有檔案 (*.*)"),
            )
            if not selected:
                return {"ok": False, "cancelled": True}

            path = selected if isinstance(selected, str) else selected[0]
            target = Path(path)
            if target.suffix.lower() != ".json":
                target = target.with_suffix(".json")
            target.write_text(content, encoding="utf-8")
            return {"ok": True, "path": str(target)}
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

    def export_word_template(self, payload: str) -> dict:
        try:
            import webview

            if self.window is None:
                return {"ok": False, "error": "視窗尚未準備完成"}

            data = json.loads(payload)
            title = data.get("title") or "Kobe強強照片黏貼"
            selected = self.window.create_file_dialog(
                webview.SAVE_DIALOG,
                save_filename=f"{title}.docx",
                file_types=("Word 文件 (*.docx)", "所有檔案 (*.*)"),
            )
            if not selected:
                return {"ok": False, "cancelled": True}

            path = selected if isinstance(selected, str) else selected[0]
            target = Path(path)
            if target.suffix.lower() != ".docx":
                target = target.with_suffix(".docx")

            generated = export_word(
                template_dir=resource_path("templates"),
                output_dir=target.parent,
                mode=str(data.get("mode") or ""),
                title=target.stem,
                photos=list(data.get("photos") or []),
            )
            if generated != target:
                if target.exists():
                    target.unlink()
                generated.rename(target)
            return {"ok": True, "path": str(target)}
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

    def check_update(self, current_version: str) -> dict:
        try:
            release = fetch_json(REPO_API)
            latest = release.get("tag_name", "")
            asset = pick_windows_asset(release.get("assets", []))
            if not latest or not asset:
                return {"ok": False, "error": "找不到可用的 Windows 版本"}
            return {
                "ok": True,
                "update": version_tuple(latest) > version_tuple(current_version),
                "version": latest,
                "download_url": asset["browser_download_url"],
                "asset_name": asset.get("label") or asset.get("name", ""),
            }
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

    def install_update(self, download_url: str) -> dict:
        try:
            if not getattr(sys, "frozen", False):
                return {"ok": False, "error": "開發模式不能自動替換 exe，請先打包後再測試更新。"}

            target = Path(sys.executable)
            tmp_exe = Path(tempfile.gettempdir()) / "Kobe強強照片黏貼_update.exe"
            req = urllib.request.Request(download_url, headers={"User-Agent": "KobePhotoPasteUpdater"})
            with urllib.request.urlopen(req, timeout=120) as response:
                tmp_exe.write_bytes(response.read())

            if tmp_exe.stat().st_size < 1024 * 1024:
                return {"ok": False, "error": "下載檔案異常，已取消更新。"}

            bat = Path(tempfile.gettempdir()) / "Kobe強強照片黏貼_update.cmd"
            bat.write_text(
                "\n".join([
                    "@echo off",
                    "chcp 65001 >nul",
                    f'set "SRC={tmp_exe}"',
                    f'set "DST={target}"',
                    ":wait",
                    "timeout /t 1 /nobreak >nul",
                    'copy /Y "%SRC%" "%DST%" >nul',
                    "if errorlevel 1 goto wait",
                    'start "" "%DST%"',
                    'del "%SRC%" >nul 2>nul',
                    'del "%~f0" >nul 2>nul',
                ]),
                encoding="utf-8",
            )
            subprocess.Popen(["cmd", "/c", str(bat)], creationflags=subprocess.CREATE_NO_WINDOW)
            if self.window is not None:
                threading.Timer(0.3, self.window.destroy).start()
            return {"ok": True}
        except Exception as exc:
            return {"ok": False, "error": str(exc)}


def fetch_json(url: str) -> dict:
    req = urllib.request.Request(url, headers={"User-Agent": "KobePhotoPasteUpdater"})
    with urllib.request.urlopen(req, timeout=30) as response:
        return json.loads(response.read().decode("utf-8"))


def version_tuple(version: str) -> tuple[int, ...]:
    parts = version.strip().lower().lstrip("v").split(".")
    nums = []
    for part in parts:
        digits = "".join(ch for ch in part if ch.isdigit())
        nums.append(int(digits or 0))
    return tuple(nums)


def pick_windows_asset(assets: list[dict]) -> dict | None:
    exe_assets = [asset for asset in assets if asset.get("name", "").lower().endswith(".exe")]
    if not exe_assets:
        return None
    for asset in exe_assets:
        name = (asset.get("label") or asset.get("name", "")).lower()
        if "kobe" in name or "照片" in name:
            return asset
    return exe_assets[0]


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

    api = ProjectApi()
    window = webview.create_window(
        APP_TITLE,
        html=src.read_text(encoding="utf-8"),
        js_api=api,
        width=1280,
        height=860,
        min_size=(980, 680),
        text_select=True,
    )
    api.window = window
    webview.start()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
