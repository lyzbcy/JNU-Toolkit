# -*- coding: utf-8 -*-
"""
静默更新(每天首次使用时由 SKILL.md 触发):
  1. 比对本地 VERSION 与远端 raw VERSION
  2. 不一致 -> 下载仓库 zip, 原子替换本 skill 目录(排除本地配置)
  3. 全程静默: 失败则继续用旧版, 绝不打断本次使用
本地配置(提交配置.json)在更新中被保留, 不被远端覆盖。
"""
import os
import shutil
import sys
import tempfile
import urllib.request
import zipfile
from datetime import date
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

SKILL_DIR = Path(__file__).resolve().parent.parent
REPO = "lyzbcy/JNU-Toolkit"
BRANCH = "main"
REMOTE_VERSION_URL = f"https://raw.githubusercontent.com/{REPO}/{BRANCH}/tools/zongce-fill/VERSION"
REMOTE_ZIP_URL = f"https://github.com/{REPO}/archive/refs/heads/{BRANCH}.zip"
STAMP = SKILL_DIR / ".last_update_check"
KEEP_LOCAL = ["提交配置.json"]          # 更新时保留的本地文件
STALE_DAYS = 1


def _get(url, timeout=10):
    req = urllib.request.Request(url, headers={"User-Agent": "zongce-fill-updater"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def checked_today():
    if not STAMP.exists():
        return False
    try:
        return STAMP.read_text(encoding="utf-8").strip() == date.today().isoformat()
    except Exception:
        return False


def local_version():
    v = SKILL_DIR / "VERSION"
    return v.read_text(encoding="utf-8").strip() if v.exists() else "0.0.0"


def main():
    if checked_today():
        return
    try:
        STAMP.write_text(date.today().isoformat(), encoding="utf-8")
    except Exception:
        pass
    try:
        remote = _get(REMOTE_VERSION_URL).decode("utf-8").strip()
    except Exception as e:
        print(f"[更新检查跳过: 无法访问远端 ({type(e).__name__}), 继续使用本地版本 {local_version()}]")
        return
    if remote == local_version():
        return
    print(f"[检测到新版本 {local_version()} -> {remote}, 正在静默更新...]")
    try:
        with tempfile.TemporaryDirectory() as td:
            zpath = Path(td) / "repo.zip"
            zpath.write_bytes(_get(REMOTE_ZIP_URL, timeout=60))
            with zipfile.ZipFile(zpath) as z:
                z.extractall(td)
            src = Path(td) / f"{REPO.split('/')[1]}-{BRANCH}" / "tools" / "zongce-fill"
            if not (src / "SKILL.md").exists():
                raise FileNotFoundError("远端包结构异常")
            backup = Path(td) / "backup"
            shutil.copytree(SKILL_DIR, backup)
            for item in src.iterdir():
                if item.name in KEEP_LOCAL:
                    continue
                dst = SKILL_DIR / item.name
                if dst.is_dir():
                    shutil.rmtree(dst, ignore_errors=True)
                elif dst.exists():
                    dst.unlink()
                if item.is_dir():
                    shutil.copytree(item, dst)
                else:
                    shutil.copy2(item, dst)
        print(f"[已更新到 {remote}, 本次任务继续使用新版本]")
    except Exception as e:
        print(f"[静默更新失败({type(e).__name__}: {e}), 继续使用旧版本 {local_version()}, 下次自动重试]")


if __name__ == "__main__":
    main()
