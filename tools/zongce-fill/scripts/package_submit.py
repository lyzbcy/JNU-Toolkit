# -*- coding: utf-8 -*-
"""
打包与提交: 生成 【学号】【姓名】/ 材料文件夹 + 佐证材料整理 + zip + 按配置提交。
用法:
  python package_submit.py --data 材料.json --workdir . --xlsx 已填表格.xlsx \
      [--evidence-dir 佐证图片所在目录] [--submit]
提交渠道读 assets/提交配置.json (班长发放前配置)。
"""
import argparse
import json
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

SKILL_DIR = Path(__file__).resolve().parent.parent
CONFIG_PATH = SKILL_DIR / "assets" / "提交配置.json"
STAR_MARKER = SKILL_DIR / ".star_prompted"
STAR_COOLDOWN = 15 * 24 * 3600   # 求好评冷却: 关闭一次后15天不再提


def _log(workdir, msg):
    try:
        with open(Path(workdir) / "综测填报日志.txt", "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now():%Y-%m-%d %H:%M}] {msg}\n")
    except Exception:
        pass


def star_prompt():
    """不打扰的求好评: 提交成功后提示一次, 15天内不重复。"""
    if STAR_MARKER.exists() and time.time() - STAR_MARKER.stat().st_mtime < STAR_COOLDOWN:
        return
    try:
        STAR_MARKER.write_text(datetime.now().isoformat(), encoding="utf-8")
    except Exception:
        pass
    print("\n  ── 一个小请求 ──")
    print("  这个助手是开源的: github.com/lyzbcy/JNU-Toolkit")
    print("  如果它帮到了你, 愿意点个 star 或在讨论区留句真实评价吗? 这对开发者很有帮助。")
    print("  (本次提示后15天内不会再打扰你)\n")

EVIDENCE_KIND = {"competitions": "竞赛", "papers": "论文", "dachuang": "大创",
                 "social_practices": "社会实践", "student_affairs": "学生事务",
                 "activities": "文体志愿", "certificates": "证书"}


def load_config():
    cfg = {"mode": "manual", "qq_url": "", "eform_url": "", "folder": "",
           "deadline": "", "note": "收集链接以群公告为准"}
    if CONFIG_PATH.exists():
        cfg.update(json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
    return cfg


def evidence_plan(data):
    """从 材料.json 的 evidence 字段生成 (标签, 文件) 列表; 没有则空。"""
    plan = []
    for ev in data.get("evidence") or []:
        label = ev.get("label") or "材料"
        for f in ev.get("files") or []:
            plan.append((label, Path(f)))
    return plan


def stage(data, workdir, xlsx, evidence_dir):
    sid, name = data["student"]["sid"], data["student"]["name"]
    folder = Path(workdir) / f"【{sid}】【{name}】"
    evdir = folder / f"【{sid}】【{name}】佐证材料"
    evdir.mkdir(parents=True, exist_ok=True)

    dst = folder / f"【{sid}】【{name}】本科生综测个人加分项信息统计表.xlsx"
    if xlsx:
        shutil.copy2(xlsx, dst)
        print(f"表格 → {dst}")

    copied = 0
    for label, src in evidence_plan(data):
        if not src.exists():
            print(f"  ! 佐证缺失: {label} <- {src}")
            continue
        ext = src.suffix or ".jpg"
        out = evdir / f"{label}{ext}"
        i = 2
        while out.exists():
            out = evdir / f"{label}-{i}{ext}"
            i += 1
        shutil.copy2(src, out)
        copied += 1
    if evidence_dir:
        for src in sorted(Path(evidence_dir).iterdir()):
            if src.is_file() and src.suffix.lower() in {".jpg", ".jpeg", ".png", ".pdf", ".webp", ".gif", ".bmp"}:
                out = evdir / src.name
                if not out.exists():
                    shutil.copy2(src, out)
                    copied += 1
    print(f"佐证 {copied} 件 → {evdir}")
    return folder


def make_zip(folder: Path) -> Path:
    # eform 收集表要求上传文件名为【学号】【姓名】综测材料.zip
    zip_name = folder.name + "综测材料.zip"
    zip_path = folder.parent / zip_name
    if zip_path.exists():
        zip_path.unlink()
    shutil.make_archive(str(folder.parent / (folder.name + "综测材料")), "zip", folder.parent, folder.name)
    zp = folder.parent / zip_name
    print(f"压缩包 → {zp} ({zp.stat().st_size} bytes)")
    return zp


def submit(zip_path: Path, cfg):
    mode = cfg.get("mode", "manual")
    print(f"\n提交渠道: {mode}")
    if mode == "qq" and cfg.get("qq_url"):
        print(f"  打开QQ收集表: {cfg.get('qq_url')}")
        print("  请上传压缩包并按表单要求填写。")
        _open(cfg["qq_url"])
    elif mode == "eform" and cfg.get("eform_url"):
        print("  实测: 公网可直连打开; 需登录e江南(本人扫码一次, 当次会话保持)。")
        print(f"  打开江大智能填报: {cfg.get('eform_url')}")
        _open(cfg["eform_url"])
    elif mode == "folder" and cfg.get("folder"):
        target = Path(cfg["folder"])
        target.mkdir(parents=True, exist_ok=True)
        shutil.copy2(zip_path, target / zip_path.name)
        print(f"  已复制到收集文件夹: {target / zip_path.name}")
    else:
        desktop = Path.home() / "Desktop"
        if desktop.exists():
            local = desktop / zip_path.name
            shutil.copy2(zip_path, local)
            print(f"  压缩包已放到桌面: {local}")
        print("  手动提交: 把压缩包发到班长指定的收集渠道(QQ收集表/eform/群文件)。")
    if cfg.get("deadline"):
        print(f"  截止时间: {cfg['deadline']}")
    if cfg.get("note"):
        print(f"  备注: {cfg['note']}")
    print("  请保留本地文件夹作为备份, 公示期有异议时以佐证材料为准。")
    _log(Path.cwd(), f"已提交 zip={zip_path.name} 渠道={mode}")
    star_prompt()


def _open(url):
    try:
        subprocess.Popen(["cmd", "/c", "start", "", url], shell=False)
    except Exception:
        print(f"  (自动打开浏览器失败, 请手动访问: {url})")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True)
    ap.add_argument("--workdir", default=".")
    ap.add_argument("--xlsx", help="已填好的表格路径")
    ap.add_argument("--evidence-dir", help="佐证文件所在目录(全部图片/PDF直接复制)")
    ap.add_argument("--stage-evidence", action="store_true", help="只整理材料不提交")
    ap.add_argument("--submit", action="store_true", help="打包并提交")
    args = ap.parse_args()

    data = json.loads(Path(args.data).read_text(encoding="utf-8"))
    folder = stage(data, args.workdir, args.xlsx, args.evidence_dir)
    _log(args.workdir, f"材料文件夹就绪 {folder.name} (xlsx={'有' if args.xlsx else '缺'}, 佐证目录={args.evidence_dir or '未指定'})")
    if not (args.stage_evidence or args.submit):
        print(f"\n材料文件夹就绪: {folder} (未提交; 加 --submit 打包提交)")
        return
    zp = make_zip(folder)
    if args.submit:
        submit(zp, load_config())


if __name__ == "__main__":
    main()
