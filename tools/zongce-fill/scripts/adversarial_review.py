# -*- coding: utf-8 -*-
"""
对抗性评审: 假装自己是来挑刺的复查者, 对已填表格+材料.json 发起结构化攻击。
设计思想来自"再派一个子Agent监测每一步操作"的回归测试招法——这里用确定性规则
把攻击固化成脚本, 任何运行环境都能跑。

攻击向量:
  A 结构攻击   模板被改动?(隐藏表/表头/下拉校验与官方模板比对)
  B 篡改攻击   已填文件 vs 由材料.json 重新生成的期望文件 逐格比对, 分数虚报=ERROR
  C 分数攻击   重算每项分数, written > computed = 虚报(ERROR), < = 漏分(WARN)
  D 佐证攻击   每个计分项都要有对应佐证, 缺佐证的项禁止进入提交
  E 学年攻击   时间字段超出 2025-2026 学年 = ERROR
  F 重复攻击   同一项目/同一竞赛重复申报, 违反就高原则
  G 可疑模式   学时恰好40无截图/清一色一等奖/边界分数/缺队内顺位
用法:
  python adversarial_review.py --data 材料.json --file 已填.xlsx --template 模板.xlsx [--evidence-dir 佐证目录]
退出码: 0=评审通过(可有提醒) / 2=存在ERROR, 禁止提交
"""
import argparse
import json
import re
import sys
import tempfile
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

sys.path.insert(0, str(Path(__file__).resolve().parent))
from fill_form import (Calculator, CERT_NAMES, norm_text, load_workbook, write_workbook)

YEAR_LO, YEAR_HI = "2025-09", "2026-08"   # 2025-2026 学年窗口(粗口径)


class Review:
    def __init__(self):
        self.vectors = {}   # vector -> (defended_bool, [findings])

    def hit(self, vec, level, msg):
        """level: ERROR / WARN / INFO"""
        defended = level != "ERROR"
        self.vectors.setdefault(vec, [True, []])
        if not defended:
            self.vectors[vec][0] = False
        self.vectors[vec][1].append(f"{level} {msg}")

    def report(self):
        print("\n" + "=" * 62)
        print("对抗性评审报告")
        print("=" * 62)
        n_err = n_warn = n_info = 0
        for vec, (defended, finds) in self.vectors.items():
            status = "◎ 防守成功" if defended else "× 攻击得手"
            print(f"\n[{vec}] {status}")
            for f in finds:
                print(f"   {f}")
                if f.startswith("ERROR"): n_err += 1
                elif f.startswith("WARN"): n_warn += 1
                else: n_info += 1
        print(f"\n合计: ERROR {n_err} / WARN {n_warn} / INFO {n_info}")
        if n_err:
            print("结论: 评审不通过 —— 存在虚报/缺佐证/超学年等硬伤, 禁止提交。")
            print("      修复 材料.json 后重新生成, 不要手改 xlsx。")
        else:
            print("结论: 评审通过。WARN/INFO 项请向同学核实后放行。")
        return 2 if n_err else 0


def _cells(ws, r1, r2, c1, c2):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None:
                yield f"{ws.cell(row=r, column=c).coordinate}", v


def vector_structure(review, tpl_path, stu_path):
    tpl, stu = load_workbook(tpl_path), load_workbook(stu_path)
    if tpl.sheetnames != stu.sheetnames:
        review.hit("A 结构攻击", "ERROR", f"sheet 列表被改动: {stu.sheetnames}")
        return
    if stu["Sheet1"].sheet_state != "hidden":
        review.hit("A 结构攻击", "ERROR", "隐藏竞赛目录 Sheet1 被取消隐藏")
    t1, s1 = tpl["Sheet1"], stu["Sheet1"]
    diff = 0
    for r in range(1, max(t1.max_row, s1.max_row) + 1):
        tv, sv = t1.cell(row=r, column=1).value, s1.cell(row=r, column=1).value
        bv, cv = t1.cell(row=r, column=2).value, s1.cell(row=r, column=2).value
        if (tv, bv) != (sv, cv):
            diff += 1
            if diff <= 3:
                review.hit("A 结构攻击", "ERROR", f"竞赛目录第{r}行被改动: {sv}/{cv} != {tv}/{bv}")
    if diff > 3:
        review.hit("A 结构攻击", "ERROR", f"竞赛目录共{diff}行被改动")
    for name in tpl.sheetnames:
        tdv = len(tpl[name].data_validations.dataValidation)
        sdv = len(stu[name].data_validations.dataValidation)
        if tdv != sdv:
            review.hit("A 结构攻击", "WARN", f"{name} 下拉校验组数 {tdv}->{sdv}, 可能破坏了模板")
        for r in range(1, 3):
            for c in range(1, 10):
                tv, sv = tpl[name].cell(row=r, column=c).value, stu[name].cell(row=r, column=c).value
                if tv != sv and str(tv or "").strip() and str(sv or "").strip() not in ("", "None"):
                    review.hit("A 结构攻击", "WARN", f"{name}!{tpl[name].cell(row=r, column=c).coordinate} 表头被改: {sv!r} != {tv!r}")
    if not any(v[0] is False for v in review.vectors.values()) and "A 结构攻击" not in review.vectors:
        review.hit("A 结构攻击", "INFO", "模板结构完整(目录/表头/校验未被改动)")


def vector_tamper(review, data, tpl_path, stu_path):
    """用 材料.json 在干净模板上重新生成期望结果, 与实际文件逐格比对。"""
    tpl = load_workbook(tpl_path)
    calc = _recalc(data, tpl)
    if calc is None:
        return
    with tempfile.TemporaryDirectory() as td:
        expected_path = Path(td) / "expected.xlsx"
        write_workbook(load_workbook(tpl_path), calc, data, expected_path)
        exp, stu = load_workbook(expected_path), load_workbook(stu_path)
        zones = [("基础分", 4, 12, 1, 6), ("加分项-竞赛", 3, 40, 1, 9),
                 ("加分项-论文著作专利", 3, 40, 1, 7), ("加分项-社会实践", 3, 60, 1, 5),
                 ("加分项-学生事务", 3, 40, 1, 4), ("加分项-文体志愿类", 4, 40, 1, 7),
                 ("加分项-其他", 4, 40, 1, 9)]
        diff = 0
        for sn, r1, r2, c1, c2 in zones:
            for coord, v in _cells(stu[sn], r1, r2, c1, c2):
                ev = exp[sn][coord].value
                if ev != v:
                    diff += 1
                    col = re.sub(r"\d+", "", coord)
                    level = "ERROR" if col in ("I", "G", "E", "D") and _num(v) > _num(ev) else "WARN"
                    review.hit("B 篡改攻击", level,
                               f"{sn}!{coord}: 文件值{v!r} 与生成期望{ev!r} 不符" +
                               ("(分数高于应得=虚报)" if level == "ERROR" else "(疑似手工修改)"))
        if diff == 0:
            review.hit("B 篡改攻击", "INFO", "与AI生成结果逐格一致, 无手工篡改痕迹")


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return -1e9


def _recalc(data, tpl_wb):
    ws1 = tpl_wb["Sheet1"]
    catalog = {}
    for row in ws1.iter_rows(min_row=1, max_row=ws1.max_row, max_col=2):
        if row[0].value and row[1].value:
            catalog[str(row[0].value)] = str(row[1].value)
    calc = Calculator(catalog)
    for fn in (calc.calc_basic, calc.calc_competitions, calc.calc_papers, calc.calc_practices,
               calc.calc_dachuang, calc.calc_affairs, calc.calc_activities, calc.calc_others):
        fn(data)
    return calc


def _match_norm(s):
    """匹配用归一化: 去空白/全半角差异/标点(连字符、括号等)。"""
    return re.sub(r"[^\w\u4e00-\u9fff]", "", norm_text(s))


def _common_sub(a, b, minlen=4):
    a, b = _match_norm(a), _match_norm(b)
    if len(a) < minlen or len(b) < minlen:
        return a == b and a != ""
    # 最长公共子串(短串足够小)
    best = 0
    for i in range(len(a)):
        for j in range(len(b)):
            k = 0
            while i + k < len(a) and j + k < len(b) and a[i + k] == b[j + k]:
                k += 1
            best = max(best, k)
    return best >= minlen


def vector_evidence(review, data, evidence_dir):
    pool = [ev.get("label", "") for ev in data.get("evidence") or []]
    if evidence_dir:
        pool += [p.stem for p in Path(evidence_dir).iterdir() if p.is_file()]
    pool = [p for p in pool if p]
    items = []
    for c in data.get("competitions") or []:
        items.append(("竞赛", c.get("name", "")))
    for p in data.get("papers") or []:
        items.append(("论文", p.get("title", "")))
    for x in data.get("dachuang") or []:
        if x.get("passed"):
            items.append(("大创", x.get("title", "")))
    for s in data.get("social_practices") or []:
        items.append(("社会实践", s.get("name", "")))
    for a in (data.get("student_affairs") or []):
        items.append(("学生事务", a.get("position", "")))
    for a in data.get("activities") or []:
        items.append(("文体志愿", a.get("name", "")))
    for c in data.get("certificates") or []:
        items.append(("证书", c.get("name", "")))
    if not items:
        review.hit("C 佐证攻击", "INFO", "无计分项, 无需佐证")
        return
    missing = []
    waivers = data.get("evidence_waivers") or {}
    for kind, name in items:
        if any(_common_sub(name, p) for p in pool):
            continue
        key = f"{kind}-{name}"
        if key in waivers:
            review.hit("C 佐证攻击", "WARN", f"[{key}] 缺个人佐证, 已按核定豁免: {waivers[key]}")
            continue
        missing.append(key)
    if missing:
        for m in missing:
            review.hit("C 佐证攻击", "ERROR", f"[{m}] 找不到对应佐证(证据池{len(pool)}件)——补佐证或删行")
    else:
        review.hit("C 佐证攻击", "INFO", f"{len(items)}个计分项全部有佐证对应(证据池{len(pool)}件)")


def _extract_year(s):
    ys = re.findall(r"20\d{2}", str(s or ""))
    return [int(y) for y in ys]


def vector_year(review, data):
    checks = []
    for c in data.get("competitions") or []:
        checks.append(("竞赛", c.get("name"), c.get("time", "")))
    for c in (data.get("basic") or {}).get("competitions_participated") or []:
        checks.append(("科创基础", c.get("name"), c.get("time", "")))
    for a in data.get("activities") or []:
        checks.append((a.get("type"), a.get("name"), a.get("time", "")))
    for c in data.get("certificates") or []:
        checks.append(("证书", c.get("name"), c.get("date", "")))
    for a in data.get("student_affairs") or []:
        checks.append(("学生事务", a.get("position"), a.get("period", "")))
    for v in data.get("overseas") or []:
        checks.append(("海外", v.get("school"), v.get("period", "")))
    bad, unknown = [], []
    for kind, name, t in checks:
        ys = _extract_year(t)
        if not ys:
            unknown.append(f"{kind}-{name}")
        elif any(y < 2025 or y > 2026 for y in ys):
            bad.append(f"{kind}-{name}({t})")
    for b in bad:
        review.hit("E 学年攻击", "ERROR", f"{b} 超出2025-2026学年窗口")
    if unknown:
        review.hit("E 学年攻击", "WARN", f"{len(unknown)}项时间无法自动判定: " + "; ".join(unknown[:5]))
    if not bad and not unknown:
        review.hit("E 学年攻击", "INFO", "全部时间字段落在2025-2026学年")


def vector_dup(review, data):
    names = [norm_text(c.get("name", "")) for c in data.get("competitions") or []]
    dups = {n for n in names if names.count(n) > 1}
    for d in dups:
        review.hit("F 重复攻击", "ERROR", f"竞赛[{d}]申报了多次——同一项目只按最高值计一次, 请合并")
    titles = [norm_text(p.get("title", "")) for p in data.get("papers") or []]
    tdup = {t for t in titles if titles.count(t) > 1}
    for t in tdup:
        review.hit("F 重复攻击", "ERROR", f"论文[{t}]重复申报")
    acts = [(norm_text(a.get("name", "")), a.get("award_level"), a.get("award_grade"))
            for a in data.get("activities") or []]
    adup = {a for a in acts if acts.count(a) > 1}
    for a in adup:
        review.hit("F 重复攻击", "WARN", f"活动[{a[0]}]同级别重复, 确认是否同一项")
    certs = {}
    for c in data.get("certificates") or []:
        k = c.get("name")
        certs.setdefault(k, []).append(c.get("score_or_grade"))
    for k, vs in certs.items():
        if len(vs) > 1:
            review.hit("F 重复攻击", "WARN", f"{k} 考取{len(vs)}次({vs}), 同证书按最高一次计")
    if not any(not v[0] for v in review.vectors.values()) and "F 重复攻击" not in review.vectors:
        review.hit("F 重复攻击", "INFO", "无重复申报")


def vector_suspicious(review, data):
    hours = (data.get("basic") or {}).get("second_classroom_hours")
    if hours in (40, "40"):
        review.hit("G 可疑模式", "WARN", "学时恰好=40(满分线), 请让同学出示第二课堂学时截图")
    comps = data.get("competitions") or []
    if len(comps) >= 2 and all(c.get("award_grade") == "一等奖" for c in comps):
        review.hit("G 可疑模式", "INFO", "多项竞赛均为一等奖——逐项核对证书原件后再放行")
    for c in comps:
        if c.get("team_size", 1) and int(c.get("team_size", 1)) > 1 and not c.get("team_rank"):
            review.hit("G 可疑模式", "WARN", f"[{c.get('name')}] 团队{c.get('team_size')}人但缺队内顺位, 分摊系数无法计算")
    for c in data.get("certificates") or []:
        try:
            s = float(c.get("score_or_grade"))
        except (TypeError, ValueError):
            continue
        if c.get("name") in ("CET-4", "CET-6") and s in (425, 500, 560):
            review.hit("G 可疑模式", "INFO", f"{c.get('name')}={s} 恰在档位边界, 已按就高档计, 请核对成绩单")
    if not any(not v[0] for v in review.vectors.values()) and "G 可疑模式" not in review.vectors:
        review.hit("G 可疑模式", "INFO", "无可疑模式")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True)
    ap.add_argument("--file", required=True)
    ap.add_argument("--template", required=True)
    ap.add_argument("--evidence-dir")
    args = ap.parse_args()

    data = json.loads(Path(args.data).read_text(encoding="utf-8"))
    review = Review()
    vector_structure(review, args.template, args.file)
    vector_tamper(review, data, args.template, args.file)
    vector_evidence(review, data, args.evidence_dir)
    vector_year(review, data)
    vector_dup(review, data)
    vector_suspicious(review, data)
    sys.exit(review.report())


if __name__ == "__main__":
    main()
