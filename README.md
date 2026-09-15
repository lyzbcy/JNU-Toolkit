# JNU实用工具包（JNU-Toolkit）

江南大学同学的 AI 实用工具集合。第一个工具：**综测个人加分项自动填报助手（zongce-fill）**——把每年开学最痛苦的综测填表，从一下午压缩到十分钟。

## 🧰 工具列表

| 工具 | 用途 | 状态 |
|---|---|---|
| [zongce-fill](tools/zongce-fill/) | 人工智能与计算机学院综测《个人加分项信息统计表》自动填报：访谈 → 自动算分 → 生成表格 → 校验 → 对抗性评审 → 经本人同意后打包提交 | ✅ v1.1.0 可用 |

## 🚀 快速开始（以 zongce-fill 为例）

### 前置条件
- 任一支持 `.agents/skills/` 目录的 AI 命令行工具（ZCode、Claude Code 兼容 CLI 等）
- Python 3.9+，`pip install openpyxl`

### 安装

```bash
git clone https://github.com/lyzbcy/JNU-Toolkit.git
# 把工具复制进 AI CLI 的技能目录
xcopy /e /i /y "JNU-Toolkit\tools\zongce-fill" "%USERPROFILE%\.agents\skills\zongce-fill"   (Windows)
cp -r JNU-Toolkit/tools/zongce-fill ~/.agents/skills/                                         (macOS/Linux)
```

### 使用
1. 新建文件夹，放入学院下发的《本科生综测个人加分项信息统计表》空白模板
2. 在该文件夹启动你的 AI CLI
3. 对它说：**"帮我填综测表"**，然后把获奖证书/成绩单截图直接丢给它

## 🛡️ 可靠性设计

这个工具替你填报，也替你把关。每份产出在提交前要过四道关：

1. **脚本算分**——分值表直接内置自《学院综测实施细则（2025年7月修订）》，AI 不心算
2. **格式校验**——下拉选项、必填项、表格结构逐格核对，不破坏学院模板
3. **对抗性评审**——独立复查视角发起 7 类攻击：模板篡改 / 生成后手工改分 / 分数虚报 / 缺佐证 / 超学年 / 重复申报 / 可疑模式。任一攻击得手即禁止提交
4. **双重同意**——写入表格前、提交前各需本人明确确认；缺佐证的项要么补材料要么删行

另有：静默自动更新（每天首次使用自动对齐远端版本）、填报日志留痕、15 天冷却的单次 star 提示。

## 📐 目录结构

```
JNU-Toolkit/
├── tools/               ← 各工具（每个是一个可独立安装的 skill）
│   └── zongce-fill/
│       ├── SKILL.md     ← 工作流主文档
│       ├── VERSION      ← 版本号（静默更新锚点）
│       ├── references/  ← 评分规则、表格结构规格
│       ├── scripts/     ← fill_form / validate / adversarial_review / package_submit / silent_update
│       └── assets/      ← 空白模板、材料示例、提交配置
├── guides/              ← 面向人的填报指南（HTML）
└── docs/                ← roadmap 与开发记录
```

## ⚠️ 说明

- 分值规则以学院当年正式文件为准；细则未明确的口径（如部分竞赛三等奖档位、论文分区）一律标记"待班长/辅导员核定"，本工具不擅自拍板。
- 空白模板来自学院下发的通知材料包，每年请以最新官方版本为准替换 `assets/` 中的模板。
- 本工具只做信息整理与算分，**不替你编造任何材料**——弄虚作假会取消当学年评奖评优资格，不值得。

## 📄 License

MIT © [lyzbcy](https://lyzbcy.github.io/)

## 关于作者

🐟 **捞鱼** —— 一个弱小但有梦想的开发者
个人主页：[lyzbcy.github.io](https://lyzbcy.github.io/)

如果这个工具帮到了你，欢迎点一个 ⭐ star，或在 [Discussions](https://github.com/lyzbcy/JNU-Toolkit/discussions) 留下你的真实评价——这对我很有帮助。
