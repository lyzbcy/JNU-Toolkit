---
name: lyzbcy-cg
description: 通过 Edge 已登录会话检查江南大学 CG 在线作业、截止时间和个人提交状态。
metadata:
  version: "1.0.1"
---

# CG 在线作业

用已经连接 Edge 的 kimi-webbridge。请在 Edge 登录 WebVPN 和 CG，再从「选择课程」复制课程列表地址。不要在命令、截图或公开材料里保存密码、Cookie 或令牌。

```powershell
python scripts/cg.py --url "<课程列表地址>" login-status
python scripts/cg.py --url "<课程列表地址>" assignments
```

也可由本机环境变量 `CG_WORKSPACE_URL` 提供地址，或在当前交接目录的 `agents.md` 写入课程列表地址。脚本只读取用户本人的课程和提交记录。搜索全部课程时会顺次切换站点记住的当前课程；运行中不要用另一标签切换课程。如果扫描报错，单独读取对应页，空白页面不能证明没有作业。

通过「在线作业 → 当前作业」复查有效期、多个题目各自的状态及文件要求。列出存在未完成题目的作业，由用户确定范围。实验报告可配合 [experiment-report-skill](https://github.com/lyzbcy/experiment-report-skill) 和课程提供的模板。文件上传使用平台正式表单，经用户授权后提交。等待超时先回看该题的下载入口与最近提交时间。CG 以服务端回执为准，核对前不要重复上传。

脚本每日本地首次运行后在后台检查同仓库公开发行版。下载包必须与 GitHub Release 的 SHA-256 和文件白名单相符，更新在下次运行生效。失败时继续使用现有版本。
