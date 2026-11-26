---
name: lyzbcy-cg
description: 通过 Edge 已登录会话检查江南大学 CG 在线作业、截止时间和个人提交状态；可按本人授权只读检索校园邮箱。
metadata:
  version: "1.0.2"
---

# CG 在线作业

用已经连接 Edge 的 kimi-webbridge。请在 Edge 登录 WebVPN 和 CG，再从「选择课程」复制课程列表地址。不要在命令、截图或公开材料里保存密码、Cookie 或令牌。

```powershell
python scripts/cg.py --url "<课程列表地址>" login-status
python scripts/cg.py --url "<课程列表地址>" assignments
```

也可由本机环境变量 `CG_WORKSPACE_URL` 提供地址，或在当前交接目录的 `agents.md` 写入课程列表地址。脚本只读取用户本人的课程和提交记录。搜索全部课程时会顺次切换站点记住的当前课程；运行中不要用另一标签切换课程。如果扫描报错，单独读取对应页，空白页面不能证明没有作业。

通过「在线作业 → 当前作业」复查有效期、多个题目各自的状态及文件要求。列出存在未完成题目的作业，由用户确定范围。实验报告可配合 [experiment-report-skill](https://github.com/lyzbcy/experiment-report-skill) 和课程提供的模板。文件上传使用平台正式表单，经用户授权后提交。等待超时先回看该题的下载入口与最近提交时间。CG 以服务端回执为准，核对前不要重复上传。

## 本人校园邮箱（可选，只读）

官方入口：https://mail.jiangnan.edu.cn/ 。学校[学生信息化指南](https://nic.jiangnan.edu.cn/info/1019/3539.htm)和[邮箱客户端配置](https://nic.jiangnan.edu.cn/info/1007/1817.htm)说明校园邮箱及IMAP；代码使用学校文档支持的通用服务 imap.exmail.qq.com，TLS证书校验开启。本人先激活邮箱并按官方指引开启客户端功能，按需要获取客户端专用密码。

```powershell
# 交互输入本人邮箱和密码（不回显）
python scripts/cg.py mail --days 7 --limit 20
# 只在主题/发件人中检索关键词；可选正文预览
python scripts/cg.py mail --days 30 --keyword "通知" --limit 15 --preview-chars 1500
```

非交互环境需要操作者明确授权、提供 `--account`，并通过其可信本地秘密管理器将密码送入 `--password-stdin`，不要将密码写入命令、公开脚本、日志或配置。不从浏览器提取凭证，不绕过认证。只查询本人 INBOX，使用 IMAP EXAMINE/readonly 和 BODY.PEEK，不修改已读标记、不发送、转发、删除邮件。默认仅输出邮件头；正文预览仅取有界的文本内容，不加载外部链接或下载附件；预览并非完整原文。输出为本地 JSON，属于私人内容，禁止加入公开仓库或发行包。 `scan_incomplete:true` 表示日期范围内尚未逐封扫描或存在无法读取的邮件，不能据此声称不存在指定邮件。当前版本采用合成 IMAP 测试，不提供对每个用户真实账户的登录成功保证；实际登录按本人状态核验。

## 静默更新

脚本每日本地首次运行后在后台检查同仓库公开发行版。下载包必须与 GitHub Release 的 SHA-256 和文件白名单相符，更新在下次运行生效。失败时继续使用现有版本。
