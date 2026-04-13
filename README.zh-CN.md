# Email Automation Based on Browser

[English](./README.md) | [简体中文](./README.zh-CN.md)

这是一个基于浏览器会话的 Outlook Web 自动化项目，用于周期性邮件简报和日历更新。

这个项目不依赖桌面版 Outlook 的 COM/MAPI，也不需要单独创建 Microsoft Graph 应用注册。它会复用已经登录的 Outlook Web 会话，从 Playwright 的 storage state 里提取 Outlook Web access token，并调用网页端实际使用的 Outlook REST 接口。

## 仓库内容

- `scripts/outlook_helper.py`：用于登录刷新、会话检查、近期邮件抓取、日历事件创建的 CLI
- `scripts/outlook_briefing_data.py`：给 Codex 或其他自动化 agent 用的紧凑交接数据生成器
- `skills/outlook-browser-email-automation`：对上述脚本与工作流做封装的可安装 Codex skill
- `automation-examples/outlook-browser-email.toml`：调用 skill 的简短自动化示例，而不是把整套流程硬编码进 prompt

## 安装与配置

安装依赖：

```powershell
python -m pip install -r requirements.txt
```

只有默认配置不够用时，才复制并修改示例配置：

```powershell
Copy-Item .\scripts\mail_automation_web_config.example.json .\scripts\mail_automation_web_config.json
```

默认状态目录在 `%LOCALAPPDATA%\codex-mail-automation-web` 下。不要提交生成出来的配置、浏览器状态文件或调试输出。

常用配置字段：

- `local_time_zone`：解释本地邮件与日历时间用的 IANA 时区，例如 `Europe/London`
- `outlook_time_zone`：创建事件时使用的 Outlook 或 Exchange 时区名，例如 `GMT Standard Time`
- `edge_user_data_dir` 和 `edge_profile_directory`：首次网页登录所使用的浏览器 profile

## 首次登录

```powershell
python .\scripts\outlook_helper.py web-login
```

然后在浏览器里完成微软登录流程。Outlook Mail 加载完成后，helper 会退出并保存本地 storage state。

## 常用命令

检查会话：

```powershell
python .\scripts\outlook_helper.py web-status
```

确保会话至少还能再有效三小时：

```powershell
python .\scripts\outlook_helper.py ensure-session --min-valid-seconds 10800
```

抓取近期邮件：

```powershell
python .\scripts\outlook_helper.py recent-mail --hours 24 --max-items 250
```

生成自动化交接数据：

```powershell
python .\scripts\outlook_briefing_data.py --ensure-session --hours 24 --max-items 250
```

## 安装 Codex Skill

把 `skills/outlook-browser-email-automation` 复制到你的 Codex skills 目录：

```powershell
$codexHome = if ($env:CODEX_HOME) { $env:CODEX_HOME } else { Join-Path $env:USERPROFILE ".codex" }
New-Item -ItemType Directory -Force -Path (Join-Path $codexHome "skills") | Out-Null
Copy-Item .\skills\outlook-browser-email-automation (Join-Path $codexHome "skills") -Recurse -Force
```

然后这样调用：

```text
Use $outlook-browser-email-automation to review recent Outlook mail and update calendar items.
```

## 说明

- 这是浏览器会话自动化，不是官方 Microsoft Graph 集成。
- 如果 Outlook 需要密码、Windows Hello 或 MFA，请手动执行 `web-login`。
- 已保存的浏览器状态具备邮件访问能力，务必妥善保管。
- 日历创建逻辑偏保守，会对同标题且同时间的事件做去重。

## 许可证

见 `LICENSE`。
