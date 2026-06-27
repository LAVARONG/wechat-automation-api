# 微信自动化发信 Skill / HTTP API

这是一个面向 Windows 桌面微信的自动化发信工具，核心能力是通过 `uiautomation` 控制已登录的微信客户端发送文本、图片 URL 或本地文件。项目提供两种入口：

- **Skill / CLI 入口**：推荐给各类 Agent 使用。一次命令完成一次发送，执行后立即退出；只发送消息时无需启动后台 HTTP API。
- **HTTP API 入口**：适合外部系统长期集成、批量投递和队列处理，需要启动 Flask 后台服务。

本项目不使用 Hook、不接入微信协议，只通过 Windows UI 自动化操作本机微信窗口。

## Skill 用法

这是最推荐的使用方式：你只需要用自然语言告诉智能体要把什么内容发给谁，智能体会自动调用 Skill，不需要你手写命令，也不需要启动后台 HTTP API。

### 安装

```text
请帮我安装部署这个skill https://github.com/LAVARONG/wechat-automation-api ，完成后给微信联系人“文件传输助手”发送“这是来着微信自动化发信 Skill 的测试信息”
```

智能体会自动拉取仓库、读取 `SKILL.md`，并按项目说明准备运行环境。如果正常，会操作微信给“文件传输助手”发送测试信息。请提前确保 PC 版微信已经登录。

### 玩法示例

安装完成后，可以直接这样对智能体说：

```text
查询今天关于 AI 的最新新闻，整理成简洁的总结，然后通过微信发送到联系人“文件传输助手”。
```

```text
把这次运行结果整理成 5 条要点，通过微信发送到“项目通知群”。
```

```text
下载这张图片并通过微信发给“文件传输助手”。
```

```text
给微信的“文件传输助手”发送我下载目录下最新的一张图片。
```

```text
把刚生成的 C:\tmp\report.xlsx 通过微信发送给“文件传输助手”。
```

```text
把当前项目的测试结果总结一下，然后通过微信发给“项目通知群”。
```

普通发消息、发图片、发本地文件都走 Skill 即可，不需要运行 `run.bat`，也不需要启动 `scripts/app.py`。

## CLI 用法

CLI 适合开发者手动测试、排查环境问题，或在本机脚本里直接调用。

### 手动安装环境

先确保满足以下条件：

- Windows 10/11
- Python 3.7+
- 微信 PC 客户端已启动并登录
- 微信窗口可被系统 UI 自动化识别

在项目根目录安装依赖：

```powershell
pip install -r requirements.txt
```

### 发送文本

在项目根目录执行：

```powershell
python scripts/skill_cli.py --to "文件传输助手" --content "这是一条 Skill CLI 测试消息"
```

成功时 stdout 会输出：

```text
发送成功: 已向「文件传输助手」发送文本消息。
```

失败时会输出明确原因和错误代码，例如：

```text
发送失败: 未找到联系人或群组: 文件传输助手。请确认名称与微信备注/群名完全一致。
错误代码: CONTACT_NOT_FOUND
```

这通常表示微信窗口未找到、联系人名称不匹配、系统 UI 自动化不可用或剪贴板不可用。

### 发送图片 URL

图片发送使用 URL，脚本会下载图片并通过剪贴板粘贴到微信：

```powershell
python scripts/skill_cli.py --to "文件传输助手" --content "https://example.com/image.png" --action "sendpic"
```

### 发送本地图片

```powershell
python scripts/skill_cli.py --to "文件传输助手" --content "C:\tmp\screenshot.png" --action "sendpic"
```

如果传入的是本地路径，`sendpic` 会自动跳过下载，按本地文件粘贴发送。

### 发送本地文件

发送本机上的任意文件时使用 `sendfile`：

```powershell
python scripts/skill_cli.py --to "文件传输助手" --content "C:\tmp\report.xlsx" --action "sendfile"
```

建议 Agent 使用绝对路径，避免相对路径因工作目录不同而找不到文件。

### JSON 输出

需要稳定解析时可追加 `--json`：

```powershell
python scripts/skill_cli.py --to "文件传输助手" --content "测试消息" --json
```

输出包含 `success`、`code`、`message`、`to`、`action`。退出码 `0` 表示成功，非 `0` 表示失败。

## 系统辅助功能启用

若运行环境为精简版或 Ghost 系统，需确保 Windows 辅助功能正常运行。通过键盘 Win 键唤出开始菜单，输入“讲述人”并开启该功能（开启后可立即关闭），此操作可确保系统底层 UI 自动化接口处于激活状态，避免程序无法识别元素。

## 远程桌面与锁屏

本项目依赖 Windows UI Automation 操作微信窗口。Windows 锁屏状态下 UIA 通常不可用；服务器通过远程桌面 RDP 登录后，如果直接关闭远程桌面窗口，也可能导致鼠标、键盘和 UI 自动化失效，从而无法继续操作微信。

在服务器远程桌面场景中，建议使用项目自带的 `disconnect_rdp.bat` 断开远程桌面连接。它会尝试执行 `tscon <会话ID> /dest:console`，将当前 RDP 会话切回本地控制台并保持会话运行，用于解决关闭远程桌面后鼠标键盘失效的问题。

使用方式：

```powershell
.\disconnect_rdp.bat
```

如执行失败，请尝试以管理员身份运行，并确认当前确实处于远程桌面会话中。

## HTTP API 模式

HTTP API 模式适合长期运行的服务化场景，例如外部系统通过 HTTP 投递消息、批量发送到多个联系人、使用队列控制发送间隔。

### 准备配置

复制配置示例：

```powershell
Copy-Item config.json.example config.json
notepad config.json
```

配置示例：

```json
{
    "token": "your_secret_token_here",
    "host": "127.0.0.1",
    "port": 8808,
    "message_interval": 1,
    "log_level": "INFO",
    "log_file": "wechat_automation.log"
}
```

`config.json` 已加入 `.gitignore`，可用于保存本地 token。

### 启动服务

```powershell
.\run.bat
```

或直接运行：

```powershell
python scripts/app.py
```

启动后默认监听：

```text
POST http://127.0.0.1:8808/
GET  http://127.0.0.1:8808/status
GET  http://127.0.0.1:8808/health
```

### HTTP 发送文本

```powershell
$body = @{
    token = "your_secret_token_here"
    action = "sendtext"
    to = @("文件传输助手")
    content = "这是一条 HTTP API 测试消息"
} | ConvertTo-Json

Invoke-RestMethod -Uri "http://127.0.0.1:8808/" -Method Post -Body $body -ContentType "application/json"
```

### HTTP 发送图片 URL

```powershell
$body = @{
    token = "your_secret_token_here"
    action = "sendpic"
    to = @("文件传输助手")
    content = "https://example.com/image.png"
} | ConvertTo-Json

Invoke-RestMethod -Uri "http://127.0.0.1:8808/" -Method Post -Body $body -ContentType "application/json"
```

### HTTP 发送本地文件

```powershell
$body = @{
    token = "your_secret_token_here"
    action = "sendfile"
    to = @("文件传输助手")
    content = "C:\tmp\report.xlsx"
} | ConvertTo-Json

Invoke-RestMethod -Uri "http://127.0.0.1:8808/" -Method Post -Body $body -ContentType "application/json"
```

成功响应示例：

```json
{
    "success": true,
    "message": "消息已加入队列",
    "queued_count": 1,
    "queue_size": 1
}
```

## 项目结构

```text
wechat-automation-api/
├── scripts/
│   ├── skill_cli.py            # Agent Skill / CLI 同步发送入口
│   ├── wechat_controller.py    # 微信窗口查找、联系人搜索、文本/图片发送逻辑
│   ├── app.py                  # Flask HTTP API 服务入口
│   ├── message_queue.py        # HTTP 模式消息队列
│   └── monitor.py              # 微信状态监控脚本
├── examples/
│   └── wx.py                   # uiautomation 最小验证示例
├── test/
│   └── test_api.py             # HTTP API 测试脚本
├── docs/
│   └── changelog.md            # 更新日志
├── SKILL.md                    # Agent Skill 使用说明
├── requirements.txt            # Python 依赖
├── config.json.example         # HTTP API 配置示例
├── run.bat                     # HTTP API 快捷启动脚本
└── README.md                   # 项目说明
```

## 工作原理

Skill / CLI 模式：

1. Agent 或用户执行 `python scripts/skill_cli.py ...`
2. 脚本创建 `WeChatController`
3. 控制器查找并激活微信窗口
4. 优先从会话列表定位联系人，失败后使用搜索框
5. 通过剪贴板粘贴文本、图片或本地文件并按 Enter 发送
6. 命令返回 `发送成功` 或失败原因并退出

HTTP API 模式：

1. Flask 接收 HTTP 请求并验证 token
2. 请求被加入消息队列
3. 后台线程按顺序调用 `WeChatController`
4. 按配置的 `message_interval` 控制发送间隔

## 开发者辅助工具

编写或调整元素查找逻辑时，建议准备 UI 自动化查看工具：

- 使用 Visual Studio SDK 自带的 `Inspect.exe` 工具，或 Windows SDK 工具包中的 Inspect 工具。
- 用于实时查看微信窗口的 UI 元素树结构、`Name`、`AutomationID`、`ClassName` 等关键属性。
- 当微信版本更新导致控件名称或层级变化时，优先用 Inspect 重新确认元素属性，再修改 `scripts/wechat_controller.py` 中的查找逻辑。

当前代码中比较关键的定位点包括：

- 微信主窗口：`Name="微信"`，`ClassName="mmui::MainWindow"`
- 会话项：`ClassName="mmui::ChatSessionCell"`，`AutomationId="session_item_<联系人名>"`
- 搜索框：`EditControl(Name="搜索")`
- 聊天输入框：`EditControl(AutomationId="chat_input_field", ClassName="mmui::ChatInputField")`
  - 旧版 `EditControl(foundIndex=1)` 在新版微信打开公众号文章/视频后会错位（右侧内置浏览器有自己的输入框），已弃用。

## 常见问题

### 提示未找到微信窗口

确认微信 PC 客户端已启动并登录。程序会尝试通过微信默认快捷键 `Ctrl+Alt+W` 唤醒窗口；如果仍失败，请检查微信快捷键设置、窗口标题和系统 UI 自动化能力。

### 精简版或 Ghost 系统无法识别元素

按“系统辅助功能启用”章节操作一次：Win 键打开开始菜单，搜索并开启“讲述人”，开启后可立即关闭。这可以激活底层 UI 自动化接口。

### 关闭远程桌面后无法操作微信

不要直接锁屏或直接关闭 RDP 窗口后期待 UIA 继续稳定工作。请使用 `disconnect_rdp.bat` 断开远程桌面连接并保持会话回到本地控制台；必要时以管理员身份运行该脚本。

### 找不到联系人或发错会话

检查 `--to` 或 HTTP 请求中的 `to` 是否与微信里的联系人名称、群名称或备注完全一致。建议先用“文件传输助手”完成发送测试，再切换到真实联系人。

### 文本包含中文、换行或特殊字符是否支持

支持。文本发送优先使用剪贴板粘贴，不依赖逐字键盘输入，因此比 `SendKeys` 更适合中文和特殊符号。

### 什么时候需要 HTTP API

只有当你需要外部系统通过 HTTP 调用、后台队列、批量发送、状态查询或长期服务化运行时才需要 HTTP API。普通 Agent 发消息直接使用 Skill / CLI。

### CLI 常见错误代码

- `WECHAT_WINDOW_NOT_FOUND`：未找到微信窗口。确认微信 PC 已启动并登录；精简版或 Ghost 系统可开启一次“讲述人”激活辅助功能。
- `CONTACT_NOT_FOUND`：未找到联系人或群组。确认名称与微信备注、联系人名或群名完全一致。
- `SEARCH_BOX_NOT_FOUND`：未找到微信搜索框。可能是微信 UI 变化、窗口未加载完成或系统 UI 自动化不可用。
- `CHAT_INPUT_NOT_FOUND`：未找到聊天输入框。可能未进入目标会话或微信 UI 结构变化。
- `CLIPBOARD_TEXT_FAILED`：文本写入剪贴板失败。检查系统剪贴板、`pyperclip` 或远程桌面会话状态。
- `IMAGE_DOWNLOAD_FAILED` / `IMAGE_URL_NOT_IMAGE`：图片 URL 下载失败或返回内容不是图片。
- `IMAGE_CLIPBOARD_FAILED`：图片复制到剪贴板失败。
- `LOCAL_FILE_NOT_FOUND`：本地文件不存在。检查路径是否正确，建议使用绝对路径。
- `LOCAL_FILE_INVALID`：路径不是文件，当前不支持发送目录。
- `FILE_CLIPBOARD_FAILED`：本地文件复制到剪贴板失败。检查系统剪贴板或远程桌面会话状态。
- `SEND_FILE_ERROR`：发送本地文件异常。
- `DEPENDENCY_MISSING`：缺少 Python 依赖。请在项目根目录执行 `pip install -r requirements.txt`。
- `CLI_EXCEPTION`：CLI 自身执行异常。

## 日志

HTTP API 模式会写入 `wechat_automation.log`：

```powershell
Get-Content wechat_automation.log -Tail 50
Get-Content wechat_automation.log -Wait
```

Skill / CLI 模式主要通过 stdout 返回执行结果。

## 许可证

本项目仅供学习和研究使用。请遵守微信客户端使用规范，并在自己的账号和设备上谨慎使用。
