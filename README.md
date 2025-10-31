# 微信自动化 HTTP 服务

基于 Flask 和 uiautomation 的微信消息自动化发送系统。通过 HTTP API 接收请求，使用队列管理，自动控制微信 PC 客户端发送消息。

## ✨ 项目特性

- 🌐 **HTTP API 接口** - RESTful API，易于集成
- 🔐 **Token 身份验证** - 保证接口安全
- 📋 **消息队列管理** - 按顺序可靠发送
- 📨 **批量发送支持** - 一次请求发送给多个联系人
- 💬 **支持文本消息** - 发送文本内容（支持换行）
- 🖼️ **支持图片消息** - 通过 URL 下载图片后发送
- ⏱️ **自动间隔控制** - 每条消息间隔 1 秒，可配置
- 📝 **完善的日志系统** - 记录所有操作和错误
- 🛡️ **错误容错机制** - 单条失败不影响后续消息

## 🚀 快速开始

### 1. 环境要求

- Windows 10/11
- Python 3.7+
- 微信 PC 客户端（已登录）

### 2. 安装依赖

```powershell
pip install -r requirements.txt
```

### 3. 配置文件

```powershell
# 复制配置文件示例
Copy-Item config.json.example config.json

# 编辑配置文件，修改 token 等配置
notepad config.json
```

### 4. 启动服务

```powershell
# 确保微信已启动并登录
python app.py
```

看到以下输出表示启动成功：

```
========================================
微信自动化服务已启动
监听地址: http://127.0.0.1:8808
API 端点: POST http://127.0.0.1:8808/
========================================
```

### 5. 发送测试消息

#### PowerShell 示例

```powershell
# 发送文本消息
$body = @{
    token = "123123"
    action = "sendtext"
    to = @("线报转发")
    content = "你好，这是测试消息"
} | ConvertTo-Json

Invoke-RestMethod -Uri "http://127.0.0.1:8808/" -Method Post -Body $body -ContentType "application/json"

# 发送图片消息
$body = @{
    token = "123123"
    action = "sendpic"
    to = @("线报转发")
    content = "https://example.com/image.jpg"
} | ConvertTo-Json

Invoke-RestMethod -Uri "http://127.0.0.1:8808/" -Method Post -Body $body -ContentType "application/json"
```

#### Python 示例

```python
import requests

url = "http://127.0.0.1:8808/"

# 发送文本消息
data = {
    "token": "123123",
    "action": "sendtext",
    "to": ["线报转发", "LAVA"],  # 可以同时发给多人
    "content": "你好，这是自动化测试消息"
}
response = requests.post(url, json=data)
print(response.json())

# 发送图片消息
data = {
    "token": "123123",
    "action": "sendpic",
    "to": ["线报转发"],
    "content": "https://example.com/image.jpg"
}
response = requests.post(url, json=data)
print(response.json())
```

#### 使用测试脚本

```powershell
python test/test_api.py
```

## 📖 API 文档

### 发送消息

**端点**: `POST http://127.0.0.1:8808/`

#### 发送文本消息

**请求体**:
```json
{
    "token": "123123",
    "action": "sendtext",
    "to": ["联系人1", "联系人2"],
    "content": "消息内容"
}
```

#### 发送图片消息

**请求体**:
```json
{
    "token": "123123",
    "action": "sendpic",
    "to": ["联系人1", "联系人2"],
    "content": "https://example.com/image.jpg"
}
```

**成功响应** (200):
```json
{
    "success": true,
    "message": "消息已加入队列",
    "queued_count": 2,
    "queue_size": 2
}
```

**失败响应** (401):
```json
{
    "success": false,
    "error": "无效的 token"
}
```

### 查询状态

**端点**: `GET http://127.0.0.1:8808/status`

**响应**:
```json
{
    "status": "running",
    "queue_size": 5
}
```

### 健康检查

**端点**: `GET http://127.0.0.1:8808/health`

## 📁 项目结构

```
winappdriver/
├── app.py                      # Flask 主服务
├── wechat_controller.py        # 微信控制器
├── message_queue.py            # 消息队列管理
├── config.json.example         # 配置文件示例
├── requirements.txt            # Python 依赖
├── disconnect_rdp.bat          # RDP 断开脚本
├── test/                       # 测试文件目录
│   ├── test_api.py            # API 测试脚本
│   └── README.md              # 测试说明
├── examples/                   # 示例代码目录
│   ├── wx.py                  # uiautomation 最小示例
│   └── README.md              # 示例说明
├── docs/                       # 文档目录
│   └── changelog.md           # 更新日志
├── README.md                   # 本文件
├── 使用说明.md                 # 详细使用说明
├── 快速启动指南.md             # 快速入门
├── .gitignore                  # Git 忽略文件
├── .cursorrules                # Cursor 规则文件
└── wechat_automation.log       # 日志文件（运行时生成）
```

## ⚙️ 配置说明

首次使用需要复制 `config.json.example` 为 `config.json` 并编辑：

```powershell
Copy-Item config.json.example config.json
```

配置项说明：

```json
{
    "token": "your_secret_token_here",  // API 访问令牌（请修改为自己的密钥）
    "host": "127.0.0.1",                // 服务监听地址
    "port": 8808,                       // 服务监听端口
    "message_interval": 1,              // 消息发送间隔（秒）
    "log_level": "INFO",                // 日志级别（DEBUG/INFO/WARNING/ERROR）
    "log_file": "wechat_automation.log" // 日志文件路径
}
```

**注意**：`config.json` 已加入 `.gitignore`，不会被提交到 Git，可以安全地存储您的 token。

## 🔧 工作原理

1. **接收请求** - HTTP API 接收消息发送请求
2. **验证 Token** - 验证请求的身份令牌
3. **加入队列** - 消息立即加入队列，返回成功响应
4. **后台处理** - 独立线程按顺序处理队列中的消息
5. **控制微信** - 使用 uiautomation 搜索联系人并发送消息
6. **间隔控制** - 每条消息发送后等待指定时间（默认 1 秒）

## 📚 使用场景

- 📢 **消息群发** - 一键发送通知给多个联系人
- 🤖 **自动回复** - 结合其他系统实现自动化回复
- 📊 **报警通知** - 监控系统发送报警消息到微信
- 🔔 **定时提醒** - 设置定时任务发送提醒消息
- 🌐 **系统集成** - 集成到现有系统中实现微信通知

## 📝 注意事项

### 使用前准备
- ✅ 确保微信 PC 客户端已启动并登录
- ✅ 微信可以最小化，但不能关闭
- ✅ 确保联系人名称准确（区分大小写）

### 消息发送逻辑
- 消息会立即加入队列并返回成功响应
- 后台线程按顺序处理队列中的消息
- 每条消息发送后自动等待 1 秒（可配置）
- 某个联系人发送失败会跳过并继续处理下一条

### 日志查看

```powershell
# 实时查看日志
Get-Content wechat_automation.log -Wait

# 查看最后 50 行
Get-Content wechat_automation.log -Tail 50
```

## 🛠️ 技术栈

- **Flask** - 轻量级 Web 框架
- **uiautomation** - Windows UI 自动化库
- **threading** - 多线程支持
- **queue** - 线程安全队列
- **logging** - 日志系统
- **requests** - HTTP 请求库，用于下载图片
- **Pillow** - 图片处理库
- **pywin32** - Windows API 支持，用于剪贴板操作

## 📖 完整文档

- [使用说明.md](使用说明.md) - 详细的使用说明和 API 文档
- [快速启动指南.md](快速启动指南.md) - 5 分钟快速上手指南

## ❓ 常见问题

**Q: 提示"未找到微信窗口"**
> 确保微信已启动并且窗口标题为"微信"

**Q: 找不到联系人**
> 检查联系人名称是否正确，注意大小写

**Q: 如何修改 Token**
> 编辑 `config.json` 文件中的 `token` 字段

**Q: 如何修改消息间隔时间**
> 编辑 `config.json` 文件中的 `message_interval` 字段（单位：秒）

**Q: 可以发送图片或文件吗**
> 已支持通过 URL 发送图片（使用 action: "sendpic"），文件功能将在后续版本添加

**Q: 图片支持哪些格式**
> 支持常见图片格式：JPG、PNG、GIF、BMP 等

## 🎯 未来计划

- [x] 支持发送图片（已完成）
- [ ] 支持发送文件
- [ ] 消息发送状态查询
- [ ] Web 管理界面
- [ ] 消息发送历史记录
- [ ] 支持群聊消息

## 📜 版本历史

### v1.1.0 (2025-10-30)
- ✅ 新增图片发送功能（sendpic action）
- ✅ 支持通过 URL 下载图片并发送
- ✅ 自动剪贴板操作发送图片
- ✅ 自动清理临时文件

### v1.0.0 (2025-10-30)
- ✅ 实现基础 HTTP API 服务
- ✅ 实现消息队列管理
- ✅ 实现微信自动化控制
- ✅ 添加 Token 身份验证
- ✅ 添加完善的日志系统
- ✅ 支持批量发送消息

## 📄 许可证

本项目仅供学习和研究使用。

---

## 附加项目

本仓库还包含其他 Windows 自动化测试案例：

### WinAppDriver 测试案例

基于 WinAppDriver 的 Windows 应用程序自动化测试示例（位于 `temp/` 目录）。

- 记事本自动化测试
- 计算器自动化测试
- 元素检查工具

详见 [temp/README.md](temp/README.md)

