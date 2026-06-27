# 更新日志

## 2026-05-30

### 优化：消息队列按联系人聚合调度，减少会话切换（v2.3.0）

**修改文件：** `scripts/message_queue.py`

**问题描述：**
- 同时给多个群发「一文一图」时，旧队列按入队顺序 `a文字 → b文字 → a图片 → b图片` 串行发送，会话切换路径为 `a → b → a → b`（切换 3 次），每次切换都要重新搜索联系人，耗时明显。

**解决方案：**
- ✅ 队列存储由 `queue.Queue`（FIFO）改为 `list + threading.Condition`，支持线程安全的动态重排。
- ✅ 采用贪心「粘连当前联系人」调度：优先把发往同一联系人的消息连续发完再切换；同一联系人内部保持入队顺序（先文字后图片）。上述场景被重排为 `a文字 → a图片 → b文字 → b图片`，切换次数降为 1 次。
- ✅ 增强：同联系人连发时跳过重复的联系人搜索，仅在确认会话停留时跳过，发送失败/异常则强制下一条重新搜索。
- ✅ 零延迟：依赖「单条发送耗时远大于入队间隔」，发送期间新入队的消息会被自动纳入重排，无需等待。

**兼容性：**
- `MessageQueue` 对外接口（`start` / `add_message` / `get_queue_size` / `stop`）保持不变，`app.py` 无需改动；发送总条数不变，仅改变发送顺序与搜索次数。

**详细文档：** `docs/2605/30队列调度优化.md`

---

### 新增：skill_cli 自动优先 API 队列模式（v2.2.0）

**修改文件：** `scripts/skill_cli.py`、`SKILL.md`

**背景：**
- 多个发信请求同时直接走 CLI 会争抢微信窗体/剪贴板，容易互相干扰。
- 后台 HTTP API（`app.py`）自带串行队列，多请求入队后逐条发送，不会冲突。

**改动内容：**
- ✅ `skill_cli.py` 新增执行模式自动判定：默认检测后台 API（`GET /health`）是否在线，在线则自动改走 HTTP 队列模式（`POST /`），离线时回退原有同步 CLI 模式。
- ✅ 新增 `--api`（强制 API 队列）与 `--no-api`（强制同步 CLI）开关。
- ✅ 自动模式下仅当 `config.json` 配齐 `token` 且 API 在线才走 API，否则静默回退 CLI，避免因缺 token 导致发送直接失败。
- ✅ 同步更新 `SKILL.md`：新增「执行模式与结果判定」一节，说明两种模式、`--api/--no-api` 参数，以及 API 队列模式下“已提交队列=受理成功、真实送达需查 /status 或日志”的语义差异。

**注意事项：**
- API 队列模式为异步：HTTP 返回成功仅代表消息已入队受理，不代表已成功送达。需要同步确认真实发送结果时，请用 `--no-api` 强制 CLI 模式，或查询后台 `/status` 与服务端日志。

---

## 2026-05-29

### 修复：uiautomation 日志写项目目录被拦截 & Windows 中文输出乱码（v2.1.5）

**修改文件：** `scripts/wechat_controller.py`、`scripts/skill_cli.py`

**问题描述：**
- `uiautomation` 默认在项目目录写 `@AutomationLog.txt`，在沙箱/受限环境下被拦截导致报错 `FileNotFoundError: can't write the log`
- `skill_cli.py` 在 Windows 下 stdout 使用 GBK 编码，中文输出乱码（显示为 `????`）

**解决方案：**
- ✅ `wechat_controller.py` 初始化时将 `auto.Logger.SetLogDir()` 重定向到系统临时目录 `%TEMP%\wechat_automation_logs`
- ✅ `skill_cli.py` 开头新增 `sys.stdout.reconfigure(encoding="utf-8")`，修复 Windows 下中文输出乱码

**优化效果：**
- 沙箱/受限环境下不再因日志写入失败而报错
- `skill_cli.py --help` 及所有中文输出正常显示

---

## 2026-05-27

### 修复：search_contact 剪贴板降级仍可能丢失中文（v2.1.4）

**修改文件：** `scripts/wechat_controller.py`

**问题描述：**
- v2.1.1（66be608）只把"剪贴板失败时含中文不再降级 SendKeys"的修复加到了 `send_message`，但 `search_contact` 里搜索框输入联系人名称的剪贴板兜底逻辑没有同步修复。
- 后续 v2.1.3（caa8843）大重构改 `SendResult` 返回值时，`search_contact_result` 这块剪贴板降级路径被原样搬过来，依然是 `interval=0.01` 的 `SendKeys` + 无中文检查。
- 实际表现：剪贴板偶发抢占时，中文联系人名（如「线报转发」「文件传输助手」）会丢字 → 搜不到联系人 → 整个 `search_and_send` 链路失败，但失败原因被掩盖为「CONTACT_NOT_FOUND」。

**解决方案：**
- ✅ `search_contact_result` 的剪贴板兜底逻辑对齐 `send_message_result`：
  - 联系人名称含中文/多字节字符且剪贴板失败时，直接返回 `CLIPBOARD_TEXT_FAILED`，不再降级 `SendKeys` 漏字搜错人
  - 纯 ASCII 名称才允许 `SendKeys` 兜底，`interval` 从 `0.01` 调整为 `0.05`

**优化效果：**
- 中文联系人搜索的剪贴板异常路径不再产生「假性找不到」错误
- 错误码更精确，便于上游 CLI/Skill 给出"请检查剪贴板"的针对性提示

---

## 2026-05-26

### 修复：sendpic 不支持本地图片路径（v2.1.2）

**修改文件：** `scripts/wechat_controller.py`、`scripts/skill_cli.py`

**问题描述：**
- `sendpic` 动作只支持图片 URL，传入本地路径时报错 `No connection adapters were found for 'C:\...'`
- `skill_cli.py` 的 `--content` 帮助信息写的是"图片URL"，未说明支持本地路径

**解决方案：**
- ✅ `send_picture()` 方法新增本地文件检测：`os.path.exists(image_url)` 为 True 时跳过下载，直接使用本地路径
- ✅ 更新 `skill_cli.py` 帮助信息：`--content` 说明改为"消息内容(发文本)或图片路径/URL(发图片)"

**优化效果：**
- 本地图片可直接发送，不再需要先上传获取 URL
- 调用示例：`python skill_cli.py --to "文件传输助手" --content "C:\path\to\pic.png" --action sendpic`

---

## 2026-05-23

### 修复：剪贴板中文失效导致消息残缺（v2.1.1）

**修改文件：** `scripts/wechat_controller.py`

**问题描述：**
- `uiautomation.SetClipboardText()` 对中文 Unicode 支持不好，设置后 `GetClipboardText()` 验证失败，触发 3 次重试后放弃
- 降级到 `SendKeys` 后 `interval=0.01`（10ms）太快，IME 来不及组字，导致中文/多字节字符丢失
- 实际表现：发送 "🦞 龙虾测试：WorkBuddy 发信功能正常！" 收到 "🦞 龙虾测试：orkBuddy 发信功能正常！"（W 丢失）

**解决方案：**
- ✅ 剪贴板操作改用 `pyperclip.copy()` / `pyperclip.paste()`，对 Unicode 支持完整
- ✅ 含中文/多字节字符时，剪贴板失败后**拒绝降级 SendKeys**，直接报错退出（避免产生残缺消息）
- ✅ 纯英文消息仍保留 SendKeys 降级路径，`interval` 从 0.01 调整为 0.05
- ✅ 新增 `import pyperclip`

**优化效果：**
- 中文消息 100% 完整发送
- 剪贴板验证通过率显著提升，不再有 "剪贴板内容验证失败，重试中..." 日志

---
## 2026-03-12

### 新增：微信掉线独立监控与 wpush 预警通知功能

**修改与新增内容：**
- **架构升级**：在 `scripts/` 下新增独立的守护式进程 `monitor.py`，专门负责监控微信自动化状态。
- **平滑集成**：在 `scripts/app.py` 启动阶段集成 `subprocess`，智能拉起 `monitor.py` 并在主程序退出时妥善回收（终止）监控进程资源。
- **自定义配置**：在 `config.json` 中扩展 `monitor_interval`、`monitor_max_retries` 和 `wpush` 配置项，允许用户设定检测间隔（例如60秒），最大重试次数（例如3次即静默不发送）以及 Wpush 的消息标题与内容。
- **防骚扰设计**：实现了监控失败次数的累加机制，只有在连续失败且次数未超过 `monitor_max_retries` 时才调用 API 发送请求；在微信窗口重新恢复获取时自动重置防骚扰计数。

## 2026-03-11

### 重点重构：适配 OpenClaw 等智能体的 Skill 化改造

**修改与新增内容：**
- **架构重构**：将 `app.py`, `wechat_controller.py`, `message_queue.py` 等核心逻辑抽离至 `scripts/` 目录下，与项目根目录配置产生解耦。
- **新增入口 `skill_cli.py`**：专为大语言模型或智能体系统（如 OpenClaw）打造的极简命令行调用入口，无缝支持即用即弃的 Skill 调用模式，极大简化调用链路并提供明确执行退出码反馈。
- **启动脚本 `run.bat`**：根目录新增便捷启动脚本，双击即可拉起原本的 HTTP 队列发送监听服务，保留向后兼容的独立运行能力。
- **配置优化**：升级 `app.py` 的绝对路径寻址逻辑，保证无论从哪个工作目录触发都能正确挂载 `config.json` 与 `wechat_automation.log`。
- **接入文档 `SKILL.md`**：新增标准的 Skill 调用定义规范文档。

---## 2025-11-29

### 修复：SendKeys 特殊字符丢失问题 & 健壮性增强

**修改文件：** `wechat_controller.py`

**问题描述：**
- 使用 `SendKeys` 发送消息时，中文特殊字符（如 `【`、`】`、`￥` 等）后面的字符会丢失
- 例如：`【拼多多】蝶安芬` 发送后变成 `【【多多】】安芬`
- 原因是 `SendKeys` 会将 `{` 和 `}` 等字符误解析为特殊按键序列

**解决方案：**
- ✅ 将文本发送方式从 `SendKeys` 改为**剪贴板粘贴**（`Ctrl+V`）
- ✅ 使用 `uiautomation.SetClipboardText()` 设置剪贴板内容
- ✅ 搜索联系人时也改用剪贴板粘贴方式输入

**优化效果：**
- 彻底解决特殊字符丢失问题
- 发送速度显著提升（粘贴比逐字符输入快得多）
- 支持所有 Unicode 字符

**健壮性增强：**
1. 新增 `_set_clipboard_text()` 方法：
   - 带 3 次重试机制
   - 自动验证剪贴板内容是否设置成功
   - 失败时回退到 `SendKeys`（并转义特殊字符）

2. 优化 `_copy_image_to_clipboard()` 方法：
   - 添加重试机制
   - 确保剪贴板被正确关闭（即使发生异常）

3. 优化 `_download_image()` 方法：
   - 添加下载重试机制（3次）
   - 添加缓存文件有效性验证
   - 添加 User-Agent 请求头
   - 使用 `image.verify()` 验证图片完整性

4. 所有发送方法添加窗口激活步骤 `wx.SetActive()`，确保焦点正确

5. 关键方法的异常处理添加 `exc_info=True`，记录完整堆栈信息

---

## 2025-11-05

### 优化：会话列表智能激活机制

**修改文件：** `wechat_controller.py`

**功能描述：**
- 优化联系人查找策略，优先从左侧会话列表直接激活对话（更快）
- 添加会话选中状态检测，避免重复点击已选中的会话
- 保留搜索框作为降级方案，确保兼容性

**优化效果：**
- 对于最近聊天的联系人，响应速度显著提升
- 避免误操作：已选中的会话不会被重复点击（防止取消选择）
- 双重验证：点击/搜索后都会验证会话是否真正选中
- 降低搜索框使用频率，减少UI操作步骤

**技术实现：**
1. 新增 `_is_session_selected()` 方法：
   - 通过 `GetPattern(10010)` 获取 SelectionItemPattern
   - 检查 `IsSelected` 属性判断会话是否已选中

2. 新增 `_activate_from_session_list()` 方法：
   - 通过 `AutomationId="session_item_{联系人名}"` 和 `ClassName="mmui::ChatSessionCell"` 定位会话项
   - 检查会话选中状态，已选中则跳过点击
   - 点击后验证是否成功选中

3. 优化 `search_contact()` 方法：
   - **策略1（优先）**：尝试从会话列表直接激活
   - **策略2（降级）**：找不到时使用搜索框
   - 搜索后也验证会话选中状态

## 2025-10-31

### 优化：图片发送缓存机制

**修改文件：** `wechat_controller.py`

**功能描述：**
- 使用 MD5(url) 作为缓存文件名，避免相同URL的图片重复下载
- 发送图片前先检查缓存目录是否已存在该文件
- 如果文件存在则直接使用缓存，不存在才下载
- 缓存目录：系统临时目录下的 `wechat_image_cache` 文件夹

**优化效果：**
- 相同URL的图片只下载一次，显著提高重复发送速度
- 减少网络带宽消耗
- 改善用户体验

**技术实现：**
1. 导入 `hashlib` 模块用于计算MD5
2. 在 `__init__` 中创建缓存目录
3. 修改 `_download_image` 方法添加缓存检查逻辑
4. 移除 `send_picture` 方法中删除临时文件的逻辑


