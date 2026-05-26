---
name: "wechat-automation"
description: "微信发信能力：控制本机发送微信文本或图片。触发词：给[某人]发微信、通知[某人]、把这个用微信发给[某人]、用微信发个图片"
---

# 微信自动发信 Skill

通过命令行脚本控制 Windows 桌面微信向联系人或群组发送文本或图片消息。Use when the user asks to send a WeChat message, notify a contact, or forward content via WeChat.

## 执行流程

1. **确认收件人** — 向用户确认 `--to` 参数所用的联系人/群名，必须与微信内备注或网名**完全一致**（含大小写和空格）。发错人不可撤回，务必先核实。
2. **构建并执行命令** — 根据消息类型选择对应命令：

### 发送文本消息

```bash
python scripts/skill_cli.py --to "联系人或群组名称" --content "你要发送的具体文本"
```

### 发送图片消息

```bash
python scripts/skill_cli.py --to "联系人或群组名称" --content "https://example.com/image.png" --action "sendpic"
```

3. **检查结果** — 读取 stdout 判断执行结果：
   - `发送成功` → 报告成功
   - `发送失败` → **原样**反馈错误信息给用户，不自动重试
   - `执行异常` → 报告异常详情给用户

## 约束

- `--to` 名称必须与微信内备注/网名**完全匹配**，否则可能发错人或失败
- 不要使用视觉模型分析微信界面 — 底层脚本已封装所有窗体操作
- 内容可含换行符（`\n`）和表情符号，脚本自动处理转义
