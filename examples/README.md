# 示例代码

本目录包含一些简单的示例代码，用于演示和验证基础功能。

## wx.py - uiautomation 最小验证示例

这是一个最小的 uiautomation 使用示例，用于验证基础的 UI 自动化功能。

### 功能
- 激活微信窗口
- 在搜索框中搜索联系人
- 发送简单的测试消息

### 使用方法

```powershell
# 确保微信已启动并登录
python examples/wx.py
```

### 注意事项
- 需要修改代码中的联系人名称（默认为 "线报转发"）
- 这是一个最小示例，主要用于学习和验证 uiautomation 的基本用法
- 实际项目使用的是 `wechat_controller.py` 中更完善的实现

## 适用场景
- 学习 uiautomation 库的基本用法
- 验证 Windows UI 自动化环境是否正常
- 作为开发新功能的起点参考

