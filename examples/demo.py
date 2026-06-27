import uiautomation as auto

# 方案 A：通过代码尝试“唤醒”UI 树
def wake_up_wechat_uia():
    wechat_win = auto.WindowControl(searchDepth=1, Name="微信")
    if wechat_win.Exists(0):
        # 尝试切换焦点或发送特定指令触发 UI 暴露
        wechat_win.SetFocus()
        # 某些版本按下 Ctrl+Shift+Alt+D 可能会开启调试/完整 UI 模式
        auto.SendKeys('测试') 
        auto.SendKeys('{Enter}')

if __name__ == "__main__":
    wake_up_wechat_uia()