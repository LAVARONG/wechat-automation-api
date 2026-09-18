import sys
import ctypes

# 确保线程绑定至用户交互式 default 桌面（避免自动化/任务环境下找不到窗口）
if sys.platform == "win32":
    try:
        h_desk = ctypes.windll.user32.OpenDesktopW("default", 0, False, 0x1FF)
        if h_desk:
            ctypes.windll.user32.SetThreadDesktop(h_desk)
    except Exception:
        pass

import uiautomation as auto
import time

# 微信主窗口：微信 4.x 架构下类名为 mmui::MainWindow，窗口名称为登录用户昵称（不能限定 Name="微信"）
wx = auto.WindowControl(searchDepth=1, ClassName='mmui::MainWindow')
if not wx.Exists(0, 0):
    wx = auto.WindowControl(searchDepth=1, ClassName='WeChatMainWndForPC')
if not wx.Exists(0, 0):
    wx = auto.WindowControl(searchDepth=1, Name="微信")
wx.SetActive()
time.sleep(1)

# 左上角搜索框，Name 在中文系统下为 "搜索"，微信 4.x 深度通常需 >= 20
search_box = wx.EditControl(Name='搜索', searchDepth=25)
search_box.Click()
search_box.SendKeys('线报转发{Enter}')
time.sleep(1)

# 聊天输入框：用稳定的 AutomationId + ClassName 精确定位
# 注意：新版微信打开公众号文章/视频时右侧会出现内置浏览器面板，
# 旧写法 wx.EditControl(foundIndex=1) 会被浏览器里的输入框抢走，必须改用 AutomationId
chat_edit = wx.EditControl(
    AutomationId="chat_input_field",
    ClassName="mmui::ChatInputField",
)
chat_edit.Click()
chat_edit.SendKeys('你好，这是自动化测试消息{Enter}')