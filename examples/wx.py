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

# 激活搜索框：优先通用快捷键 Ctrl+F（避免新版微信带有“小微AI”按钮或未聚焦态为 GroupControl 的控件匹配失效）
auto.SendKeys('{Ctrl}f')
time.sleep(0.3)
search_box = auto.GetFocusedControl()
if not search_box or search_box.ControlTypeName != 'EditControl':
    search_ctrl = wx.EditControl(Name='搜索', searchDepth=25)
    if not search_ctrl.Exists(0, 0):
        search_ctrl = wx.GroupControl(Name='搜索', searchDepth=25)
    if search_ctrl.Exists(0, 0):
        search_ctrl.Click()
        time.sleep(0.3)
        search_box = wx.EditControl(ClassName='mmui::XValidatorTextEdit', searchDepth=25)

if search_box and search_box.Exists(0, 0):
    search_box.SendKeys('{Ctrl}a{Back}')
    search_box.SendKeys('线报转发{Enter}')
    time.sleep(1)

# 聊天输入框：用稳定的 AutomationId + ClassName 精确定位
# 注意：新版微信打开公众号文章/视频时右侧会出现内置浏览器面板，
# 旧写法 wx.EditControl(foundIndex=1) 会被浏览器里的输入框抢走，必须改用 AutomationId
chat_edit = wx.EditControl(
    AutomationId="chat_input_field",
    ClassName="mmui::ChatInputField",
)
# 聚焦输入框：优先使用 SetFocus 避免移动用户鼠标光标
if not chat_edit.SetFocus():
    chat_edit.Click()
chat_edit.SendKeys('你好，这是自动化测试消息{Enter}')