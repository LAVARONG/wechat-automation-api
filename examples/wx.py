import uiautomation as auto
import time

wx = auto.WindowControl(searchDepth=1, Name="微信", ClassName='mmui::MainWindow')
wx.SetActive()
time.sleep(1)

# 左上角搜索框，Name 在中文系统下为 "搜索"
search_box = wx.EditControl(Name='搜索')
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