import uiautomation as auto
import time

wx = auto.WindowControl(searchDepth=1, Name="微信", ClassName='mmui::MainWindow')
wx.SetActive()  # 激活窗口
time.sleep(1)
# 微信主界面左上角有一个搜索框，通常是 "搜索" 或 "Search"
search_box = wx.EditControl(Name='搜索')
search_box.Click()
search_box.SendKeys('线报转发{Enter}')
time.sleep(1)


# 4. 找到输入框并发送消息
chat_edit = wx.EditControl(foundIndex=1)  # 聊天输入框
chat_edit.Click()
chat_edit.SendKeys('你好，这是自动化测试消息{Enter}')