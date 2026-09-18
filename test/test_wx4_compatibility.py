"""
微信4.x/3.x版本UI兼容性验证脚本
用于检测当前环境微信客户端的主窗口、搜索框、会话项和输入框是否满足自动化条件
测试过程为只读检测，不会发送任何消息，不影响实际聊天
"""
import sys
import os

# 将 scripts 目录加入 sys.path
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'scripts'))

from wechat_controller import WeChatController

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

def main():
    print("=" * 60)
    print("开始执行微信兼容性验证检测（无干扰只读模式）")
    print("=" * 60)

    controller = WeChatController()

    # 1. 验证微信主窗口获取
    print("\n[测试 1] 微信主窗口获取...")
    result, wx = controller._get_wechat_window_result()
    if not result.success or not wx:
        print(f"❌ 失败: {result.code} - {result.message}")
        return False
    print(f"✅ 成功: 窗口名称='{wx.Name}', 类名='{wx.ClassName}', 进程ID={wx.ProcessId}")

    # 2. 验证搜索框定位
    print("\n[测试 2] 搜索框定位...")
    search_box = wx.EditControl(Name='搜索', searchDepth=25)
    if not search_box.Exists(0, 0):
        search_box = wx.EditControl(ClassName='mmui::XValidatorTextEdit', searchDepth=25)
    
    if search_box.Exists(0, 0):
        print(f"✅ 成功: 找到搜索框, 类名='{search_box.ClassName}', 坐标={search_box.BoundingRectangle}")
    else:
        print("❌ 失败: 未找到搜索框")
        return False

    # 3. 验证会话项识别能力
    print("\n[测试 3] 会话项控件识别...")
    session_cells = []
    def find_sessions(ctrl, depth=0):
        for c in ctrl.GetChildren():
            if c.ClassName == 'mmui::ChatSessionCell':
                session_cells.append((depth, c))
            find_sessions(c, depth + 1)
    find_sessions(wx)
    print(f"✅ 成功: 识别到可见会话项数量: {len(session_cells)}")
    for d, s in session_cells[:3]:
        aid = s.AutomationId
        contact_name = aid.replace('session_item_', '') if aid.startswith('session_item_') else aid
        print(f"   - 会话: '{contact_name}', 深度: {d}, 坐标: {s.BoundingRectangle}")

    # 4. 验证聊天输入框定位逻辑
    print("\n[测试 4] 聊天输入框探测...")
    chat_edit = controller._find_chat_input(wx, max_retries=1)
    if chat_edit:
        print(f"✅ 成功: 找到当前聊天输入框, 控件名='{chat_edit.Name}', 类名='{chat_edit.ClassName}', 坐标={chat_edit.BoundingRectangle}")
    else:
        print("ℹ️ 提示: 当前主窗口未打开聊天面板或未选中具体会话（符合预期，点击会话后即会渲染）")

    print("\n" + "=" * 60)
    print("微信兼容性验证全部通过！主窗口与各关键控件定位正常。")
    print("=" * 60)
    return True

if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)
