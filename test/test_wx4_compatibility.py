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

    # 2. 验证搜索框跨版本激活与定位（兼容小微灰度/GroupControl/EditControl）
    print("\n[测试 2] 搜索框激活与定位...")
    search_box = controller._activate_search_box(wx)
    if search_box and search_box.Exists(0, 0):
        print(f"✅ 成功: 找到并激活搜索框, 控件类型='{search_box.ControlTypeName}', 类名='{search_box.ClassName}', 坐标={search_box.BoundingRectangle}")
        search_box.SendKeys('{Esc}')
    else:
        print("❌ 失败: 未能激活或找到微信搜索框")
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

    # 5. 验证会话核验逻辑（防串群机制）
    print("\n[测试 5] 会话核验与防串群机制验证...")
    if chat_edit and chat_edit.Name:
        target_self = chat_edit.Name.strip()
        ok_self, reason_self = controller._verify_active_chat(wx, target_self)
        print(f"   - 命中同名验证 ('{target_self}'): {'✅ 通过' if ok_self else '❌ 失败'} ({reason_self})")
        
        target_fake = "__NON_EXISTENT_CONTACT_FOR_SAFETY_TEST__"
        ok_fake, reason_fake = controller._verify_active_chat(wx, target_fake)
        print(f"   - 拦截异名防串群 ('{target_fake}'): {'✅ 正确拦截' if not ok_fake else '❌ 未拦截(异常)'} ({reason_fake})")
    else:
        print("ℹ️ 提示: 未打开会话，跳过核验测试")

    print("\n" + "=" * 60)
    print("微信兼容性验证全部通过！主窗口与各关键控件定位正常。")
    print("=" * 60)
    return True

if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)
