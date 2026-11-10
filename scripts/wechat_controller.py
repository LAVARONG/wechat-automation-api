"""
微信控制器模块
封装微信操作，提供搜索联系人和发送消息的功能
"""
import sys
import ctypes

# 缓存 default 桌面句柄，避免多线程或重复调用造成句柄泄露
_cached_desktop_handle = None


def _attach_to_default_desktop():
    """
    确保当前线程绑定到 Windows 交互式桌面 'default'。
    在子进程、Agent 或隔离桌面环境下，线程可能默认关联到隔离虚拟桌面，
    必须在导入任何 UI/COM 库（如 uiautomation）之前切换，否则线程会被锁定在隔离桌面。
    """
    global _cached_desktop_handle
    if sys.platform != "win32":
        return
    try:
        user32 = ctypes.windll.user32
        if _cached_desktop_handle is None:
            # DESKTOP_ALL_ACCESS 权限掩码为 0x1FF
            _cached_desktop_handle = user32.OpenDesktopW("default", 0, False, 0x1FF)
        if _cached_desktop_handle:
            user32.SetThreadDesktop(_cached_desktop_handle)
    except Exception:
        pass


# 必须在导入 uiautomation 之前切换桌面
_attach_to_default_desktop()

import uiautomation as auto
import time
import logging
import requests
import os
import sys
import json
import struct
import tempfile
from dataclasses import dataclass
from contextlib import contextmanager
from functools import wraps
from PIL import Image
from io import BytesIO
import win32clipboard
import win32con
from pathlib import Path
import hashlib
import pyperclip

# 将 uiautomation 日志目录重定向到系统临时目录，避免在项目目录写文件被拦截
_auto_log_dir = os.path.join(tempfile.gettempdir(), 'wechat_automation_logs')
if not os.path.exists(_auto_log_dir):
    os.makedirs(_auto_log_dir, exist_ok=True)
try:
    auto.Logger.SetLogDir(_auto_log_dir)
except Exception:
    pass

# 禁用 uiautomation 默认的模拟慢速鼠标移动（默认 1 秒慢速滑行），实现瞬时定位
auto.MAX_MOVE_SECOND = 0

# 配置日志
logger = logging.getLogger(__name__)


class _POINT(ctypes.Structure):
    """Windows API 坐标点结构"""
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


class _RECT(ctypes.Structure):
    """Windows API 矩形范围结构"""
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


def _get_project_root():
    """获取项目根目录绝对路径"""
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _load_lock_mouse_config():
    """从 config.json 加载 lock_mouse_during_send 配置，默认 False"""
    try:
        config_path = os.path.join(_get_project_root(), "config.json")
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                return bool(cfg.get("lock_mouse_during_send", False))
    except Exception:
        pass
    return False


@dataclass
class SendResult:
    """一次发送操作的结果，供 CLI/Skill 输出明确反馈。"""
    success: bool
    code: str
    message: str


def _with_input_lock(func):
    """方法装饰器：为操作添加安全输入保护锁（方案 2 可选安全锁）。"""
    @wraps(func)
    def wrapper(self, *args, **kwargs):
        with self._safe_input_lock():
            return func(self, *args, **kwargs)
    return wrapper


class WeChatController:
    """微信控制器类，用于自动化控制微信发送消息"""
    
    def __init__(self, lock_mouse_during_send=None):
        """
        初始化微信控制器
        
        Args:
            lock_mouse_during_send: 可选布尔值。发送期间是否临时锁定鼠标，防止用户误触干扰。
                                    若为 None，则自动从 config.json 的 lock_mouse_during_send 项加载。
        """
        self.wx = None
        if lock_mouse_during_send is None:
            self.lock_mouse_during_send = _load_lock_mouse_config()
        else:
            self.lock_mouse_during_send = bool(lock_mouse_during_send)

        self._input_lock_depth = 0
        self._is_clipped = False
        # 创建图片缓存目录
        self.cache_dir = os.path.join(tempfile.gettempdir(), 'wechat_image_cache')
        if not os.path.exists(self.cache_dir):
            os.makedirs(self.cache_dir)
            logger.info(f"创建图片缓存目录: {self.cache_dir}")

    @contextmanager
    def _safe_input_lock(self):
        """
        可重入的安全输入保护锁上下文（方案 2 可选安全锁）。
        当 self.lock_mouse_during_send 为 True 时：
        - 进入时：记录当前光标位置，使用 ClipCursor 锚定光标至 1x1 像素（普通权限即生效），
                  并尝试 BlockInput(True) 系统级阻断输入（需管理员权限）。
        - 退出时：在 finally 中 100% 解除 BlockInput(False) 与 ClipCursor(None)，
                  并将鼠标光标瞬间复位，杜绝死锁。
        """
        if not self.lock_mouse_during_send or sys.platform != "win32":
            yield
            return

        should_lock = (self._input_lock_depth == 0)
        self._input_lock_depth += 1
        orig_pos = None

        try:
            if should_lock:
                orig_pos = self._lock_input()
            yield
        finally:
            self._input_lock_depth -= 1
            if should_lock:
                self._unlock_input(orig_pos)

    def _lock_input(self):
        """执行硬件/系统级输入锁定"""
        orig_pos = None
        try:
            user32 = ctypes.windll.user32
            pt = _POINT()
            if user32.GetCursorPos(ctypes.byref(pt)):
                orig_pos = (pt.x, pt.y)
                # 约束光标至当前像素点，防止鼠标被用户意外滑动
                rect = _RECT(pt.x, pt.y, pt.x + 1, pt.y + 1)
                user32.ClipCursor(ctypes.byref(rect))
                self._is_clipped = True

            # 尝试硬件输入级阻断（若有管理员提权，键盘+鼠标完全阻断）
            user32.BlockInput(True)
            logger.debug("已激活发信输入锁定保护")
        except Exception as e:
            logger.debug(f"激活输入锁定异常: {e}")
        return orig_pos

    def _unlock_input(self, orig_pos=None):
        """解除输入锁定并恢复鼠标原始坐标"""
        try:
            user32 = ctypes.windll.user32
            user32.BlockInput(False)
            user32.ClipCursor(None)
            self._is_clipped = False
            if orig_pos:
                user32.SetCursorPos(int(orig_pos[0]), int(orig_pos[1]))
            logger.debug("已解除发信输入锁定保护并复位光标")
        except Exception as e:
            logger.debug(f"解除输入锁定异常: {e}")

    def _get_cursor_pos(self):
        """获取当前物理鼠标光标位置"""
        try:
            if sys.platform == "win32":
                pt = _POINT()
                if ctypes.windll.user32.GetCursorPos(ctypes.byref(pt)):
                    return pt.x, pt.y
        except Exception:
            pass
        return None

    def _set_cursor_pos(self, x, y):
        """设置物理鼠标光标位置"""
        try:
            if sys.platform == "win32":
                ctypes.windll.user32.SetCursorPos(int(x), int(y))
        except Exception:
            pass

    def _safe_click(self, control, wait_time=0.05):
        """
        安全点击控件（方案 1 核心优化）：
        1. 记录用户当前鼠标位置；
        2. 若处于 ClipCursor 锁定状态，临时解开约束以允许光标移动至目标控件；
        3. 执行瞬发无滑行点击（simulateMove=False 彻底关闭慢速轨迹移动，瞬时到达）；
        4. 在 finally 块中立即将光标原路瞬时复位到用户原位置；
        5. 若处于锁定状态，重新对原位置施加 ClipCursor 约束。
        """
        if control is None or not control.Exists(0, 0):
            return False

        orig_pos = self._get_cursor_pos()
        try:
            # 若正处于鼠标范围锁定状态，必须先临时解除 ClipCursor，否则 Windows 会阻止光标移向目标控件
            if self._is_clipped and sys.platform == "win32":
                ctypes.windll.user32.ClipCursor(None)

            # simulateMove=False 禁用模拟慢速滑行动画，毫秒级瞬发点击
            control.Click(simulateMove=False, waitTime=wait_time)
            return True
        except Exception as e:
            logger.warning(f"安全点击控件异常: {e}")
            return False
        finally:
            if orig_pos:
                self._set_cursor_pos(orig_pos[0], orig_pos[1])
            # 若之前开启了锁定，复位后重新加锁
            if self._is_clipped and orig_pos and sys.platform == "win32":
                try:
                    rect = _RECT(orig_pos[0], orig_pos[1], orig_pos[0] + 1, orig_pos[1] + 1)
                    ctypes.windll.user32.ClipCursor(ctypes.byref(rect))
                except Exception:
                    pass

    def _focus_chat_input(self, chat_edit):
        """
        聚焦聊天输入框（方案 1 核心优化）：
        优先使用 UIA 原生 SetFocus() 避免移动或抢夺物理鼠标光标；
        若 SetFocus 未能成功或抛出异常，降级使用毫秒级坐标快照安全点击兜底。
        """
        try:
            if chat_edit and chat_edit.Exists(0, 0):
                if chat_edit.SetFocus():
                    time.sleep(0.05)
                    focused = auto.GetFocusedControl()
                    if focused and (focused == chat_edit or focused.AutomationId == chat_edit.AutomationId):
                        return True
        except Exception as e:
            logger.debug(f"chat_edit.SetFocus 失败: {e}")

        # 兜底：安全点击（保存原鼠标位置，点击后瞬间复位）
        return self._safe_click(chat_edit)

        # 兜底：安全点击（保存原鼠标位置，点击后瞬间复位）
        return self._safe_click(chat_edit)

    def _ok(self, message):
        return SendResult(True, "OK", message)

    def _fail(self, code, message):
        return SendResult(False, code, message)

    def _is_url(self, value):
        return value.lower().startswith(("http://", "https://"))

    def _resolve_local_file_result(self, file_path):
        """解析并校验本地文件路径。"""
        try:
            expanded_path = os.path.expandvars(os.path.expanduser(file_path.strip().strip('"')))
            path = Path(expanded_path)
            if not path.is_absolute():
                path = Path.cwd() / path
            path = path.resolve()

            if not path.exists():
                return self._fail("LOCAL_FILE_NOT_FOUND", f"本地文件不存在: {path}"), None
            if not path.is_file():
                return self._fail("LOCAL_FILE_INVALID", f"路径不是文件，暂不支持发送目录: {path}"), None

            return self._ok("本地文件路径有效"), str(path)
        except Exception as e:
            return self._fail("LOCAL_FILE_PATH_ERROR", f"解析本地文件路径异常: {str(e)}"), None
        
    def _find_main_window(self):
        """
        跨版本定位微信主窗口。
        兼容策略：
        1. 微信 4.x (Qt 架构)：ClassName='mmui::MainWindow'。
           注意：微信 4.x 窗口 Name 为当前登录账号昵称（如 "LAVA"），绝不能限定 Name="微信"！
        2. 微信 3.x (旧版 Win32/CEF 架构)：ClassName='WeChatMainWndForPC'。
        3. 兜底策略：Name='微信' 顶层窗口。
        """
        try:
            # 策略1: 优先匹配微信 4.x 主窗口类名 mmui::MainWindow（不限定 Name，兼容任意昵称）
            wx = auto.WindowControl(searchDepth=1, ClassName='mmui::MainWindow')
            if wx.Exists(0, 0):
                return wx

            # 策略2: 兼容微信 3.x 旧版主窗口类名 WeChatMainWndForPC
            wx = auto.WindowControl(searchDepth=1, ClassName='WeChatMainWndForPC')
            if wx.Exists(0, 0):
                return wx

            # 策略3: 兼容通过窗口名称 "微信" 查找（部分旧版或未重命名版本）
            wx = auto.WindowControl(searchDepth=1, Name="微信")
            if wx.Exists(0, 0):
                return wx

            return None
        except Exception as e:
            logger.debug(f"定位微信主窗口异常: {e}")
            return None

    def _get_wechat_window_result(self):
        """获取微信窗口对象"""
        try:
            # 确保当前执行线程已关联交互式 default 桌面
            _attach_to_default_desktop()

            # 第一次尝试查找微信窗口
            wx = self._find_main_window()
            if wx and wx.Exists(0, 0):
                return self._ok("已找到微信窗口"), wx
            
            # 第一次找不到，尝试用快捷键唤醒微信窗口（Ctrl+Alt+W 是微信的默认快捷键）
            logger.info("未找到微信窗口，尝试使用快捷键 Ctrl+Alt+W 唤醒微信...")
            auto.SendKeys('{Ctrl}{Alt}w', waitTime=0.1)
            time.sleep(1.0)  # 等待窗口显示
            
            # 第二次尝试查找微信窗口
            wx = self._find_main_window()
            if wx and wx.Exists(0, 0):
                logger.info("成功通过快捷键唤醒微信窗口")
                return self._ok("已唤醒并找到微信窗口"), wx
            else:
                message = "未找到微信窗口，请确认微信 PC 客户端已启动并登录；如使用精简版或 Ghost 系统，请先开启一次 Windows“讲述人”以激活辅助功能。"
                logger.error(message)
                return self._fail("WECHAT_WINDOW_NOT_FOUND", message), None
        except Exception as e:
            message = f"获取微信窗口异常: {str(e)}"
            logger.error(message)
            return self._fail("WECHAT_WINDOW_ERROR", message), None

    def _get_wechat_window(self):
        """获取微信窗口对象，保留给旧调用方使用。"""
        result, wx = self._get_wechat_window_result()
        return wx if result.success else None
    
    def _is_session_selected(self, session_item):
        """
        检查会话项是否已被选中
        
        Args:
            session_item: 会话项控件
            
        Returns:
            bool: 是否已选中
        """
        try:
            # 通过 GetPattern 获取 SelectionItemPattern
            # PatternId.SelectionItemPattern = 10010
            pattern = session_item.GetPattern(10010)
            if pattern and hasattr(pattern, 'IsSelected'):
                is_selected = pattern.IsSelected
                logger.debug(f"会话选中状态: {is_selected}")
                return is_selected
            
            logger.debug("无法获取选中状态")
            return False
            
        except Exception as e:
            logger.debug(f"检查选中状态失败: {str(e)}")
            return False
    
    def _verify_active_chat(self, wx, target_name, max_retries=3, retry_interval=0.2):
        """
        核对当前打开的聊天会话是否与目标名称一致（防串群核心校验机制）。
        
        校验策略：
        1. 微信 4.x 聊天输入框 Name: 经过实测，微信 4.x 中 chat_input_field 的 Name 属性
           严格等于当前联系人/群聊名称（不包含群人数后缀），如 '就买酒 新群'、'文件传输助手'。
        2. 微信 4.x 顶部大标题 current_chat_name_label: 位于窗口顶部中央的 TextControl，
           包含当前打开会话的标题。
        3. 左侧会话项选中态: AutomationId='session_item_{target_name}' 且处于选中状态。
        4. 微信 3.x 兼容: 检查窗口标题或顶层会话名称。
        
        Args:
            wx: 微信主窗口对象
            target_name: 期望的目标联系人/群聊名称
            max_retries: 最大重试轮次（适配界面切换微小渲染时延）
            retry_interval: 重试间隔（秒）
            
        Returns:
            tuple[bool, str]: (是否匹配, 判定描述或不匹配原因)
        """
        target = target_name.strip()
        last_found_name = "未识别"

        for attempt in range(max_retries):
            # 策略1: 优先检查右侧聊天输入框 Name
            try:
                chat_input = self._find_chat_input(wx, max_retries=1)
                if chat_input and chat_input.Name:
                    curr_name = chat_input.Name.strip()
                    last_found_name = curr_name
                    if curr_name == target or target in curr_name:
                        return True, f"输入框目标匹配: '{curr_name}'"
            except Exception:
                pass

            # 策略2: 检查顶部大标题控件（current_chat_name_label）
            try:
                title_label = wx.TextControl(
                    AutomationId='content_view.top_content_view.title_h_view.left_v_view.left_content_v_view.left_ui_.big_title_line_h_view.current_chat_name_label',
                    searchDepth=25
                )
                if title_label.Exists(0, 0) and title_label.Name:
                    curr_title = title_label.Name.strip()
                    last_found_name = curr_title
                    if curr_title == target or target in curr_title:
                        return True, f"顶部标题目标匹配: '{curr_title}'"
            except Exception:
                pass

            # 策略3: 检查左侧会话列表项选中状态
            try:
                session_item = wx.Control(
                    ClassName="mmui::ChatSessionCell",
                    AutomationId=f"session_item_{target}",
                    searchDepth=25
                )
                if session_item.Exists(0, 0) and self._is_session_selected(session_item):
                    return True, f"会话列表项「{target}」处于选中激活状态"
            except Exception:
                pass

            if attempt < max_retries - 1:
                time.sleep(retry_interval)

        return False, f"当前活动窗口为「{last_found_name}」，与目标「{target}」不一致"

    def _activate_search_box(self, wx):
        """
        跨版本激活并定位搜索框输入控件。
        
        兼容策略：
        1. 优先使用全局通用快捷键 Ctrl+F（微信 4.x / 3.x 全版本支持）。
           在微信 4.x 中，即使左侧有“小微AI”按钮或未聚焦态为 GroupControl，Ctrl+F 会瞬间
           实例化并聚焦真实的 EditControl (mmui::XValidatorTextEdit)，彻底杜绝鼠标坐标误触。
        2. 若快捷键未捕获，采用多态控件探测并模拟点击：
           - EditControl(Name='搜索') (微信 3.x 或已聚焦态)
           - EditControl(ClassName='mmui::XValidatorTextEdit')
           - GroupControl(Name='搜索') (微信 4.x 未聚焦常态容器 mmui::XHBoxView)
           - Control(Name='搜索') (兜底)
           点击后再次捕获 EditControl。
        3. 清空搜索框内容（Ctrl+A + Backspace），确保无历史残留。
        
        Returns:
            EditControl 或 None
        """
        try:
            # 激活主窗口
            wx.SetActive()
            time.sleep(0.2)

            search_edit = None

            # 方法1: 尝试通用快捷键 Ctrl+F
            auto.SendKeys('{Ctrl}f')
            time.sleep(0.3)
            focused = auto.GetFocusedControl()
            if focused and focused.ControlTypeName == 'EditControl' and (
                focused.Name == '搜索' or 'Edit' in focused.ClassName or 'Validator' in focused.ClassName
            ):
                search_edit = focused

            # 方法2: 控件探测与点击兜底（兼容 4.x 小微灰度及无快捷键场景）
            if search_edit is None or not search_edit.Exists(0, 0):
                search_ctrl = wx.EditControl(Name='搜索', searchDepth=25)
                if not search_ctrl.Exists(0, 0):
                    search_ctrl = wx.EditControl(ClassName='mmui::XValidatorTextEdit', searchDepth=25)
                if not search_ctrl.Exists(0, 0):
                    search_ctrl = wx.GroupControl(Name='搜索', searchDepth=25)
                if not search_ctrl.Exists(0, 0):
                    search_ctrl = wx.Control(Name='搜索', searchDepth=25)

                if search_ctrl.Exists(0, 0):
                    self._safe_click(search_ctrl)
                    time.sleep(0.3)
                    # 点击激活后再次查找真实编辑框
                    search_edit = wx.EditControl(ClassName='mmui::XValidatorTextEdit', searchDepth=25)
                    if not search_edit.Exists(0, 0):
                        search_edit = wx.EditControl(Name='搜索', searchDepth=25)
                    if not search_edit.Exists(0, 0):
                        search_edit = auto.GetFocusedControl()

            if search_edit and search_edit.Exists(0, 0):
                # 全选并清空搜索框内原有残留文本
                search_edit.SendKeys('{Ctrl}a{Back}')
                time.sleep(0.1)
                return search_edit

            return None
        except Exception as e:
            logger.warning(f"激活微信搜索框异常: {e}")
            return None
    
    def _activate_from_session_list_result(self, contact_name):
        """
        从左侧会话列表直接激活对话（快速方法）
        
        Args:
            contact_name: 联系人名称
            
        Returns:
            SendResult: 激活结果
        """
        try:
            # 获取微信窗口
            result, wx = self._get_wechat_window_result()
            if not result.success:
                return result

            # 激活主窗口，确保获得输入与鼠标焦点
            wx.SetActive()
            time.sleep(0.2)
            
            # 查找会话列表中的联系人
            # AutomationId 格式: session_item_[联系人名]
            automation_id = f"session_item_{contact_name}"
            
            # 查找会话项（微信 4.x 实际深度探测为 14，适当加深至 25 避免层级微调失效）
            session_item = wx.Control(
                ClassName="mmui::ChatSessionCell",
                AutomationId=automation_id,
                searchDepth=25
            )
            
            if session_item.Exists(0, 0):
                # 检查是否已经选中
                if not self._is_session_selected(session_item):
                    logger.info(f"点击激活会话: {contact_name}")
                    self._safe_click(session_item)
                    time.sleep(0.3)
                else:
                    logger.info(f"会话 '{contact_name}' 已经处于选中状态")
                
                # 双重校验：核验当前活动会话是否确实是目标联系人/群
                verified, reason = self._verify_active_chat(wx, contact_name)
                if verified:
                    logger.info(f"从会话列表成功激活并核验联系人: {contact_name} ({reason})")
                    return self._ok(f"已从会话列表激活联系人/群组: {contact_name}")
                else:
                    logger.warning(f"会话列表点击后核验未通过: {reason}")
                    return self._fail("SESSION_VERIFY_FAILED", f"会话列表选中后核对失败: {reason}")
            else:
                logger.debug(f"会话列表中未找到 {contact_name}，将使用搜索方式")
                return self._fail("SESSION_NOT_FOUND", f"会话列表中未找到: {contact_name}")
                
        except Exception as e:
            message = f"从会话列表激活失败: {str(e)}"
            logger.debug(message)
            return self._fail("SESSION_ACTIVATE_ERROR", message)

    def _activate_from_session_list(self, contact_name):
        """从会话列表激活对话，保留布尔返回给旧调用方使用。"""
        return self._activate_from_session_list_result(contact_name).success
    
    @_with_input_lock
    def search_contact_result(self, contact_name):
        """
        搜索联系人（双重策略：优先从会话列表激活，找不到再通过搜索框精准匹配）
        
        Args:
            contact_name: 联系人名称
            
        Returns:
            SendResult: 搜索与定位结果
        """
        try:
            target = contact_name.strip()

            # 策略1: 优先尝试从会话列表直接激活（更快）
            session_result = self._activate_from_session_list_result(target)
            if session_result.success:
                return session_result
            if session_result.code in ("WECHAT_WINDOW_NOT_FOUND", "WECHAT_WINDOW_ERROR"):
                return session_result
            
            # 策略2: 降级使用搜索框进行精准匹配（兜底方案）
            logger.info(f"使用搜索框查找联系人/群组: {target}")
            
            # 获取微信窗口
            result, wx = self._get_wechat_window_result()
            if not result.success:
                return result
            
            # 激活搜索框（兼容小微灰度、GroupControl/EditControl 及全版本 Ctrl+F）
            search_box = self._activate_search_box(wx)
            if not search_box or not search_box.Exists(0, 0):
                message = "未找到微信搜索框，可能是微信版本 UI 结构变化、窗口未完全加载或系统 UI 自动化不可用。"
                logger.error(message)
                return self._fail("SEARCH_BOX_NOT_FOUND", message)
            
            # 使用剪贴板粘贴方式输入搜索内容（更快且避免特殊字符问题）
            if self._set_clipboard_text(target):
                search_box.SendKeys('{Ctrl}v')
                time.sleep(0.8)
            else:
                # 剪贴板设置失败：含中文/多字节字符直接报错，避免 SendKeys 漏字搜不到联系人
                if any(ord(c) > 127 for c in target):
                    error_message = (
                        f"搜索时剪贴板设置失败，且联系人名称「{target}」包含中文或多字节字符，"
                        "SendKeys 漏字会导致搜不到目标。请检查 pyperclip/系统剪贴板是否正常。"
                    )
                    logger.error(error_message)
                    return self._fail("CLIPBOARD_TEXT_FAILED", error_message)
                logger.warning("搜索时剪贴板设置失败，使用 SendKeys 方式（仅推荐英文）")
                escaped_name = target.replace('{', '{{').replace('}', '}}')
                search_box.SendKeys(escaped_name, interval=0.05)
                time.sleep(0.8)
            
            # 等待搜索结果弹窗（微信 4.x 的 SearchContentPopover 浮层）
            popover = wx.Control(ClassName='mmui::SearchContentPopover', searchDepth=2)
            if not popover.Exists(0, 0):
                popover = auto.WindowControl(ClassName='mmui::SearchContentPopover')

            # 允许匹配的业务分组白名单：最常使用、联系人、群聊、功能（如文件传输助手）、服务号
            allowed_categories = {'最常使用', '联系人', '群聊', '功能', '服务号'}
            # 严禁匹配的分组黑名单：绝不能匹配网络搜索、聊天记录或添加外部账号
            forbidden_categories = {'聊天记录', '搜索网络结果', '网络查找微信号：', '收藏'}

            matched_item = None

            if popover.Exists(0, 0):
                search_list = popover.ListControl(AutomationId='search_list')
                if search_list.Exists(0, 0):
                    current_category = None
                    for item in search_list.GetChildren():
                        name = item.Name.strip() if item.Name else ""
                        aid = item.AutomationId or ""
                        cname = item.ClassName or ""

                        # 判断是否为分组标题行
                        if not aid and (name in allowed_categories or name in forbidden_categories):
                            current_category = name
                            continue

                        # 过滤非目标类型：聊天记录行、查看全部、空行等
                        if cname == 'mmui::SearchMessageCellView' or name.startswith('查看全部') or not name:
                            continue

                        # 仅在白名单分组下寻找完全匹配项
                        if current_category in allowed_categories:
                            is_exact = (
                                name == target or
                                aid == f"search_item_{target}" or
                                aid == f"search_item_function_{target}" or
                                aid == f"search_item_contact_{target}" or
                                aid == f"search_item_group_{target}" or
                                aid.endswith(f"_{target}")
                            )
                            if is_exact:
                                matched_item = item
                                logger.info(f"在搜索分组「{current_category}」中找到完全匹配项: {name} (aid={aid})")
                                break

            # 分支 1: 在 Popover 中找到明确的完全匹配项，直接点击该项
            if matched_item:
                self._safe_click(matched_item)
                time.sleep(0.5)
            elif popover.Exists(0, 0):
                # 分支 2: Popover 已渲染，但在合法业务分组中没有完全符合的选项
                # （可能仅有“搜索网络结果”或“聊天记录”等），严禁按回车！立即按 Esc 安全关闭，防止串群或误发！
                logger.warning(f"搜索「{target}」未在联系人/群聊/功能分组中找到完全匹配项，安全退出搜索浮层")
                auto.SendKeys('{Esc}')
                time.sleep(0.2)
                message = f"未在微信联系人/群聊/功能分组中找到与「{target}」完全匹配的会话，为防止误发已安全取消操作。"
                return self._fail("CONTACT_NOT_FOUND", message)
            else:
                # 分支 3: 未检测到 Popover 弹窗（微信 3.x 旧版架构兜底），按 Enter 尝试确认
                logger.info("未检测到 Popover 浮层，使用 Enter 确认旧版搜索...")
                search_box.SendKeys('{Enter}')
                time.sleep(0.8)

            # 核心防串群拦截：核验切换后的当前聊天窗口是否确实为需要发送的目标
            verified, reason = self._verify_active_chat(wx, target)
            if not verified:
                message = f"会话激活核验未通过: {reason}。已拦截并拒绝进入发送流程，以防向错误群聊发送内容！"
                logger.error(message)
                return self._fail("TARGET_VERIFY_FAILED", message)

            logger.info(f"搜索并核验通过，当前处于目标会话: {target} ({reason})")
            return self._ok(f"已搜索并确认进入会话: {target}")
            
        except Exception as e:
            message = f"搜索联系人 '{contact_name}' 异常: {str(e)}"
            logger.error(message)
            return self._fail("CONTACT_SEARCH_ERROR", message)

    def search_contact(self, contact_name):
        """搜索联系人，保留布尔返回给旧调用方使用。"""
        return self.search_contact_result(contact_name).success
    
    def _find_chat_input(self, wx, max_retries=3, retry_interval=0.15):
        """
        定位聊天输入框（全面兼容微信 4.x Qt 架构与微信 3.x 旧版架构）。

        微信 4.x (Qt 架构):
            AutomationId = "chat_input_field"
            ClassName    = "mmui::ChatInputField"
            实测深度为 17 层。在切换会话时，右侧聊天面板渲染有轻微毫秒级延迟，因此提供短暂重试。
        微信 3.x (旧版 Win32/CEF 架构):
            部分版本为 EditControl 且 Name='输入' 或基础 EditControl。
        """
        for attempt in range(max_retries):
            try:
                # 策略1: 优先 AutomationId + ClassName 精确匹配（微信 4.x 推荐，searchDepth=25）
                chat_edit = wx.EditControl(
                    AutomationId="chat_input_field",
                    ClassName="mmui::ChatInputField",
                    searchDepth=25,
                )
                if chat_edit.Exists(0, 0):
                    return chat_edit

                # 策略2: 仅通过 AutomationId 匹配
                chat_edit = wx.EditControl(AutomationId="chat_input_field", searchDepth=25)
                if chat_edit.Exists(0, 0):
                    logger.warning("仅以 AutomationId 匹配到聊天输入框，ClassName 可能已变化")
                    return chat_edit

                # 策略3: 仅通过 ClassName 匹配
                chat_edit = wx.EditControl(ClassName="mmui::ChatInputField", searchDepth=25)
                if chat_edit.Exists(0, 0):
                    logger.warning("仅以 ClassName 匹配到聊天输入框，AutomationId 可能已变化")
                    return chat_edit

                # 策略4: 兼容微信 3.x 旧版输入框 (Name='输入')
                chat_edit = wx.EditControl(Name="输入", searchDepth=25)
                if chat_edit.Exists(0, 0):
                    logger.info("匹配到旧版微信聊天输入框 (Name='输入')")
                    return chat_edit

            except Exception as e:
                logger.debug(f"查找聊天输入框异常 (尝试 {attempt + 1}/{max_retries}): {str(e)}")

            if attempt < max_retries - 1:
                time.sleep(retry_interval)

        logger.error("未找到聊天输入框，可能未成功进入目标会话或微信 UI 结构已变化")
        return None

    def _set_clipboard_text(self, text, max_retries=3):
        """
        安全地设置剪贴板文本（使用 pyperclip，正确支持 Unicode/中文）
        验证方式改为：先获取→清空→设置→再获取→比对
        
        Args:
            text: 要设置的文本内容
            max_retries: 最大重试次数
            
        Returns:
            bool: 操作是否成功
        """
        for attempt in range(max_retries):
            try:
                pyperclip.copy(text)
                time.sleep(0.05)
                clipboard_content = pyperclip.paste()
                if clipboard_content == text:
                    return True
                else:
                    logger.warning(f"剪贴板内容验证失败，重试中... (尝试 {attempt + 1}/{max_retries}) 期望长度={len(text)} 实际长度={len(clipboard_content)}")
            except Exception as e:
                logger.warning(f"设置剪贴板失败: {e}，重试中... (尝试 {attempt + 1}/{max_retries})")
            time.sleep(0.1)
        
        logger.error("设置剪贴板文本失败，已达最大重试次数")
        return False
    
    @_with_input_lock
    def send_message_result(self, message, expected_contact=None):
        """
        发送消息（使用剪贴板粘贴方式，解决 SendKeys 特殊字符问题）
        
        Args:
            message: 要发送的消息内容（支持换行符 \n）
            expected_contact: 可选，期望发送的目标会话名称。若传入，将在物理发送前进行二次防串群终极拦截校验。
            
        Returns:
            SendResult: 发送结果
        """
        try:
            # 获取微信窗口
            result, wx = self._get_wechat_window_result()
            if not result.success:
                return result

            # 激活主窗口，确保获得输入与鼠标焦点
            wx.SetActive()
            time.sleep(0.2)
            
            # 若指定了目标联系人，执行物理发送前的二次核验拦截
            if expected_contact:
                verified, reason = self._verify_active_chat(wx, expected_contact)
                if not verified:
                    abort_msg = f"防串群终极拦截: {reason}，已紧急终止发送！"
                    logger.error(abort_msg)
                    return self._fail("TARGET_MISMATCH_ABORT", abort_msg)

            # 通过 AutomationId="chat_input_field" 精确定位聊天输入框
            # 旧版 foundIndex=1 在打开公众号文章/视频面板后会错位
            chat_edit = self._find_chat_input(wx)
            if chat_edit is None:
                message = "未找到聊天输入框，可能未成功进入目标会话或微信 UI 结构已变化。"
                logger.error(message)
                return self._fail("CHAT_INPUT_NOT_FOUND", message)
            
            # 聚焦输入框获取焦点（优先 SetFocus 避免移动鼠标；兜底安全点击）
            self._focus_chat_input(chat_edit)
            time.sleep(0.2)
            
            # 使用剪贴板粘贴方式发送消息
            # 这种方式比 SendKeys 快得多，且不会出现特殊字符（如【】￥等）被误解析的问题
            if not self._set_clipboard_text(message):
                # 剪贴板设置失败，回退到 SendKeys 方式（作为兜底）
                # 注意：SendKeys 对中文支持较差，仅建议用于纯英文短消息
                if any(ord(c) > 127 for c in message):
                    error_message = "剪贴板设置失败，且消息包含中文或多字节字符，无法可靠发送。请检查 pyperclip/系统剪贴板是否正常。"
                    logger.error(error_message)
                    return self._fail("CLIPBOARD_TEXT_FAILED", error_message)
                logger.warning("剪贴板设置失败，尝试使用 SendKeys 方式（仅推荐英文）")
                # 转义特殊字符，避免被 SendKeys 误解析
                escaped_message = message.replace('{', '{{').replace('}', '}}')
                formatted_message = escaped_message.replace('\n', '{Shift}{Enter}')
                # interval 至少 0.05s，中文/多字节字符容易丢失
                chat_edit.SendKeys(formatted_message + '{Enter}', interval=0.05)
            else:
                # 粘贴消息（Ctrl+V）
                chat_edit.SendKeys('{Ctrl}v')
                time.sleep(0.2)
                
                # 发送消息（Enter）
                chat_edit.SendKeys('{Enter}')
            
            time.sleep(0.3)
            
            # 日志中显示原始消息（包含换行符）
            log_preview = message.replace('\n', '\\n')[:50]
            logger.info(f"成功发送消息: {log_preview}...")
            return self._ok("文本消息发送成功")
            
        except Exception as e:
            message = f"发送文本消息异常: {str(e)}"
            logger.error(message, exc_info=True)
            return self._fail("SEND_TEXT_ERROR", message)

    def send_message(self, message):
        """发送文本消息，保留布尔返回给旧调用方使用。"""
        return self.send_message_result(message).success
    
    def _download_image_result(self, url, max_retries=3):
        """
        从 URL 下载图片到缓存文件（使用 MD5 作为文件名避免重复下载，带重试机制）
        
        Args:
            url: 图片的 URL
            max_retries: 最大重试次数
            
        Returns:
            str: 缓存文件路径，失败返回 None
        """
        try:
            # 计算 URL 的 MD5 作为文件名
            url_md5 = hashlib.md5(url.encode('utf-8')).hexdigest()
            cache_path = os.path.join(self.cache_dir, f"{url_md5}.png")
            
            # 检查缓存文件是否已存在且有效
            if os.path.exists(cache_path):
                # 验证缓存文件是否有效（大小大于0）
                if os.path.getsize(cache_path) > 0:
                    logger.info(f"使用缓存图片: {cache_path}")
                    return self._ok("已使用缓存图片"), cache_path
                else:
                    # 缓存文件无效，删除后重新下载
                    logger.warning(f"缓存文件无效，删除后重新下载: {cache_path}")
                    os.remove(cache_path)
            
            logger.info(f"开始下载图片: {url}")
            
            # 带重试的下载
            last_error = None
            for attempt in range(max_retries):
                try:
                    # 下载图片
                    response = requests.get(url, timeout=30, headers={
                        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                    })
                    response.raise_for_status()
                    
                    # 判断内容类型（某些服务器可能不返回正确的 content-type）
                    content_type = response.headers.get('content-type', '')
                    if content_type and not content_type.startswith('image/'):
                        # 如果服务器明确返回非图片类型，则报错
                        if 'text/' in content_type or 'application/json' in content_type:
                            logger.error(f"URL 返回的不是图片类型: {content_type}")
                            return self._fail("IMAGE_URL_NOT_IMAGE", f"URL 返回的不是图片类型: {content_type}"), None
                    
                    # 尝试打开图片验证有效性
                    image = Image.open(BytesIO(response.content))
                    image.verify()  # 验证图片完整性
                    
                    # 重新打开图片（verify 后需要重新打开）
                    image = Image.open(BytesIO(response.content))
                    
                    # 保存图片到缓存目录
                    image.save(cache_path, 'PNG')
                    logger.info(f"图片已下载并缓存到: {cache_path}")
                    
                    return self._ok("图片下载成功"), cache_path
                    
                except requests.RequestException as e:
                    last_error = e
                    if attempt < max_retries - 1:
                        logger.warning(f"下载图片失败: {e}，重试中... (尝试 {attempt + 1}/{max_retries})")
                        time.sleep(1)  # 等待一秒后重试
                    continue
                except Exception as e:
                    last_error = e
                    logger.error(f"处理图片失败: {e}")
                    break
            
            message = f"下载图片失败，已达最大重试次数: {last_error}"
            logger.error(message)
            return self._fail("IMAGE_DOWNLOAD_FAILED", message), None
            
        except Exception as e:
            message = f"下载图片过程中发生异常: {str(e)}"
            logger.error(message, exc_info=True)
            return self._fail("IMAGE_DOWNLOAD_ERROR", message), None

    def _download_image(self, url, max_retries=3):
        """下载图片，保留原返回形式给旧调用方使用。"""
        result, cache_path = self._download_image_result(url, max_retries=max_retries)
        return cache_path if result.success else None
    
    def _copy_image_to_clipboard_result(self, image_path, max_retries=3):
        """
        将图片复制到剪贴板（带重试机制和安全的资源释放）
        
        Args:
            image_path: 图片文件路径
            max_retries: 最大重试次数
            
        Returns:
            bool: 操作是否成功
        """
        for attempt in range(max_retries):
            clipboard_opened = False
            try:
                # 打开图片
                image = Image.open(image_path)
                
                # 转换为 BMP 格式（Windows 剪贴板需要）
                output = BytesIO()
                image.convert('RGB').save(output, 'BMP')
                data = output.getvalue()[14:]  # BMP 文件头是 14 字节，剪贴板不需要
                output.close()
                
                # 复制到剪贴板（使用标志位确保正确关闭）
                win32clipboard.OpenClipboard()
                clipboard_opened = True
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
                win32clipboard.CloseClipboard()
                clipboard_opened = False
                
                logger.info("图片已复制到剪贴板")
                return self._ok("图片已复制到剪贴板")
                
            except Exception as e:
                # 确保剪贴板被关闭
                if clipboard_opened:
                    try:
                        win32clipboard.CloseClipboard()
                    except Exception:
                        pass
                
                if attempt < max_retries - 1:
                    logger.warning(f"复制图片到剪贴板失败: {e}，重试中... (尝试 {attempt + 1}/{max_retries})")
                    time.sleep(0.2)
                else:
                    logger.error(f"复制图片到剪贴板失败，已达最大重试次数: {str(e)}")
        
        return self._fail("IMAGE_CLIPBOARD_FAILED", "复制图片到剪贴板失败，请检查图片文件、系统剪贴板或微信窗口状态。")

    def _copy_image_to_clipboard(self, image_path, max_retries=3):
        """复制图片到剪贴板，保留布尔返回给旧调用方使用。"""
        return self._copy_image_to_clipboard_result(image_path, max_retries=max_retries).success

    def _copy_file_to_clipboard_result(self, file_path, max_retries=3):
        """
        将本地文件复制到剪贴板，用于在微信输入框中粘贴发送文件。
        Windows 文件剪贴板使用 CF_HDROP 格式。
        """
        path_result, resolved_path = self._resolve_local_file_result(file_path)
        if not path_result.success:
            return path_result

        # DROPFILES: pFiles=20, pt=(0,0), fNC=0, fWide=1
        dropfiles_header = struct.pack("IiiII", 20, 0, 0, 0, 1)
        file_list = (resolved_path + "\0\0").encode("utf-16le")
        clipboard_data = dropfiles_header + file_list

        for attempt in range(max_retries):
            clipboard_opened = False
            try:
                win32clipboard.OpenClipboard()
                clipboard_opened = True
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32con.CF_HDROP, clipboard_data)
                win32clipboard.CloseClipboard()
                clipboard_opened = False
                logger.info(f"文件已复制到剪贴板: {resolved_path}")
                return self._ok(f"文件已复制到剪贴板: {resolved_path}")
            except Exception as e:
                if clipboard_opened:
                    try:
                        win32clipboard.CloseClipboard()
                    except Exception:
                        pass
                if attempt < max_retries - 1:
                    logger.warning(f"复制文件到剪贴板失败: {e}，重试中... (尝试 {attempt + 1}/{max_retries})")
                    time.sleep(0.2)
                else:
                    logger.error(f"复制文件到剪贴板失败，已达最大重试次数: {str(e)}")

        return self._fail("FILE_CLIPBOARD_FAILED", f"复制文件到剪贴板失败: {resolved_path}")

    def _copy_file_to_clipboard(self, file_path, max_retries=3):
        """复制本地文件到剪贴板，保留布尔返回给旧调用方使用。"""
        return self._copy_file_to_clipboard_result(file_path, max_retries=max_retries).success
    
    @_with_input_lock
    def send_picture_result(self, image_url, expected_contact=None):
        """
        发送图片（URL 图片下载后发送；本地路径直接作为文件粘贴发送）
        
        Args:
            image_url: 图片的 URL 或本地文件路径
            expected_contact: 可选，期望发送的目标会话名称。若传入，将在物理发送前进行二次防串群终极拦截校验。
            
        Returns:
            SendResult: 发送结果
        """
        try:
            if not self._is_url(image_url):
                return self.send_file_result(image_url, expected_contact=expected_contact)

            # 下载图片（或使用缓存）
            download_result, cache_file = self._download_image_result(image_url)
            if not download_result.success:
                return download_result
            
            # 复制图片到剪贴板
            clipboard_result = self._copy_image_to_clipboard_result(cache_file)
            if not clipboard_result.success:
                return clipboard_result
            
            # 获取微信窗口
            result, wx = self._get_wechat_window_result()
            if not result.success:
                return result
            
            # 激活窗口确保焦点正确
            wx.SetActive()
            time.sleep(0.2)
            
            # 防串群终极拦截校验
            if expected_contact:
                verified, reason = self._verify_active_chat(wx, expected_contact)
                if not verified:
                    abort_msg = f"防串群终极拦截: {reason}，已紧急终止图片发送！"
                    logger.error(abort_msg)
                    return self._fail("TARGET_MISMATCH_ABORT", abort_msg)

            # 通过 AutomationId="chat_input_field" 精确定位聊天输入框
            chat_edit = self._find_chat_input(wx)
            if chat_edit is None:
                message = "未找到聊天输入框，可能未成功进入目标会话或微信 UI 结构已变化。"
                logger.error(message)
                return self._fail("CHAT_INPUT_NOT_FOUND", message)
            
            # 聚焦输入框获取焦点（优先 SetFocus 避免移动鼠标；兜底安全点击）
            self._focus_chat_input(chat_edit)
            time.sleep(0.2)
            
            # 粘贴图片（Ctrl+V）
            chat_edit.SendKeys('{Ctrl}v')
            time.sleep(0.5)
            
            # 发送（Enter）
            chat_edit.SendKeys('{Enter}')
            time.sleep(0.3)
            
            logger.info(f"成功发送图片: {image_url}")
            return self._ok("图片消息发送成功")
            
        except Exception as e:
            message = f"发送图片异常: {str(e)}"
            logger.error(message, exc_info=True)
            return self._fail("SEND_IMAGE_ERROR", message)

    def send_picture(self, image_url):
        """发送图片，保留布尔返回给旧调用方使用。"""
        return self.send_picture_result(image_url).success

    @_with_input_lock
    def send_file_result(self, file_path, expected_contact=None):
        """
        发送本地文件或本地图片：复制文件到剪贴板，粘贴到微信输入框并回车发送。
        """
        try:
            clipboard_result = self._copy_file_to_clipboard_result(file_path)
            if not clipboard_result.success:
                return clipboard_result

            result, wx = self._get_wechat_window_result()
            if not result.success:
                return result

            wx.SetActive()
            time.sleep(0.2)

            # 防串群终极拦截校验
            if expected_contact:
                verified, reason = self._verify_active_chat(wx, expected_contact)
                if not verified:
                    abort_msg = f"防串群终极拦截: {reason}，已紧急终止文件发送！"
                    logger.error(abort_msg)
                    return self._fail("TARGET_MISMATCH_ABORT", abort_msg)

            # 通过 AutomationId="chat_input_field" 精确定位聊天输入框
            chat_edit = self._find_chat_input(wx)
            if chat_edit is None:
                message = "未找到聊天输入框，可能未成功进入目标会话或微信 UI 结构已变化。"
                logger.error(message)
                return self._fail("CHAT_INPUT_NOT_FOUND", message)

            # 聚焦输入框获取焦点（优先 SetFocus 避免移动鼠标；兜底安全点击）
            self._focus_chat_input(chat_edit)
            time.sleep(0.2)
            chat_edit.SendKeys('{Ctrl}v')
            time.sleep(0.8)
            chat_edit.SendKeys('{Enter}')
            time.sleep(0.5)

            _, resolved_path = self._resolve_local_file_result(file_path)
            logger.info(f"成功发送文件: {resolved_path}")
            return self._ok(f"发送成功: 已发送本地文件「{resolved_path}」。")
        except Exception as e:
            message = f"发送本地文件异常: {str(e)}"
            logger.error(message, exc_info=True)
            return self._fail("SEND_FILE_ERROR", message)

    def send_file(self, file_path):
        """发送本地文件，保留布尔返回给旧调用方使用。"""
        return self.send_file_result(file_path).success
    
    @_with_input_lock
    def search_and_send_result(self, contact_name, message):
        """
        搜索联系人并发送消息（组合操作，内置严格防串群校验）
        
        Args:
            contact_name: 联系人名称
            message: 要发送的消息内容
            
        Returns:
            SendResult: 操作结果
        """
        logger.info(f"开始向 '{contact_name}' 发送消息")
        
        # 搜索联系人并严格核验当前活动窗口
        search_result = self.search_contact_result(contact_name)
        if not search_result.success:
            logger.warning(f"跳过向 '{contact_name}' 发送消息（搜索失败: {search_result.message}）")
            return search_result
        
        # 发送消息（传入 expected_contact 进行物理发送前的二次拦截拦截）
        send_result = self.send_message_result(message, expected_contact=contact_name)
        if not send_result.success:
            logger.warning(f"向 '{contact_name}' 发送消息失败: {send_result.message}")
            return send_result
        
        logger.info(f"成功向 '{contact_name}' 发送消息")
        return self._ok(f"发送成功: 已向「{contact_name}」发送文本消息。")

    def search_and_send(self, contact_name, message):
        """搜索联系人并发送文本，保留布尔返回给旧调用方使用。"""
        return self.search_and_send_result(contact_name, message).success
    
    @_with_input_lock
    def search_and_send_picture_result(self, contact_name, image_url):
        """
        搜索联系人并发送图片（组合操作，内置严格防串群校验）
        
        Args:
            contact_name: 联系人名称
            image_url: 图片的 URL 或本地文件路径
            
        Returns:
            SendResult: 操作结果
        """
        logger.info(f"开始向 '{contact_name}' 发送图片")
        
        # 搜索联系人并严格核验当前活动窗口
        search_result = self.search_contact_result(contact_name)
        if not search_result.success:
            logger.warning(f"跳过向 '{contact_name}' 发送图片（搜索失败: {search_result.message}）")
            return search_result
        
        # 发送图片（传入 expected_contact 进行物理发送前的二次拦截校验）
        send_result = self.send_picture_result(image_url, expected_contact=contact_name)
        if not send_result.success:
            logger.warning(f"向 '{contact_name}' 发送图片失败: {send_result.message}")
            return send_result
        
        logger.info(f"成功向 '{contact_name}' 发送图片")
        if not self._is_url(image_url):
            _, resolved_path = self._resolve_local_file_result(image_url)
            return self._ok(f"发送成功: 已向「{contact_name}」发送本地文件「{resolved_path}」。")
        return self._ok(f"发送成功: 已向「{contact_name}」发送图片消息。")

    def search_and_send_picture(self, contact_name, image_url):
        """搜索联系人并发送图片，保留布尔返回给旧调用方使用。"""
        return self.search_and_send_picture_result(contact_name, image_url).success

    @_with_input_lock
    def search_and_send_file_result(self, contact_name, file_path):
        """
        搜索联系人并发送本地文件（组合操作，内置严格防串群校验）
        """
        logger.info(f"开始向 '{contact_name}' 发送文件")

        # 搜索联系人并严格核验当前活动窗口
        search_result = self.search_contact_result(contact_name)
        if not search_result.success:
            logger.warning(f"跳过向 '{contact_name}' 发送文件（搜索失败: {search_result.message}）")
            return search_result

        # 发送文件（传入 expected_contact 进行物理发送前的二次拦截校验）
        send_result = self.send_file_result(file_path, expected_contact=contact_name)
        if not send_result.success:
            logger.warning(f"向 '{contact_name}' 发送文件失败: {send_result.message}")
            return send_result

        logger.info(f"成功向 '{contact_name}' 发送文件")
        _, resolved_path = self._resolve_local_file_result(file_path)
        return self._ok(f"发送成功: 已向「{contact_name}」发送本地文件「{resolved_path}」。")

    def search_and_send_file(self, contact_name, file_path):
        """搜索联系人并发送本地文件，保留布尔返回给旧调用方使用。"""
        return self.search_and_send_file_result(contact_name, file_path).success


# 测试代码
if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # 创建控制器并测试
    controller = WeChatController()
    controller.search_and_send("线报转发", "这是一条测试消息")

