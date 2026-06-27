"""
微信控制器模块
封装微信操作，提供搜索联系人和发送消息的功能
"""
import uiautomation as auto
import time
import logging
import requests
import os
import struct
import tempfile
from dataclasses import dataclass
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

# 配置日志
logger = logging.getLogger(__name__)


@dataclass
class SendResult:
    """一次发送操作的结果，供 CLI/Skill 输出明确反馈。"""
    success: bool
    code: str
    message: str


class WeChatController:
    """微信控制器类，用于自动化控制微信发送消息"""
    
    def __init__(self):
        """初始化微信控制器"""
        self.wx = None
        # 创建图片缓存目录
        self.cache_dir = os.path.join(tempfile.gettempdir(), 'wechat_image_cache')
        if not os.path.exists(self.cache_dir):
            os.makedirs(self.cache_dir)
            logger.info(f"创建图片缓存目录: {self.cache_dir}")

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
        
    def _get_wechat_window_result(self):
        """获取微信窗口对象"""
        try:
            # 第一次尝试查找微信窗口
            wx = auto.WindowControl(searchDepth=1, Name="微信", ClassName='Qt51514QWindowIcon')
            if wx.Exists(0, 0):
                return self._ok("已找到微信窗口"), wx
            
            # 第一次找不到，尝试用快捷键唤醒微信窗口（Ctrl+Alt+W 是微信的默认快捷键）
            logger.info("未找到微信窗口，尝试使用快捷键 Ctrl+Alt+W 唤醒微信...")
            auto.SendKeys('{Ctrl}{Alt}w', waitTime=0.1)
            time.sleep(2.0)  # 等待窗口显示
            
            # 第二次尝试查找微信窗口
            wx = auto.WindowControl(searchDepth=1, Name="微信", ClassName='Qt51514QWindowIcon')
            if wx.Exists(0, 0):
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
    
    def _activate_from_session_list_result(self, contact_name):
        """
        从左侧会话列表直接激活对话（快速方法）
        
        Args:
            contact_name: 联系人名称
            
        Returns:
            bool: 是否成功激活
        """
        try:
            # 获取微信窗口
            result, wx = self._get_wechat_window_result()
            if not result.success:
                return result
            
            # 查找会话列表中的联系人
            # AutomationId 格式: session_item_[联系人名]
            automation_id = f"session_item_{contact_name}"
            
            # 查找会话项
            session_item = wx.Control(
                ClassName="mmui::ChatSessionCell",
                AutomationId=automation_id,
                searchDepth=15
            )
            
            if session_item.Exists(0, 0):
                # 检查是否已经选中
                if self._is_session_selected(session_item):
                    logger.info(f"会话 '{contact_name}' 已经处于选中状态，无需点击")
                    return self._ok(f"已从会话列表选中联系人/群组: {contact_name}")
                
                # 未选中，需要点击激活
                logger.info(f"点击激活会话: {contact_name}")
                session_item.Click()
                time.sleep(0.3)
                
                # 验证是否选中成功
                if self._is_session_selected(session_item):
                    logger.info(f"从会话列表成功激活联系人: {contact_name}")
                    return self._ok(f"已从会话列表激活联系人/群组: {contact_name}")
                else:
                    logger.warning(f"点击后会话 '{contact_name}' 未被选中")
                    return self._fail("SESSION_NOT_SELECTED", f"找到会话但未能选中: {contact_name}")
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
    
    def search_contact_result(self, contact_name):
        """
        搜索联系人（双重策略：优先从会话列表激活，找不到再搜索）
        
        Args:
            contact_name: 联系人名称
            
        Returns:
            bool: 搜索是否成功
        """
        try:
            # 策略1: 优先尝试从会话列表直接激活（更快）
            session_result = self._activate_from_session_list_result(contact_name)
            if session_result.success:
                return session_result
            if session_result.code == "WECHAT_WINDOW_NOT_FOUND" or session_result.code == "WECHAT_WINDOW_ERROR":
                return session_result
            
            # 策略2: 降级使用搜索框（兜底方案）
            logger.info(f"使用搜索框查找联系人: {contact_name}")
            
            # 获取微信窗口
            result, wx = self._get_wechat_window_result()
            if not result.success:
                return result
            
            # 激活窗口
            wx.SetActive()
            time.sleep(0.5)
            
            # 查找搜索框
            search_box = wx.EditControl(Name='搜索')
            if not search_box.Exists(0, 0):
                message = "未找到微信搜索框，可能是微信版本 UI 结构变化、窗口未完全加载或系统 UI 自动化不可用。"
                logger.error(message)
                return self._fail("SEARCH_BOX_NOT_FOUND", message)
            
            # 点击搜索框
            search_box.Click()
            time.sleep(0.3)
            
            # 使用剪贴板粘贴方式输入搜索内容（更快且避免特殊字符问题）
            if self._set_clipboard_text(contact_name):
                search_box.SendKeys('{Ctrl}v')
                time.sleep(0.2)
            else:
                # 剪贴板设置失败：含中文/多字节字符直接报错，避免 SendKeys 漏字搜不到联系人
                if any(ord(c) > 127 for c in contact_name):
                    error_message = (
                        f"搜索时剪贴板设置失败，且联系人名称「{contact_name}」包含中文或多字节字符，"
                        "SendKeys 漏字会导致搜不到目标。请检查 pyperclip/系统剪贴板是否正常。"
                    )
                    logger.error(error_message)
                    return self._fail("CLIPBOARD_TEXT_FAILED", error_message)
                # 纯 ASCII 名称才允许 SendKeys 兜底，interval 至少 0.05s 防字符丢失
                logger.warning("搜索时剪贴板设置失败，使用 SendKeys 方式（仅推荐英文）")
                escaped_name = contact_name.replace('{', '{{').replace('}', '}}')
                search_box.SendKeys(escaped_name, interval=0.05)
                time.sleep(0.2)
            
            # 按 Enter 确认搜索
            search_box.SendKeys('{Enter}')
            time.sleep(0.8)
            
            # 搜索后，尝试验证会话是否已选中
            automation_id = f"session_item_{contact_name}"
            session_item = wx.Control(
                ClassName="mmui::ChatSessionCell",
                AutomationId=automation_id,
                searchDepth=15
            )
            
            if session_item.Exists(0, 0):
                if self._is_session_selected(session_item):
                    logger.info(f"搜索后确认会话 '{contact_name}' 已选中")
                    return self._ok(f"已搜索并选中联系人/群组: {contact_name}")
                else:
                    message = f"搜索后找到会话但未能选中: {contact_name}"
                    logger.warning(message)
                    return self._fail("CONTACT_NOT_SELECTED", message)
            else:
                message = f"未找到联系人或群组: {contact_name}。请确认名称与微信备注/群名完全一致。"
                logger.warning(message)
                return self._fail("CONTACT_NOT_FOUND", message)
            
        except Exception as e:
            message = f"搜索联系人 '{contact_name}' 异常: {str(e)}"
            logger.error(message)
            return self._fail("CONTACT_SEARCH_ERROR", message)

    def search_contact(self, contact_name):
        """搜索联系人，保留布尔返回给旧调用方使用。"""
        return self.search_contact_result(contact_name).success
    
    def _find_chat_input(self, wx):
        """
        定位聊天输入框（新版微信兼容）。

        旧版用 wx.EditControl(foundIndex=1) 依赖深度优先第二个 EditControl，
        但新版微信打开公众号文章/视频后右侧会出现独立内置浏览器面板，
        浏览器里的输入框会让 foundIndex 错位，从而点不到真正的聊天输入框。

        正确做法是用稳定的 AutomationId + ClassName 精确匹配：
            AutomationId = "chat_input_field"
            ClassName    = "mmui::ChatInputField"
        Name 字段是当前会话名称（联系人/群名），会动态变化，因此不参与匹配。
        """
        try:
            chat_edit = wx.EditControl(
                AutomationId="chat_input_field",
                ClassName="mmui::ChatInputField",
                searchDepth=20,
            )
            if chat_edit.Exists(0, 0):
                return chat_edit

            chat_edit = wx.EditControl(ClassName="mmui::ChatInputField", searchDepth=20)
            if chat_edit.Exists(0, 0):
                logger.warning("仅以 ClassName 匹配到聊天输入框，AutomationId 可能已变化")
                return chat_edit

            chat_edit = wx.EditControl(AutomationId="chat_input_field", searchDepth=20)
            if chat_edit.Exists(0, 0):
                logger.warning("仅以 AutomationId 匹配到聊天输入框，ClassName 可能已变化")
                return chat_edit

            return None
        except Exception as e:
            logger.error(f"查找聊天输入框异常: {str(e)}")
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
    
    def send_message_result(self, message):
        """
        发送消息（使用剪贴板粘贴方式，解决 SendKeys 特殊字符问题）
        
        Args:
            message: 要发送的消息内容（支持换行符 \n）
            
        Returns:
            bool: 发送是否成功
        """
        try:
            # 获取微信窗口
            result, wx = self._get_wechat_window_result()
            if not result.success:
                return result
            
            # 激活窗口确保焦点正确
            wx.SetActive()
            time.sleep(0.2)
            
            # 通过 AutomationId="chat_input_field" 精确定位聊天输入框
            # 旧版 foundIndex=1 在打开公众号文章/视频面板后会错位
            chat_edit = self._find_chat_input(wx)
            if chat_edit is None:
                message = "未找到聊天输入框，可能未成功进入目标会话或微信 UI 结构已变化。"
                logger.error(message)
                return self._fail("CHAT_INPUT_NOT_FOUND", message)
            
            # 点击输入框获取焦点
            chat_edit.Click()
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
    
    def send_picture_result(self, image_url):
        """
        发送图片（URL 图片下载后发送；本地路径直接作为文件粘贴发送）
        
        Args:
            image_url: 图片的 URL 或本地文件路径
            
        Returns:
            bool: 发送是否成功
        """
        try:
            if not self._is_url(image_url):
                return self.send_file_result(image_url)

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
            
            # 通过 AutomationId="chat_input_field" 精确定位聊天输入框
            chat_edit = self._find_chat_input(wx)
            if chat_edit is None:
                message = "未找到聊天输入框，可能未成功进入目标会话或微信 UI 结构已变化。"
                logger.error(message)
                return self._fail("CHAT_INPUT_NOT_FOUND", message)
            
            # 点击输入框获取焦点
            chat_edit.Click()
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

    def send_file_result(self, file_path):
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

            # 通过 AutomationId="chat_input_field" 精确定位聊天输入框
            chat_edit = self._find_chat_input(wx)
            if chat_edit is None:
                message = "未找到聊天输入框，可能未成功进入目标会话或微信 UI 结构已变化。"
                logger.error(message)
                return self._fail("CHAT_INPUT_NOT_FOUND", message)

            chat_edit.Click()
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
    
    def search_and_send_result(self, contact_name, message):
        """
        搜索联系人并发送消息（组合操作）
        
        Args:
            contact_name: 联系人名称
            message: 要发送的消息内容
            
        Returns:
            bool: 操作是否成功
        """
        logger.info(f"开始向 '{contact_name}' 发送消息")
        
        # 搜索联系人
        search_result = self.search_contact_result(contact_name)
        if not search_result.success:
            logger.warning(f"跳过向 '{contact_name}' 发送消息（搜索失败）")
            return search_result
        
        # 发送消息
        send_result = self.send_message_result(message)
        if not send_result.success:
            logger.warning(f"向 '{contact_name}' 发送消息失败")
            return send_result
        
        logger.info(f"成功向 '{contact_name}' 发送消息")
        return self._ok(f"发送成功: 已向「{contact_name}」发送文本消息。")

    def search_and_send(self, contact_name, message):
        """搜索联系人并发送文本，保留布尔返回给旧调用方使用。"""
        return self.search_and_send_result(contact_name, message).success
    
    def search_and_send_picture_result(self, contact_name, image_url):
        """
        搜索联系人并发送图片（组合操作）
        
        Args:
            contact_name: 联系人名称
            image_url: 图片的 URL
            
        Returns:
            bool: 操作是否成功
        """
        logger.info(f"开始向 '{contact_name}' 发送图片")
        
        # 搜索联系人
        search_result = self.search_contact_result(contact_name)
        if not search_result.success:
            logger.warning(f"跳过向 '{contact_name}' 发送图片（搜索失败）")
            return search_result
        
        # 发送图片
        send_result = self.send_picture_result(image_url)
        if not send_result.success:
            logger.warning(f"向 '{contact_name}' 发送图片失败")
            return send_result
        
        logger.info(f"成功向 '{contact_name}' 发送图片")
        if not self._is_url(image_url):
            _, resolved_path = self._resolve_local_file_result(image_url)
            return self._ok(f"发送成功: 已向「{contact_name}」发送本地文件「{resolved_path}」。")
        return self._ok(f"发送成功: 已向「{contact_name}」发送图片消息。")

    def search_and_send_picture(self, contact_name, image_url):
        """搜索联系人并发送图片，保留布尔返回给旧调用方使用。"""
        return self.search_and_send_picture_result(contact_name, image_url).success

    def search_and_send_file_result(self, contact_name, file_path):
        """
        搜索联系人并发送本地文件（组合操作）
        """
        logger.info(f"开始向 '{contact_name}' 发送文件")

        search_result = self.search_contact_result(contact_name)
        if not search_result.success:
            logger.warning(f"跳过向 '{contact_name}' 发送文件（搜索失败）")
            return search_result

        send_result = self.send_file_result(file_path)
        if not send_result.success:
            logger.warning(f"向 '{contact_name}' 发送文件失败")
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

