"""
微信控制器模块
封装微信操作，提供搜索联系人和发送消息的功能
"""
import uiautomation as auto
import time
import logging
import requests
import os
import tempfile
from PIL import Image
from io import BytesIO
import win32clipboard
from pathlib import Path
import hashlib

# 配置日志
logger = logging.getLogger(__name__)


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
        
    def _get_wechat_window(self):
        """获取微信窗口对象"""
        try:
            # 第一次尝试查找微信窗口
            wx = auto.WindowControl(searchDepth=1, Name="微信", ClassName='mmui::MainWindow')
            if wx.Exists(0, 0):
                return wx
            
            # 第一次找不到，尝试用快捷键唤醒微信窗口（Ctrl+Alt+W 是微信的默认快捷键）
            logger.info("未找到微信窗口，尝试使用快捷键 Ctrl+Alt+W 唤醒微信...")
            auto.SendKeys('{Ctrl}{Alt}w', waitTime=0.1)
            time.sleep(1.0)  # 等待窗口显示
            
            # 第二次尝试查找微信窗口
            wx = auto.WindowControl(searchDepth=1, Name="微信", ClassName='mmui::MainWindow')
            if wx.Exists(0, 0):
                logger.info("成功通过快捷键唤醒微信窗口")
                return wx
            else:
                logger.error("未找到微信窗口，请确保微信已启动并设置了 Ctrl+Alt+W 快捷键")
                return None
        except Exception as e:
            logger.error(f"获取微信窗口失败: {str(e)}")
            return None
    
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
    
    def _activate_from_session_list(self, contact_name):
        """
        从左侧会话列表直接激活对话（快速方法）
        
        Args:
            contact_name: 联系人名称
            
        Returns:
            bool: 是否成功激活
        """
        try:
            # 获取微信窗口
            wx = self._get_wechat_window()
            if not wx:
                return False
            
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
                    return True
                
                # 未选中，需要点击激活
                logger.info(f"点击激活会话: {contact_name}")
                session_item.Click()
                time.sleep(0.3)
                
                # 验证是否选中成功
                if self._is_session_selected(session_item):
                    logger.info(f"从会话列表成功激活联系人: {contact_name}")
                    return True
                else:
                    logger.warning(f"点击后会话 '{contact_name}' 未被选中")
                    return False
            else:
                logger.debug(f"会话列表中未找到 {contact_name}，将使用搜索方式")
                return False
                
        except Exception as e:
            logger.debug(f"从会话列表激活失败: {str(e)}")
            return False
    
    def search_contact(self, contact_name):
        """
        搜索联系人（双重策略：优先从会话列表激活，找不到再搜索）
        
        Args:
            contact_name: 联系人名称
            
        Returns:
            bool: 搜索是否成功
        """
        try:
            # 策略1: 优先尝试从会话列表直接激活（更快）
            if self._activate_from_session_list(contact_name):
                return True
            
            # 策略2: 降级使用搜索框（兜底方案）
            logger.info(f"使用搜索框查找联系人: {contact_name}")
            
            # 获取微信窗口
            wx = self._get_wechat_window()
            if not wx:
                return False
            
            # 激活窗口
            wx.SetActive()
            time.sleep(0.5)
            
            # 查找搜索框
            search_box = wx.EditControl(Name='搜索')
            if not search_box.Exists(0, 0):
                logger.error("未找到搜索框")
                return False
            
            # 点击搜索框
            search_box.Click()
            time.sleep(0.3)
            
            # 使用剪贴板粘贴方式输入搜索内容（更快且避免特殊字符问题）
            if self._set_clipboard_text(contact_name):
                search_box.SendKeys('{Ctrl}v')
                time.sleep(0.2)
            else:
                # 剪贴板设置失败，回退到 SendKeys
                logger.warning("搜索时剪贴板设置失败，使用 SendKeys 方式")
                escaped_name = contact_name.replace('{', '{{').replace('}', '}}')
                search_box.SendKeys(escaped_name, interval=0.01)
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
                    return True
                else:
                    logger.warning(f"搜索后会话 '{contact_name}' 未被选中")
                    return False
            else:
                # 搜索框可能找到的是其他类型的结果（如公众号、小程序等）
                # 这种情况下暂时认为搜索成功，保持向下兼容
                logger.info(f"完成搜索联系人: {contact_name}（无法验证选中状态）")
                return True
            
        except Exception as e:
            logger.error(f"搜索联系人 '{contact_name}' 失败: {str(e)}")
            return False
    
    def _set_clipboard_text(self, text, max_retries=3):
        """
        安全地设置剪贴板文本（带重试机制）
        
        Args:
            text: 要设置的文本内容
            max_retries: 最大重试次数
            
        Returns:
            bool: 操作是否成功
        """
        for attempt in range(max_retries):
            try:
                auto.SetClipboardText(text)
                # 验证剪贴板内容是否设置成功
                time.sleep(0.05)
                clipboard_content = auto.GetClipboardText()
                if clipboard_content == text:
                    return True
                else:
                    logger.warning(f"剪贴板内容验证失败，重试中... (尝试 {attempt + 1}/{max_retries})")
            except Exception as e:
                logger.warning(f"设置剪贴板失败: {e}，重试中... (尝试 {attempt + 1}/{max_retries})")
            time.sleep(0.1)
        
        logger.error("设置剪贴板文本失败，已达最大重试次数")
        return False
    
    def send_message(self, message):
        """
        发送消息（使用剪贴板粘贴方式，解决 SendKeys 特殊字符问题）
        
        Args:
            message: 要发送的消息内容（支持换行符 \n）
            
        Returns:
            bool: 发送是否成功
        """
        try:
            # 获取微信窗口
            wx = self._get_wechat_window()
            if not wx:
                return False
            
            # 激活窗口确保焦点正确
            wx.SetActive()
            time.sleep(0.2)
            
            # 查找聊天输入框（foundIndex=1 表示第二个 EditControl）
            chat_edit = wx.EditControl(foundIndex=1)
            if not chat_edit.Exists(0, 0):
                logger.error("未找到聊天输入框")
                return False
            
            # 点击输入框获取焦点
            chat_edit.Click()
            time.sleep(0.2)
            
            # 使用剪贴板粘贴方式发送消息
            # 这种方式比 SendKeys 快得多，且不会出现特殊字符（如【】￥等）被误解析的问题
            if not self._set_clipboard_text(message):
                # 剪贴板设置失败，回退到 SendKeys 方式（作为兜底）
                logger.warning("剪贴板设置失败，尝试使用 SendKeys 方式")
                # 转义特殊字符，避免被 SendKeys 误解析
                escaped_message = message.replace('{', '{{').replace('}', '}}')
                formatted_message = escaped_message.replace('\n', '{Shift}{Enter}')
                chat_edit.SendKeys(formatted_message + '{Enter}', interval=0.01)
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
            return True
            
        except Exception as e:
            logger.error(f"发送消息失败: {str(e)}", exc_info=True)
            return False
    
    def _download_image(self, url, max_retries=3):
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
                    return cache_path
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
                            return None
                    
                    # 尝试打开图片验证有效性
                    image = Image.open(BytesIO(response.content))
                    image.verify()  # 验证图片完整性
                    
                    # 重新打开图片（verify 后需要重新打开）
                    image = Image.open(BytesIO(response.content))
                    
                    # 保存图片到缓存目录
                    image.save(cache_path, 'PNG')
                    logger.info(f"图片已下载并缓存到: {cache_path}")
                    
                    return cache_path
                    
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
            
            logger.error(f"下载图片失败，已达最大重试次数: {last_error}")
            return None
            
        except Exception as e:
            logger.error(f"下载图片过程中发生错误: {str(e)}", exc_info=True)
            return None
    
    def _copy_image_to_clipboard(self, image_path, max_retries=3):
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
                return True
                
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
        
        return False
    
    def send_picture(self, image_url):
        """
        发送图片（通过 URL 下载后粘贴发送，使用缓存避免重复下载）
        
        Args:
            image_url: 图片的 URL
            
        Returns:
            bool: 发送是否成功
        """
        try:
            # 下载图片（或使用缓存）
            cache_file = self._download_image(image_url)
            if not cache_file:
                return False
            
            # 复制图片到剪贴板
            if not self._copy_image_to_clipboard(cache_file):
                return False
            
            # 获取微信窗口
            wx = self._get_wechat_window()
            if not wx:
                return False
            
            # 激活窗口确保焦点正确
            wx.SetActive()
            time.sleep(0.2)
            
            # 查找聊天输入框
            chat_edit = wx.EditControl(foundIndex=1)
            if not chat_edit.Exists(0, 0):
                logger.error("未找到聊天输入框")
                return False
            
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
            return True
            
        except Exception as e:
            logger.error(f"发送图片失败: {str(e)}", exc_info=True)
            return False
    
    def search_and_send(self, contact_name, message):
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
        if not self.search_contact(contact_name):
            logger.warning(f"跳过向 '{contact_name}' 发送消息（搜索失败）")
            return False
        
        # 发送消息
        if not self.send_message(message):
            logger.warning(f"向 '{contact_name}' 发送消息失败")
            return False
        
        logger.info(f"成功向 '{contact_name}' 发送消息")
        return True
    
    def search_and_send_picture(self, contact_name, image_url):
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
        if not self.search_contact(contact_name):
            logger.warning(f"跳过向 '{contact_name}' 发送图片（搜索失败）")
            return False
        
        # 发送图片
        if not self.send_picture(image_url):
            logger.warning(f"向 '{contact_name}' 发送图片失败")
            return False
        
        logger.info(f"成功向 '{contact_name}' 发送图片")
        return True


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

