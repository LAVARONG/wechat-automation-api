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
    
    def search_contact(self, contact_name):
        """
        搜索联系人
        
        Args:
            contact_name: 联系人名称
            
        Returns:
            bool: 搜索是否成功
        """
        try:
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
            
            # 点击搜索框并输入
            search_box.Click()
            time.sleep(0.3)
            search_box.SendKeys('{Ctrl}a')  # 清空搜索框
            time.sleep(0.1)
            search_box.SendKeys(contact_name + '{Enter}')
            time.sleep(0.8)
            
            logger.info(f"成功搜索联系人: {contact_name}")
            return True
            
        except Exception as e:
            logger.error(f"搜索联系人 '{contact_name}' 失败: {str(e)}")
            return False
    
    def send_message(self, message):
        """
        发送消息
        
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
            
            # 查找聊天输入框（foundIndex=1 表示第二个 EditControl）
            chat_edit = wx.EditControl(foundIndex=1)
            if not chat_edit.Exists(0, 0):
                logger.error("未找到聊天输入框")
                return False
            
            # 点击输入框
            chat_edit.Click()
            time.sleep(0.3)
            
            # 处理换行符：将 \n 替换为 {Shift}{Enter}
            # 因为微信中 Enter 是发送，Shift+Enter 才是换行
            formatted_message = message.replace('\n', '{Shift}{Enter}')
            
            # 发送消息（最后加上 {Enter} 发送）
            chat_edit.SendKeys(formatted_message + '{Enter}')
            time.sleep(0.5)
            
            # 日志中显示原始消息（包含换行符）
            log_preview = message.replace('\n', '\\n')[:50]
            logger.info(f"成功发送消息: {log_preview}...")
            return True
            
        except Exception as e:
            logger.error(f"发送消息失败: {str(e)}")
            return False
    
    def _download_image(self, url):
        """
        从 URL 下载图片到缓存文件（使用 MD5 作为文件名避免重复下载）
        
        Args:
            url: 图片的 URL
            
        Returns:
            str: 缓存文件路径，失败返回 None
        """
        try:
            # 计算 URL 的 MD5 作为文件名
            url_md5 = hashlib.md5(url.encode('utf-8')).hexdigest()
            cache_path = os.path.join(self.cache_dir, f"{url_md5}.png")
            
            # 检查缓存文件是否已存在
            if os.path.exists(cache_path):
                logger.info(f"使用缓存图片: {cache_path}")
                return cache_path
            
            logger.info(f"开始下载图片: {url}")
            
            # 下载图片
            response = requests.get(url, timeout=30)
            response.raise_for_status()
            
            # 判断内容类型
            content_type = response.headers.get('content-type', '')
            if not content_type.startswith('image/'):
                logger.error(f"URL 返回的不是图片类型: {content_type}")
                return None
            
            # 打开图片
            image = Image.open(BytesIO(response.content))
            
            # 保存图片到缓存目录
            image.save(cache_path, 'PNG')
            logger.info(f"图片已下载并缓存到: {cache_path}")
            
            return cache_path
            
        except Exception as e:
            logger.error(f"下载图片失败: {str(e)}")
            return None
    
    def _copy_image_to_clipboard(self, image_path):
        """
        将图片复制到剪贴板
        
        Args:
            image_path: 图片文件路径
            
        Returns:
            bool: 操作是否成功
        """
        try:
            # 打开图片
            image = Image.open(image_path)
            
            # 转换为 BMP 格式（Windows 剪贴板需要）
            output = BytesIO()
            image.convert('RGB').save(output, 'BMP')
            data = output.getvalue()[14:]  # BMP 文件头是 14 字节，剪贴板不需要
            output.close()
            
            # 复制到剪贴板
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            win32clipboard.CloseClipboard()
            
            logger.info("图片已复制到剪贴板")
            return True
            
        except Exception as e:
            logger.error(f"复制图片到剪贴板失败: {str(e)}")
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
            
            # 查找聊天输入框
            chat_edit = wx.EditControl(foundIndex=1)
            if not chat_edit.Exists(0, 0):
                logger.error("未找到聊天输入框")
                return False
            
            # 点击输入框
            chat_edit.Click()
            time.sleep(0.3)
            
            # 粘贴图片（Ctrl+V）
            chat_edit.SendKeys('{Ctrl}v')
            time.sleep(0.5)
            
            # 发送（Enter）
            chat_edit.SendKeys('{Enter}')
            time.sleep(0.5)
            
            logger.info(f"成功发送图片: {image_url}")
            return True
            
        except Exception as e:
            logger.error(f"发送图片失败: {str(e)}")
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

