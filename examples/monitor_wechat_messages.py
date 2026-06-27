"""
微信消息监听脚本
用于监听微信独立聊天窗口的消息变化，检测并输出新消息

功能说明：
- 自动查找所有独立的微信聊天窗口（ClassName: mmui::ChatSingleWindow）
- 每秒轮询一次消息列表，检测新消息
- 将新消息输出到控制台和日志文件
- 支持多窗口同时监听（每个窗口一个独立线程）
- 自动检测新打开的窗口并开始监听

使用方法：
1. 确保微信已启动并登录
2. 打开一个或多个独立的聊天窗口（右键联系人选择"打开聊天窗口"或双击）
3. 运行脚本：python examples/monitor_wechat_messages.py
4. 按 Ctrl+C 停止监听

依赖要求：
- uiautomation
- threading（Python 标准库）
- logging（Python 标准库）

注意事项：
- 只监听独立弹出的聊天窗口，不监听主窗口内的聊天
- 消息内容通过 UI 元素的 Name 属性读取
- 首次运行会将所有现有消息视为"旧消息"
- 每个窗口在独立线程中运行，互不影响
- 动态检测新窗口，每5秒扫描一次
"""

import uiautomation as auto
import time
import logging
import hashlib
import threading
from typing import List, Set, Dict

# 配置日志系统
def setup_logging():
    """配置日志系统"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('wechat_monitor.log', encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)

logger = setup_logging()


class WechatMessageMonitor:
    """微信消息监听器"""
    
    def __init__(self):
        """初始化监听器"""
        self.logger = logger
        # 每个窗口独立维护自己的消息历史（使用窗口句柄作为key）
        self.window_messages: Dict[int, Set[str]] = {}
        # 记录正在监听的窗口句柄
        self.monitoring_windows: Set[int] = set()
        # 线程锁，用于保护共享数据
        self.lock = threading.Lock()
        # 是否继续运行的标志
        self.running = True
        # 记录上次扫描到的窗口数量，用于检测变化
        self.last_window_count = 0
        self.logger.info("微信消息监听器初始化完成")
    
    def _hash_message(self, message: str) -> str:
        """
        计算消息的哈希值
        
        Args:
            message: 消息内容
            
        Returns:
            str: 消息的 MD5 哈希值
        """
        return hashlib.md5(message.encode('utf-8')).hexdigest()
    
    def find_chat_windows(self, log_always: bool = False) -> List[auto.WindowControl]:
        """
        查找所有独立的微信聊天窗口
        
        Args:
            log_always: 是否总是输出日志（默认False，只在数量变化时输出）
        
        Returns:
            List[auto.WindowControl]: 找到的聊天窗口列表
        """
        try:
            chat_windows = []
            # 查找所有 ClassName 为 mmui::ChatSingleWindow 的窗口
            # searchDepth=1 表示只在顶层窗口中搜索
            for window in auto.GetRootControl().GetChildren():
                if window.ClassName == "mmui::ChatSingleWindow":
                    chat_windows.append(window)
            
            # 只在数量变化时输出日志
            current_count = len(chat_windows)
            if log_always or current_count != self.last_window_count:
                self.logger.info(f"找到 {current_count} 个独立聊天窗口")
                self.last_window_count = current_count
            
            return chat_windows
            
        except Exception as e:
            self.logger.error(f"查找聊天窗口时发生错误: {e}", exc_info=True)
            return []
    
    def get_messages(self, window: auto.WindowControl) -> List[str]:
        """
        获取窗口中的所有消息
        
        Args:
            window: 聊天窗口对象
            
        Returns:
            List[str]: 消息内容列表
        """
        messages = []
        try:
            # 查找所有消息列表项
            # ClassName: mmui::ChatTextItemView
            # ControlType: UIA_ListItemControlTypeId
            
            # 查找消息列表容器
            list_control = window.ListControl()
            if not list_control.Exists(0, 0):
                return messages
            
            message_items = list_control.GetChildren()
            
            for item in message_items:
                # 检查是否是聊天消息项
                if "ChatTextItemView" in item.ClassName or "ChatBubbleItemView" in item.ClassName:
                    # 获取消息内容（通过 Name 属性）
                    message_text = item.Name
                    if message_text and message_text.strip():
                        messages.append(message_text.strip())
            
            return messages
            
        except Exception as e:
            # 降低日志级别，避免过多错误信息
            self.logger.debug(f"读取消息列表时发生错误: {e}")
            return []
    
    def monitor_single_window(self, window: auto.WindowControl, window_handle: int):
        """
        监听单个窗口的消息变化（在独立线程中运行）
        
        Args:
            window: 要监听的聊天窗口
            window_handle: 窗口句柄（用于标识窗口）
        """
        try:
            # 获取窗口标题（通常是联系人名称）
            window_title = window.Name
            self.logger.info(f"[线程 {threading.current_thread().name}] 开始监听窗口: {window_title}")
            
            # 为这个窗口初始化消息集合
            with self.lock:
                if window_handle not in self.window_messages:
                    self.window_messages[window_handle] = set()
            
            # 首次读取，将所有现有消息标记为已读
            initial_messages = self.get_messages(window)
            with self.lock:
                for msg in initial_messages:
                    msg_hash = self._hash_message(msg)
                    self.window_messages[window_handle].add(msg_hash)
            
            self.logger.info(f"[{window_title}] 初始化完成，已有 {len(initial_messages)} 条历史消息")
            
            # 开始监听循环
            while self.running:
                try:
                    # 检查窗口是否还存在
                    if not window.Exists(0, 0):
                        self.logger.warning(f"[{window_title}] 窗口已关闭，停止监听")
                        break
                    
                    # 读取当前所有消息
                    current_messages = self.get_messages(window)
                    
                    # 检查新消息
                    new_messages_found = []
                    with self.lock:
                        for msg in current_messages:
                            msg_hash = self._hash_message(msg)
                            if msg_hash not in self.window_messages[window_handle]:
                                # 发现新消息
                                new_messages_found.append(msg)
                                self.window_messages[window_handle].add(msg_hash)
                    
                    # 输出新消息（在锁外执行，避免阻塞）
                    for msg in new_messages_found:
                        self.logger.info(f"[{window_title}] [新消息] {msg}")
                    
                    # 等待1秒后继续
                    time.sleep(1)
                    
                except Exception as e:
                    self.logger.debug(f"[{window_title}] 监听循环中发生错误: {e}")
                    time.sleep(1)  # 出错后等待一秒继续
            
            self.logger.info(f"[{window_title}] 监听线程退出")
                    
        except Exception as e:
            self.logger.error(f"监听窗口时发生严重错误: {e}", exc_info=True)
        finally:
            # 清理：从监听列表中移除
            with self.lock:
                self.monitoring_windows.discard(window_handle)
                if window_handle in self.window_messages:
                    del self.window_messages[window_handle]
    
    def start_monitoring_window(self, window: auto.WindowControl):
        """
        为指定窗口启动监听线程
        
        Args:
            window: 聊天窗口对象
        """
        try:
            # 获取窗口句柄作为唯一标识
            window_handle = window.NativeWindowHandle
            
            with self.lock:
                # 检查是否已经在监听这个窗口
                if window_handle in self.monitoring_windows:
                    return
                
                # 标记为正在监听
                self.monitoring_windows.add(window_handle)
            
            # 创建并启动监听线程
            thread_name = f"Monitor-{window.Name[:10]}"
            thread = threading.Thread(
                target=self.monitor_single_window,
                args=(window, window_handle),
                name=thread_name,
                daemon=True
            )
            thread.start()
            
        except Exception as e:
            self.logger.error(f"启动监听线程失败: {e}", exc_info=True)
            with self.lock:
                self.monitoring_windows.discard(window_handle)
    
    def scan_and_start_monitors(self):
        """扫描所有窗口并为新窗口启动监听"""
        chat_windows = self.find_chat_windows()
        
        if not chat_windows:
            return
        
        # 为每个窗口启动监听线程
        for window in chat_windows:
            self.start_monitoring_window(window)
    
    def run(self):
        """主运行逻辑（多线程版本）"""
        try:
            self.logger.info("=" * 50)
            self.logger.info("开始监听微信消息...")
            self.logger.info("功能：多窗口同时监听（多线程）")
            self.logger.info("提示：请确保至少有一个独立的聊天窗口打开")
            self.logger.info("系统会自动检测新打开的窗口并开始监听")
            self.logger.info("按 Ctrl+C 停止监听")
            self.logger.info("=" * 50)
            
            # 首次扫描并启动监听（总是输出日志）
            chat_windows = self.find_chat_windows(log_always=True)
            for window in chat_windows:
                self.start_monitoring_window(window)
            
            if not self.monitoring_windows:
                self.logger.warning("未找到任何独立聊天窗口！")
                self.logger.info("请打开一个独立的聊天窗口（双击联系人或右键选择'打开聊天窗口'）")
                self.logger.info("程序将每5秒扫描一次...")
            
            # 主循环：定期扫描新窗口
            while self.running:
                try:
                    time.sleep(5)  # 每5秒扫描一次
                    self.scan_and_start_monitors()
                    
                except KeyboardInterrupt:
                    raise
                except Exception as e:
                    self.logger.error(f"主循环发生错误: {e}", exc_info=True)
            
        except KeyboardInterrupt:
            self.logger.info("\n用户中断，正在停止所有监听线程...")
            self.running = False
            # 等待所有线程结束
            time.sleep(2)
        except Exception as e:
            self.logger.error(f"程序运行时发生错误: {e}", exc_info=True)
        finally:
            self.running = False
            self.logger.info("监听已停止")
            self.logger.info(f"共监听了 {len(self.window_messages)} 个窗口")


if __name__ == "__main__":
    monitor = WechatMessageMonitor()
    monitor.run()

