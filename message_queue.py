"""
消息队列管理模块
实现线程安全的消息队列，支持后台自动处理消息发送
"""
import queue
import threading
import time
import logging
import uiautomation as auto
from wechat_controller import WeChatController

# 配置日志
logger = logging.getLogger(__name__)


class MessageQueue:
    """消息队列管理类，负责管理和处理微信消息发送队列"""
    
    def __init__(self, message_interval=1):
        """
        初始化消息队列
        
        Args:
            message_interval: 消息发送间隔时间（秒），默认1秒
        """
        self.queue = queue.Queue()
        self.message_interval = message_interval
        self.worker_thread = None
        self.running = False
        self.lock = threading.Lock()
        
    def start(self):
        """启动消息队列处理线程"""
        if self.worker_thread is not None and self.worker_thread.is_alive():
            logger.warning("消息队列处理线程已在运行")
            return
        
        self.running = True
        self.worker_thread = threading.Thread(target=self._process_queue, daemon=True)
        self.worker_thread.start()
        logger.info("消息队列处理线程已启动")
    
    def stop(self):
        """停止消息队列处理线程"""
        self.running = False
        if self.worker_thread is not None:
            self.worker_thread.join(timeout=5)
        logger.info("消息队列处理线程已停止")
    
    def add_message(self, to_list, content, action='sendtext'):
        """
        添加消息到队列
        
        Args:
            to_list: 接收者列表，可以包含多个联系人
            content: 消息内容（文本或图片 URL）
            action: 消息类型，'sendtext' 或 'sendpic'
            
        Returns:
            int: 添加到队列的消息数量
        """
        if not isinstance(to_list, list):
            to_list = [to_list]
        
        count = 0
        for contact in to_list:
            message_item = {
                'to': contact,
                'content': content,
                'action': action
            }
            self.queue.put(message_item)
            count += 1
            
            if action == 'sendpic':
                logger.info(f"图片消息已加入队列: 接收者={contact}, URL={content}")
            else:
                logger.info(f"文本消息已加入队列: 接收者={contact}, 内容长度={len(content)}")
        
        return count
    
    def get_queue_size(self):
        """
        获取当前队列中待处理的消息数量
        
        Returns:
            int: 队列大小
        """
        return self.queue.qsize()
    
    def _process_queue(self):
        """
        后台处理队列中的消息
        这是一个持续运行的线程函数
        """
        logger.info("消息处理线程开始运行")
        
        # 在线程中初始化 COM，这是使用 uiautomation 在子线程中的必需步骤
        with auto.UIAutomationInitializerInThread():
            # 在线程中创建微信控制器
            wechat_controller = WeChatController()
            logger.info("微信控制器已在线程中初始化")
            
            while self.running:
                try:
                    # 尝试从队列获取消息，超时时间1秒
                    try:
                        message_item = self.queue.get(timeout=1)
                    except queue.Empty:
                        # 队列为空，继续循环
                        continue
                    
                    # 提取消息信息
                    contact = message_item['to']
                    content = message_item['content']
                    action = message_item.get('action', 'sendtext')  # 默认为文本消息
                    
                    # 根据 action 类型发送消息
                    if action == 'sendpic':
                        logger.info(f"开始处理图片消息: 接收者={contact}")
                        success = wechat_controller.search_and_send_picture(contact, content)
                        
                        if success:
                            logger.info(f"图片发送成功: 接收者={contact}")
                        else:
                            logger.error(f"图片发送失败: 接收者={contact}")
                    else:
                        logger.info(f"开始处理文本消息: 接收者={contact}")
                        success = wechat_controller.search_and_send(contact, content)
                        
                        if success:
                            logger.info(f"文本消息发送成功: 接收者={contact}")
                        else:
                            logger.error(f"文本消息发送失败: 接收者={contact}")
                    
                    # 标记任务完成
                    self.queue.task_done()
                    
                    # 等待指定的间隔时间再处理下一条消息
                    time.sleep(self.message_interval)
                    
                except Exception as e:
                    logger.error(f"处理消息时发生错误: {str(e)}", exc_info=True)
        
        logger.info("消息处理线程已退出")
    
    def wait_until_empty(self, timeout=None):
        """
        等待队列处理完所有消息
        
        Args:
            timeout: 超时时间（秒），None 表示无限等待
            
        Returns:
            bool: True 表示队列已空，False 表示超时
        """
        try:
            if timeout:
                self.queue.join()
                return True
            else:
                # 使用超时等待
                start_time = time.time()
                while not self.queue.empty():
                    if time.time() - start_time > timeout:
                        return False
                    time.sleep(0.1)
                return True
        except Exception as e:
            logger.error(f"等待队列清空时发生错误: {str(e)}")
            return False


# 测试代码
if __name__ == "__main__":
    # 配置日志
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # 创建消息队列并测试
    mq = MessageQueue(message_interval=1)
    mq.start()
    
    # 添加测试消息
    mq.add_message(["线报转发"], "这是一条测试消息 1")
    mq.add_message(["线报转发"], "这是一条测试消息 2")
    
    print(f"当前队列大小: {mq.get_queue_size()}")
    
    # 等待处理完成
    time.sleep(5)
    mq.stop()

