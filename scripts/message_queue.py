"""
消息队列管理模块
实现线程安全的消息队列，支持后台自动处理消息发送

调度策略：贪心"粘连当前联系人"——优先把发往同一联系人的消息连续发完，
再切换到下一个联系人，从而减少在不同会话间来回切换的次数；同一联系人
内部仍按入队先后顺序发送（保证"先文字后图片"等顺序）。
同一联系人连续发送时还会跳过重复的联系人搜索，进一步节省时间。
"""
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
        self.message_interval = message_interval
        self.worker_thread = None
        self.running = False
        # _pending 按入队顺序保存待发送消息；_cond 同时充当互斥锁与条件变量
        self._pending = []
        self._cond = threading.Condition()
        # 标记 worker 是否正在处理某条消息（取出后到发送完成之间为 True）
        self._processing = False

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
        with self._cond:
            self.running = False
            # 唤醒可能正在等待新消息的 worker，使其尽快退出
            self._cond.notify_all()
        if self.worker_thread is not None:
            self.worker_thread.join(timeout=5)
        logger.info("消息队列处理线程已停止")

    def add_message(self, to_list, content, action='sendtext'):
        """
        添加消息到队列

        Args:
            to_list: 接收者列表，可以包含多个联系人
            content: 消息内容（文本、图片 URL 或本地文件路径）
            action: 消息类型，'sendtext'、'sendpic' 或 'sendfile'

        Returns:
            int: 添加到队列的消息数量
        """
        if not isinstance(to_list, list):
            to_list = [to_list]

        count = 0
        with self._cond:
            for contact in to_list:
                self._pending.append({
                    'to': contact,
                    'content': content,
                    'action': action
                })
                count += 1

                if action == 'sendfile':
                    logger.info(f"文件消息已加入队列: 接收者={contact}, 路径={content}")
                elif action == 'sendpic':
                    logger.info(f"图片消息已加入队列: 接收者={contact}, URL={content}")
                else:
                    logger.info(f"文本消息已加入队列: 接收者={contact}, 内容长度={len(content)}")
            # 唤醒处理线程
            self._cond.notify_all()

        return count

    def get_queue_size(self):
        """
        获取当前队列中待处理的消息数量（不含正在发送的那一条）

        Returns:
            int: 队列大小
        """
        with self._cond:
            return len(self._pending)

    def _select_next_locked(self, last_contact):
        """
        从待发送列表中选出下一条要发送的消息（必须在持有 _cond 锁时调用）。

        调度规则：
        1. 若队列中存在发往"上一个联系人"的消息，则取其中最早入队的那条，
           这样同一联系人的多条消息会被连续发送，减少会话切换次数。
        2. 否则取队头（最早入队的消息），保证不同联系人之间的整体先后顺序。

        Args:
            last_contact: 上一条已处理消息的联系人，None 表示尚未发送过

        Returns:
            dict: 选中的消息项
        """
        if last_contact is not None:
            for index, item in enumerate(self._pending):
                if item['to'] == last_contact:
                    return self._pending.pop(index)
        return self._pending.pop(0)

    def _deliver(self, controller, contact, content, action, skip_search):
        """
        发送单条消息。

        Args:
            controller: WeChatController 实例
            contact: 联系人名称
            content: 消息内容
            action: 消息类型（sendtext / sendpic / sendfile）
            skip_search: 是否跳过联系人搜索（当前会话已停留在该联系人时为 True）

        Returns:
            tuple: (success, session_valid)
                success      —— 本条是否发送成功
                session_valid —— 发送后当前会话是否仍确定停留在该联系人，
                                 供下一条判断能否跳过搜索
        """
        send_funcs = {
            'sendfile': controller.send_file_result,
            'sendpic': controller.send_picture_result,
            'sendtext': controller.send_message_result,
        }
        send_func = send_funcs.get(action, controller.send_message_result)

        # 仅当确定当前会话已停留在该联系人时才跳过搜索
        if skip_search:
            logger.info(f"复用当前会话，跳过联系人搜索: 接收者={contact}")
        else:
            search_result = controller.search_contact_result(contact)
            if not search_result.success:
                logger.error(f"搜索联系人失败: 接收者={contact}, code={search_result.code}")
                return False, False

        send_result = send_func(content)
        if send_result.success:
            logger.info(f"消息发送成功: 接收者={contact}, action={action}")
            return True, True

        logger.error(f"消息发送失败: 接收者={contact}, action={action}, code={send_result.code}")
        # 发送失败时当前会话状态不确定，下一条强制重新搜索
        return False, False

    def _process_queue(self):
        """
        后台处理队列中的消息（持续运行的线程函数）。
        采用"粘连当前联系人"调度，并在同一联系人连发时跳过重复搜索。
        """
        logger.info("消息处理线程开始运行")

        # 在线程中初始化 COM，这是使用 uiautomation 在子线程中的必需步骤
        with auto.UIAutomationInitializerInThread():
            wechat_controller = WeChatController()
            logger.info("微信控制器已在线程中初始化")

            last_contact = None     # 上一条处理的联系人（用于粘连排序）
            session_contact = None  # 当前会话已确认停留的联系人（用于跳过搜索）

            while self.running:
                # 取出下一条待发送消息（队列为空时阻塞等待，最多 1 秒以便检查退出标志）
                with self._cond:
                    while self.running and not self._pending:
                        self._cond.wait(timeout=1)
                    if not self.running:
                        break
                    if not self._pending:
                        continue
                    message_item = self._select_next_locked(last_contact)
                    self._processing = True

                try:
                    contact = message_item['to']
                    content = message_item['content']
                    action = message_item.get('action', 'sendtext')

                    # 仅当当前会话确定停留在同一联系人时才跳过搜索
                    skip_search = (session_contact == contact)

                    _, session_valid = self._deliver(
                        wechat_controller, contact, content, action, skip_search
                    )

                    last_contact = contact
                    session_contact = contact if session_valid else None
                except Exception as e:
                    logger.error(f"处理消息时发生错误: {str(e)}", exc_info=True)
                    # 异常后无法确定会话状态，下一条强制重新搜索
                    session_contact = None
                finally:
                    with self._cond:
                        self._processing = False
                        self._cond.notify_all()

                # 间隔一段时间再处理下一条，降低发送过快被风控的风险
                time.sleep(self.message_interval)

        logger.info("消息处理线程已退出")

    def wait_until_empty(self, timeout=None):
        """
        等待队列处理完所有消息（含正在发送的那一条）。

        Args:
            timeout: 超时时间（秒），None 表示无限等待

        Returns:
            bool: True 表示队列已清空，False 表示超时
        """
        start_time = time.time()
        with self._cond:
            while self._pending or self._processing:
                if timeout is not None:
                    remaining = timeout - (time.time() - start_time)
                    if remaining <= 0:
                        return False
                    self._cond.wait(timeout=remaining)
                else:
                    self._cond.wait(timeout=1)
            return True


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

    # 添加测试消息：交错入队两个联系人的文字与图片，
    # 调度后应被重排为「同一联系人连续发送」
    mq.add_message(["线报转发", "文件传输助手"], "这是一条测试文字消息")
    mq.add_message(["线报转发", "文件传输助手"], "https://example.com/test.png", action="sendpic")

    print(f"当前队列大小: {mq.get_queue_size()}")

    # 等待处理完成
    mq.wait_until_empty(timeout=30)
    mq.stop()
