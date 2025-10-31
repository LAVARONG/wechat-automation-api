"""
微信自动化 HTTP 服务
提供 RESTful API 接口，接收消息发送请求并加入队列处理
"""
from flask import Flask, request, jsonify
import json
import logging
import os
from message_queue import MessageQueue

# 创建 Flask 应用
app = Flask(__name__)

# 全局变量
message_queue = None
config = None

# 配置日志
def setup_logging():
    """配置日志系统"""
    log_level = config.get('log_level', 'INFO')
    log_file = config.get('log_file', 'wechat_automation.log')
    
    logging.basicConfig(
        level=getattr(logging, log_level),
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    
    logger = logging.getLogger(__name__)
    logger.info("日志系统已初始化")
    return logger

# 加载配置
def load_config():
    """从 config.json 加载配置"""
    config_file = 'config.json'
    if not os.path.exists(config_file):
        raise FileNotFoundError(f"配置文件 {config_file} 不存在")
    
    with open(config_file, 'r', encoding='utf-8') as f:
        return json.load(f)

# 验证 token
def verify_token(token):
    """
    验证请求的 token 是否正确
    
    Args:
        token: 请求中的 token
        
    Returns:
        bool: token 是否有效
    """
    return token == config.get('token')

# 验证请求参数
def validate_request_data(data):
    """
    验证请求数据的完整性
    
    Args:
        data: 请求的 JSON 数据
        
    Returns:
        tuple: (是否有效, 错误信息)
    """
    # 检查必需字段
    required_fields = ['token', 'action', 'to', 'content']
    for field in required_fields:
        if field not in data:
            return False, f"缺少必需字段: {field}"
    
    # 验证 action
    if data['action'] not in ['sendtext', 'sendpic']:
        return False, f"不支持的操作: {data['action']}"
    
    # 验证 to 字段是否为列表
    if not isinstance(data['to'], list):
        return False, "'to' 字段必须是数组"
    
    if len(data['to']) == 0:
        return False, "'to' 字段不能为空"
    
    # 验证 content 字段
    if not isinstance(data['content'], str) or len(data['content']) == 0:
        return False, "'content' 字段必须是非空字符串"
    
    return True, None


@app.route('/', methods=['POST'])
def send_message():
    """
    处理消息发送请求
    
    请求格式 (发送文本):
    {
        "token": "123123",
        "action": "sendtext",
        "to": ["联系人1", "联系人2"],
        "content": "消息内容"
    }
    
    请求格式 (发送图片):
    {
        "token": "123123",
        "action": "sendpic",
        "to": ["联系人1", "联系人2"],
        "content": "图片的URL"
    }
    
    响应格式:
    {
        "success": true,
        "message": "消息已加入队列",
        "queued_count": 2
    }
    """
    logger = logging.getLogger(__name__)
    
    try:
        # 获取 JSON 数据
        data = request.get_json()
        
        if not data:
            logger.warning("收到空请求")
            return jsonify({
                'success': False,
                'error': '请求体不能为空'
            }), 400
        
        logger.info(f"收到消息发送请求: to={data.get('to', [])}, content_length={len(data.get('content', ''))}")
        
        # 验证请求参数
        is_valid, error_msg = validate_request_data(data)
        if not is_valid:
            logger.warning(f"请求参数验证失败: {error_msg}")
            return jsonify({
                'success': False,
                'error': error_msg
            }), 400
        
        # 验证 token
        if not verify_token(data['token']):
            logger.warning(f"Token 验证失败")
            return jsonify({
                'success': False,
                'error': '无效的 token'
            }), 401
        
        # 将消息加入队列
        to_list = data['to']
        content = data['content']
        action = data['action']
        queued_count = message_queue.add_message(to_list, content, action)
        
        logger.info(f"消息已加入队列: 接收者数量={queued_count}, 队列大小={message_queue.get_queue_size()}")
        
        # 返回成功响应
        return jsonify({
            'success': True,
            'message': '消息已加入队列',
            'queued_count': queued_count,
            'queue_size': message_queue.get_queue_size()
        }), 200
        
    except Exception as e:
        logger.error(f"处理请求时发生错误: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': f'服务器内部错误: {str(e)}'
        }), 500


@app.route('/status', methods=['GET'])
def get_status():
    """
    获取服务状态
    
    响应格式:
    {
        "status": "running",
        "queue_size": 5
    }
    """
    return jsonify({
        'status': 'running',
        'queue_size': message_queue.get_queue_size()
    }), 200


@app.route('/health', methods=['GET'])
def health_check():
    """健康检查端点"""
    return jsonify({
        'status': 'healthy'
    }), 200


def main():
    """主函数，初始化并启动服务"""
    global message_queue, config
    
    # 加载配置
    try:
        config = load_config()
        print("配置加载成功")
    except Exception as e:
        print(f"加载配置失败: {str(e)}")
        return
    
    # 配置日志
    logger = setup_logging()
    
    # 创建并启动消息队列
    message_interval = config.get('message_interval', 1)
    message_queue = MessageQueue(message_interval=message_interval)
    message_queue.start()
    logger.info("消息队列已启动")
    
    # 获取服务器配置
    host = config.get('host', '127.0.0.1')
    port = config.get('port', 8808)
    
    # 启动 Flask 服务
    logger.info(f"微信自动化服务启动在 http://{host}:{port}")
    print(f"\n========================================")
    print(f"微信自动化服务已启动")
    print(f"监听地址: http://{host}:{port}")
    print(f"API 端点: POST http://{host}:{port}/")
    print(f"状态查询: GET http://{host}:{port}/status")
    print(f"健康检查: GET http://{host}:{port}/health")
    print(f"========================================\n")
    
    try:
        app.run(host=host, port=port, debug=False)
    except KeyboardInterrupt:
        logger.info("收到退出信号")
    finally:
        # 停止消息队列
        logger.info("正在停止消息队列...")
        message_queue.stop()
        logger.info("服务已停止")


if __name__ == '__main__':
    main()

