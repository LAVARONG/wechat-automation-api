import os
import sys
import json
import logging
import requests

def get_project_root():
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def load_config():
    config_file = os.path.join(get_project_root(), 'config.json')
    if not os.path.exists(config_file):
        raise FileNotFoundError(f"配置文件 {config_file} 不存在")
    
    with open(config_file, 'r', encoding='utf-8') as f:
        return json.load(f)

def send_test_notification():
    # 简单的日志配置，输出到控制台
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    logger = logging.getLogger(__name__)

    try:
        config = load_config()
    except Exception as e:
        logger.error(f"加载配置失败: {e}")
        return

    wpush_config = config.get('wpush', {})
    apikey = wpush_config.get('apikey')
    
    if not apikey or apikey == '你的apikey':
        logger.warning(f"未配置有效的 WPush apikey，当前读取到的值为: '{apikey}'")
        return
        
    url = "https://api.wpush.cn/api/v1/send"
    headers = {'Content-Type': 'application/json'}
    payload = {
        "apikey": apikey,
        "title": wpush_config.get('title', '微信监控通知') + " - 测试消息",
        "content": wpush_config.get('content', '微信可能已掉线。') + "\n\n这是一条测试消息，如果您收到此消息，说明 WPush 通知配置正常！"
    }
    
    logger.info("========== 发送通知配置信息 ==========")
    logger.info(f"APIURL: {url}")
    logger.info(f"APIKEY: {apikey}")
    logger.info(f"Title: {payload['title']}")
    logger.info(f"Content: {payload['content']}")
    logger.info("=====================================")
    logger.info("正在发送请求中，请稍候...")
    
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=10)
        
        logger.info(f"服务器返回状态码: {response.status_code}")
        logger.info(f"服务器返回内容: {response.text}")
        
        response.raise_for_status()
        logger.info("✅ WPush 测试通知发送成功！请检查您的手机或相关设备。")
    except requests.exceptions.RequestException as e:
        logger.error(f"❌ WPush 通知发送失败，请求异常: {str(e)}")
    except Exception as e:
        logger.error(f"❌ WPush 通知发送网络请求报错: {str(e)}")

if __name__ == '__main__':
    print("开始运行 WPush 通知发送测试脚本...\n")
    send_test_notification()
    print("\n测试脚本运行结束。")
