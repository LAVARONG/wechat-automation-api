import argparse
import json
import sys
import os

# 修复 Windows 下 stdout 中文乱码问题
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

# 确保能导入同目录的其他模块
sys.path.append(os.path.dirname(os.path.abspath(__file__)))


def get_project_root():
    """获取项目根目录"""
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def load_config():
    """加载 config.json，失败返回 None"""
    config_file = os.path.join(get_project_root(), 'config.json')
    if not os.path.exists(config_file):
        return None
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return None


def check_api_available(host, port, timeout=1):
    """检测 API 服务是否可访问（自动检测用）"""
    try:
        import requests as req
        resp = req.get(f"http://{host}:{port}/health", timeout=timeout)
        data = resp.json()
        return resp.status_code == 200 and data.get('status') == 'healthy'
    except ImportError:
        return False
    except Exception:
        return False


def send_via_api(host, port, token, to, content, action):
    """通过 HTTP API 发送消息，返回解析后的 JSON 响应"""
    try:
        import requests as req
    except ImportError:
        raise RuntimeError("API 模式需要 requests 库，请执行: pip install requests")
    resp = req.post(
        f"http://{host}:{port}/",
        json={'token': token, 'action': action, 'to': [to], 'content': content},
        timeout=10
    )
    return resp.json()


def main():
    parser = argparse.ArgumentParser(description="微信发信命令行工具(Skill入口)")
    parser.add_argument("--to", required=True, help="接收者联系人名称")
    parser.add_argument("--content", required=True, help="消息内容、图片URL或本地文件路径")
    parser.add_argument("--action", choices=["sendtext", "sendpic", "sendfile"], default="sendtext", help="类型: sendtext(默认)、sendpic 或 sendfile")
    parser.add_argument("--json", action="store_true", help="以 JSON 格式输出结果，便于 Agent 稳定解析")
    parser.add_argument("--api", action="store_true", help="强制使用 API 模式（需先启动 app.py 服务）")
    parser.add_argument("--no-api", action="store_true", help="强制使用 CLI 模式（跳过 API 自动检测）")

    args = parser.parse_args()

    def output(success, code, message):
        if args.json:
            print(json.dumps({
                "success": success,
                "code": code,
                "message": message,
                "to": args.to,
                "action": args.action
            }, ensure_ascii=False))
        else:
            if success:
                print(message)
            else:
                print(f"发送失败: {message}")
                print(f"错误代码: {code}")

    # 决定是否使用 API 模式
    config = load_config()
    api_host = '127.0.0.1'
    api_port = 8808
    use_api = False

    if args.api:
        # 强制 API 模式：缺少 config/token 时直接报错，不回退 CLI
        use_api = True
        if config:
            api_host = config.get('host', '127.0.0.1')
            api_port = config.get('port', 8808)
    elif not args.no_api and config:
        # 自动模式：仅当 token 已配置且 API 在线才走队列，否则静默回退 CLI
        api_host = config.get('host', '127.0.0.1')
        api_port = config.get('port', 8808)
        if config.get('token') and check_api_available(api_host, api_port):
            use_api = True

    # API 模式：通过 HTTP 请求发送（经后台队列串行处理，多请求不冲突）
    # 注意：队列为异步，成功仅代表“已入队受理”，真实发送结果需查 /status 或日志
    if use_api:
        if not config:
            output(False, "NO_CONFIG", "API 模式需要 config.json 配置文件（无法读取配置）")
            sys.exit(1)
        token = config.get('token', '')
        if not token:
            output(False, "NO_TOKEN", "API 模式需要在 config.json 中配置 token 字段")
            sys.exit(1)

        try:
            api_result = send_via_api(api_host, api_port, token, args.to, args.content, args.action)
            if api_result.get('success'):
                queue_size = api_result.get('queue_size', 0)
                output(True, "OK", f"消息已提交发送队列受理（当前排队 {queue_size} 条），由后台异步发送")
                sys.exit(0)
            else:
                error_msg = api_result.get('error', 'API 返回未知错误')
                output(False, "API_ERROR", error_msg)
                sys.exit(1)
        except Exception as e:
            output(False, "API_REQUEST_FAILED", f"API 请求失败: {str(e)}")
            sys.exit(1)

    # CLI 模式：直接调用 WeChatController（原有逻辑）
    try:
        try:
            from wechat_controller import WeChatController
        except ModuleNotFoundError as e:
            output(False, "DEPENDENCY_MISSING", f"缺少 Python 依赖: {e.name}。请在项目根目录执行 pip install -r requirements.txt。")
            sys.exit(1)

        controller = WeChatController()
        if args.action == "sendfile":
            result = controller.search_and_send_file_result(args.to, args.content)
        elif args.action == "sendpic":
            result = controller.search_and_send_picture_result(args.to, args.content)
        else:
            result = controller.search_and_send_result(args.to, args.content)

        output(result.success, result.code, result.message)
        if result.success:
            sys.exit(0)
        sys.exit(1)

    except Exception as e:
        output(False, "CLI_EXCEPTION", f"Skill CLI 执行异常: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
