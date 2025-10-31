"""
API 测试脚本
用于测试微信自动化 HTTP 服务的各个功能
"""
import requests
import json
import time

# 配置
BASE_URL = "http://127.0.0.1:8808"
TOKEN = "123123"


def test_health_check():
    """测试健康检查端点"""
    print("\n" + "="*50)
    print("测试 1: 健康检查")
    print("="*50)
    
    try:
        response = requests.get(f"{BASE_URL}/health")
        print(f"状态码: {response.status_code}")
        print(f"响应: {json.dumps(response.json(), ensure_ascii=False, indent=2)}")
        return response.status_code == 200
    except Exception as e:
        print(f"错误: {str(e)}")
        return False


def test_status():
    """测试状态查询端点"""
    print("\n" + "="*50)
    print("测试 2: 查询状态")
    print("="*50)
    
    try:
        response = requests.get(f"{BASE_URL}/status")
        print(f"状态码: {response.status_code}")
        print(f"响应: {json.dumps(response.json(), ensure_ascii=False, indent=2)}")
        return response.status_code == 200
    except Exception as e:
        print(f"错误: {str(e)}")
        return False


def test_invalid_token():
    """测试无效 token"""
    print("\n" + "="*50)
    print("测试 3: 无效 Token（应该返回 401）")
    print("="*50)
    
    data = {
        "token": "wrong_token",
        "action": "sendtext",
        "to": ["测试联系人"],
        "content": "测试消息"
    }
    
    try:
        response = requests.post(f"{BASE_URL}/", json=data)
        print(f"状态码: {response.status_code}")
        print(f"响应: {json.dumps(response.json(), ensure_ascii=False, indent=2)}")
        return response.status_code == 401
    except Exception as e:
        print(f"错误: {str(e)}")
        return False


def test_missing_fields():
    """测试缺少必需字段"""
    print("\n" + "="*50)
    print("测试 4: 缺少必需字段（应该返回 400）")
    print("="*50)
    
    data = {
        "token": TOKEN,
        "action": "sendtext"
        # 缺少 to 和 content 字段
    }
    
    try:
        response = requests.post(f"{BASE_URL}/", json=data)
        print(f"状态码: {response.status_code}")
        print(f"响应: {json.dumps(response.json(), ensure_ascii=False, indent=2)}")
        return response.status_code == 400
    except Exception as e:
        print(f"错误: {str(e)}")
        return False


def test_send_single_message():
    """测试发送单个消息"""
    print("\n" + "="*50)
    print("测试 5: 发送单个消息")
    print("="*50)
    
    data = {
        "token": TOKEN,
        "action": "sendtext",
        "to": ["线报转发"],
        "content": "这是一条API测试消息 - " + time.strftime("%H:%M:%S")
    }
    
    print(f"请求数据: {json.dumps(data, ensure_ascii=False, indent=2)}")
    
    try:
        response = requests.post(f"{BASE_URL}/", json=data)
        print(f"状态码: {response.status_code}")
        print(f"响应: {json.dumps(response.json(), ensure_ascii=False, indent=2)}")
        return response.status_code == 200
    except Exception as e:
        print(f"错误: {str(e)}")
        return False


def test_send_multiple_recipients():
    """测试发送消息给多个接收者"""
    print("\n" + "="*50)
    print("测试 6: 发送消息给多个接收者")
    print("="*50)
    
    data = {
        "token": TOKEN,
        "action": "sendtext",
        "to": ["线报转发", "LAVA"],  # 多个接收者
        "content": "批量测试消息 - " + time.strftime("%H:%M:%S")
    }
    
    print(f"请求数据: {json.dumps(data, ensure_ascii=False, indent=2)}")
    
    try:
        response = requests.post(f"{BASE_URL}/", json=data)
        print(f"状态码: {response.status_code}")
        print(f"响应: {json.dumps(response.json(), ensure_ascii=False, indent=2)}")
        
        if response.status_code == 200:
            result = response.json()
            print(f"\n已将 {result.get('queued_count', 0)} 条消息加入队列")
            print(f"当前队列大小: {result.get('queue_size', 0)}")
        
        return response.status_code == 200
    except Exception as e:
        print(f"错误: {str(e)}")
        return False


def run_all_tests():
    """运行所有测试"""
    print("\n" + "="*70)
    print(" 微信自动化 HTTP 服务 - API 测试套件")
    print("="*70)
    print(f"测试目标: {BASE_URL}")
    print(f"Token: {TOKEN}")
    print("\n注意：在运行实际发送消息的测试前，请确保：")
    print("  1. 微信客户端已启动并登录")
    print("  2. 测试用的联系人存在")
    print("  3. app.py 服务已启动")
    print("\n按 Enter 开始测试，或 Ctrl+C 取消...")
    try:
        input()
    except KeyboardInterrupt:
        print("\n\n测试已取消")
        return
    
    tests = [
        ("健康检查", test_health_check),
        ("状态查询", test_status),
        ("无效 Token", test_invalid_token),
        ("缺少字段", test_missing_fields),
        ("发送单个消息", test_send_single_message),
        ("批量发送", test_send_multiple_recipients),
    ]
    
    results = []
    
    for name, test_func in tests:
        try:
            result = test_func()
            results.append((name, result))
            time.sleep(1)  # 测试间隔
        except Exception as e:
            print(f"\n测试 '{name}' 发生异常: {str(e)}")
            results.append((name, False))
    
    # 输出测试结果汇总
    print("\n" + "="*70)
    print(" 测试结果汇总")
    print("="*70)
    
    for name, result in results:
        status = "✓ 通过" if result else "✗ 失败"
        print(f"{status} - {name}")
    
    passed = sum(1 for _, result in results if result)
    total = len(results)
    print(f"\n总计: {passed}/{total} 个测试通过")
    print("="*70)


if __name__ == "__main__":
    try:
        run_all_tests()
    except KeyboardInterrupt:
        print("\n\n测试已中断")

