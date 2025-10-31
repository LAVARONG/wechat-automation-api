# 测试文件

本目录包含项目的测试脚本。

## test_api.py - API 功能测试

完整的 API 测试套件，用于测试微信自动化 HTTP 服务的各个功能。

### 测试内容

1. **健康检查** - 测试 `/health` 端点
2. **状态查询** - 测试 `/status` 端点
3. **无效 Token** - 验证 Token 验证功能
4. **缺少字段** - 验证请求参数验证
5. **发送单个消息** - 测试发送单条消息
6. **批量发送** - 测试发送多条消息

### 使用方法

```powershell
# 1. 确保服务已启动
python app.py

# 2. 在新的终端窗口运行测试
python test/test_api.py
```

### 配置

在运行测试前，请确保：
- 微信客户端已启动并登录
- 修改 `test_api.py` 中的联系人名称（默认为 "线报转发" 和 "LAVA"）
- 确认 `BASE_URL` 和 `TOKEN` 配置正确

### 测试结果

测试会输出详细的请求和响应信息，最后显示测试结果汇总：

```
✓ 通过 - 健康检查
✓ 通过 - 状态查询
✓ 通过 - 无效 Token
...
总计: 6/6 个测试通过
```

## 添加新测试

添加新测试函数时，请遵循以下格式：

```python
def test_your_feature():
    """测试描述"""
    print("\n" + "="*50)
    print("测试 N: 测试名称")
    print("="*50)
    
    try:
        # 测试逻辑
        response = requests.post(...)
        print(f"状态码: {response.status_code}")
        print(f"响应: {json.dumps(response.json(), ensure_ascii=False, indent=2)}")
        return response.status_code == 200
    except Exception as e:
        print(f"错误: {str(e)}")
        return False
```

然后将测试函数添加到 `run_all_tests()` 中的 `tests` 列表。

