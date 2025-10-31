@echo off
chcp 65001 >nul
REM ====================================
REM 断开远程桌面连接（保持会话）
REM 解决关闭远程桌面后鼠标键盘失效的问题
REM ====================================

echo ========================================
echo 正在查询当前会话信息...
echo ========================================
echo.

REM 显示当前所有会话
query session

echo.
echo ========================================
echo 正在查找 RDP 会话...
echo ========================================
echo.

REM 查找 rdp-tcp 开头的会话ID
for /f "tokens=2,3" %%a in ('query session ^| findstr /i "rdp-tcp"') do (
    set SESSION_ID=%%b
    set SESSION_NAME=%%a
    echo 找到 RDP 会话: %%a (ID: %%b^)
    goto :disconnect
)

echo 未找到活动的 RDP 会话！
echo 可能原因：
echo 1. 当前不在远程桌面会话中
echo 2. 需要以管理员身份运行此脚本
echo.
pause
exit /b 1

:disconnect
echo.
echo ========================================
echo 正在断开远程桌面连接...
echo ========================================
echo.
echo 执行命令: tscon %SESSION_ID% /dest:console
echo.

REM 执行断开命令（使用 session ID）
tscon %SESSION_ID% /dest:console

REM 检查命令是否执行成功
if %errorlevel% equ 0 (
    echo 远程桌面已成功断开！
    echo 会话已返回到本地控制台。
) else (
    echo 断开失败！错误代码: %errorlevel%
    echo.
    echo 请确保：
    echo 1. 以管理员身份运行此脚本
    echo 2. 当前在远程桌面会话中
    echo 3. 拥有足够的权限执行此操作
    echo.
    pause
)

exit /b %errorlevel%

