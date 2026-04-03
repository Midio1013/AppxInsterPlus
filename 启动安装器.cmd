@echo off
chcp 65001 >nul 2>&1
title Fluent Appx 批量安装器

echo ============================================
echo   Fluent Appx 批量安装器
echo ============================================
echo.
echo 正在启动安装器...
echo.

powershell -ExecutionPolicy Bypass -File "%~dp0AppxBatchInstaller.ps1"

echo.
echo 安装器已关闭。
pause
