# Fluent Appx 批量安装器

一个具有 Fluent Design 风格的 `.appx`、`.msix`、`.appxbundle`、`.msixbundle` 批量安装工具。

## 功能特性

- ✅ **Fluent Design 风格界面** - 现代化 UI，视觉体验优秀
- ✅ **批量安装** - 支持同时安装多个 Appx 包文件
- ✅ **拖拽操作** - 直接拖拽文件到窗口即可添加
- ✅ **进度显示** - 实时显示安装进度和状态
- ✅ **多种格式支持** - 支持 `.appx`, `.msix`, `.appxbundle`, `.msixbundle`
- ✅ **开发者模式检测** - 自动检测并提示开启开发者模式

## 系统要求

- Windows 10/11
- PowerShell 5.1 或更高版本
- 建议开启开发者模式（设置 → 更新和安全 → 开发者选项）

## 使用方法

### 方法一：直接运行

1. 右键点击 `AppxBatchInstaller.ps1`
2. 选择 **"使用 PowerShell 运行"**

### 方法二：命令行运行

```powershell
# 以管理员身份打开 PowerShell，然后执行：
powershell -ExecutionPolicy Bypass -File "D:\Projects\AppxInstaller\AppxBatchInstaller.ps1"
```

### 方法三：创建快捷方式

创建一个 `.bat` 批处理文件，内容如下：

```batch
@echo off
powershell -ExecutionPolicy Bypass -File "%~dp0AppxBatchInstaller.ps1"
pause
```

## 操作说明

1. **添加文件** - 点击"➕ 添加文件"按钮选择要安装的包文件
2. **添加文件夹** - 点击"📁 添加文件夹"批量添加某个文件夹内的所有包文件
3. **拖拽添加** - 直接拖拽文件到文件列表区域
4. **开始安装** - 点击"▶️ 开始安装"按钮执行批量安装
5. **查看进度** - 安装过程中实时显示进度和每个文件的状态

## 注意事项

- ⚠️ 部分应用可能需要**管理员权限**才能安装
- ⚠️ 如果安装失败，请检查是否已开启**开发者模式**
- ⚠️ 某些企业签名的应用可能需要信任相应的证书

## 常见问题

### Q: 安装时提示"需要开发人员许可"？
A: 请在 Windows 设置中开启开发者模式：
   设置 → 更新和安全 → 开发者选项 → 开发人员模式

### Q: 提示"证书不受信任"？
A: 需要先安装并信任应用的签名证书，或使用自签名证书的应用。

### Q: 拖拽功能不生效？
A: 请确保以正常方式运行脚本，某些受限的 PowerShell 环境可能不支持拖拽。

## 文件结构

```
AppxInstaller/
├── AppxBatchInstaller.ps1    # 主程序
├── 启动安装器.cmd             # 快捷启动脚本
└── README.md                  # 说明文档
```

## 技术实现

- **UI 框架**: WPF (Windows Presentation Foundation)
- **安装命令**: `Add-AppxPackage`
- **设计风格**: Microsoft Fluent Design System

---

📦 **Fluent Appx 批量安装器** - 让 Appx 安装更简单！
