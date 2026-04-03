# =============================================================================
# Fluent 风格 Appx/Msix 批量安装器
# 支持：.appx, .msix, .appxbundle, .msixbundle
# =============================================================================

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Drawing

# XAML 界面定义
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Fluent Appx 批量安装器" 
        Height="600" Width="800"
        WindowStartupLocation="CenterScreen"
        Background="#F3F3F3"
        FontFamily="Segoe UI Variable, Segoe UI"
        AllowsTransparency="False"
        WindowStyle="SingleBorderWindow">
    
    <Window.Resources>
        <!-- Fluent 风格按钮 -->
        <Style x:Key="FluentButton" TargetType="Button">
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="20,10"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                CornerRadius="4" 
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#106EBE"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#C8C8C8"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- Fluent 风格列表项 -->
        <Style x:Key="FileListItem" TargetType="ListBoxItem">
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="Margin" Value="0,2"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBoxItem">
                        <Border x:Name="Bd" 
                                Background="{TemplateBinding Background}" 
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#E5F3FF"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#CCE4F7"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- 进度条样式 -->
        <Style x:Key="FluentProgressBar" TargetType="ProgressBar">
            <Setter Property="Height" Value="4"/>
            <Setter Property="Background" Value="#E0E0E0"/>
            <Setter Property="Foreground" Value="#0078D4"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- 标题区域 -->
        <StackPanel Grid.Row="0" Margin="0,0,0,20">
            <TextBlock Text="📦 Fluent Appx 批量安装器" 
                       FontSize="28" FontWeight="SemiBold" 
                       Foreground="#202020"/>
            <TextBlock Text="支持 .appx, .msix, .appxbundle, .msixbundle 格式" 
                       FontSize="14" Foreground="#606060" 
                       Margin="0,5,0,0"/>
        </StackPanel>
        
        <!-- 文件列表区域 -->
        <Border Grid.Row="1" 
                Background="White" 
                CornerRadius="8" 
                BorderBrush="#E0E0E0" 
                BorderThickness="1"
                AllowDrop="True"
                x:Name="DropZone">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                
                <!-- 工具栏 -->
                <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="15,15,15,10">
                    <Button x:Name="BtnAddFiles" Content="➕ 添加文件" 
                            Style="{StaticResource FluentButton}" Margin="0,0,10,0"/>
                    <Button x:Name="BtnAddFolder" Content="📁 添加文件夹" 
                            Style="{StaticResource FluentButton}" Margin="0,0,10,0"/>
                    <Button x:Name="BtnClear" Content="🗑️ 清空列表" 
                            Style="{StaticResource FluentButton}" 
                            Background="#E8E8E8" Foreground="#202020"/>
                    <TextBlock x:Name="TxtFileCount" 
                               Text="0 个文件待安装" 
                               VerticalAlignment="Center" 
                               Margin="15,0,0,0"
                               Foreground="#606060"/>
                </StackPanel>
                
                <!-- 文件列表 -->
                <ListBox Grid.Row="1" 
                         x:Name="FileList" 
                         Margin="15,0,15,15"
                         Background="Transparent"
                         BorderThickness="0"
                         ItemContainerStyle="{StaticResource FileListItem}">
                    <ListBox.ItemTemplate>
                        <DataTemplate>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="📄" Margin="0,0,10,0"/>
                                <TextBlock Grid.Column="1" Text="{Binding FileName}" 
                                           TextTrimming="CharacterEllipsis"
                                           VerticalAlignment="Center"/>
                                <TextBlock Grid.Column="2" Text="{Binding Status}" 
                                           Margin="10,0,0,0"
                                           FontWeight="SemiBold"/>
                            </Grid>
                        </DataTemplate>
                    </ListBox.ItemTemplate>
                </ListBox>
                
                <!-- 拖拽提示 -->
                <Border x:Name="DropOverlay" 
                        Background="#E5F3FF" 
                        CornerRadius="8"
                        BorderBrush="#0078D4"
                        BorderThickness="2"
                        Visibility="Collapsed">
                    <TextBlock Text="⬇️ 拖拽文件到此处" 
                               FontSize="24" 
                               HorizontalAlignment="Center" 
                               VerticalAlignment="Center"
                               Foreground="#0078D4"/>
                </Border>
            </Grid>
        </Border>
        
        <!-- 进度区域 -->
        <StackPanel Grid.Row="2" Margin="0,15,0,15">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="安装进度：" 
                           FontSize="14" 
                           VerticalAlignment="Center"
                           Foreground="#202020"/>
                <ProgressBar Grid.Column="1" 
                             x:Name="ProgressBar" 
                             Style="{StaticResource FluentProgressBar}"
                             Margin="10,0"/>
                <TextBlock Grid.Column="2" 
                           x:Name="ProgressText" 
                           Text="0 / 0" 
                           FontSize="14"
                           Foreground="#606060"/>
            </Grid>
            <TextBlock x:Name="CurrentFileText" 
                       Text="当前文件：-" 
                       FontSize="12" 
                       Foreground="#606060"
                       Margin="0,8,0,0"
                       TextTrimming="CharacterEllipsis"/>
        </StackPanel>
        
        <!-- 操作按钮区域 -->
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="BtnInstall" 
                    Content="▶️ 开始安装" 
                    Style="{StaticResource FluentButton}"
                    IsEnabled="False"
                    Width="150"/>
            <Button x:Name="BtnExit" 
                    Content="退出" 
                    Style="{StaticResource FluentButton}"
                    Background="#E8E8E8" 
                    Foreground="#202020"
                    Margin="10,0,0,0"
                    Width="100"/>
        </StackPanel>
    </Grid>
</Window>
"@

# 加载 XAML
try {
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    Write-Host "❌ 加载 XAML 界面失败：$_" -ForegroundColor Red
    exit 1
}

# 全局变量
$script:PackageFiles = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$script:IsInstalling = $false

# 获取控件引用
$btnAddFiles = $window.FindName("BtnAddFiles")
$btnAddFolder = $window.FindName("BtnAddFolder")
$btnClear = $window.FindName("BtnClear")
$btnInstall = $window.FindName("BtnInstall")
$btnExit = $window.FindName("BtnExit")
$fileList = $window.FindName("FileList")
$txtFileCount = $window.FindName("TxtFileCount")
$progressBar = $window.FindName("ProgressBar")
$progressText = $window.FindName("ProgressText")
$currentFileText = $window.FindName("CurrentFileText")
$dropZone = $window.FindName("DropZone")
$dropOverlay = $window.FindName("DropOverlay")

# 支持的扩展名
$supportedExtensions = @(".appx", ".msix", ".appxbundle", ".msixbundle")

# 更新文件计数显示
function Update-FileCount {
    $count = $PackageFiles.Count
    $txtFileCount.Text = "$count 个文件待安装"
    $btnInstall.IsEnabled = ($count -gt 0) -and (-not $IsInstalling)
}

# 添加文件到列表
function Add-FilesToList {
    param([string[]]$Paths)
    
    foreach ($path in $Paths) {
        $ext = [System.IO.Path]::GetExtension($path).ToLower()
        if ($supportedExtensions -contains $ext) {
            # 检查是否已存在
            $exists = $false
            foreach ($item in $PackageFiles) {
                if ($item.Path -eq $path) {
                    $exists = $true
                    break
                }
            }
            if (-not $exists) {
                $PackageFiles.Add([PSCustomObject]@{
                    Path = $path
                    FileName = [System.IO.Path]::GetFileName($path)
                    Status = "⏳ 等待中"
                })
            }
        }
    }
    Update-FileCount
}

# 添加文件按钮事件
$btnAddFiles.Add_Click({
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.Filter = "Appx 包文件|*.appx;*.msix;*.appxbundle;*.msixbundle|所有文件|*.*"
    $dialog.Multiselect = $true
    $dialog.Title = "选择要安装的 Appx 包文件"
    
    if ($dialog.ShowDialog() -eq $true) {
        Add-FilesToList -Paths $dialog.Filenames
    }
})

# 添加文件夹按钮事件
$btnAddFolder.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "选择包含 Appx 包文件的文件夹"
    $dialog.ShowNewFolderButton = $false
    
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $files = Get-ChildItem -Path $dialog.SelectedPath -File | 
                 Where-Object { $supportedExtensions -contains $_.Extension.ToLower() } |
                 Select-Object -ExpandProperty FullName
        Add-FilesToList -Paths $files
    }
})

# 清空列表按钮事件
$btnClear.Add_Click({
    if (-not $IsInstalling) {
        $PackageFiles.Clear()
        Update-FileCount
    }
})

# 退出按钮事件
$btnExit.Add_Click({
    if ($IsInstalling) {
        $result = [System.Windows.MessageBox]::Show(
            "正在安装中，确定要退出吗？",
            "确认退出",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            $window.Close()
        }
    } else {
        $window.Close()
    }
})

# 拖拽功能事件
$dropZone.Add_DragOver({
    if ($_.Data.GetDataPresent([Windows.DataObject]::FileDrop)) {
        $_.Effects = [Windows.DragDropEffects]::Copy
        $dropOverlay.Visibility = [System.Windows.Visibility]::Visible
    } else {
        $_.Effects = [Windows.DragDropEffects]::None
    }
    $_.Handled = $true
})

$dropZone.Add_DragLeave({
    $dropOverlay.Visibility = [System.Windows.Visibility]::Collapsed
})

$dropZone.Add_Drop({
    $dropOverlay.Visibility = [System.Windows.Visibility]::Collapsed
    $files = $_.Data.GetData([Windows.DataObject]::FileDrop)
    if ($files) {
        Add-FilesToList -Paths $files
    }
    $_.Handled = $true
})

# 安装函数
function Install-Package {
    param([string]$Path)
    
    try {
        # 使用 Add-AppxPackage 安装
        Add-AppxPackage -Path $Path -ErrorAction Stop
        return $true
    } catch {
        Write-Host "❌ 安装失败：$_" -ForegroundColor Red
        return $false
    }
}

# 更新列表项状态
function Update-ItemStatus {
    param([string]$Path, [string]$Status)
    
    foreach ($item in $PackageFiles) {
        if ($item.Path -eq $Path) {
            $item.Status = $Status
            break
        }
    }
}

# 开始安装按钮事件
$btnInstall.Add_Click({
    if ($PackageFiles.Count -eq 0 -or $IsInstalling) {
        return
    }
    
    $IsInstalling = $true
    $btnInstall.IsEnabled = $false
    $btnAddFiles.IsEnabled = $false
    $btnAddFolder.IsEnabled = $false
    $btnClear.IsEnabled = $false
    
    $total = $PackageFiles.Count
    $success = 0
    $failed = 0
    $current = 0
    
    # 检查开发者模式
    try {
        $devModeKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" -Name "AllowDevelopmentWithoutDevLicense" -ErrorAction Stop
        if ($devModeKey.AllowDevelopmentWithoutDevLicense -ne 1) {
            $result = [System.Windows.MessageBox]::Show(
                "检测到未开启开发者模式，可能导致安装失败。`n`n是否继续安装？",
                "开发者模式提示",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning
            )
            if ($result -eq [System.Windows.MessageBoxResult]::No) {
                $IsInstalling = $false
                $btnInstall.IsEnabled = $true
                $btnAddFiles.IsEnabled = $true
                $btnAddFolder.IsEnabled = $true
                $btnClear.IsEnabled = $true
                return
            }
        }
    } catch {
        # 注册表项不存在，继续安装
    }
    
    foreach ($package in $PackageFiles) {
        $current++
        $progress = [math]::Round(($current / $total) * 100)
        
        # 更新 UI
        $progressBar.Value = $progress
        $progressText.Text = "$current / $total ($progress%)"
        $currentFileText.Text = "当前文件：$($package.FileName)"
        
        Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
        Write-Host "[$current/$total] 正在安装：$($package.FileName)" -ForegroundColor Cyan
        
        # 检查文件是否存在
        if (-not (Test-Path $package.Path)) {
            Update-ItemStatus -Path $package.Path -Status "❌ 文件不存在"
            Write-Host "❌ 文件不存在：$($package.Path)" -ForegroundColor Red
            $failed++
            continue
        }
        
        # 执行安装
        $result = Install-Package -Path $package.Path
        
        if ($result) {
            Update-ItemStatus -Path $package.Path -Status "✅ 成功"
            Write-Host "✅ 安装成功：$($package.FileName)" -ForegroundColor Green
            $success++
        } else {
            Update-ItemStatus -Path $package.Path -Status "❌ 失败"
            $failed++
        }
        
        # 刷新列表显示
        $fileList.Items.Refresh()
    }
    
    # 安装完成
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    Write-Host "安装完成！成功：$success, 失败：$failed" -ForegroundColor $(if ($failed -eq 0) { "Green" } else { "Yellow" })
    
    $progressText.Text = "完成 - 成功：$success, 失败：$failed"
    $currentFileText.Text = "所有任务已完成"
    
    [System.Windows.MessageBox]::Show(
        "批量安装完成！`n`n成功：$success`n失败：$failed",
        "安装完成",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    )
    
    $IsInstalling = $false
    Update-FileCount
    $btnAddFiles.IsEnabled = $true
    $btnAddFolder.IsEnabled = $true
    $btnClear.IsEnabled = $true
})

# 显示窗口
$window.ShowDialog()
