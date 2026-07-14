# VB_ERP 本地调试配置指南

## 概述

本文档介绍了如何配置和使用 VB_ERP 项目的本地调试环境。项目已配置为支持 VS Code 和 Visual Studio 2022 进行调试。

## 快速开始

### 1. 环境要求

- **.NET 8.0 SDK** 或更高版本
- **Visual Studio 2022** (推荐) 或 **VS Code** + C# 扩展
- **MySQL 8.0** 数据库服务器 (localhost:3306)

### 2. 使用 VS Code 调试

#### 安装必要的扩展
1. 打开 VS Code
2. 安装以下扩展：
   - **C# Dev Kit** (ms-dotnettools.csdevkit)
   - **C#** (ms-dotnettools.csharp)

#### 调试步骤
1. 在 VS Code 中打开 `e:\ERPV4` 文件夹
2. 按 `F5` 或点击调试菜单中的 "Launch VB_ERP"
3. 选择调试配置：
   - **Launch VB_ERP** - 调试 Debug 配置
   - **Launch VB_ERP (Release)** - 调试 Release 配置
   - **Launch VB_ERP (with arguments)** - 带参数调试

#### 可用任务
- `Ctrl+Shift+B` - 执行默认构建任务
- `Tasks: Run Task` - 选择其他任务：
  - `build` - 构建 Debug 配置
  - `build-release` - 构建 Release 配置
  - `clean` - 清理项目
  - `restore` - 还原 NuGet 包
  - `watch` - 启用文件监视

### 3. 使用 Visual Studio 2022 调试

1. 打开 Visual Studio 2022
2. 选择 "打开项目或解决方案"
3. 导航到 `e:\ERPV4\VB_ERP` 并选择 `VB_ERP.vbproj`
4. 在工具栏中选择调试配置 (Debug/Release)
5. 按 `F5` 开始调试

### 4. 调试配置详解

#### launch.json 配置
- **类型**: `coreclr` (.NET Core 调试器)
- **程序**: `VB_ERP/bin/Debug/net8.0-windows/VB_ERP.dll`
- **环境变量**: `DOTNET_ENVIRONMENT=Development`
- **控制台**: 内部控制台

#### 任务配置
- **预调试任务**: `build` - 自动在调试前构建项目
- **问题匹配器**: `$msCompile` - 自动识别编译错误

## 高级调试技巧

### 1. 断点调试
- 在代码行号左侧点击设置断点
- 支持条件断点、命中计数断点、命中时断点
- 使用 `F10` 单步跳过，`F11` 单步进入

### 2. 监视变量
- 在调试时使用 "监视" 窗口
- 右键变量选择 "添加到监视"
- 支持表达式监视

### 3. 调试控制台
- 在调试时使用集成终端
- 查看应用程序输出
- 执行调试命令

### 4. 远程调试
如需远程调试，修改 `launch.json` 中的 `program` 路径指向远程机器的 DLL 路径。

## 数据库调试

### 连接字符串
项目使用两个 MySQL 数据库：
1. **认证数据库**: `erpshouquan`
   - 服务器: localhost:3306
   - 用户名: erpshouquan
   - 密码: erpshouquan

2. **业务数据库**: `dldata`
   - 服务器: localhost:3306
   - 用户名: dldata
   - 密码: yt19880925!@#

### 调试数据库连接
1. 在 `DatabaseModule.vb` 中设置断点
2. 检查连接字符串和连接状态
3. 使用 `App.config` 中的配置

## 常见问题解决

### 1. 构建失败
```bash
# 清理并重新构建
dotnet clean
dotnet restore
dotnet build
```

### 2. 调试器无法启动
- 确保项目已成功构建
- 检查 `VB_ERP.dll` 是否存在于 `bin/Debug/net8.0-windows/` 目录
- 确认 .NET 8.0 SDK 已安装

### 3. 数据库连接失败
- 确保 MySQL 服务正在运行
- 检查 `App.config` 或 `DatabaseModule.vb` 中的连接字符串
- 验证数据库用户权限

### 4. 命名空间错误
如果遇到命名空间错误，检查项目文件中的 `<RootNamespace>VB_ERP</RootNamespace>` 配置。

## 调试日志

### 启用详细日志
在 `App.config` 中启用跟踪：
```xml
<system.diagnostics>
  <switches>
    <add name="DatabaseTrace" value="4" />
  </switches>
</system.diagnostics>
```

### 输出调试信息
使用 `Debug.WriteLine()` 或 `Trace.WriteLine()` 输出调试信息：
```vb
Debug.WriteLine("数据库连接状态: " & connection.State)
Trace.WriteLine("用户登录: " & username)
```

## 性能调试

### 性能分析器
1. 使用 Visual Studio 的性能分析器
2. 选择 "CPU 使用率" 或 ".NET 对象分配"
3. 运行应用程序并分析性能瓶颈

### 内存调试
1. 使用 "内存使用率" 工具
2. 检查对象分配和释放
3. 识别内存泄漏

## 部署调试

### 发布配置
```bash
dotnet publish --configuration Release --output ./publish
```

### 远程部署调试
1. 在目标机器上安装 .NET 8.0 运行时
2. 复制发布文件夹
3. 使用远程调试器连接

## 支持与反馈

如遇到调试问题，请检查：
1. 项目文档 `VB.NET开发说明文档.md`
2. 系统配置文件 `配置/系统.config.json`
3. 数据库配置

## 更新日志

**2026-07-13**: 初始调试配置创建
- 添加 VS Code 调试配置
- 添加 Visual Studio 启动配置
- 添加构建任务配置
- 添加 App.config 配置文件