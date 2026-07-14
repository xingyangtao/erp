# VB_ERP 调试测试脚本
# 测试调试配置是否正常工作

Write-Host "=== VB_ERP 调试配置测试 ===" -ForegroundColor Green

# 检查 .NET SDK
Write-Host "`n1. 检查 .NET SDK 版本:" -ForegroundColor Yellow
dotnet --version

# 检查项目文件
Write-Host "`n2. 检查项目文件:" -ForegroundColor Yellow
if (Test-Path "VB_ERP/VB_ERP.vbproj") {
    Write-Host "✓ 项目文件存在" -ForegroundColor Green
} else {
    Write-Host "✗ 项目文件不存在" -ForegroundColor Red
    exit 1
}

# 检查调试配置
Write-Host "`n3. 检查调试配置文件:" -ForegroundColor Yellow
$debugFiles = @(
    ".vscode/launch.json",
    ".vscode/tasks.json",
    "VB_ERP/Properties/launchSettings.json",
    "VB_ERP/App.config"
)

foreach ($file in $debugFiles) {
    if (Test-Path $file) {
        Write-Host "✓ $file" -ForegroundColor Green
    } else {
        Write-Host "✗ $file" -ForegroundColor Red
    }
}

# 测试构建
Write-Host "`n4. 测试项目构建:" -ForegroundColor Yellow
Set-Location "VB_ERP"
$buildResult = dotnet build --configuration Debug 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "✓ 构建成功" -ForegroundColor Green
} else {
    Write-Host "✗ 构建失败" -ForegroundColor Red
    Write-Host $buildResult
}

# 检查输出文件
Write-Host "`n5. 检查构建输出:" -ForegroundColor Yellow
if (Test-Path "bin/Debug/net8.0-windows/VB_ERP.dll") {
    Write-Host "✓ 调试输出文件存在" -ForegroundColor Green
} else {
    Write-Host "✗ 调试输出文件不存在" -ForegroundColor Red
}

# 检查启动配置
Write-Host "`n6. 检查启动配置:" -ForegroundColor Yellow
$launchConfig = Get-Content "../.vscode/launch.json" | ConvertFrom-Json
Write-Host "可用调试配置:" -ForegroundColor Cyan
foreach ($config in $launchConfig.configurations) {
    Write-Host "  - $($config.name)" -ForegroundColor White
}

Write-Host "`n=== 测试完成 ===" -ForegroundColor Green
Write-Host "调试配置已准备就绪！" -ForegroundColor Cyan
Write-Host "使用 VS Code 按 F5 开始调试" -ForegroundColor White
Write-Host "使用 Visual Studio 2022 打开 VB_ERP.vbproj 开始调试" -ForegroundColor White