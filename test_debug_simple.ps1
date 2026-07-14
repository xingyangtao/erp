# VB_ERP 调试配置简单测试

Write-Host "=== VB_ERP 调试配置测试 ===" -ForegroundColor Green

# 检查关键文件
Write-Host "`n检查调试配置文件:" -ForegroundColor Yellow
$files = @(
    ".vscode\launch.json",
    ".vscode\tasks.json",
    "VB_ERP\Properties\launchSettings.json",
    "VB_ERP\App.config",
    "VB_ERP\VB_ERP.vbproj"
)

foreach ($file in $files) {
    if (Test-Path $file) {
        Write-Host "✓ $file" -ForegroundColor Green
    } else {
        Write-Host "✗ $file" -ForegroundColor Red
    }
}

# 测试构建
Write-Host "`n测试项目构建:" -ForegroundColor Yellow
Set-Location "VB_ERP"
$buildResult = dotnet build --configuration Debug 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "✓ 构建成功" -ForegroundColor Green
} else {
    Write-Host "✗ 构建失败" -ForegroundColor Red
}

Write-Host "`n=== 测试完成 ===" -ForegroundColor Green