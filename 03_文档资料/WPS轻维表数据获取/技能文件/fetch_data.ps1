# PowerShell 脚本：调用 WPS AirScript Webhook 并保存数据
# 用法：.\fetch_data.ps1 -WebhookUrl "..." -Token "..." -OutFile "output.json"

param(
    [Parameter(Mandatory=$true)]
    [string]$WebhookUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$Token,
    
    [string]$OutFile = "wps_data_raw.json",
    
    [int]$TimeoutSec = 120
)

Write-Host "正在调用 WPS AirScript Webhook..." -ForegroundColor Cyan
Write-Host "URL: $WebhookUrl"

try {
    $response = Invoke-RestMethod `
        -Uri    $WebhookUrl `
        -Method POST `
        -Headers @{
            "Content-Type"    = "application/json"
            "AirScript-Token" = $Token
        } `
        -Body '{"Context":{"argv":{}}}' `
        -TimeoutSec $TimeoutSec

    $result = $response.data.result
    
    if ($result.success) {
        Write-Host "获取成功！总条数: $($result.total)" -ForegroundColor Green
        $result.data | ConvertTo-Json -Depth 10 | Out-File $OutFile -Encoding utf8
        Write-Host "数据已保存到: $OutFile" -ForegroundColor Green
    } else {
        Write-Host "脚本执行失败: $($result.error)" -ForegroundColor Red
    }

} catch {
    Write-Host "调用失败: $_" -ForegroundColor Red
}
