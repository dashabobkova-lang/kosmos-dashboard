Write-Host '=== PODKLYUCHENIE DISKA Z (Bitrix) ===' -ForegroundColor Cyan
Write-Host 'Vvedite login i parol ot Bitrix24 (portal bylulu.bitrix24.ru)' -ForegroundColor Cyan
Write-Host ''
net use Z: /delete /y | Out-Null
net use Z: https://bylulu.bitrix24.ru/docs/path/ /persistent:yes
Write-Host ''
if (Test-Path 'Z:\') {
    Write-Host 'OK: disk Z podklyuchen' -ForegroundColor Green
} else {
    Write-Host 'ERROR: disk ne podklyuchen' -ForegroundColor Red
}
