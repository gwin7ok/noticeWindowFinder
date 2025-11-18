param(
    [string]$Title = 'Google Chrome',
    [string]$Line1 = '🔴 ライブ配信が開始されました',
    [string]$Line2 = 'www.youtube.com - ExampleChannel',
    [string]$Attribution = 'www.youtube.com'
)

Write-Output '=== send_burnttoast_attribution.ps1 ==='
Write-Output "Title=$Title Line1=$Line1 Line2=$Line2 Attribution=$Attribution"

Write-Output 'Importing BurntToast...'
try {
    Import-Module BurntToast -ErrorAction Stop
    Write-Output 'BurntToast imported.'
} catch {
    Write-Output 'Failed to import BurntToast: ' + $_.Exception.Message
    Write-Output 'Attempting to install BurntToast (CurrentUser)...'
    try {
        Install-Module -Name BurntToast -Scope CurrentUser -Force -ErrorAction Stop
        Import-Module BurntToast -ErrorAction Stop
        Write-Output 'BurntToast installed and imported.'
    } catch {
        Write-Output 'Failed to install/import BurntToast: ' + $_.Exception.Message
        exit 1
    }
}

Write-Output 'Sending BurntToast notification with -Attribution...'
try {
    New-BurntToastNotification -Text $Title, $Line1, $Line2 -Attribution $Attribution -ErrorAction Stop
    Write-Output 'Sent BurntToast notification (with Attribution).'
    exit 0
} catch {
    Write-Output 'Failed to send BurntToast notification: ' + $_.Exception.Message
    exit 1
}
