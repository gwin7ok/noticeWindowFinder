param(
    [string]$AppId = $null
)

Write-Output '=== send_test_toast_adopted_route.ps1 ==='
Write-Output "AppId=$AppId"

# Simple BurntToast baseline (if available)
try {
    if (Get-Command New-BurntToastNotification -ErrorAction SilentlyContinue) {
        Write-Output 'Sending baseline BurntToast -Text...'
        New-BurntToastNotification -Text 'Google Chrome', '🔴 ライブ配信が開始されました', 'www.youtube.com - ExampleChannel' -ErrorAction Stop
        Write-Output 'Baseline BurntToast sent.'
    } else {
        Write-Output 'BurntToast command not available; skipping baseline send.'
    }
} catch {
    Write-Output 'Baseline BurntToast send failed: ' + $_.Exception.Message
}

# WinRT send with explicit <attribution>
Write-Output 'Sending WinRT toast with <attribution>www.youtube.com</attribution>...'
try {
    $title = 'Google Chrome'
    $line1 = '🔴 ライブ配信が開始されました'
    $attribution = 'www.youtube.com'
    $xml = "<toast><visual><binding template='ToastGeneric'><text>$title</text><text>$line1</text><attribution>$attribution</attribution></binding></visual></toast>"
    $doc = New-Object -TypeName 'Windows.Data.Xml.Dom.XmlDocument' -ErrorAction Stop
    $doc.LoadXml($xml)
    $toastType = [Windows.UI.Notifications.ToastNotification]
    $managerType = [Windows.UI.Notifications.ToastNotificationManager]
    $toast = $toastType::new($doc)
    if ([string]::IsNullOrEmpty($AppId)) {
        $notifier = $managerType::CreateToastNotifier()
    } else {
        $notifier = $managerType::CreateToastNotifier($AppId)
    }
    $notifier.Show($toast)
    Write-Output 'WinRT toast send attempted (no exception).' 
} catch {
    Write-Output 'WinRT toast send failed: ' + $_.Exception.Message
    exit 1
}

Write-Output 'Done.'
