param(
    [string]$AppId = $null,
    [string]$Title = 'Google Chrome',
    [string]$Line1 = '🔴 ライブ配信が開始されました',
    [string]$Attribution = 'www.youtube.com'
)

Write-Output '=== Sending WINRT-only test toast (attribution only) ==='
Write-Output "AppId=$AppId Title=$Title Line1=$Line1 Attribution=$Attribution"

function Send-ReminderToast-WinRTOnly {
    param($title, $line1, $attribution, $appId = $null)
    try {
        $xml = "<toast><visual><binding template='ToastGeneric'><text>$title</text><text>$line1</text><attribution>$attribution</attribution></binding></visual></toast>"
        $docType = 'Windows.Data.Xml.Dom.XmlDocument'
        $doc = New-Object -TypeName $docType -ErrorAction Stop
        $doc.LoadXml($xml)
        $toastType = [Windows.UI.Notifications.ToastNotification]
        $notifierType = [Windows.UI.Notifications.ToastNotificationManager]
        $toast = $toastType::new($doc)
        if ([string]::IsNullOrEmpty($appId)) {
            $notifier = $notifierType::CreateToastNotifier()
        } else {
            $notifier = $notifierType::CreateToastNotifier($appId)
        }
        $notifier.Show($toast)
        Write-Output 'Sent WinRT toast with explicit <attribution>.'
        return $true
    } catch {
        Write-Output 'WinRT send failed: ' + $_.Exception.Message
        return $false
    }
}

$ok = Send-ReminderToast-WinRTOnly -title $Title -line1 $Line1 -attribution $Attribution -appId $AppId
if (-not $ok) { exit 1 } else { exit 0 }
