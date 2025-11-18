param(
    [string]$AppId = $null
)

Write-Output '=== send_test_toast_dual_labeled.ps1 ==='
Write-Output "AppId=$AppId"

# Payload base
$baseTitle = 'Google Chrome'
$baseLine1 = '🔴 ライブ配信が開始されました'
$baseLine2 = 'www.youtube.com - ExampleChannel'
$attribution = 'www.youtube.com'

function Send-BurntToast-Labeled {
    param($label)
    try {
        $title = "$($baseTitle) ($label)"
        Write-Output "Sending BurntToast labeled: $label"
        if (Get-Command New-BurntToastNotification -ErrorAction SilentlyContinue) {
            New-BurntToastNotification -Text $title, $baseLine1, $baseLine2 -ErrorAction Stop
            Write-Output 'BurntToast send succeeded.'
            return $true
        } else {
            Write-Output 'BurntToast not available.'
            return $false
        }
    } catch {
        Write-Output 'BurntToast send failed: ' + $_.Exception.Message
        return $false
    }
}

function Send-WinRT-Labeled {
    param($label)
    try {
        $title = "$($baseTitle) ($label)"
        Write-Output "Sending WinRT labeled: $label"
        $xml = "<toast><visual><binding template='ToastGeneric'><text>$title</text><text>$baseLine1</text><attribution>$attribution</attribution></binding></visual></toast>"
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
        Write-Output 'WinRT send attempted (no exception).' 
        return $true
    } catch {
        Write-Output 'WinRT send failed: ' + $_.Exception.Message
        return $false
    }
}

# 1) Send BurntToast labeled
Send-BurntToast-Labeled -label 'BurntToast'

Start-Sleep -Milliseconds 500

# 2) Send WinRT labeled
Send-WinRT-Labeled -label 'WinRT'

Write-Output 'Done.'
