Write-Output '=== Sending test toast (no ToastCloser) ==='

Write-Output 'Checking BurntToast module...'
$m = Get-Module -ListAvailable -Name BurntToast
if ($m) {
    Write-Output 'BurntToast found:'
    $m | Format-Table -AutoSize
} else {
    Write-Output 'BurntToast not found'
}

if (-not $m) {
    Write-Output 'Installing BurntToast (CurrentUser) - may prompt for NuGet/repository trust...'
    Install-Module -Name BurntToast -Scope CurrentUser -Force -ErrorAction SilentlyContinue
}

Import-Module BurntToast -ErrorAction SilentlyContinue
Write-Output 'Module import result:'
Get-Module -Name BurntToast -ErrorAction SilentlyContinue | Format-Table -AutoSize

Write-Output 'Attempting WinRT raw reminder toast (may work even if BurntToast lacks -Scenario)...'
function Send-ReminderToast-WinRT {
    param($title, $line1, $line2, $appId = $null)
    try {
        $xml = "<toast scenario='reminder'><visual><binding template='ToastGeneric'><text>$title</text><text>$line1</text><text>$line2</text></binding></visual></toast>"
        $docType = 'Windows.Data.Xml.Dom.XmlDocument'
        $doc = New-Object -TypeName $docType -ErrorAction Stop
        $doc.LoadXml($xml)
        $toastType = [Windows.UI.Notifications.ToastNotification]
        $notifierType = [Windows.UI.Notifications.ToastNotificationManager]
        $toast = $toastType::new($doc)
        if ([string]::IsNullOrEmpty($appId)) {
            # Use default notifier for current process (no AppUserModelID)
            $notifier = $notifierType::CreateToastNotifier()
        } else {
            $notifier = $notifierType::CreateToastNotifier($appId)
        }
        $notifier.Show($toast)
        Write-Output 'Sent reminder toast via WinRT.'
        return $true
    } catch {
        Write-Output 'WinRT reminder send failed: ' + $_.Exception.Message
        return $false
    }
}

 $title = 'Google Chrome'
 $line1 = '🔴 ライブ配信が開始されました'
 $line2 = 'www.youtube.com - ExampleChannel'
 if (Send-ReminderToast-WinRT -title $title -line1 $line1 -line2 $line2) {
    Write-Output 'WinRT reminder toast sent; exiting.'
    return
 } else {
    Write-Output 'WinRT approach failed; falling back to BurntToast method.'
 }

if (Get-Command New-BurntToastNotification -ErrorAction SilentlyContinue) {
    Write-Output 'Sending test notification now (try with -Scenario first)...'
    $title = 'Google Chrome'
    $line1 = '🔴 ライブ配信が開始されました'
    $line2 = 'www.youtube.com - ExampleChannel'
    $sent = $false
    try {
        New-BurntToastNotification -Text $title, $line1, $line2 -Scenario Reminder -ErrorAction Stop
        Write-Output 'Sent BurntToast notification with -Scenario.'
        $sent = $true
    } catch {
        Write-Output 'Sending with -Scenario failed: ' + $_.Exception.Message
        Write-Output 'Retrying without -Scenario...'
        try {
            New-BurntToastNotification -Text $title, $line1, $line2 -ErrorAction Stop
            Write-Output 'Sent BurntToast notification without -Scenario.'
            $sent = $true
        } catch {
            Write-Output 'Retry without -Scenario also failed: ' + $_.Exception.Message
        }
    }
    if (-not $sent) { Write-Output 'Failed to send notification.' }
} else {
    Write-Output 'New-BurntToastNotification command not available. Install/import may have failed.'
}
