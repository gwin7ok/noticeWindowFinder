param(
    [bool]$BurntToastOnly = $false,
    [bool]$EnsureAttribution = $true,
    [string]$AppId = $null
)

Write-Output '=== VERBOSE ONESHOT: send_test_toast_verbose_once.ps1 ==='
Write-Output "Parameters: BurntToastOnly=$BurntToastOnly EnsureAttribution=$EnsureAttribution AppId=$AppId"

function LogStage { param($m) Write-Output "[STAGE] $m" }
function LogDetail { param($m) Write-Output "  - $m" }

LogStage 'Checking BurntToast availability'
$m = Get-Module -ListAvailable -Name BurntToast
if ($m) { LogDetail 'BurntToast found'; $m | Format-Table -AutoSize } else { LogDetail 'BurntToast not found' }

LogStage 'Attempting to import BurntToast (if present)'
try { Import-Module BurntToast -ErrorAction Stop; LogDetail 'Import-Module BurntToast succeeded' } catch { LogDetail "Import failed: $($_.Exception.Message)" }

LogStage 'Preparing payload'
$Title = 'Google Chrome'
$Line1 = '🔴 ライブ配信が開始されました'
$Line2 = 'www.youtube.com - ExampleChannel'
$Attribution = 'www.youtube.com'
LogDetail "Title='$Title' Line1='$Line1' Line2='$Line2' Attribution='$Attribution'"

if (-not $BurntToastOnly) {
    LogStage 'Baseline: sending simple BurntToast -Text (if available)'
    if (Get-Command New-BurntToastNotification -ErrorAction SilentlyContinue) {
        try {
            LogDetail 'Calling New-BurntToastNotification -Text ...'
            New-BurntToastNotification -Text $Title, $Line1, $Line2 -ErrorAction Stop
            LogDetail 'BurntToast -Text succeeded (baseline send)'
        } catch {
            LogDetail "BurntToast -Text failed: $($_.Exception.Message)"
        }
    } else { LogDetail 'New-BurntToastNotification not available' }
} else { LogStage 'BurntToastOnly mode: skipping baseline BurntToast' }

LogStage 'Determining send strategy'
LogDetail "BurntToastOnly=$BurntToastOnly EnsureAttribution=$EnsureAttribution"

function TryWinRT {
    param($title, $line1, $attrib, $appId)
    LogStage 'WINRT: building XML and attempting send'
    $xml = "<toast><visual><binding template='ToastGeneric'><text>$title</text><text>$line1</text><attribution>$attrib</attribution></binding></visual></toast>"
    LogDetail "XML: $xml"
    try {
        LogDetail 'Creating Windows.Data.Xml.Dom.XmlDocument'
        $doc = New-Object -TypeName 'Windows.Data.Xml.Dom.XmlDocument' -ErrorAction Stop
        LogDetail 'XmlDocument created'
        $doc.LoadXml($xml)
        LogDetail 'XML loaded into XmlDocument'
        $toastType = [Windows.UI.Notifications.ToastNotification]
        $managerType = [Windows.UI.Notifications.ToastNotificationManager]
        LogDetail 'Constructing ToastNotification'
        $toast = $toastType::new($doc)
        if ([string]::IsNullOrEmpty($appId)) { LogDetail 'Using default CreateToastNotifier()'; $notifier = $managerType::CreateToastNotifier() } else { LogDetail "Using CreateToastNotifier(AppId=$appId)"; $notifier = $managerType::CreateToastNotifier($appId) }
        LogDetail 'Calling notifier.Show(toast)'
        $notifier.Show($toast)
        LogDetail 'notifier.Show() completed without exception'
        return $true
    } catch {
        LogDetail "WinRT send exception: $($_.Exception.Message)"
        return $false
    }
}

if ($EnsureAttribution) {
    LogStage 'EnsureAttribution requested -> trying WinRT with explicit <attribution>'
    $ok = TryWinRT -title $Title -line1 $Line1 -attrib $Attribution -appId $AppId
    if ($ok) { LogStage 'WinRT send reported success -> DONE'; exit 0 } else { LogStage 'WinRT send failed -> will try BurntToast fallbacks' }
} elseif (-not $BurntToastOnly) {
    LogStage 'No EnsureAttribution: trying WinRT without separate attribution'
    $ok = TryWinRT -title $Title -line1 $Line1 -attrib "$Line2" -appId $AppId
    if ($ok) { LogStage 'WinRT send reported success -> DONE'; exit 0 } else { LogStage 'WinRT failed -> will try BurntToast fallbacks' }
} else { LogStage 'BurntToastOnly: skipping WinRT attempts' }

LogStage 'BurntToast fallbacks: -Scenario, then -Xml with <attribution>, then -Text'
if (Get-Command New-BurntToastNotification -ErrorAction SilentlyContinue) {
    try {
        LogDetail 'Trying New-BurntToastNotification -Scenario Reminder'
        New-BurntToastNotification -Text $Title, $Line1, $Line2 -Scenario Reminder -ErrorAction Stop
        LogDetail 'BurntToast -Scenario succeeded'
        exit 0
    } catch {
        LogDetail "-Scenario failed: $($_.Exception.Message)"
        $xml = "<toast><visual><binding template='ToastGeneric'><text>$Title</text><text>$Line1</text><text>$Line2</text><attribution>$Attribution</attribution></binding></visual></toast>"
        try {
            LogDetail 'Trying New-BurntToastNotification -Xml with <attribution>'
            New-BurntToastNotification -Xml $xml -ErrorAction Stop
            LogDetail 'BurntToast -Xml succeeded'
            exit 0
        } catch {
            LogDetail "-Xml failed: $($_.Exception.Message)"
            try {
                LogDetail 'Falling back to New-BurntToastNotification -Text'
                New-BurntToastNotification -Text $Title, $Line1, $Line2 -ErrorAction Stop
                LogDetail 'BurntToast -Text fallback succeeded'
                exit 0
            } catch {
                LogDetail "BurntToast -Text fallback failed: $($_.Exception.Message)"
                LogStage 'ALL SEND ATTEMPTS FAILED'
                exit 1
            }
        }
    }
} else {
    LogDetail 'New-BurntToastNotification not available -> no fallback'
    LogStage 'ALL SEND ATTEMPTS FAILED (no BurntToast)'
    exit 1
}
