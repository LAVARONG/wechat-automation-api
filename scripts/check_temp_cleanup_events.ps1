param(
    [datetime]$CenterTime = [datetime]'2026-06-27 14:27:53.997',
    [int]$BeforeSeconds = 60,
    [int]$AfterSeconds = 60,
    [string]$OutputPath = ""
)

$ErrorActionPreference = 'Stop'

$start = $CenterTime.AddSeconds(-$BeforeSeconds)
$end = $CenterTime.AddSeconds($AfterSeconds)

$logs = @(
    'Application',
    'System',
    'Microsoft-Windows-TaskScheduler/Operational',
    'Microsoft-Windows-Windows Defender/Operational'
)

$pattern = '_MEI|Temp|certifi|cacert|delete|cleanup|Defender|Storage|清理|Temporary'

function Write-Section {
    param(
        [string]$Title
    )

    Write-Host ''
    Write-Host "===== $Title =====" -ForegroundColor Cyan
}

function Format-Events {
    param(
        [string]$LogName,
        [array]$Events
    )

    if (-not $Events -or $Events.Count -eq 0) {
        "[$LogName] No matching events found."
        return
    }

    foreach ($event in $Events) {
        @(
            "TimeCreated      : $($event.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss.fff'))"
            "LogName          : $LogName"
            "ProviderName     : $($event.ProviderName)"
            "Id               : $($event.Id)"
            "LevelDisplayName : $($event.LevelDisplayName)"
            "Message          :"
            $event.Message
            ('-' * 80)
        ) -join [Environment]::NewLine
    }
}

Write-Host "CenterTime   : $($CenterTime.ToString('yyyy-MM-dd HH:mm:ss.fff'))" -ForegroundColor Yellow
Write-Host "StartTime    : $($start.ToString('yyyy-MM-dd HH:mm:ss.fff'))" -ForegroundColor Yellow
Write-Host "EndTime      : $($end.ToString('yyyy-MM-dd HH:mm:ss.fff'))" -ForegroundColor Yellow
Write-Host "KeywordRegex : $pattern" -ForegroundColor Yellow

$allOutput = New-Object System.Collections.Generic.List[string]
$allOutput.Add("CenterTime   : $($CenterTime.ToString('yyyy-MM-dd HH:mm:ss.fff'))")
$allOutput.Add("StartTime    : $($start.ToString('yyyy-MM-dd HH:mm:ss.fff'))")
$allOutput.Add("EndTime      : $($end.ToString('yyyy-MM-dd HH:mm:ss.fff'))")
$allOutput.Add("KeywordRegex : $pattern")

foreach ($logName in $logs) {
    Write-Section -Title $logName
    $allOutput.Add('')
    $allOutput.Add("===== $logName =====")

    try {
        $events = Get-WinEvent -FilterHashtable @{
            LogName   = $logName
            StartTime = $start
            EndTime   = $end
        } | Where-Object {
            $_.Message -match $pattern
        }

        $formatted = Format-Events -LogName $logName -Events $events
        $formatted | ForEach-Object { Write-Host $_ }
        $formatted | ForEach-Object { $allOutput.Add($_) }
    }
    catch {
        $errorText = "[$logName] Query failed: $($_.Exception.Message)"
        Write-Host $errorText -ForegroundColor Red
        $allOutput.Add($errorText)
    }
}

if ($OutputPath) {
    $directory = Split-Path -Parent $OutputPath
    if ($directory -and -not (Test-Path $directory)) {
        New-Item -ItemType Directory -Path $directory | Out-Null
    }

    $allOutput | Out-File -FilePath $OutputPath -Encoding utf8
    Write-Host ''
    Write-Host "Exported to: $OutputPath" -ForegroundColor Green
}
