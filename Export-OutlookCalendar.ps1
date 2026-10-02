<#
.SYNOPSIS
  Export the default Outlook calendar for a date range to CSV (recurring meetings expanded).

.DESCRIPTION
  Uses the classic Outlook desktop app through COM, so it needs no admin rights,
  app registration or Microsoft 365 connector. The "new Outlook" app has no COM
  interface; see the README for the .ics fallback.

.EXAMPLE
  .\Export-OutlookCalendar.ps1 -Start 2026-09-01 -End 2026-10-01 -OutFile calendar.csv
#>
param(
    [Parameter(Mandatory)][datetime]$Start,
    [Parameter(Mandatory)][datetime]$End,   # exclusive
    [string]$OutFile = "calendar.csv",
    [switch]$IncludeBody
)

$ErrorActionPreference = "Stop"
$showAs = @{ 0 = "Free"; 1 = "Tentative"; 2 = "Busy"; 3 = "OutOfOffice"; 4 = "WorkingElsewhere" }
$response = @{ 0 = "None"; 1 = "Organizer"; 2 = "Tentative"; 3 = "Accepted"; 4 = "Declined"; 5 = "NotResponded" }

$outlook = New-Object -ComObject Outlook.Application
$calendar = $outlook.GetNamespace("MAPI").GetDefaultFolder(9)  # olFolderCalendar
$items = $calendar.Items
# Sort before IncludeRecurrences, then Restrict: that's what makes Outlook
# return each occurrence of a recurring meeting instead of the series master.
$items.Sort("[Start]")
$items.IncludeRecurrences = $true
# Restrict expects dates in the machine's own short date/time format.
$filter = "[Start] >= '{0}' AND [Start] < '{1}'" -f $Start.ToString("g"), $End.ToString("g")

$rows = foreach ($item in $items.Restrict($filter)) {
    $row = [ordered]@{
        Subject    = $item.Subject
        Start      = $item.Start.ToString("yyyy-MM-ddTHH:mm:ss")
        End        = $item.End.ToString("yyyy-MM-ddTHH:mm:ss")
        AllDay     = [string]$item.AllDayEvent
        ShowAs     = $showAs[[int]$item.BusyStatus]
        Response   = $response[[int]$item.ResponseStatus]
        Location   = $item.Location
        Organizer  = $item.Organizer
        Categories = $item.Categories
        Recurring  = [string]$item.IsRecurring
    }
    if ($IncludeBody) {
        $body = ($item.Body -replace '\s+', ' ').Trim()
        $row.Body = $body.Substring(0, [Math]::Min(300, $body.Length))
    }
    [pscustomobject]$row
}

$rows | Export-Csv -Path $OutFile -NoTypeInformation -Encoding UTF8
Write-Output "Exported $(@($rows).Count) calendar items to $OutFile"
