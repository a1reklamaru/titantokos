param(
  [switch]$NoCommit,
  [switch]$NoPush
)

$ErrorActionPreference = "Stop"

$SpreadsheetId = "1anw9pGO6y7emBZlbvgUJUycLtC54KMz5d-i3vXmnji8"
$ClientsConfig = @(
  @{ Key = "titan"; Gid = "0" },
  @{ Key = "tokos"; Gid = "1190677190" }
)
$DashboardPath = Join-Path $PSScriptRoot "..\titan-tokos-dashboard.html"

$MonthNames = @{
  "января" = 1; "февраля" = 2; "марта" = 3; "апреля" = 4; "мая" = 5; "июня" = 6;
  "июля" = 7; "августа" = 8; "сентября" = 9; "октября" = 10; "ноября" = 11; "декабря" = 12
}

function Convert-ToNumber($Value) {
  if ($null -eq $Value) { return 0 }
  $text = [string]$Value
  $text = $text.Trim()
  if (-not $text -or $text -eq "#DIV/0!") { return 0 }
  $text = $text -replace "[^\d,\.\-]", ""
  if (-not $text) { return 0 }
  $text = $text -replace ",", "."
  $number = 0.0
  if ([double]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
    return $number
  }
  return 0
}

function Convert-ToInt($Value) {
  return [int][Math]::Round((Convert-ToNumber $Value), 0, [MidpointRounding]::AwayFromZero)
}

function Get-FirstFilledCell($Row) {
  foreach ($cell in $Row) {
    if (-not [string]::IsNullOrWhiteSpace($cell)) { return ([string]$cell).Trim() }
  }
  return ""
}

function Read-CsvRows($Path) {
  Add-Type -AssemblyName Microsoft.VisualBasic
  $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($Path, [Text.Encoding]::UTF8)
  $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
  $parser.SetDelimiters(",")
  $parser.HasFieldsEnclosedInQuotes = $true
  $rows = @()
  try {
    while (-not $parser.EndOfData) {
      $rows += ,$parser.ReadFields()
    }
  }
  finally {
    $parser.Close()
  }
  return $rows
}

function Get-PeriodStartDate($Period, $Year) {
  $clean = ($Period -replace "\(.*?\)", "").Trim().ToLowerInvariant()
  if ($clean -notmatch "^\s*(\d{1,2})") {
    throw "Не удалось распознать начало периода: $Period"
  }
  $day = [int]$Matches[1]
  $monthMatches = [regex]::Matches($clean, "(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)")
  if ($monthMatches.Count -eq 0) {
    throw "Не удалось распознать месяц периода: $Period"
  }
  $monthName = $monthMatches[0].Value
  $month = $MonthNames[$monthName]
  return Get-Date -Year $Year -Month $month -Day $day -Hour 0 -Minute 0 -Second 0
}

function New-MetricRow($Period, $Date, $Impressions, $Clicks, $Cost, $Conversions) {
  $impressionsRounded = [int][Math]::Round($Impressions, 0, [MidpointRounding]::AwayFromZero)
  $clicksRounded = [int][Math]::Round($Clicks, 0, [MidpointRounding]::AwayFromZero)
  $costRounded = [Math]::Round($Cost, 2)
  $conversionsRounded = [int][Math]::Round($Conversions, 0, [MidpointRounding]::AwayFromZero)
  $ctr = if ($impressionsRounded) { $clicksRounded / $impressionsRounded } else { 0 }
  $cpc = if ($clicksRounded) { $costRounded / $clicksRounded } else { 0 }
  $cpa = if ($conversionsRounded) { $costRounded / $conversionsRounded } else { 0 }
  return [ordered]@{
    period = $Period
    date = $Date.ToString("yyyy-MM-dd")
    year = [int]$Date.Year
    impressions = $impressionsRounded
    clicks = $clicksRounded
    ctr = [Math]::Round($ctr, 6)
    cost = $costRounded
    cpc = [Math]::Round($cpc, 2)
    conversions = $conversionsRounded
    cpa = [Math]::Round($cpa, 2)
  }
}

function Parse-ClientSheet($Path) {
  $rows = Read-CsvRows $Path
  $currentYear = $null
  $currentPeriod = $null
  $currentDate = $null
  $inCampaignBlock = $false
  $currentCampaignRows = @()
  $weeklyByDate = [ordered]@{}
  $campaignsByDate = [ordered]@{}

  foreach ($row in $rows) {
    $first = Get-FirstFilledCell $row
    if (-not $first) { continue }

    if ($first -match "^\d{4}$") {
      $currentYear = [int]$first
      $inCampaignBlock = $false
      $currentCampaignRows = @()
      continue
    }

    if ($first -eq "Кампания") {
      $inCampaignBlock = $true
      continue
    }

    if ($first -like "Остаток*") {
      $inCampaignBlock = $false
      continue
    }

    if ($first -eq "Итого") {
      if ($null -eq $currentDate) { continue }
      $dateKey = $currentDate.ToString("yyyy-MM-dd")
      $weeklyByDate[$dateKey] = New-MetricRow $currentPeriod $currentDate (Convert-ToInt $row[1]) (Convert-ToInt $row[2]) (Convert-ToNumber $row[5]) (Convert-ToInt $row[6])
      $campaignsByDate[$dateKey] = @($currentCampaignRows)
      $inCampaignBlock = $false
      continue
    }

    if ($inCampaignBlock -and $null -ne $currentDate) {
      $campaignRow = [ordered]@{
        campaign = $first
        date = $currentDate.ToString("yyyy-MM-dd")
        year = [int]$currentDate.Year
        impressions = Convert-ToInt $row[1]
        clicks = Convert-ToInt $row[2]
        ctr = 0
        cost = [Math]::Round((Convert-ToNumber $row[5]), 2)
        cpc = 0
        conversions = Convert-ToInt $row[6]
        cpa = 0
      }
      $campaignRow.ctr = if ($campaignRow.impressions) { [Math]::Round($campaignRow.clicks / $campaignRow.impressions, 6) } else { 0 }
      $campaignRow.cpc = if ($campaignRow.clicks) { [Math]::Round($campaignRow.cost / $campaignRow.clicks, 2) } else { 0 }
      $campaignRow.cpa = if ($campaignRow.conversions) { [Math]::Round($campaignRow.cost / $campaignRow.conversions, 2) } else { 0 }
      $currentCampaignRows += $campaignRow
      continue
    }

    if ($first -match "ИТОГО") {
      $inCampaignBlock = $false
      continue
    }

    if ($null -ne $currentYear -and $first -match "^\d{1,2}" -and $first -match "(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)") {
      $currentPeriod = $first
      $currentDate = Get-PeriodStartDate $first $currentYear
      $currentCampaignRows = @()
      $inCampaignBlock = $false
    }
  }

  $weekly = @()
  foreach ($dateKey in ($weeklyByDate.Keys | Sort-Object)) {
    $weekly += $weeklyByDate[$dateKey]
  }

  $campaignsWeekly = @()
  foreach ($dateKey in ($campaignsByDate.Keys | Sort-Object)) {
    $campaignsWeekly += $campaignsByDate[$dateKey]
  }

  return @{
    weekly = $weekly
    campaignsWeekly = $campaignsWeekly
  }
}

function Add-ToBucket($Map, $Key, $Template, $Weight) {
  if (-not $Map.ContainsKey($Key)) {
    $Map[$Key] = [ordered]@{
      period = $Template.period
      date = $Template.date
      year = $Template.year
      impressions = 0.0
      clicks = 0.0
      cost = 0.0
      conversions = 0.0
    }
    if ($Template.Contains("campaign")) {
      $Map[$Key].campaign = $Template.campaign
    }
  }
  $Map[$Key].impressions += $Template.impressions * $Weight
  $Map[$Key].clicks += $Template.clicks * $Weight
  $Map[$Key].cost += $Template.cost * $Weight
  $Map[$Key].conversions += $Template.conversions * $Weight
}

function Get-RussianMonthLabel($Date) {
  $names = @("Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь")
  return "$($names[$Date.Month - 1]) $($Date.Year)"
}

function Build-MonthlyRows($Rows, [switch]$ByCampaign) {
  $map = @{}
  foreach ($item in $Rows) {
    $weekStart = [datetime]::ParseExact($item.date, "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture)
    for ($i = 0; $i -lt 7; $i++) {
      $day = $weekStart.AddDays($i)
      $monthStart = Get-Date -Year $day.Year -Month $day.Month -Day 1 -Hour 0 -Minute 0 -Second 0
      $monthKey = $monthStart.ToString("yyyy-MM")
      $key = if ($ByCampaign) { "$($item.campaign)|$monthKey" } else { $monthKey }
      $template = [ordered]@{
        period = Get-RussianMonthLabel $monthStart
        date = $monthStart.ToString("yyyy-MM-dd")
        year = [int]$monthStart.Year
        impressions = $item.impressions
        clicks = $item.clicks
        cost = $item.cost
        conversions = $item.conversions
      }
      if ($ByCampaign) { $template.campaign = $item.campaign }
      Add-ToBucket $map $key $template (1 / 7)
    }
  }

  $result = @()
  foreach ($entry in $map.GetEnumerator() | Sort-Object { $_.Value.date }, { $_.Value.campaign }) {
    $value = $entry.Value
    $metric = New-MetricRow $value.period ([datetime]::ParseExact($value.date, "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture)) $value.impressions $value.clicks $value.cost $value.conversions
    if ($ByCampaign) {
      $campaignRow = [ordered]@{ campaign = $value.campaign }
      foreach ($key in $metric.Keys) { $campaignRow[$key] = $metric[$key] }
      $result += $campaignRow
    }
    else {
      $result += $metric
    }
  }
  return $result
}

$clients = [ordered]@{}
$tempFiles = @()
try {
  foreach ($client in $ClientsConfig) {
    $url = "https://docs.google.com/spreadsheets/d/$SpreadsheetId/export?format=csv&gid=$($client.Gid)"
    $tempFile = Join-Path $env:TEMP "dashboard-$($client.Key)-$([guid]::NewGuid()).csv"
    $tempFiles += $tempFile
    Invoke-WebRequest -Uri $url -OutFile $tempFile -UseBasicParsing
    $parsed = Parse-ClientSheet $tempFile
    $clients[$client.Key] = [ordered]@{
      weekly = $parsed.weekly
      monthly = Build-MonthlyRows $parsed.weekly
      campaignsWeekly = $parsed.campaignsWeekly
      campaignsMonthly = Build-MonthlyRows $parsed.campaignsWeekly -ByCampaign
    }
  }
}
finally {
  foreach ($file in $tempFiles) {
    if (Test-Path $file) { Remove-Item -LiteralPath $file -Force }
  }
}

$json = $clients | ConvertTo-Json -Depth 100 -Compress
$html = [IO.File]::ReadAllText((Resolve-Path $DashboardPath), [Text.Encoding]::UTF8)
$pattern = '(?s)const clients = .*?;\s+const clientLabels ='
$replacement = "const clients = $json;`r`n    const clientLabels ="
$match = [regex]::Match($html, $pattern)
if (-not $match.Success) {
  throw "Не найден блок const clients в titan-tokos-dashboard.html"
}
$updated = [regex]::Replace($html, $pattern, $replacement, 1)
[IO.File]::WriteAllText((Resolve-Path $DashboardPath), $updated, [Text.UTF8Encoding]::new($false))

$status = git status --short -- titan-tokos-dashboard.html
if (-not $status) {
  Write-Host "Дашборд уже актуален: изменений нет."
  exit 0
}

Write-Host "Данные обновлены в titan-tokos-dashboard.html"
if ($NoCommit) { exit 0 }

function Test-CanWriteGitDir() {
  $gitDir = Join-Path $PSScriptRoot "..\.git"
  if (-not (Test-Path -LiteralPath $gitDir)) { return $false }
  $probe = Join-Path $gitDir "__codex_write_probe.tmp"
  try {
    New-Item -ItemType File -LiteralPath $probe -Force -ErrorAction Stop | Out-Null
    Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
    return $true
  }
  catch {
    return $false
  }
}

if (-not (Test-CanWriteGitDir)) {
  Write-Warning "Нет прав записи в .git: пропускаю git add/commit/push."
  exit 0
}

git add titan-tokos-dashboard.html
git -c user.name="a1reklamaru" -c user.email="a1reklamaru@users.noreply.github.com" commit -m "Update dashboard data"

if (-not $NoPush) {
  try { git push }
  catch { Write-Warning "git push: $($_.Exception.Message)" }
}

