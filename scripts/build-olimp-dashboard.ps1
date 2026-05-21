param(
  [string]$SpreadsheetId = "1GIvNz_RrA0tMMU9hJ33aotnLgIcv9m4CM8xn_rmZ_oE",
  [string]$Gid = "1880532691",
  [string]$OutputPath = ".\olimpiyskaya-derevnya-dashboard.html",
  [string]$TemplatePath = ".\titan-tokos-dashboard.html"
)

$ErrorActionPreference = "Stop"

$MonthNames = @{
  "января" = 1; "январь" = 1;
  "февраля" = 2; "февраль" = 2;
  "марта" = 3; "март" = 3;
  "апреля" = 4; "апрель" = 4;
  "мая" = 5; "май" = 5;
  "июня" = 6; "июнь" = 6;
  "июля" = 7; "июль" = 7;
  "августа" = 8; "август" = 8;
  "сентября" = 9; "сентябрь" = 9;
  "октября" = 10; "октябрь" = 10;
  "ноября" = 11; "ноябрь" = 11;
  "декабря" = 12; "декабрь" = 12
}

function Convert-ToNumber($Value) {
  if ($null -eq $Value) { return 0 }
  $text = ([string]$Value).Trim()
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

function Get-MonthNumber($Text) {
  $lower = ([string]$Text).ToLowerInvariant()
  foreach ($name in $MonthNames.Keys) {
    if ($lower -match [regex]::Escape($name)) { return $MonthNames[$name] }
  }
  return $null
}

function Test-PeriodHeader($Row) {
  $first = Get-FirstFilledCell $Row
  if ($first -notmatch "^\s*\d{1,2}" -or $null -eq (Get-MonthNumber $first)) { return $false }
  return ($Row.Count -gt 1 -and ([string]$Row[1]).Trim() -eq "Показы")
}

function Get-PeriodBounds($Period, $SectionYear) {
  $clean = ([string]$Period) -replace "\(.*?\)", " "
  $clean = $clean -replace "[-–—]", " - "
  $clean = $clean -replace "\s+", " "
  $lower = $clean.Trim().ToLowerInvariant()
  $monthMatches = [regex]::Matches($lower, "(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)")
  if ($monthMatches.Count -eq 0) { throw "Не удалось распознать месяц периода: $Period" }

  $numbers = [regex]::Matches($lower, "\d{1,2}") | ForEach-Object { [int]$_.Value }
  if ($numbers.Count -eq 0) { throw "Не удалось распознать день периода: $Period" }

  $startDay = $numbers[0]
  $endDay = if ($numbers.Count -ge 2) { $numbers[$numbers.Count - 1] } else { $startDay + 6 }
  $startMonth = $MonthNames[$monthMatches[0].Value]
  $endMonth = if ($monthMatches.Count -ge 2) { $MonthNames[$monthMatches[$monthMatches.Count - 1].Value] } else { $startMonth }

  $startYear = $SectionYear
  $endYear = $SectionYear
  if ($startMonth -eq 12 -and $endMonth -eq 1) {
    $startYear = $SectionYear - 1
  }
  elseif ($startMonth -gt $endMonth) {
    $endYear = $SectionYear + 1
  }

  $start = Get-Date -Year $startYear -Month $startMonth -Day $startDay -Hour 0 -Minute 0 -Second 0
  $end = Get-Date -Year $endYear -Month $endMonth -Day $endDay -Hour 0 -Minute 0 -Second 0
  return @{ Start = $start; End = $end }
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
    if ($Template.Contains("campaign")) { $Map[$Key].campaign = $Template.campaign }
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
    $start = [datetime]::ParseExact($item.date, "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture)
    $days = if ($item.Contains("days")) { [int]$item.days } else { 7 }
    for ($i = 0; $i -lt $days; $i++) {
      $day = $start.AddDays($i)
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
      Add-ToBucket $map $key $template (1 / $days)
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

function Build-YearlyRows($Rows, [switch]$ByCampaign) {
  $map = @{}
  foreach ($item in $Rows) {
    $yearStart = Get-Date -Year $item.year -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
    $key = if ($ByCampaign) { "$($item.campaign)|$($item.year)" } else { [string]$item.year }
    $template = [ordered]@{
      period = "$($item.year) год"
      date = $yearStart.ToString("yyyy-MM-dd")
      year = [int]$item.year
      impressions = $item.impressions
      clicks = $item.clicks
      cost = $item.cost
      conversions = $item.conversions
    }
    if ($ByCampaign) { $template.campaign = $item.campaign }
    Add-ToBucket $map $key $template 1
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

function Parse-OlimpSheet($Path) {
  $rows = Read-CsvRows $Path
  $sectionYear = $null
  $currentPeriod = $null
  $currentBounds = $null
  $campaignRows = @()
  $weeklyByDate = [ordered]@{}
  $campaignsByDate = [ordered]@{}

  foreach ($row in $rows) {
    $first = Get-FirstFilledCell $row
    if (-not $first) { continue }
    if ($first -like "Остаток*") { continue }

    if ($first -match "(20\d{2})(?:\D|$)") {
      $sectionYear = [int]$Matches[1]
      if (-not (Test-PeriodHeader $row)) { continue }
    }

    if (Test-PeriodHeader $row) {
      if ($null -eq $sectionYear) { throw "Не найден год перед периодом: $first" }
      $currentPeriod = $first
      $currentBounds = Get-PeriodBounds $first $sectionYear
      $campaignRows = @()
      continue
    }

    if ($first -eq "Всего" -and $null -ne $currentBounds) {
      $days = [int](($currentBounds.End - $currentBounds.Start).TotalDays + 1)
      $dateKey = $currentBounds.Start.ToString("yyyy-MM-dd")
      $weeklyByDate[$dateKey] = New-MetricRow $currentPeriod $currentBounds.Start (Convert-ToInt $row[1]) (Convert-ToInt $row[2]) (Convert-ToNumber $row[4]) (Convert-ToInt $row[6])
      $weeklyByDate[$dateKey].days = $days
      foreach ($campaignRow in $campaignRows) { $campaignRow.days = $days }
      $campaignsByDate[$dateKey] = @($campaignRows)
      continue
    }

    if ($null -ne $currentBounds -and $row.Count -gt 6 -and (Convert-ToNumber $row[1]) -gt 0 -and $first -ne "ИТОГО") {
      $campaignRow = [ordered]@{
        campaign = ($first -replace "\s+", " ").Trim()
        date = $currentBounds.Start.ToString("yyyy-MM-dd")
        year = [int]$currentBounds.Start.Year
        impressions = Convert-ToInt $row[1]
        clicks = Convert-ToInt $row[2]
        ctr = 0
        cost = [Math]::Round((Convert-ToNumber $row[4]), 2)
        cpc = 0
        conversions = Convert-ToInt $row[6]
        cpa = 0
      }
      $campaignRow.ctr = if ($campaignRow.impressions) { [Math]::Round($campaignRow.clicks / $campaignRow.impressions, 6) } else { 0 }
      $campaignRow.cpc = if ($campaignRow.clicks) { [Math]::Round($campaignRow.cost / $campaignRow.clicks, 2) } else { 0 }
      $campaignRow.cpa = if ($campaignRow.conversions) { [Math]::Round($campaignRow.cost / $campaignRow.conversions, 2) } else { 0 }
      $campaignRows += $campaignRow
    }
  }

  $weekly = @()
  foreach ($dateKey in ($weeklyByDate.Keys | Sort-Object)) { $weekly += $weeklyByDate[$dateKey] }

  $campaignsWeekly = @()
  foreach ($dateKey in ($campaignsByDate.Keys | Sort-Object)) { $campaignsWeekly += $campaignsByDate[$dateKey] }

  return @{ weekly = $weekly; campaignsWeekly = $campaignsWeekly }
}

function Set-RegexOnce($Text, $Pattern, $Replacement, $ErrorMessage) {
  $match = [regex]::Match($Text, $Pattern)
  if (-not $match.Success) { throw $ErrorMessage }
  return [regex]::Replace($Text, $Pattern, $Replacement, 1)
}

$tempFile = Join-Path $env:TEMP "dashboard-olimp-$([guid]::NewGuid()).csv"
try {
  $url = "https://docs.google.com/spreadsheets/d/$SpreadsheetId/export?format=csv&gid=$Gid"
  Invoke-WebRequest -Uri $url -OutFile $tempFile -UseBasicParsing
  $parsed = Parse-OlimpSheet $tempFile
}
finally {
  if (Test-Path -LiteralPath $tempFile) { Remove-Item -LiteralPath $tempFile -Force }
}

$monthly = Build-MonthlyRows $parsed.weekly
$campaignsMonthly = Build-MonthlyRows $parsed.campaignsWeekly -ByCampaign
$clients = [ordered]@{
  olimp = [ordered]@{
    weekly = $parsed.weekly
    monthly = $monthly
    yearly = Build-YearlyRows $monthly
    campaignsWeekly = $parsed.campaignsWeekly
    campaignsMonthly = $campaignsMonthly
    campaignsYearly = Build-YearlyRows $campaignsMonthly -ByCampaign
  }
}
$labels = [ordered]@{ olimp = 'База отдыха "Олимпийская деревня"' }

$resolvedTemplate = Resolve-Path $TemplatePath
$html = [IO.File]::ReadAllText($resolvedTemplate, [Text.Encoding]::UTF8)
$clientsJson = $clients | ConvertTo-Json -Depth 100 -Compress
$labelsJson = $labels | ConvertTo-Json -Depth 20 -Compress

$html = Set-RegexOnce $html '(?s)const clients = .*?;\s+const clientLabels =' "const clients = $clientsJson;`r`n    const clientLabels =" "Не найден блок const clients."
$html = Set-RegexOnce $html '(?s)const clientLabels = .*?;\s+const colors =' "const clientLabels = $labelsJson;`r`n    const colors =" "Не найден блок const clientLabels."
$html = Set-RegexOnce $html 'const state = \{ client: ".*?",' 'const state = { client: "olimp",' "Не найден блок const state."
$html = Set-RegexOnce $html '(?s)\s*<div class="segmented" role="group" aria-label="Клиент">\s*.*?\s*</div>' "" "Не найден блок кнопок клиентов."
$html = Set-RegexOnce $html '<title>.*?</title>' '<title>Дашборд База отдыха "Олимпийская деревня"</title>' "Не найден title."
$html = Set-RegexOnce $html '<h1 id="clientTitle">.*?</h1>' '<h1 id="clientTitle">База отдыха "Олимпийская деревня"</h1>' "Не найден clientTitle."
$html = Set-RegexOnce $html '<p>Динамика рекламных показателей по неделям и месяцам в Яндекс Директе</p>' '<p>Динамика рекламных показателей по неделям, месяцам и годам в Яндекс Директе</p>' "Не найден subtitle."
$html = Set-RegexOnce $html '<button type="button" data-mode="monthly">Месяцы</button>' "<button type=""button"" data-mode=""monthly"">Месяцы</button>`r`n          <button type=""button"" data-mode=""yearly"">Годы</button>" "Не найден переключатель месяцев."
$html = $html.Replace('function campaignData() { return clients[state.client][state.mode === "monthly" ? "campaignsMonthly" : "campaignsWeekly"]; }', 'function campaignData() { return clients[state.client][state.mode === "yearly" ? "campaignsYearly" : state.mode === "monthly" ? "campaignsMonthly" : "campaignsWeekly"]; }')
$html = $html.Replace('const periodName = state.mode === "monthly" ? "месяцы" : "недели";', 'const periodName = state.mode === "yearly" ? "годы" : state.mode === "monthly" ? "месяцы" : "недели";')
$html = $html.Replace('document.getElementById("periodFilterLabel").textContent = state.mode === "monthly" ? "Месяц" : "Неделя";', 'document.getElementById("periodFilterLabel").textContent = state.mode === "yearly" ? "Год" : state.mode === "monthly" ? "Месяц" : "Неделя";')
$html = $html.Replace('document.getElementById("tableTitle").textContent = state.mode === "monthly" ? "Данные по месяцам" : "Данные по неделям";', 'document.getElementById("tableTitle").textContent = state.mode === "yearly" ? "Данные по годам" : state.mode === "monthly" ? "Данные по месяцам" : "Данные по неделям";')
$html = $html.Replace('document.getElementById("periodColumnTitle").textContent = state.mode === "monthly" ? "Месяц" : "Неделя";', 'document.getElementById("periodColumnTitle").textContent = state.mode === "yearly" ? "Год" : state.mode === "monthly" ? "Месяц" : "Неделя";')
$html = $html.Replace('function periodUnitShort() { return state.mode === "monthly" ? "мес." : "нед."; }', 'function periodUnitShort() { return state.mode === "yearly" ? "г." : state.mode === "monthly" ? "мес." : "нед."; }')
$html = $html.Replace('    .select-wrap {', @'
    .quick-range {
      box-shadow: inset 0 0 0 1px rgba(91, 111, 138, 0.12);
    }

    .quick-range button {
      min-width: 84px;
    }

    .select-wrap {
'@)
$html = [regex]::Replace($html, '(?s)(<label class="select-wrap">\s*<span id="periodFilterLabel">.*?</span>\s*<select id="weekFilter">.*?</select>\s*</label>)\s*</div>', "`$1`r`n        <div class=""segmented quick-range"" role=""group"" aria-label=""Быстрый период"">`r`n          <button type=""button"" data-quick-range=""last-week"">Неделя</button>`r`n          <button type=""button"" data-quick-range=""last-2-weeks"">2 недели</button>`r`n          <button type=""button"" data-quick-range=""last-month"">Месяц</button>`r`n        </div>`r`n      </div>", 1)
$html = $html.Replace('const state = { client: "olimp", mode: "weekly", range: "all", week: "all" };', 'const state = { client: "olimp", mode: "weekly", range: "all", week: "all", quickRange: "all" };')
$html = $html.Replace('const state = { client: "olimp", mode: "weekly", range: "all", week: "all", quickRange: "all" };', 'const state = { client: "olimp", mode: "weekly", range: "all", week: "all", quickRange: "last-week" };')
$html = $html.Replace('state.week = "all";
        monthDetailsOpen = true;', 'state.week = "all";
        state.quickRange = "all";
        monthDetailsOpen = true;
        updateQuickRangeButtons();')
$html = $html.Replace('state.week = event.target.value;
      monthDetailsOpen = true;', 'state.week = event.target.value;
      state.quickRange = "all";
      monthDetailsOpen = true;
      updateQuickRangeButtons();')
$html = [regex]::Replace($html, 'state\.week = "all";\s*monthDetailsOpen = true;', "state.week = ""all"";`r`n        state.quickRange = ""all"";`r`n        monthDetailsOpen = true;`r`n        updateQuickRangeButtons();")
$html = [regex]::Replace($html, 'state\.week = event\.target\.value;\s*monthDetailsOpen = true;', "state.week = event.target.value;`r`n      state.quickRange = ""all"";`r`n      monthDetailsOpen = true;`r`n      updateQuickRangeButtons();")
$html = $html.Replace('    window.addEventListener("resize", debounce(() => {', @'
    document.querySelectorAll("[data-quick-range]").forEach((button) => {
      button.addEventListener("click", () => {
        applyQuickRange(button.dataset.quickRange);
      });
    });

    window.addEventListener("resize", debounce(() => {
'@)
$html = $html.Replace('function currentCampaignData() {
      return clients[state.client][state.mode === "monthly" ? "campaignsMonthly" : "campaignsWeekly"];
    }', 'function currentCampaignData() {
      return clients[state.client][state.mode === "yearly" ? "campaignsYearly" : state.mode === "monthly" ? "campaignsMonthly" : "campaignsWeekly"];
    }')
$html = [regex]::Replace($html, 'function currentCampaignData\(\) \{\s*return clients\[state\.client\]\[state\.mode === "monthly" \? "campaignsMonthly" : "campaignsWeekly"\];\s*\}', 'function currentCampaignData() {
      return clients[state.client][state.mode === "yearly" ? "campaignsYearly" : state.mode === "monthly" ? "campaignsMonthly" : "campaignsWeekly"];
    }')
$html = $html.Replace('    function updateQuickRangeButtons() {', @'
    function initializeDefaultPeriod() {
      setMode("weekly");
      state.range = "all";
      state.quickRange = "last-week";
      const latestWeek = latestRows(currentData(), 1)[0];
      state.week = latestWeek ? latestWeek.date : "all";
      monthDetailsOpen = true;
      updateQuickRangeButtons();
    }

    function updateQuickRangeButtons() {
'@)
$html = $html.Replace('    function renderYearButtons() {', @'
    function setMode(mode) {
      state.mode = mode;
      document.querySelectorAll("[data-mode]").forEach((item) => {
        item.classList.toggle("active", item.dataset.mode === mode);
      });
    }

    function latestRows(rows, count) {
      return [...rows].sort((a, b) => a.date.localeCompare(b.date)).slice(-count);
    }

    function applyQuickRange(range) {
      state.quickRange = range;
      state.range = "all";
      state.week = "all";
      monthDetailsOpen = true;
      if (range === "last-month") {
        setMode("monthly");
        renderYearButtons();
        const latestMonth = latestRows(currentData(), 1)[0];
        state.week = latestMonth ? latestMonth.date : "all";
      }
      else {
        setMode("weekly");
        renderYearButtons();
        const latestWeek = latestRows(currentData(), 1)[0];
        state.week = range === "last-week" && latestWeek ? latestWeek.date : "all";
      }
      updateQuickRangeButtons();
      updateWeekFilter();
      render();
    }

    function updateQuickRangeButtons() {
      document.querySelectorAll("[data-quick-range]").forEach((button) => {
        button.classList.toggle("active", state.quickRange === button.dataset.quickRange);
      });
    }

    function renderYearButtons() {
'@)
$html = $html.Replace('    function updateQuickRangeButtons() {', @'
    function initializeDefaultPeriod() {
      setMode("weekly");
      state.range = "all";
      state.quickRange = "last-week";
      const latestWeek = latestRows(currentData(), 1)[0];
      state.week = latestWeek ? latestWeek.date : "all";
      monthDetailsOpen = true;
      updateQuickRangeButtons();
    }

    function updateQuickRangeButtons() {
'@)
$html = $html.Replace('state.week = "all";
          monthDetailsOpen = true;', 'state.week = "all";
          state.quickRange = "all";
          monthDetailsOpen = true;
          updateQuickRangeButtons();')
$html = $html.Replace('      return currentData().filter((item) => {
        const yearMatches = state.range === "all" || String(item.year) === state.range;
        const weekMatches = state.week === "all" || item.date === state.week;
        return yearMatches && weekMatches;
      });', '      const baseRows = currentData().filter((item) => {
        const yearMatches = state.range === "all" || String(item.year) === state.range;
        const weekMatches = state.week === "all" || item.date === state.week;
        return yearMatches && weekMatches;
      });
      if (state.quickRange === "last-2-weeks") return latestRows(baseRows, 2);
      return baseRows;')
$html = [regex]::Replace($html, 'return currentData\(\)\.filter\(\(item\) => \{\s*const yearMatches = state\.range === "all" \|\| String\(item\.year\) === state\.range;\s*const weekMatches = state\.week === "all" \|\| item\.date === state\.week;\s*return yearMatches && weekMatches;\s*\}\);', 'const baseRows = currentData().filter((item) => {
        const yearMatches = state.range === "all" || String(item.year) === state.range;
        const weekMatches = state.week === "all" || item.date === state.week;
        return yearMatches && weekMatches;
      });
      if (state.quickRange === "last-2-weeks") return latestRows(baseRows, 2);
      return baseRows;')
$html = $html.Replace('const periodMatches = state.week === "all" || selectedDates.has(item.date);', 'const periodMatches = state.week === "all" && state.quickRange === "all" ? true : selectedDates.has(item.date);')
$html = $html.Replace('state.week = range === "last-week" && latestWeek ? latestWeek.date : "all";', 'state.week = range === "last-2-weeks" ? "quick:last-2-weeks" : latestWeek ? latestWeek.date : "all";')
$html = $html.Replace('document.getElementById("weekFilter").innerHTML = [
        `<option value="all">Все ${periodName}</option>`,
        ...periods.map((item) => `<option value="${item.date}">${item.period}</option>`)
      ].join("");', 'const quickOptions = state.quickRange === "last-2-weeks" ? [''<option value="quick:last-2-weeks">Последние 2 недели</option>''] : [];
      document.getElementById("weekFilter").innerHTML = [
        `<option value="all">Все ${periodName}</option>`,
        ...quickOptions,
        ...periods.map((item) => `<option value="${item.date}">${item.period}</option>`)
      ].join("");')
$html = [regex]::Replace($html, 'document\.getElementById\("periodFilterLabel"\)\.textContent = state\.mode === "yearly" \? "Год" : state\.mode === "monthly" \? "Месяц" : "Неделя";\s*document\.getElementById\("weekFilter"\)\.innerHTML = \[\s*`<option value="all">Все \$\{periodName\}</option>`,\s*\.\.\.periods\.map\(\(item\) => `<option value="\$\{item\.date\}">\$\{item\.period\}</option>`\)\s*\]\.join\(""\);', 'document.getElementById("periodFilterLabel").textContent = state.mode === "yearly" ? "Год" : state.mode === "monthly" ? "Месяц" : "Неделя";
      const quickOptions = state.quickRange === "last-2-weeks" ? [''<option value="quick:last-2-weeks">Последние 2 недели</option>''] : [];
      document.getElementById("weekFilter").innerHTML = [
        `<option value="all">Все ${periodName}</option>`,
        ...quickOptions,
        ...periods.map((item) => `<option value="${item.date}">${item.period}</option>`)
      ].join("");')
$html = [regex]::Replace($html, 'function getFilteredData\(\) \{\s*const baseRows = currentData\(\)\.filter\(\(item\) => \{\s*const yearMatches = state\.range === "all" \|\| String\(item\.year\) === state\.range;\s*const weekMatches = state\.week === "all" \|\| item\.date === state\.week;\s*return yearMatches && weekMatches;\s*\}\);\s*if \(state\.quickRange === "last-2-weeks"\) return latestRows\(baseRows, 2\);\s*return baseRows;\s*\}', 'function getFilteredData() {
      if (state.week === "quick:last-2-weeks" || state.quickRange === "last-2-weeks") {
        const periodRows = currentData().filter((item) => state.range === "all" || String(item.year) === state.range);
        return latestRows(periodRows, 2);
      }
      if (state.quickRange === "last-week" && state.week === "all") {
        const periodRows = currentData().filter((item) => state.range === "all" || String(item.year) === state.range);
        return latestRows(periodRows, 1);
      }
      const baseRows = currentData().filter((item) => {
        const yearMatches = state.range === "all" || String(item.year) === state.range;
        const weekMatches = state.week === "all" || item.date === state.week;
        return yearMatches && weekMatches;
      });
      return baseRows;
    }')
$html = $html.Replace('if (!rows.length) return "период";', 'if (!rows.length) return "период";
      if (state.quickRange === "last-week") return `последняя неделя: ${rows[0].period}`;
      if (state.quickRange === "last-2-weeks") return "последние 2 недели";
      if (state.quickRange === "last-month") return `последний месяц: ${rows[0].period}`;')
$html = $html.Replace("updateQuickRangeButtons();`r`n        initializeDefaultPeriod();`r`n        renderYearButtons();", "updateQuickRangeButtons();`r`n        renderYearButtons();")
$html = [regex]::Replace($html, '\s*(?:initializeDefaultPeriod\(\);\s*)?renderYearButtons\(\);\s*updateWeekFilter\(\);\s*render\(\);\s*</script>', "`r`n    initializeDefaultPeriod();`r`n    renderYearButtons();`r`n    updateWeekFilter();`r`n    render();`r`n  </script>", 1)

$resolvedOutput = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
[IO.File]::WriteAllText($resolvedOutput, $html, [Text.UTF8Encoding]::new($false))

Write-Host "Дашборд создан: $resolvedOutput"
Write-Host "Недель: $($parsed.weekly.Count)"
Write-Host "Месяцев: $($monthly.Count)"
Write-Host "Годов: $($clients.olimp.yearly.Count)"




