<#
Generate-TeamsNet-RTT-ByHost.ps1
- Bucketed & locale-safe (no worksheet formulas)
- Always releases Excel COM objects in finally, regardless of success/failure.
#>

param(
  [string]$InputDir   = (Join-Path $Env:LOCALAPPDATA 'TeamsNet'),
  [Parameter(Mandatory=$true)][string]$Output,
  [string]$TargetsFile,
  [int]$ThresholdMs = 100,
  [int]$BucketMinutes = 60,
  [switch]$Visible
)

# Excel consts
$xlDelimited=1; $xlYes=1; $xlLine=4; $xlLegendBottom=-4107
$xlSrcRange=1; $xlInsertDeleteCells=2; $xlCellTypeVisible=12
$xlUp=-4162; $xlCategory=1; $xlValue=2; $xlTimeScale=3
$msoLineDash=4

$ErrorActionPreference='Stop'

function Sanitize-SheetName([string]$name){
  if(-not $name){ return 'Host' }
  $n = $name -replace '[:\\/\?\*\[\]]','_'
  if($n.Length -gt 31){ $n = $n.Substring(0,31) }
  if($n -match '^\s*$'){ $n = 'Host' }
  return $n
}
function Get-HostColor([string]$HostName){
  $palette=@(
    @{r= 33; g=150; b=243}, @{r= 76; g=175; b= 80}, @{r=244; g= 67; b= 54},
    @{r=255; g=193; b=  7}, @{r=156; g= 39; b=176}, @{r=  0; g=188; b=212},
    @{r=121; g= 85; b= 72}, @{r= 63; g= 81; b=181}, @{r=205; g=220; b= 57},
    @{r=233; g= 30; b= 99}
  )
  $sum=0; $HostName.ToCharArray() | ForEach-Object { $sum += [int]$_ }
  $c=$palette[$sum % $palette.Count]
  return [int]($c.r + ($c.g -shl 8) + ($c.b -shl 16))
}
function Write-Column2D($ws, [string]$addr, [object[]]$arr){
  $n = if($arr){ [int]$arr.Count } else { 0 }
  if($n -le 0){ return }
  $data = New-Object 'object[,]' $n,1
  for($i=0;$i -lt $n;$i++){ $data[$i,0]=$arr[$i] }
  $ws.Range($addr).Resize($n,1).Value2 = $data
}
function New-RepeatedArray([object]$value, [int]$count){
  if($count -le 0){ return @() }
  $a = New-Object object[] $count
  for($i=0;$i -lt $count;$i++){ $a[$i] = $value }
  return $a
}
function Release-Com([object]$obj){
  if($null -ne $obj -and [System.Runtime.InteropServices.Marshal]::IsComObject($obj)){
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) } catch {}
  }
}

# ---------- inputs ----------
$csv = Join-Path $InputDir 'teams_net_quality.csv'
if(-not (Test-Path $csv)){ throw 'CSV not found: ' + $csv }

if(-not $TargetsFile){
  $TargetsFile = Join-Path $PSScriptRoot 'target.txt'
  if(-not (Test-Path $TargetsFile)){
    $TargetsFile = Join-Path $InputDir 'target.txt'
  }
}
if(-not (Test-Path $TargetsFile)){ throw 'Targets file not found: ' + $TargetsFile }

$targets = Get-Content -Raw -Encoding UTF8 $TargetsFile |
  ForEach-Object { $_ -split "`r?`n" } |
  Where-Object { $_ -and (-not $_.Trim().StartsWith('#')) } |
  ForEach-Object { $_.Trim() } | Select-Object -Unique
if(-not $targets -or $targets.Count -eq 0){ throw 'No valid hosts in target.txt' }

# ---------- main with guaranteed cleanup ----------
$excel=$null; $wb=$null; $wsAll=$null
try {
  # Excel
  try { $excel = New-Object -ComObject Excel.Application } catch { throw 'Cannot start Excel COM. Is Excel installed?' }
  $excel.Visible = [bool]$Visible
  $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Add()
  while($wb.Worksheets.Count -gt 1){ $wb.Worksheets.Item(1).Delete() }
  $wb.Worksheets.Item(1).Name = 'AllData'
  $wsAll = $wb.Worksheets('AllData')

  # import CSV
  try { foreach($qt in @($wsAll.QueryTables())){ $qt.Delete() } } catch {}
  try { foreach($lo in @($wsAll.ListObjects())){ $lo.Unlist() } } catch {}
  $wsAll.Cells.Clear()

  $qt = $wsAll.QueryTables.Add('TEXT;' + $csv, $wsAll.Range('A1'))
  $qt.TextFileParseType = $xlDelimited
  $qt.TextFileCommaDelimiter = $true
  $qt.TextFilePlatform = 65001
  $qt.TextFileTrailingMinusNumbers = $true
  $qt.AdjustColumnWidth = $true
  $qt.RefreshStyle = $xlInsertDeleteCells
  $null = $qt.Refresh()
  $rng = $qt.ResultRange
  if(-not $rng){ throw 'CSV import failed or empty: ' + $csv }
  $qt.Delete(); $qt=$null

  $loAll = $wsAll.ListObjects.Add($xlSrcRange, $rng, $null, $xlYes)
  $loAll.Name = 'tblAll'
  $null = $wsAll.Columns.AutoFit()

  function Try-GetColIndex($listObject,[string]$col){ try { return $listObject.ListColumns($col).Index } catch { return $null } }
  $colHost = Try-GetColIndex $loAll 'host'
  $colTime = Try-GetColIndex $loAll 'timestamp'
  $colRtt  = Try-GetColIndex $loAll 'icmp_avg_ms'
  if($null -eq $colHost -or $null -eq $colTime -or $null -eq $colRtt){
    throw 'Required columns missing: host, timestamp, icmp_avg_ms'
  }
  try { $loAll.ListColumns('timestamp').DataBodyRange.NumberFormatLocal = 'yyyy/mm/dd hh:mm' } catch {}
  try { $loAll.ListColumns('icmp_avg_ms').DataBodyRange.NumberFormatLocal = '0.0' } catch {}

  # per host
  $created=@()
  $frac = [double]$BucketMinutes / 1440.0
  $ciInv = [System.Globalization.CultureInfo]::InvariantCulture
  $ciCur = [System.Globalization.CultureInfo]::CurrentCulture

  foreach($h in $targets){
    $null = $loAll.Range.AutoFilter($colHost, $h)
    $timeVis = $loAll.ListColumns('timestamp').DataBodyRange.SpecialCells($xlCellTypeVisible)
    $rttVis  = $loAll.ListColumns('icmp_avg_ms').DataBodyRange.SpecialCells($xlCellTypeVisible)

    $sn = Sanitize-SheetName $h
    try {
      $ws = $wb.Worksheets.Item($sn); $ws.Cells.Clear()
      try { foreach($co in @($ws.ChartObjects())){ $co.Delete() } } catch {}
    } catch { $ws = $wb.Worksheets.Add(); $ws.Name=$sn }

    # headers
    $ws.Cells(1,1).Value2='timestamp'
    $ws.Cells(1,2).Value2='icmp_avg_ms'
    $ws.Cells(1,3).Value2='threshold_ms'

    # raw copy
    $null = $timeVis.Copy($ws.Range('A2'))
    $null = $rttVis.Copy($ws.Range('B2'))
    $lastRow = $ws.Cells($ws.Rows.Count,1).End($xlUp).Row
    if($lastRow -lt 2){ try { $null = $loAll.AutoFilter.ShowAllData() } catch {}; continue }

    $ws.Range("C2:C$lastRow").Value2 = $ThresholdMs
    try { $ws.Range("A2:A$lastRow").NumberFormatLocal='yyyy/mm/dd hh:mm' } catch {}
    try { $ws.Range("B2:B$lastRow").NumberFormatLocal='0.0' } catch {}

    # read raw and aggregate
    $rngAraw = $ws.Range("A2:A$lastRow").Value2
    $rngBraw = $ws.Range("B2:B$lastRow").Value2
    if(-not ($rngAraw -is [Array])){ $rngA = New-Object 'object[,]' 1,1; $rngA[0,0]=$rngAraw } else { $rngA=$rngAraw }
    if(-not ($rngBraw -is [Array])){ $rngB = New-Object 'object[,]' 1,1; $rngB[0,0]=$rngBraw } else { $rngB=$rngBraw }

    $rows = $rngA.GetLength(0)
    $times = New-Object double[] $rows
    $rtts  = New-Object double[] $rows
    for($i=0;$i -lt $rows;$i++){
      $t = $rngA[$i,0]
      if($t -is [double]){ $tNum = [double]$t }
      else {
        $tStr = [string]$t
        try { $dt = [datetime]::Parse($tStr, $ciCur) } catch { try { $dt = [datetime]::Parse($tStr, $ciInv) } catch { continue } }
        $tNum = $dt.ToOADate()
      }
      $v = $rngB[$i,0]
      if($v -is [double]){ $rNum = [double]$v }
      else {
        $tmp=0.0
        if([double]::TryParse([string]$v, [System.Globalization.NumberStyles]::Float, $ciInv, [ref]$tmp)){ $rNum = $tmp }
        elseif([double]::TryParse([string]$v, [System.Globalization.NumberStyles]::Float, $ciCur, [ref]$tmp)){ $rNum = $tmp }
        else { continue }
      }
      $times[$i]=$tNum; $rtts[$i]=$rNum
    }

    $agg = @{}  # bucket -> {sum,cnt}
    for($i=0;$i -lt $rows;$i++){
      $t=[double]$times[$i]; $r=[double]$rtts[$i]
      if([double]::IsNaN($t) -or [double]::IsNaN($r)){ continue }
      $b = [math]::Floor($t / $frac) * $frac
      $entry = $agg[$b]
      if(-not $entry){ $entry = @{sum=0.0; cnt=0}; $agg[$b]=$entry }
      $entry.sum = $entry.sum + $r
      $entry.cnt = $entry.cnt + 1
    }

    $useBuckets = ($agg.Keys.Count -ge 2)
    if($useBuckets){
      $keys = $agg.Keys | Sort-Object
      $xs = New-Object object[] $keys.Count
      $ys = New-Object object[] $keys.Count
      for($k=0;$k -lt $keys.Count;$k++){
        $b = [double]$keys[$k]
        $avg = if($agg[$b].cnt -gt 0){ $agg[$b].sum / $agg[$b].cnt } else { [double]::NaN }
        $xs[$k] = $b; $ys[$k] = $avg
      }
      $ws.Cells(1,6).Value2='bucket'
      $ws.Cells(1,7).Value2='avg_rtt'
      $ws.Cells(1,8).Value2='threshold_bucket'
      Write-Column2D $ws "F2" $xs
      Write-Column2D $ws "G2" $ys
      $thr = New-RepeatedArray -value $ThresholdMs -count $xs.Count
      Write-Column2D $ws "H2" $thr
      try { $ws.Range(("F2:F{0}" -f (1+$xs.Count))).NumberFormatLocal='yyyy/mm/dd hh:mm' } catch {}
      try { $ws.Range(("G2:G{0}" -f (1+$ys.Count))).NumberFormatLocal='0.0' } catch {}
      $null = $ws.Columns("A:H").AutoFit()

      try { foreach($co in @($ws.ChartObjects())){ $co.Delete() } } catch {}
      $ch = $ws.ChartObjects().Add(300,10,900,320)
      $c = $ch.Chart; $c.ChartType = $xlLine; $c.HasTitle=$true
      $c.ChartTitle.Text = 'RTT hourly avg (icmp_avg_ms) - ' + $h
      $c.Legend.Position = $xlLegendBottom
      try { $c.SeriesCollection().Delete() } catch {}
      $rgb = Get-HostColor $h
      $endRow = 1 + $xs.Count
      $s1 = $c.SeriesCollection().NewSeries(); $s1.Name=$h
      $s1.XValues = $ws.Range(("F2:F{0}" -f $endRow)); $s1.Values = $ws.Range(("G2:G{0}" -f $endRow))
      try { $s1.Format.Line.ForeColor.RGB = $rgb; $s1.Format.Line.Weight = 2 } catch {}
      $s2 = $c.SeriesCollection().NewSeries(); $s2.Name='threshold ' + $ThresholdMs + ' ms'
      $s2.XValues = $ws.Range(("F2:F{0}" -f $endRow)); $s2.Values = $ws.Range(("H2:H{0}" -f $endRow))
      try { $s2.Format.Line.ForeColor.RGB = 255; $s2.Format.Line.Weight = 1.5; $s2.Format.Line.DashStyle = $msoLineDash } catch {}
      try {
        $v = $c.Axes($xlValue); $v.MinimumScale=0; $v.MaximumScale=300; $v.MajorUnit=50; $v.TickLabels.NumberFormat='0.0'
        $x = $c.Axes($xlCategory); $x.CategoryType=$xlTimeScale; $x.MajorUnit=1/24; $x.TickLabels.NumberFormat='mm/dd hh:mm'
      } catch {}
    } else {
      try { foreach($co in @($ws.ChartObjects())){ $co.Delete() } } catch {}
      $ch = $ws.ChartObjects().Add(300,10,900,320)
      $c = $ch.Chart; $c.ChartType = $xlLine; $c.HasTitle=$true
      $c.ChartTitle.Text = 'RTT (raw) - ' + $h
      $c.Legend.Position = $xlLegendBottom
      try { $c.SeriesCollection().Delete() } catch {}
      $rgb = Get-HostColor $h
      $s1 = $c.SeriesCollection().NewSeries(); $s1.Name=$h
      $s1.XValues = $ws.Range("A2:A$lastRow"); $s1.Values = $ws.Range("B2:B$lastRow")
      try { $s1.Format.Line.ForeColor.RGB = $rgb; $s1.Format.Line.Weight = 2 } catch {}
      $s2 = $c.SeriesCollection().NewSeries(); $s2.Name='threshold ' + $ThresholdMs + ' ms'
      $s2.XValues = $ws.Range("A2:A$lastRow"); $s2.Values = $ws.Range("C2:C$lastRow")
      try { $s2.Format.Line.ForeColor.RGB = 255; $s2.Format.Line.Weight = 1.5; $s2.Format.Line.DashStyle = $msoLineDash } catch {}
      try {
        $v = $c.Axes($xlValue); $v.MinimumScale=0; $v.MaximumScale=300; $v.MajorUnit=50; $v.TickLabels.NumberFormat='0.0'
        $x = $c.Axes($xlCategory); $x.CategoryType=$xlTimeScale; $x.MajorUnit=1/24; $x.TickLabels.NumberFormat='mm/dd hh:mm'
      } catch {}
    }

    $created += $sn
    try { $null = $loAll.AutoFilter.ShowAllData() } catch {}
    $ws=$null
  }

  if(-not $created -or $created.Count -eq 0){ throw 'No data matched hosts in target.txt' }

  try {
    $wsIdx = $wb.Worksheets.Add(); $wsIdx.Name='INDEX'
    $wsIdx.Cells(1,1).Value2='Hosts'
    $r=2; foreach($sn in $created){ $wsIdx.Hyperlinks.Add($wsIdx.Cells($r,1),'',"'$sn'!A1",'',$sn) | Out-Null; $r++ }
    $wsIdx=$null
  } catch {}

  try { $wb.Worksheets('AllData').Move($wb.Worksheets.Item($wb.Worksheets.Count)) } catch {}

  $wb.SaveAs($Output)
}
catch {
  Write-Error $_
  throw
}
finally {
  # close workbook first
  if($wb){
    try { $wb.Close($false) } catch {}
    Release-Com $wb
    $wb = $null
  }
  # quit Excel last
  if($excel){
    try { $excel.Quit() } catch {}
    Release-Com $excel
    $excel = $null
  }
  # extra GC passes help tear down COM RCWs
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host ('Output: ' + $Output)