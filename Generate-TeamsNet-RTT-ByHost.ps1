<#
Generate-TeamsNet-RTT-ByHost.ps1
- CSV(teams_net_quality.csv)のhost列から、target.txt に書いた IP or ホスト名に一致する行だけ抽出してグラフ化
- 表記ゆれ許容: URL/ポート/括弧/末尾ドット除去、必要ならDNSで名前→IPv4も同一視
- 1ホスト=1シート。A/B/C 列（timestamp/icmp_avg_ms/threshold）のみをグラフ化
- XY(Scatter with lines)で X=時刻(OADate, 昇順)、Y=RTT。閾値線(既定100ms, 赤破線)
- Y: 0..300(50刻み), X: 1時間刻み
- 例外は詳細出力、成功/失敗に関わらず Excel COM を確実に解放

Usage:
  powershell -NoProfile -ExecutionPolicy Bypass `
    -File .\Generate-TeamsNet-RTT-ByHost.ps1 `
    -Output ".\Output\TeamsNet-RTT-ByHost.xlsx" `
    -TargetsFile ".\target.txt" `
    -BucketMinutes 60
#>

param(
  [string]$InputDir   = (Join-Path $Env:LOCALAPPDATA 'TeamsNet'),
  [Parameter(Mandatory=$true)][string]$Output,
  [string]$TargetsFile,
  [int]$ThresholdMs = 100,
  [int]$BucketMinutes = 60,
  [switch]$Visible
)

# ---- error reporting ----
$Error.Clear()
$ErrorActionPreference = 'Stop'
$PSDefaultParameterValues['*:ErrorAction'] = 'Stop'
$global:ErrorView = 'NormalView'

function Format-ErrorRecord {
  param([System.Management.Automation.ErrorRecord]$Err)
  $ex = $Err.Exception
  $lines = New-Object System.Collections.Generic.List[string]
  $lines.Add('ERROR -----')
  $lines.Add(('FQID   : {0}' -f $Err.FullyQualifiedErrorId))
  if($Err.CategoryInfo){ $lines.Add(('Category: {0}' -f $Err.CategoryInfo.ToString())) }
  if($Err.TargetObject){ $lines.Add(('Target  : {0}' -f $Err.TargetObject)) }
  if($Err.InvocationInfo){
    $ii=$Err.InvocationInfo
    $lines.Add(('Script  : {0}' -f $ii.ScriptName))
    $lines.Add(('Line    : {0}  Char : {1}' -f $ii.ScriptLineNumber,$ii.OffsetInLine))
    if($ii.Line){ $lines.Add(('Code    : {0}' -f $ii.Line.Trim())) }
    if($ii.PositionMessage){ $lines.Add($ii.PositionMessage) }
  }
  if($ex){
    $lines.Add(('Type    : {0}' -f $ex.GetType().FullName))
    try { if($ex.HResult -ne $null){ $lines.Add(('HResult : 0x{0:X8}' -f $ex.HResult)) } } catch {}
    if($ex.Message){ $lines.Add(('Message : {0}' -f $ex.Message)) }
    if($ex.StackTrace){ $lines.Add('StackTrace:'); $lines.Add($ex.StackTrace) }
    if($Err.ScriptStackTrace){ $lines.Add('ScriptStackTrace:'); $lines.Add($Err.ScriptStackTrace) }
    $ix=$ex.InnerException; $n=1
    while($ix){
      $lines.Add(('Inner[{0}] Type   : {1}' -f $n,$ix.GetType().FullName))
      $lines.Add(('Inner[{0}] Message: {1}' -f $n,$ix.Message))
      if($ix.StackTrace){ $lines.Add($ix.StackTrace) }
      $ix=$ix.InnerException; $n++
    }
  }
  return ($lines -join [Environment]::NewLine)
}

# ---- helpers ----
function Sanitize-SheetName([string]$name){
  if(-not $name){ return 'Host' }
  $n = $name -replace '[:\\/\?\*\[\]]','_'
  if($n.Length -gt 31){ $n=$n.Substring(0,31) }
  if($n -match '^\s*$'){ $n='Host' }
  return $n
}
function Get-HostColor([string]$HostName){
  $palette=@(
    @{r=33; g=150; b=243}, @{r=76; g=175; b=80}, @{r=244; g=67; b=54},
    @{r=255; g=193; b=7},  @{r=156; g=39;  b=176}, @{r=0;  g=188; b=212},
    @{r=121; g=85;  b=72}, @{r=63; g=81;  b=181}, @{r=205; g=220; b=57},
    @{r=233; g=30;  b=99}
  )
  $sum=0; $HostName.ToCharArray() | ForEach-Object { $sum+=[int]$_ }
  $c=$palette[$sum % $palette.Count]
  [int]$rgb = ([int]$c.r) -bor (([int]$c.g) -shl 8) -bor (([int]$c.b) -shl 16)
  return $rgb
}
function Write-Column2D($ws,[string]$addr,[object[]]$arr){
  [int]$n = if($arr){ [int]$arr.Count } else { 0 }
  if($n -le 0){ return }
  $data = New-Object 'object[,]' ([int]$n),([int]1)
  for($i=0;$i -lt $n;$i++){ $data[$i,0]=$arr[$i] }
  $ws.Range($addr).Resize([int]$n,1).Value2=$data
}
function New-RepeatedArray([object]$value,[int]$count){
  if($count -le 0){ return @() }
  $count=[int]$count
  $a=New-Object object[] $count
  for($i=0;$i -lt $count;$i++){ $a[$i]=$value }
  return $a
}
function Release-Com([object]$obj){
  if($null -ne $obj -and [System.Runtime.InteropServices.Marshal]::IsComObject($obj)){
    try{ [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) }catch{}
  }
}

# ★ host表記を正規化（URL→host、:port除去、(ip)抽出、末尾.除去、[]除去、Lower/Trim）
function Normalize-Host([string]$s){
  if(-not $s){ return '' }
  $t = $s.Trim().Trim('"',"'").ToLowerInvariant()
  if($t -match '\(([0-9]{1,3}(?:\.[0-9]{1,3}){3})\)'){ return $Matches[1] }
  try{
    $uri = $null
    if([System.Uri]::TryCreate($t, [System.UriKind]::Absolute, [ref]$uri) -and $uri.Host){ $t = $uri.Host }
  }catch{}
  $t = $t.TrimEnd('.').Trim('[',']')
  if($t -match '^(.+?):\d+$'){ $t = $Matches[1] }
  return $t
}
# ★ 名前→IPv4の別名も追加（失敗は無視）
function Expand-Aliases([string]$hostNorm){
  $set = New-Object System.Collections.Generic.HashSet[string]
  if([string]::IsNullOrWhiteSpace($hostNorm)){ return $set }
  [void]$set.Add($hostNorm)
  try{
    if($hostNorm -notmatch '^\d{1,3}(\.\d{1,3}){3}$'){
      $ips = [System.Net.Dns]::GetHostAddresses($hostNorm) | Where-Object { $_.AddressFamily -eq 'InterNetwork' }
      foreach($ip in $ips){ [void]$set.Add($ip.ToString().ToLowerInvariant()) }
    }
  }catch{}
  return $set
}

# ---- Excel consts ----
[int]$xlDelimited=1; [int]$xlYes=1; [int]$xlLegendBottom=-4107
[int]$xlSrcRange=1; [int]$xlInsertDeleteCells=2; [int]$xlCellTypeVisible=12
[int]$xlUp=-4162; [int]$xlCategory=1; [int]$xlValue=2
[int]$xlXYScatterLinesNoMarkers = 75
[int]$msoLineDash=4

# ---- inputs ----
$csv = Join-Path $InputDir 'teams_net_quality.csv'
if(-not (Test-Path $csv)){ throw 'CSV not found: ' + $csv }

if(-not $TargetsFile){
  $TargetsFile = Join-Path $PSScriptRoot 'target.txt'
  if(-not (Test-Path $TargetsFile)){ $TargetsFile = Join-Path $InputDir 'target.txt' }
}
if(-not (Test-Path $TargetsFile)){ throw 'Targets file not found: ' + $TargetsFile }

# targets 読み込み（コメント/空行スキップ）
$targets = Get-Content -Encoding UTF8 $TargetsFile |
  ForEach-Object { $_.Trim() } |
  Where-Object { $_ -and (-not $_.StartsWith('#')) } |
  Select-Object -Unique
if(-not $targets -or $targets.Count -eq 0){ throw 'No valid hosts in target.txt' }

# normalize bucket
$BucketMinutes=[int]$BucketMinutes
if($BucketMinutes -lt 1){ $BucketMinutes=60 }
[double]$frac = [double]$BucketMinutes / 1440.0

# ---- main with guaranteed cleanup ----
$excel=$null; $wb=$null; $wsAll=$null
try{
  # Excel
  try{ $excel=New-Object -ComObject Excel.Application }catch{ throw 'Cannot start Excel COM. Is Excel installed?' }
  $excel.Visible=[bool]$Visible
  $excel.DisplayAlerts=$false
  $wb=$excel.Workbooks.Add()
  while($wb.Worksheets.Count -gt 1){ $wb.Worksheets.Item(1).Delete() }
  $wb.Worksheets.Item(1).Name='AllData'
  $wsAll=$wb.Worksheets('AllData')

  # import csv -> AllData
  try{ foreach($qt in @($wsAll.QueryTables())){ $qt.Delete() } }catch{}
  try{ foreach($lo in @($wsAll.ListObjects())){ $lo.Unlist() } }catch{}
  $wsAll.Cells.Clear()

  $qt=$wsAll.QueryTables.Add('TEXT;' + $csv,$wsAll.Range('A1'))
  $qt.TextFileParseType=$xlDelimited
  $qt.TextFileCommaDelimiter=$true
  $qt.TextFilePlatform=65001
  $qt.TextFileTrailingMinusNumbers=$true
  $qt.AdjustColumnWidth=$true
  $qt.RefreshStyle=$xlInsertDeleteCells
  $null=$qt.Refresh()
  $rng=$qt.ResultRange
  if(-not $rng){ throw 'CSV import failed or empty: ' + $csv }
  $qt.Delete(); $qt=$null

  $loAll=$wsAll.ListObjects.Add($xlSrcRange,$rng,$null,$xlYes)
  $loAll.Name='tblAll'
  $null=$wsAll.Columns.AutoFit()

  function Try-GetColIndex($listObject,[string]$col){ try{ return $listObject.ListColumns($col).Index }catch{ return $null } }
  $colHost=Try-GetColIndex $loAll 'host'
  $colTime=Try-GetColIndex $loAll 'timestamp'
  $colRtt =Try-GetColIndex $loAll 'icmp_avg_ms'
  if($null -eq $colHost -or $null -eq $colTime -or $null -eq $colRtt){ throw 'Required columns missing: host, timestamp, icmp_avg_ms' }
  try{ $loAll.ListColumns('timestamp').DataBodyRange.NumberFormatLocal='yyyy/mm/dd hh:mm' }catch{}
  try{ $loAll.ListColumns('icmp_avg_ms').DataBodyRange.NumberFormatLocal='0.0' }catch{}

  # ★ 全行を一括で 2D 配列として取得（以降、PowerShell側で対象行を選別）
  $hostCol = $loAll.ListColumns('host').DataBodyRange.Value2
  $timeCol = $loAll.ListColumns('timestamp').DataBodyRange.Value2
  $rttCol  = $loAll.ListColumns('icmp_avg_ms').DataBodyRange.Value2
  if(-not ($hostCol -is [Array])){ $tmp=New-Object 'object[,]' 1,1; $tmp[0,0]=$hostCol; $hostCol=$tmp }
  if(-not ($timeCol -is [Array])){ $tmp=New-Object 'object[,]' 1,1; $tmp[0,0]=$timeCol; $timeCol=$tmp }
  if(-not ($rttCol  -is [Array])){ $tmp=New-Object 'object[,]' 1,1; $tmp[0,0]=$rttCol;  $rttCol=$tmp }
  [int]$rowCount = $hostCol.GetLength(0)

  $created=@()
  $ciInv=[System.Globalization.CultureInfo]::InvariantCulture
  $ciCur=[System.Globalization.CultureInfo]::CurrentCulture

  foreach($h in $targets){
    # target を正規化 & 名前→IPv4 も別名として許容
    $targetNorm = Normalize-Host $h
    if([string]::IsNullOrWhiteSpace($targetNorm)){ continue }
    $aliases = Expand-Aliases $targetNorm

    # マッチ行インデックスを収集（厳密一致 or 片包含フォールバック）
    $rowsIdx = New-Object System.Collections.Generic.List[int]
    for($i=0; $i -lt $rowCount; $i++){
      $hn = Normalize-Host ([string]$hostCol[$i,0])
      if([string]::IsNullOrWhiteSpace($hn)){ continue }
      if($aliases.Contains($hn)){ $rowsIdx.Add($i); continue }
      if($hn.Contains($targetNorm) -or $targetNorm.Contains($hn)){ $rowsIdx.Add($i); continue }
    }
    if($rowsIdx.Count -lt 1){ continue }  # このターゲットでは該当なし

    # 対象行を数値に変換（時刻→OADate, RTT→Double）
    $timesRaw = New-Object System.Collections.Generic.List[double]
    $rttsRaw  = New-Object System.Collections.Generic.List[double]
    foreach($ix in $rowsIdx){
      $t = $timeCol[$ix,0]; $r = $rttCol[$ix,0]
      if($t -is [double]){ [double]$tNum=[double]$t } else {
        $tStr=[string]$t
        try{ $dt=[datetime]::Parse($tStr,$ciCur) }catch{ try{ $dt=[datetime]::Parse($tStr,$ciInv) }catch{ continue } }
        [double]$tNum=$dt.ToOADate()
      }
      if($r -is [double]){ [double]$rNum=[double]$r } else {
        $d=0.0
        if([double]::TryParse([string]$r,[System.Globalization.NumberStyles]::Float,$ciInv,[ref]$d)){ [double]$rNum=$d }
        elseif([double]::TryParse([string]$r,[System.Globalization.NumberStyles]::Float,$ciCur,[ref]$d)){ [double]$rNum=$d }
        else{ continue }
      }
      $timesRaw.Add($tNum); $rttsRaw.Add($rNum)
    }

    [int]$rows = $timesRaw.Count
    if($rows -lt 2){ continue }
    $times = $timesRaw.ToArray()
    $rtts  = $rttsRaw.ToArray()

    # 集計 or 生データ（ともに時刻昇順に）
    $xs=@(); $ys=@()
    $agg=@{}
    for($i=0;$i -lt $rows;$i++){
      [double]$t=[double]$times[$i]; [double]$r=[double]$rtts[$i]
      if([double]::IsNaN($t) -or [double]::IsNaN($r)){ continue }
      [double]$bkt=[math]::Floor($t / ([double]$frac)) * ([double]$frac)
      if($agg[$bkt]){ $agg[$bkt].sum=[double]($agg[$bkt].sum+$r); $agg[$bkt].cnt=[int]($agg[$bkt].cnt+1) }
      else{ $agg[$bkt]=@{sum=[double]$r; cnt=[int]1} }
    }
    $useBuckets = ($agg.Keys.Count -ge 2)
    if($useBuckets){
      $keys=@($agg.Keys) | Sort-Object {[double]$_}
      foreach($k in $keys){
        [double]$avg = $agg[$k].sum / [math]::Max(1,[int]$agg[$k].cnt)
        $xs += [double]$k
        $ys += [double]$avg
      }
    } else {
      $pairs=@()
      for($i=0;$i -lt $rows;$i++){
        [double]$t=[double]$times[$i]; [double]$r=[double]$rtts[$i]
        if([double]::IsNaN($t) -or [double]::IsNaN($r)){ continue }
        $pairs += [pscustomobject]@{ t=$t; r=$r }
      }
      $pairs = $pairs | Sort-Object t
      foreach($p in $pairs){ $xs += [double]$p.t; $ys += [double]$p.r }
    }

    # シート作成/クリア
    $sn=Sanitize-SheetName $h
    try{ $ws=$wb.Worksheets.Item($sn); $ws.Cells.Clear(); try{ foreach($co in @($ws.ChartObjects())){ $co.Delete() } }catch{} }
    catch{ $ws=$wb.Worksheets.Add(); $ws.Name=$sn }

    # A/B/Cへ書き出し（A:timestamp/OADate, B:RTT, C:threshold）
    [int]$n=[int]$xs.Count
    if($n -lt 2){ continue }
    $ws.Cells(1,1).Value2='timestamp'
    $ws.Cells(1,2).Value2='icmp_avg_ms'
    $ws.Cells(1,3).Value2='threshold_ms'
    Write-Column2D $ws 'A2' $xs
    Write-Column2D $ws 'B2' $ys
    Write-Column2D $ws 'C2' (New-RepeatedArray -value ([double]$ThresholdMs) -count $n)
    try{ $ws.Range(("A2:A{0}" -f (1+$n))).NumberFormatLocal='yyyy/mm/dd hh:mm' }catch{}
    try{ $ws.Range(("B2:B{0}" -f (1+$n))).NumberFormatLocal='0.0' }catch{}
    $null=$ws.Columns("A:C").AutoFit()

    # グラフ（XY）
    try{ foreach($co in @($ws.ChartObjects())){ $co.Delete() } }catch{}
    $ch=$ws.ChartObjects().Add(300,10,900,320)
    $c=$ch.Chart; $c.ChartType=$xlXYScatterLinesNoMarkers; $c.HasTitle=$true
    $titlePrefix = 'RTT (raw) - '
    if ($useBuckets) { $titlePrefix = 'RTT hourly avg (icmp_avg_ms) - ' }
    $c.ChartTitle.Text = ($titlePrefix + $h)
    $c.Legend.Position=$xlLegendBottom
    try{ $c.SeriesCollection().Delete() }catch{}
    [int]$endRow=1+[int]$n
    $rgb=Get-HostColor $h
    $s1=$c.SeriesCollection().NewSeries(); $s1.Name=$h
    $s1.XValues=$ws.Range(("A2:A{0}" -f $endRow)); $s1.Values=$ws.Range(("B2:B{0}" -f $endRow))
    try{ $s1.Format.Line.ForeColor.RGB=$rgb; $s1.Format.Line.Weight=2 }catch{}
    $s2=$c.SeriesCollection().NewSeries(); $s2.Name=('threshold ' + [int]$ThresholdMs + ' ms')
    $s2.XValues=$ws.Range(("A2:A{0}" -f $endRow)); $s2.Values=$ws.Range(("C2:C{0}" -f $endRow))
    try{ $s2.Format.Line.ForeColor.RGB=255; $s2.Format.Line.Weight=1.5; $s2.Format.Line.DashStyle=$msoLineDash }catch{}
    try{
      $v=$c.Axes($xlValue);    $v.MinimumScale=[double]0; $v.MaximumScale=[double]300; $v.MajorUnit=[double]50; $v.TickLabels.NumberFormat='0.0'
      $x=$c.Axes($xlCategory); $x.MajorUnit=[double](1.0/24.0);                             $x.TickLabels.NumberFormat='mm/dd hh:mm'
    }catch{}

    $created += $sn
    $ws=$null
  }

  if(-not $created -or $created.Count -eq 0){ throw 'No data matched hosts in target.txt' }

  try{
    $wsIdx=$wb.Worksheets.Add(); $wsIdx.Name='INDEX'
    $wsIdx.Cells(1,1).Value2='Hosts'
    $r=2; foreach($sn in $created){ $wsIdx.Hyperlinks.Add($wsIdx.Cells($r,1),'',"'$sn'!A1",'',$sn) | Out-Null; $r++ }
    $wsIdx=$null
  }catch{}

  try{ $wb.Worksheets('AllData').Move($wb.Worksheets.Item($wb.Worksheets.Count)) }catch{}

  $wb.SaveAs($Output)
  Write-Host ('Output: ' + $Output)
}
catch{
  Write-Error (Format-ErrorRecord -Err $_)
  throw
}
finally{
  if($wb){ try{ $wb.Close($false) }catch{}; Release-Com $wb; $wb=$null }
  if($excel){ try{ $excel.Quit() }catch{}; Release-Com $excel; $excel=$null }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}