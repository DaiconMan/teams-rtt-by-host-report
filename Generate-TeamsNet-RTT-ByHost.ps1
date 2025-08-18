<#
Generate-TeamsNet-RTT-ByHost.ps1

目的:
- CSV(teams_net_quality.csv)の host 列から、target.txt に書かれた IP/ホストに合致する行のみ抽出し、時系列グラフを作成
- target.txt で「IP, ラベル」のカンマ区切りをサポート（ラベルはシート名・グラフ表示に使用）
- 表記ゆれ対策: URL/ポート/括弧/末尾ドット/[] の除去による正規化 + 必要に応じ DNS で名前→IP の別名も許容
- 1ホスト=1シート。グラフは A/B/C 列のみをソース (A: timestamp(OADate), B: icmp_avg_ms, C: threshold)
- XY(散布・折れ線, マーカー付) で X=時刻(数値), Y=RTT。Y: 0..300ms(50刻み), X: 1時間刻み。閾値線(既定100ms, 赤破線)
- PS 5.1 互換。Excel COM は成功/失敗に関わらず必ず解放
- 一致0件のときは DEBUG シートに CSV 側の host 分布を出力してからエラー終了

使い方:
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

# 逆引き（失敗はnull）
function Try-ReverseDns([string]$ip){
  try{
    $addr=[System.Net.IPAddress]::Parse($ip)
    $he=[System.Net.Dns]::GetHostEntry($addr)
    if($he.HostName){ return $he.HostName.ToLowerInvariant() }
  }catch{}
  return $null
}

# シート存在チェック
function Has-Sheet($wb,[string]$name){
  try{ $null=$wb.Worksheets.Item($name); return $true }catch{ return $false }
}

# 正規化: URL→host、(ip)抽出、末尾.除去、[]除去、:port除去、lower/trim
function Normalize-Host([string]$s){
  if(-not $s){ return '' }
  $t = $s.Trim().Trim('"',"'").ToLowerInvariant()
  if($t -match '\(([0-9]{1,3}(?:\.[0-9]{1,3}){3})\)'){ return $Matches[1] }
  if($t -match '\[([0-9a-f:]+)\]'){ return $Matches[1] }
  try{
    $uri = $null
    if([System.Uri]::TryCreate($t, [System.UriKind]::Absolute, [ref]$uri) -and $uri.Host){
      $t = $uri.Host.ToLowerInvariant()
    }
  }catch{}
  $t = $t.TrimEnd('.').Trim('[',']')
  $isIPv6 = $false
  try{
    $ipRef = $null
    if([System.Net.IPAddress]::TryParse($t,[ref]$ipRef)){ $isIPv6 = ($ipRef.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6) }
  }catch{}
  if(-not $isIPv6){
    if($t -match '^(.+?):(\d+)$'){ $t = $Matches[1] }
  }
  if($t -match '(^|\s)(\d{1,3}(?:\.\d{1,3}){3})(\s|$)'){ return $Matches[2] }
  return $t
}

# 名前→IP の別名も許容（失敗は無視）
function Expand-Aliases([string]$hostNorm){
  $set = New-Object System.Collections.Generic.HashSet[string]
  if([string]::IsNullOrWhiteSpace($hostNorm)){ return $set }
  [void]$set.Add($hostNorm)
  try{
    $ipRef = $null
    if(-not [System.Net.IPAddress]::TryParse($hostNorm,[ref]$ipRef)){
      $ips = [System.Net.Dns]::GetHostAddresses($hostNorm)
      foreach($ip in $ips){ [void]$set.Add($ip.ToString().ToLowerInvariant()) }
    }
  }catch{}
  return $set
}

# 列探索（ヘッダー表記ゆれ対策）
function Find-ColumnIndex($listObject, [string[]]$candidates){
  $cols = @($listObject.ListColumns)
  foreach($cand in $candidates){
    foreach($col in $cols){
      $name = [string]$col.Name
      if($name -and ($name.Trim().ToLowerInvariant() -eq $cand)){ return $col.Index }
    }
  }
  foreach($cand in $candidates){
    foreach($col in $cols){
      $name = [string]$col.Name
      if($name -and ($name.Trim().ToLowerInvariant() -like "*$cand*")){ return $col.Index }
    }
  }
  return $null
}

# target.txt パーサ: 「キー, ラベル」形式（ラベル省略可）
function Parse-Targets([string]$file){
  $list = New-Object System.Collections.Generic.List[object]
  $lines = Get-Content -Encoding UTF8 $file
  foreach($line in $lines){
    $t = ($line -replace '^\ufeff','').Trim()
    if(-not $t -or $t.StartsWith('#')){ continue }
    $parts = $t.Split(@(','),2,[System.StringSplitOptions]::None)
    $key = $parts[0].Trim()
    $label = if($parts.Count -ge 2){ $parts[1].Trim() } else { '' }
    if($key){
      $list.Add([pscustomobject]@{ Raw=$key; Label=$label })
    }
  }
  return $list
}

# ---- Excel consts ----
[int]$xlDelimited=1; [int]$xlYes=1; [int]$xlLegendBottom=-4107
[int]$xlSrcRange=1; [int]$xlInsertDeleteCells=2
[int]$xlUp=-4162; [int]$xlCategory=1; [int]$xlValue=2
[int]$xlXYScatterLines = 74
[int]$msoLineDash=4

# ---- inputs ----
$csv = Join-Path $InputDir 'teams_net_quality.csv'
Write-Host "[INFO] CSV : $csv"
if(-not (Test-Path $csv)){ throw 'CSV not found: ' + $csv }

if(-not $TargetsFile){
  $TargetsFile = Join-Path $PSScriptRoot 'target.txt'
  if(-not (Test-Path $TargetsFile)){ $TargetsFile = Join-Path $InputDir 'target.txt' }
}
Write-Host "[INFO] TGT : $TargetsFile"
if(-not (Test-Path $TargetsFile)){ throw 'Targets file not found: ' + $TargetsFile }

# targets 読み込み（キー, ラベル）
$targetItems = Parse-Targets $TargetsFile
if(-not $targetItems -or $targetItems.Count -eq 0){ throw 'No valid entries in target.txt' }
Write-Host "[INFO] Targets:"
$targetItems | ForEach-Object { Write-Host ("  - key='{0}' label='{1}'" -f $_.Raw, $_.Label) }

# バケット幅（日単位）
$BucketMinutes=[int]$BucketMinutes
if($BucketMinutes -lt 1){ $BucketMinutes=60 }
[double]$frac = [double]$BucketMinutes / 1440.0

# ---- main with guaranteed cleanup ----
$excel=$null; $wb=$null; $wsAll=$null
try{
  # Excel 起動
  try{ $excel=New-Object -ComObject Excel.Application }catch{ throw 'Cannot start Excel COM. Is Excel installed?' }
  $excel.Visible=[bool]$Visible
  $excel.DisplayAlerts=$false
  $wb=$excel.Workbooks.Add()
  while($wb.Worksheets.Count -gt 1){ $wb.Worksheets.Item(1).Delete() }
  $wb.Worksheets.Item(1).Name='AllData'
  $wsAll=$wb.Worksheets('AllData')

  # CSV インポート → テーブル化
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

  # 列インデックス特定
  $colHost=Find-ColumnIndex $loAll @('host','hostname','target','dst_host','dest','remote_host')
  $colTime=Find-ColumnIndex $loAll @('timestamp','time','datetime','date')
  $colRtt =Find-ColumnIndex $loAll @('icmp_avg_ms','rtt_ms','avg_rtt','avg_rtt_ms','icmp_avg','icmp_rtt_ms')
  if($null -eq $colHost -or $null -eq $colTime -or $null -eq $colRtt){
    $names = (@($loAll.ListColumns) | ForEach-Object { $_.Name }) -join ', '
    throw ("Required columns missing: host/timestamp/icmp_avg_ms (found: {0})" -f $names)
  }
  try{ $loAll.ListColumns($colTime).DataBodyRange.NumberFormatLocal='yyyy/mm/dd hh:mm' }catch{}
  try{ $loAll.ListColumns($colRtt ).DataBodyRange.NumberFormatLocal='0.0' }catch{}

  # ---- 列値を 1始まり2D で取得し、境界を把握 ----
  $hostCol = $loAll.ListColumns($colHost).DataBodyRange.Value2
  $timeCol = $loAll.ListColumns($colTime).DataBodyRange.Value2
  $rttCol  = $loAll.ListColumns($colRtt ).DataBodyRange.Value2

  if(-not ($hostCol -is [Array])){ $tmp = New-Object 'object[,]' 1,1; $tmp[1,1] = $hostCol; $hostCol = $tmp }
  if(-not ($timeCol -is [Array])){ $tmp = New-Object 'object[,]' 1,1; $tmp[1,1] = $timeCol; $timeCol = $tmp }
  if(-not ($rttCol  -is [Array])){ $tmp = New-Object 'object[,]' 1,1; $tmp[1,1] = $rttCol;  $rttCol  = $tmp }

  $Rlo = $hostCol.GetLowerBound(0); $Rhi = $hostCol.GetUpperBound(0)
  $Clo = $hostCol.GetLowerBound(1); $Chi = $hostCol.GetUpperBound(1)
  [int]$rowCount = $Rhi - $Rlo + 1

  [int]$hostColIndex = $Clo
  [int]$timeColIndex = $timeCol.GetLowerBound(1)
  [int]$rttColIndex  = $rttCol.GetLowerBound(1)

  Write-Host "[DEBUG] bounds host: rows $Rlo..$Rhi cols $Clo..$Chi ; rowCount=$rowCount"

  # DEBUG 用: CSV 側の正規化 host 分布
  $csvNorm = @{}
  for($i=$Rlo; $i -le $Rhi; $i++){
    $raw  = [string]$hostCol[$i, $hostColIndex]
    $norm = Normalize-Host $raw
    if(-not $csvNorm.ContainsKey($norm)){ $csvNorm[$norm] = @{count=0; sample=$raw} }
    $csvNorm[$norm].count++
  }

  $created=@()
  $ciInv=[System.Globalization.CultureInfo]::InvariantCulture
  $ciCur=[System.Globalization.CultureInfo]::CurrentCulture

  foreach($item in $targetItems){
    $h = $item.Raw
    $labelHint = $item.Label
    $targetNorm = Normalize-Host $h
    if([string]::IsNullOrWhiteSpace($targetNorm)){ continue }
    $aliases = Expand-Aliases $targetNorm

    # ---- マッチ行抽出 ----
    $rowsIdx = New-Object System.Collections.Generic.List[int]
    for($i=$Rlo; $i -le $Rhi; $i++){
      $raw = [string]$hostCol[$i, $hostColIndex]
      $hn  = Normalize-Host $raw
      if([string]::IsNullOrWhiteSpace($hn)){ continue }
      if($aliases.Contains($hn)){ $rowsIdx.Add($i); continue }
      if($hn.Contains($targetNorm) -or $targetNorm.Contains($hn)){ $rowsIdx.Add($i); continue }
      $rawL = ('' + $raw).Trim().ToLowerInvariant()
      if($rawL.Contains($targetNorm)){ $rowsIdx.Add($i); continue }
    }
    Write-Host ("[MATCH] key='{0}' norm='{1}' aliases=[{2}] -> rows={3}" -f $h, $targetNorm, ([string]::Join(',', $aliases)), $rowsIdx.Count)
    if($rowsIdx.Count -lt 1){ continue }

    # ---- ラベル（表示名）決定 ----
    $label = $labelHint
    if([string]::IsNullOrWhiteSpace($label)){
      # keyがIPなら逆引き→無ければCSV最頻の名前部分、keyがホストならそれを使う
      if($targetNorm -match '^\d{1,3}(?:\.\d{1,3}){3}$'){
        $rdns = Try-ReverseDns $targetNorm
        if($rdns){ $label = $rdns } else {
          $counts=@{}
          foreach($ix in $rowsIdx){
            $raw = ('' + $hostCol[$ix, $hostColIndex])
            if(-not $counts.ContainsKey($raw)){ $counts[$raw]=0 }
            $counts[$raw]++
          }
          if($counts.Count -gt 0){
            $topRaw = ($counts.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
            $namePart = ('' + $topRaw) -replace '\s*\(.*$',''
            if($namePart){ $label = $namePart.ToLowerInvariant() }
          }
          if([string]::IsNullOrWhiteSpace($label)){ $label = $h }
        }
      } else {
        $label = $h
      }
    }

    # ---- 数値化（時刻→OADate, RTT→Double）----
    $timesRaw = New-Object System.Collections.Generic.List[double]
    $rttsRaw  = New-Object System.Collections.Generic.List[double]
    foreach($ix in $rowsIdx){
      $t = $timeCol[$ix, $timeColIndex]
      $r = $rttCol[$ix, $rttColIndex]
      if($t -is [double]){ [double]$tNum=[double]$t }
      else{
        $tStr=[string]$t
        try{ $dt=[datetime]::Parse($tStr,$ciCur) }catch{ try{ $dt=[datetime]::Parse($tStr,$ciInv) }catch{ continue } }
        [double]$tNum=$dt.ToOADate()
      }
      if($r -is [double]){ [double]$rNum=[double]$r }
      else{
        $d=0.0
        if([double]::TryParse([string]$r,[System.Globalization.NumberStyles]::Float,$ciInv,[ref]$d)){ [double]$rNum=$d }
        elseif([double]::TryParse([string]$r,[System.Globalization.NumberStyles]::Float,$ciCur,[ref]$d)){ [double]$rNum=$d }
        else{ continue }
      }
      $timesRaw.Add($tNum); $rttsRaw.Add($rNum)
    }

    [int]$rowsM = $timesRaw.Count
    if($rowsM -lt 1){ continue }
    $times = $timesRaw.ToArray()
    $rtts  = $rttsRaw.ToArray()

    # ---- 集計 or 生データ（昇順）----
    $xs=@(); $ys=@()
    $agg=@{}
    for($i2=0;$i2 -lt $rowsM;$i2++){
      [double]$t=[double]$times[$i2]; [double]$r=[double]$rtts[$i2]
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
      for($i3=0;$i3 -lt $rowsM;$i3++){
        [double]$t=[double]$times[$i3]; [double]$r=[double]$rtts[$i3]
        if([double]::IsNaN($t) -or [double]::IsNaN($r)){ continue }
        $pairs += [pscustomobject]@{ t=$t; r=$r }
      }
      $pairs = $pairs | Sort-Object t
      foreach($p in $pairs){ $xs += [double]$p.t; $ys += [double]$p.r }
    }

    # ---- シート作成 ----
    [int]$n=[int]$xs.Count
    if($n -lt 1){ continue }

    $snBase = Sanitize-SheetName $label
    $sn = $snBase
    $suffix = 2
    while(Has-Sheet $wb $sn){ $sn = Sanitize-SheetName ($snBase + "_" + $suffix); $suffix++ }

    try{ $ws=$wb.Worksheets.Add(); $ws.Name=$sn }
    catch{ $ws=$wb.Worksheets.Add() }

    # A/B/C へ書き出し
    $ws.Cells(1,1).Value2='timestamp'
    $ws.Cells(1,2).Value2='icmp_avg_ms'
    $ws.Cells(1,3).Value2='threshold_ms'
    Write-Column2D $ws 'A2' $xs
    Write-Column2D $ws 'B2' $ys
    Write-Column2D $ws 'C2' (New-RepeatedArray -value ([double]$ThresholdMs) -count $n)
    try{ $ws.Range(("A2:A{0}" -f (1+$n))).NumberFormatLocal='yyyy/mm/dd hh:mm' }catch{}
    try{ $ws.Range(("B2:B{0}" -f (1+$n))).NumberFormatLocal='0.0' }catch{}
    $null=$ws.Columns("A:C").AutoFit()

    # ---- グラフ（XY 散布・折れ線・マーカー付）----
    try{ foreach($co in @($ws.ChartObjects())){ $co.Delete() } }catch{}
    $ch=$ws.ChartObjects().Add(300,10,900,320)
    $c=$ch.Chart; $c.ChartType=$xlXYScatterLines; $c.HasTitle=$true
    $titlePrefix = 'RTT (raw) - '
    if ($useBuckets) { $titlePrefix = 'RTT hourly avg (icmp_avg_ms) - ' }
    $c.ChartTitle.Text = ($titlePrefix + $label)
    $c.Legend.Position=$xlLegendBottom
    try{ $c.SeriesCollection().Delete() }catch{}
    [int]$endRow=1+[int]$n
    $rgb=Get-HostColor $label

    $s1=$c.SeriesCollection().NewSeries(); $s1.Name=$label
    $s1.XValues=$ws.Range(("A2:A{0}" -f $endRow)); $s1.Values=$ws.Range(("B2:B{0}" -f $endRow))
    try{
      $s1.Format.Line.ForeColor.RGB=$rgb; $s1.Format.Line.Weight=2
      $s1.MarkerStyle = 8; $s1.MarkerSize = 5
    }catch{}

    $s2=$c.SeriesCollection().NewSeries(); $s2.Name=('threshold ' + [int]$ThresholdMs + ' ms')
    $s2.XValues=$ws.Range(("A2:A{0}" -f $endRow)); $s2.Values=$ws.Range(("C2:C{0}" -f $endRow))
    try{ $s2.Format.Line.ForeColor.RGB=255; $s2.Format.Line.Weight=1.5; $s2.Format.Line.DashStyle=$msoLineDash }catch{}

    try{
      $v=$c.Axes($xlValue);    $v.MinimumScale=[double]0; $v.MaximumScale=[double]300; $v.MajorUnit=[double]50; $v.TickLabels.NumberFormat='0.0'
      $x=$c.Axes($xlCategory); $x.MajorUnit=[double](1.0/24.0);                             $x.TickLabels.NumberFormat='mm/dd hh:mm'
    }catch{}

    $created += $sn
    $ws=$null
    Write-Host ("[INFO] label='{0}': rowsIdx={1}, plotted points={2}" -f $label, $rowsIdx.Count, $n)
  }

  if(-not $created -or $created.Count -eq 0){
    # DEBUG シートを書き出してからエラー終了
    try{
      $wsDbg = $wb.Worksheets.Add(); $wsDbg.Name='DEBUG'
      $wsDbg.Cells(1,1).Value2='csv_host_sample'
      $wsDbg.Cells(1,2).Value2='normalized'
      $wsDbg.Cells(1,3).Value2='count'
      $r=2
      $items = $csvNorm.GetEnumerator() | Sort-Object { $_.Value.count } -Descending
      foreach($it in $items){
        $wsDbg.Cells($r,1).Value2 = $it.Value.sample
        $wsDbg.Cells($r,2).Value2 = $it.Key
        $wsDbg.Cells($r,3).Value2 = [int]$it.Value.count
        $r++
      }
      $wsDbg.Columns.AutoFit() | Out-Null
      $wb.SaveAs($Output)
      Write-Warning ("No data matched keys in target.txt. DEBUG sheet written to: {0}" -f $Output)
    }catch{}
    throw 'No data matched keys in target.txt'
  }

  # INDEX
  try{
    $wsIdx=$wb.Worksheets.Add(); $wsIdx.Name='INDEX'
    $wsIdx.Cells(1,1).Value2='Hosts'
    $r=2; foreach($sn in $created){ $wsIdx.Hyperlinks.Add($wsIdx.Cells($r,1),'',"'$sn'!A1",'',$sn) | Out-Null; $r++ }
    $wsIdx=$null
  }catch{}

  # AllData を末尾へ
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