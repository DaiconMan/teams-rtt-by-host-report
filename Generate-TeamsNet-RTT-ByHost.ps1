<#
Generate-TeamsNet-RTT-ByHost.ps1
- ピボット非依存。target.txt のホストだけ抽出し、RTT(icmp_avg_ms)を「ホストごと1枚」の折れ線で出力。
- 仕様:
  * 日付表示 mm/dd hh:mm（X軸は時間スケール）
  * ホストごとに線色固定（安定ハッシュ）
  * 閾値線 100ms（赤・破線）
  * 縦軸 0〜300ms 固定（主目盛 50ms）
  * 文字列として取り込まれた timestamp/RTT を強制数値化してグラフ化
#>

param(
  [string]$InputDir   = (Join-Path $Env:LOCALAPPDATA "TeamsNet"),
  [Parameter(Mandatory=$true)][string]$Output,
  [string]$TargetsFile,   # 省略時: スクリプト隣の target.txt → %LOCALAPPDATA%\TeamsNet\target.txt
  [int]$ThresholdMs = 100,
  [switch]$Visible
)

# ===== Excel 定数 =====
$xlDelimited            = 1
$xlYes                  = 1
$xlLine                 = 4
$xlLegendPositionBottom = -4107
$xlSrcRange             = 1
$xlInsertDeleteCells    = 2
$xlCellTypeVisible      = 12
$xlUp                   = -4162
$xlCategory             = 1
$xlValue                = 2
$xlTimeScale            = 3
$msoLineDash            = 4

$ErrorActionPreference = "Stop"

# 入力CSV
$csv = Join-Path $InputDir "teams_net_quality.csv"
if(-not (Test-Path $csv)){ throw "CSV が見つかりません: $csv" }

# target.txt 探索
if(-not $TargetsFile){
  $TargetsFile = Join-Path $PSScriptRoot "target.txt"
  if(-not (Test-Path $TargetsFile)){
    $TargetsFile = Join-Path $InputDir "target.txt"
  }
}
if(-not (Test-Path $TargetsFile)){ throw "TargetsFile が見つかりません: $TargetsFile" }

# 対象ホストの読み込み
$targets = Get-Content -Raw -Encoding UTF8 $TargetsFile |
  ForEach-Object { $_ -split "`r?`n" } |
  Where-Object { $_ -and (-not $_.Trim().StartsWith("#")) } |
  ForEach-Object { $_.Trim() } | Select-Object -Unique
if(-not $targets -or $targets.Count -eq 0){ throw "target.txt に有効なホストがありません。" }

# 出力フォルダ
$outDir = Split-Path -Parent $Output
if($outDir -and -not (Test-Path $outDir)){ New-Item -ItemType Directory -Path $outDir | Out-Null }

# Excel 起動
try { $excel = New-Object -ComObject Excel.Application } catch { throw "Excel の COM を起動できません（Excel が必要）。" }
$excel.Visible = [bool]$Visible
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Add()

# 既定シート整理
while($wb.Worksheets.Count -gt 1){ $wb.Worksheets.Item(1).Delete() }
$wb.Worksheets.Item(1).Name = "AllData"

function Sanitize-SheetName([string]$name){
  if(-not $name){ return "Host" }
  $n = $name -replace '[:\\/\?\*\[\]]','_'
  if($n.Length -gt 31){ $n = $n.Substring(0,31) }
  if($n -match '^\s*$'){ $n = "Host" }
  return $n
}

# $host 自動変数衝突を避けるため、引数名は $HostName に
function Get-HostColor([string]$HostName){
  $palette = @(
    @{r= 33; g=150; b=243}, @{r= 76; g=175; b= 80}, @{r=244; g= 67; b= 54},
    @{r=255; g=193; b=  7}, @{r=156; g= 39; b=176}, @{r=  0; g=188; b=212},
    @{r=121; g= 85; b= 72}, @{r= 63; g= 81; b=181}, @{r=205; g=220; b= 57},
    @{r=233; g= 30; b= 99}
  )
  $sum = 0; $HostName.ToCharArray() | ForEach-Object { $sum += [int]$_ }
  $idx = $sum % $palette.Count
  $c = $palette[$idx]
  return [int]($c.r + ($c.g -shl 8) + ($c.b -shl 16))
}

# AllData 取り込み（QueryTable → データ残し → テーブル化）
$wsAll = $wb.Worksheets("AllData")
try { foreach($qt in @($wsAll.QueryTables())){ $qt.Delete() } } catch {}
try { foreach($lo in @($wsAll.ListObjects())){ $lo.Unlist() } } catch {}
$wsAll.Cells.Clear()

$qt = $wsAll.QueryTables.Add("TEXT;" + $csv, $wsAll.Range("A1"))
$qt.TextFileParseType            = $xlDelimited
$qt.TextFileCommaDelimiter       = $true
$qt.TextFilePlatform             = 65001
$qt.TextFileTrailingMinusNumbers = $true
$qt.AdjustColumnWidth            = $true
$qt.RefreshStyle                 = $xlInsertDeleteCells
$qt.Refresh() | Out-Null

$rng = $qt.ResultRange
if(-not $rng){ throw "CSV にデータが無いか、取り込みに失敗しました: $csv" }
$qt.Delete()

$loAll = $wsAll.ListObjects.Add($xlSrcRange, $rng, $null, $xlYes)
$loAll.Name = "tblAll"
$wsAll.Columns.AutoFit() | Out-Null

# 必須列
function Try-GetColIndex($listObject, [string]$colName){ try { return $listObject.ListColumns($colName).Index } catch { return $null } }
$colHost = Try-GetColIndex $loAll "host"
$colTime = Try-GetColIndex $loAll "timestamp"
$colRtt  = Try-GetColIndex $loAll "icmp_avg_ms"
if($null -eq $colHost -or $null -eq $colTime -or $null -eq $colRtt){
  throw "必要列 'host','timestamp','icmp_avg_ms' が見つかりません。"
}

# 表示形式だけ先に
try { $loAll.ListColumns("timestamp").DataBodyRange.NumberFormat = "mm/dd hh:mm" } catch {}
try { $loAll.ListColumns("icmp_avg_ms").DataBodyRange.NumberFormat = "0.0" } catch {}

# ===== 各ホストごと =====
$created = @()
foreach($h in $targets){
  # フィルター
  $loAll.Range.AutoFilter($colHost, $h) | Out-Null
  $visTime = $loAll.ListColumns("timestamp").DataBodyRange.SpecialCells($xlCellTypeVisible)
  $visRtt  = $loAll.ListColumns("icmp_avg_ms").DataBodyRange.SpecialCells($xlCellTypeVisible)

  # 可視行数
  $rowCount = 0
  try { $rowCount = $visTime.Areas | ForEach-Object { $_.Rows.Count } | Measure-Object -Sum | Select-Object -ExpandProperty Sum } catch { $rowCount = 0 }

  if($rowCount -gt 0){
    $sheetName = Sanitize-SheetName $h
    try {
      $ws = $wb.Worksheets.Item($sheetName)
      $ws.Cells.Clear()
      try { foreach($co in @($ws.ChartObjects())){ $co.Delete() } } catch {}
    } catch {
      $ws = $wb.Worksheets.Add(); $ws.Name = $sheetName
    }

    # 見出し
    $ws.Cells(1,1).Value2 = "timestamp"
    $ws.Cells(1,2).Value2 = "icmp_avg_ms"
    $ws.Cells(1,3).Value2 = "threshold_ms"

    # 可視セルを貼付
    $visTime.Copy($ws.Range("A2")) | Out-Null
    $visRtt.Copy($ws.Range("B2"))  | Out-Null

    # 最終行
    $lastRow = $ws.Cells($ws.Rows.Count, 1).End($xlUp).Row
    if($lastRow -lt 2){
      $loAll.AutoFilter.ShowAllData() | Out-Null 2>$null
      continue
    }

    # === 重要: 文字列→数値へ強制変換（X:日時、Y:RTT） ===
    # A列: "yyyy-mm-dd hh:mm:ss" などを LEFT/MID + DATEVALUE/TIMEVALUE で確実に数値化
    $ws.Range("D1").Value2 = "ts_numeric"
    $ws.Range("D2:D$lastRow").FormulaR1C1 = "=DATEVALUE(LEFT(RC[-3],10))+TIMEVALUE(MID(RC[-3],12,8))"
    $ws.Range("A2:A$lastRow").Value2 = $ws.Range("D2:D$lastRow").Value2
    $ws.Range("D:D").Clear()

    # B列: RTT も VALUE() で数値化（千区切りや文字列対策）
    $ws.Range("E1").Value2 = "rtt_numeric"
    $ws.Range("E2:E$lastRow").FormulaR1C1 = "=IF(RC[-3]="""","""",VALUE(RC[-3]))"
    $ws.Range("B2:B$lastRow").Value2 = $ws.Range("E2:E$lastRow").Value2
    $ws.Range("E:E").Clear()

    # 閾値列
    $ws.Range("C2:C$lastRow").Value2 = $ThresholdMs

    # 表示形式
    $ws.Range("A2:A$lastRow").NumberFormat = "mm/dd hh:mm"
    $ws.Range("B2:B$lastRow").NumberFormat = "0.0"

    # グラフ
    try { foreach($co in @($ws.ChartObjects())){ $co.Delete() } } catch {}
    $ch  = $ws.ChartObjects().Add(300, 10, 900, 320)
    $ch.Name = "chtRTT"
    $chC = $ch.Chart
    $chC.ChartType = $xlLine
    $chC.HasTitle  = $true
    $chC.ChartTitle.Text = "RTT (icmp_avg_ms) - " + $h
    $chC.Legend.Position = $xlLegendPositionBottom
    try { $chC.SeriesCollection().Delete() } catch {}

    # RTTシリーズ（ホスト色）
    $sColor = Get-HostColor $h
    $s1 = $chC.SeriesCollection().NewSeries()
    $s1.Name    = $h
    $s1.XValues = $ws.Range("A2:A$lastRow")
    $s1.Values  = $ws.Range("B2:B$lastRow")
    try {
      $s1.Format.Line.ForeColor.RGB = $sColor
      $s1.Format.Line.Weight = 2
    } catch {}

    # 閾値（赤・破線）
    $s2 = $chC.SeriesCollection().NewSeries()
    $s2.Name    = "threshold " + $ThresholdMs + "ms"
    $s2.XValues = $ws.Range("A2:A$lastRow")
    $s2.Values  = $ws.Range("C2:C$lastRow")
    try {
      $s2.Format.Line.ForeColor.RGB = 255
      $s2.Format.Line.Weight = 1.5
      $s2.Format.Line.DashStyle = $msoLineDash
    } catch {}

    # 軸：Y=0..300 / X=1時間
    try {
      $valAx = $chC.Axes($xlValue)
      $valAx.MinimumScale = 0
      $valAx.MaximumScale = 300
      $valAx.MajorUnit    = 50
      $valAx.TickLabels.NumberFormat = "0.0"
    } catch {}
    try {
      $catAx = $chC.Axes($xlCategory)
      $catAx.CategoryType = $xlTimeScale
      $catAx.MajorUnit    = 1/24
      $catAx.TickLabels.NumberFormat = "mm/dd hh:mm"
    } catch {}

    $created += $sheetName
  }

  # フィルタ解除
  $loAll.AutoFilter.ShowAllData() | Out-Null 2>$null
}

if(-not $created -or $created.Count -eq 0){
  throw "target.txt のホストに一致するデータがありませんでした。"
}

# INDEX
try {
  $wsIdx = $wb.Worksheets.Add()
  $wsIdx.Name = "INDEX"
  $wsIdx.Cells(1,1).Value2 = "Hosts"
  $r = 2
  foreach($sn in $created){
    $wsIdx.Hyperlinks.Add($wsIdx.Cells($r,1), "", "'$sn'!A1", "", $sn) | Out-Null
    $r++
  }
} catch {}

# AllData を右端へ
try { $wb.Worksheets("AllData").Move($wb.Worksheets.Item($wb.Worksheets.Count)) } catch {}

# 保存・終了
$wb.SaveAs($Output)
$wb.Close($true)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)    | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

Write-Host "出力しました: $Output"
