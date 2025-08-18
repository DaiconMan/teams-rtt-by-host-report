<#
Generate-TeamsNet-LayerDelta.ps1

目的:
- 既存の teams_net_quality.csv から、各レイヤ(L2/L3/RTR_LAN/RTR_WAN/ZSCALER/SAAS)のRTTを時間バケット集計
- 役割ごとに「有効指標」を自動選択 (ICMP/TCP/HTTP) → FQDNのICMP不可でも分析可能
- ΔL3, ΔRtrLAN, ΔRtrWAN, ΔCloud を算出し、Excelに「レイヤ系列」と「Δ系列」のグラフを出力
- PowerShell 5.1 互換。Excel COM は成功/失敗に関わらず必ず解放

使い方:
  powershell -NoProfile -ExecutionPolicy Bypass `
    -File .\Generate-TeamsNet-LayerDelta.ps1 `
    -CsvPath "$Env:LOCALAPPDATA\TeamsNet\teams_net_quality.csv" `
    -TargetsCsv ".\targets.csv" `
    -Output ".\LayerDelta-Report.xlsx" `
    -BucketMinutes 5
#>

param(
  [Parameter(Mandatory=$true)][string]$CsvPath,
  [Parameter(Mandatory=$true)][string]$TargetsCsv,
  [Parameter(Mandatory=$true)][string]$Output,
  [int]$BucketMinutes = 5,
  [switch]$Visible
)

# ------------- 共通ユーティリティ -------------
$ErrorActionPreference='Stop'
$PSDefaultParameterValues['*:ErrorAction']='Stop'

function Release-Com([object]$obj){
  if($null -ne $obj -and [System.Runtime.InteropServices.Marshal]::IsComObject($obj)){
    try{ [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) }catch{}
  }
}
function Sanitize-SheetName([string]$name){
  if(-not $name){ return 'Sheet' }
  $n = $name -replace '[:\\/\?\*\[\]]','_'
  if($n.Length -gt 31){ $n=$n.Substring(0,31) }
  if($n -match '^\s*$'){ $n='Sheet' }
  return $n
}
function Normalize-Host([string]$s){
  if(-not $s){ return '' }
  $t = $s.Trim().Trim('"',"'").ToLowerInvariant()
  if($t -match '\(([0-9]{1,3}(?:\.[0-9]{1,3}){3})\)'){ return $Matches[1] }  # name (ip)
  if($t -match '\[([0-9a-f:]+)\]'){ return $Matches[1] }                    # name [ipv6]
  try{ $uri=$null; if([System.Uri]::TryCreate($t,[System.UriKind]::Absolute,[ref]$uri) -and $uri.Host){ $t=$uri.Host.ToLowerInvariant() } }catch{}
  $t = $t.TrimEnd('.').Trim('[',']')
  $isIPv6=$false; try{ $ip=$null; if([System.Net.IPAddress]::TryParse($t,[ref]$ip)){ $isIPv6=($ip.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6) } }catch{}
  if(-not $isIPv6){ if($t -match '^(.+?):(\d+)$'){ $t=$Matches[1] } }
  if($t -match '(^|\s)(\d{1,3}(?:\.\d{1,3}){3})(\s|$)'){ return $Matches[2] }
  return $t
}
function Format-ErrorRecord([System.Management.Automation.ErrorRecord]$Err){
  $ii=$Err.InvocationInfo
  $msg=@()
  $msg+="[ERROR] $($Err.FullyQualifiedErrorId)"
  if($ii){ $msg+=" at line $($ii.ScriptLineNumber) char $($ii.OffsetInLine): $($ii.Line)" }
  $ex=$Err.Exception
  if($ex){ $msg+=" $($ex.GetType().FullName): $($ex.Message)"; if($ex.InnerException){ $msg+=" -> $($ex.InnerException.Message)" } }
  return ($msg -join "`r`n")
}

# ------------- targets.csv 読み込み（プレースホルダ解決） -------------
function Get-DefaultGatewayIPv4(){
  try{
    $gw = Get-NetIPConfiguration | Where-Object { $_.IPv4DefaultGateway -and $_.NetAdapter.Status -eq 'Up' } | Select-Object -First 1
    if($gw){ return $gw.IPv4DefaultGateway.NextHop }
  }catch{}
  return $null
}
function Get-HopN([int]$n){
  try{
    $out = tracert -4 -d -h $n 8.8.8.8 2>$null
    # 1行に "  n   <ms> <ms> <ms> ip" のような行が来る想定
    foreach($line in $out){
      if($line -match "^\s*$n\s+\S+\s+\S+\s+\S+\s+(\d{1,3}(?:\.\d{1,3}){3})\s*$"){ return $Matches[1] }
      if($line -match "^\s*$n\s+(\d{1,3}(?:\.\d{1,3}){3})\s*$"){ return $Matches[1] }
    }
  }catch{}
  return $null
}
function Parse-TargetsCsv([string]$path){
  if(-not (Test-Path $path)){ throw "targets.csv not found: $path" }
  $rows = Import-Csv -Path $path -Encoding UTF8
  if(-not $rows -or $rows.Count -eq 0){ throw "targets.csv is empty: $path" }

  $gw = Get-DefaultGatewayIPv4
  $hop2 = Get-HopN 2
  $hop3 = Get-HopN 3

  $list = New-Object System.Collections.Generic.List[object]
  foreach($r in $rows){
    $role  = ('' + $r.role).Trim().ToUpperInvariant()
    $key   = ('' + $r.key ).Trim()
    $label = ('' + $r.label).Trim()
    if(-not $role -or -not $key){ continue }

    # プレースホルダ置換
    if($key -eq '{GATEWAY}' -and $gw){ $key=$gw }
    if($key -eq '{HOP2}'    -and $hop2){ $key=$hop2 }
    if($key -eq '{HOP3}'    -and $hop3){ $key=$hop3 }

    $list.Add([pscustomobject]@{
      Role = $role
      Key  = $key
      KeyNorm = Normalize-Host $key
      Label = (if($label){ $label } else { $key })
    })
  }
  if($list.Count -eq 0){ throw "No valid entries in targets.csv (after placeholders)" }
  return $list
}

# ------------- CSV 読み込み -------------
if(-not (Test-Path $CsvPath)){ throw "CSV not found: $CsvPath" }
$data = Import-Csv -Path $CsvPath -Encoding UTF8
if(-not $data -or $data.Count -eq 0){ throw "CSV is empty: $CsvPath" }

# 列名の同義語
$COL_HOST = @('host','hostname','target','dst_host','dest','remote_host')
$COL_TIME = @('timestamp','time','datetime','date')
$COL_ICMP = @('icmp_avg_ms','rtt_ms','avg_rtt','avg_rtt_ms','icmp_avg','icmp_rtt_ms')
$COL_TCP  = @('tcp_ms','tcp_connect_ms','tcp443_ms')
$COL_HTTP = @('http_ms','http_head_ms','http_head_rtt_ms')
$COL_DNS  = @('dns_ms','dns_lookup_ms','dns_rtt_ms')

# 実在ヘッダーを小文字化して取得名を決める
$headers = @{}
$data[0].PSObject.Properties.Name | ForEach-Object { $headers[$_.ToLowerInvariant()] = $_ }

function Resolve-Col([string[]]$cands){
  foreach($c in $cands){ if($headers.ContainsKey($c)){ return $headers[$c] } }
  foreach($c in $cands){ # 部分一致
    foreach($k in $headers.Keys){ if($k -like "*$c*"){ return $headers[$k] } }
  }
  return $null
}

$hn = Resolve-Col $COL_HOST; if(-not $hn){ throw "host column not found" }
$tn = Resolve-Col $COL_TIME; if(-not $tn){ throw "timestamp column not found" }
$in = Resolve-Col $COL_ICMP
$tc = Resolve-Col $COL_TCP
$ht = Resolve-Col $COL_HTTP
$dn = Resolve-Col $COL_DNS

# ------------- targets マップ構築 -------------
$targets = Parse-TargetsCsv $TargetsCsv
# 役割→(keynorm,label)配列
$roleKeys = @{}
foreach($t in $targets){
  if(-not $roleKeys.ContainsKey($t.Role)){ $roleKeys[$t.Role] = New-Object System.Collections.Generic.List[object] }
  $roleKeys[$t.Role].Add($t)
}

# ------------- 有効RTTの選択ロジック -------------
function To-DoubleOrNull($v){
  if($v -is [double]){ return [double]$v }
  $s = ('' + $v).Trim()
  if(-not $s){ return $null }
  $d=0.0
  if([double]::TryParse($s,[System.Globalization.NumberStyles]::Float,[System.Globalization.CultureInfo]::InvariantCulture,[ref]$d)){ return [double]$d }
  if([double]::TryParse($s,[System.Globalization.NumberStyles]::Float,[System.Globalization.CultureInfo]::CurrentCulture,[ref]$d)){ return [double]$d }
  return $null
}
function Pick-EffRtt([string]$role, [object]$row){
  # 優先順位: 役割ごと
  # L2/L3/RTR_* => ICMP > TCP > HTTP
  # ZSCALER/SAAS => TCP > HTTP > ICMP
  $icmp = if($in){ To-DoubleOrNull $row.$in } else { $null }
  $tcp  = if($tc){ To-DoubleOrNull $row.$tc } else { $null }
  $http = if($ht){ To-DoubleOrNull $row.$ht } else { $null }

  $cat = $role
  if($role -like 'RTR*' -or $role -eq 'L2' -or $role -eq 'L3'){
    if($icmp -ne $null){ return ,@($icmp,'icmp') }
    if($tcp  -ne $null){ return ,@($tcp ,'tcp') }
    if($http -ne $null){ return ,@($http,'http') }
  } else { # SAAS/ZSCALER/その他
    if($tcp  -ne $null){ return ,@($tcp ,'tcp') }
    if($http -ne $null){ return ,@($http,'http') }
    if($icmp -ne $null){ return ,@($icmp,'icmp') }
  }
  return ,@($null,'')
}

# ------------- マッチ & 時間バケット集計 -------------
if($BucketMinutes -lt 1){ $BucketMinutes = 5 }
[double]$frac = [double]$BucketMinutes / 1440.0 # 日単位

# バケット: key=OADate(double)，値= { role => [values...] }
$buckets = @{}

$ciCur=[System.Globalization.CultureInfo]::CurrentCulture
$ciInv=[System.Globalization.CultureInfo]::InvariantCulture

# host 正規化して targets に照合
foreach($row in $data){
  $hraw = '' + $row.$hn
  if(-not $hraw){ continue }
  $hnorm = Normalize-Host $hraw

  $ts = '' + $row.$tn
  if(-not $ts){ continue }
  try{ $dt=[datetime]::Parse($ts,$ciCur) }catch{ try{ $dt=[datetime]::Parse($ts,$ciInv) }catch{ continue } }
  [double]$tOa = $dt.ToOADate()
  [double]$bucket = [math]::Floor($tOa / $frac) * $frac

  foreach($role in $roleKeys.Keys){
    $matched=$false; $label=''
    foreach($t in $roleKeys[$role]){
      if($hnorm -eq $t.KeyNorm -or $hnorm.Contains($t.KeyNorm) -or $t.KeyNorm.Contains($hnorm)){
        $matched=$true; $label=$t.Label; break
      }
    }
    if(-not $matched){ continue }

    $pair = Pick-EffRtt $role $row
    $val = $pair[0]; $src = $pair[1]
    if($val -eq $null){ continue }

    if(-not $buckets.ContainsKey($bucket)){ $buckets[$bucket]=@{} }
    if(-not $buckets[$bucket].ContainsKey($role)){ $buckets[$bucket][$role]=New-Object System.Collections.Generic.List[double] }
    $buckets[$bucket][$role].Add([double]$val)
  }
}

if($buckets.Count -eq 0){ throw "No data matched targets.csv (check key normalization and headers)" }

# ------------- 平均化 → レイヤ系列/Δ系列を作成 -------------
# 使用役割の順序（存在するもののみ）
$roleOrder = @('L2','L3','RTR_LAN','RTR_WAN','ZSCALER','SAAS') | Where-Object { $roleKeys.ContainsKey($_) }

$X = @()
$series = @{}  # role => values aligned to X
foreach($r in $roleOrder){ $series[$r]=@() }

# Δ 系列
$DeltaNames = @('DELTA_L3','DELTA_RTR_LAN','DELTA_RTR_WAN','DELTA_CLOUD')
$delta = @{}
foreach($d in $DeltaNames){ $delta[$d]=@() }

$keys = @($buckets.Keys) | Sort-Object {[double]$_}
foreach($k in $keys){
  $X += [double]$k

  # 役割ごとに平均
  $valsThis = @{}
  foreach($r in $roleOrder){
    if($buckets[$k].ContainsKey($r)){
      $arr = $buckets[$k][$r].ToArray()
      [double]$avg = ($arr | Measure-Object -Average).Average
      $series[$r] += [double]$avg
      $valsThis[$r] = [double]$avg
    }else{
      $series[$r] += $null
      $valsThis[$r] = $null
    }
  }

  # Δ計算（存在するものだけ）
  $l2  = if($valsThis.ContainsKey('L2')){ $valsThis['L2'] }else{$null}
  $l3  = if($valsThis.ContainsKey('L3')){ $valsThis['L3'] }else{$null}
  $lan = if($valsThis.ContainsKey('RTR_LAN')){ $valsThis['RTR_LAN'] }else{$null}
  $wan = if($valsThis.ContainsKey('RTR_WAN')){ $valsThis['RTR_WAN'] }else{$null}
  $saas= if($valsThis.ContainsKey('SAAS')){ $valsThis['SAAS'] }else{$null}

  function SubOrNull($a,$b){ if($a -ne $null -and $b -ne $null){ return [double]($a-$b) } else { return $null } }

  $delta['DELTA_L3']       += (SubOrNull $l3  $l2)
  $delta['DELTA_RTR_LAN']  += (SubOrNull $lan $l3)
  $delta['DELTA_RTR_WAN']  += (SubOrNull $wan $lan)
  $delta['DELTA_CLOUD']    += (SubOrNull $saas $wan)
}

# ------------- Excel 出力 -------------
[int]$xlXYScatterLines=74; [int]$xlLegendBottom=-4107; [int]$xlCategory=1; [int]$xlValue=2
[int]$msoLineDash=4

$excel=$null; $wb=$null
try{
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = [bool]$Visible
  $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Add()
  while($wb.Worksheets.Count -gt 1){ $wb.Worksheets.Item(1).Delete() }

  # ---- LayerSeries ----
  $ws1 = $wb.Worksheets.Item(1); $ws1.Name = 'LayerSeries'
  $ws1.Cells(1,1).Value2='timestamp'
  # X軸（OADate）を書き込み
  $n=$X.Count
  if($n -lt 1){ throw "No aggregated buckets" }
  $arrX = New-Object 'object[,]' $n,1
  for($i=0;$i -lt $n;$i++){ $arrX[$i,0]=$X[$i] }
  $ws1.Range('A2').Resize($n,1).Value2=$arrX
  $ws1.Range(("A2:A{0}" -f (1+$n))).NumberFormatLocal='yyyy/mm/dd hh:mm'

  # 役割列
  $col=2
  foreach($r in $roleOrder){
    $ws1.Cells(1,$col).Value2=$r
    $vals=$series[$r]
    $arr = New-Object 'object[,]' $n,1
    for($i=0;$i -lt $n;$i++){ $arr[$i,0]=$vals[$i] }
    $ws1.Range($ws1.Cells(2,$col),$ws1.Cells(1+$n,$col)).Value2=$arr
    $col++
  }
  $ws1.Columns.AutoFit() | Out-Null

  # グラフ（レイヤ）
  $ch1 = $ws1.ChartObjects().Add(320,10,900,330)
  $c1 = $ch1.Chart
  $c1.ChartType = $xlXYScatterLines
  $c1.HasTitle = $true
  $c1.ChartTitle.Text = 'Layer RTT (bucket avg)'
  $c1.Legend.Position = $xlLegendBottom
  try{ $c1.SeriesCollection().Delete() }catch{}
  $endRow = 1 + $n
  $sCol = 2
  foreach($r in $roleOrder){
    $s = $c1.SeriesCollection().NewSeries()
    $s.Name = $r
    $s.XValues = $ws1.Range(("A2:A{0}" -f $endRow))
    $s.Values  = $ws1.Range(($ws1.Cells(2,$sCol).Address()+":"+$ws1.Cells($endRow,$sCol).Address()))
    $sCol++
  }
  try{
    $v=$c1.Axes($xlValue);    $v.MinimumScale=0; $v.MaximumScale=300; $v.MajorUnit=50
    $x=$c1.Axes($xlCategory); $x.MajorUnit=(1.0/24.0); $x.TickLabels.NumberFormat='mm/dd hh:mm'
  }catch{}

  # ---- DeltaSeries ----
  $ws2 = $wb.Worksheets.Add()
  $ws2.Name = 'DeltaSeries'
  $ws2.Cells(1,1).Value2='timestamp'
  $ws2.Range('A2').Resize($n,1).Value2=$arrX
  $ws2.Range(("A2:A{0}" -f (1+$n))).NumberFormatLocal='yyyy/mm/dd hh:mm'
  $col=2
  foreach($d in $DeltaNames){
    $ws2.Cells(1,$col).Value2=$d
    $vals=$delta[$d]
    $arr = New-Object 'object[,]' $n,1
    for($i=0;$i -lt $n;$i++){ $arr[$i,0]=$vals[$i] }
    $ws2.Range($ws2.Cells(2,$col),$ws2.Cells(1+$n,$col)).Value2=$arr
    $col++
  }
  $ws2.Columns.AutoFit() | Out-Null

  # グラフ（Δ）
  $ch2 = $ws2.ChartObjects().Add(320,10,900,330)
  $c2 = $ch2.Chart
  $c2.ChartType = $xlXYScatterLines
  $c2.HasTitle = $true
  $c2.ChartTitle.Text = 'Layer Δ (bucket avg)'
  $c2.Legend.Position = $xlLegendBottom
  try{ $c2.SeriesCollection().Delete() }catch{}
  $endRow = 1 + $n
  $sCol = 2
  foreach($d in $DeltaNames){
    $s = $c2.SeriesCollection().NewSeries()
    $s.Name = $d
    $s.XValues = $ws2.Range(("A2:A{0}" -f $endRow))
    $s.Values  = $ws2.Range(($ws2.Cells(2,$sCol).Address()+":"+$ws2.Cells($endRow,$sCol).Address()))
    $sCol++
  }
  try{
    $v=$c2.Axes($xlValue);    $v.MinimumScale=0; $v.MaximumScale=300; $v.MajorUnit=50
    $x=$c2.Axes($xlCategory); $x.MajorUnit=(1.0/24.0); $x.TickLabels.NumberFormat='mm/dd hh:mm'
  }catch{}

  $wb.SaveAs($Output)
  Write-Host "Output: $Output"
}
catch{
  Write-Error (Format-ErrorRecord $_)
  throw
}
finally{
  if($wb){ try{ $wb.Close($false) }catch{}; Release-Com $wb; $wb=$null }
  if($excel){ try{ $excel.Quit() }catch{}; Release-Com $excel; $excel=$null }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
