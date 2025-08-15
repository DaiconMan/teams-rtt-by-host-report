# Teams RTT by Host Report

`teams_net_quality.csv` から **ターゲット(host)ごとにRTTグラフ**を作成する PowerShell スクリプト。  
- ピボット非依存（PowerPointに貼っても連動しない普通のグラフ）
- 縦軸 0–300ms 固定、横軸 1時間刻み
- 100ms 赤の破線の閾値線
- ホストごとに線色固定

## 使い方
```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\scripts\windows\Generate-TeamsNet-RTT-ByHost.ps1 `
  -Output "$HOME\Desktop\TeamsNet-RTT-ByHost.xlsx" `
  -TargetsFile ".\samples\target.txt"