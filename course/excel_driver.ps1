# course/excel_driver.ps1
# Excel（COM）で xlsx を開き、JSON のジョブに従って「値を入れる・再計算・読む・保存」を行う検証用ドライバ。
# check_course_xlsx.py / verify_excel.py から呼ばれる。このファイルは ASCII のみ（日本語は JSON 側）。
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File excel_driver.ps1 -Job job.json
#
# job.json: { "file": "...xlsx", "out": "result.json", "saveAs": "optional.xlsx",
#             "steps": [ {op:"set",sheet,addr,value} | {op:"calc"} | {op:"get",sheet,addr,key,text}
#                       | {op:"dv",sheet,addr,key} | {op:"info",key} | {op:"save"} ... ] }
# 元ファイルは上書きしない（呼び出し側でコピーを渡す）。保存は saveAs か op:"saveas" のみ。
param([Parameter(Mandatory=$true)][string]$JobFile)

$ErrorActionPreference = "Stop"
$job = Get-Content -Raw -Encoding UTF8 $JobFile | ConvertFrom-Json
$result = [ordered]@{}
$sw = [System.Diagnostics.Stopwatch]::StartNew()

$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false
$xl.Visible = $false
$xl.ScreenUpdating = $false
try {
    $wb = $xl.Workbooks.Open($job.file, 0, $false)
    $result["open_ms"] = $sw.ElapsedMilliseconds
    $result["excel_version"] = $xl.Version
    $result["errors"] = @()
    foreach ($s in $job.steps) {
      try {
        switch ($s.op) {
            "set" {
                $ws = $wb.Worksheets.Item([string]$s.sheet)
                $rng = $ws.Range([string]$s.addr)
                if ($null -eq $s.value) { $rng.ClearContents() | Out-Null }
                elseif ($s.value -is [string]) { $rng.Value2 = [string]$s.value }
                else { $rng.Value2 = [double]$s.value }
            }
            "calc" {
                $t = $sw.ElapsedMilliseconds
                $xl.CalculateFull()
                $result["calc_ms_" + $s.key] = $sw.ElapsedMilliseconds - $t
            }
            "get" {
                $ws = $wb.Worksheets.Item([string]$s.sheet)
                $rng = $ws.Range([string]$s.addr)
                $rows = @()
                for ($r = 1; $r -le $rng.Rows.Count; $r++) {
                    $row = @()
                    for ($c = 1; $c -le $rng.Columns.Count; $c++) {
                        $cell = $rng.Cells.Item($r, $c)
                        if ($s.text) { $row += [string]$cell.Text } else { $row += $cell.Value2 }
                    }
                    $rows += ,$row
                }
                $result[[string]$s.key] = $rows
            }
            "dv" {
                $ws = $wb.Worksheets.Item([string]$s.sheet)
                $cell = $ws.Range([string]$s.addr)
                $o = [ordered]@{}
                try {
                    $v = $cell.Validation
                    $o["type"] = $v.Type
                    $o["formula1"] = [string]$v.Formula1
                    $o["formula2"] = [string]$v.Formula2
                    $o["operator"] = $v.Operator
                    $o["alert_style"] = $v.AlertStyle
                    $o["show_error"] = $v.ShowError
                    $o["error_message"] = [string]$v.ErrorMessage
                    $o["in_cell_dropdown"] = $v.InCellDropdown
                    $o["current_value_valid"] = $cell.Validation.Value
                } catch { $o["error"] = $_.Exception.Message }
                $result[[string]$s.key] = $o
            }
            "testvalid" {
                # セルに値を入れて、Validation.Value（入力規則を満たすか）を返す。値は元に戻す
                $ws = $wb.Worksheets.Item([string]$s.sheet)
                $cell = $ws.Range([string]$s.addr)
                $old = $cell.Value2
                $stage = "set"
                $ok = $false
                try {
                    if ($s.value -is [decimal]) { $cell.Value2 = [double]$s.value }
                    elseif ($s.value -is [string]) { $cell.Value2 = [string]$s.value }
                    else { $cell.Value2 = [double]$s.value }
                    $stage = "validate"
                    try { $ok = [bool]($cell.Validation.Value) } catch { $ok = $false }
                } finally {
                    if ($null -eq $old) { $cell.ClearContents() | Out-Null }
                    elseif ($old -is [string]) { $cell.Value2 = [string]$old }
                    else { $cell.Value2 = [double]$old }
                }
                $result[[string]$s.key] = $ok
            }
            "info" {
                $o = [ordered]@{}
                $o["sheets"] = @()
                foreach ($w in $wb.Worksheets) {
                    $o["sheets"] += ,@([string]$w.Name, [int]$w.Visible, [int]$w.ChartObjects().Count)
                }
                $o["calc_mode"] = $xl.Calculation
                $o["names"] = @()
                foreach ($n in $wb.Names) { $o["names"] += ,@([string]$n.Name, [string]$n.RefersTo) }
                $result[[string]$s.key] = $o
            }
            "chart" {
                $ws = $wb.Worksheets.Item([string]$s.sheet)
                $o = @()
                foreach ($co in $ws.ChartObjects()) {
                    $ch = $co.Chart
                    $series = @()
                    foreach ($sr in $ch.SeriesCollection()) { $series += ,@([string]$sr.Name, [string]$sr.Formula) }
                    $o += ,@([string]$co.Name, [string]$co.TopLeftCell.Address(), [string]$co.BottomRightCell.Address(), [int]$ch.ChartType, $series)
                }
                $result[[string]$s.key] = $o
            }
            "rowinfo" {
                $ws = $wb.Worksheets.Item([string]$s.sheet)
                $o = @()
                foreach ($rr in $s.rows) { $o += ,@([int]$rr, [double]$ws.Rows.Item([int]$rr).RowHeight, [bool]$ws.Rows.Item([int]$rr).Hidden) }
                $result[[string]$s.key] = $o
            }
            "exportpdf" {
                # 指定シートだけ PDF に出す（見た目の確認用）
                $ws = $wb.Worksheets.Item([string]$s.sheet)
                $ws.ExportAsFixedFormat(0, [string]$s.path)
            }
            "saveas" {
                $t = $sw.ElapsedMilliseconds
                $wb.SaveAs([string]$s.path, 51)
                $result["saveas_ms"] = $sw.ElapsedMilliseconds - $t
            }
            default { throw "unknown op: $($s.op)" }
        }
      } catch {
        $result["errors"] += ,@([string]$s.op, [string]$s.key, [string]$s.addr, [string]$_.Exception.Message)
      }
    }
    $wb.Close($false)
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
$result["total_ms"] = $sw.ElapsedMilliseconds
$json = $result | ConvertTo-Json -Depth 8
[System.IO.File]::WriteAllText($job.out, $json, (New-Object System.Text.UTF8Encoding($false)))
