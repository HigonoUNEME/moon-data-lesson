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
                    # 6th element: axis and title details (title text, legend, value-axis min/max, category label spacing)
                    $ax = [ordered]@{}
                    try { $ax["title"] = $(if ($ch.HasTitle) { [string]$ch.ChartTitle.Text } else { "" }) } catch {}
                    try { $ax["legend"] = [bool]$ch.HasLegend } catch {}
                    try { $va = $ch.Axes(2); $ax["y_min"] = $va.MinimumScale; $ax["y_max"] = $va.MaximumScale; $ax["y_major"] = $va.MajorUnit
                          $ax["y_title"] = $(if ($va.HasTitle) { [string]$va.AxisTitle.Text } else { "" }) } catch {}
                    try { $ca = $ch.Axes(1); $ax["x_label_spacing"] = $ca.TickLabelSpacing
                          $ax["x_title"] = $(if ($ca.HasTitle) { [string]$ca.AxisTitle.Text } else { "" }) } catch {}
                    $o += ,@([string]$co.Name, [string]$co.TopLeftCell.Address(), [string]$co.BottomRightCell.Address(), [int]$ch.ChartType, $series, $ax)
                }
                $result[[string]$s.key] = $o
            }
            "colhidden" {
                # whether a column is hidden (e.g. the helper column M that holds the drop-down list)
                $ws = $wb.Worksheets.Item([string]$s.sheet)
                $result[[string]$s.key] = [bool]$ws.Columns.Item([string]$s.col).Hidden
            }
            "rowinfo" {
                $ws = $wb.Worksheets.Item([string]$s.sheet)
                $o = @()
                foreach ($rr in $s.rows) { $o += ,@([int]$rr, [double]$ws.Rows.Item([int]$rr).RowHeight, [bool]$ws.Rows.Item([int]$rr).Hidden) }
                $result[[string]$s.key] = $o
            }
            "exportpdf" {
                # export one sheet to PDF (for visual checks). Optional: area (print area, e.g. "A1:N40"), landscape, fit (1 = fit to one page)
                $ws = $wb.Worksheets.Item([string]$s.sheet)
                if ($s.area) { $ws.PageSetup.PrintArea = [string]$s.area }
                if ($s.landscape) { $ws.PageSetup.Orientation = 2 }
                if ($s.fit) { $ws.PageSetup.Zoom = $false; $ws.PageSetup.FitToPagesWide = 1; $ws.PageSetup.FitToPagesTall = 1 }
                $ws.ExportAsFixedFormat(0, [string]$s.path)
            }
            "exportchart" {
                # export each chart of a sheet to PNG as Excel draws it: <path>_<n>.png
                $ws = $wb.Worksheets.Item([string]$s.sheet)
                $n = 0
                foreach ($co in $ws.ChartObjects()) {
                    $n++
                    $co.Chart.Export(([string]$s.path + "_" + $n + ".png"), "PNG") | Out-Null
                }
                $result[[string]$s.key] = $n
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
