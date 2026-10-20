# Điền một file mẫu Excel bằng chính Excel (giữ nguyên logo, định dạng, công thức), lưu bản mới và xuất PDF.
# Tham số: đường dẫn file JSON
#   {template, output, pdf, pdf_sheet,
#    inserts: [[sheet, dòng, số dòng chèn], ...],     -> chèn dòng (giữ định dạng dòng phía trên)
#    cells:   [[sheet, ô, giá trị], ...],
#    hide_zero: [[sheet, dòng đầu, dòng cuối, cột số lượng, cột cuối bảng], ...]}  -> ẩn dòng hàng không đặt
param([string]$JobFile)
$ErrorActionPreference = 'Stop'
$job = Get-Content -LiteralPath $JobFile -Raw -Encoding UTF8 | ConvertFrom-Json
Copy-Item -LiteralPath $job.template -Destination $job.output -Force
$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
$excel.Visible = $false
try {
    $book = $excel.Workbooks.Open($job.output)
    foreach ($ins in @($job.inserts)) {
        if ($null -eq $ins -or [int]$ins[2] -le 0) { continue }
        $first = [int]$ins[1]; $last = $first + [int]$ins[2] - 1
        [void]$book.Worksheets.Item([string]$ins[0]).Range("$($first):$($last)").EntireRow.Insert()
    }
    foreach ($c in $job.cells) {
        $raw = $c[2]
        # giá trị từ ConvertFrom-Json bị bọc PSObject / kiểu Decimal: Excel COM báo "Specified cast is not valid"
        # -> ép rõ về [string] hoặc [double] trước khi ghi
        if ($null -eq $raw) { $v = [string]'' }
        elseif ($raw -is [string]) { $v = [string]$raw }
        else { $v = [double]$raw }
        $range = $book.Worksheets.Item([string]$c[0]).Range([string]$c[1])
        # không dùng "$range.Value2 = $v": PowerShell nhớ kiểu của lần gán đầu (chữ) rồi báo lỗi khi gán số
        [void]$range.GetType().InvokeMember('Value2', [Reflection.BindingFlags]::SetProperty, $null, $range, @($v))
    }
    $excel.CalculateFull()
    # ẩn các mặt hàng không đặt: [sheet, dòng đầu, dòng cuối, cột số lượng]
    # (không dùng AutoFilter có điều kiện: Excel tự nới vùng lọc xuống các dòng tổng và ẩn luôn dòng "Công nợ 0%")
    foreach ($h in @($job.hide_zero)) {
        if ($null -eq $h) { continue }
        $ws = $book.Worksheets.Item([string]$h[0])
        if ($ws.AutoFilterMode) { $ws.AutoFilterMode = $false }  # bộ lọc cũ trong mẫu ẩn sai dòng
        $first = [int]$h[1]; $last = [int]$h[2]; $col = [string]$h[3]
        $ws.Range("$($first):$($last)").EntireRow.Hidden = $false
        [void]$ws.Range("A$($first - 1):$([string]$h[4])$($last)").AutoFilter()  # bật lại nút lọc (không điều kiện) trên cả bảng
        for ($r = $first; $r -le $last; $r++) {
            $q = $ws.Range("$col$r").Value2
            if ($null -eq $q -or ($q -is [double] -and [math]::Abs($q) -lt 1e-9)) { $ws.Rows($r).Hidden = $true }
        }
    }
    $book.Save()
    if ($job.pdf) {
        $book.Worksheets.Item([string]$job.pdf_sheet).ExportAsFixedFormat(0, $job.pdf)
    }
    $book.Close($false)
}
finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
