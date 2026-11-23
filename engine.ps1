param([string]$file, [string]$destination)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
function Write-ReportCell($cell, $value) {
    if ($null -eq $value -or $value -is [string] -and $value.Length -eq 0) {
        $cell.MergeArea.ClearContents() | Out-Null
    } elseif ($value -is [string]) {
        $cell.Value2 = [string]$value
    } else {
        $literal = [Convert]::ToDouble($value).ToString('R', [Globalization.CultureInfo]::InvariantCulture)
        $cell.Formula = '=' + $literal
    }
}
$excel = $null
try {
    Write-Output "Starting Microsoft Excel..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $job = [Guid]::NewGuid().ToString('N')
    $temp = Join-Path ([IO.Path]::GetTempPath()) $job
    $temp = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) ('SenuraHiring/work/'+$job)
    $stage = Join-Path $temp 'bills'
    [IO.Directory]::CreateDirectory($stage) | Out-Null
    $local = Join-Path $temp 'source.xlsx'
    [IO.File]::Copy($file,$local)
    $source = $excel.Workbooks.Open($local,0,$true)
    $setup = $source.Worksheets.Item('SETUP')
    $summary = $source.Worksheets.Item('SUMMARY')
    $excel.Calculation = -4105
    Write-Output "Recalculating monthly report..."
    $excel.CalculateFullRebuild()
    $config = $setup.UsedRange.Value2
    $data = $summary.UsedRange.Value2
    $site = [string]$config.GetValue(3,2)
    $month = '{0:D4}-{1:D2}' -f [int]$config.GetValue(3,5),[int]$config.GetValue(3,8)
    if ($site -ne [string]$data.GetValue(3,2) -or $config.GetValue(3,5) -ne $data.GetValue(3,4) -or $config.GetValue(3,8) -ne $data.GetValue(3,6)) { throw 'Setup and summary identify different reporting periods or sites.' }
    if ([string]$data.GetValue(7,27) -notmatch 'FINAL PAYABLE') { throw 'Unsupported SUMMARY layout.' }
    $active = @{}
    for ($r=6; $r -le $config.GetLength(0); $r++) {
        if ($config.GetValue($r,1) -ne 'YES') { continue }
        $id = [string]$config.GetValue($r,4)
        if (!$id) { $id = [string]$config.GetValue($r,2) }
        if (!$id -or $active.ContainsKey($id)) { throw "Missing or duplicate active vehicle ID at SETUP row $r" }
        $active[$id]=$r
    }
    $seen=@{}
    $line=2
    $reviews=0
    for ($r=8;$r -le $data.GetLength(0);$r++) {
        if ($data.GetValue($r,1) -eq 'TOTAL') { break }
        $id=[string]$data.GetValue($r,3)
        if (!$id) { $id=[string]$data.GetValue($r,1) }
        if (!$id) { continue }
        if (!$active.ContainsKey($id) -or $seen.ContainsKey($id)) { throw "Unexpected vehicle $id in SUMMARY. Check the source workbook." }
        $seen[$id]=$true
        $line++
        $status=[string]$data.GetValue($r,28)
        $ready=($status -eq 'READY' -and $data.GetValue($r,24) -is [double] -and $data.GetValue($r,24) -gt 0 -and $data.GetValue($r,27) -is [double])
        if (!$ready) { $status="REVIEW - $status"; $reviews++ }
        Write-Output "Preparing bill for $id..."
        $output=$excel.Workbooks.Open((Join-Path $PSScriptRoot 'template.xlsx'),0,$true)
        $bill=$output.Worksheets.Item('Sheet1')
        $log=$output.Worksheets.Item('Sheet2')
        $first=[datetime]::new([int]$config.GetValue(3,5),[int]$config.GetValue(3,8),1)
        $days=[datetime]::DaysInMonth($first.Year,$first.Month)
        $parts=([string]$data.GetValue($r,2)) -split ' / ',2
        $setupRow=$active[$id]
        $bill.Range('E4').Value2=$id
        $bill.Range('E5').Value2=$site
        $bill.Range('E6').Value2=$parts[0]
        $bill.Range('E7').Value2=[string](@($parts)[1])
        $bill.Range('K5').Value2=[string]$data.GetValue($r,1)
        $bill.Range('K6').Value2="$id/$month"
        $bill.Range('K7').Value2=(Get-Date -Format yyyy.MM.dd)
        $bill.Range('E11').MergeArea.ClearContents()
        $bill.Range('K11').Value2=[string]$config.GetValue($setupRow,6)
        $bill.Range('I13').Value2="Quantity ($($data.GetValue($r,5)))"
        $bill.Range('B14').Value2='A. Billable days / hours'
        Write-ReportCell ($bill.Range('I14')) ($data.GetValue($r,23))
        Write-ReportCell ($bill.Range('I15')) ($data.GetValue($r,7))
        Write-ReportCell ($bill.Range('K14')) ($data.GetValue($r,24))
        Write-ReportCell ($bill.Range('I25')) ($data.GetValue($r,13))
        Write-ReportCell ($bill.Range('K25')) ($config.GetValue(3,11))
        Write-ReportCell ($bill.Range('L25')) ($data.GetValue($r,26))
        $bill.Range('L27').Formula='=L25'
        Write-ReportCell ($bill.Range('L14')) ($data.GetValue($r,25))
        if (!$ready) {
            $bill.Range('E8').Value2=$status
            foreach ($addr in 'L14','L19','L25','L27','L29','L32','L33') { $bill.Range($addr).ClearContents() }
        }
        foreach ($addr in 'A37','H37','C41') { $bill.Range($addr).MergeArea.ClearContents() }
        $log.Range('A9:O44').ClearContents()
        $log.Range('A1').Value2="$site | $id"
        $log.Range('E2').Value2=[string]$data.GetValue($r,2)
        $log.Range('E3').Value2=$id
        $ownerName = 'SENURA CIVIL ENGINEERING (PVT) LTD'
        $bill.Range('E9').Value2=$ownerName
        $bill.Range('E9').ShrinkToFit=$true
        $log.Range('E4').Value2=$ownerName
        $log.Range('E5').Value2=[string]$data.GetValue($r,1)
        $log.Range('E6').Value2=$first.ToString('yyyy.MM.dd')+' - '+$first.AddDays($days-1).ToString('yyyy.MM.dd')
        $log.Range('H7').Value2="Run ($($data.GetValue($r,4)))"
        foreach ($addr in 'A50','D50','H50') { $log.Range($addr).MergeArea.ClearContents() }
        for ($day=1;$day -le $days;$day++) {
            $daySheet=$source.Worksheets.Item(('{0:D2}' -f $day))
            $daily=$daySheet.UsedRange.Value2
            $destRow=$day+8
            $log.Cells.Item($destRow,1)=$first.AddDays($day-1).ToString('yyyy.MM.dd')
            for ($j=5;$j -le $daily.GetLength(0);$j++) {
                if ($daily.GetValue($j,3) -ne $id -and !(!$daily.GetValue($j,3) -and $daily.GetValue($j,1) -eq $id)) { continue }
                foreach ($pair in @('2,5','3,7','4,11','6,4','7,6','8,9','9,12','10,8')) {
                    $destCol,$srcCol=$pair.Split(',')
                    $v=$daily.GetValue($j,[int]$srcCol)
                    if ($null -ne $v) { Write-ReportCell ($log.Cells.Item($destRow,[int]$destCol)) $v }
                }
                if ($daily.GetValue($j,8) -in 'WORK','WORKING','WORKED') { Write-ReportCell ($log.Cells.Item($destRow,5)) 1 }
                break
            }
        }
        $log.Range('G46').Formula='=IFERROR(H45/I45,"")'
        $rightEdge = $bill.Range('A1:L43').Borders.Item(10)
        $rightEdge.LineStyle = 1
        $rightEdge.Weight = 2
        $rightEdge.Color = 0
        $bill.Activate()
        $safeId=$id -replace '[^a-zA-Z0-9_-]','_'
        $name=('{0:D3}_{1}_{2}.xlsx' -f ($line-2),$safeId,$month)
        $output.SaveAs((Join-Path $stage $name),51)
        $output.Close($false)
    }
    if ($seen.Count -ne $active.Count -or !$seen.Count) { throw 'Active fleet and generated bill count differ or the report has no active vehicles.' }
    $label=(($site -replace '[^a-zA-Z0-9 _-]','_')+' '+$month+' Bills '+$job.Substring(0,8))
    $homeBills=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'SenuraHiring/Bills'
    [IO.Directory]::CreateDirectory($homeBills) | Out-Null
    $archive=Join-Path $homeBills $label
    [IO.Directory]::Move($stage,$archive)
    $result=$archive
    try {
        $folder=Join-Path $destination $label
        if ($folder -ne $archive) {
            [IO.Directory]::CreateDirectory($folder) | Out-Null
            foreach ($item in [IO.Directory]::GetFiles($archive)) {
                [IO.File]::Copy($item,(Join-Path $folder ([IO.Path]::GetFileName($item))))
            }
            $result=$folder
        }
        $message="Saved $($seen.Count) bills in $result."
    } catch {
        $message="Windows could not copy all bills to the selected folder. All $($seen.Count) bills are saved in $archive."
    }
    Write-Output "$message Require review: $reviews."
    Start-Process explorer.exe -ArgumentList ('"'+$result+'"') -ErrorAction SilentlyContinue
} catch {
    [Console]::Error.WriteLine(("Line {0}: {1}" -f $_.InvocationInfo.ScriptLineNumber,$_.Exception.Message))
    exit 1
} finally {
    if ($excel) { $excel.Quit() }
    foreach ($dir in @($temp,$pending)) {
        if ($dir -and [IO.Directory]::Exists($dir)) { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}
