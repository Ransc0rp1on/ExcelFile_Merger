param(
    [Parameter(Mandatory=$true)][string[]]$Files,
    [Parameter(Mandatory=$true)][string]$OutputFile
)

function Import-ExcelFile {
    param($Path)
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open((Resolve-Path $Path).Path)
        $worksheet = $workbook.Sheets.Item(1)
        $range = $worksheet.UsedRange
        
        # Get headers
        $headers = @()
        for ($col = 1; $col -le $range.Columns.Count; $col++) {
            $headers += $range.Cells.Item(1, $col).Value2
        }
        
        # Get data rows
        $data = @()
        for ($row = 2; $row -le $range.Rows.Count; $row++) {
            $rowData = [ordered]@{}
            for ($col = 1; $col -le $headers.Count; $col++) {
                $value = $range.Cells.Item($row, $col).Value2
                $rowData[$headers[$col-1]] = if ($null -eq $value) { $null } else { $value.ToString() }
            }
            $data += [PSCustomObject]$rowData
        }
        return $data
    }
    finally {
        # Clean up COM objects safely
        if ($workbook) { $workbook.Close($false) }
        if ($excel) { $excel.Quit() }
        
        # Release COM objects in reverse order
        if ($range) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($range) | Out-Null }
        if ($worksheet) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null }
        if ($workbook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null }
        if ($excel) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
        
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Import-DataFile {
    param($Path)
    $extension = [System.IO.Path]::GetExtension($Path).ToLower()
    
    switch ($extension) {
        '.csv' { return Import-Csv -Path $Path -Delimiter "," -Encoding UTF8 }
        '.xlsx' { return Import-ExcelFile -Path $Path }
        default { throw "Unsupported file format: $extension" }
    }
}

# Main script
try {
    # Collect all unique columns
    $allColumns = [System.Collections.Generic.HashSet[string]]::new()
    
    # First pass: Collect columns from all files
    foreach ($file in $Files) {
        Write-Host "Scanning columns in: $file" -ForegroundColor Cyan
        $data = Import-DataFile -Path $file
        foreach ($row in $data) {
            $row.PSObject.Properties.Name | ForEach-Object { 
                [void]$allColumns.Add($_) 
            }
        }
    }

    # Prepare final data
    $allData = [System.Collections.ArrayList]::new()
    
    # Second pass: Process files with aligned columns
    foreach ($file in $Files) {
        Write-Host "Processing file: $file" -ForegroundColor Green
        $data = Import-DataFile -Path $file
        foreach ($row in $data) {
            $newRow = [ordered]@{}
            foreach ($col in $allColumns) {
                $newRow[$col] = if ($row.PSObject.Properties[$col]) { $row.$col } else { $null }
            }
            [void]$allData.Add([PSCustomObject]$newRow)
        }
    }

    # Export results
    $allData | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8 -Force
    Write-Host "Successfully merged $($Files.Count) files! Output saved to $OutputFile" -ForegroundColor Green
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    exit 1
}
