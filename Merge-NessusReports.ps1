param(
    [Parameter(Mandatory=$true)][string]$File1,
    [Parameter(Mandatory=$true)][string]$File2,
    [Parameter(Mandatory=$true)][string]$OutputFile
)

function Import-DataFile {
    param($Path)
    $extension = [System.IO.Path]::GetExtension($Path).ToLower()
    
    switch ($extension) {
        '.csv' {
            return Import-Csv -Path $Path -Delimiter "," -Encoding UTF8
        }
        '.xlsx' {
            try {
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $false
                $excel.DisplayAlerts = $false
                $workbook = $excel.Workbooks.Open($Path)
                $worksheet = $workbook.Sheets.Item(1)
                $range = $worksheet.UsedRange
                $data = $range.Value2 | ConvertFrom-Csv -Delimiter "`t"
                return $data
            }
            finally {
                if ($workbook) { $workbook.Close($false) }
                if ($excel) { $excel.Quit() }
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($range) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            }
        }
        default { throw "Unsupported file format: $extension" }
    }
}

function Export-DataFile {
    param($Data, $Path)
    $extension = [System.IO.Path]::GetExtension($Path).ToLower()
    
    switch ($extension) {
        '.csv' {
            $Data | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Delimiter ","
        }
        '.xlsx' {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $workbook = $excel.Workbooks.Add()
            $worksheet = $workbook.Worksheets.Item(1)
            
            # Add headers
            $column = 1
            $Data[0].PSObject.Properties.Name | ForEach-Object {
                $worksheet.Cells.Item(1, $column) = $_
                $column++
            }
            
            # Add data
            $row = 2
            foreach ($item in $Data) {
                $col = 1
                $item.PSObject.Properties.Value | ForEach-Object {
                    $worksheet.Cells.Item($row, $col) = $_
                    $col++
                }
                $row++
            }
            
            $worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null
            $workbook.SaveAs($Path, 51)  # 51 = xlOpenXMLWorkbook
            $workbook.Close($false)
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
        default { throw "Unsupported output format: $extension" }
    }
}

# Main script
try {
    # Import data
    $data1 = Import-DataFile -Path $File1
    $data2 = Import-DataFile -Path $File2
    
    # Get all column names
    $columns1 = $data1[0].PSObject.Properties.Name
    $columns2 = $data2[0].PSObject.Properties.Name
    $allColumns = ($columns1 + $columns2) | Select-Object -Unique
    
    # Create combined data with all columns
    $combinedData = [System.Collections.Generic.List[object]]::new()
    
    # Process first dataset
    foreach ($row in $data1) {
        $newRow = [ordered]@{}
        foreach ($col in $allColumns) {
            $newRow[$col] = if ($row.PSObject.Properties[$col]) { $row.$col } else { $null }
        }
        $combinedData.Add([PSCustomObject]$newRow)
    }
    
    # Process second dataset
    foreach ($row in $data2) {
        $newRow = [ordered]@{}
        foreach ($col in $allColumns) {
            $newRow[$col] = if ($row.PSObject.Properties[$col]) { $row.$col } else { $null }
        }
        $combinedData.Add([PSCustomObject]$newRow)
    }
    
    # Export combined data
    Export-DataFile -Data $combinedData -Path $OutputFile
    Write-Host "Successfully merged files! Output saved to $OutputFile" -ForegroundColor Green
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    exit 1
}