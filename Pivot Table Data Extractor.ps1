# Define file paths
$excelFilePath = "YOUR PATH TO YOUR EXCEL SPREADSHEET"
$sheetName = "Sheet2" # Replace with your sheet name
$txtFilePath = "PATH TO THE FILE YOU WANT IT TO WRITE THE DATA TO. TXT WORKS BEST"

# Define filter criteria
$filterNAME OF FILTER = "YOURFILTERVALUE"           # Replace with the YOURFILTERNAME filter value
$filterNAME OF FILTER = "YOURFILTERVALUE" # Replace with the YOURFILTERNAME filter value
$filterNAME OF FILTER = "YOURFILTERVALUE"     # Replace with the YOURFILTERNAME filter value

# Define the custom text to prepend and append
$customTextStart = @"
ADD YOUR OWN TEXT OR DELETE THIS TO HAVE A HEADER TO YOUR OUTPUT
"@

$customTextEnd = "ADD YOUR OWN TEXT OR DELETE THIS TO HAVE A FOOTER ON YOUR OUTPUT"

# Create Excel COM object and open the workbook
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelFilePath)
$worksheet = $workbook.Sheets.Item($sheetName)

# Access the pivot table (assuming the first pivot table on the sheet)
$pivotTable = $worksheet.PivotTables().Item(1)

# Apply filters to the pivot table
$pivotTable.PivotFields("YOURFILTERNAME").ClearAllFilters()
$pivotTable.PivotFields("YOURFILTERNAME").CurrentPage = $filterYOURFILTERNAME

$pivotTable.PivotFields("YOURFILTERNAME").ClearAllFilters()
$pivotTable.PivotFields("YOURFILTERNAME").CurrentPage = $filterYOURFILTERNAME

$pivotTable.PivotFields("YOURFILTERNAME").ClearAllFilters()
$pivotTable.PivotFields("YOURFILTERNAME").CurrentPage = $filterYOURFILTERNAME

# Refresh the pivot table to apply the filters
$pivotTable.RefreshTable()

# Get the data from the pivot table after applying filters
$dataRange = $pivotTable.TableRange2
$data = $dataRange.Value2

# Convert the data range into a more manageable format (array of arrays)
$dataArray = @()
# Start processing from row 6
for ($i = 6; $i -le $data.GetLength(0); $i++) {
    $row = @()
    for ($j = 1; $j -le $data.GetLength(1); $j++) {
        $row += $data[$i, $j]
    }
    $dataArray += [PSCustomObject]@{ Row = $row }
}

# Sort the data by the first column (modify sorting logic as needed)
$sortedData = $dataArray | Sort-Object { $_.Row[0] }

# Write the custom text, header, and sorted data to the text file
# First write the custom text at the beginning
$customTextStart | Out-File -FilePath $txtFilePath -Encoding utf8

# Append the header and data
$headerLine = ($dataArray[0].Row -join " ") # Adjust delimiter if needed
Add-Content -Path $txtFilePath -Value $headerLine

$sortedData | ForEach-Object {
    $line = ($_.Row -join " ") # Adjust delimiter if needed
    Add-Content -Path $txtFilePath -Value $line
}

# Append custom text at the end
Add-Content -Path $txtFilePath -Value $customTextEnd

# Clean up
$workbook.Close($false)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Output "Data successfully exported, filtered, and sorted. Check the file: $txtFilePath"