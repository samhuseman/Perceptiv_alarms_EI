# Specify the folder path
$folder = 'data'

# Find the Excel file in the folder
$Values = Get-ChildItem $folder -Filter *.xlsx | Select -First 1

<# # Check if an Excel file was found
if ($file) {
    # Open the Excel file with Excel
    $excel = New-Object -ComObject Excel.Application
    $workbook = $excel.Workbooks.Open($file.FullName)
    $excel.Visible = $true
}
else {
    Write-Warning 'No Excel file was found in the specified folder.'
} #>

# Open the source worksheet
$sourceWorksheet = $Values.Worksheets.Item("Sheet1")

# Copy the contents of the source worksheet
$sourceWorksheet.Range("A1:B10").Copy()

# Open the destination workbook
$destinationWorkbook = $excel.Workbooks.Open("C:\path\to\destination\workbook.xlsx")


