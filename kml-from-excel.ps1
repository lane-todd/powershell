# Load required .NET assemblies
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.Windows.Forms

function Prompt-UserForFile {
    param (
        [string]$Filter,
        [string]$InitialDirectory
    )
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = $Filter
    $dialog.InitialDirectory = $InitialDirectory
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    } else {
        Write-Host "No file selected. Exiting script."
        exit
    }
}

function Prompt-UserToSaveFile {
    param (
        [string]$Filter,
        [string]$InitialDirectory,
        [string]$DefaultFileName
    )
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = $Filter
    $dialog.InitialDirectory = $InitialDirectory
    $dialog.FileName = $DefaultFileName
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    } else {
        Write-Host "Save operation canceled. Exiting script."
        exit
    }
}

function Cleanup-Excel {
    param (
        $worksheet, $workbook, $excel
    )
    try {
        $workbook.Close($False)
        $excel.Quit()
    } catch {
        Write-Warning "Error while closing Excel: $_"
    }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# Prompt user to select an Excel file
$excelFilePath = Prompt-UserForFile -Filter "Excel Files (*.xlsx)|*.xlsx" -InitialDirectory "C:\"

# Open Excel and process data
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Open($excelFilePath)
$worksheet = $workbook.Sheets.Item(1)
$usedRange = $worksheet.UsedRange

# Build KML content
$kmlStart = @'
<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2" xmlns:gx="http://www.google.com/kml/ext/2.2" xmlns:kml="http://www.opengis.net/kml/2.2" xmlns:atom="http://www.w3.org/2005/Atom">
<Document>
'@

# Dynamically generate schema
$schema = "<Schema name='schema0' id='schema0'>" + "`n"
for ($i = 1; $i -le $usedRange.Columns.Count; $i++) {
    $columnHeader = $worksheet.Cells.Item(1, $i).Value2
    if ($columnHeader) {
        $schema += "<SimpleField type='string' name='$columnHeader'></SimpleField>" + "`n"
    }
}
$schema += "</Schema>" + "`n"

# Add style for map pins
$kmlStyle = @"
<Style id="pdfmaps_style_red">
  <IconStyle>
    <Icon>
      <href>http://download.avenza.com/images/pdfmaps_icons/v2/pin-red-inground.png</href>
    </Icon>
  </IconStyle>
</Style>
"@

# Generate placemarks
$columnHeaders = @()
for ($i = 1; $i -le $usedRange.Columns.Count; $i++) {
    $columnHeader = $worksheet.Cells.Item(1, $i).Value2
    if ($columnHeader) { $columnHeaders += $columnHeader }
}
$allPlacemarks = ""
for ($row = 2; $row -le $usedRange.Rows.Count; $row++) {
    $placemark = "<Placemark>`n"
    foreach ($columnHeader in $columnHeaders) {
        $value = $worksheet.Cells.Item($row, $columnIndexMap[$columnHeader]).Value2
        $placemark += "<SimpleData name='$columnHeader'>$([System.Web.HttpUtility]::HtmlEncode($value))</SimpleData>`n"
    }
    $placemark += "</Placemark>`n"
    $allPlacemarks += $placemark
}

# Complete KML
$kmlContent = $kmlStart + $schema + $kmlStyle + $allPlacemarks + "</Document></kml>"

# Save the KML file
$kmlFilePath = Prompt-UserToSaveFile -Filter "KML Files (*.kml)|*.kml" -InitialDirectory "C:\" -DefaultFileName "pins-from-EXCEL.kml"
$kmlContent | Set-Content -Path $kmlFilePath -Encoding UTF8 -Force

Write-Host "KML file generated successfully: $kmlFilePath"

# Cleanup
Cleanup-Excel -worksheet $worksheet -workbook $workbook -excel $excel



##### TODO add back in SCHEMADATA elements and simpleField/SimpleData to display the column elements in a table on the pin...
