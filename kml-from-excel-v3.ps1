# Load required assemblies
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.Windows.Forms

# Define a class for handling Excel data
class ExcelData {
    [string[]]$Headers
    [array]$Rows

    ExcelData() {
        $this.Headers = @()
        $this.Rows = @()
    }

    [void]LoadData($worksheet) {
        $usedRange = $worksheet.UsedRange
        if ($usedRange.Rows.Count -lt 2) {
            Write-Host "Excel file does not contain enough data. Exiting script."
            exit
        }

        # Bulk load data
        $data = $usedRange.Value2
        $this.Headers = @($data[0]) # First row as headers

        # Extract rows as objects
        for ($i = 1; $i -lt $data.Count; $i++) { # Remaining rows as data
            $rowData = @{}
            for ($j = 0; $j -lt $this.Headers.Count; $j++) {
                $header = $this.Headers[$j]
                $rowData[$header] = $data[$i][$j]
            }
            $this.Rows += $rowData
        }
    }
}

# Define a class for generating KML
class KmlGenerator {
    [string]$Schema
    [string]$Placemarks

    KmlGenerator() {
        $this.Schema = ""
        $this.Placemarks = ""
    }

    [void]GenerateSchema($headers) {
        $this.Schema = "<Schema name='schema0' id='schema0'>`n"
        foreach ($header in $headers) {
            $this.Schema += "<SimpleField type='string' name='$header'></SimpleField>`n"
        }
        $this.Schema += "</Schema>`n"
    }

    [string]GenerateStyle($color) {
        $color = $color.ToLower()
        if ($color -notin @('red', 'yellow', 'green')) {
            $color = 'red' # Default color if invalid
        }
        return @"
<Style id="pdfmaps_style_$color">
  <IconStyle>
    <Icon>
      <href>http://download.avenza.com/images/pdfmaps_icons/v2/pin-$color-inground.png</href>
    </Icon>
  </IconStyle>
</Style>
"@
    }

    [void]GeneratePlacemarks($rows, $headers) {
        # Define columns to ignore in the <SchemaData>
        $ignoredColumns = @("PINCOLOR")

        if (-not $headers.Contains("latitude") -or -not $headers.Contains("longitude")) {
            Write-Host "Missing required 'latitude' or 'longitude' column in the spreadsheet. Exiting script."
            exit
        }

        foreach ($row in $rows) {
            if (-not $row["latitude"] -or -not $row["longitude"]) {
                Write-Warning "Row missing 'latitude' or 'longitude' values. Skipping."
                continue
            }

            # Determine the placemark name
            $name = ""
            if ($headers.Contains("STREETNO") -and $headers.Contains("STREETNAME")) {
                $streetNo = $row["STREETNO"]
                $streetName = $row["STREETNAME"]
                if ($streetNo -and $streetName) {
                    $name = "$streetNo $streetName"
                } elseif ($streetName) {
                    $name = "$streetName"
                } elseif ($streetNo) {
                    $name = "$streetNo"
                }
            }

            # Fallback to PARCELNO if name is still empty
            if (-not $name -and $headers.Contains("PARCELNO")) {
                $name = $row["PARCELNO"]
            }

            # Final fallback if PARCELNO is also empty
            if (-not $name) {
                $name = "Unnamed"
            }

            $color = if ($row.ContainsKey("PINCOLOR")) { $row["PINCOLOR"] } else { "red" }
            $styleUrl = "pdfmaps_style_$($color.ToLower())"

            # Start building the placemark
            $placemark = @"
<Placemark>
  <name>$([System.Web.HttpUtility]::HtmlEncode($name))</name>
  <styleUrl>#$styleUrl</styleUrl>
  <ExtendedData>
    <SchemaData schemaUrl="#schema0">
"@

            # Add all <SimpleData> elements for SchemaData, excluding ignored columns
            foreach ($key in $row.Keys) {
                if ($key -notin $ignoredColumns) {
                    $value = [System.Web.HttpUtility]::HtmlEncode($row[$key])
                    $placemark += "      <SimpleData name='$key'>$value</SimpleData>`n"
                }
            }

            $placemark += @"
    </SchemaData>
  </ExtendedData>
  <Point>
    <coordinates>$($row["longitude"]),$($row["latitude"])</coordinates>
  </Point>
</Placemark>
"@

            # Append the placemark to the collection
            $this.Placemarks += $placemark
        }
    }

    [string]BuildKml($styles, $metadata) {
        return @"
<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2" xmlns:gx="http://www.google.com/kml/ext/2.2" xmlns:kml="http://www.opengis.net/kml/2.2" xmlns:atom="http://www.w3.org/2005/Atom">
<Document>
<name>Generated KML</name>
$metadata
$this.Schema
$styles
$this.Placemarks
</Document>
</kml>
"@
    }
}

# Utility functions
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

# Main script
$excelFilePath = Prompt-UserForFile -Filter "Excel Files (*.xlsx)|*.xlsx" -InitialDirectory "C:\"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Open($excelFilePath)
$worksheet = $workbook.Sheets.Item(1)

# Load Excel data
$excelData = [ExcelData]::new()
$excelData.LoadData($worksheet)

# Generate KML
$kmlGenerator = [KmlGenerator]::new()
$kmlGenerator.GenerateSchema($excelData.Headers)

# Generate styles for all unique colors
$uniqueColors = ($excelData.Rows | Where-Object { $_.ContainsKey("PINCOLOR") } | ForEach-Object { $_["PINCOLOR"].ToLower() }) | Sort-Object -Unique
$styles = ""
foreach ($color in $uniqueColors) {
    $styles += $kmlGenerator.GenerateStyle($color)
}

# Generate placemarks
$kmlGenerator.GeneratePlacemarks($excelData.Rows, $excelData.Headers)

# Add metadata
$metadata = @"
<description>KML generated from file: $excelFilePath</description>
<description>Generated on: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</description>
"@

# Build KML content
$kmlContent = $kmlGenerator.BuildKml($styles, $metadata)

# Save KML file
$kmlFilePath = Prompt-UserToSaveFile -Filter "KML Files (*.kml)|*.kml" -InitialDirectory "C:\" -DefaultFileName "pins-from-EXCEL.kml"
$kmlContent | Set-Content -Path $kmlFilePath -Encoding UTF8 -Force

Write-Host "KML file generated successfully: $kmlFilePath"

# Cleanup
Cleanup-Excel -worksheet $worksheet -workbook $workbook -excel $excel
