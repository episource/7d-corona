# Copyright 2020 Philipp Serr (episource)
# 
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
# 
#     http://www.apache.org/licenses/LICENSE-2.0
# 
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

# Enable common parameters
[CmdletBinding()] 
Param(
)

$xlsx = $PSScriptRoot + "\7d-corona.xlsx"
$imgFile = $PSScriptRoot + "\7d-corona.png"
$imgWidth = $None

$sheetChart = "Schaubild"
$chartName = "Neuinfektionen"
$sheetCountryMap = @{
    "ECDC Deutschland" = "germany"
    "ECDC Frankreich" = "france"
    "ECDC Italien" = "italy"
    "ECDC Spanien" = "spain"
    "ECDC Österreich" = "austria"
    "ECDC Schweiz" = "switzerland"
}


$queryUrlEcdc = "https://opendata.ecdc.europa.eu/covid19/casedistribution/csv"


function Update-EcdcSheet($sheet, $data, $country, $lastUpdate) {
    $curRow = 3
    $data | ?{
        $_.country -eq $country 
    } | %{
        $curRow++
        $sheet.Cells.Item($curRow, 1) = $_.date.ToOADate()
        $sheet.Cells.Item($curRow, 2) = $_.cases
        $sheet.Cells.Item($curRow, 4) = $_.population
    }
    
    $sheet.Cells.Item(1, 1) = $lastUpdate
}

function Get-DataFromEcdc() {
    $r = Invoke-WebRequest -UseBasicParsing $queryUrlEcdc
    
    return [String]::new($r.Content) | %{
        $_ -split "[\r\n]+" 
    } | Select-Object -Skip 1 | ?{
        $_.Length -gt 0
    } | %{
        $row = $_.Split(",")
        [PSCustomObject]@{ 
            "date"=[DateTime]::ParseExact($row[0], "dd/MM/yyyy", $null)
            "cases"=$row[4]
            "country"=$row[6]
            "population"=$row[9]
        }
    } 
}

$now = Get-Date -format "dddd yyyy-MM-dd HH:mm"
$lastUpdate = "Datenabruf: $now"

$excel = New-Object -ComObject Excel.Application
try {
    $excel.Visible = $true
    $excel.ScreenUpdating = $False 
    $excelWb = $excel.Workbooks.Open($xlsx)
    $excelSheetChart = $excelWb.Sheets($sheetChart)
    
    
    $ecdcData = Get-DataFromEcdc
    
    $sheetCountryMap.Keys | %{
        $excelSheet = $excelWb.Sheets($_)
        Update-EcdcSheet $excelSheet $ecdcData $sheetCountryMap[$_] $lastUpdate
    }
    
    
    $excelChart = $excelSheetChart.ChartObjects($chartName)
    $excelChartXValues = $excelChart.Chart.SeriesCollection(1).XValues
    
    $wc = [Math]::floor($excelChartXValues.Length / 7)
    $lastOaDate = $excelChartXValues.Get(1)
    $lastDate = [System.DateTime]::FromOaDate($lastOaDate)
    $daysToNextMonday = (8 - $lastDate.DayOfWeek) % 7
    
    $excelChart.Chart.Axes(1).MinimumScale = $lastOaDate + $daysToNextMonday - $wc * 7
    $excelChart.Chart.Axes(1).MaximumScale = $lastOaDate + $daysToNextMonday
    $excelChart.Chart.ChartTitle.Text = "Neuinfektionen/100k (7 Tage) - $lastUpdate"
    
    $excel.ScreenUpdating = $True
    
    try {
        $excelChart.CopyPicture([Microsoft.Office.Interop.Excel.XlPictureAppearance]::xlScreen, [Microsoft.Office.Interop.Excel.XlCopyPictureFormat]::xlBitmap)
        
        $img = Get-Clipboard -Format Image
        if (-not $img) {
            throw "clipboard empty"
        }

        if ($imgWidth -eq $None) {
            $imgWidth = $img.Width
        }
        $imgHeight = [int]($img.Height / $img.Width * $imgWidth)
        $outBitmap = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $imgWidth, $imgHeight
        
        $outG = [System.Drawing.Graphics]::FromImage($outBitmap)
        $outG.SmoothingMode = "HighQuality"
        $outG.InterpolationMode = "HighQualityBicubic"
        $outG.PixelOffsetMode = "HighQuality"
        $outGRectangle = 
        $outG.DrawImage($img, [System.Drawing.Rectangle]::new(0, 0, $imgWidth, $imgHeight))
        
        $outBitmap.Save($imgFile)
    } catch {
        Write-Warning "Failed to save chart image: $_"
        throw
    }
    
    $excelWb.Save()
} finally {
    $excel.ScreenUpdating = $True
    $excel.Quit()
}

