<#
PS4DS: Visually Investigate Grouped Data
Author: Eric K. Miller
Last updated: 1 December 2025

This script contains PowerShell code for plotting data. It is part of
the PowerShell-For-DataScience module, so assumes the Math.NET Numerics
DLL is loaded, as this is integral to the histogram plotting.
#>

using namespace System.Windows.Forms
using namespace System.Windows.Forms.DataVisualization.Charting

Add-Type -AssemblyName System.Windows.Forms.DataVisualization
Add-Type -AssemblyName System.Windows.Forms

function Show-GroupedData {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        $DataObject,
        [Parameter(Mandatory)]
        [ValidateSet('Bar', 'Column', 'Pie')]
        [string]$ChartType,
        [Parameter(Mandatory)]
        [string]$XData,
        [Parameter(Mandatory)]
        [string]$YData
    )

    begin {
        # Create chart object and set its properties
        $Chart = New-Object Chart
        $Chart.Width = 700
        $Chart.Height = 500
        $ChartArea = New-Object ChartArea
        $Chart.ChartAreas.Add($ChartArea)

        # Create form to host the chart
        $Form = New-Object Form
        $Form.Width = 700
        $Form.Height = 500
        $Form.Controls.Add($Chart)
        $DataColor = 'SteelBlue'
    }
    process {
        # Create the data series and their properties for charting
        switch ($ChartType) {
            'Bar' {
                $Series = New-Object Series
                $ChartTypes = [SeriesChartType]
                $Series.ChartType = $ChartTypes::$ChartType

                $Series.Points.DataBindXY($XData, $YData)
                $Series.IsValueShownAsLabel = $true
                $Series.Color = [System.Drawing.Color]::$DataColor
                $ChartArea.AxisX.Interval = 1
                $Chart.Series.Add($Series)

                $Form.Text = 'Bar Plot'
            }
            'Column' {
                $Series = New-Object Series
                $ChartTypes = [SeriesChartType]
                $Series.ChartType = $ChartTypes::$ChartType

                $Series.Points.DataBindXY($XData, $YData)
                $Series.IsValueShownAsLabel = $true
                $Series.Color = [System.Drawing.Color]::$DataColor
                $ChartArea.AxisX.Interval = 1
                $Chart.Series.Add($Series)

                $Form.Text = 'Column Plot'
            }
            'Pie' {
                $Series = New-Object Series
                $ChartTypes = [SeriesChartType]
                $Series.ChartType = $ChartTypes::$ChartType

                for ($i = 0; $i -lt $XData.Count; $i++) {
                    [void]$Series.Points.AddXY($XData[$i], $YData[$i])
                }
            
                $Series['PieLabelStyle'] = 'Outside'  # $Series.CustomProperties
                $Series['PieLineColor'] = 'Gray'  # $Series.CustomProperties
                $Series.Label = "#AXISLABEL: #VAL (#PERCENT{P0})"
                $Chart.Series.Add($Series)
            
                $Form.Text = 'Pie Plot'
            }
    }
    end {
        $ChartTitle = New-Object Title
        $ChartTitle.Text = $ChartTitleText
        $Font = New-Object System.Drawing.Font('Lucida Console', 12, [System.Drawing.FontStyle]::Bold)
        $ChartTitle.Font = $Font
        $Chart.Titles.Add($ChartTitle)

        $ChartArea.AxisX.Title = $XAxisTitleText
        $ChartArea.AxisY.Title = $YAxisTitleText
    
        $Chart.Dock = 'Fill'  # chart adjusts to fit the entire container when the Form is resized

        #$Chart.SaveImage(...)

        $Form.Add_Shown({$Form.Activate()})  # ensures the Form gets focus
        $Form.ShowDialog()  # shows the Form
    }
}
