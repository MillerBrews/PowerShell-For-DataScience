<#
PS4DS: Visually Investigate Raw Data
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

function Show-RawData {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        $DataObject,
        [Parameter(Mandatory)]
        [string]$XData,
        [Parameter(Mandatory)]
        [string]$YData,
        [Parameter()][switch]$AddLine,
        [Parameter()]$Theta
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
        # Add points series -------------------------------
        $PointsSeries = New-Object Series
        $ChartTypes = [SeriesChartType]
        $PointsSeries.ChartType = $ChartTypes::Point
        $PointsSeries.LegendText = 'Raw data'

        $PointsSeries.Points.DataBindXY($DataObject.$XData, $DataObject.$YData)
        $PointsSeries.Color = [System.Drawing.Color]::$DataColor
        $Chart.Series.Add($PointsSeries)

        if ($AddLine) {
            # Add line series  --------------------------------
            $LineSeries = New-Object Series
            $LineSeries.ChartType = $ChartTypes::Line
            $LineSeries.LegendText = 'Regression line'

            $XData_bounds = $DataObject.$XData | Measure-Object -Minimum -Maximum
            $x_pts = $XData_bounds.Minimum..$XData_bounds.Maximum
            $y_pts = $x_pts | ForEach-Object {$Theta[0] * $_ + $Theta[1]}

            $LineSeries.Points.DataBindXY($x_pts, $y_pts)
            $lineColor = 'Goldenrod'
            $LineSeries.Color = [System.Drawing.Color]::$lineColor
            $LineSeries.BorderWidth = 3
            
            # Create a text annotation showing the line's equation
            $annotation = New-Object TextAnnotation
            $annotation.Text = "y = $([Math]::Round($Theta[0],2))*x + $([Math]::Round($Theta[1],2))"
            $annotation.Font = New-Object System.Drawing.Font('Lucida Console', 10, [System.Drawing.FontStyle]::Bold)
            $annotation.ForeColor = [System.Drawing.Color]::$lineColor
            $annotation.AxisX = $Chart.ChartAreas[0].AxisX
            $annotation.AxisY = $Chart.ChartAreas[0].AxisY
            $annotation_Xpos = 1.75 * ($DataObject.$XData | Measure-Object -Minimum).Minimum
            $annotation.AnchorX = $annotation_Xpos
            $annotation.AnchorY = 1.5 * ($Theta[0] * $annotation_Xpos + $Theta[1])
            
            $Chart.Series.Add($LineSeries)
            $Chart.Annotations.Add($annotation)

            # Create legend
            $legend = New-Object Legend
            $legend.Docking = 'Top'
            $legend.Alignment = 'Center'
            $Chart.Legends.Add($legend)
            }
        #>

        $Form.Text = 'Point and Line Plot'
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
