<#
PS4DS: Visually Investigate Statistical Data
Author: Eric K. Miller
Last updated: 2 December 2025

This script contains a PowerShell function to plot statistical data,
specifically boxplots and histograms. It is part of the
PowerShell-For-DataScience module, so assumes the Math.NET Numerics
DLL is loaded, as this is integral to the histogram plotting.
#>

#========================
#    BoxPlot function
#========================

using namespace System.Windows.Forms.DataVisualization.Charting
using namespace System.Windows.Forms

Add-Type -AssemblyName System.Windows.Forms.DataVisualization
Add-Type -AssemblyName System.Windows.Forms

function Show-BoxPlotData {
    <#
    .SYNOPSIS
        Plot boxplot statistics from a DataObject's numeric fields.

    .DESCRIPTION
        This function identifies numeric fields from a DataObject, and
    creates boxplots showing six statistical values. It uses optional
    parameters to customize the PowerShell chart's properties.
    
    .PARAMETER DataObject
        The object to use to plot data.

    .PARAMETER ChartTitle (Optional)
        A string with the title of the Chart.
    
    .PARAMETER BoxPlotPercentile (Optional)
        A string for the percentile to determine Q1 and Q3 values
    (default is '25').

    .PARAMETER BoxPlotWhiskerPercentile (Optional)
        A string for the percentile to determine the whisker values
    (default is '10').

    .PARAMETER BoxPlotShowUnusualValues (Optional)
        A Boolean string for displaying outliers--outside the ends 
    of the whiskers (default is 'True').

    .PARAMETER BoxSeriesPalette (Optional)
        A ValidateSet that enables the user to set the BoxPlot color
    palette (default is 'BrightPastel').
    
    .EXAMPLE
        $BoxPlotParams = @{
    BoxPlotWhiskerPercentile = '5'
    BoxPlotShowUnusualValues = 'False'
    BoxSeriesPalette = 'Excel'
    }
    Show-BoxPlotData $StarWars @BoxPlotParams
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [Alias('boxplot')]
    param (
        [Parameter(Mandatory)]$DataObject,
        [Parameter()][string]$ChartTitle = 'Distribution of Numeric Fields',
        [Parameter()][string]$BoxPlotPercentile        = '25',    # default
        [Parameter()][string]$BoxPlotWhiskerPercentile = '10',    # default
        [Parameter()][string]$BoxPlotShowUnusualValues = 'True',  # default
        [Parameter()]
        [ValidateSet('None', 'Bright', 'Grayscale', 'Excel', 'Light',
        'Pastel', 'EarthTones', 'SemiTransparent', 'Berry', 'Chocolate',
        'Fire', 'SeaGreen', 'BrightPastel')]
        [string]$BoxSeriesPalette = 'BrightPastel'
    )
    
    #region Boilerplate charting objects creation
    # Chart to hold the ChartArea and Series
    $Chart = New-Object Chart
    $Chart.Width = 700
    $Chart.Height = 500
    $ChartArea = New-Object ChartArea
    $ChartArea.Name = 'ChartArea'
    $Chart.ChartAreas.Add($ChartArea)

    # Form to hold the chart
    $Form = New-Object Form
    $Form.Width = 700
    $Form.Height = 500
    $Form.Controls.Add($Chart)
    $Form.Text = 'BoxPlot'
    #endregion

    #region PointSeries creation
    # Get numeric fields
    $numeric_fields = $DataObject[0].PSObject.Properties |
        ? {$_.Value -is [double] -or $_.Value -is [int]} |
        % {$_.Name}

    $fields_index = 0
    $PointSeriesNames = @()

    foreach ($field in $numeric_fields) {
        # Extract values for each field
        $field_values = $DataObject.$field
    
        # Point data series
        $PointSeries = New-Object Series
        $PointSeries.Name = "PointData_$fields_index"
        $PointSeries.ChartType = [SeriesChartType]::Point
        $PointSeries.ChartArea = $ChartArea.Name
        [double[]]$field_values | % {[void]$PointSeries.Points.AddXY($fields_index + 1, $_)}
        $PointSeries.ToolTip     = "$($field): #VALY1{0.0}"
        $PointSeries.MarkerStyle = [MarkerStyle]::Circle
        $PointSeries.MarkerSize  = 6
        $PointSeries.Color       = 'Gray'
        $Chart.Series.Add($PointSeries)

        # Collect to create automatic boxplots
        $PointSeriesNames += $PointSeries.Name
        
        # Labels for BoxPlots from the numeric fields
        $Label = New-Object CustomLabel
        $Label.FromPosition = $fields_index + 1 - 0.5
        $Label.ToPosition   = $fields_index + 1 + 0.5
        $Label.Text = $field
        $ChartArea.AxisX.CustomLabels.Add($Label)

        $fields_index++
    }
    #endregion

    #region BoxSeries creation
    $BoxSeries = New-Object Series
    $BoxSeries.ChartType = [SeriesChartType]::BoxPlot
    $BoxSeries.ChartArea = $ChartArea.Name
    <# BoxPlot data series are calculated and added for each
    Point series, delimited by semicolons #>
    $BoxSeries['BoxPlotSeries'] = $PointSeriesNames -join ';'
    $BoxSeries['BoxPlotPercentile']        = $BoxPlotPercentile 
    $BoxSeries['BoxPlotWhiskerPercentile'] = $BoxPlotWhiskerPercentile 
    $BoxSeries['BoxPlotShowUnusualValues'] = $BoxPlotShowUnusualValues
    $BoxSeries['PointWidth'] = '0.5'
    $BoxSeries.Palette     = $BoxSeriesPalette
    $BoxSeries.BorderColor = 'Black'
    $BoxSeries.BorderWidth = 1
    $BoxSeries.ToolTip = "BoxPlot Statistics:\n\n" +
                         "Max Whisker: #VALY2{0.0}\n" +
                         "Q3 (75%): #VALY4{0.0}\n" +
                         "Median (- - -): #VALY6{0.0}\n" +
                         "Avg: #VALY5{0.0}\n" +
                         "Q1 (25%): #VALY3{0.0}\n" +
                         "Min Whisker: #VALY1{0.0}\n"
    $Chart.Series.Add($BoxSeries)
    #endregion

    #region ChartArea settings
    $title_font = 'Microsoft Sans Serif'
    $label_fontSize = 11
    $label_fontColor = 'Gray'

    # Titles
    [void]$Chart.Titles.Add($ChartTitle)
    $Chart.Titles[0].Font = New-Object System.Drawing.Font($title_font, 16)
    $ChartArea.AxisX.Title = 'Numeric Fields'
    $ChartArea.AxisX.TitleFont = New-Object System.Drawing.Font($title_font, $label_fontSize)
    $ChartArea.AxisX.TitleForeColor = $label_fontColor
    $ChartArea.AxisY.Title = 'BoxPlot Statistics'
    $ChartArea.AxisY.TitleFont = New-Object System.Drawing.Font($title_font, $label_fontSize)
    $ChartArea.AxisY.TitleForeColor = $label_fontColor

    # AxisX
    $ChartArea.AxisX.LineColor = 'LightGray'
    $ChartArea.AxisX.LineWidth = 2
    $ChartArea.AxisX.MajorTickMark.LineColor = 'LightGray'
    $ChartArea.AxisX.MajorGrid.LineWidth     = 0
    $ChartArea.AxisX.Minimum  = 0
    $ChartArea.AxisX.Interval = 1
    
    # AxisY
    $ChartArea.AxisY.LineColor = 'LightGray'
    $ChartArea.AxisY.LineWidth = 2
    $ChartArea.AxisY.MajorTickMark.LineColor = 'LightGray'
    $ChartArea.AxisY.MajorGrid.LineWidth     = 0
    $ChartArea.AxisY.LabelStyle.ForeColor = $label_fontColor

    $ChartArea.BackColor = 'LightGray'

    # Chart adjusts to fit the entire container when the Form is resized
    $Chart.Dock = 'Fill'

    $Form.Add_Shown({$Form.Activate()})  # ensures the Form gets focus
    $Form.ShowDialog()
    #endregion
}