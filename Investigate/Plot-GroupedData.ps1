<#
PS4DS: Visually Investigate Grouped Data
Author: Eric K. Miller
Last updated: 5 December 2025

This script contains a PowerShell function to plot grouped data,
specifically bar plots, column plots, and pie plots. It is part of
the PowerShell-For-DataScience module.
#>

using namespace System.Windows.Forms
using namespace System.Windows.Forms.DataVisualization.Charting

Add-Type -AssemblyName System.Windows.Forms.DataVisualization
Add-Type -AssemblyName System.Windows.Forms

function Show-GroupedData {
    <#
    .SYNOPSIS
        Plot grouped values from a DataObject's nominal fields.

    .DESCRIPTION
        This function groups values from a DataObject's nominal fields
    to display one of bar, column, or pie chart. It uses optional
    parameters to customize the PowerShell chart's properties.
    
    .PARAMETER DataObject
        The object to use to plot data.

    .PARAMETER GroupProperty
        A string for the nominal field on which to group values.

    .PARAMETER ChartType
        A ValidateSet of chart types to select.
    
    .PARAMETER ChartTitle (Optional)
        A string with the title of the Chart.
    
    .PARAMETER NumCategories (Optional)
        An integer specifying the number of nominal values to show
    (default is 10).

    .PARAMETER DataColor (Optional)
        A parameter enabling the user to specify the color of the bar
    and column charts. The value can be a string or RGB values
    (default is 'SteelBlue').

    .PARAMETER SeriesPalette (Optional)
        A ValidateSet of color palettes to specify the color of the pie
    chart (default is 'Excel').

    .PARAMETER Save (Optional)
        A ValidateSet of image file types if the user wants to save the
    chart.
        
    .EXAMPLE
        $GroupedPlotParams = @{
    GroupProperty = 'homeworld'
    NumCategories = 5
    }
    Show-GroupedData -DataObject $StarWars_NoNull -ChartType Pie @GroupedPlotParams
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [Alias('groupplot')]
    param (
        [Parameter(Mandatory)]
        $DataObject,

        [Parameter(Mandatory)]
        [string]$GroupProperty,

        [Parameter(Mandatory)]
        [ValidateSet('Bar', 'Column', 'Pie')]
        [string]$ChartType,

        [Parameter()]
        [string]$ChartTitle = "Value Counts of Nominal Field $GroupProperty",

        [Parameter()]
        [int]$NumCategories = 10,

        [Parameter()]
        $DataColor = 'SteelBlue',

        [Parameter()]
        [ValidateSet('None', 'Bright', 'Grayscale', 'Excel', 'Light',
            'Pastel', 'EarthTones', 'SemiTransparent', 'Berry',
            'Chocolate', 'Fire', 'SeaGreen', 'BrightPastel')]
        [string]$SeriesPalette = 'Excel',

        [Parameter()]
        [ValidateSet('Jpeg', 'Png', 'Bmp', 'Tiff')]
        [string]$Save
    )

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
    
    $DataObject_Grouped = $DataObject |
        Group-Object -Property $GroupProperty -NoElement |
        Sort-Object -Property Count -Descending |
        Select-Object -First $NumCategories

    # Ensure data is properly typed for DataBindXY below
    $groupValues = [string[]]$DataObject_Grouped.Name
    $groupCounts = [int[]]($DataObject_Grouped | Select-Object -ExpandProperty Count)
    
    # Create the data series and their properties for charting
    switch ($ChartType) {
        'Bar' {
            $Series = New-Object Series
            $Series.ChartType = [SeriesChartType]::$ChartType

            # Re-sort so largest bar is on top
            $DataObject_Grouped = $DataObject_Grouped | Sort-Object -Property Count
            $groupValues = [string[]]$DataObject_Grouped.Name
            $groupCounts = [int[]]($DataObject_Grouped | Select-Object -ExpandProperty Count)

            $Series.Points.DataBindXY($groupValues, $groupCounts)

            $Series.IsValueShownAsLabel = $true
            $Series.Color = [System.Drawing.Color]::$DataColor
            $Chart.Series.Add($Series)
            $ChartArea.AxisX.Interval = 1
            
            $Form.Text = 'Bar Plot'
        }
        'Column' {
            $Series = New-Object Series
            $Series.ChartType = [SeriesChartType]::$ChartType

            $Series.Points.DataBindXY($groupValues, $groupCounts)

            $Series.IsValueShownAsLabel = $true
            $Series.Color = [System.Drawing.Color]::$DataColor
            $Chart.Series.Add($Series)
            $ChartArea.AxisX.Interval = 1

            $Form.Text = 'Column Plot'
        }
        'Pie' {
            # https://learn.microsoft.com/en-us/previous-versions/dd456674(v=vs.140)
            $Series = New-Object Series
            $Series.ChartType = [SeriesChartType]::$ChartType

            $Series.Points.DataBindXY($groupValues, $groupCounts)
            
            $Series['PieStartAngle'] = 0          # $Series.CustomProperties
            $Series['PieLabelStyle'] = 'Outside'  # $Series.CustomProperties
            $Series['PieLineColor']  = 'Gray'     # $Series.CustomProperties
            $Series.Palette = $SeriesPalette
            $Series.Label   = "#AXISLABEL: #VAL (#PERCENT{P0})"
            $Chart.Series.Add($Series)
            
            $Form.Text = 'Pie Plot'
        }
    }

    #region ChartArea settings
    $title_font = 'Microsoft Sans Serif'
    $label_fontSize = 11
    $label_fontColor = 'Gray'

    # Titles
    [void]$Chart.Titles.Add($ChartTitle)
    $Chart.Titles[0].Font = New-Object System.Drawing.Font($title_font, 16, [System.Drawing.FontStyle]::Bold)
    $ChartArea.AxisX.Title = 'Property Values'
    $ChartArea.AxisX.TitleForeColor = $label_fontColor
    $ChartArea.AxisY.Title = 'Property Counts'
    $ChartArea.AxisX.TitleForeColor = $label_fontColor

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

    $ChartArea.BackColor = 'White'
    
    # Chart adjusts to fit the entire container when the Form is resized
    $Chart.Dock = 'Fill'

    if ($Save) {
        $SaveFileDialog = New-Object SaveFileDialog

        switch ($Save) {
            'Jpeg' {
                $SaveFileDialog.Filter = 'JPEG Image File (*.jpeg) | *.jpeg'
                $ChartImageFormat = 0
            }
            'Png' {
                $SaveFileDialog.Filter = 'PNG Image File (*.png) | *.png'
                $ChartImageFormat = 1
            }
            'Bmp' {
                $SaveFileDialog.Filter = 'Windows Bitmap Image File (*.bmp) | *.bmp'
                $ChartImageFormat = 2
            }
            'Tiff' {
                $SaveFileDialog.Filter = 'Tag Image File Format (*.tiff) | *.tiff'
                $ChartImageFormat = 3
            }
        }

        $SaveFileDialog.ShowDialog() | Out-Null
        $imgFile = $SaveFileDialog.FileName

        if ($imgFile) {
            $Chart.SaveImage($imgFile, $ChartImageFormat)
            Write-Host "Chart saved at: $imgFile"
        }
        else {
            Write-Host "Image file not saved."
        }
    }
    
    $Form.Add_Shown({$Form.Activate()})  # ensures the Form gets focus
    $Form.ShowDialog()
    #endregion
}
