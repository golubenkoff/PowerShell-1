<#
.Synopsis
   
   Create a visual chart from a data set saved as a PNG file

.DESCRIPTION
   
   Create a visual chart from a data set saved as a PNG file
      
   Add-Type -AssemblyName System.Windows.Forms
   Add-Type -AssemblyName System.Windows.Forms.DataVisualization
   https://thwack.solarwinds.com/docs/DOC-191247

.EXAMPLE
   
   Generate-Chart -Hashtable $table -SaveFile C:\temp\graph.png -ChartType Pie

.EXAMPLE
   
   $Cities = @{London=7556900; Berlin=3429900; Madrid=3213271; Rome=2726539; Paris=2188500}
   Generate-Chart -Hashtable $Cities -SaveFile C:\temp\cities-graph.png -ChartType Column

#>
function Generate-Chart
{
    [CmdletBinding()]
    [Alias()]
    param (
        # Array of items to generate chart
        [Parameter(Mandatory = $true)]
        $Hashtable,

        # Path to save PNG file
        [Parameter(Mandatory = $true)]
        [string]$SaveFile,

        # type of chart
        [string]$ChartType = 'Column',

        # Width and height of table
        [int]$Width = 400,
        [int]$Height = 400        
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Windows.Forms.DataVisualization

    # generate report if table exists
    if (($hashtable -is [array]) -or ($hashtable -is [hashtable]) -or ($hashtable -is [System.Collections.Specialized.OrderedDictionary]))
    {
        # create chart object
        $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
        $Chart.Width = $Width
        $Chart.Height = $Height
        $Chart.Left = 20
        $Chart.Top = 20

        # create a chart area to draw on and add to chart
        $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
        $Chart.ChartAreas.Add($ChartArea)

        # add data to chart
        # $Cities = @{London=7556900; Berlin=3429900; Madrid=3213271; Rome=2726539; Paris=2188500}

        [void]$Chart.Series.Add("Data")
        $Chart.Series["Data"].Points.DataBindXY($Hashtable.Keys, $Hashtable.Values)
        $chart.Series["Data"].ChartType = $ChartType

        # Legend
        $Legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
        $Legend.IsEquallySpacedItems = $True
        $Legend.BorderColor = 'Black'

        # Customizations for pie chart
        If ($ChartType -eq "Pie")
        {
            # chart data style
            $Chart.Series["Data"]["PieLabelStyle"] = "Outside"
            $Chart.Series["Data"]["PieLineColor"] = "Black"
            #$Chart.Series["Data"]["PieDrawingStyle"] = "Concave"
            ($Chart.Series["Data"].Points.FindMaxByValue())["Exploded"] = $false
            $Chart.Series["Data"]['PieLabelStyle'] = 'Disabled'

            # legends style
            $Chart.Legends.Add($Legend)
            $chart.Series["Data"].LegendText = "#VALX (#VALY)"
        }

        # save chart to file
        $Chart.SaveImage($SaveFile, "PNG")   
    }
}
