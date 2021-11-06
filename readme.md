<div id="top"></div>

<br />
<div align="center">

![Product Name Screen Shot][product-logo]

<h3 align="center">VBA Sankey</h3>


  <p align="center">
    A VBA implementation of an interactive Sankey visualization for Excel.
  </p>

</div>

## About The Project

![Product Name Screen Shot][product-anim]

This library was created to provide a simple library to add Sankey Charts to Excel. The charts created with the library are interactive, and the appearance can be easily configured (shapes, colors&transparency, margins&dimensions).

When clicking on an element in the chart, the connected elements are highligted while all unrelated ones are hidden.
To reset the view, click on the background.

<p align="right">(<a href="#top">back to top</a>)</p>

### Installation and Usage

Import the sankeyVB.bas module into the VBA Project.

The main public function provided by the module is `sankeyDraw`.
the signature of the function is the following:

```vbnet
Public Function sankeyDraw(data As Range, targetSheet As Worksheet, chartName As String, chartTop As Double, chartLeft As Double, chartWidth As Double, chartHeight As Double, skc As sankeyConfig)
```
The parameters to this function are:
<ls>
<li><b>data:</b> a Range object set to the table containing the data to be used in the chart. This range must not contain the table headers if present, and the right-most column should contain the values for each row.</li>
<li><b>targetSheet:</b> a Worksheet object, referencing the worksheet where the chart is to be placed.</li>
<li><b>chartName:</b> a String containing the name to be used for the chart. This name will uniquely identify the chart in case multiple charts are placed in the same worksheet/workbook.</li>
<li><b>chartTop:</b> a Double value, identifying the vertical distance of the chart from the top edge of the worksheet.</li>
<li><b>chartLeft:</b> a Double value, identifying the horizontal distance of the chart from the left edge of the worksheet.</li>
<li><b>chartWidth:</b> a Double value, identifying the total width of the chart.</li>
<li><b>chartHeight:</b> a Double value, identifying the total height of the chart.</li>
<li><b>skc:</b> a sankeyConfig structure (see below), with all the configuration parameters.</li>
</ls>

<br />
<br />

The `sankeyConfig` structure is defined as follows:

```vbnet
Public Type sankeyConfig

    'geometry settings'
    verticalBlankSpacePerc As Double    '% of blank space between blocks, compared to the total available space'
    hMargin As Double                   'Horizontal Margin (in px)'
    vMargin As Double                   'Vertical Margin (in px)'
    blockColorMode As colorMode         'Mode to be used for coloring connectors. sourceColor|targetColor|gradientColor'
    lowTransparency As Double           'Low value for transparency [0-1]'
    baseTransparency As Double          'Base value for transparency [0-1]'
    highTransparency As Double          'High value for transparency [0-1]'

    'color settings'
    saturationMedian As Integer         'Median saturation for all blocks [0-255]'
    saturationVariance As Integer       'Possible saturation variance between levels [0-255]'
    luminanceMedian As Integer          'Median luminance for all blocks [0-255]'
    luminanceVariance As Integer        'Possible luminance variance between levels [0-255]'
    hueRangeForLevel As Integer         'range of hues to be spread over a single level [0-255]'
    hueVarianceBetweenLevels As Integer 'difference between base hues of one level with the next [0-255]'
    backgroundColor As colorRGB         'background color (r, g, b components)'

End Type
```
<br />

This structure can be declared manually in your code prior to calling `sankeyDraw`, otherwise the function `getDefaultSankeyConfig` can be used to get a structure pre-populated with default values.

### Sample
```vbnet
Sub testSankey()

  Dim dataRange As Range
  Dim skc As sankeyConfig
  Dim targetSheet As Worksheet

  'identify the range of data to be used
  Set dataRange = Sheet1.Range("A2:F55")

  Set targetSheet = Sheet2

  skc = getDefaultSankeyConfig

  'change the background color from the default one to white
  skc.backgroundColor.r = 255
  skc.backgroundColor.g = 255
  skc.backgroundColor.b = 255

  'create the chart

  'Public Function sankeyDraw(data As Range, targetSheet As Worksheet, chartName As String, chartTop As Double, chartLeft As Double, chartWidth As Double, chartHeight As Double, skc As sankeyConfig)
  sankeyDraw dataRange, ActiveSheet, "TEST_CHART", 50, 50, 900, 400, skc

End Sub
```
[product-logo]: img/logo.png
[product-screenshot]: img/still_preview1.png
[product-anim]: img/anim_preview1.gif
