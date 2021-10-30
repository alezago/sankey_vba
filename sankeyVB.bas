Attribute VB_Name = "sankeyVB"
Option Explicit

'Use the correct Windows API declaration based on VBA version
#If VBA7 Then
    Public Declare PtrSafe Function ColorRGBToHLS Lib "shlwapi.dll" (ByVal clrRGB As Long, pwHue As Long, pwLuminance As Long, pwSaturation As Long) As Long
    Public Declare PtrSafe Function ColorHLSToRGB Lib "shlwapi.dll" (ByVal wHue As Long, ByVal wLuminance As Long, ByVal wSaturation As Long) As Long
#Else
    'In case any error is highlighted in the next two lines, you can safely ignore it
    'Ref. to: docs.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/compile-error-editing-vba-macro
    Public Declare Function ColorRGBToHLS Lib "shlwapi.dll" (ByVal clrRGB As Long, pwHue As Long, pwLuminance As Long, pwSaturation As Long) As Long
    Public Declare Function ColorHLSToRGB Lib "shlwapi.dll" (ByVal wHue As Long, ByVal wLuminance As Long, ByVal wSaturation As Long) As Long
#End If

'Shape properties model constants
Private Const SHP_PROPERTY_SEPARATOR As String = "%-%"
Private Const SHP_PROPERTY_EQUAL As String = "%=%"
Private Const SHP_PAIR_PROPERTY As Integer = 0
Private Const SHP_PAIR_VALUE As Integer = 1

'property names for storing in each shape the currently free space for connectors
Private Const PROP_OFFSET_OUT As String = "OFFSET_OUT"
Private Const PROP_OFFSET_IN As String = "OFFSET_IN"

Private Const SHAPE_NAME_PREFIX As String = "VB_SK_"
Private Const VERTICAL_SPACING_PERC As Double = 0.3    'percentage used by blank space, compared to the whole box size
Private Const BOX_TO_BLANK_RATIO_H As Double = 0.11
Private Const LEVEL_DISTANCE As Double = 60    'not used, now dynamically based on the available space and the number of levels
Private Const HORIZONTAL_MARGIN As Double = 45
Private Const VERTICAL_MARGIN As Double = 25

Private Const BASE_TRANSPARENCY As Double = 0.6    '0=solid, 1=invisible
Private Const HIGH_TRANSPARENCY As Double = 0.85
Private Const LOW_TRANSPARENCY As Double = 0.07

'label constants
Private Const LABEL_WIDTH As Double = 150
Private Const LABEL_HEIGHT As Double = 25
Private Const LABEL_OFFSET As Double = 0    'this does NOT count the margin of between the textbox and the text

'space should be measured including (half of) the empty space between blocks
Private Const LABEL_MIN_SPACE_1L As Double = 20
Private Const LABEL_MIN_SPACE_2L As Double = 34

'background color / border
Private Const BG_COLOR_R As Integer = 15
Private Const BG_COLOR_G As Integer = 15
Private Const BG_COLOR_B As Integer = 15

Private Const BG_BORDER_COLOR_R As Integer = 10
Private Const BG_BORDER_COLOR_G As Integer = 10
Private Const BG_BORDER_COLOR_B As Integer = 10

Private Const THRESHOLD_DARK_BG As Double = 100    'if the brightness is under this value, the background is considered "dark" (label text becomes white)

'possible color modes
Private Enum colorMode
    sourceColor = 1
    targetColor = 2
    gradientColor = 3
End Enum

Private Enum vbsk_colorTheme
    vbsk_colorThemeLight = 1
    vbsk_colorThemeDark = 2
End Enum

Private Type colorHSL
    h As Single
    s As Single
    l As Single
End Type

Private Type colorRGB
    r As Single
    g As Single
    b As Single
End Type

Public Type sankeyConfig

    'geometry settings
    verticalBlankSpacePerc As Double    '% of blank space between blocks, compared to the total available space
    hMargin As Double
    vMargin As Double
    blockColorMode As colorMode
    lowTransparency As Double
    baseTransparency As Double
    highTransparency As Double
    
    'color settings
    saturationMedian As Integer
    saturationVariance As Integer
    luminanceMedian As Integer
    luminanceVariance As Integer
    hueRangeForLevel As Integer
    hueVarianceBetweenLevels As Integer
    backgroundColor As colorRGB
    
End Type

Private Type shapeGeometry
    top As Double
    bottom As Double
    left As Double
    right As Double
End Type

Private Type dataStructure
    originalDataTable As Range
    sortedDataTable As Range
    mbrSheet As Worksheet
    totalValue As Double
    levelCount As Integer
    elemCountForLevel() As Long
End Type

Private Function RGBToHLS(crgb As colorRGB) As colorHSL

  ' wrapper for ColorRGBToHLS
  Dim iHue As Long
  Dim iLum As Long
  Dim iSat As Long
  
  ColorRGBToHLS RGB(crgb.r, crgb.g, crgb.b), iHue, iLum, iSat
  RGBToHLS.h = iHue
  RGBToHLS.l = iLum
  RGBToHLS.s = iSat
End Function

Private Function HSLToRGB(chls As colorHSL) As colorRGB

  ' wrapper for ColorHLSToRGB
  Dim iRGB As Long
  
  iRGB = ColorHLSToRGB(chls.h, chls.l, chls.s)
  HSLToRGB.r = iRGB And 255
  HSLToRGB.g = (iRGB / 256) And 255
  HSLToRGB.b = (iRGB / 65536) And 255
  
End Function

Private Function getColorHSL(h_base As Integer, h_var As Integer, s_base As Integer, s_var As Integer, l_base As Integer, l_var As Integer) As colorHSL
    
    Dim s As Single, l As Single
    
    getColorHSL.h = (h_base + (h_var / 2) * (2 * Rnd() - 1)) Mod 255
    
    s = s_base + (s_var / 2) * (2 * Rnd() - 1)
    If s > 255 Then
        getColorHSL.s = 255
    Else
        getColorHSL.s = s
    End If
    
    l = l_base + (l_var / 2) * (2 * Rnd() - 1)
    If l > 255 Then
        getColorHSL.l = 255
    Else
        getColorHSL.l = l
    End If

End Function

Private Function getSankeyConfig() As sankeyConfig
    
    'general Settings
    getSankeyConfig.hMargin = Range("CONFIG_MARGIN_H").value
    getSankeyConfig.vMargin = Range("CONFIG_MARGIN_V").value
    getSankeyConfig.verticalBlankSpacePerc = Range("CONFIG_SPACING").value / 100
    getSankeyConfig.baseTransparency = Range("CONFIG_TRANSP_BASE").value / 100
    getSankeyConfig.lowTransparency = Range("CONFIG_TRANSP_LOW").value / 100
    getSankeyConfig.highTransparency = Range("CONFIG_TRANSP_HIGH").value / 100
    Select Case Range("CONFIG_COLOR_MODE").value
        Case "Gradient"
            getSankeyConfig.blockColorMode = gradientColor
        Case "Source"
            getSankeyConfig.blockColorMode = sourceColor
        Case "Target"
            getSankeyConfig.blockColorMode = targetColor
    End Select
    
    
    'Color Settings
    getSankeyConfig.hueRangeForLevel = Range("CONFIG_COLOR_HUE_VAR").value
    getSankeyConfig.hueVarianceBetweenLevels = Range("CONFIG_COLOR_HUE_JUMP").value
    getSankeyConfig.luminanceMedian = Range("CONFIG_COLOR_LUM_MED").value
    getSankeyConfig.luminanceVariance = Range("CONFIG_COLOR_LUM_VAR").value
    getSankeyConfig.saturationMedian = Range("CONFIG_COLOR_SAT_MED").value
    getSankeyConfig.saturationVariance = Range("CONFIG_COLOR_SAT_VAR").value
    
    Select Case Range("CONFIG_COLOR_BACKGROUND").value
        Case "Dark Grey"
            getSankeyConfig.backgroundColor.r = 45
            getSankeyConfig.backgroundColor.g = 45
            getSankeyConfig.backgroundColor.b = 45
            
        Case "Navy Blue"
            getSankeyConfig.backgroundColor.r = 43
            getSankeyConfig.backgroundColor.g = 67
            getSankeyConfig.backgroundColor.b = 103
            
        Case "Black"
            getSankeyConfig.backgroundColor.r = 0
            getSankeyConfig.backgroundColor.g = 0
            getSankeyConfig.backgroundColor.b = 0
        
        Case "White"
            getSankeyConfig.backgroundColor.r = 255
            getSankeyConfig.backgroundColor.g = 255
            getSankeyConfig.backgroundColor.b = 255
        
        Case "Light Grey"
            getSankeyConfig.backgroundColor.r = 242
            getSankeyConfig.backgroundColor.g = 242
            getSankeyConfig.backgroundColor.b = 242
    End Select
    
End Function


Private Function shapeExistsInSheet(shapeName As String, ws As Worksheet) As Boolean

shapeExistsInSheet = False

On Error GoTo notFound
shapeExistsInSheet = ws.Shapes(shapeName).Name = shapeName

notFound:

End Function

'Retrieve the bounding box coordinates from a Shape object
Private Function getGeomForShape(sh As Shape) As shapeGeometry

    getGeomForShape.top = sh.top
    getGeomForShape.bottom = sh.top + sh.height
    getGeomForShape.left = sh.left
    getGeomForShape.right = sh.left + sh.width

End Function

Private Function formatValue(value As Double, Optional decimalNum As Integer = 2, Optional useAbbreviations As Boolean = True) As String

Dim suffix As String
Dim valTemp As Double

valTemp = value

If useAbbreviations Then
    If valTemp > 1000000000 Then
        valTemp = valTemp / 1000000000
        suffix = "B"
    ElseIf valTemp > 1000000 Then
        valTemp = valTemp / 1000000
        suffix = "M"
    ElseIf valTemp > 1000 Then
        valTemp = valTemp / 1000
        suffix = "K"
    End If
End If

valTemp = Round(valTemp, decimalNum)
formatValue = LTrim(Str(valTemp) & suffix)

End Function

'Create the label associated with a specific block
'@TODO: clip the vertical size if it exceeds the size of the bounding box (should be done before anything else, so the evaluation of text height considers the correct label height)
Private Function addLabelForBlock(block As Shape, labelText As String, labelValue As Double, rightOfBlock As Boolean, chartID As String, vSpacing As Double, hSpacing As Double, backgroundColor As colorRGB) As Shape
    
    Dim top As Double
    Dim left As Double
    Dim text As String
    
    If block.height + vSpacing > LABEL_MIN_SPACE_2L Then
        text = labelText & vbCr & formatValue(labelValue, 2, True)
    ElseIf block.height + vSpacing > LABEL_MIN_SPACE_1L Then
        text = labelText
    Else
        text = ""
    End If

    top = block.top - vSpacing / 2
    
    If rightOfBlock Then
        left = block.left + block.width + LABEL_OFFSET
    Else
        left = block.left - hSpacing / 2
    End If
    
    Dim label As Shape
    
    Set label = block.TopLeftCell.Worksheet.Shapes.AddTextbox(msoTextOrientationHorizontal, left, top, hSpacing / 2 - LABEL_OFFSET, block.height + vSpacing)
    label.Name = block.Name & "_LBL"
    label.TextFrame.Characters.Delete
    label.TextFrame.VerticalAlignment = xlVAlignCenter
    If rightOfBlock Then
        label.TextFrame.HorizontalAlignment = xlHAlignLeft
    Else
        label.TextFrame.HorizontalAlignment = xlHAlignRight
    End If
    label.TextFrame.Characters.Insert text
    
    'set the color of the label text as white if the background luminance is lower than the threshold
    If (backgroundColor.r + backgroundColor.g + backgroundColor.b) / 3 < THRESHOLD_DARK_BG Then
        label.TextFrame.Characters.Font.Color = RGB(230, 230, 230)
    End If
    
    label.Line.Visible = msoFalse
    label.Fill.Visible = msoFalse
    
    label.ZOrder msoBringToFront
    
    setPropertyValue label, "TYPE", "LABEL"
    setPropertyValue label, "CHARTID", chartID
    
    label.OnAction = "clickHandler"
    
    'readjust the vertical position, because of autofit
    'label.top = block.top + (block.height / 2) - (label.height / 2)
    'label.top = block.top
    'label.height = block.height
    
End Function

'returns the connecting shape
Private Function connectShapesWithOffset(S1 As Shape, S2 As Shape, Optional useStraight As Boolean = False, Optional useProportional As Boolean = False, Optional verticalSize As Double = 0, Optional useOffset As Boolean = False, Optional colorMode As colorMode = sourceColor) As Shape

Dim geom1 As shapeGeometry
Dim geom2 As shapeGeometry

Dim source As shapeGeometry
Dim target As shapeGeometry

'helper function, to get coordinates and forget about the actual shapes
geom1 = getGeomForShape(S1)
geom2 = getGeomForShape(S2)

If geom2.left > geom1.right Then
    source = geom1
    target = geom2
ElseIf geom2.right < geom1.left Then
    source = geom2
    target = geom1
Else    'the two shapes are overlapping, cannot link them
        Debug.Print "Error: cannot connect the two shapes because their bounding boxes are overlapping."
        Exit Function
End If

'check if the two shapes are on the same sheet
If S1.TopLeftCell.Worksheet.Index <> S2.TopLeftCell.Worksheet.Index Then
    Debug.Print "Error: cannot connect two shapes on different worksheets."
    Exit Function
End If

Dim DISTANCE As Double
Dim offsetSource As Double
Dim offsetTarget As Double

DISTANCE = target.left - source.right

If useOffset Then
    offsetSource = CDbl(getPropertyValue(S1, PROP_OFFSET_OUT, "0"))
    offsetTarget = CDbl(getPropertyValue(S2, PROP_OFFSET_IN, "0"))
Else
    offsetSource = 0
    offsetTarget = 0
End If



Dim ff As FreeformBuilder
Dim currentSheet As Worksheet

Set currentSheet = S1.TopLeftCell.Worksheet

Set ff = currentSheet.Shapes.BuildFreeform(msoEditingCorner, source.right, source.top + offsetSource)

If useStraight Then
    ff.AddNodes msoSegmentLine, msoEditingAuto, target.left, target.top + offsetTarget
Else
    ff.AddNodes msoSegmentCurve, msoEditingCorner, source.right + (DISTANCE / 2), source.top + offsetSource, target.left - (DISTANCE / 2), target.top + offsetTarget, target.left, target.top + offsetTarget
End If

If verticalSize = 0 Then
    ff.AddNodes msoSegmentLine, msoEditingAuto, target.left, target.bottom    'this case is pretty useless, but we keep the possibility anyway
    If useStraight Then
        ff.AddNodes msoSegmentLine, msoEditingAuto, source.right, source.bottom
    Else
        ff.AddNodes msoSegmentCurve, msoEditingCorner, target.left - (DISTANCE / 2), target.bottom, source.right + (DISTANCE / 2), source.bottom, source.right, source.bottom
    End If
Else
    'validate the vertical size, never exceed the size of the blocks
    If verticalSize > S1.height - offsetSource Then
        verticalSize = S1.height - offsetSource
    End If
    
    If verticalSize > S2.height - offsetTarget Then
        verticalSize = S2.height - offsetTarget
    End If
    
    
    ff.AddNodes msoSegmentLine, msoEditingAuto, target.left, target.top + offsetTarget + verticalSize
    If useStraight Then
        ff.AddNodes msoSegmentLine, msoEditingAuto, source.right, source.top + offsetSource + verticalSize
    Else
        ff.AddNodes msoSegmentCurve, msoEditingCorner, target.left - (DISTANCE / 2), target.top + offsetTarget + verticalSize, source.right + (DISTANCE / 2), source.top + offsetSource + verticalSize, source.right, source.top + offsetSource + verticalSize
    End If
End If

'close the line
ff.AddNodes msoSegmentLine, msoEditingAuto, source.right, source.top + offsetSource

'if we were using the offset, update the shape properties
If useOffset Then
    If verticalSize = 0 Then
        setPropertyValue S1, PROP_OFFSET_OUT, S1.height
        setPropertyValue S2, PROP_OFFSET_IN, S2.height
    Else
        setPropertyValue S1, PROP_OFFSET_OUT, offsetSource + verticalSize
        setPropertyValue S2, PROP_OFFSET_IN, offsetTarget + verticalSize
    End If
End If

Set connectShapesWithOffset = ff.ConvertToShape

'color the shapes based on the selected color mode
Select Case colorMode
    Case sourceColor:
        connectShapesWithOffset.Fill.ForeColor.RGB = S1.Fill.ForeColor.RGB
        connectShapesWithOffset.Line.ForeColor.RGB = S1.Line.ForeColor.RGB
    Case targetColor:
        connectShapesWithOffset.Fill.ForeColor.RGB = S2.Fill.ForeColor.RGB
        connectShapesWithOffset.Line.ForeColor.RGB = S2.Line.ForeColor.RGB
    Case gradientColor:
        
        connectShapesWithOffset.Fill.TwoColorGradient msoGradientVertical, 2
        connectShapesWithOffset.Fill.GradientStops(1).Color = S2.Fill.ForeColor.RGB
        connectShapesWithOffset.Fill.GradientStops(2).Color = S1.Fill.ForeColor.RGB
        connectShapesWithOffset.Line.Visible = msoFalse
        'TODO:Implement
        'connectShapesWithOffset.Fill.ForeColor.RGB = S1.Fill.ForeColor.RGB
        'connectShapesWithOffset.Line.ForeColor.RGB = S1.Line.ForeColor.RGB
        
End Select

End Function


Private Function getPropertyValue(sh As Shape, propertyName As String, Optional valueNotFound As String = "") As String

Dim altText As String
Dim spl1() As String
Dim spl2() As String
Dim i As Long

getPropertyValue = valueNotFound

altText = sh.AlternativeText

spl1 = Split(altText, SHP_PROPERTY_SEPARATOR)

For i = LBound(spl1) To UBound(spl1)
    spl2 = Split(spl1(i), SHP_PROPERTY_EQUAL)
    If UBound(spl2) = 1 Then
        If spl2(0) = propertyName Then
            getPropertyValue = spl2(1)
            Exit Function
        End If
    End If
Next i

End Function



Private Function setPropertyValue(sh As Shape, propertyName As String, ByVal propertyValue As String)

Dim altText As String
Dim newText As String
Dim textSplit() As String
Dim propertyValuePair As Integer
Dim propValue() As String
Dim replaced As Boolean

altText = sh.AlternativeText

If altText = "" Then
    sh.AlternativeText = propertyName & SHP_PROPERTY_EQUAL & propertyValue
    Exit Function
End If

textSplit = Split(altText, SHP_PROPERTY_SEPARATOR)


'TODO: change the approach to use Mid() and replace directly in string, instead of using up all this memory and time
replaced = False

For propertyValuePair = LBound(textSplit) To UBound(textSplit)
    
    propValue = Split(textSplit(propertyValuePair), SHP_PROPERTY_EQUAL)
    
    If propValue(SHP_PAIR_PROPERTY) <> propertyName Then
        newText = newText & propValue(SHP_PAIR_PROPERTY) & SHP_PROPERTY_EQUAL & propValue(SHP_PAIR_VALUE) & SHP_PROPERTY_SEPARATOR
    Else
        replaced = True
        newText = newText & propValue(SHP_PAIR_PROPERTY) & SHP_PROPERTY_EQUAL & propertyValue & SHP_PROPERTY_SEPARATOR
    End If
    
Next propertyValuePair

If replaced = False Then
    sh.AlternativeText = altText & SHP_PROPERTY_SEPARATOR & propertyName & SHP_PROPERTY_EQUAL & propertyValue
Else
    newText = left(newText, Len(newText) - Len(SHP_PROPERTY_SEPARATOR))
    sh.AlternativeText = newText
End If

End Function


Private Function printAllPropertyValues(sh As Shape, Optional useMsgBox As Boolean = False)

    If useMsgBox Then
        MsgBox Replace(Replace(sh.AlternativeText, SHP_PROPERTY_SEPARATOR, vbCrLf), SHP_PROPERTY_EQUAL, "="), vbOKOnly
    Else
        Debug.Print Replace(Replace(sh.AlternativeText, SHP_PROPERTY_SEPARATOR, vbCrLf), SHP_PROPERTY_EQUAL, "=")
    End If

End Function


Sub testSankeyVB()
    
    'Check if there is already a chart present

    If shapeExistsInSheet(SHAPE_NAME_PREFIX & "TSK_CHARTAREA", ActiveSheet) Then
        MsgBox "Please delete the existing chart before creating a new one."
        Exit Sub
        
    End If
    
    
    Dim sourceData As Range
    Dim chartRange As Range
    
    Dim top As Double
    Dim left As Double
    Dim width As Double
    Dim height As Double
    
    Dim skConfig As sankeyConfig

    'get all the configuration from the config sheet
    skConfig = getSankeyConfig
    
    'On Error GoTo errExit
    Sheet3.Shapes("BTN_CRE").Fill.ForeColor.RGB = RGB(225, 250, 225)
    
    Application.Wait Now() + TimeValue("00:00:01")
    
    On Error GoTo errExit
    Set sourceData = Application.InputBox("Select the data range you want to chart. Please do NOT include any header.", "Select Dataset", Type:=8)
    
    On Error GoTo 0
    Set chartRange = Range("TEST_CHART_RANGE")
    
    If sourceData Is Nothing Or sourceData Is Nothing Then
        MsgBox "Please make sure the selections are correct!", vbCritical, "Error"
        Exit Sub
    End If
    
    top = chartRange.top
    left = chartRange.left
    width = chartRange.width
    height = chartRange.height
    
    drawSankey sourceData, chartRange.Worksheet, "TSK", top, left, width, height, skConfig
    
    Sheet3.Shapes("BTN_CRE").Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    Exit Sub
    
errExit:
    MsgBox "Error! Please check your selections."
    Sheet3.Shapes("BTN_CRE").Fill.ForeColor.RGB = RGB(255, 255, 255)
    
End Sub

Private Function setupDataStructures(originalDataTable As Range, skc As sankeyConfig) As dataStructure
    Set setupDataStructures.originalDataTable = originalDataTable
    setupDataStructures.totalValue = WorksheetFunction.Sum(originalDataTable.Columns(originalDataTable.Columns.Count))
    setupDataStructures.levelCount = originalDataTable.Columns.Count - 1
    
    Dim sortedDataSheet As Worksheet
    
    Set sortedDataSheet = Sheets.Add
    sortedDataSheet.Name = "SANKEYVB_SORTED"
    sortedDataSheet.Visible = xlSheetVeryHidden
    
    Set setupDataStructures.sortedDataTable = Range(sortedDataSheet.Cells(1, 1), sortedDataSheet.Cells(originalDataTable.Rows.Count, originalDataTable.Columns.Count))
    setupDataStructures.sortedDataTable.value = originalDataTable.value
    
    ReDim setupDataStructures.elemCountForLevel(1 To setupDataStructures.levelCount)
    
    Dim memberDataSheet As Worksheet
    
    Set memberDataSheet = Worksheets.Add
    memberDataSheet.Name = "SANKEYVB_MEMBERS"
    memberDataSheet.Visible = xlSheetVeryHidden
    
    Set setupDataStructures.mbrSheet = memberDataSheet
    
    Dim i As Integer
    Dim j As Integer
    Dim sortRange As Range
    Dim startColor As colorRGB
    Dim endColor As colorRGB
    Dim gradient As Variant
    Dim h As Integer    'hue
    Dim cHSL As colorHSL
    
    'starting hue is a random value
    h = Int(255 * Rnd())
    
    For i = 1 To setupDataStructures.levelCount
    
        h = (h + skc.hueVarianceBetweenLevels) Mod 255
        
        Range(memberDataSheet.Cells(1, i), memberDataSheet.Cells(setupDataStructures.sortedDataTable.Rows.Count, i)).value = setupDataStructures.sortedDataTable.Columns(i).value
        Range(memberDataSheet.Cells(1, i), memberDataSheet.Cells(setupDataStructures.sortedDataTable.Rows.Count, i)).RemoveDuplicates 1, xlNo
        
        If memberDataSheet.Cells(2, i).value = "" Then
            setupDataStructures.elemCountForLevel(i) = 1
            
            'add color
            'TODO: Palettes
            'single Color
            cHSL = getColorHSL(h, 0, skc.saturationMedian, skc.saturationVariance, skc.luminanceMedian, skc.luminanceVariance)
            startColor = HSLToRGB(cHSL)
            memberDataSheet.Cells(1, i).Interior.Color = RGB(startColor.r, startColor.g, startColor.b)
            
            
        Else
            Set sortRange = Range(memberDataSheet.Cells(1, i), memberDataSheet.Cells(1, i).End(xlDown))
            setupDataStructures.elemCountForLevel(i) = sortRange.Cells.Count
            memberDataSheet.Sort.SortFields.Clear
            memberDataSheet.Sort.SortFields.Add sortRange, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            memberDataSheet.Sort.SetRange sortRange
            memberDataSheet.Sort.Header = xlNo
            memberDataSheet.Sort.MatchCase = True
            memberDataSheet.Sort.Orientation = xlTopToBottom
            memberDataSheet.Sort.SortMethod = xlPinYin
            memberDataSheet.Sort.Apply
            
            cHSL = getColorHSL(h, 0, skc.saturationMedian, skc.saturationVariance, skc.luminanceMedian, skc.luminanceVariance)
            startColor = HSLToRGB(cHSL)
            'startColor.r = Int(255 * Rnd)
            'startColor.g = Int(255 * Rnd)
            'startColor.b = Int(255 * Rnd)
            
            cHSL = getColorHSL(h + skc.hueRangeForLevel, 0, skc.saturationMedian, skc.saturationVariance, skc.luminanceMedian, skc.luminanceVariance)
            endColor = HSLToRGB(cHSL)
            'endColor.r = Int(255 * Rnd)
            'endColor.g = Int(255 * Rnd)
            'endColor.b = Int(255 * Rnd)
            
            Debug.Print "Colors for level " & i & ": from (" & startColor.r & ", " & startColor.g & ", " & startColor.b & ") to (" & endColor.r & ", " & endColor.g & ", " & endColor.b & ")"
            
            gradient = getGradientBetweenColors(startColor, endColor, sortRange.Cells.Count)
            
            For j = 1 To sortRange.Cells.Count
                sortRange.Cells(j, 1).Interior.Color = RGB(Int(gradient(j, 1)), Int(gradient(j, 2)), Int(gradient(j, 3)))
            Next j
            
        End If
        
    Next i
    
End Function

Private Function cleanupDataStructures()

'clear all sheets
Application.DisplayAlerts = False

Dim sh As Worksheet

For Each sh In ActiveWorkbook.Worksheets
    If sh.Name = "SANKEYVB_SORTED" Or sh.Name = "SANKEYVB_MEMBERS" Then
        sh.Visible = xlSheetHidden    'cannot delete very hidden sheets
        sh.Delete
    End If
Next sh

Application.DisplayAlerts = True

End Function

Private Function drawSankey(data As Range, targetSheet As Worksheet, chartName As String, chartTop As Double, chartLeft As Double, chartWidth As Double, chartHeight As Double, skc As sankeyConfig)
    
    Dim ds As dataStructure
    
    Dim verticalSpacing As Double
    Dim blockSpaceForLevel As Double
    Dim blockWidth As Double
    Dim spacing As Double
    Dim boxWidth As Double    'the area occupied by the chart is smaller than the total area of the chart
    Dim boxHeight As Double    'the area occupied by the chart is smaller than the total area of the chart
    Dim top As Double
    Dim left As Double
    
    boxWidth = chartWidth - (skc.hMargin * 2)
    boxHeight = chartHeight - (skc.vMargin * 2)
    top = chartTop + skc.vMargin
    left = chartLeft + skc.hMargin
    
    'On Error GoTo cleanup
    
    'total blank space in each level
    verticalSpacing = skc.verticalBlankSpacePerc * boxHeight
    blockSpaceForLevel = boxHeight - verticalSpacing
    
    'get all necessary properties and build the support sheets
    ds = setupDataStructures(data, skc)
    targetSheet.Activate
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    'guess what the block width will be, considering a constant % of the box will be blank
    '@TODO: implement a second way of handling this by setting a fixed block width as a param.
    blockWidth = BOX_TO_BLANK_RATIO_H * boxWidth / ds.levelCount
    spacing = boxWidth * (1 - BOX_TO_BLANK_RATIO_H) / (ds.levelCount - 1)    'div.0 should never occur, doesn't make sense anyway to build a diagram with a single dim.
    
    Dim level As Integer
    Dim idx As Long
    Dim blockValue As Double
    Dim blockHeight As Double
    Dim blockShape As Shape
    Dim currentVerticalPos As Double
    Dim verticalSpacingSize As Double
    Dim shapeCounter As Long
    Dim nameCollection() As String    'we store all the names, in order to build a group for each level later
    Dim nameCollection2() As String
    Dim nameCollection3() As String
    Dim shapeGroup As Shape
    
    shapeCounter = 0
    
    ReDim nameCollection2(1 To ds.levelCount)
    
    'build the blocks, and prepare properties for each to handle connectors from multiple sources
    For level = 1 To ds.levelCount
        
        Debug.Print "Building blocks for level " & level & "(" & ds.elemCountForLevel(level) & " elements)"
        
    
        If ds.elemCountForLevel(level) = 1 Then
            ReDim nameCollection(1 To 3)
            blockValue = ds.totalValue
            currentVerticalPos = top + (verticalSpacing / 2)   'single member, center it
            shapeCounter = shapeCounter + 1
            Set blockShape = targetSheet.Shapes.AddShape(msoShapeRectangle, left + (level - 1) * (blockWidth + spacing), currentVerticalPos, blockWidth, blockSpaceForLevel)
            blockShape.Name = SHAPE_NAME_PREFIX & chartName & "_L" & level & "_BLOCK_1"
            nameCollection(1) = blockShape.Name
            
            setPropertyValue blockShape, "ELEMENT", ds.mbrSheet.Cells(1, level).value
            setPropertyValue blockShape, "VALUE", ds.totalValue
            setPropertyValue blockShape, "TYPE", "BLOCK"
            setPropertyValue blockShape, "LEVEL", CStr(level)
            setPropertyValue blockShape, "CHARTID", chartName
            
            'TODO: testing only
            blockShape.Fill.ForeColor.RGB = ds.mbrSheet.Cells(1, level).Interior.Color
            blockShape.Line.ForeColor.RGB = blockShape.Fill.ForeColor.RGB
            'blockShape.OnAction = "displayPropertiesForCaller"
            blockShape.OnAction = "clickHandler"
            blockShape.ZOrder msoSendToBack
            
            If level <= (ds.levelCount / 2) Then
                addLabelForBlock blockShape, ds.mbrSheet.Cells(1, level).value, blockValue, True, chartName, 0, spacing, skc.backgroundColor
            Else
                addLabelForBlock blockShape, ds.mbrSheet.Cells(1, level).value, blockValue, False, chartName, 0, spacing, skc.backgroundColor
            End If
            
            nameCollection(2) = blockShape.Name & "_LBL"
            
            Set blockShape = targetSheet.Shapes.AddShape(msoShapeRectangle, left + (level - 1) * (blockWidth + spacing), currentVerticalPos, 0, 0)
            blockShape.Name = SHAPE_NAME_PREFIX & chartName & "_L" & level & "_SPACER"
            blockShape.Visible = msoFalse
            blockShape.Fill.Transparency = 1
            blockShape.Line.Transparency = 1
            blockShape.ZOrder msoSendToBack
            nameCollection(3) = blockShape.Name
            
        Else
            ReDim nameCollection(1 To 2 * (ds.elemCountForLevel(level)))
            currentVerticalPos = top    'reset vertical position to the top
            verticalSpacingSize = verticalSpacing / (ds.elemCountForLevel(level) - 1)
            
            For idx = 1 To ds.elemCountForLevel(level)
                blockValue = WorksheetFunction.SumIf(ds.sortedDataTable.Columns(level), ds.mbrSheet.Cells(idx, level), ds.sortedDataTable.Columns(ds.levelCount + 1))
                blockHeight = blockSpaceForLevel * (blockValue / ds.totalValue)
                
                shapeCounter = shapeCounter + 1
                Set blockShape = targetSheet.Shapes.AddShape(msoShapeRectangle, left + (level - 1) * (blockWidth + spacing), currentVerticalPos, blockWidth, blockHeight)
                blockShape.Name = SHAPE_NAME_PREFIX & chartName & "_L" & level & "_BLOCK_" & idx
                nameCollection(idx) = blockShape.Name
                
                setPropertyValue blockShape, "ELEMENT", ds.mbrSheet.Cells(idx, level).value
                setPropertyValue blockShape, "VALUE", blockValue
                setPropertyValue blockShape, "TYPE", "BLOCK"
                setPropertyValue blockShape, "LEVEL", CStr(level)
                setPropertyValue blockShape, "CHARTID", chartName
                currentVerticalPos = currentVerticalPos + blockHeight + verticalSpacingSize
                
                'TODO: testing only
                blockShape.Fill.ForeColor.RGB = ds.mbrSheet.Cells(idx, level).Interior.Color
                blockShape.Line.ForeColor.RGB = blockShape.Fill.ForeColor.RGB
                'blockShape.OnAction = "displayPropertiesForCaller"
                blockShape.OnAction = "clickHandler"
                blockShape.ZOrder msoSendToBack
                
                If level <= (ds.levelCount / 2) Then
                    addLabelForBlock blockShape, ds.mbrSheet.Cells(idx, level).value, blockValue, True, chartName, verticalSpacingSize, spacing, skc.backgroundColor
                Else
                    addLabelForBlock blockShape, ds.mbrSheet.Cells(idx, level).value, blockValue, False, chartName, verticalSpacingSize, spacing, skc.backgroundColor
                End If
                
                nameCollection(idx + ds.elemCountForLevel(level)) = blockShape.Name & "_LBL"
                
            Next idx
        End If
        
        Set shapeGroup = targetSheet.Shapes.Range(nameCollection).group
        shapeGroup.Name = SHAPE_NAME_PREFIX & chartName & "_L" & level & "_BLOCK_GROUP"
        
        'put this in a group too
        nameCollection2(level) = shapeGroup.Name
    Next level
    
    Set shapeGroup = targetSheet.Shapes.Range(nameCollection2).group
    shapeGroup.Name = SHAPE_NAME_PREFIX & chartName & "_BLOCKS"
    
    
    
    Dim sourceShape As Shape
    Dim targetShape As Shape
    Dim connector As Shape
    Dim levelIdx As Integer
    
    ReDim nameCollection2(1 To ds.levelCount - 1)
    
    'build the connectors
    For level = 1 To ds.levelCount - 1    'there is no outgoing connection from the last level
        
        ReDim nameCollection(1 To ds.sortedDataTable.Rows.Count)
        
        Debug.Print "building connection between level " & level & " and level " & level + 1
        
        ds.sortedDataTable.Worksheet.Sort.SortFields.Clear
        ds.sortedDataTable.Worksheet.Sort.SortFields.Add ds.sortedDataTable.Columns(level), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ds.sortedDataTable.Worksheet.Sort.SortFields.Add ds.sortedDataTable.Columns(level + 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ds.sortedDataTable.Worksheet.Sort.SetRange ds.sortedDataTable
        ds.sortedDataTable.Worksheet.Sort.Header = xlNo
        ds.sortedDataTable.Worksheet.Sort.MatchCase = True
        ds.sortedDataTable.Worksheet.Sort.Orientation = xlTopToBottom
        ds.sortedDataTable.Worksheet.Sort.SortMethod = xlPinYin
        ds.sortedDataTable.Worksheet.Sort.Apply
        
        
        For idx = 1 To ds.sortedDataTable.Rows.Count
            
            If ds.sortedDataTable.Cells(idx, ds.levelCount + 1).value = 0 Then
                Debug.Print "suppressing connector because value is zero"
            Else
        
                Set sourceShape = getShapeForElementAndLevelInChart(targetSheet, chartName, level, ds.sortedDataTable.Cells(idx, level).value)
                Set targetShape = getShapeForElementAndLevelInChart(targetSheet, chartName, level + 1, ds.sortedDataTable.Cells(idx, level + 1))
                
                Set connector = connectShapesWithOffset(sourceShape, targetShape, False, False, ds.sortedDataTable.Cells(idx, ds.levelCount + 1).value * (blockSpaceForLevel / ds.totalValue), True, skc.blockColorMode)
                connector.Name = SHAPE_NAME_PREFIX & chartName & "_L" & level & "_L" & level + 1 & "_CONN_" & idx
                nameCollection(idx) = connector.Name
                
                connector.Line.Transparency = skc.baseTransparency
                connector.Fill.Transparency = skc.baseTransparency
                
                connector.ZOrder msoSendToBack
                
                'store information on ALL the members associated with this connector
                For levelIdx = 1 To ds.levelCount
                    setPropertyValue connector, "LEVEL" & levelIdx, ds.sortedDataTable.Cells(idx, levelIdx).value
                Next levelIdx
                
                setPropertyValue connector, "VALUE", ds.sortedDataTable.Cells(idx, ds.levelCount + 1).value
                setPropertyValue connector, "TYPE", "CONNECTOR"
                setPropertyValue connector, "CHARTID", chartName
                setPropertyValue connector, "FROMLEVEL", level
                setPropertyValue connector, "TOLEVEL", level + 1
                setPropertyValue connector, "SOURCEBLOCK", sourceShape.Name    'we store these to have a simple way to move from connector shape to block shape
                setPropertyValue connector, "TARGETBLOCK", targetShape.Name    'we store these to have a simple way to move from connector shape to block shape
                
                connector.OnAction = "clickHandler"
            End If
        Next idx
        
        If UBound(nameCollection) - LBound(nameCollection) = 0 Then
            Set shapeGroup = targetSheet.Shapes(nameCollection(LBound(nameCollection)))
        Else
            Set shapeGroup = targetSheet.Shapes.Range(nameCollection).group
        End If
        shapeGroup.Name = SHAPE_NAME_PREFIX & chartName & "_L" & level & "_L" & level + 1 & "_CONN_GROUP"
        
        nameCollection2(level) = shapeGroup.Name
        
    Next level
    
    'We fail here, if there is only one connector group
    If UBound(nameCollection2) - LBound(nameCollection2) = 0 Then
        Set shapeGroup = targetSheet.Shapes(nameCollection2(1))
        shapeGroup.Name = SHAPE_NAME_PREFIX & chartName & "_CONNECTORS"
    Else
        Set shapeGroup = targetSheet.Shapes.Range(nameCollection2).group
        shapeGroup.Name = SHAPE_NAME_PREFIX & chartName & "_CONNECTORS"
    End If

    
    'Create Background
    '@TODO: Cleanup
    Set shapeGroup = targetSheet.Shapes.AddShape(msoShapeRectangle, chartLeft, chartTop, chartWidth, chartHeight)
    shapeGroup.ZOrder msoSendToBack
    shapeGroup.Fill.ForeColor.RGB = RGB(skc.backgroundColor.r, skc.backgroundColor.g, skc.backgroundColor.b)
    shapeGroup.Line.ForeColor.RGB = RGB(0, 0, 0)
    shapeGroup.Name = SHAPE_NAME_PREFIX & chartName & "_BACKGROUND"
    setPropertyValue shapeGroup, "TYPE", "BACKGROUND"
    setPropertyValue shapeGroup, "CHARTID", chartName
    shapeGroup.OnAction = "clickHandler"
    
    Set shapeGroup = targetSheet.Shapes.Range(Array(SHAPE_NAME_PREFIX & chartName & "_BLOCKS", SHAPE_NAME_PREFIX & chartName & "_CONNECTORS", SHAPE_NAME_PREFIX & chartName & "_BACKGROUND")).group
    shapeGroup.Name = SHAPE_NAME_PREFIX & chartName & "_CHARTAREA"
    
    cleanupDataStructures
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Exit Function

cleanup:
    
    cleanupDataStructures
    clearAllSKeyShapes
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
End Function

Private Function getShapeForElementAndLevelInChart(ws As Worksheet, chartID As String, level As Integer, element As String) As Shape

    Dim group As Shape
    
    On Error Resume Next
    Set group = ws.Shapes(SHAPE_NAME_PREFIX & chartID & "_CHARTAREA")
    
    If group Is Nothing Then
        Set group = ws.Shapes(SHAPE_NAME_PREFIX & chartID & "_BLOCKS")
    End If
    
    If group Is Nothing Then
        Set group = ws.Shapes(SHAPE_NAME_PREFIX & chartID & "_L" & level & "_BLOCK_GROUP")
    End If
    
    If group Is Nothing Then
        Debug.Print "getShapeForElementAndLevelInChart could not find a block for element " & element & " on level " & level & " in chart " & chartID & "."
    End If
    
    On Error GoTo 0
    
    
    Dim sh As Shape
    
    For Each sh In group.GroupItems
        If getPropertyValue(sh, "ELEMENT", "") = element And getPropertyValue(sh, "LEVEL", "0") = level And getPropertyValue(sh, "TYPE", "") = "BLOCK" Then
            Set getShapeForElementAndLevelInChart = sh
            Exit Function
        End If
    Next sh

End Function

Sub displayPropertiesForCaller()

Dim sh As Shape
On Error GoTo noShape
Set sh = ActiveSheet.Shapes(Application.caller)
On Error GoTo 0

printAllPropertyValues sh, True

Exit Sub

noShape:
Debug.Print "Caller is not a shape."
End Sub


Sub clickHandler()

Dim caller As String
Dim shp As Shape

On Error GoTo invalidCaller

caller = Application.caller
Set shp = ActiveSheet.Shapes(caller)

On Error GoTo 0

'check if it's actually a sankey shape
If left(caller, Len(SHAPE_NAME_PREFIX)) = SHAPE_NAME_PREFIX Then
    
    Dim shapeType As String
    
    shapeType = getPropertyValue(shp, "TYPE")
    
    Select Case shapeType
        Case "BLOCK":
            handleBlockClick shp
        
        Case "CONNECTOR":
            handleConnectorClick shp
        
        Case "BACKGROUND":
            handleBackgroundClick shp
        
        Case "LABEL":
            Set shp = ActiveSheet.Shapes(left(shp.Name, Len(shp.Name) - 4))
            handleBlockClick shp
        Case Else:
        
    End Select
    
End If

Exit Sub

invalidCaller:
Debug.Print "Error: invalid caller for clickHandler() Sub."

End Sub

'@TODO: find an implementation that makes sense
Private Function handleConnectorClick(conn As Shape)
    
    Debug.Print "connector click not yet implemented"
    
End Function

Private Function handleBackgroundClick(background As Shape)
    
    Dim chartID As String
    
    chartID = getPropertyValue(background, "CHARTID")
    
    Dim groupShp As Shape
    Dim sh As Shape
    
    Set groupShp = ActiveSheet.Shapes(SHAPE_NAME_PREFIX & chartID & "_CHARTAREA")
    
    If groupShp Is Nothing Then
        Debug.Print "handleClickBlock: couldn't find the main group (_CHARTAREA) for the specified chart (" & chartID & ")"
        Exit Function
    End If
    
    Dim shapeType As String
    
    For Each sh In groupShp.GroupItems
        shapeType = getPropertyValue(sh, "TYPE", "")
        
        If shapeType = "CONNECTOR" Then
            sh.Fill.Transparency = BASE_TRANSPARENCY
            sh.Line.Transparency = BASE_TRANSPARENCY
        ElseIf shapeType = "BLOCK" Then
            sh.Fill.Transparency = LOW_TRANSPARENCY
            sh.Line.Transparency = LOW_TRANSPARENCY
        ElseIf shapeType = "LABEL" Then
            sh.Visible = msoTrue
        End If
    Next sh
    
End Function

Private Function handleBlockClick(block As Shape)
    Dim chartID As String
    Dim element As String
    Dim level As Integer
    Dim skc As sankeyConfig
    
    skc = getSankeyConfig
    
    chartID = getPropertyValue(block, "CHARTID")
    element = getPropertyValue(block, "ELEMENT", "")
    level = getPropertyValue(block, "LEVEL", "0")
    
    block.Fill.Transparency = skc.lowTransparency
    block.Line.Transparency = skc.highTransparency
    
    Dim groupShp As Shape
    
    Set groupShp = ActiveSheet.Shapes(SHAPE_NAME_PREFIX & chartID & "_CHARTAREA")
    
    If groupShp Is Nothing Then
        Debug.Print "handleClickBlock: couldn't find the main group (_CHARTAREA) for the specified chart (" & chartID & ")"
        Exit Function
    End If
    
    Dim sh As Shape
    Dim sourceBlock As Shape
    Dim targetBlock As Shape
    Dim blockType As String
    
    For Each sh In groupShp.GroupItems
        blockType = getPropertyValue(sh, "TYPE")
        
        If blockType = "BLOCK" Then
         sh.Fill.Transparency = skc.highTransparency
         sh.Line.Transparency = skc.highTransparency
        ElseIf blockType = "LABEL" Then
         sh.Visible = msoFalse
        End If
    Next sh
    
    For Each sh In groupShp.GroupItems
        If getPropertyValue(sh, "TYPE") = "CONNECTOR" Then
            
            If getPropertyValue(sh, "LEVEL" & level) = element Then
                sh.Fill.Transparency = skc.lowTransparency
                sh.Line.Transparency = skc.lowTransparency
                
                Set sourceBlock = sh.TopLeftCell.Worksheet.Shapes(getPropertyValue(sh, "SOURCEBLOCK", ""))
                Set targetBlock = sh.TopLeftCell.Worksheet.Shapes(getPropertyValue(sh, "TARGETBLOCK", ""))
                
                
                sourceBlock.Fill.Transparency = skc.lowTransparency
                sourceBlock.Line.Transparency = skc.lowTransparency
                
                targetBlock.Fill.Transparency = skc.lowTransparency
                targetBlock.Line.Transparency = skc.lowTransparency
                sh.TopLeftCell.Worksheet.Shapes(sourceBlock.Name & "_LBL").Visible = msoTrue
                sh.TopLeftCell.Worksheet.Shapes(targetBlock.Name & "_LBL").Visible = msoTrue
            Else
                sh.Fill.Transparency = skc.highTransparency
                sh.Line.Transparency = skc.highTransparency
            End If
        End If
    Next sh
    
    
End Function

Sub displayProperties()
Dim sh As Shape
On Error GoTo noShape
Set sh = ActiveSheet.Shapes(Selection.Name)
On Error GoTo 0
printAllPropertyValues sh, True
Exit Sub
noShape:
Debug.Print "Selection is not a shape."
End Sub

Sub clearAllSKeyShapes()

Sheet3.Shapes("BTN_DEL").Fill.ForeColor.RGB = RGB(250, 225, 225)

'this in normal cases should not be necessary, but we do it anyway in case the generation of the chart was stopped for some reason
cleanupDataStructures

Dim sh As Shape

For Each sh In ActiveSheet.Shapes
    
    If Len(sh.Name) > Len(SHAPE_NAME_PREFIX) Then
        If left(sh.Name, Len(SHAPE_NAME_PREFIX)) = SHAPE_NAME_PREFIX Then
            Debug.Print "Removing shape: " & sh.Name
            sh.Delete
        End If
    End If
Next sh

Sheet3.Shapes("BTN_DEL").Fill.ForeColor.RGB = RGB(255, 255, 255)
Sheet3.Shapes("BTN_CRE").Fill.ForeColor.RGB = RGB(255, 255, 255)

End Sub


Private Function getGradientBetweenColors(startColor As colorRGB, endColor As colorRGB, numSteps As Integer) As Variant
    
    Dim finalArray() As Integer
    
    If numSteps < 2 Then
        'getGradientBetweenColors = Nothing
        Debug.Print "Cannot generate a gradient of less than 2 colors."
        Exit Function
    End If
    
    ReDim finalArray(1 To numSteps, 1 To 3)
    
    Dim incrementR As Double
    Dim incrementG As Double
    Dim incrementB As Double
    
    incrementR = (endColor.r - startColor.r) / (numSteps - 1)
    incrementG = (endColor.g - startColor.g) / (numSteps - 1)
    incrementB = (endColor.b - startColor.b) / (numSteps - 1)


    Dim i As Integer
    
    For i = 1 To numSteps
        finalArray(i, 1) = startColor.r + incrementR * (i - 1)
        finalArray(i, 2) = startColor.g + incrementG * (i - 1)
        finalArray(i, 3) = startColor.b + incrementB * (i - 1)
    Next i
    
    getGradientBetweenColors = finalArray
    
End Function
