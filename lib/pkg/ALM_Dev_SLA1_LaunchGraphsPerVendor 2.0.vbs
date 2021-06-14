Sub Macro1()
'
' Macro1 Macro
'

'
iTopRow = 1
iCurrenSheetName = ActiveSheet.Name

Dim oShapes() As Shape

Dim sRngHeader As String: sRngHeader = "A" & iTopRow & ":E" & iTopRow
Dim sRngHeaderAbs As String: sRngHeaderAbs = "$A$" & iTopRow & ":$E$" & iTopRow

Dim sRngThisRow As String
Dim sRngThisRowAbs As String

imaxChart = 0

For Index = iTopRow + 1 To ActiveSheet.UsedRange.Rows.Count

sRngThisRow = "A" & Index & ":E" & Index
sRngThisRowAbs = "$A$" & Index & ":$E$" & Index

    Range(sRngHeader & "," & sRngThisRow).Select
    
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.SetSourceData Source:=Range( _
        iCurrenSheetName & "!" & sRngThisRowAbs & "," & iCurrenSheetName & "!" & sRngHeaderAbs)
        
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).ApplyDataLabels
    
    ReDim Preserve oShapes(imaxChart) As Shape
    Set oShapes(imaxChart) = ActiveSheet.Shapes(Replace(ActiveChart.Name, ActiveSheet.Name & " ", ""))
    imaxChart = imaxChart + 1
    
    If imaxChart = 1 Then
        Call FirstShape(oShapes(0))
    End If
    
Next

Dim bHorizontally As Boolean
Dim oShareSource As Shape

For i = 1 To UBound(oShapes)
'0,1
'2,3
'4,5

iEven = i Mod 2

    If iEven <> 0 Then
        bHorizontally = True
        Set oShareSource = oShapes(i - 1)
        'iFirstShapeIndexInThisRow = i - 1
    Else
        bHorizontally = False
        Set oShareSource = oShapes(i - 2)
        'iFirstShapeIndexInThisRow = i - 1
    End If
    
        Call Macro2(oShareSource, oShapes(i), bHorizontally)
Next

End Sub
Sub FirstShape(oShapeSource As Shape)

iRowCount = 6
iTop = 87 + (iRowCount * 15)
iLeft = 20

    oShapeSource.Left = iLeft
    oShapeSource.Top = iTop
    
End Sub
Sub Macro2(oShapeSource As Shape, oShapeTarget As Shape, bHorizontally As Boolean)
'
' Macro2 Macro
'

'
iColGap = 20
iRowGap = 20

    With oShapeSource
        iRefTop = .Top
        iRefLeft = .Left
        iRefWidth = .Width
        iRefHeight = .Height
    End With

If bHorizontally = True Then
    'Horizontal
    oShapeTarget.Left = iRefLeft + iRefWidth + iColGap
    oShapeTarget.Top = iRefTop
Else

    'Vertically
    oShapeTarget.Left = iRefLeft
    oShapeTarget.Top = iRefTop + iRefHeight + iRowGap
End If

'    ActiveSheet.Shapes("Chart 11").Left = iLeft
'    ActiveSheet.Shapes("Chart 11").Top = iTop
'
'    iwidth = ActiveSheet.Shapes("Chart 11").Width
'    iheight = ActiveSheet.Shapes("Chart 11").Height
'
'    'Horizontal
'    ActiveSheet.Shapes("Chart 10").Left = iLeft + iwidth + iColGap
'    ActiveSheet.Shapes("Chart 10").Top = iTop
'
'    'Vertically
'    'ActiveSheet.Shapes("Chart 10").Left = iLeft
'    'ActiveSheet.Shapes("Chart 10").Top = iTop + iheight + iRowGap
End Sub
