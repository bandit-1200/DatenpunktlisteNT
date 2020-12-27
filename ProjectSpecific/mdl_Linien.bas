Attribute VB_Name = "mdl_Linien"


Sub Trennlinie()
'Filter_alles_an
Application.ScreenUpdating = False
Sheets(1).UsedRange.AutoFilter Field:=1
For X = 1 To 600 Step 16
On Error Resume Next
    Sheets(1).Rows(X).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
Next
Application.ScreenUpdating = True

Sheets(1).Rows(1).Select


    Cells.Select
    Selection.EntireColumn.Hidden = False


End Sub



