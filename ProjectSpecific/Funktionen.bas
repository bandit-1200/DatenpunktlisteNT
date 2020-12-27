Attribute VB_Name = "Funktionen"
Function SheetExists(ByVal SheetName As String) As Boolean
Dim i As Integer
    For i = 1 To Sheets.Count
        If Sheets(i).Name = SheetName Then SheetExists = True: Exit Function
    Next i
    SheetExists = False
End Function





