Attribute VB_Name = "mdl_Balttmanager"

Public Sub Test()
Dim R               As Integer
Dim G               As Integer
Dim B               As Integer
    'GetRGB ActiveCell.Interior.Color, R, G, B
    GetRGB "116737", R, G, B
    
    
    
    MsgBox "Rot: " & R & vbCrLf & _
            "Grün: " & G & vbCrLf & _
            "Blau: " & B
End Sub



Sub GetRGB(RGB As Long, ByRef Red As Integer, _
        ByRef Green As Integer, ByRef Blue As Integer)
    Red = RGB And 255
    Green = RGB \ 256 And 255
    Blue = RGB \ 256 ^ 2 And 255
End Sub


