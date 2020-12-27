Attribute VB_Name = "mdl_nurWert"
Sub FormatfreiEinfuegen()
Attribute FormatfreiEinfuegen.VB_Description = "nur Wert eintragen"
Attribute FormatfreiEinfuegen.VB_ProcData.VB_Invoke_Func = "V\n14"
On Error Resume Next

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
    
On Error GoTo 0
End Sub
