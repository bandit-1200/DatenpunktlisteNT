Attribute VB_Name = "mdl_Projekt"


'Balttschutz - Test ******************************************************************

Sub CellLocker()
ActiveSheet.Unprotect
'Cells.Select
' unlock all the cells
'Selection.Locked = False
' next, select the cells (or range) that you want to make read only,
' here I used simply A1
Range("A:A,B:B,C:C,D:D,E:E,1:1,L:L,N:N,P:P,R:R,S:S").Select
' lock those cells
Selection.Locked = True
' now we need to protect the sheet to restrict access to the cells.
' I protected only the contents you can add whatever you want
ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False
End Sub


Sub Schutz_aufheben()

    ActiveSheet.Unprotect

End Sub

