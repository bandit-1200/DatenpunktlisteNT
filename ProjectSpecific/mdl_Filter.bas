Attribute VB_Name = "mdl_Filter"
Sub Filter_alles_an()
Attribute Filter_alles_an.VB_ProcData.VB_Invoke_Func = " \n14"
    Cells.Select
    Selection.EntireColumn.Hidden = False
    Cells(1, 1).Select
End Sub
Sub Filter_BMKZ()
    Filter_alles_an
    'Columns("L:AC").Select
    Range("L:AC,AF:AF").Select
    Selection.EntireColumn.Hidden = True
    Range("A1:B1").Select
End Sub
Sub Filter_PG5()
    Filter_alles_an
    Range("L:M,O:O").Select
    Range("O1").Activate
    Range("L:M,O:O,P:P,Q:Q,S:S").Select
    Range("S1").Activate
    Selection.EntireColumn.Hidden = True
    Range("A1:B1").Select
End Sub

Sub Filter_PromosNT()
    Filter_alles_an
    Columns("M:M").Select
    Range("M:M,N:N,O:O,Q:R").Select
    Range("Q1").Activate
    Selection.EntireColumn.Hidden = True
    Range("A1:B1").Select
End Sub
Sub Filter_Inbetriebnahme()
    Filter_alles_an
    Range("S:W,Q:Q,P:P,O:O,M:M,L:L,AC:AC,AB:AB,AA:AA,AD:AD,AE:AE").Select
    Selection.EntireColumn.Hidden = True
    Range("A1").Activate

End Sub
Sub Filter_PromosObjekte()
    Filter_alles_an
    Range("N:N,O:O,P:P,R:R,S:S").Select
    Range("S1").Activate
    ActiveWindow.SmallScroll ToRight:=8
    Range("N:N,O:O,P:P,R:R,S:S,T:T,U:U,V:V,W:W,X:AF").Select
    Range("X1").Activate
    Selection.EntireColumn.Hidden = True
End Sub

