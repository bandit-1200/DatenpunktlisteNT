Attribute VB_Name = "mdl_Namen"
Sub copy_namen()
  
  
  letztezeile_namen2PL = Sheets("Namen_cfg").Cells(Rows.Count, 1).End(xlUp).Row

  For n2pl = 1 To letztezeile_namen2PL
  

    
    suchen_N2PL = Sheets("Namen_cfg").Cells(n2pl, 1)
    MsgBox suchen_N2PL
    
    If suchen_N2PL = "" Then Exit Sub
    
      'Filter rücksetzen
      With Sheets(suchen_N2PL)
        If .FilterMode Then .ShowAllData
      End With
      
      
    Sheets(suchen_N2PL).Select
    
    '    Range("F10,F:F,S:S").Select
    '    Range("S1").Activate
        Columns("F:S").Select
        Selection.Copy
        Sheets("tmp").Select
        ActiveWindow.SmallScroll Down:=-9
        Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
     namen_suchen
        
Next

End Sub

Sub namen_suchen()

'letztezeile = Sheets("tmp").UsedRange.SpecialCells(xlCellTypeLastCell).Row
letztezeile_tmp = Sheets("tmp").Cells(Rows.Count, 1).End(xlUp).Row
letztezeile_Plist = Sheets("Name_aus_PList").Cells(Rows.Count, 2).End(xlUp).Row


'MsgBox letztezeile_Plist
'spalte 1 = Name
'spalte 14 = AKS

For pl = 1 To letztezeile_Plist

    aks_Plist = Sheets("Name_aus_PList").Cells(pl, 2)

    For tmp = 1 To letztezeile_tmp
    
        aks_tmp = Sheets("tmp").Cells(tmp, 14)
        name_tmp = Sheets("tmp").Cells(tmp, 1)
        
        If aks_tmp = aks_Plist Then
        
        Sheets("Name_aus_PList").Cells(pl, 1) = name_tmp
        
        End If
        
    
    Next

Next


Sheets("Name_aus_PList").Select



End Sub
