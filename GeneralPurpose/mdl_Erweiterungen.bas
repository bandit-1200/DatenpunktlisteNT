Attribute VB_Name = "mdl_Erweiterungen"
'Sloterweiterungen

Sub Erweiterungen_Eintragen()

'MsgBox Form_Belegung.List_Erweiterung.ListIndex



    With Form_Slots.List_Erweiterung
        'sMsg = "Use TextColumn" & vbNewLine & String(15, "-") & vbNewLine
        'sMsg = sMsg & .Value & vbTab & .Text
        'sMsg = sMsg & vbNewLine & vbNewLine
        'sMsg = sMsg & "No TextColumn" & vbNewLine & String(15, "-") & vbNewLine
        'sMsg = sMsg & .Value & vbTab & .List(.ListIndex, 1)
        sMsgT = .Value
        sMsgV = .List(.ListIndex, 1)
        
    End With
'MsgBox "Type: " & sMsgT & " Slots: " & sMsgV

If Form_Slots.Option_laufend = True Then

    letztezeileDB_E = Sheets("Erweiterungen").Cells(Rows.Count, 1).End(xlUp).Row
    
   ' MsgBox letztezeileDB_E
    
    If letztezeileDB_E >= 16 Then
        MsgBox "Maximale Anzahl der Erweiterungsmöglichkeiten erreicht!", vbInformation
    Exit Sub
    
    End If
    
    
    If (letztezeileDB_E = 1 And Sheets("Erweiterungen").Cells(1, 1) <> "") Or letztezeileDB_E > 1 Then
        letztezeileDB_E = letztezeileDB_E + 1
    End If
    
    
    
    
    Sheets("Erweiterungen").Cells(letztezeileDB_E, 1) = sMsgT
    Sheets("Erweiterungen").Cells(letztezeileDB_E, 2) = sMsgV

End If



'auf Markierung eintragen
If Form_Slots.Option_markierung = True Then
   uli = Form_Belegung.List_Erweiterung.ListIndex + 1
   If Sheets("Erweiterungen").Cells(uli, 1) <> "" Then
        Erweiterung_ueberschreiben = MsgBox("Erweiterung ersetzen?", vbYesNoCancel, "Löschen?")
        
   Else
        uli = Form_Belegung.List_Erweiterung.ListIndex + 1
        Sheets("Erweiterungen").Cells(uli, 1) = sMsgT
        Sheets("Erweiterungen").Cells(uli, 2) = sMsgV
   End If
   

    If Erweiterung_ueberschreiben = 6 Then
        uli = Form_Belegung.List_Erweiterung.ListIndex + 1
        Sheets("Erweiterungen").Cells(uli, 1) = sMsgT
        Sheets("Erweiterungen").Cells(uli, 2) = sMsgV
    Else
        Exit Sub
   
   End If
   
'Liste der Erweiterungen aktualisieren




End If




End Sub


Sub Erweiterungen_Refresh()

Form_Belegung.List_Erweiterung.Clear
letztezeileDB_Er = Sheets("Erweiterungen").Cells(Rows.Count, 1).End(xlUp).Row

For erw = 1 To letztezeileDB_Er
    SlotErweiterung = Sheets("Erweiterungen").Cells(erw, 1)
    Form_Belegung.List_Erweiterung.AddItem SlotErweiterung
Next


End Sub

Sub Erweiterungen_2_Modulliste()

max_Erweiterungen = Form_Belegung.List_Erweiterung.ListCount ' größe der Liste ermitteln
ListAnfang = 7

For maxE = 1 To max_Erweiterungen ' Schleife bis Ende der Liste durchlaufen
    maxEWM = Sheets("Erweiterungen").Cells(maxE, 2) ' Anzahl der Steckplätze lesen
        
        For iM = 1 To maxEWM
       ' MsgBox iM
            Sheets("Modulliste").Cells(ListAnfang, 6) = Sheets("Erweiterungen").Cells(maxE, 1) ' Erweiterung in Modulliste eintragen
            ListAnfang = ListAnfang + 1 'Zähler bei jedem durchlauf um N erhöhen
            
        Next
        
   ' ListAnfang = ListAnfang + maxEWM
    'MsgBox maxEWM
    
    

Next


End Sub

