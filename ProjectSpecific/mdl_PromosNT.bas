Attribute VB_Name = "mdl_PromosNT"
Sub PromosNT_Objekte_2DMS()
    Form_Namen.Frame_ExportOptionen.Visible = False
    Form_Namen.cmd_Namen_Plist.Caption = "Ausführen"
    Form_Namen.Caption = "Datenpunkte in PromosNT anlegen"
    Form_Namen.Show

End Sub


Sub Export_Objekte_PromosNT()
'Objekte nach PrpmosNT per json
  letztezeile_namen2PL = Sheets("Namen_cfg").Cells(Rows.Count, 1).End(xlUp).Row

  For n2pl = 1 To letztezeile_namen2PL
  
  
    suchen_N2PL = Sheets("Namen_cfg").Cells(n2pl, 1)
    frageObjekte = MsgBox("Soll der Inhalt der Tabelle: " & suchen_N2PL & " übertragen werden?", vbYesNoCancel, "Übertragen?")
    
    If frageObjekte = vbNo Then GoTo nochweiter
    If frageObjekte = vbCancel Then Exit Sub
    
    If suchen_N2PL = "" Then Exit Sub
    
      'Filter rücksetzen
      With Sheets(suchen_N2PL)
        If .FilterMode Then .ShowAllData
      End With

    BlattName = suchen_N2PL
    
    letztezeile = Sheets(BlattName).Cells(Rows.Count, 1).End(xlUp).Row

    'Json String erstellen
    For O_DMS = 1 To 15
        If Sheets(BlattName).Cells(1, O_DMS) = "DMS-NAME" Then GoTo weiter
        If Sheets(BlattName).Cells(1, O_DMS) = "" Then GoTo weiter
            For O_Value_Z = 2 To letztezeile
                O_Zusatz = Sheets(BlattName).Cells(1, O_DMS)
                O_Name = Sheets(BlattName).Cells(O_Value_Z, 1)
                O_Path = Sheets(BlattName).Cells(O_Value_Z, 2)
                'O_Objekt = Sheets("MES01").Cells(2, 3)
                O_Value = Sheets(BlattName).Cells(O_Value_Z, O_DMS)
                
                O_Path = O_Path & ":" & O_Zusatz
                
                JsonString = "{""whois"":""XLS"",""user"":""XLS"",""set"":[{""path"":""" & O_Path & """ ,""value"":""" & O_Value & """,""type"":""string"",""create"":true}]}"
                
                'Debug.Print JsonString
                DMS_Anfrage (JsonString)
            Next
weiter:
    
    Next
    
nochweiter:

Next

End Sub



