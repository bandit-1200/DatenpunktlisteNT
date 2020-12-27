Attribute VB_Name = "mdl_BMKZ_sync"
Sub BMKZ_Sync()
    Application.ScreenUpdating = False

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim db As Workbook
    Set db = Application.Workbooks.Open("C:\BILZ\Projekte\V O R L A G E N\BMKZ_db.xlsm")
    'MsgBox "Derzeit nur auf fester Datenbank möglich! " & db.Path, vbInformation
    
    
    If Form_BMKZ_sync.Option_wb_db Then
    
        letzterEintrag = wb.Sheets("BMKZ-Belegung").UsedRange.SpecialCells(xlCellTypeLastCell).Row
    Else
    
        letzterEintrag = db.Sheets("BMKZ-Belegung").UsedRange.SpecialCells(xlCellTypeLastCell).Row
    End If
    
    
    
    'MsgBox letzterEintrag
    
    'Exit Sub
    
    wb.Sheets("BMKZ-Belegung").Activate
    db.Sheets(1).Activate
    
    
    
    
    
    
 'If Not Form_BMKZ_sync.Check_Fragen Then
    'ProzessBar*******************************************************************
        ProzessBarCSV.Show
        ProzessBarCSV.lbl_warten.Caption = "Bitte warten....Tranfer am laufen..."
        Dim Pausenlänge, Start, Ende, Gesamtdauer
    
        Pausenlänge = 0.1 ' Dauer festlegen.
        Start = Timer    ' Anfangszeit setzen.
        Do While Timer < Start + Pausenlänge
            DoEvents    ' Steuerung an andere Prozesse
                ' abgeben.
        Loop
        Ende = Timer    ' Ende festlegen.
        Gesamtdauer = Ende - Start    ' Gesamtdauer berechnen.
     '   MsgBox "Die Pause dauerte " & Gesamtdauer & " Sekunden"
    'ProzessBar*******************************************************************
'End If



   ' db.Sheets(1).Cells(1, 1) = "test"
For spalte = 1 To 200

    'ProzessBar*******************************************************************
    ProzessBarCSV.csvBar.Value = spalte / 200 * 100
    'ProzessBar*******************************************************************


    For zeile = 1 To letzterEintrag
        wb_BMKZ = wb.Sheets("BMKZ-Belegung").Cells(zeile, spalte)
        db_BMKZ = db.Sheets("BMKZ-Belegung").Cells(zeile, spalte)
        
        If Form_BMKZ_sync.Option_wb_db Then ueberschreiben = wb_BMKZ
        If Form_BMKZ_sync.Option_db_wb Then ueberschreiben = db_BMKZ
        
        'Debug.Print wb_BMKZ & " - " & db_BMKZ
        
        If wb_BMKZ <> "" Or db_BMKZ <> "" Then
            If wb_BMKZ = db_BMKZ Then
                'MsgBox "="
            Else
                Form_BMKZ_sync.lbl_wb_value = wb_BMKZ
                Form_BMKZ_sync.lbl_db_value = db_BMKZ
                    If Form_BMKZ_sync.Check_Fragen Then
                        Frage = MsgBox("Eintag erstellen? >" & ueberschreiben & " <", vbYesNoCancel)
                    Else
                        Frage = vbYes
                    End If
                    If Frage = vbCancel Then Exit Sub
                    If Frage = vbYes Then
                        If Form_BMKZ_sync.Option_wb_db Then
                            db.Sheets("BMKZ-Belegung").Cells(zeile, spalte) = wb_BMKZ
                        Else
                            wb.Sheets("BMKZ-Belegung").Cells(zeile, spalte) = db_BMKZ
                        End If
                    End If
            End If
        End If
    Next
Next
 'If Not Form_BMKZ_sync.Check_Fragen Then
    'ProzessBar*******************************************************************
    ProzessBarCSV.lbl_warten.Caption = "SYNC Fertig!..."
      Application.Wait (Now + TimeValue("0:00:01"))
    Unload ProzessBarCSV
    'ProzessBar*******************************************************************
'End If
Application.ScreenUpdating = True
End Sub
