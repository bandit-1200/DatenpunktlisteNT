Attribute VB_Name = "AUTO_AKS"
Sub AKS_AutoFill()

'ProzessBar*******************************************************************
    ProzessBarCSV.Show
    ProzessBarCSV.lbl_warten.Caption = "Bitte warten....Import-BMKZ..."
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









asi = Application.ActiveSheet.Index
Sinx = Sheets("Import_CFG").Cells(2, 30)
'MsgBox asi

For BMKZ_insert = 2 To 577
fund = 0
    'ProzessBar*******************************************************************
    ProzessBarCSV.csvBar.Value = BMKZ_insert / 577 * 100
    'ProzessBar*******************************************************************
    
    
    
    BMKZ_T0 = Sheets(asi).Cells(BMKZ_insert, 7)
    BMKZ_T1 = Sheets(asi).Cells(BMKZ_insert, 30)
    BMKZ_T2 = Sheets(asi).Cells(BMKZ_insert, 31)
    
    'MsgBox BMKZ_T0
    'MsgBox BMKZ_T1
    'MsgBox BMKZ_T2
    
    If BMKZ_T2 <> "" Then
     
        For suchen = 1 To 200
            S_BMKZ_T1 = Sheets("BMKZ-Belegung").Cells(1, suchen)
            
            If S_BMKZ_T1 = BMKZ_T1 Then
                
                For suche_t2 = 2 To 50
                    
                    S_BMKZ_T2 = Sheets("BMKZ-Belegung").Cells(suche_t2, suchen)
                    
                    If fund = 0 Then
                        
                        If S_BMKZ_T2 = BMKZ_T2 Then
                                BMKZ_Erg_T1 = Sheets("BMKZ-Belegung").Cells(suche_t2, suchen + 1)
                                Sheets(asi).Cells(BMKZ_insert, Sinx) = BMKZ_Erg_T1
                                fund = 1
                                'MsgBox "FUND " & BMKZ_Erg_T1
                        End If
                    End If
                    
                        
                Next
            
            End If
            
        Next
    End If
Next

'BMKZ-Belegung

'ProzessBar*******************************************************************
ProzessBarCSV.lbl_warten.Caption = "BMKZ-Import Fertig!..."
  Application.Wait (Now + TimeValue("0:00:02"))
Unload ProzessBarCSV
'ProzessBar*******************************************************************
End Sub
