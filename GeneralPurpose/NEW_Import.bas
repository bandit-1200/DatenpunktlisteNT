Attribute VB_Name = "NEW_Import"
Sub Import_perUserForm() 'Importieren



'ProzessBar*******************************************************************
    ProzessBarCSV.Show
    ProzessBarCSV.lbl_warten.Caption = "Bitte warten....Import..."
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







quell_Blatt_BlattName = Sheets("Import_CFG").Cells(1, 1)
quell_Blatt_Adresse = Sheets("Import_CFG").Cells(2, 3)

quell_Blatt_EIN_Name = Sheets("Import_CFG").Cells(3, 1)
quell_Blatt_Name = Sheets("Import_CFG").Cells(3, 3)

quell_Blatt_EIN_AKS = Sheets("Import_CFG").Cells(4, 1)

quell_Blatt_EIN_AKS_T1 = Sheets("Import_CFG").Cells(5, 1)
quell_Blatt_POS_AKS_T1 = Sheets("Import_CFG").Cells(5, 3)

quell_Blatt_EIN_AKS_T2 = Sheets("Import_CFG").Cells(6, 1)
quell_Blatt_POS_AKS_T2 = Sheets("Import_CFG").Cells(6, 3)

quell_Blatt_EIN_AKS_T3 = Sheets("Import_CFG").Cells(7, 1)
quell_Blatt_POS_AKS_T3 = Sheets("Import_CFG").Cells(7, 3)

quell_Blatt_EIN_AKS_T4 = Sheets("Import_CFG").Cells(8, 1)
quell_Blatt_POS_AKS_T4 = Sheets("Import_CFG").Cells(8, 3)

quell_Blatt_EIN_AKS_T5 = Sheets("Import_CFG").Cells(9, 1)
quell_Blatt_POS_AKS_T5 = Sheets("Import_CFG").Cells(9, 3)

quell_Blatt_EIN_AKS_T6 = Sheets("Import_CFG").Cells(10, 1)
quell_Blatt_POS_AKS_T6 = Sheets("Import_CFG").Cells(10, 3)


Ziel_Blatt_BlattName = Sheets("Import_CFG").Cells(1, 10)
Ziel_Blatt_Adresse = Sheets("Import_CFG").Cells(2, 12)
Ziel_Blatt_Name = Sheets("Import_CFG").Cells(3, 12)
Ziel_Blatt_POS_AKS_T1 = Sheets("Import_CFG").Cells(5, 12)
Ziel_Blatt_POS_AKS_T2 = Sheets("Import_CFG").Cells(6, 12)
Ziel_Blatt_POS_AKS_T3 = Sheets("Import_CFG").Cells(7, 12)
Ziel_Blatt_POS_AKS_T4 = Sheets("Import_CFG").Cells(8, 12)
Ziel_Blatt_POS_AKS_T5 = Sheets("Import_CFG").Cells(9, 12)
Ziel_Blatt_POS_AKS_T6 = Sheets("Import_CFG").Cells(10, 12)


  With Sheets(quell_Blatt_BlattName)
    If .FilterMode Then .ShowAllData
  End With
  
  
    With Sheets(Ziel_Blatt_BlattName)
    If .FilterMode Then .ShowAllData
  End With
If quell_Blatt_EIN_AKS = False And quell_Blatt_EIN_Name = False Then

MsgBox "ÄÄÄ - Es ist nix zum Importieren ausgewählt!", vbInformation

Exit Sub
End If




letzteZeileQuelle = Sheets(quell_Blatt_BlattName).Cells(Rows.Count, quell_Blatt_Name).End(xlUp).Row


    For ZeileZielBlatt = 2 To 600
    'ProzessBar*******************************************************************
    ProzessBarCSV.csvBar.Value = ZeileZielBlatt / 600 * 100
    'ProzessBar*******************************************************************
    
    
        ZielBlattAdresse = CStr(Sheets(Ziel_Blatt_BlattName).Cells(ZeileZielBlatt, Ziel_Blatt_Adresse))
        
        For ZeileQuellBlattZAE = 2 To letzteZeileQuelle
            QuellBlattAdresse = CStr(Sheets(quell_Blatt_BlattName).Cells(ZeileQuellBlattZAE, quell_Blatt_Adresse))
            If QuellBlattAdresse = "" Then QuellBlattAdresse = 10000
            
            'Application.StatusBar = ZielBlattAdresse

            'Application.StatusBar = "VISIO: " & QuellBlattAdresse & " - " & "SAIA: " & ZielBlattAdresse
            
            
            If QuellBlattAdresse = ZielBlattAdresse Then
                
      
                quellBlattName = Sheets(quell_Blatt_BlattName).Cells(ZeileQuellBlattZAE, quell_Blatt_Name)
            
                QuelleAKS_T1 = Sheets(quell_Blatt_BlattName).Cells(ZeileQuellBlattZAE, quell_Blatt_POS_AKS_T1)
                QuelleAKS_T2 = Sheets(quell_Blatt_BlattName).Cells(ZeileQuellBlattZAE, quell_Blatt_POS_AKS_T2)
                QuelleAKS_T3 = Sheets(quell_Blatt_BlattName).Cells(ZeileQuellBlattZAE, quell_Blatt_POS_AKS_T3)
                QuelleAKS_T4 = Sheets(quell_Blatt_BlattName).Cells(ZeileQuellBlattZAE, quell_Blatt_POS_AKS_T4)
                QuelleAKS_T5 = Sheets(quell_Blatt_BlattName).Cells(ZeileQuellBlattZAE, quell_Blatt_POS_AKS_T5)
                QuelleAKS_T6 = Sheets(quell_Blatt_BlattName).Cells(ZeileQuellBlattZAE, quell_Blatt_POS_AKS_T6)
                
                
                If quell_Blatt_EIN_Name = True Then
                    Sheets(Ziel_Blatt_BlattName).Cells(ZeileZielBlatt, Ziel_Blatt_Name) = quellBlattName
                End If
                
                
                If quell_Blatt_EIN_AKS = True Then
                    If quell_Blatt_EIN_AKS_T1 = True Then
                        Sheets(Ziel_Blatt_BlattName).Cells(ZeileZielBlatt, Ziel_Blatt_POS_AKS_T1) = QuelleAKS_T1
                    End If
                    
                    If quell_Blatt_EIN_AKS_T2 = True Then
                        Sheets(Ziel_Blatt_BlattName).Cells(ZeileZielBlatt, Ziel_Blatt_POS_AKS_T2) = QuelleAKS_T2
                    End If
                    
                    If quell_Blatt_EIN_AKS_T3 = True Then
                        Sheets(Ziel_Blatt_BlattName).Cells(ZeileZielBlatt, Ziel_Blatt_POS_AKS_T3) = QuelleAKS_T3
                    End If
                    
                    If quell_Blatt_EIN_AKS_T4 = True Then
                        Sheets(Ziel_Blatt_BlattName).Cells(ZeileZielBlatt, Ziel_Blatt_POS_AKS_T4) = QuelleAKS_T4
                    End If
                    
                    If quell_Blatt_EIN_AKS_T5 = True Then
                        Sheets(Ziel_Blatt_BlattName).Cells(ZeileZielBlatt, Ziel_Blatt_POS_AKS_T5) = QuelleAKS_T5
                    End If
                    
                    If quell_Blatt_EIN_AKS_T6 = True Then
                        Sheets(Ziel_Blatt_BlattName).Cells(ZeileZielBlatt, Ziel_Blatt_POS_AKS_T6) = QuelleAKS_T6
                    End If
                End If
                
            End If
            
        Next
    
    Next
    
'ProzessBar*******************************************************************
ProzessBarCSV.lbl_warten.Caption = "Import Fertig!..."
  Application.Wait (Now + TimeValue("0:00:03"))
Unload ProzessBarCSV
'ProzessBar*******************************************************************
End Sub
