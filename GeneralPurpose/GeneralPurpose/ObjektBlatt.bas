Attribute VB_Name = "ObjektBlatt"
Sub NeuesSheet()
'** Neues benanntes Tabellenblatt einfügen
'** einfügen als letztes Blatt
 Application.ScreenUpdating = False
 
'** Dimensionierung der Variablen
Dim blatt As Object
Dim BlattName As String
Dim bolFlg As Boolean

 
 
'Worksheets("Objektliste").Range("A1:C2000").Clear

letztezeile = Sheets(1).UsedRange.SpecialCells(xlCellTypeLastCell).Row

For o = 2 To letztezeile
    bolFlg = False

    OBJECT_Name = Sheets(1).Cells(o, 13)
    OBJECT_Name2 = Replace(OBJECT_Name, "_", "")
    
    If OBJECT_Name2 <> "" Then
    
        '** Blattname festlegen
        BlattName = OBJECT_Name2
         
        '** Prüfen, ob das Blatt, welches eingefügt werden soll bereits vorhanden ist
        '** Nur einfügen, wenn Blatt noch nicht vorhanden ist
        For Each blatt In Sheets
          If blatt.Name = BlattName Then bolFlg = True
        Next blatt
         
        '** Blatt nur einfügen, wenn noch nicht vorhanden
        If bolFlg = False Then
          With ThisWorkbook
            .Sheets.Add after:=Sheets(Worksheets.Count)
            .ActiveSheet.Name = OBJECT_Name2
          End With
        End If
         
         Sheets(OBJECT_Name2).Cells.ClearContents
         
        Sheets(OBJECT_Name2).Cells(1, 1) = "NAME"
        Sheets(OBJECT_Name2).Cells(1, 2) = "DMS-NAME"
        Sheets(OBJECT_Name2).Cells(1, 3) = "OBJECT"
        
        'Zusatz eintragen
        For zz = 1 To 100
           If Sheets("DB2").Cells(1, zz) = OBJECT_Name Then
               For s = 2 To 20
                'Sheets("DB2").Cells(s, zz).Select
                 DMSZusatz = Sheets("DB2").Cells(s, zz)
                 Sheets(OBJECT_Name2).Cells(1, s + 2) = DMSZusatz
                Next
               
           End If
           
        Next
 
      End If
     
 
Next
 objekte_anlegen
 Application.ScreenUpdating = True
End Sub

Sub objekte_anlegen()
Application.ScreenUpdating = False
letztezeile = Sheets(1).UsedRange.SpecialCells(xlCellTypeLastCell).Row


For i = 2 To letztezeile
    
    If Sheets(1).Cells(i, 13) <> "" Then
        OBJECT_Name = Sheets(1).Cells(i, 13)
        OBJECT_Name2 = Replace(OBJECT_Name, "_", "")
        letztezeile_OBJECT_Name2 = Sheets(OBJECT_Name2).UsedRange.SpecialCells(xlCellTypeLastCell).Row
        
        Z_Name = Sheets(1).Cells(i, 6)
        Z_AKS = Sheets(1).Cells(i, 12)
        Z_IO = Sheets(1).Cells(i, 16)
        Z_Zusatz = Sheets(1).Cells(i, 17)
        
        AKS_Fund = False
        
        For AKSCheck = 2 To letztezeile_OBJECT_Name2 + 1
            ZP_AKS = Sheets(OBJECT_Name2).Cells(AKSCheck, 2)
               
            If ZP_AKS = Z_AKS And Z_AKS <> "" Then
            
                For ZusatzCheck = 4 To 50
                    ZP_Zusatz = Sheets(OBJECT_Name2).Cells(1, ZusatzCheck)
                        
                    If Z_Zusatz = ZP_Zusatz And ZP_Zusatz <> "" Then
                        Sheets(OBJECT_Name2).Cells(AKSCheck, ZusatzCheck) = Z_IO
                    End If
                Next
                AKS_Fund = True
            End If
            
        Next
        
        If AKS_Fund = False Then
                    Sheets(OBJECT_Name2).Cells(letztezeile_OBJECT_Name2 + 1, 1) = Z_Name
                    Sheets(OBJECT_Name2).Cells(letztezeile_OBJECT_Name2 + 1, 2) = Z_AKS
                    Sheets(OBJECT_Name2).Cells(letztezeile_OBJECT_Name2 + 1, 3) = OBJECT_Name2
                    
                    For ZusatzCheck = 4 To 50
                        ZP_Zusatz = Sheets(OBJECT_Name2).Cells(1, ZusatzCheck)
                            
                        If Z_Zusatz = ZP_Zusatz And ZP_Zusatz <> "" Then
                            Sheets(OBJECT_Name2).Cells(letztezeile_OBJECT_Name2 + 1, ZusatzCheck) = Z_IO
                        End If
                    Next

        End If
        

       
    End If




Next



'Lücken füllen
For lu = 1 To 50

    l_sn = Sheets("DB2").Cells(1, lu)
    l_sn2 = Replace(l_sn, "_", "")
     
   ' MsgBox l_sn2

    
    If SheetExists(l_sn2) = True Then
        letztezeile_l_sn2 = Sheets(l_sn2).UsedRange.SpecialCells(xlCellTypeLastCell).Row
        
        For LS = 1 To 50
            
            For LZ = 1 To letztezeile_l_sn2
                If Sheets(l_sn2).Cells(1, LS) <> "" Then
                    Luecke = Sheets(l_sn2).Cells(LZ, LS)
                    L_Zusatz = Sheets(l_sn2).Cells(1, LS)
                    
                    If Luecke = "" Then
                        For L_DB = 1 To 50
                           L_Objekt = Sheets("DB2").Cells(1, L_DB)
                           
                           If L_Objekt = l_sn Then
                                'MsgBox L_Objekt
                                For zusatz = 2 To 20
                                    F_Zusatz = Sheets("DB2").Cells(zusatz, L_DB + 1)
                                    If F_Zusatz <> "" Then
                                        Sheets(l_sn2).Cells(LZ, LS) = F_Zusatz
                                    End If
                                    
                                Next
                                
                           End If
                           
                           'L_Zusatz
                           'L_Fueller
                           
                        Next
                        
                    End If
                    
                
                End If
             Next
             
             
                
            
        Next
     
    End If


Next



 letztezeile_Farbe = Sheets(1).UsedRange.SpecialCells(xlCellTypeLastCell).Row

For farb = 2 To letztezeile_Farbe
    If Sheets(1).Cells(farb, 17) <> "" Then
        Sheets(1).Cells(farb, 6).Font.ColorIndex = xlAutomatic
    Else
        Sheets(1).Cells(farb, 6).Font.ColorIndex = 3
    End If
    

Next


 Application.ScreenUpdating = True

End Sub



Sub DblFind()
    Sheets("Objektliste").Select
    
    Dim lngZeile As Long
    Application.ScreenUpdating = False
    For lngZeile = 1 To Cells(65536, 1).End(xlUp).Row
        If Cells(lngZeile, 2) = Cells(lngZeile + 1, 2) Then
            With Range(Cells(lngZeile, 2), Cells(lngZeile + 1, 2))
                .Font.Bold = True
                .Font.ColorIndex = 3
            End With
            MsgBox "Doppelte Einträge in Objektliste gefunden!", vbCritical, "Objektliste"
            
            
        End If
    Next
    Application.ScreenUpdating = True
End Sub














