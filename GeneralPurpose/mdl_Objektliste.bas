Attribute VB_Name = "mdl_Objektliste"
Public OBJECT_Name As String
Public plausiebel As Integer

Sub Open_Slotmanager()
'Schutz_aufheben
Sheets(1).Select
'If Application.ActiveSheet.Index = 1 Then

Form_Belegung.Show
'End If





'CellLocker
End Sub
Sub open_ImportCFG()

Form_Visio.Show

End Sub


Sub open_PromosNT_Menu()
Form_Menu.Show


End Sub

Sub open_Import_perUserForm()
I_von = Sheets("Import_CFG").Cells(1, 1)
I_nach = Sheets("Import_CFG").Cells(1, 10)
fragen = MsgBox("Sind die Import - Einstellungen vorher geprüft worden?" & vbCrLf & vbCrLf & "Von: " & I_von & " nach: " & I_nach & " Importieren?", vbQuestion & vbYesNo)

If fragen = vbYes Then
    Import_perUserForm
    MsgBox "Import Fertig!!", vbInformation
    
Else
    MsgBox "Importieren gestoppt!!", vbInformation


End If

End Sub


Sub ObjektListe_erstellen()


    Worksheets("Objektliste").Range("A1:C2000").Clear



    Sheets("Objektliste").Cells(1, 1) = "NAME"
    Sheets("Objektliste").Cells(1, 2) = "DMS-NAME"
    Sheets("Objektliste").Cells(1, 3) = "OBJECT"

OLZ = 2

letztezeile = Sheets(1).UsedRange.SpecialCells(xlCellTypeLastCell).Row

For o = 2 To letztezeile

    If Sheets(1).Cells(o, 13) <> "" Then
    
        OBJECT_Name = Sheets(1).Cells(o, 13)
        
        plausiebel = 0
        
        prüfung
        
        If plausiebel = 1 Then
        
        
            DMS_Name = Sheets(1).Cells(o, 12)
            NAME_ = Sheets(1).Cells(o, 6)
            
            OBJECT_Name = Replace(OBJECT_Name, "_", "")
         
            Sheets("Objektliste").Cells(OLZ, 3) = OBJECT_Name
            Sheets("Objektliste").Cells(OLZ, 2) = DMS_Name
            Sheets("Objektliste").Cells(OLZ, 1) = NAME_
            
            OLZ = OLZ + 1
            
            plausiebel = 0
            
        Else
            MsgBox OBJECT_Name & "  - ist kein Objekt!!"
        End If
        
        
    End If
    
Next


DblFind


End Sub

Sub prüfung()

'    letztezeile_DB = Sheets("DB2").Range("A1").SpecialCells(xlCellTypeLastCell).Address
    'MsgBox OBJECT_Name
    
 '       If Sheets("DB2").Cells(1, i) = OBJECT_Name Then
            plausiebel = 1
   '     End If
        
   ' Next
    

End Sub

Sub Test()
    If SheetExists("MEL01") = True Then
        MsgBox "Da!"
    Else
        MsgBox "Nicht da!"
    End If
End Sub

Sub csvSpeichern()
    'ChDrive ActiveWorkbook.Path 'Arbeitsverzeichnis auf Verzeichnis, in dem Excel- Datei liegt
    'Arbeitsverzeichnis = ActiveWorkbook.Path
     
    'BMO - Name suchen
    For n = 1 To 50
    
        sn = Sheets("DB2").Cells(1, n)
        sn = Replace(sn, "_", "")
        
        If SheetExists(sn) = True Then
            'MsgBox sn
            
            WBN = ThisWorkbook.Name
             
            exportFolder = ActiveWorkbook.Path
            exportfile = ActiveWorkbook.Path & "\" & sn & ".csv"
            
            'MsgBox exportfile
            
            Dateinummer = FreeFile
            Set TB = ThisWorkbook.Worksheets(sn)
            
            Open exportfile For Output As #Dateinummer
            
            For z = 1 To TB.UsedRange.Rows.Count
                'If Cells(Z, 1).Value = Text Then SL = 10 Else SL = 6
                If Cells(z, 1).Value = Text Then SL = 30 Else SL = 20
                    For s = 1 To SL
                        tmp = tmp & CStr(TB.Cells(z, s).Text) & ";"
                    Next s
                    tmp = Left(tmp, Len(tmp) - 1)
                    Print #Dateinummer, tmp
                    tmp = ""
            Next z
            Close #Dateinummer
        
        End If
    
    Next

Style = vbYesNo + vbQuestion

If exportFolder = "" Then
    MsgBox "Es wurde keine csv-Datei erstellt!" & Chr(13) & "Keine Objekte Angelegt!"
Else
    Frage = MsgBox("csv Dateien wurden im Verzeichnis:" & Chr(13) & exportFolder & " gespeicert... " & Chr(13) & Chr(13) & "Verzeichnis öffnen?", Style, "CSV - Export")
End If
    

If Frage = vbYes Then
    Shell "explorer.exe /e," & exportFolder, vbNormalFocus
Else
    'do nothing
End If

End Sub

Sub reset_inhalt()

'Module löschen
For ml = 2 To 600 Step 16
    Sheets(1).Cells(ml, 1) = "x"
Next

End Sub

Sub SlotListe_new()

Form_Belegung.ListBox_Slot_Modul.Clear


   For LS = 2 To 540 Step 16
    
    ModulType = Sheets(1).Cells(LS, 1)
    SlotNr = Sheets(1).Cells(LS, 3)
    
        With Form_Belegung.ListBox_Slot_Modul
            .ColumnCount = 2
            .ColumnWidths = "60;60"
            .AddItem
            .List(i, 0) = SlotNr
            .List(i, 1) = ModulType
            i = i + 1
        
        End With

    Next

End Sub


Sub IOs_anpassen()


Application.ScreenUpdating = False


    Sheets(1).Select
    Columns("R:R").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("tmp").Select
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        

    Sheets(1).Select
    Columns("P:P").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("tmp").Select
    Columns("C:C").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        


    Sheets(1).Select
    Columns("F:F").Select
    Selection.Copy
    Sheets("tmp").Select
    Columns("A:A").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
        
        Sheets("Arbeitsblatt").Select
     letztezeile_A_Blatt = Sheets("Arbeitsblatt").UsedRange.SpecialCells(xlCellTypeLastCell).Row
     letztezeile_tmp_Blatt = Sheets("tmp").UsedRange.SpecialCells(xlCellTypeLastCell).Row
     
     'MsgBox letztezeile_A_Blatt
     For IOs = 1 To letztezeile_A_Blatt
        Promos_S_AKS = Sheets("Arbeitsblatt").Cells(IOs, 2)
     
        For S_IOs = 1 To letztezeile_tmp_Blatt
           Promos_D_AKS = Sheets("tmp").Cells(S_IOs, 2)
             
           If Promos_S_AKS = Promos_D_AKS Then
                Sheets("Arbeitsblatt").Cells(IOs, 1) = Sheets("tmp").Cells(S_IOs, 3)
                
                Sheets(1).Select
                
            Sheets(1).Cells(S_IOs, 16).Select
            With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
    End With
                
                
                
                
           End If
         Next

     Next
    
        
     
     
     
     
     
     
     
     Sheets("Arbeitsblatt").Select
     
 Application.ScreenUpdating = True

End Sub



Sub pruefen()
    'Fehlende Objekte suchen
    
    letztezeile_obj = Sheets("Objektliste").UsedRange.SpecialCells(xlCellTypeLastCell).Row
    letztezeile_pro = Sheets("promos").UsedRange.SpecialCells(xlCellTypeLastCell).Row
    
    For i = 2 To letztezeile_obj
        AKS_obj = Sheets("Objektliste").Cells(i, 2)
            
            For ii = 1 To letztezeile_pro
                If Sheets("promos").Cells(ii, 2) = AKS_obj Then
                
                Sheets("Objektliste").Cells(i, 2).Select
                
                With Selection.Interior
                   .Pattern = xlSolid
                   .PatternColorIndex = xlAutomatic
                   .Color = 5296274
                   .TintAndShade = 0
                   .PatternTintAndShade = 0
                End With
                
                End If
            Next
    Next
    
                
    

End Sub


Sub BMKZ_cfg_open()

    Form_BMKZ_cfg.Show

End Sub

