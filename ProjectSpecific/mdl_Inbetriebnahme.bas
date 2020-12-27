Attribute VB_Name = "mdl_Inbetriebnahme"
Sub InbetriebnahmeProtokoll()

Form_Namen.Caption = "Datenpunktlisten auswählen"
Form_Namen.cmd_Namen_Plist.Visible = False
Form_Namen.Option_json.Visible = False
Form_Namen.Option_Plist.Visible = False

Form_Namen.cmd_exit.Caption = "OK"
Form_Namen.Show


End Sub

Sub InbetriebnahmeProtokoll_erstellen()


Application.ScreenUpdating = False
Sheets("InbetriebnahmeProtokoll").Select
    Range("A11:L957").Select
    ActiveWindow.SmallScroll Down:=60
    Rows("11:1010").Select
    Range("A1010").Activate
    Selection.Delete Shift:=xlUp
    Range("A11").Select


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



ibnp = 10

Sheets("InbetriebnahmeProtokoll").Cells(1, 1) = "Inbetreibnahmeprotokoll"
Sheets("InbetriebnahmeProtokoll").Cells(ibnp, 1) = "Anlagenteil"
Sheets("InbetriebnahmeProtokoll").Cells(ibnp, 2) = "Prüfdatum"
Sheets("InbetriebnahmeProtokoll").Cells(ibnp, 3) = "Prüfer"
Sheets("InbetriebnahmeProtokoll").Cells(ibnp, 4) = "Bemerkung"
Sheets("InbetriebnahmeProtokoll").Cells(ibnp, 5) = "E-Schema"
Sheets("InbetriebnahmeProtokoll").Cells(ibnp, 8) = "AKST1"
Sheets("InbetriebnahmeProtokoll").Cells(ibnp, 9) = "AKST2"
Sheets("InbetriebnahmeProtokoll").Cells(ibnp, 10) = "AKST3"
Sheets("InbetriebnahmeProtokoll").Cells(ibnp, 11) = "AKST4"
Sheets("InbetriebnahmeProtokoll").Cells(ibnp, 12) = "AKST5"




ibnp = ibnp + 1

  letztezeile_namenCFG = Sheets("Namen_cfg").UsedRange.SpecialCells(xlCellTypeLastCell).Row

  For n2pl = 1 To letztezeile_namenCFG
   BerichtDatenpunkte = Sheets("Namen_cfg").Cells(n2pl, 1)
   
    
    
    'Filter rücksetzen
    With Sheets(BerichtDatenpunkte)
      If .FilterMode Then .ShowAllData
    End With


MsgBox "Blatt: " & BerichtDatenpunkte & " ausgewählt!"

For VEi = 2 To 577


    'ProzessBar*******************************************************************
    ProzessBarCSV.csvBar.Value = VEi / 577 * 100
    'ProzessBar*******************************************************************


    If Sheets(BerichtDatenpunkte).Cells(VEi, 6) <> "" Then
    
        
      
        IBNP_Adresse = Sheets(BerichtDatenpunkte).Cells(VEi, 5)
        IBNP_Name = Sheets(BerichtDatenpunkte).Cells(VEi, 6)
        
        IBNP_AKST1 = Sheets(BerichtDatenpunkte).Cells(VEi, 7)
        IBNP_AKST2 = Sheets(BerichtDatenpunkte).Cells(VEi, 8)
        IBNP_AKST3 = Sheets(BerichtDatenpunkte).Cells(VEi, 9)
        IBNP_AKST4 = Sheets(BerichtDatenpunkte).Cells(VEi, 10)
        IBNP_AKST5 = Sheets(BerichtDatenpunkte).Cells(VEi, 11)
                
        IBNP_PruefName = Sheets(BerichtDatenpunkte).Cells(VEi, 24)
        IBNP_Datum = Sheets(BerichtDatenpunkte).Cells(VEi, 25)
        IBNP_Bemerkung = Sheets(BerichtDatenpunkte).Cells(VEi, 26)
        
        IBNP_E_SCHEMA = Sheets(BerichtDatenpunkte).Cells(VEi, 33)
        
        
        AKS_komplett = IBNP_AKST1 & IBNP_AKST2 & IBNP_AKST3 & IBNP_AKST4 & IBNP_AKST5
        AKS_Referenz = IBNP_AKST1 & IBNP_AKST2 & IBNP_AKST3
        AKS_Zusatz = IBNP_AKST4


        Sheets("InbetriebnahmeProtokoll").Cells(ibnp, 1) = IBNP_Name
        Sheets("InbetriebnahmeProtokoll").Cells(ibnp, 2) = IBNP_Datum
        Sheets("InbetriebnahmeProtokoll").Cells(ibnp, 3) = IBNP_PruefName
        Sheets("InbetriebnahmeProtokoll").Cells(ibnp, 4) = IBNP_Bemerkung
        Sheets("InbetriebnahmeProtokoll").Cells(ibnp, 5) = IBNP_E_SCHEMA
        Sheets("InbetriebnahmeProtokoll").Cells(ibnp, 8) = IBNP_AKST1
        Sheets("InbetriebnahmeProtokoll").Cells(ibnp, 9) = IBNP_AKST2
        Sheets("InbetriebnahmeProtokoll").Cells(ibnp, 10) = IBNP_AKST3
        Sheets("InbetriebnahmeProtokoll").Cells(ibnp, 11) = IBNP_AKST4
        Sheets("InbetriebnahmeProtokoll").Cells(ibnp, 12) = IBNP_AKST5




        ibnp = ibnp + 1

    End If




Next

Rows("10:10").Select
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "Calibri"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Cells.Select
    Cells.EntireColumn.AutoFit
        Rows("1:9").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
        Sheets("InbetriebnahmeProtokoll").Range("A11:L1100").Select
    ActiveWorkbook.Worksheets("InbetriebnahmeProtokoll").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("InbetriebnahmeProtokoll").Sort.SortFields.Add2 Key _
        :=Range("H11:H175"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("InbetriebnahmeProtokoll").Sort.SortFields.Add2 Key _
        :=Range("I11:I175"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("InbetriebnahmeProtokoll").Sort.SortFields.Add2 Key _
        :=Range("J11:J175"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("InbetriebnahmeProtokoll").Sort.SortFields.Add2 Key _
        :=Range("K11:K175"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("InbetriebnahmeProtokoll").Sort.SortFields.Add2 Key _
        :=Range("L11:L175"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("InbetriebnahmeProtokoll").Sort
        .SetRange Range("A11:L175")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("InbetriebnahmeProtokoll").Cells(10, 1).Select
Next

Application.ScreenUpdating = True

'ProzessBar*******************************************************************
ProzessBarCSV.lbl_warten.Caption = "Import Fertig!..."
  Application.Wait (Now + TimeValue("0:00:02"))
Unload ProzessBarCSV
'ProzessBar*******************************************************************
 
End Sub
Sub Projekt_Beschriftung()
    ProjektName = Application.Worksheets("InbetriebnahmeProtokoll").lbl_projekt
    
    ProjektName = Replace(ProjektName, "Projekt: ", "")
    
    
    ProjektName = InputBox("Projektname?", "Projektname...", ProjektName)
    ProjektName = "Projekt: " & ProjektName
    
     Application.Worksheets("InbetriebnahmeProtokoll").lbl_projekt = ProjektName
     Application.Worksheets(1).lbl_projekt = ProjektName

End Sub
