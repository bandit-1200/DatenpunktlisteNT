VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Belegung 
   Caption         =   "Datenpunktmanager"
   ClientHeight    =   10860
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15780
   OleObjectBlob   =   "Form_Belegung.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Form_Belegung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_abgleich_Click()
    IOs_anpassen
End Sub

Private Sub cmd_csv_Click()
  With Sheets(1)
    If .FilterMode Then .ShowAllData
  End With
    'ObjektListe_erstellen
    NeuesSheet
   
    csvSpeichern
End Sub

Private Sub cmd_del_SlotErweiterung_Click()

'Solot_Del = MsgBox("Markierte Erweiterung löschen?", vbYesNoCancel, "Löschen?")
'If Solot_Del = 6 Then
'    MsgBox List_Erweiterung.ListIndex
'    ei = List_Erweiterung.ListIndex
'    Sheets("Erweiterungen").Cells(ei, 1) = ""
'    Sheets("Erweiterungen").Cells(ei, 2) = ""
'End If
Application.ScreenUpdating = False

For i = 1 To 50
    Sheets("Erweiterungen").Cells(i, 1) = ""
    Sheets("Modulliste").Cells(i + 6, 6) = ""
    
Next

Erweiterungen_Refresh

Application.ScreenUpdating = True

End Sub

Private Sub cmd_exit_Click()
    Unload Me

End Sub

Private Sub cmd_format_Click()
    Trennlinie
End Sub

Private Sub cmd_Import_Click()

Form_Import.Show

End Sub

Private Sub cmd_Modul_reset_Click()
frage_module = MsgBox("Module löschen?", vbYesNoCancel, "Löschen?")
'MsgBox frage_module

If frage_module = 6 Then
    reset_inhalt
End If
UserForm_Initialize
End Sub

Private Sub cmd_Namen_Plist_Click()
    copy_namen
End Sub

Private Sub cmd_neu_Click()
Form_Slots.Show

End Sub

Private Sub cmd_objekte_Click()
  
  With Sheets(1)
    If .FilterMode Then .ShowAllData
  End With
    'ObjektListe_erstellen
    NeuesSheet
   
    csvSpeichern
    
End Sub

Private Sub cmd_transfer_Click()
    
   ' MsgBox List_Module.Text
    
    'If ListSlot.ListIndex <> "" Then
     '   ListSlot.ListIndex = 0
   ' End If
    
    Slot_auswahl = Me.ListSlot.ListIndex * 16 + 2
    On Error GoTo fehler
    
    Sheets(1).Cells(Slot_auswahl, 1) = List_Module.Text
   
   
    'ListSlot_Change
    'SlotListe_new
    ListSlot.ListIndex = ListSlot.ListIndex + 1
    Exit Sub:
fehler:
    MsgBox "Bitte eine Auswahl treffen!"
    
End Sub

Private Sub Command_Info_Click()
Erweiterungen_2_Modulliste

End Sub

Private Sub CommandButton1_Click()
    csvSpeichern
End Sub

Private Sub List_EM_Change()
Me.ListBox_Slot_Modul.ListIndex = Me.List_EM.ListIndex

End Sub

Private Sub List_Erweiterung_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 
 Form_Slots.Show
 
 



End Sub

Private Sub List_Module_Click()
lbl_Module = List_Module.Text


End Sub



Private Sub List_Module_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmd_transfer_Click
End Sub

Private Sub ListBox_Slot_Modul_Change()
    'Me.List_EM.ListIndex = Me.ListBox_Slot_Modul.ListIndex
End Sub

Private Sub ListBox_Slot_Modul_Click()
'Me.List_EM.ListIndex = Me.ListBox_Slot_Modul.ListIndex
Me.ListSlot.ListIndex = Me.ListBox_Slot_Modul.ListIndex
  
End Sub

Private Sub ListSlot_Change()
    'M_Pos = 2
   ' MsgBox ListSlot.ListIndex

    'lbl_Module
    
    
    
    Slot_auswahl = ListSlot.ListIndex * 16 + 2
    On Error Resume Next
    Modul_in_Slot = Sheets(1).Cells(Slot_auswahl, 1)

    lbl_Module = Modul_in_Slot

    SlotListe_new
    
End Sub



Private Sub ListSlot_Click()
 ListBox_Slot_Modul.ListIndex = ListSlot.ListIndex
End Sub

Private Sub UserForm_Initialize()

'Sheets(1).UsedRange.AutoFilter
'Sheets(1).UsedRange.AutoFilter Field:=1

    Trennlinie ' Format anpassen
    
    'Filter rücksetzen
  With Sheets(1)
    If .FilterMode Then .ShowAllData
  End With


Me.ListBox_Slot_Modul.Clear
Me.List_Module.Clear
Me.ListSlot.Clear
Me.List_EM.Clear
Me.cmd_exit.Picture = Application.CommandBars.GetImageMso("MailDelete", 20, 20)

Me.cmd_del_SlotErweiterung.Picture = Application.CommandBars.GetImageMso("RecordsDeleteRecord", 20, 20)
Me.cmd_Modul_reset.Picture = Application.CommandBars.GetImageMso("ClearMenu", 20, 20)
Me.cmd_transfer.Picture = Application.CommandBars.GetImageMso("PivotExpandField", 40, 40)


'DataFormAddRecord
Me.List_EM.RowSource = "Modulliste!" & "E7:G40"



Me.List_Erweiterung.Clear
letztezeileDB_Er = Sheets("Erweiterungen").Cells(Rows.Count, 1).End(xlUp).Row

For erw = 1 To letztezeileDB_Er
    SlotErweiterung = Sheets("Erweiterungen").Cells(erw, 1)
    Me.List_Erweiterung.AddItem SlotErweiterung
Next


'Me.List_Erweiterung.RowSource = "Erweiterungen!" & "A1:A17"


SlotListe_new


letztezeileDB_A = Sheets("DB").UsedRange.SpecialCells(xlCellTypeLastCell).Row


For lf = 0 To letztezeileDB_A

    ModulName = Sheets("DB").Cells(lf + 1, 1)
    List_Module.AddItem ModulName


Next


For LS = 0 To 100
    ListSlot.AddItem "Slot" & LS
    
    
Next



End Sub
