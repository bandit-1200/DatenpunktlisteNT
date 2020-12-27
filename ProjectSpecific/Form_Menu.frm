VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Menu 
   Caption         =   "MENU"
   ClientHeight    =   2520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7740
   OleObjectBlob   =   "Form_Menu.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Form_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_exit_Click()
Unload Me

End Sub

Private Sub cmd_Import_Click()
Form_Visio.Show

End Sub

Private Sub cmd_Namen_json_Click()
    Form_Namen.Show
    
End Sub

Private Sub cmd_Namen_Plist_Click()
Form_Namen.Show

    'copy_namen
End Sub


Private Sub cmd_objekte_Click()
  
  With Sheets(1)
    If .FilterMode Then .ShowAllData
  End With
    'ObjektListe_erstellen
    NeuesSheet
   
    csvSpeichern
    
End Sub


Private Sub cmd_Slot_Click()
Form_Belegung.Show

End Sub

Private Sub UserForm_Activate()
    'Filter rücksetzen
  With Sheets(1)
    If .FilterMode Then .ShowAllData
  End With

End Sub



Private Sub UserForm_Initialize()
Me.cmd_exit.Picture = Application.CommandBars.GetImageMso("MailDelete", 20, 20)
Me.cmd_Import.Picture = Application.CommandBars.GetImageMso("MailDelete", 20, 20)
Me.cmd_Namen_json.Picture = Application.CommandBars.GetImageMso("TableExportTableToSharePointList", 20, 20)
    
    'Filter rücksetzen
  With Sheets(1)
    If .FilterMode Then .ShowAllData
  End With

End Sub
