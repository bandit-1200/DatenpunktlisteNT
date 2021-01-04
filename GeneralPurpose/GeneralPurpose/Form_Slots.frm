VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Slots 
   Caption         =   "UserForm2"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5010
   OleObjectBlob   =   "Form_Slots.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Form_Slots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmd_eintragen_Click()
Erweiterungen_Eintragen
Erweiterungen_Refresh
Erweiterungen_2_Modulliste

End Sub

Private Sub cmd_ok_Click()
Erweiterungen_Refresh
Unload Me
End Sub

Private Sub cmd_test_index_Click()

End Sub

Private Sub List_Erweiterung_DblClick(ByVal Cancel As MSForms.ReturnBoolean)


Erweiterungen_Eintragen
Erweiterungen_Refresh
Erweiterungen_2_Modulliste

End Sub

Private Sub UserForm_Initialize()
'SmartArtAddShapeSplitMenu
Me.cmd_ok.Picture = Application.CommandBars.GetImageMso("MailDelete", 20, 20)
Me.cmd_eintragen.Picture = Application.CommandBars.GetImageMso("SmartArtAddShapeSplitMenu", 20, 20)
letztezeileDB_E = Sheets("DB").Cells(Rows.Count, 5).End(xlUp).Row



'List_Erweiterung.ColumnCount = 2
For le = 1 To letztezeileDB_E
    ErweiterungName = Sheets("DB").Cells(le + 1, 5)
    ErweiterungAnzahl = Sheets("DB").Cells(le + 1, 6)
     
    Me.List_Erweiterung.ColumnHeads = True
    Me.List_Erweiterung.ColumnCount = 2
      
    Me.List_Erweiterung.RowSource = "DB!" & "E2:F10"
    Me.List_Erweiterung.Selected(1) = True
    
    
     
     
     
     
   ' With List_Erweiterung
   '     .AddItem ErweiterungName
   '     .List(.ListCount - 1, 1) = "ErweiterungName"
   '     .List(.ListCount - 1, 2) = "ErweiterungName"
        
   ' End With
Next
End Sub
