VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Import 
   Caption         =   "Import Maske"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8040
   OleObjectBlob   =   "Form_Import.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Form_Import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_ausKabelzugliste_Click()
Form_Warten.Show

ImportKabelzug
Unload Form_Warten
End Sub

Private Sub cmd_exit_Click()

Unload Me

End Sub

Private Sub cmd_visioImport_Click()
Form_Visio.Show

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
Me.cmd_exit.Picture = Application.CommandBars.GetImageMso("MailDelete", 20, 20)
End Sub
