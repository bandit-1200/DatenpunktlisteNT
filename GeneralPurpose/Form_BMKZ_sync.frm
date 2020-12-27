VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_BMKZ_sync 
   Caption         =   "BMKZ Sync"
   ClientHeight    =   6720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8010
   OleObjectBlob   =   "Form_BMKZ_sync.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Form_BMKZ_sync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_exit_Click()
Unload Me

End Sub



Private Sub cmd_sync_Click()
    BMKZ_Sync
End Sub


Private Sub UserForm_Initialize()

    Me.cmd_sync.Picture = Application.CommandBars.GetImageMso("Repeat", 20, 20)
    Me.cmd_exit.Picture = Application.CommandBars.GetImageMso("MailDelete", 20, 20)


End Sub
