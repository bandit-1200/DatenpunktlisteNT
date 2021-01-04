VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_BMKZ_cfg 
   Caption         =   "BMKZ CFG"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7845
   OleObjectBlob   =   "Form_BMKZ_cfg.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Form_BMKZ_cfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_exit_Click()
Unload Me

End Sub

Private Sub ComboBMKZ_Change()
Sheets("Import_CFG").Cells(2, 30) = Me.ComboBMKZ.ListIndex + 1


End Sub



Private Sub UserForm_Initialize()
Me.cmd_exit.Picture = Application.CommandBars.GetImageMso("MailDelete", 20, 20)

    Me.ComboBMKZ.Clear
    
    Me.ComboBMKZ.AddItem "A"
    Me.ComboBMKZ.AddItem "B"
    Me.ComboBMKZ.AddItem "C"
    Me.ComboBMKZ.AddItem "D"
    Me.ComboBMKZ.AddItem "E"
    Me.ComboBMKZ.AddItem "F"
    Me.ComboBMKZ.AddItem "G"
    Me.ComboBMKZ.AddItem "H"
    Me.ComboBMKZ.AddItem "I"
    Me.ComboBMKZ.AddItem "J"
    Me.ComboBMKZ.AddItem "K"
    Me.ComboBMKZ.AddItem "L"
    Me.ComboBMKZ.AddItem "M"
    Me.ComboBMKZ.AddItem "N"
    Me.ComboBMKZ.AddItem "O"
    Me.ComboBMKZ.AddItem "P"
    Me.ComboBMKZ.AddItem "Q"
    Me.ComboBMKZ.AddItem "R"
    Me.ComboBMKZ.AddItem "S"
    Me.ComboBMKZ.AddItem "T"
    Me.ComboBMKZ.AddItem "U"
    Me.ComboBMKZ.AddItem "V"
    Me.ComboBMKZ.AddItem "W"
    Me.ComboBMKZ.AddItem "X"
    Me.ComboBMKZ.AddItem "Y"
    Me.ComboBMKZ.AddItem "Z"
    
    
    If Sheets("Import_CFG").Cells(2, 30).Value < 1 Then
        Me.ComboBMKZ.ListIndex = 8
    Else
       Me.ComboBMKZ.ListIndex = Sheets("Import_CFG").Cells(2, 30) - 1
    End If
    
End Sub
