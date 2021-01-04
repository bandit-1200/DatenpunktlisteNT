VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Projekt 
   Caption         =   "Projekteinstellungen"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10785
   OleObjectBlob   =   "Form_Projekt.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Form_Projekt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_DDC_OK_Click()

Sheets(1).Cells(1, 7) = Me.txt_DDCName.Text
Sheets(1).Name = Me.txt_DDCName.Text
MsgBox "DDC Name wurde geändert", vbInformation

End Sub

Private Sub cmd_DDC_Type_Click()
 Sheets(1).Cells(1, 1).Value = Me.txt_DDC_Type.Text
End Sub

Private Sub cmd_exit_Click()
    Unload Me

End Sub

Private Sub cmd_ProjName_Click()
    ProjektName = Me.txt_ProjName.Text
    ProjektName = "Projekt: " & ProjektName
    
    Application.Worksheets("InbetriebnahmeProtokoll").lbl_projekt = ProjektName
    Application.Worksheets(1).lbl_projekt = ProjektName
End Sub



Private Sub UserForm_Initialize()
    Me.cmd_exit.Picture = Application.CommandBars.GetImageMso("MailDelete", 20, 20)
    Me.cmd_DDC_OK.Picture = Application.CommandBars.GetImageMso("FrameSaveCurrentAs", 20, 20)
    Me.cmd_ProjName.Picture = Application.CommandBars.GetImageMso("FrameSaveCurrentAs", 20, 20)
    Me.cmd_DDC_Type.Picture = Application.CommandBars.GetImageMso("FrameSaveCurrentAs", 20, 20)
    'AcceptInvitation
    
    Me.txt_DDCName.Text = Sheets(1).Name
    Me.txt_DDC_Type.Text = Sheets(1).Cells(1, 1).Value
    

    ProjektName = Application.Worksheets("InbetriebnahmeProtokoll").lbl_projekt
    
    ProjektName = Replace(ProjektName, "Projekt: ", "")
    Me.txt_ProjName = ProjektName



End Sub
