VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Visio 
   Caption         =   "Importer.."
   ClientHeight    =   4320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9345
   OleObjectBlob   =   "Form_Visio.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Form_Visio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_exit_Click()
Unload Me
End Sub





Private Sub cmd_settings_Click()
On Error GoTo fehler
POP_import_frm.Show
On Error GoTo 0
Exit Sub
fehler:
MsgBox "Blattname prüfen!", vbInformation

End Sub

Private Sub ComboBox_Quelle_Click()
Sheets("Import_CFG").Cells(1, 1) = Me.ComboBox_Quelle.Text
'POP_import_001.Show
End Sub



Private Sub ComboBox_Ziel_Click()
Sheets("Import_CFG").Cells(1, 10) = Me.ComboBox_Ziel.Text
End Sub



Private Sub Command_Start_Import_Click()
Import_perUserForm
End Sub

Private Sub UserForm_Initialize()
Me.cmd_exit.Picture = Application.CommandBars.GetImageMso("MailDelete", 20, 20)
'Me.cmd_settings.Picture = Application.CommandBars.GetImageMso("MailMergeWizard", 20, 20)
Me.Image_Wizard.Picture = Application.CommandBars.GetImageMso("MailMergeWizard", 25, 25)
Me.Image_Import.Picture = Application.CommandBars.GetImageMso("XmlImport", 25, 25)
Me.Image_Export.Picture = Application.CommandBars.GetImageMso("TableExportMenu", 30, 30)


'TableExportMenu
'MailMergeWizard
'XmlImport

    q_sci = 0
    sc = Sheets.Count - 1
    For q_sc = 1 To sc
    'MsgBox sc
    
    

    
    ComboBox_Quelle.AddItem Sheets(q_sc).Name
    ComboBox_Ziel.AddItem Sheets(q_sc).Name
    
    Next
    
    
    
    If Sheets("Import_CFG").Cells(1, 1) <> "" Then
        
        ComboBox_Quelle.Text = Sheets("Import_CFG").Cells(1, 1)
        
    Else
        ComboBox_Quelle.Text = "Visio_Import"
        
    End If

    
        If Sheets("Import_CFG").Cells(1, 10) <> "" Then
        
        Me.ComboBox_Ziel.Text = Sheets("Import_CFG").Cells(1, 10)
        
    Else
        Me.ComboBox_Ziel.ListIndex = 0
        
    End If
        
   ' Me.ComboBox_Quelle.ListIndex = q_sci
   ' Me.ComboBox_Ziel.ListIndex = 0
    

    
End Sub
