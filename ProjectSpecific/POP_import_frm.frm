VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} POP_import_frm 
   Caption         =   "Import config"
   ClientHeight    =   13080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17100
   OleObjectBlob   =   "POP_import_frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "POP_import_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckBox_AKS_Import_Click()
If Me.CheckBox_AKS_Import = True Then
    Me.ComboBox_AKS.Enabled = True
    Me.lbl_AKS.Enabled = True
    Me.Label_Spalte_AKS.Enabled = True
    Sheets("Import_CFG").Cells(4, 1) = True

Else
    Me.ComboBox_AKS.Enabled = False
    Me.lbl_AKS.Enabled = False
    Me.Label_Spalte_AKS.Enabled = False
    Sheets("Import_CFG").Cells(4, 1) = False
End If
refrech
End Sub

Private Sub CheckBox_AKS_T1_Click()
If Me.CheckBox_AKS_T1 = True Then
    Me.lbl_AKS_T1.Enabled = True
    Sheets("Import_CFG").Cells(5, 1) = True

Else
    Me.lbl_AKS_T1.Enabled = False
    Sheets("Import_CFG").Cells(5, 1) = False
End If

refrech
End Sub
Private Sub CheckBox_AKS_T2_Click()
If Me.CheckBox_AKS_T2 = True Then
    Me.lbl_AKS_T2.Enabled = True
    Sheets("Import_CFG").Cells(6, 1) = True

Else
    Me.lbl_AKS_T2.Enabled = False
    Sheets("Import_CFG").Cells(6, 1) = False
End If
refrech
End Sub
Private Sub CheckBox_AKS_T3_Click()
If Me.CheckBox_AKS_T3 = True Then
    Me.lbl_AKS_T3.Enabled = True
    Sheets("Import_CFG").Cells(7, 1) = True

Else
    Me.lbl_AKS_T3.Enabled = False
    Sheets("Import_CFG").Cells(7, 1) = False
End If
refrech
End Sub
Private Sub CheckBox_AKS_T4_Click()
If Me.CheckBox_AKS_T4 = True Then
    Me.lbl_AKS_T4.Enabled = True
    Sheets("Import_CFG").Cells(8, 1) = True

Else
    Me.lbl_AKS_T4.Enabled = False
    Sheets("Import_CFG").Cells(8, 1) = False
End If
refrech
End Sub
Private Sub CheckBox_AKS_T5_Click()
If Me.CheckBox_AKS_T5 = True Then
    Me.lbl_AKS_T5.Enabled = True
    Sheets("Import_CFG").Cells(9, 1) = True

Else
    Me.lbl_AKS_T5.Enabled = False
    Sheets("Import_CFG").Cells(9, 1) = False
End If
refrech
End Sub
Private Sub CheckBox_AKS_T6_Click()
If Me.CheckBox_AKS_T6 = True Then
    Me.lbl_AKS_T6.Enabled = True
    Sheets("Import_CFG").Cells(10, 1) = True

Else
    Me.lbl_AKS_T6.Enabled = False
    Sheets("Import_CFG").Cells(10, 1) = False
End If
refrech
End Sub

Private Sub CheckBox_Name_Import_Click()
If Me.CheckBox_Name_Import = True Then
    Me.ComboBox_Name.Enabled = True
    Me.lbl_Name.Enabled = True
    Me.Label_Spalte_Name.Enabled = True
    Sheets("Import_CFG").Cells(3, 1) = True
Else
    Me.ComboBox_Name.Enabled = False
    Me.lbl_Name.Enabled = False
    Me.Label_Spalte_Name.Enabled = False
    Sheets("Import_CFG").Cells(3, 1) = False
End If

refrech
End Sub

Private Sub cmd_exit_Click()
Unload Me

End Sub




Private Sub ComboBox_AKS_Click()

   ' Sheets("Import_CFG").Cells(4, 2) = Me.ComboBox_AKS.ListIndex ' + 1 'AKS
    
   ' Sheets("Import_CFG").Cells(4, 3) = Sheets("Import_CFG").Cells(4, 2) + 1
    
   ' Import_Blatt_SName = Sheets("Import_CFG").Cells(1, 1)
   ' Import_Blatt_AKS = Sheets("Import_CFG").Cells(4, 2)
    
   ' Me.lbl_AKS = Sheets(Import_Blatt_SName).Cells(1, Import_Blatt_AKS)
   ' refrech
End Sub

Private Sub ComboBox_Name_Click()
Sheets("Import_CFG").Cells(3, 2) = Me.ComboBox_Name.ListIndex '+ 1 'Name
Sheets("Import_CFG").Cells(3, 3) = Sheets("Import_CFG").Cells(3, 2) + 1

Import_Blatt_SName = Sheets("Import_CFG").Cells(1, 1)
Import_Blatt_Name = Sheets("Import_CFG").Cells(3, 3)

Me.lbl_Name = Sheets(Import_Blatt_SName).Cells(1, Import_Blatt_Name)



refrech
End Sub

Private Sub ComboBox_Adresse_Click()


Sheets("Import_CFG").Cells(2, 2) = Me.ComboBox_Adresse.ListIndex '+ 1 'Adresse
Sheets("Import_CFG").Cells(2, 3) = Sheets("Import_CFG").Cells(2, 2) + 1
Import_Blatt_SName = Sheets("Import_CFG").Cells(1, 1)
Import_Blatt_Adresse = Sheets("Import_CFG").Cells(2, 3)

Me.lbl_Adresse = Sheets(Import_Blatt_SName).Cells(1, Import_Blatt_Adresse)

refrech
End Sub




Private Sub ScrollBar_AKS_Click()

End Sub



Private Sub ComboBox_Quelle_AKS_T1_Change()
Sheets("Import_CFG").Cells(5, 2) = Me.ComboBox_Quelle_AKS_T1.ListIndex '+ 1 'QuelleAdresse
Sheets("Import_CFG").Cells(5, 3) = Sheets("Import_CFG").Cells(5, 2) + 1
Quelle_Blatt_SName = Sheets("Import_CFG").Cells(1, 1)
Quelle_Blatt_QuelleAdresse = Sheets("Import_CFG").Cells(5, 3)
If Quelle_Blatt_QuelleAdresse > 1 Then Quelle_Blatt_QuelleAdresse = 1
Me.CheckBox_AKS_T1.Caption = Sheets(Quelle_Blatt_SName).Cells(1, Quelle_Blatt_QuelleAdresse)

refrech
End Sub
Private Sub ComboBox_Quelle_AKS_T2_Change()
Sheets("Import_CFG").Cells(6, 2) = Me.ComboBox_Quelle_AKS_T2.ListIndex '+ 1 'QuelleAdresse
Sheets("Import_CFG").Cells(6, 3) = Sheets("Import_CFG").Cells(6, 2) + 1
Quelle_Blatt_SName = Sheets("Import_CFG").Cells(1, 1)
Quelle_Blatt_QuelleAdresse = Sheets("Import_CFG").Cells(6, 3)
If Quelle_Blatt_QuelleAdresse > 1 Then Quelle_Blatt_QuelleAdresse = 1
Me.CheckBox_AKS_T2.Caption = Sheets(Quelle_Blatt_SName).Cells(1, Quelle_Blatt_QuelleAdresse)

refrech
End Sub
Private Sub ComboBox_Quelle_AKS_T3_Change()
Sheets("Import_CFG").Cells(7, 2) = Me.ComboBox_Quelle_AKS_T3.ListIndex '+ 1 'QuelleAdresse
Sheets("Import_CFG").Cells(7, 3) = Sheets("Import_CFG").Cells(7, 2) + 1
Quelle_Blatt_SName = Sheets("Import_CFG").Cells(1, 1)
Quelle_Blatt_QuelleAdresse = Sheets("Import_CFG").Cells(7, 3)
If Quelle_Blatt_QuelleAdresse > 1 Then Quelle_Blatt_QuelleAdresse = 1
Me.CheckBox_AKS_T3.Caption = Sheets(Quelle_Blatt_SName).Cells(1, Quelle_Blatt_QuelleAdresse)

refrech
End Sub
Private Sub ComboBox_Quelle_AKS_T4_Change()
Sheets("Import_CFG").Cells(8, 2) = Me.ComboBox_Quelle_AKS_T4.ListIndex '+ 1 'QuelleAdresse
Sheets("Import_CFG").Cells(8, 3) = Sheets("Import_CFG").Cells(8, 2) + 1
Quelle_Blatt_SName = Sheets("Import_CFG").Cells(1, 1)
Quelle_Blatt_QuelleAdresse = Sheets("Import_CFG").Cells(8, 3)
If Quelle_Blatt_QuelleAdresse > 1 Then Quelle_Blatt_QuelleAdresse = 1
Me.CheckBox_AKS_T4.Caption = Sheets(Quelle_Blatt_SName).Cells(1, Quelle_Blatt_QuelleAdresse)

refrech
End Sub
Private Sub ComboBox_Quelle_AKS_T5_Change()
Sheets("Import_CFG").Cells(9, 2) = Me.ComboBox_Quelle_AKS_T5.ListIndex '+ 1 'QuelleAdresse
Sheets("Import_CFG").Cells(9, 3) = Sheets("Import_CFG").Cells(9, 2) + 1
Quelle_Blatt_SName = Sheets("Import_CFG").Cells(1, 1)
Quelle_Blatt_QuelleAdresse = Sheets("Import_CFG").Cells(9, 3)
If Quelle_Blatt_QuelleAdresse > 1 Then Quelle_Blatt_QuelleAdresse = 1
Me.CheckBox_AKS_T5.Caption = Sheets(Quelle_Blatt_SName).Cells(1, Quelle_Blatt_QuelleAdresse)

refrech
End Sub
Private Sub ComboBox_Quelle_AKS_T6_Change()
Sheets("Import_CFG").Cells(10, 2) = Me.ComboBox_Quelle_AKS_T6.ListIndex '+ 1 'QuelleAdresse
Sheets("Import_CFG").Cells(10, 3) = Sheets("Import_CFG").Cells(10, 2) + 1
Quelle_Blatt_SName = Sheets("Import_CFG").Cells(1, 1)
Quelle_Blatt_QuelleAdresse = Sheets("Import_CFG").Cells(10, 3)
If Quelle_Blatt_QuelleAdresse > 1 Then Quelle_Blatt_QuelleAdresse = 1
Me.CheckBox_AKS_T6.Caption = Sheets(Quelle_Blatt_SName).Cells(1, Quelle_Blatt_QuelleAdresse)

refrech
End Sub


Private Sub ComboBox_Zeile_Change()
Sheets("Import_CFG").Cells(11, 2) = Me.ComboBox_Zeile.Text

End Sub

Private Sub ComboBox_Ziel_AKS_T1_Change()
Sheets("Import_CFG").Cells(5, 11) = Me.ComboBox_Ziel_AKS_T1.ListIndex '+ 1 'ZielAdresse
Sheets("Import_CFG").Cells(5, 12) = Sheets("Import_CFG").Cells(5, 11) + 1
Ziel_Blatt_SName = Sheets("Import_CFG").Cells(1, 10)
Ziel_Blatt_ZielAdresse = Sheets("Import_CFG").Cells(5, 12)
On Error Resume Next
Me.lbl_Ziel_AKS_T1 = Sheets(Ziel_Blatt_SName).Cells(2000, Ziel_Blatt_ZielAdresse)

refrech



End Sub
Private Sub ComboBox_Ziel_AKS_T2_Change()
Sheets("Import_CFG").Cells(6, 11) = Me.ComboBox_Ziel_AKS_T2.ListIndex '+ 1 'ZielAdresse
Sheets("Import_CFG").Cells(6, 12) = Sheets("Import_CFG").Cells(6, 11) + 1
Ziel_Blatt_SName = Sheets("Import_CFG").Cells(1, 10)
Ziel_Blatt_ZielAdresse = Sheets("Import_CFG").Cells(6, 12)
On Error Resume Next
Me.lbl_Ziel_AKS_T2 = Sheets(Ziel_Blatt_SName).Cells(2000, Ziel_Blatt_ZielAdresse)

refrech



End Sub
Private Sub ComboBox_Ziel_AKS_T3_Change()
Sheets("Import_CFG").Cells(7, 11) = Me.ComboBox_Ziel_AKS_T3.ListIndex '+ 1 'ZielAdresse
Sheets("Import_CFG").Cells(7, 12) = Sheets("Import_CFG").Cells(7, 11) + 1
Ziel_Blatt_SName = Sheets("Import_CFG").Cells(1, 10)
Ziel_Blatt_ZielAdresse = Sheets("Import_CFG").Cells(7, 12)
On Error Resume Next
Me.lbl_Ziel_AKS_T3 = Sheets(Ziel_Blatt_SName).Cells(2000, Ziel_Blatt_ZielAdresse)

refrech



End Sub
Private Sub ComboBox_Ziel_AKS_T4_Change()
Sheets("Import_CFG").Cells(8, 11) = Me.ComboBox_Ziel_AKS_T4.ListIndex '+ 1 'ZielAdresse
Sheets("Import_CFG").Cells(8, 12) = Sheets("Import_CFG").Cells(8, 11) + 1
Ziel_Blatt_SName = Sheets("Import_CFG").Cells(1, 10)
Ziel_Blatt_ZielAdresse = Sheets("Import_CFG").Cells(8, 12)
On Error Resume Next
Me.lbl_Ziel_AKS_T4 = Sheets(Ziel_Blatt_SName).Cells(2000, Ziel_Blatt_ZielAdresse)

refrech



End Sub
Private Sub ComboBox_Ziel_AKS_T5_Change()
Sheets("Import_CFG").Cells(9, 11) = Me.ComboBox_Ziel_AKS_T5.ListIndex '+ 1 'ZielAdresse
Sheets("Import_CFG").Cells(9, 12) = Sheets("Import_CFG").Cells(9, 11) + 1
Ziel_Blatt_SName = Sheets("Import_CFG").Cells(1, 10)
Ziel_Blatt_ZielAdresse = Sheets("Import_CFG").Cells(9, 12)
On Error Resume Next
Me.lbl_Ziel_AKS_T5 = Sheets(Ziel_Blatt_SName).Cells(2000, Ziel_Blatt_ZielAdresse)

refrech



End Sub
Private Sub ComboBox_Ziel_AKS_T6_Change()
Sheets("Import_CFG").Cells(10, 11) = Me.ComboBox_Ziel_AKS_T6.ListIndex '+ 1 'ZielAdresse
Sheets("Import_CFG").Cells(10, 12) = Sheets("Import_CFG").Cells(10, 11) + 1
Ziel_Blatt_SName = Sheets("Import_CFG").Cells(1, 10)
Ziel_Blatt_ZielAdresse = Sheets("Import_CFG").Cells(10, 12)
On Error Resume Next
Me.lbl_Ziel_AKS_T6 = Sheets(Ziel_Blatt_SName).Cells(2000, Ziel_Blatt_ZielAdresse)

refrech



End Sub



Private Sub ComboBox_ZielAdresse_Change()
Sheets("Import_CFG").Cells(2, 11) = Me.ComboBox_ZielAdresse.ListIndex '+ 1 'ZielAdresse
Sheets("Import_CFG").Cells(2, 12) = Sheets("Import_CFG").Cells(2, 11) + 1
Ziel_Blatt_SName = Sheets("Import_CFG").Cells(1, 10)
Ziel_Blatt_ZielAdresse = Sheets("Import_CFG").Cells(2, 12)

Me.lbl_ZielAdresse = Sheets(Ziel_Blatt_SName).Cells(1, Ziel_Blatt_ZielAdresse)

refrech


End Sub

Private Sub ComboBox_ZielName_Change()
Sheets("Import_CFG").Cells(3, 11) = Me.ComboBox_ZielName.ListIndex '+ 1 'ZielName
Sheets("Import_CFG").Cells(3, 12) = Sheets("Import_CFG").Cells(3, 11) + 1
Ziel_Blatt_SName = Sheets("Import_CFG").Cells(1, 10)
Ziel_Blatt_ZielName = Sheets("Import_CFG").Cells(3, 12)

Me.lbl_ZielName = Sheets(Ziel_Blatt_SName).Cells(1, Ziel_Blatt_ZielName)

refrech
End Sub

Private Sub ScrollBar_AKS_Change()
Me.lbl_ScrollBar_AKS = Me.ScrollBar_AKS.Value


refrech
End Sub

Private Sub UserForm_Initialize()
Me.cmd_exit.Picture = Application.CommandBars.GetImageMso("MailDelete", 20, 20)
Me.Label_QuellBlatt.Caption = Sheets("Import_CFG").Cells(1, 1) 'Quelle - BlattName
Me.Label_ZiellBlatt.Caption = Sheets("Import_CFG").Cells(1, 10) 'Ziel - BlattName
ZielBlattName = Sheets("Import_CFG").Cells(1, 10)


Sheets(ZielBlattName).Cells(2000, 7) = "AKS_T1"
Sheets(ZielBlattName).Cells(2000, 8) = "AKS_T2"
Sheets(ZielBlattName).Cells(2000, 9) = "AKS_T3"
Sheets(ZielBlattName).Cells(2000, 10) = "AKS_T4"
Sheets(ZielBlattName).Cells(2000, 11) = "AKS_T5"

Me.ComboBox_Quelle_AKS_T1.AddItem "A"
Me.ComboBox_Quelle_AKS_T1.AddItem "B"
Me.ComboBox_Quelle_AKS_T1.AddItem "C"
Me.ComboBox_Quelle_AKS_T1.AddItem "D"
Me.ComboBox_Quelle_AKS_T1.AddItem "E"
Me.ComboBox_Quelle_AKS_T1.AddItem "F"
Me.ComboBox_Quelle_AKS_T1.AddItem "G"
Me.ComboBox_Quelle_AKS_T1.AddItem "H"
Me.ComboBox_Quelle_AKS_T1.AddItem "I"
Me.ComboBox_Quelle_AKS_T1.AddItem "J"
Me.ComboBox_Quelle_AKS_T1.AddItem "K"
Me.ComboBox_Quelle_AKS_T1.AddItem "L"
Me.ComboBox_Quelle_AKS_T1.AddItem "M"
Me.ComboBox_Quelle_AKS_T1.AddItem "N"
Me.ComboBox_Quelle_AKS_T1.AddItem "O"
Me.ComboBox_Quelle_AKS_T1.AddItem "P"
Me.ComboBox_Quelle_AKS_T1.AddItem "Q"
Me.ComboBox_Quelle_AKS_T1.AddItem "R"
Me.ComboBox_Quelle_AKS_T1.AddItem "S"
Me.ComboBox_Quelle_AKS_T1.AddItem "T"
Me.ComboBox_Quelle_AKS_T1.AddItem "U"
Me.ComboBox_Quelle_AKS_T1.AddItem "V"
Me.ComboBox_Quelle_AKS_T1.AddItem "W"
Me.ComboBox_Quelle_AKS_T1.AddItem "X"
Me.ComboBox_Quelle_AKS_T1.AddItem "Y"
Me.ComboBox_Quelle_AKS_T1.AddItem "Z"

If Sheets("Import_CFG").Cells(5, 2).Value < 1 Then
    Me.ComboBox_Quelle_AKS_T1.ListIndex = 7
Else
   Me.ComboBox_Quelle_AKS_T1.ListIndex = Sheets("Import_CFG").Cells(5, 2)
End If

Me.ComboBox_Quelle_AKS_T2.AddItem "A"
Me.ComboBox_Quelle_AKS_T2.AddItem "B"
Me.ComboBox_Quelle_AKS_T2.AddItem "C"
Me.ComboBox_Quelle_AKS_T2.AddItem "D"
Me.ComboBox_Quelle_AKS_T2.AddItem "E"
Me.ComboBox_Quelle_AKS_T2.AddItem "F"
Me.ComboBox_Quelle_AKS_T2.AddItem "G"
Me.ComboBox_Quelle_AKS_T2.AddItem "H"
Me.ComboBox_Quelle_AKS_T2.AddItem "I"
Me.ComboBox_Quelle_AKS_T2.AddItem "J"
Me.ComboBox_Quelle_AKS_T2.AddItem "K"
Me.ComboBox_Quelle_AKS_T2.AddItem "L"
Me.ComboBox_Quelle_AKS_T2.AddItem "M"
Me.ComboBox_Quelle_AKS_T2.AddItem "N"
Me.ComboBox_Quelle_AKS_T2.AddItem "O"
Me.ComboBox_Quelle_AKS_T2.AddItem "P"
Me.ComboBox_Quelle_AKS_T2.AddItem "Q"
Me.ComboBox_Quelle_AKS_T2.AddItem "R"
Me.ComboBox_Quelle_AKS_T2.AddItem "S"
Me.ComboBox_Quelle_AKS_T2.AddItem "T"
Me.ComboBox_Quelle_AKS_T2.AddItem "U"
Me.ComboBox_Quelle_AKS_T2.AddItem "V"
Me.ComboBox_Quelle_AKS_T2.AddItem "W"
Me.ComboBox_Quelle_AKS_T2.AddItem "X"
Me.ComboBox_Quelle_AKS_T2.AddItem "Y"
Me.ComboBox_Quelle_AKS_T2.AddItem "Z"

If Sheets("Import_CFG").Cells(6, 2).Value < 1 Then
    Me.ComboBox_Quelle_AKS_T2.ListIndex = 8
Else
   Me.ComboBox_Quelle_AKS_T2.ListIndex = Sheets("Import_CFG").Cells(6, 2)
End If

Me.ComboBox_Quelle_AKS_T3.AddItem "A"
Me.ComboBox_Quelle_AKS_T3.AddItem "B"
Me.ComboBox_Quelle_AKS_T3.AddItem "C"
Me.ComboBox_Quelle_AKS_T3.AddItem "D"
Me.ComboBox_Quelle_AKS_T3.AddItem "E"
Me.ComboBox_Quelle_AKS_T3.AddItem "F"
Me.ComboBox_Quelle_AKS_T3.AddItem "G"
Me.ComboBox_Quelle_AKS_T3.AddItem "H"
Me.ComboBox_Quelle_AKS_T3.AddItem "I"
Me.ComboBox_Quelle_AKS_T3.AddItem "J"
Me.ComboBox_Quelle_AKS_T3.AddItem "K"
Me.ComboBox_Quelle_AKS_T3.AddItem "L"
Me.ComboBox_Quelle_AKS_T3.AddItem "M"
Me.ComboBox_Quelle_AKS_T3.AddItem "N"
Me.ComboBox_Quelle_AKS_T3.AddItem "O"
Me.ComboBox_Quelle_AKS_T3.AddItem "P"
Me.ComboBox_Quelle_AKS_T3.AddItem "Q"
Me.ComboBox_Quelle_AKS_T3.AddItem "R"
Me.ComboBox_Quelle_AKS_T3.AddItem "S"
Me.ComboBox_Quelle_AKS_T3.AddItem "T"
Me.ComboBox_Quelle_AKS_T3.AddItem "U"
Me.ComboBox_Quelle_AKS_T3.AddItem "V"
Me.ComboBox_Quelle_AKS_T3.AddItem "W"
Me.ComboBox_Quelle_AKS_T3.AddItem "X"
Me.ComboBox_Quelle_AKS_T3.AddItem "Y"
Me.ComboBox_Quelle_AKS_T3.AddItem "Z"

If Sheets("Import_CFG").Cells(7, 2).Value < 1 Then
    Me.ComboBox_Quelle_AKS_T3.ListIndex = 9
Else
   Me.ComboBox_Quelle_AKS_T3.ListIndex = Sheets("Import_CFG").Cells(7, 2)
End If


Me.ComboBox_Quelle_AKS_T4.AddItem "A"
Me.ComboBox_Quelle_AKS_T4.AddItem "B"
Me.ComboBox_Quelle_AKS_T4.AddItem "C"
Me.ComboBox_Quelle_AKS_T4.AddItem "D"
Me.ComboBox_Quelle_AKS_T4.AddItem "E"
Me.ComboBox_Quelle_AKS_T4.AddItem "F"
Me.ComboBox_Quelle_AKS_T4.AddItem "G"
Me.ComboBox_Quelle_AKS_T4.AddItem "H"
Me.ComboBox_Quelle_AKS_T4.AddItem "I"
Me.ComboBox_Quelle_AKS_T4.AddItem "J"
Me.ComboBox_Quelle_AKS_T4.AddItem "K"
Me.ComboBox_Quelle_AKS_T4.AddItem "L"
Me.ComboBox_Quelle_AKS_T4.AddItem "M"
Me.ComboBox_Quelle_AKS_T4.AddItem "N"
Me.ComboBox_Quelle_AKS_T4.AddItem "O"
Me.ComboBox_Quelle_AKS_T4.AddItem "P"
Me.ComboBox_Quelle_AKS_T4.AddItem "Q"
Me.ComboBox_Quelle_AKS_T4.AddItem "R"
Me.ComboBox_Quelle_AKS_T4.AddItem "S"
Me.ComboBox_Quelle_AKS_T4.AddItem "T"
Me.ComboBox_Quelle_AKS_T4.AddItem "U"
Me.ComboBox_Quelle_AKS_T4.AddItem "V"
Me.ComboBox_Quelle_AKS_T4.AddItem "W"
Me.ComboBox_Quelle_AKS_T4.AddItem "X"
Me.ComboBox_Quelle_AKS_T4.AddItem "Y"
Me.ComboBox_Quelle_AKS_T4.AddItem "Z"

If Sheets("Import_CFG").Cells(8, 2).Value < 1 Then
    Me.ComboBox_Quelle_AKS_T4.ListIndex = 10
Else
   Me.ComboBox_Quelle_AKS_T4.ListIndex = Sheets("Import_CFG").Cells(8, 2)
End If


Me.ComboBox_Quelle_AKS_T5.AddItem "A"
Me.ComboBox_Quelle_AKS_T5.AddItem "B"
Me.ComboBox_Quelle_AKS_T5.AddItem "C"
Me.ComboBox_Quelle_AKS_T5.AddItem "D"
Me.ComboBox_Quelle_AKS_T5.AddItem "E"
Me.ComboBox_Quelle_AKS_T5.AddItem "F"
Me.ComboBox_Quelle_AKS_T5.AddItem "G"
Me.ComboBox_Quelle_AKS_T5.AddItem "H"
Me.ComboBox_Quelle_AKS_T5.AddItem "I"
Me.ComboBox_Quelle_AKS_T5.AddItem "J"
Me.ComboBox_Quelle_AKS_T5.AddItem "K"
Me.ComboBox_Quelle_AKS_T5.AddItem "L"
Me.ComboBox_Quelle_AKS_T5.AddItem "M"
Me.ComboBox_Quelle_AKS_T5.AddItem "N"
Me.ComboBox_Quelle_AKS_T5.AddItem "O"
Me.ComboBox_Quelle_AKS_T5.AddItem "P"
Me.ComboBox_Quelle_AKS_T5.AddItem "Q"
Me.ComboBox_Quelle_AKS_T5.AddItem "R"
Me.ComboBox_Quelle_AKS_T5.AddItem "S"
Me.ComboBox_Quelle_AKS_T5.AddItem "T"
Me.ComboBox_Quelle_AKS_T5.AddItem "U"
Me.ComboBox_Quelle_AKS_T5.AddItem "V"
Me.ComboBox_Quelle_AKS_T5.AddItem "W"
Me.ComboBox_Quelle_AKS_T5.AddItem "X"
Me.ComboBox_Quelle_AKS_T5.AddItem "Y"
Me.ComboBox_Quelle_AKS_T5.AddItem "Z"

If Sheets("Import_CFG").Cells(9, 2).Value < 1 Then
    Me.ComboBox_Quelle_AKS_T5.ListIndex = 11
Else
   Me.ComboBox_Quelle_AKS_T5.ListIndex = Sheets("Import_CFG").Cells(9, 2)
End If

Me.ComboBox_Quelle_AKS_T6.AddItem "A"
Me.ComboBox_Quelle_AKS_T6.AddItem "B"
Me.ComboBox_Quelle_AKS_T6.AddItem "C"
Me.ComboBox_Quelle_AKS_T6.AddItem "D"
Me.ComboBox_Quelle_AKS_T6.AddItem "E"
Me.ComboBox_Quelle_AKS_T6.AddItem "F"
Me.ComboBox_Quelle_AKS_T6.AddItem "G"
Me.ComboBox_Quelle_AKS_T6.AddItem "H"
Me.ComboBox_Quelle_AKS_T6.AddItem "I"
Me.ComboBox_Quelle_AKS_T6.AddItem "J"
Me.ComboBox_Quelle_AKS_T6.AddItem "K"
Me.ComboBox_Quelle_AKS_T6.AddItem "L"
Me.ComboBox_Quelle_AKS_T6.AddItem "M"
Me.ComboBox_Quelle_AKS_T6.AddItem "N"
Me.ComboBox_Quelle_AKS_T6.AddItem "O"
Me.ComboBox_Quelle_AKS_T6.AddItem "P"
Me.ComboBox_Quelle_AKS_T6.AddItem "Q"
Me.ComboBox_Quelle_AKS_T6.AddItem "R"
Me.ComboBox_Quelle_AKS_T6.AddItem "S"
Me.ComboBox_Quelle_AKS_T6.AddItem "T"
Me.ComboBox_Quelle_AKS_T6.AddItem "U"
Me.ComboBox_Quelle_AKS_T6.AddItem "V"
Me.ComboBox_Quelle_AKS_T6.AddItem "W"
Me.ComboBox_Quelle_AKS_T6.AddItem "X"
Me.ComboBox_Quelle_AKS_T6.AddItem "Y"
Me.ComboBox_Quelle_AKS_T6.AddItem "Z"

If Sheets("Import_CFG").Cells(10, 2).Value < 1 Then
    Me.ComboBox_Quelle_AKS_T6.ListIndex = 12
Else
   Me.ComboBox_Quelle_AKS_T6.ListIndex = Sheets("Import_CFG").Cells(10, 2)
End If




'quelle
Me.ComboBox_Adresse.AddItem "A"
Me.ComboBox_Adresse.AddItem "B"
Me.ComboBox_Adresse.AddItem "C"
Me.ComboBox_Adresse.AddItem "D"
Me.ComboBox_Adresse.AddItem "E"
Me.ComboBox_Adresse.AddItem "F"
Me.ComboBox_Adresse.AddItem "G"
Me.ComboBox_Adresse.AddItem "H"
Me.ComboBox_Adresse.AddItem "I"
Me.ComboBox_Adresse.AddItem "J"
Me.ComboBox_Adresse.AddItem "K"
Me.ComboBox_Adresse.AddItem "L"
Me.ComboBox_Adresse.AddItem "M"
Me.ComboBox_Adresse.AddItem "N"
Me.ComboBox_Adresse.AddItem "O"
Me.ComboBox_Adresse.AddItem "P"
Me.ComboBox_Adresse.AddItem "Q"
Me.ComboBox_Adresse.AddItem "R"
Me.ComboBox_Adresse.AddItem "S"
Me.ComboBox_Adresse.AddItem "T"
Me.ComboBox_Adresse.AddItem "U"
Me.ComboBox_Adresse.AddItem "V"
Me.ComboBox_Adresse.AddItem "W"
Me.ComboBox_Adresse.AddItem "X"
Me.ComboBox_Adresse.AddItem "Y"
Me.ComboBox_Adresse.AddItem "Z"


If Sheets("Import_CFG").Cells(2, 2).Value < 1 Then
    Me.ComboBox_Adresse.ListIndex = 18
Else
   Me.ComboBox_Adresse.ListIndex = Sheets("Import_CFG").Cells(2, 2)
End If



Me.ComboBox_ZielAdresse.AddItem "A"
Me.ComboBox_ZielAdresse.AddItem "B"
Me.ComboBox_ZielAdresse.AddItem "C"
Me.ComboBox_ZielAdresse.AddItem "D"
Me.ComboBox_ZielAdresse.AddItem "E"
Me.ComboBox_ZielAdresse.AddItem "F"
Me.ComboBox_ZielAdresse.AddItem "G"
Me.ComboBox_ZielAdresse.AddItem "H"
Me.ComboBox_ZielAdresse.AddItem "I"
Me.ComboBox_ZielAdresse.AddItem "J"
Me.ComboBox_ZielAdresse.AddItem "K"
Me.ComboBox_ZielAdresse.AddItem "L"
Me.ComboBox_ZielAdresse.AddItem "M"
Me.ComboBox_ZielAdresse.AddItem "N"
Me.ComboBox_ZielAdresse.AddItem "O"
Me.ComboBox_ZielAdresse.AddItem "P"
Me.ComboBox_ZielAdresse.AddItem "Q"
Me.ComboBox_ZielAdresse.AddItem "R"
Me.ComboBox_ZielAdresse.AddItem "S"
Me.ComboBox_ZielAdresse.AddItem "T"
Me.ComboBox_ZielAdresse.AddItem "U"
Me.ComboBox_ZielAdresse.AddItem "V"
Me.ComboBox_ZielAdresse.AddItem "W"
Me.ComboBox_ZielAdresse.AddItem "X"
Me.ComboBox_ZielAdresse.AddItem "Y"
Me.ComboBox_ZielAdresse.AddItem "Z"



If Sheets("Import_CFG").Cells(2, 11).Value < 1 Then
    Me.ComboBox_ZielAdresse.ListIndex = 4
Else
   Me.ComboBox_ZielAdresse.ListIndex = Sheets("Import_CFG").Cells(2, 11)
End If



Me.ComboBox_Name.AddItem "A"
Me.ComboBox_Name.AddItem "B"
Me.ComboBox_Name.AddItem "C"
Me.ComboBox_Name.AddItem "D"
Me.ComboBox_Name.AddItem "E"
Me.ComboBox_Name.AddItem "F"
Me.ComboBox_Name.AddItem "G"
Me.ComboBox_Name.AddItem "H"
Me.ComboBox_Name.AddItem "I"
Me.ComboBox_Name.AddItem "J"
Me.ComboBox_Name.AddItem "K"
Me.ComboBox_Name.AddItem "L"
Me.ComboBox_Name.AddItem "M"
Me.ComboBox_Name.AddItem "N"
Me.ComboBox_Name.AddItem "O"
Me.ComboBox_Name.AddItem "P"
Me.ComboBox_Name.AddItem "Q"
Me.ComboBox_Name.AddItem "R"
Me.ComboBox_Name.AddItem "S"
Me.ComboBox_Name.AddItem "T"
Me.ComboBox_Name.AddItem "U"
Me.ComboBox_Name.AddItem "V"
Me.ComboBox_Name.AddItem "W"
Me.ComboBox_Name.AddItem "X"
Me.ComboBox_Name.AddItem "Y"
Me.ComboBox_Name.AddItem "Z"



If Sheets("Import_CFG").Cells(3, 2).Value < 1 Then
    Me.ComboBox_Name.ListIndex = 4
Else
   Me.ComboBox_Name.ListIndex = Sheets("Import_CFG").Cells(3, 2)
End If









Me.ComboBox_ZielName.AddItem "A"
Me.ComboBox_ZielName.AddItem "B"
Me.ComboBox_ZielName.AddItem "C"
Me.ComboBox_ZielName.AddItem "D"
Me.ComboBox_ZielName.AddItem "E"
Me.ComboBox_ZielName.AddItem "F"
Me.ComboBox_ZielName.AddItem "G"
Me.ComboBox_ZielName.AddItem "H"
Me.ComboBox_ZielName.AddItem "I"
Me.ComboBox_ZielName.AddItem "J"
Me.ComboBox_ZielName.AddItem "K"
Me.ComboBox_ZielName.AddItem "L"
Me.ComboBox_ZielName.AddItem "M"
Me.ComboBox_ZielName.AddItem "N"
Me.ComboBox_ZielName.AddItem "O"
Me.ComboBox_ZielName.AddItem "P"
Me.ComboBox_ZielName.AddItem "Q"
Me.ComboBox_ZielName.AddItem "R"
Me.ComboBox_ZielName.AddItem "S"
Me.ComboBox_ZielName.AddItem "T"
Me.ComboBox_ZielName.AddItem "U"
Me.ComboBox_ZielName.AddItem "V"
Me.ComboBox_ZielName.AddItem "W"
Me.ComboBox_ZielName.AddItem "X"
Me.ComboBox_ZielName.AddItem "Y"
Me.ComboBox_ZielName.AddItem "Z"



If Sheets("Import_CFG").Cells(3, 11).Value < 1 Then
    Me.ComboBox_ZielName.ListIndex = 5
Else
   Me.ComboBox_ZielName.ListIndex = Sheets("Import_CFG").Cells(3, 11)
End If

'Me.ComboBox_AKS.AddItem "A"
'Me.ComboBox_AKS.AddItem "B"
'Me.ComboBox_AKS.AddItem "C"
'Me.ComboBox_AKS.AddItem "D"
'Me.ComboBox_AKS.AddItem "E"
'Me.ComboBox_AKS.AddItem "F"
'Me.ComboBox_AKS.AddItem "G"
'Me.ComboBox_AKS.AddItem "H"
'Me.ComboBox_AKS.AddItem "I"
'Me.ComboBox_AKS.AddItem "J"
'Me.ComboBox_AKS.AddItem "K"
'Me.ComboBox_AKS.AddItem "L"
'Me.ComboBox_AKS.AddItem "M"
'Me.ComboBox_AKS.AddItem "N"
'Me.ComboBox_AKS.AddItem "O"
'Me.ComboBox_AKS.AddItem "P"
'Me.ComboBox_AKS.AddItem "Q"
'Me.ComboBox_AKS.AddItem "R"
'Me.ComboBox_AKS.AddItem "S"
'Me.ComboBox_AKS.AddItem "T"
'Me.ComboBox_AKS.AddItem "U"
'Me.ComboBox_AKS.AddItem "V"
'Me.ComboBox_AKS.AddItem "W"
'Me.ComboBox_AKS.AddItem "X"
'Me.ComboBox_AKS.AddItem "Y"
'Me.ComboBox_AKS.AddItem "Z"




'If Sheets("Import_CFG").Cells(4, 2).Value < 1 Then
'    Me.ComboBox_AKS.ListIndex = 6
'Else
'   Me.ComboBox_AKS.ListIndex = Sheets("Import_CFG").Cells(4, 2)
'End If


Me.CheckBox_Name_Import.Value = Sheets("Import_CFG").Cells(3, 1).Value



Me.CheckBox_AKS_T1.Value = Sheets("Import_CFG").Cells(5, 1).Value
Me.CheckBox_AKS_T2.Value = Sheets("Import_CFG").Cells(6, 1).Value
Me.CheckBox_AKS_T3.Value = Sheets("Import_CFG").Cells(7, 1).Value
Me.CheckBox_AKS_T4.Value = Sheets("Import_CFG").Cells(8, 1).Value
Me.CheckBox_AKS_T5.Value = Sheets("Import_CFG").Cells(9, 1).Value
Me.CheckBox_AKS_T6.Value = Sheets("Import_CFG").Cells(10, 1).Value




Me.ComboBox_Ziel_AKS_T1.AddItem "A"
Me.ComboBox_Ziel_AKS_T1.AddItem "B"
Me.ComboBox_Ziel_AKS_T1.AddItem "C"
Me.ComboBox_Ziel_AKS_T1.AddItem "D"
Me.ComboBox_Ziel_AKS_T1.AddItem "E"
Me.ComboBox_Ziel_AKS_T1.AddItem "F"
Me.ComboBox_Ziel_AKS_T1.AddItem "G"
Me.ComboBox_Ziel_AKS_T1.AddItem "H"
Me.ComboBox_Ziel_AKS_T1.AddItem "I"
Me.ComboBox_Ziel_AKS_T1.AddItem "J"
Me.ComboBox_Ziel_AKS_T1.AddItem "K"
Me.ComboBox_Ziel_AKS_T1.AddItem "L"
Me.ComboBox_Ziel_AKS_T1.AddItem "M"
Me.ComboBox_Ziel_AKS_T1.AddItem "N"
Me.ComboBox_Ziel_AKS_T1.AddItem "O"
Me.ComboBox_Ziel_AKS_T1.AddItem "P"
Me.ComboBox_Ziel_AKS_T1.AddItem "Q"
Me.ComboBox_Ziel_AKS_T1.AddItem "R"
Me.ComboBox_Ziel_AKS_T1.AddItem "S"
Me.ComboBox_Ziel_AKS_T1.AddItem "T"
Me.ComboBox_Ziel_AKS_T1.AddItem "U"
Me.ComboBox_Ziel_AKS_T1.AddItem "V"
Me.ComboBox_Ziel_AKS_T1.AddItem "W"
Me.ComboBox_Ziel_AKS_T1.AddItem "X"
Me.ComboBox_Ziel_AKS_T1.AddItem "Y"
Me.ComboBox_Ziel_AKS_T1.AddItem "Z"

If Sheets("Import_CFG").Cells(5, 11).Value < 1 Then
    Me.ComboBox_Ziel_AKS_T1.ListIndex = 6
Else
   Me.ComboBox_Ziel_AKS_T1.ListIndex = Sheets("Import_CFG").Cells(5, 11)
End If

Me.ComboBox_Ziel_AKS_T2.AddItem "A"
Me.ComboBox_Ziel_AKS_T2.AddItem "B"
Me.ComboBox_Ziel_AKS_T2.AddItem "C"
Me.ComboBox_Ziel_AKS_T2.AddItem "D"
Me.ComboBox_Ziel_AKS_T2.AddItem "E"
Me.ComboBox_Ziel_AKS_T2.AddItem "F"
Me.ComboBox_Ziel_AKS_T2.AddItem "G"
Me.ComboBox_Ziel_AKS_T2.AddItem "H"
Me.ComboBox_Ziel_AKS_T2.AddItem "I"
Me.ComboBox_Ziel_AKS_T2.AddItem "J"
Me.ComboBox_Ziel_AKS_T2.AddItem "K"
Me.ComboBox_Ziel_AKS_T2.AddItem "L"
Me.ComboBox_Ziel_AKS_T2.AddItem "M"
Me.ComboBox_Ziel_AKS_T2.AddItem "N"
Me.ComboBox_Ziel_AKS_T2.AddItem "O"
Me.ComboBox_Ziel_AKS_T2.AddItem "P"
Me.ComboBox_Ziel_AKS_T2.AddItem "Q"
Me.ComboBox_Ziel_AKS_T2.AddItem "R"
Me.ComboBox_Ziel_AKS_T2.AddItem "S"
Me.ComboBox_Ziel_AKS_T2.AddItem "T"
Me.ComboBox_Ziel_AKS_T2.AddItem "U"
Me.ComboBox_Ziel_AKS_T2.AddItem "V"
Me.ComboBox_Ziel_AKS_T2.AddItem "W"
Me.ComboBox_Ziel_AKS_T2.AddItem "X"
Me.ComboBox_Ziel_AKS_T2.AddItem "Y"
Me.ComboBox_Ziel_AKS_T2.AddItem "Z"

If Sheets("Import_CFG").Cells(6, 11).Value < 1 Then
    Me.ComboBox_Ziel_AKS_T2.ListIndex = 7
Else
   Me.ComboBox_Ziel_AKS_T2.ListIndex = Sheets("Import_CFG").Cells(6, 11)
End If


Me.ComboBox_Ziel_AKS_T3.AddItem "A"
Me.ComboBox_Ziel_AKS_T3.AddItem "B"
Me.ComboBox_Ziel_AKS_T3.AddItem "C"
Me.ComboBox_Ziel_AKS_T3.AddItem "D"
Me.ComboBox_Ziel_AKS_T3.AddItem "E"
Me.ComboBox_Ziel_AKS_T3.AddItem "F"
Me.ComboBox_Ziel_AKS_T3.AddItem "G"
Me.ComboBox_Ziel_AKS_T3.AddItem "H"
Me.ComboBox_Ziel_AKS_T3.AddItem "I"
Me.ComboBox_Ziel_AKS_T3.AddItem "J"
Me.ComboBox_Ziel_AKS_T3.AddItem "K"
Me.ComboBox_Ziel_AKS_T3.AddItem "L"
Me.ComboBox_Ziel_AKS_T3.AddItem "M"
Me.ComboBox_Ziel_AKS_T3.AddItem "N"
Me.ComboBox_Ziel_AKS_T3.AddItem "O"
Me.ComboBox_Ziel_AKS_T3.AddItem "P"
Me.ComboBox_Ziel_AKS_T3.AddItem "Q"
Me.ComboBox_Ziel_AKS_T3.AddItem "R"
Me.ComboBox_Ziel_AKS_T3.AddItem "S"
Me.ComboBox_Ziel_AKS_T3.AddItem "T"
Me.ComboBox_Ziel_AKS_T3.AddItem "U"
Me.ComboBox_Ziel_AKS_T3.AddItem "V"
Me.ComboBox_Ziel_AKS_T3.AddItem "W"
Me.ComboBox_Ziel_AKS_T3.AddItem "X"
Me.ComboBox_Ziel_AKS_T3.AddItem "Y"
Me.ComboBox_Ziel_AKS_T3.AddItem "Z"

If Sheets("Import_CFG").Cells(7, 11).Value < 1 Then
    Me.ComboBox_Ziel_AKS_T3.ListIndex = 8
Else
   Me.ComboBox_Ziel_AKS_T3.ListIndex = Sheets("Import_CFG").Cells(7, 11)
End If


Me.ComboBox_Ziel_AKS_T4.AddItem "A"
Me.ComboBox_Ziel_AKS_T4.AddItem "B"
Me.ComboBox_Ziel_AKS_T4.AddItem "C"
Me.ComboBox_Ziel_AKS_T4.AddItem "D"
Me.ComboBox_Ziel_AKS_T4.AddItem "E"
Me.ComboBox_Ziel_AKS_T4.AddItem "F"
Me.ComboBox_Ziel_AKS_T4.AddItem "G"
Me.ComboBox_Ziel_AKS_T4.AddItem "H"
Me.ComboBox_Ziel_AKS_T4.AddItem "I"
Me.ComboBox_Ziel_AKS_T4.AddItem "J"
Me.ComboBox_Ziel_AKS_T4.AddItem "K"
Me.ComboBox_Ziel_AKS_T4.AddItem "L"
Me.ComboBox_Ziel_AKS_T4.AddItem "M"
Me.ComboBox_Ziel_AKS_T4.AddItem "N"
Me.ComboBox_Ziel_AKS_T4.AddItem "O"
Me.ComboBox_Ziel_AKS_T4.AddItem "P"
Me.ComboBox_Ziel_AKS_T4.AddItem "Q"
Me.ComboBox_Ziel_AKS_T4.AddItem "R"
Me.ComboBox_Ziel_AKS_T4.AddItem "S"
Me.ComboBox_Ziel_AKS_T4.AddItem "T"
Me.ComboBox_Ziel_AKS_T4.AddItem "U"
Me.ComboBox_Ziel_AKS_T4.AddItem "V"
Me.ComboBox_Ziel_AKS_T4.AddItem "W"
Me.ComboBox_Ziel_AKS_T4.AddItem "X"
Me.ComboBox_Ziel_AKS_T4.AddItem "Y"
Me.ComboBox_Ziel_AKS_T4.AddItem "Z"

If Sheets("Import_CFG").Cells(8, 11).Value < 1 Then
    Me.ComboBox_Ziel_AKS_T4.ListIndex = 9
Else
   Me.ComboBox_Ziel_AKS_T4.ListIndex = Sheets("Import_CFG").Cells(8, 11)
End If



Me.ComboBox_Ziel_AKS_T5.AddItem "A"
Me.ComboBox_Ziel_AKS_T5.AddItem "B"
Me.ComboBox_Ziel_AKS_T5.AddItem "C"
Me.ComboBox_Ziel_AKS_T5.AddItem "D"
Me.ComboBox_Ziel_AKS_T5.AddItem "E"
Me.ComboBox_Ziel_AKS_T5.AddItem "F"
Me.ComboBox_Ziel_AKS_T5.AddItem "G"
Me.ComboBox_Ziel_AKS_T5.AddItem "H"
Me.ComboBox_Ziel_AKS_T5.AddItem "I"
Me.ComboBox_Ziel_AKS_T5.AddItem "J"
Me.ComboBox_Ziel_AKS_T5.AddItem "K"
Me.ComboBox_Ziel_AKS_T5.AddItem "L"
Me.ComboBox_Ziel_AKS_T5.AddItem "M"
Me.ComboBox_Ziel_AKS_T5.AddItem "N"
Me.ComboBox_Ziel_AKS_T5.AddItem "O"
Me.ComboBox_Ziel_AKS_T5.AddItem "P"
Me.ComboBox_Ziel_AKS_T5.AddItem "Q"
Me.ComboBox_Ziel_AKS_T5.AddItem "R"
Me.ComboBox_Ziel_AKS_T5.AddItem "S"
Me.ComboBox_Ziel_AKS_T5.AddItem "T"
Me.ComboBox_Ziel_AKS_T5.AddItem "U"
Me.ComboBox_Ziel_AKS_T5.AddItem "V"
Me.ComboBox_Ziel_AKS_T5.AddItem "W"
Me.ComboBox_Ziel_AKS_T5.AddItem "X"
Me.ComboBox_Ziel_AKS_T5.AddItem "Y"
Me.ComboBox_Ziel_AKS_T5.AddItem "Z"

If Sheets("Import_CFG").Cells(9, 11).Value < 1 Then
    Me.ComboBox_Ziel_AKS_T5.ListIndex = 10
Else
   Me.ComboBox_Ziel_AKS_T5.ListIndex = Sheets("Import_CFG").Cells(9, 11)
End If



Me.ComboBox_Ziel_AKS_T6.AddItem "A"
Me.ComboBox_Ziel_AKS_T6.AddItem "B"
Me.ComboBox_Ziel_AKS_T6.AddItem "C"
Me.ComboBox_Ziel_AKS_T6.AddItem "D"
Me.ComboBox_Ziel_AKS_T6.AddItem "E"
Me.ComboBox_Ziel_AKS_T6.AddItem "F"
Me.ComboBox_Ziel_AKS_T6.AddItem "G"
Me.ComboBox_Ziel_AKS_T6.AddItem "H"
Me.ComboBox_Ziel_AKS_T6.AddItem "I"
Me.ComboBox_Ziel_AKS_T6.AddItem "J"
Me.ComboBox_Ziel_AKS_T6.AddItem "K"
Me.ComboBox_Ziel_AKS_T6.AddItem "L"
Me.ComboBox_Ziel_AKS_T6.AddItem "M"
Me.ComboBox_Ziel_AKS_T6.AddItem "N"
Me.ComboBox_Ziel_AKS_T6.AddItem "O"
Me.ComboBox_Ziel_AKS_T6.AddItem "P"
Me.ComboBox_Ziel_AKS_T6.AddItem "Q"
Me.ComboBox_Ziel_AKS_T6.AddItem "R"
Me.ComboBox_Ziel_AKS_T6.AddItem "S"
Me.ComboBox_Ziel_AKS_T6.AddItem "T"
Me.ComboBox_Ziel_AKS_T6.AddItem "U"
Me.ComboBox_Ziel_AKS_T6.AddItem "V"
Me.ComboBox_Ziel_AKS_T6.AddItem "W"
Me.ComboBox_Ziel_AKS_T6.AddItem "X"
Me.ComboBox_Ziel_AKS_T6.AddItem "Y"
Me.ComboBox_Ziel_AKS_T6.AddItem "Z"

If Sheets("Import_CFG").Cells(10, 11).Value < 1 Then
    Me.ComboBox_Ziel_AKS_T6.ListIndex = 11
Else
   Me.ComboBox_Ziel_AKS_T6.ListIndex = Sheets("Import_CFG").Cells(10, 11)
End If


For iz = 1 To 100
    Me.ComboBox_Zeile.AddItem iz
Next

If Sheets("Import_CFG").Cells(11, 2).Value < 1 Then
    Me.ComboBox_Zeile.ListIndex = 0
Else
    Me.ComboBox_Zeile.Text = Sheets("Import_CFG").Cells(11, 2)
End If






refrech


End Sub
Private Sub refrech()
Import_Blatt_SName = Sheets("Import_CFG").Cells(1, 1)

'Import_Blatt_AKS = Sheets("Import_CFG").Cells(4, 3)
If Import_Blatt_AKS < 1 Then Import_Blatt_AKS = 1




AKS_Zeile = Me.lbl_ScrollBar_AKS.Caption

Blatt_Name_Zeile = Sheets("Import_CFG").Cells(3, 3)


'Adresse


Import_Blatt_Adresse = Sheets("Import_CFG").Cells(2, 2) + 1
Sheets("Import_CFG").Cells(2, 3) = Sheets("Import_CFG").Cells(2, 2) + 1


'Beschreibung
Me.lbl_Beschreibung = Sheets(Import_Blatt_SName).Cells(AKS_Zeile, Blatt_Name_Zeile)

'Adresse
Me.llb_Adresse_vorschau = Sheets(Import_Blatt_SName).Cells(AKS_Zeile, Import_Blatt_Adresse)


'AKS_T1

Import_Blatt_AKS_T1 = Sheets("Import_CFG").Cells(5, 1)
Import_Blatt_AKS_Spalt_T1 = Sheets("Import_CFG").Cells(5, 3)
Me.lbl_AKS_T1.Caption = Sheets(Import_Blatt_SName).Cells(AKS_Zeile, Import_Blatt_AKS_Spalt_T1)
Me.CheckBox_AKS_T1.Caption = Sheets(Import_Blatt_SName).Cells(1, Import_Blatt_AKS_Spalt_T1)
'AKS_T2

Import_Blatt_AKS_T2 = Sheets("Import_CFG").Cells(6, 1)
Import_Blatt_AKS_Spalt_T2 = Sheets("Import_CFG").Cells(6, 3)
Me.lbl_AKS_T2.Caption = Sheets(Import_Blatt_SName).Cells(AKS_Zeile, Import_Blatt_AKS_Spalt_T2)
Me.CheckBox_AKS_T2.Caption = Sheets(Import_Blatt_SName).Cells(1, Import_Blatt_AKS_Spalt_T2)
'AKS_T3

Import_Blatt_AKS_T3 = Sheets("Import_CFG").Cells(7, 1)
Import_Blatt_AKS_Spalt_T3 = Sheets("Import_CFG").Cells(7, 3)
Me.lbl_AKS_T3.Caption = Sheets(Import_Blatt_SName).Cells(AKS_Zeile, Import_Blatt_AKS_Spalt_T3)
Me.CheckBox_AKS_T3.Caption = Sheets(Import_Blatt_SName).Cells(1, Import_Blatt_AKS_Spalt_T3)
'AKS_T4

Import_Blatt_AKS_T4 = Sheets("Import_CFG").Cells(8, 1)
Import_Blatt_AKS_Spalt_T4 = Sheets("Import_CFG").Cells(8, 3)
Me.lbl_AKS_T4.Caption = Sheets(Import_Blatt_SName).Cells(AKS_Zeile, Import_Blatt_AKS_Spalt_T4)
Me.CheckBox_AKS_T4.Caption = Sheets(Import_Blatt_SName).Cells(1, Import_Blatt_AKS_Spalt_T4)

'AKS_T5

Import_Blatt_AKS_T5 = Sheets("Import_CFG").Cells(9, 1)
Import_Blatt_AKS_Spalt_T5 = Sheets("Import_CFG").Cells(9, 3)
Me.lbl_AKS_T5.Caption = Sheets(Import_Blatt_SName).Cells(AKS_Zeile, Import_Blatt_AKS_Spalt_T5)
Me.CheckBox_AKS_T5.Caption = Sheets(Import_Blatt_SName).Cells(1, Import_Blatt_AKS_Spalt_T5)

'AKS_T6

Import_Blatt_AKS_T6 = Sheets("Import_CFG").Cells(10, 1)
Import_Blatt_AKS_Spalt_T6 = Sheets("Import_CFG").Cells(10, 3)
Me.lbl_AKS_T6.Caption = Sheets(Import_Blatt_SName).Cells(AKS_Zeile, Import_Blatt_AKS_Spalt_T6)
Me.CheckBox_AKS_T6.Caption = Sheets(Import_Blatt_SName).Cells(1, Import_Blatt_AKS_Spalt_T6)















' Name Import Checkbox

If Sheets("Import_CFG").Cells(3, 1).Value = True Then
    Me.CheckBox_Name_Import.Value = True
    Me.ComboBox_Name.Enabled = True
    Me.lbl_Name.Enabled = True
    Me.Label_Spalte_Name.Enabled = True
    Me.lbl_ZielName.Enabled = True
    Me.lbl_ZielNameBeschriftug.Enabled = True
    Me.ComboBox_ZielName.Enabled = True
Else
    Me.ComboBox_Name.Enabled = False
    Me.lbl_Name.Enabled = False
    Me.Label_Spalte_Name.Enabled = False
    Me.lbl_ZielName.Enabled = False
    Me.lbl_ZielNameBeschriftug.Enabled = False
    Me.ComboBox_ZielName.Enabled = False

End If



'AKS Import Checkbox
If Sheets("Import_CFG").Cells(4, 1).Value = True Then
    Me.CheckBox_AKS_Import.Value = True
    Me.ComboBox_AKS.Enabled = True
    Me.lbl_AKS.Enabled = True
    Me.Label_Spalte_AKS.Enabled = True
    Me.Frame_AKS_Mapping.Enabled = True


Else
    Me.CheckBox_AKS_Import.Value = False
    Me.ComboBox_AKS.Enabled = False
    Me.lbl_AKS.Enabled = False
    Me.Label_Spalte_AKS.Enabled = False
    Me.Frame_AKS_Mapping.Enabled = False

    
End If


If Sheets("Import_CFG").Cells(5, 1).Value = True Then
Me.ComboBox_Ziel_AKS_T1.Enabled = True
Else
Me.ComboBox_Ziel_AKS_T1.Enabled = False
End If

If Sheets("Import_CFG").Cells(6, 1).Value = True Then
Me.ComboBox_Ziel_AKS_T2.Enabled = True
Else
Me.ComboBox_Ziel_AKS_T2.Enabled = False
End If

If Sheets("Import_CFG").Cells(7, 1).Value = True Then
Me.ComboBox_Ziel_AKS_T3.Enabled = True
Else
Me.ComboBox_Ziel_AKS_T3.Enabled = False
End If

If Sheets("Import_CFG").Cells(8, 1).Value = True Then
Me.ComboBox_Ziel_AKS_T4.Enabled = True
Else
Me.ComboBox_Ziel_AKS_T4.Enabled = False
End If

If Sheets("Import_CFG").Cells(9, 1).Value = True Then
Me.ComboBox_Ziel_AKS_T5.Enabled = True
Else
Me.ComboBox_Ziel_AKS_T5.Enabled = False
End If


If Sheets("Import_CFG").Cells(10, 1).Value = True Then
Me.ComboBox_Ziel_AKS_T6.Enabled = True
Else
Me.ComboBox_Ziel_AKS_T6.Enabled = False
End If


If Sheets("Import_CFG").Cells(4, 1).Value = True Then
Me.Frame_AKS_Mapping.Visible = True
Else
Me.Frame_AKS_Mapping.Visible = False
End If
End Sub
