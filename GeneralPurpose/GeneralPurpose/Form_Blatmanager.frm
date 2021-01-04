VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Blatmanager 
   Caption         =   "Blattmanager"
   ClientHeight    =   9585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12735
   OleObjectBlob   =   "Form_Blatmanager.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Form_Blatmanager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function Farbe_ScrollBar(ByRef Farbcode As Variant)
Dim R As Integer
Dim G As Integer
Dim B As Integer
    
    
    
    GetRGB (Farbcode), R, G, B
    Me.ScrollBar_R.Value = R
    Me.ScrollBar_G.Value = G
    Me.ScrollBar_B.Value = B
    
    
    
End Function


Private Sub cmd_cfg_Click()
Application.ActiveSheet.Tab.Color = Me.cmd_cfg.BackColor
Farbe_ScrollBar (Me.cmd_cfg.BackColor)
End Sub

Private Sub cmd_database_Click()
Application.ActiveSheet.Tab.Color = Me.cmd_database.BackColor
Farbe_ScrollBar (Me.cmd_database.BackColor)
End Sub

Private Sub cmd_DP_Liste_Click()

Application.ActiveSheet.Tab.Color = Me.cmd_DP_Liste.BackColor
Farbe_ScrollBar (Me.cmd_DP_Liste.BackColor)
End Sub

Private Sub cmd_Import_Click()
Application.ActiveSheet.Tab.Color = Me.cmd_Import.BackColor
Farbe_ScrollBar (Me.cmd_Import.BackColor)
End Sub

Private Sub cmd_Inbetriebnahme_Click()
Application.ActiveSheet.Tab.Color = Me.cmd_Inbetriebnahme.BackColor
Farbe_ScrollBar (Me.cmd_Inbetriebnahme.BackColor)
End Sub

Private Sub cmd_Info_Click()
Application.ActiveSheet.Tab.Color = Me.cmd_Info.BackColor
Farbe_ScrollBar (Me.cmd_Info.BackColor)
End Sub

Private Sub cmd_Modulliste_Click()
Application.ActiveSheet.Tab.Color = Me.cmd_Modulliste.BackColor
Farbe_ScrollBar (Me.cmd_Modulliste.BackColor)
End Sub

Private Sub cmd_Promos_Click()
Application.ActiveSheet.Tab.Color = Me.cmd_Promos.BackColor
Farbe_ScrollBar (Me.cmd_Promos.BackColor)
End Sub

Private Sub cmd_Sonstige_Click()
Application.ActiveSheet.Tab.Color = Me.cmd_Sonstige.BackColor
Farbe_ScrollBar (Me.cmd_Sonstige.BackColor)
End Sub

Private Sub ScrollBar_B_Change()
Farbe_einstellen
End Sub

Private Sub ScrollBar_G_Change()
Farbe_einstellen
End Sub

Private Sub ScrollBar_R_Change()
Farbe_einstellen
End Sub

Private Sub cmd_Farbmix_Click()

    GRPZ = Me.ListBox_Gruppe.ListIndex + 1
    On Error GoTo fehler
    auswahl = Sheets("Import_CFG").Cells(GRPZ, 42)
    On Error GoTo 0
    Me.Controls(auswahl).BackColor = Me.cmd_Farbmix.BackColor
    Sheets("Import_CFG").Cells(GRPZ, 41) = Me.cmd_Farbmix.BackColor
Exit Sub
fehler:
MsgBox "Auswahl treffen!", vbInformation, "Auswahl"

End Sub

Private Sub GruppenBehandlung()

Sheets("Import_CFG").Cells(1, 43) = Me.Check_grp1.Value
Sheets("Import_CFG").Cells(2, 43) = Me.Check_grp2.Value
Sheets("Import_CFG").Cells(3, 43) = Me.Check_grp3.Value
Sheets("Import_CFG").Cells(4, 43) = Me.Check_grp4.Value
Sheets("Import_CFG").Cells(5, 43) = Me.Check_grp5.Value
Sheets("Import_CFG").Cells(6, 43) = Me.Check_grp6.Value
Sheets("Import_CFG").Cells(7, 43) = Me.Check_grp7.Value
Sheets("Import_CFG").Cells(8, 43) = Me.Check_grp8.Value
Sheets("Import_CFG").Cells(9, 43) = Me.Check_grp9.Value


End Sub

Sub cmd_Gruppe_refresh_Click()
    GruppenBehandlung
    sc = Sheets.Count
    
    farbgeGrp1 = cmd_DP_Liste.BackColor
    farbgeGrp2 = cmd_Modulliste.BackColor
    farbgeGrp3 = cmd_Inbetriebnahme.BackColor
    farbgeGrp4 = cmd_database.BackColor
    farbgeGrp5 = cmd_Promos.BackColor
    farbgeGrp6 = cmd_Import.BackColor
    farbgeGrp7 = cmd_Info.BackColor
    farbgeGrp8 = cmd_cfg.BackColor
    farbgeGrp9 = cmd_Sonstige.BackColor



    For n = 1 To sc
        If Sheets(n).Tab.Color = farbgeGrp1 Then
            If Me.Check_grp1 Then
                Sheets(n).Visible = True
            Else
                Sheets(n).Visible = False
            End If
        End If
        
        If Sheets(n).Tab.Color = farbgeGrp2 Then
            If Me.Check_grp2 Then
                Sheets(n).Visible = True
            Else
                Sheets(n).Visible = False
            End If
        End If
        
        If Sheets(n).Tab.Color = farbgeGrp3 Then
            If Me.Check_grp3 Then
                Sheets(n).Visible = True
            Else
                Sheets(n).Visible = False
            End If
        End If
        
        If Sheets(n).Tab.Color = farbgeGrp4 Then
            If Me.Check_grp4 Then
                Sheets(n).Visible = True
            Else
                Sheets(n).Visible = False
            End If
        End If
        
        If Sheets(n).Tab.Color = farbgeGrp5 Then
            If Me.Check_grp5 Then
                Sheets(n).Visible = True
            Else
                Sheets(n).Visible = False
            End If
        End If
        
        If Sheets(n).Tab.Color = farbgeGrp6 Then
            If Me.Check_grp6 Then
                Sheets(n).Visible = True
            Else
                Sheets(n).Visible = False
            End If
        End If
        
        If Sheets(n).Tab.Color = farbgeGrp7 Then
            If Me.Check_grp7 Then
                Sheets(n).Visible = True
            Else
                Sheets(n).Visible = False
            End If
        End If
        
        If Sheets(n).Tab.Color = farbgeGrp8 Then
            If Me.Check_grp8 Then
                Sheets(n).Visible = True
            Else
                Sheets(n).Visible = False
            End If
        End If
        
        If Sheets(n).Tab.Color = farbgeGrp9 Then
            If Me.Check_grp9 Then
                Sheets(n).Visible = True
            Else
                Sheets(n).Visible = False
            End If
        End If
        
    Next


End Sub
Private Sub Farbe_einstellen()
    
    FR = Me.ScrollBar_R.Value
    FG = Me.ScrollBar_G.Value
    FB = Me.ScrollBar_B.Value
    Me.lbl_R = FR
    Me.lbl_G = FG
    Me.lbl_B = FB
    




    Me.cmd_Farbmix.BackColor = RGB(FR, FG, FB)
    Me.lbl_Farbcode = "Farbcode: " & Me.cmd_Farbmix.BackColor



End Sub


Private Sub UserForm_Initialize()
    Dim farbe(10) As Integer
    Dim R(10)               As Integer
    Dim G(10)               As Integer
    Dim B(10)               As Integer
     
     
   ' Me.Image1.Picture = Application.CommandBars.GetImageMso("FontColorMoreColorsDialog", 25, 25)
    Me.Picture = Application.CommandBars.GetImageMso("FontColorMoreColorsDialog", 25, 25)
    
    
   
    Me.cmd_exit.Picture = Application.CommandBars.GetImageMso("MailDelete", 20, 20)
    Me.cmd_Farbmix.Picture = Application.CommandBars.GetImageMso("ShapeFillColorPickerClassic", 40, 40)
    Me.cmd_Gruppe_refresh.Picture = Application.CommandBars.GetImageMso("VisibilityVisible", 40, 40)
    
    Sheets("Import_CFG").Cells(1, 40) = Me.cmd_DP_Liste.Caption
    Sheets("Import_CFG").Cells(2, 40) = Me.cmd_Modulliste.Caption
    Sheets("Import_CFG").Cells(3, 40) = Me.cmd_Inbetriebnahme.Caption
    Sheets("Import_CFG").Cells(4, 40) = Me.cmd_database.Caption
    Sheets("Import_CFG").Cells(5, 40) = Me.cmd_Promos.Caption
    Sheets("Import_CFG").Cells(6, 40) = Me.cmd_Import.Caption
    Sheets("Import_CFG").Cells(7, 40) = Me.cmd_Info.Caption
    Sheets("Import_CFG").Cells(8, 40) = Me.cmd_cfg.Caption
    Sheets("Import_CFG").Cells(9, 40) = Me.cmd_Sonstige.Caption
    
    Sheets("Import_CFG").Cells(1, 42) = Me.cmd_DP_Liste.Name
    Sheets("Import_CFG").Cells(2, 42) = Me.cmd_Modulliste.Name
    Sheets("Import_CFG").Cells(3, 42) = Me.cmd_Inbetriebnahme.Name
    Sheets("Import_CFG").Cells(4, 42) = Me.cmd_database.Name
    Sheets("Import_CFG").Cells(5, 42) = Me.cmd_Promos.Name
    Sheets("Import_CFG").Cells(6, 42) = Me.cmd_Import.Name
    Sheets("Import_CFG").Cells(7, 42) = Me.cmd_Info.Name
    Sheets("Import_CFG").Cells(8, 42) = Me.cmd_cfg.Name
    Sheets("Import_CFG").Cells(9, 42) = Me.cmd_Sonstige.Name
    
    Me.Check_grp1.Value = Sheets("Import_CFG").Cells(1, 43).Value
    Me.Check_grp2.Value = Sheets("Import_CFG").Cells(2, 43).Value
    Me.Check_grp3.Value = Sheets("Import_CFG").Cells(3, 43).Value
    Me.Check_grp4.Value = Sheets("Import_CFG").Cells(4, 43).Value
    Me.Check_grp5.Value = Sheets("Import_CFG").Cells(5, 43).Value
    Me.Check_grp6.Value = Sheets("Import_CFG").Cells(6, 43).Value
    Me.Check_grp7.Value = Sheets("Import_CFG").Cells(7, 43).Value
    Me.Check_grp8.Value = Sheets("Import_CFG").Cells(8, 43).Value
    Me.Check_grp9.Value = Sheets("Import_CFG").Cells(9, 43).Value
    
    
    
    
    Me.ListBox_Gruppe.Clear
    For i = 1 To 10
        Gruppe = Sheets("Import_CFG").Cells(i, 40)
        GetRGB Sheets("Import_CFG").Cells(i, 41), R(i), G(i), B(i)
        Me.ListBox_Gruppe.AddItem Gruppe
        
    Next




    'GetRGB "116737", R, G, B

    
    'MsgBox "Rot: " & R(1) & vbCrLf & _
            "Grün: " & G(1) & vbCrLf & _
            "Blau: " & B(1)





    Me.cmd_DP_Liste.BackColor = RGB(R(1), G(1), B(1))
    Me.cmd_Modulliste.BackColor = RGB(R(2), G(2), B(2))
    Me.cmd_Inbetriebnahme.BackColor = RGB(R(3), G(3), B(3))
    Me.cmd_database.BackColor = RGB(R(4), G(4), B(4))
    Me.cmd_Promos.BackColor = RGB(R(5), G(5), B(5))
    Me.cmd_Import.BackColor = RGB(R(6), G(6), B(6))
    Me.cmd_Info.BackColor = RGB(R(7), G(7), B(7))
    Me.cmd_cfg.BackColor = RGB(R(8), G(8), B(8))
    Me.cmd_Sonstige.BackColor = RGB(R(9), G(9), B(9))





End Sub





Private Sub cmd_exit_Click()
    Unload Me
End Sub




























