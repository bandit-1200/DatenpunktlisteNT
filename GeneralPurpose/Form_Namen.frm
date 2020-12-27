VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_Namen 
   Caption         =   "Namen_abgleich"
   ClientHeight    =   10005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8910
   OleObjectBlob   =   "Form_Namen.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Form_Namen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim z As Integer
Dim letztezeile_namen As Integer


Private Sub cmd_del_all_Click()
    letztezeile_namen = Sheets("Namen_cfg").Cells(Rows.Count, 1).End(xlUp).Row
    For l = 1 To letztezeile_namen
        Sheets("Namen_cfg").Cells(l, 1) = ""
    Next
    
    ListBox_Namen.Clear
    
    z = 1
    Call UserForm_Initialize
    
    
End Sub



Private Sub cmd_DMS_add_Click()
    letztezeile_DMS = Sheets("Namen_cfg").Cells(Rows.Count, 3).End(xlUp).Row
    aktuelleDMS = Sheets("Namen_cfg").Cells(1, 2).Value
    neueDMS = InputBox("Neue DMS Eintragen", "Neue DMS", aktuelleDMS)
    Sheets("Namen_cfg").Cells(letztezeile_DMS + 1, 3) = neueDMS
    Call UserForm_Initialize
End Sub

Private Sub cmd_DMS_del_Click()
    'MsgBox Me.ListBox_DMS_Auswahl.ListIndex
    Sheets("Namen_cfg").Cells(Me.ListBox_DMS_Auswahl.ListIndex + 1, 3) = ""
    Call UserForm_Initialize

End Sub

Private Sub cmd_exit_Click()
Unload Me

If Me.cmd_exit.Caption = "OK" Then
    InbetriebnahmeProtokoll_erstellen
End If

End Sub


Private Sub cmd_Namen_Plist_Click()
 Dim auswahl As Integer
 If Me.Option_json = True Then auswahl = 1
 If Me.Option_Plist = True Then auswahl = 2
 If cmd_Namen_Plist.Caption = "Ausführen" Then auswahl = 3
 

 Select Case auswahl
    Case 1 To 2
        'json
        If Me.Option_json = True Then Abgleich_DMS_Namen
        'Call Abgleich_DMS_Namen
    
    'Case 2
        'Call copy_namen
        'plist
        If Me.Option_Plist = True Then copy_namen
    
    Case 3
        Call Export_Objekte_PromosNT
    
    Case Else
        Debug.Print auswahl
 
 
 End Select
 
End Sub



Private Sub ListBox_DMS_Auswahl_Change()
    If Me.ListBox_DMS_Auswahl.Text <> "" Then
        Sheets("Namen_cfg").Cells(1, 2).Value = Me.ListBox_DMS_Auswahl.Text
        Me.cmd_DMS_del.Visible = True
    Else
        Me.cmd_DMS_del.Visible = False
    End If
    
    
End Sub



Private Sub ListBox_Namen_auswahl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

letztezeile_namen = Sheets("Namen_cfg").Cells(Rows.Count, 1).End(xlUp).Row
If Sheets("Namen_cfg").Cells(1, 1) <> "" Then letztezeile_namen = letztezeile_namen + 1
Sheets("Namen_cfg").Cells(letztezeile_namen, 1) = ListBox_Namen_auswahl.Text
Debug.Print letztezeile_namen

ListBox_Namen.Clear

For List2Fill = 1 To letztezeile_namen
    sn_add = Sheets("Namen_cfg").Cells(List2Fill, 1)
    ListBox_Namen.AddItem sn_add
Next
End Sub

Private Sub Option_json_Click()
Call UserForm_Initialize
End Sub

Private Sub Option_Plist_Click()

Call UserForm_Initialize
End Sub

Private Sub UserForm_Deactivate()

Call Form_Blatmanager.cmd_Gruppe_refresh_Click
    
End Sub

Private Sub UserForm_Initialize()
asn = ActiveSheet.Name
If Me.Option_json.Value = True Then
    Me.Frame_DMS_Auswahl.Visible = True
Else
    Me.Frame_DMS_Auswahl.Visible = False

End If

Call DblFind
Call Sortieren

Sheets(asn).Select
Me.cmd_exit.Picture = Application.CommandBars.GetImageMso("MailDelete", 20, 20)

Me.cmd_del_all.Picture = Application.CommandBars.GetImageMso("InkEraseMode", 20, 20)

Me.cmd_Namen_Plist.Picture = Application.CommandBars.GetImageMso("ServerConnection", 20, 20)
Me.cmd_DMS_add.Picture = Application.CommandBars.GetImageMso("OutlineExpand", 20, 20)
Me.cmd_DMS_del.Picture = Application.CommandBars.GetImageMso("OutlineCollapse", 20, 20)
'Me.Image1.Picture = Application.CommandBars.GetImageMso("OutlineExpand", 20, 20)

'ServerConnection

'InkEraseMode
    Me.ListBox_Namen_auswahl.Clear
    Me.ListBox_Namen.Clear
    Me.ListBox_DMS_Auswahl.Clear
    
    letztezeile_namen = Sheets("Namen_cfg").Cells(Rows.Count, 1).End(xlUp).Row
    letztezeile_DMS = Sheets("Namen_cfg").Cells(Rows.Count, 3).End(xlUp).Row
    
z = letztezeile_namen + 1

sc = Sheets.Count

For ListFill = 1 To sc
    sn = Sheets(ListFill).Name
    ListBox_Namen_auswahl.AddItem sn
Next

For List2Fill = 1 To letztezeile_namen
    sn_add = Sheets("Namen_cfg").Cells(List2Fill, 1)
    ListBox_Namen.AddItem sn_add
Next


For DMS_Count = 1 To letztezeile_DMS
    
    Me.ListBox_DMS_Auswahl.AddItem Sheets("Namen_cfg").Cells(DMS_Count, 3)
 
Next
On Error GoTo fehler_listbox
Me.ListBox_DMS_Auswahl.Text = Sheets("Namen_cfg").Cells(1, 2)
On Error GoTo 0
'ListBox_Namen
Exit Sub

fehler_listbox:
If Sheets("Namen_cfg").Cells(1, 3) <> "" Then
    Me.ListBox_DMS_Auswahl.Text = Sheets("Namen_cfg").Cells(1, 3)
Else

End If

End Sub


Private Sub DblFind()
Application.ScreenUpdating = False

    Sheets("Namen_cfg").Visible = True

Sheets("Namen_cfg").Select
   Dim iRow As Integer, iRowL As Integer
   iRowL = Sheets("Namen_cfg").Cells(Cells.Rows.Count, 3).End(xlUp).Row
   For iRow = iRowL To 1 Step -1
      If WorksheetFunction.CountIf(Sheets("Namen_cfg").Columns(3), Cells(iRow, 3)) > 1 Then
         Rows(iRow).Delete
      End If
   Next iRow
   
   Application.ScreenUpdating = True
End Sub

Private Sub Sortieren()
Application.ScreenUpdating = False
    ActiveWorkbook.Worksheets("Namen_cfg").Select
    Columns("C:C").Select
    ActiveWorkbook.Worksheets("Namen_cfg").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Namen_cfg").Sort.SortFields.Add2 Key:=Range( _
        "C1:C200"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Namen_cfg").Sort
        .SetRange Range("C1:C200")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Application.ScreenUpdating = True
End Sub










Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call Form_Blatmanager.cmd_Gruppe_refresh_Click
End Sub
