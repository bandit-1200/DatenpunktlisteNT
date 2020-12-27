Attribute VB_Name = "mdl_ESchema"
Sub E_SchemaEinmischen()



'ProzessBar*******************************************************************
    ProzessBarCSV.Show
    ProzessBarCSV.lbl_warten.Caption = "Bitte warten....Import..."
    Dim Pausenlänge, Start, Ende, Gesamtdauer

    Pausenlänge = 0.1 ' Dauer festlegen.
    Start = Timer    ' Anfangszeit setzen.
    Do While Timer < Start + Pausenlänge
        DoEvents    ' Steuerung an andere Prozesse
            ' abgeben.
    Loop
    Ende = Timer    ' Ende festlegen.
    Gesamtdauer = Ende - Start    ' Gesamtdauer berechnen.
 '   MsgBox "Die Pause dauerte " & Gesamtdauer & " Sekunden"
'ProzessBar*******************************************************************



  Application.ScreenUpdating = False
  
  
  With ActiveSheet
    If .FilterMode Then .ShowAllData
  End With

E_Schema_loeschen

letztezeileES = ActiveSheet.Cells(Rows.Count, 6).End(xlUp).Row

For ES = 2 To letztezeileES

    'ProzessBar*******************************************************************
    ProzessBarCSV.csvBar.Value = ES / letztezeileES * 100
    'ProzessBar*******************************************************************

    Name_puffer = ActiveSheet.Cells(ES, 6)
    ESchema = ActiveSheet.Cells(ES, 33)
    
    'Debug.Print Name_puffer
    
    If ESchema <> "" Then
        Name_puffer = Name_puffer & " (" & ESchema & ")"
    'Debug.Print Name_puffer
    
    ActiveSheet.Cells(ES, 6) = Name_puffer
    
    End If
Next
Application.ScreenUpdating = True

'ProzessBar*******************************************************************
ProzessBarCSV.lbl_warten.Caption = "Import Fertig!..."
  Application.Wait (Now + TimeValue("0:00:02"))
Unload ProzessBarCSV
'ProzessBar*******************************************************************

End Sub

Sub E_Schema_loeschen()
'E-Schema aus Namen löschen


Dim strText As String
Dim vntArray As Variant
Dim strInDerKlammer(2) As String
Dim lngI As Long
 Application.ScreenUpdating = False
letztezeileES = ActiveSheet.Cells(Rows.Count, 6).End(xlUp).Row


For ES = 2 To letztezeileES



    Name_puffer = ActiveSheet.Cells(ES, 6)
    ESchema = ActiveSheet.Cells(ES, 33)
    
    strText = Name_puffer
    
    vntArray = Split(strText, "(", -1, 1)
    
    
    
    For lngI = 0 To UBound(vntArray)
    On Error Resume Next
     strInDerKlammer(lngI) = Split(vntArray(lngI), ")", -1, 1)(0)
    'MsgBox strInDerKlammer, , "Demo: In der Klammer"
    On Error GoTo 0
    Next
    
    'MsgBox strInDerKlammer(0)
    'MsgBox strInDerKlammer(1)
    On Error Resume Next
    If strInDerKlammer(1) <> "" Then
        If strInDerKlammer(1) = ESchema Then
            neuerName = strInDerKlammer(0)
            'Debug.Print strInDerKlammer(0)
            neuerName = Left(neuerName, Len(neuerName) - 1)
            ActiveSheet.Cells(ES, 6) = neuerName
        End If
        On Error GoTo 0
    End If
     'Debug.Print neuerName
 
Next
Application.ScreenUpdating = True




End Sub
