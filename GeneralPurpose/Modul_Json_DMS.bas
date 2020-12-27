Attribute VB_Name = "Modul_Json_DMS"
'Namen per Json ins DMS laden
Public fehler_send2DMS As Boolean
Sub Abgleich_DMS_Namen()
fehler_send2DMS = False
  letztezeile_namenCFG = Sheets("Namen_cfg").UsedRange.SpecialCells(xlCellTypeLastCell).Row

  For n2pl = 1 To letztezeile_namenCFG
        
        namen_via_json = Sheets("Namen_cfg").Cells(n2pl, 1)
        If fehler_send2DMS Then namen_via_json = ""
        'MsgBox namen_via_json
      
      'ProzessBar*******************************************************************
        ProzessBarCSV.Show
        ProzessBarCSV.lbl_warten.Caption = "Bitte warten....exportiere..." & namen_via_json
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

    If namen_via_json = "" Then
    
        'ProzessBar*******************************************************************
        ProzessBarCSV.lbl_warten.Caption = "Export Fertig!..."
        ProzessBarCSV.csvBar.Value = 100
        Application.Wait (Now + TimeValue("0:00:02"))
        Unload ProzessBarCSV
        'ProzessBar*******************************************************************
    
        Exit Sub
    End If
    
    'Filter rücksetzen
    With Sheets(namen_via_json)
      If .FilterMode Then .ShowAllData
    End With

letztezeileNamen = Sheets(namen_via_json).Cells(Rows.Count, 6).End(xlUp).Row

    For na = 2 To letztezeileNamen
        If fehler_send2DMS Then
            Unload ProzessBarCSV
            Exit Sub
        End If
        'ProzessBar*******************************************************************
        ProzessBarCSV.csvBar.Value = na / letztezeileNamen * 100
        'ProzessBar*******************************************************************
    
        DMS_Name = Sheets(namen_via_json).Cells(na, 6)
        DMS_Name = Replace(DMS_Name, "ä", "ae", , , vbBinaryCompare)
        DMS_Name = Replace(DMS_Name, "Ä", "Ae", , , vbBinaryCompare)
        DMS_Name = Replace(DMS_Name, "ö", "oe", , , vbBinaryCompare)
        DMS_Name = Replace(DMS_Name, "Ö", "Oe", , , vbBinaryCompare)
        DMS_Name = Replace(DMS_Name, "Ü", "Ue", , , vbBinaryCompare)
        DMS_Name = Replace(DMS_Name, "ü", "ue", , , vbBinaryCompare)
    
        DMS_AKS = Sheets(namen_via_json).Cells(na, 19)
        
        If DMS_Name <> "" Then
        
            'Debug.Print DMS_Name & " -- " & DMS_AKS
            JsonString = "{""whois"":""XLS"",""user"":""XLS"",""set"":[{""path"":""" & DMS_AKS & """ ,""value"":""" & DMS_Name & """,""type"":""string""}]}"
          
          DMS_Anfrage (JsonString)
        
        End If
    Next
    'ProzessBar*******************************************************************
        ProzessBarCSV.lbl_warten.Caption = "Export Fertig!..."
        Application.Wait (Now + TimeValue("0:00:00"))
        Unload ProzessBarCSV
    'ProzessBar*******************************************************************
Next
    'ProzessBar*******************************************************************
        ProzessBarCSV.lbl_warten.Caption = "Export Fertig!..."
        Application.Wait (Now + TimeValue("0:00:00"))
        Unload ProzessBarCSV
    'ProzessBar*******************************************************************
End Sub

Public Function DMS_Anfrage(ByVal JsonString As String)

    Url = "http://"
    Url = Url & Sheets("Namen_cfg").Cells(1, 2)
    Url = Url & ":9020/json_data"
    'wert = Replace(wert, ",", ".")
    
   ' Url = "http://localhost:9020/json_data"
         
        Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")

        objHTTP.Open "POST", Url, False
        objHTTP.setRequestHeader "accept", "application/json"
        objHTTP.setRequestHeader "Content-Type", "application/json"
        objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
        
        On Error GoTo end_function
            objHTTP.send (JsonString)
            Debug.Print objHTTP.responseText
        Exit Function
end_function:
        'DMS_cfg.ListBox_debug.BorderColor = RGB(255, 0, 0)
        MsgBox "Fehler bei der Übertragung", vbCritical, "Fehler..."
        fehler_send2DMS = True
End Function
