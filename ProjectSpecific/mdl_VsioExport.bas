Attribute VB_Name = "mdl_VsioExport"
Sub exportVisio()

MsgBox "Funktioniert nicht!!!", vbCritical
Exit Sub


VEdz = 1
VEds = 12

For VEi = 2 To 577
    regerenz_gefunden = False
    'Debug.Print letzteVE & " <> " & VEi
    If Sheets(1).Cells(VEi, 6) <> "" Then
    
        letzteVE = Sheets("Visio_Export").UsedRange.SpecialCells(xlCellTypeLastCell).Row
      
        VE_Adresse = Sheets(1).Cells(VEi, 5)
        VE_Name = Sheets(1).Cells(VEi, 6)
        VE_AKST1 = Sheets(1).Cells(VEi, 7)
        VE_AKST2 = Sheets(1).Cells(VEi, 8)
        VE_AKST3 = Sheets(1).Cells(VEi, 9)
        VE_AKST4 = Sheets(1).Cells(VEi, 10)
        VE_AKST5 = Sheets(1).Cells(VEi, 11)
        
        
        AKS_komplett = VE_AKST1 & VE_AKST2 & VE_AKST3 & VE_AKST4 & VE_AKST5
        AKS_Referenz = VE_AKST1 & VE_AKST2 & VE_AKST3
        AKS_Zusatz = VE_AKST4
    
    For sRef = 2 To letzteVE
        If Sheets("Visio_Export").Cells(sRef, 100) = AKS_Referenz Then
            regerenz_gefunden = True
            'Debug.Print AKS_Referenz
            
            VEdz = sRef
            Debug.Print letzteVE & " Zeile: " & sRef & " --> " & AKS_Referenz & " <-- " & regerenz_gefunden; ""
          
            For DP_belegt_check = 12 To 60 Step 5
                
                'Debug.Print VEdz & " - " & VEds
                
                If Sheets("Visio_Export").Cells(VEdz, DP_belegt_check).Value = "" Then
                    VEds = DP_belegt_check
                    Debug.Print VEdz & " - " & VEds
                    'MsgBox "belegt Frei" & DP_belegt_check & " Name: " & VE_Name & " AKS: " & AKS_komplett
                    Exit For
                'End If
        
                 Else
                      'VEdz = letzteVE + 1
                      'Exit For
                End If

            Next DP_belegt_check
  
        
        End If
    Next sRef
 
    'Exit Sub
  
    If regerenz_gefunden = False Then VEdz = VEdz + 1

        
        Sheets("Visio_Export").Cells(VEdz, 1) = "-"
        Sheets("Visio_Export").Cells(VEdz, 2) = "Benutzer"
        Sheets("Visio_Export").Cells(VEdz, 3) = "Benutzer"
        Sheets("Visio_Export").Cells(VEdz, 4) = VE_AKST3
        Sheets("Visio_Export").Cells(VEdz, 5) = "-"
        Sheets("Visio_Export").Cells(VEdz, 6) = "-"
        
        If regerenz_gefunden = False Then Sheets("Visio_Export").Cells(VEdz, 7) = VE_Name
        
        Sheets("Visio_Export").Cells(VEdz, 8) = "-"
        Sheets("Visio_Export").Cells(VEdz, 9) = "-"
        'Sheets("Visio_Export").Cells(VEdz,VEds+ 10) = "-"
        'Sheets("Visio_Export").Cells(VEdz,VEds+ 11) = "-"
        Sheets("Visio_Export").Cells(VEdz, VEds) = AKS_komplett
        Sheets("Visio_Export").Cells(VEdz, VEds + 1) = VE_Name
        Sheets("Visio_Export").Cells(VEdz, VEds + 2) = VE_Name
        
        Sheets("Visio_Export").Cells(VEdz, 100) = AKS_Referenz
       ' VEdz = VEdz + 1

    End If

Next VEi


End Sub

