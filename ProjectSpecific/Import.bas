Attribute VB_Name = "Import"
Sub ImportKabelzug() 'Namen aus Kabelzugliste in Belegungsliste importieren
Application.ScreenUpdating = False

letztezeileKZL = Sheets("Kabelzugliste").Cells(Rows.Count, 1).End(xlUp).Row

'2 bis 577
For sa = 2 To 577

    PCD_SaiaAdress = CStr(Sheets(1).Cells(sa, 5))

    For KZLZ = 2 To letztezeileKZL
    
        KZL_SaiaAdress = CStr(Sheets("Kabelzugliste").Cells(KZLZ, 4))
        If KZL_SaiaAdress = "" Then KZL_SaiaAdress = 10000
        
        KZL_NAME = Sheets("Kabelzugliste").Cells(KZLZ, 3)
        
        If KZL_SaiaAdress = PCD_SaiaAdress Then
            Sheets(1).Cells(sa, 6) = KZL_NAME
            
        End If
        
    
    Next
    

Next sa

Application.ScreenUpdating = True


End Sub




Sub Import_VISIO() 'AKS - Import aus VISO - Blatt


MsgBox "Import"



letztezeileVISIO = Sheets("Visio_Import").Cells(Rows.Count, 4).End(xlUp).Row


    For sad = 2 To 577
        PCD_SaiaAdress = CStr(Sheets(1).Cells(sad, 5))
        
        For VISIOz = 2 To letztezeileVISIO
            VISIO_SaiaAdress = CStr(Sheets("Visio_Import").Cells(VISIOz, 19))
            If VISIO_SaiaAdress = "" Then VISIO_SaiaAdress = 10000
            
            'Application.StatusBar = PCD_SaiaAdress

            'Application.StatusBar = "VISIO: " & VISIO_SaiaAdress & " - " & "SAIA: " & PCD_SaiaAdress
            
            
            If VISIO_SaiaAdress = PCD_SaiaAdress Then
            'MsgBox "="
                VISIO_AKS_T1 = Sheets("Visio_Import").Cells(VISIOz, 8)
                VISIO_AKS_T2 = Sheets("Visio_Import").Cells(VISIOz, 9)
                VISIO_AKS_T3 = Sheets("Visio_Import").Cells(VISIOz, 10)
                VISIO_AKS_T4 = Sheets("Visio_Import").Cells(VISIOz, 11)
                VISIO_AKS_T5 = Sheets("Visio_Import").Cells(VISIOz, 12)
                Sheets(1).Cells(sad, 7) = VISIO_AKS_T1
                Sheets(1).Cells(sad, 8) = VISIO_AKS_T2
                Sheets(1).Cells(sad, 9) = VISIO_AKS_T3
                Sheets(1).Cells(sad, 10) = VISIO_AKS_T4
                Sheets(1).Cells(sad, 11) = VISIO_AKS_T5
            End If
            
        
        
        Next
    
    Next


End Sub
