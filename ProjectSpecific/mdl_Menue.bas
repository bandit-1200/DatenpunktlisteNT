Attribute VB_Name = "mdl_Menue"

Sub Ribbon_Filter_alles_an(contol As IRibbonControl)
    Call Filter_alles_an
End Sub

Sub Ribbon_Filter_PG5(contol As IRibbonControl)
    Call Filter_PG5
End Sub

Sub Ribbon_Filter_BMKZ(contol As IRibbonControl)
    Call Filter_BMKZ
End Sub

Sub Ribbon_Filter_PromosNT(contol As IRibbonControl)
    Call Filter_PromosNT
End Sub

Sub Ribbon_Filter_Inbetriebnahme(contol As IRibbonControl)
    Call Filter_Inbetriebnahme
End Sub
'refresh_tabelle
Sub Ribbon_refresh_tabelle(contol As IRibbonControl)
    Call refresh_tabelle
End Sub

Sub Ribbon_Open_Slotmanager(contol As IRibbonControl)
    Call Open_Slotmanager
End Sub

Sub Ribbon_open_ImportCFG(contol As IRibbonControl)
    Call open_ImportCFG
End Sub

Sub Ribbon_open_PromosNT_Menu(contol As IRibbonControl)
    Call open_PromosNT_Menu
End Sub

Sub Ribbon_BMKZ_cfg_open(contol As IRibbonControl)
    Call BMKZ_cfg_open
End Sub

Sub Ribbon_AKS_AutoFill(contol As IRibbonControl)
    Call AKS_AutoFill
End Sub

Sub Ribbon_open_Import_perUserForm(contol As IRibbonControl)
    Call open_Import_perUserForm
End Sub

Sub Ribbon_Trennlinie(contol As IRibbonControl)
    Call Trennlinie
End Sub

Sub Ribbon_InbetriebnahmeProtokoll(contol As IRibbonControl)
    Call InbetriebnahmeProtokoll
End Sub

Sub Ribbon_Projekt_Beschriftung(contol As IRibbonControl)
    Call Projekt_Beschriftung
End Sub

Sub Ribbon_Form_BMKZ_sync(contol As IRibbonControl)
     Form_BMKZ_sync.Show
End Sub

Sub Ribbon_Form_Projekt(contol As IRibbonControl)
     Form_Projekt.Show
End Sub

Sub Ribbon_Form_Blatmanager(contol As IRibbonControl)
     Form_Blatmanager.Show
End Sub

Sub Ribbon_E_SchemaEinmischen(contol As IRibbonControl)
     Call E_SchemaEinmischen
End Sub

Sub Ribbon_E_Schema_loeschen(contol As IRibbonControl)
     Call E_Schema_loeschen
End Sub

Sub Ribbon_Filter_PromosObjekte(contol As IRibbonControl)
     Call Filter_PromosObjekte
End Sub

Sub Ribbon_Update(contol As IRibbonControl)
     MsgBox "Funktion noch in Arbeit!", vbInformation, "Update"
     'call ImportModulesWarn
End Sub









