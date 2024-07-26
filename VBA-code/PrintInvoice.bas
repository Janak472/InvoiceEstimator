Attribute VB_Name = "Module4"
Sub PrintInvoiceA5()
Attribute PrintInvoiceA5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PrintInvoiceA5 Macro
'

'
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
End Sub
