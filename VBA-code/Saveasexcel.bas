Attribute VB_Name = "Module9"
Sub Saveinvoiceexcel()
Dim path As String
path = "C:\Users\USER\Desktop\invoice data\XlSX\"

Dim Invo As Long
Invo = Range("G4")

Dim fname As String
fname = Range("B5") & "_" & Invo
Application.DisplayAlerts = False

Sheet1.Copy

With ActiveWorkbook
    .SaveAs Filename:=path & fname, FileFormat:=51
    .Close
End With

Application.DisplayAlerts = True

End Sub
