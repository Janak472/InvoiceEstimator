Attribute VB_Name = "Module11"
Option Explicit

Sub Recordofinvoice()
Dim Invo As Long
Dim Custname As String
Dim amt As Currency
Dim dt_issue As Date
Dim nextrec As Range



Invo = Range("G4")
Custname = Range("B5")
amt = Range("G50")
dt_issue = Range("E5")
Set nextrec = Sheet5.Range("A1048576").End(xlUp).Offset(1, 0)

nextrec = Invo
nextrec.Offset(0, 1) = Custname
nextrec.Offset(0, 2) = dt_issue
nextrec.Offset(0, 3) = amt

ThisWorkbook.RefreshAll
 
End Sub
