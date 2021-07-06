Sub PmtCalc()
Dim IntRate As Double
Dim LoanAmt As Double
Dim Periods As Long
IntRate = 0.0625 / 12
Periods = 30 * 12
LoanAmt = 150000
MsgBox WorksheetFunction.PMT(IntRate, Periods, -LoanAmt)
End Sub
