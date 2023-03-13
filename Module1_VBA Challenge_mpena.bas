Attribute VB_Name = "Module1"
Sub Sheet_Loop()
 
 ' Declare Current as a worksheet object variable.
         Dim ws As Worksheet
         
         ' Loop through all of the worksheets in the active workbook.
         For Each ws In Worksheets
    
    'code that you want to run on each sheet-------------------------------------------------------------------

'Begin Project Coding

'declarations
Dim rowcount As LongLong 'declares - rowcount - as a long integer (64 bit)
Dim tickercount As LongLong 'declares - tickercount - as a long integer (64 bit)

'variable assignments
rowcount = ws.Cells(Rows.Count, 1).End(xlUp).Row 'assigns rowcount to an expression; is like pressing Ctrl + up
tickercount = 2 'assigns tickercount to a value

'Starts For loop
For I = 2 To rowcount
    If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1) Then
    ws.Cells(tickercount, 9).Value = ws.Cells(I, 1).Value ' returns ticker value
    ws.Cells(tickercount, 10).Value = ws.Cells(I, 6).Value - ws.Cells(I - 250, 3).Value ' returns yearly Chg value
    ws.Cells(tickercount, 11).Value = (ws.Cells(I, 6).Value / ws.Cells(I - 250, 3).Value) - 1 ' returns yearly % Chg value
    ws.Cells(tickercount, 12).Value = WorksheetFunction.Sum(ws.Range(ws.Cells(I, 7), ws.Cells(I - 250, 7))) 'sums vol

'Conditional color formatting
    CondColor = ws.Cells(tickercount, 10).Value
    Select Case CondColor
    Case Is > 0
    ws.Cells(tickercount, 10).Interior.ColorIndex = 4
    Case Is < 0
    ws.Cells(tickercount, 10).Interior.ColorIndex = 3
    Case Else
    ws.Cells(tickercount, 10).Interior.ColorIndex = 0
    End Select
    tickercount = tickercount + 1
End If

'Number formatting
    ws.Cells(I, 10).NumberFormat = "$#,##0.00"
    ws.Cells(I, 11).NumberFormat = "0.0%"
    ws.Cells(I, 12).NumberFormat = "#,##0"

'Ends For loop
Next I

'Column headings
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Chg."
    ws.Range("K1").Value = "% Chg."
    ws.Range("L1").Value = "Vol."

'Summary Data headings
    ws.Range("N2").Value = "Greatest % Inc."
    ws.Range("N3").Value = "Greatest % Dec."
    ws.Range("N4").Value = "Greatest Tot. Vol."
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"

'Summary Data bumber formating
    ws.Cells(2, 16).NumberFormat = "0.0%"
    ws.Cells(3, 16).NumberFormat = "0.0%"
    ws.Cells(4, 16).NumberFormat = "#,##0"

'Summary Data Greatest % Inc/Dec and Vol.: Values
    ws.Range("P2") = Application.WorksheetFunction.Max(ws.Range("K:K"))
    ws.Range("P3") = Application.WorksheetFunction.Min(ws.Range("K:K"))
    ws.Range("P4") = Application.WorksheetFunction.Max(ws.Range("L:L"))

'Summary Data Greatest % Inc/Dec and Vol.: Ticker
    MaxIndex = WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Range("K:K")), ws.Range("K:K"), 0)
    ws.Range("O2") = ws.Cells(MaxIndex, 9)
    MinIndex = WorksheetFunction.Match(Application.WorksheetFunction.Min(ws.Range("K:K")), ws.Range("K:K"), 0)
    ws.Range("O3") = ws.Cells(MinIndex, 9)
    VolIndex = WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Range("L:L")), ws.Range("L:L"), 0)
    ws.Range("O4") = ws.Cells(VolIndex, 9)

'End project coding

'-----------------------------------------------------------------------------------------------------------------
 Next ws

End Sub

