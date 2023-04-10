Attribute VB_Name = "Module1"
Sub MultipleYearTesting():

For Each ws In Worksheets

Dim WorksheetName As String
Dim x As Long
Dim y As Long
Dim TicketCount As Long
Dim LastARow As Long
Dim LastIRow As Long
Dim PercentChange As Double
Dim GreatIn As Double
Dim GreatDe As Double
Dim GreatVolTotal As Double

WorksheetName = ws.Name

ws.Cells(1, 9).Value = "Ticker Symbol"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

ws.Cells(2, 15).Value = "Greatest Percent Increase"
ws.Cells(3, 15).Value = "Greate Percent Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"


TicketCount = 2     'where data starts
y = 2

'data in last cell of that column
LastARow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For x = 2 To LastARow
If ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then
ws.Cells(TicketCount, 9).Value = ws.Cells(x, 1).Value

ws.Cells(TicketCount, 10).Value = ws.Cells(x, 6).Value - ws.Cells(y, 3).Value

'conditional formatting
If ws.Cells(TicketCount, 10).Value < 0 Then
ws.Cells(TicketCount, 10).Interior.ColorIndex = 3 'set it red

Else
ws.Cells(TicketCount, 10).Interior.ColorIndex = 4 'set it green

End If

'percent change
If ws.Cells(y, 3).Value <> 0 Then
PercentChange = ((ws.Cells(x, 6).Value - ws.Cells(y, 3).Value) / ws.Cells(y, 3).Value)


ws.Cells(TicketCount, 11).Value = Format(PercentChange, "Percent")

Else
ws.Cells(TicketCount, 11).Value = Format(0, "Percent")

End If

'total volume
ws.Cells(TicketCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(y, 7), ws.Cells(x, 7)))

TicketCount = TicketCount + 1

y = x + 1

End If

Next x

'getting last cell in column with data
LastIRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

GreatIn = ws.Cells(2, 12).Value
GreatDe = ws.Cells(2, 11).Value
GreatVolTotal = ws.Cells(2, 11).Value

For x = 2 To LastIRow


If ws.Cells(x, 11).Value > GreatIn Then
GreatIn = ws.Cells(x, 11).Value
ws.Cells(2, 16).Value = ws.Cells(x, 9).Value

Else

GreatIn = GreatIn

End If

If ws.Cells(x, 11).Value < GreatDe Then
GreatDe = ws.Cells(x, 11).Value
ws.Cells(3, 16).Value = ws.Cells(x, 9).Value

Else

GreatDe = GreatDe

End If

If ws.Cells(x, 12).Value > GreatVolTotal Then
GreatVolTotal = ws.Cells(x, 12).Value
ws.Cells(4, 16).Value = ws.Cells(x, 9).Value

Else

GreatVolTotal = GreatVolTotal

End If

ws.Cells(2, 17).Value = Format(GreatIn, "Percent")      'putting a percentage sign
ws.Cells(3, 17).Value = Format(GreatDe, "Percent")
ws.Cells(4, 17).Value = Format(GreatVolTotal, "Scientific")     'scientific is expotential notation

Next x


Next ws

End Sub

