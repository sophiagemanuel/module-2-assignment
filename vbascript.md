
Sub multipleyeardata():
'Making our code reach all of our worksheets
For Each ws In Worksheets
Dim worksheetname As String

    'creates all our variable

Dim i As Long
Dim j As Long
Dim lastrowa As Long
Dim lastrowi As Long
Dim percentchange As Double
Dim greatestinc As Double
Dim greatestdec As Double
Dim greatestvol As Double

worksheetname = ws.Name

'creating our column headers for our scripts
ws.Cells(1, 9).Value = "ticker"
ws.Cells(1, 10).Value = "yearly change"
ws.Cells(1, 11).Value = "percent change"
ws.Cells(1, 12).Value = "total stock volume"

'creates headers for the greatest inc, dec, and total vol
ws.Cells(1, 16).Value = "ticker"
ws.Cells(1, 17).Value = "value"
ws.Cells(2, 15).Value = "greatest % increase"
ws.Cells(3, 15).Value = "greatest % decrease"
ws.Cells(4, 15).Value = "greatest total volume"

lastrowa = ws.Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox (lastrowa)
tickcount = 2
j = 2
For i = 2 To lastrowa

'checking if the ticker name changed and if it did wrote it in column j
If Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ws.Cells(tickcount, 9).Value = ws.Cells(i, 1).Value
ws.Cells(tickcount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value

'sets cell color to red if negative percent and green if positive percent
    If ws.Cells(tickcount, 10).Value < 0 Then
    ws.Cells(tickcount, 10).Interior.ColorIndex = 3
    Else
    ws.Cells(tickcount, 10).Interior.ColorIndex = 4
    End If
'calculating percent change and formatting it into percentages
    If ws.Cells(j, 3).Value <> 0 Then
    percentchange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
    ws.Cells(tickcount, 11).Value = Format(percentchange, "percent")
    Else
    ws.Cells(tickcount, 11).Value = Format(0, "percent")
    End If
'calculating total volume
ws.Cells(tickcount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))

'increasing tickcount by one and new start row of ticker
tickcount = tickcount + 1
j = i + 1
    End If
Next i


'finding last cell in column i and preparting for summmary
lastrowi = ws.Cells(Rows.Count, 9).End(xlUp).Row
greatestvol = ws.Cells(2, 12).Value
greatestinc = ws.Cells(2, 11).Value
greatestdec = ws.Cells(2, 11).Value

'creating the loop for the summary:

For i = 2 To lastrowi
    'creating if statement where for the greatest volume we check if the next is larger and if it is take the new value and put it in cell (4,16)
    If ws.Cells(i, 12).Value > greatestvol Then
    greatestvol = ws.Cells(i, 12).Value
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    Else
    greatestvol = greatestvol
    End If
   'creating if statement where for the greatest decrease we check if the next is smallerr and if it is take the new value and put it in cell (2,16)
    If ws.Cells(i, 11).Value < greatestdec Then
    greatestdec = ws.Cells(i, 11).Value
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    Else
    greatestdec = greatestdec
    End If
   'creating if statement where for the greatest increase we check if the next is larger and if it is take the new value and put it in cell (3,16)
    If ws.Cells(i, 11).Value > greatestinc Then
    greatestdec = ws.Cells(i, 11).Value
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    Else
    greatestinc = greatestinc
    End If
    
    
   'writing our summary results in our cells with the greatest increase and decrease as percentages and the greatest volume as scientific as it is a larger number
    ws.Cells(2, 17).Value = Format(greatestinc, "percent")
    ws.Cells(3, 17).Value = Format(greatestdec, "percent")
    ws.Cells(4, 17).Value = Format(greatestvol, "scientific")
    Next i
   
   'Adjusting column width automatically to fit
    Worksheets(worksheetname).Columns("A:Z").AutoFit
   

Next ws
End Sub
