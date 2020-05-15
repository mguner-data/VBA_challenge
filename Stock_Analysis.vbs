Sub stock()


'Loop Through All Sheets


For Each ws In Worksheets

'Determine last row in each worksheet
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim column As Integer
Dim i As Long
Dim o As Long
Dim j As Long
Dim k As Long
Dim tickercounter As Long
Dim close_price As Double
Dim start_price As Double
Dim myRange
Dim Results
Dim percent_change

'Printing Column headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'Loop identifying tickers and saving

tickercounter = 0

column = 1


'tickercounter counts the number of tickers
'Loop through the rows in the column A and combine them in column I
tickercounter = 0
    For i = 3 To lastrow
    

        If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
        tickercounter = tickercounter + 1
        ws.Cells((tickercounter + 1), 9).Value = ws.Cells(i, column).Value
        End If
   
    Next i
 

'Calculations for yearly change, percent change and total volume
'Storing indices for calculations
'o counts the rows for new table where the tickers are reserved
o = 2
'k counts the start point of each ticker
k = 2

start_price = ws.Cells(2, 3)

For j = 2 To lastrow
    
    If ws.Cells(j, 1).Value <> ws.Cells(j + 1, 1).Value Then

    close_price = ws.Cells(j, 6)
'If statement to get rid of divided by zero error
        If start_price = 0 Then
        difference = close_price - start_price
        percent_change = "  NULL"
        
        Else
        difference = close_price - start_price
        percent_change = (difference / start_price)
        End If
    
    ws.Cells(o, 10).Value = difference
'If statement to color the yearly change column
 
    If difference <= 0 Then
    ws.Cells(o, 10).Interior.ColorIndex = 3
    ElseIf difference > 0 Then
    ws.Cells(o, 10).Interior.ColorIndex = 4
    End If
        
     
    ws.Cells(o, 11).Value = percent_change
    
'Summing the total volumes per ticker name
    
    myRange = ws.Range("G" & k, "G" & j)
        
        Results = WorksheetFunction.Sum(myRange)
        ws.Range("L" & o) = Results
       

   
    o = o + 1
'k updates the start point of each ticker
    k = j + 1
    
    start_price = ws.Cells(k, 3)
    
    End If
    
Next j

'BONUS CHALLENGE

'Headers for Greatest Selection
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Finding the Greatest Yearly %Increase and Greatest %Decrease

g = 2
newlastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

For m = g To newlastrow

    myMAXrange = ws.Range("K" & g, "K" & newlastrow)
    Max = WorksheetFunction.Max(myMAXrange)
    Min = WorksheetFunction.Min(myMAXrange)
    ws.Range("Q2") = Max
    ws.Range("Q3") = Min
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    myMAXvalue = ws.Range("L" & g, "K" & newlastrow)
    TotalValue = WorksheetFunction.Max(myMAXvalue)
    ws.Range("Q4") = TotalValue

Next m

w = 2
'Finding Ticker for Greatest Increase

For t = w To newlastrow

If ws.Range("Q2").Value = ws.Cells(w, 11).Value Then

ws.Range("P2").Value = ws.Cells(w, 9).Value

Exit For

End If

w = w + 1

Next t
'Finding Ticker for Greatest Decrease
w = 2

For t = w To newlastrow

If ws.Range("Q3").Value = ws.Cells(w, 11).Value Then

ws.Range("P3").Value = ws.Cells(w, 9).Value

Exit For

End If

w = w + 1

Next t

'Finding Ticker for Greatest Total Volume

w = 2

For t = w To newlastrow

If ws.Range("Q4").Value = ws.Cells(w, 12).Value Then

ws.Range("P4").Value = ws.Cells(w, 9).Value

Exit For

End If

w = w + 1

Next t


'AutoFit Every WorksheetColumn
ws.Cells.EntireColumn.AutoFit

Next ws


End Sub






