Sub StockData()
Dim Ticker As Integer
Dim TotalSV As Double
Dim Lastrow As String
Dim Lastcol As Long
Dim ws As Worksheet
Dim WS_Count As Integer
    For Each ws In ActiveWorkbook.Worksheets
    	ws.activate	
	    Lastrow = Range("A1").End(xlDown).Row
		'Lastcol will be useful only if I want to automatize the document'
		Lastcol = Range("A1").End(xlToRight).Select
		Ticker = 1
		TotalSV = 0
		Cells(1, 9).Value = "Ticker"
		Cells(1, 16).Value = "Ticker"
		Cells(1, 10).Value = "Yearly Change"
		Cells(1, 11).Value = "Percent Change"
		Cells(1, 12).Value = "Total Stock Volume"
		Cells(1, 17).Value = "Value"
		Cells(2, 15).Value = "Greatest % Increase"
		Cells(3, 15).Value = "Greatest % Decrease"
		Cells(4, 15).Value = "Greatest Total Volume"
	    For i = 2 To Lastrow
	        If Cells(i, 1).Value = Cells(i + 1, 1) Then
	        	TotalSV = TotalSV + Cells(i, 7).Value
	    	Else
	    		TotalSV = TotalSV + Cells(i, 7).Value
		        Ticker = Ticker + 1
		        Cells(Ticker, 9).Value = Cells(i, 1).Value
		        Cells(Ticker, 10).Value = Cells(Ticker, 3).Value - Cells(i, 3).Value
	            If Cells(Ticker, 3).Value - Cells(i, 3).Value < 0 Then
	            	Cells(Ticker, 10).Interior.ColorIndex = 3
                Else
                	Cells(Ticker, 10).Interior.ColorIndex = 4
            	End If
		        Cells(Ticker, 11).Value = (Cells(i, 3).Value / Cells(Ticker, 3).Value) - 1
		        Cells(Ticker, 11).NumberFormat = "0.00%"
		        Cells(Ticker, 12).Value = TotalSV
      		End If
    	Next i
		Range("Q2").Value = WorksheetFunction.Max(Range("K2:K290"))
		Range("Q2").NumberFormat = "0.00%"
		Range("Q3").Value = WorksheetFunction.Min(Range("K2:K290"))
		Range("Q3").NumberFormat = "0.00%"
		Range("Q4").Value = WorksheetFunction.Max(Range("L2:L290"))
    	For j = 2 To Lastrow
        	If Cells(j, 11).Value = Cells(2, 17).Value Then
        		Cells(2, 16).Value = Cells(j, 9).Value
            ElseIf Cells(j, 11).Value = Cells(3, 17).Value Then
            	Cells(3, 16).Value = Cells(j, 9).Value
         	ElseIf Cells(j, 12).Value = Cells(4, 17).Value Then
                Cells(4, 16).Value = Cells(j, 9).Value
        	End If
    	Next j
	Next ws    
End Sub
