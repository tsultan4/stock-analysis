# Module 2 | Assignment - Wall Street

Explore green energy stock performance by analyzing financial data using VBA.

1. Overview of Project: Explain the purpose of this analysis.

Steve is doing research for his parents portfolio and he is analyzsing yearly returns for different
stock.  Steve's code works well for a few stock but to see the returns for multiple stocks he will need to 
refactor the code so it is faster when analyzsing many stocks. 

2. Results

The performance of the market between 2017 and 2018 was very different.  Green stocks performed well in 2017 and
gave up most of the gains in 2018, except for ENPH. Refactoring the code reduced the execution times from ~3 
seconds to around .6 second for both years.
![2017 results](/VBA_Challenge_2017.png)
![2018 results](/VBA_Challenge_2018.png)


3. Summary: In a summary statement, address the following questions.

The advantages or disadvantages of refactoring the original VBA code respectively are as follows

	1. The refactoring the code makes it easier to follow.
	2. The code can be changed and is easir to maintain
	3. If done correctly it can reduce the time to run the code.

	1. The refactored code can take a long time to debug if it is
	   not working correctly.  In the end you might end up with the same code
	   and you will have lost time working.



Code
Sub AllStocksAnalysis()


    Dim startTime As Single
    Dim endTime  As Single
    

    yearValue = InputBox("What year would you like to run the analysis on?")

       startTime = Timer
       
   '1.Format the output sheet on the "All Stocks Analysis" worksheet.
    
    Worksheets("AllStocksAnalysis").Activate
    
        Range("A1").Value = "All Stocks (" + yearValue + ")"
    
          'Create a header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"
   
   '2.Initialize an array of all tickers.
   
   Dim tickerIndex As Single
   tickerIndex = 0
   Dim tickers(12) As String
   Dim tickerVolumes(12) As Long
   For k = 0 To 11
   tickerVolumes(k) = 0
   Next k
   Dim tickerStartingPrices(12) As Single
   For l = 0 To 11
   tickerStartingPrices(l) = 0
   Next l
   Dim tickerEndingPrices(12) As Single
   For m = 0 To 11
   tickerEndingPrices(m) = 0
   Next m
   
   
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
   '3.Prepare for the analysis of tickers.
        'Initialize variables for the starting price and ending price.
        
        Dim startingPrice As Double
        Dim endingPrice As Double
        
        'Activate the data worksheet.
        
        Worksheets(yearValue).Activate
        
        'Find the number of rows to loop over.
        
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
   
   '4.Loop through the tickers.
   
   For i = 0 To 11

       ticker = tickers(i)
       totalVolume = 0
   
  
   '5.Loop through rows in the data.
        
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
        
        'Find the total volume for the current ticker.
        
        If Cells(j, 1).Value = ticker Then
        
            totalVolume = totalVolume + Cells(j, 8).Value
         
         End If
        
        'Find the starting price for the current ticker.
        
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
            startingPrice = Cells(j, 6).Value
            
         End If
        
        'Find the ending price for the current ticker.
    
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
            endingPrice = Cells(j, 6).Value
            
         End If
         
         Next j
    
   '6.Output the data for the current ticker.


        Worksheets("AllStocksAnalysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

        Next i
'Formatting
    Worksheets("AllStocksAnalysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
'Formatting
    Worksheets("AllStocksAnalysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
        dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If
        
    Next i

endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
 
  
End Sub

Sub ClearWorksheet()

    Cells.Clear

End Sub

Sub yearValueAnalysis()
 
 yearValue = InputBox("What year would you like to run the analysis on?")
 
 Range("A1").Value = "All Stocks (" + yearValue + ")"
 
 Sheets(yearValue).Activate
 
End Sub
   
