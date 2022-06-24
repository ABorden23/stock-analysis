## VBA Stock Analysis

![This is an Image](https://github.com/ABorden23/stock-analysis/blob/main/istockphoto-1150316505-612x612.png)

# Overview of Project: 

In this project we have prepared a high level workbook for a client named Steve. At the click of a button, we can analyze an entire dataset. After presenting this analysis, Steve now wants a new an improved dataset that can anazlyze thousands of stocks instead of hundreds. Using the Visual Basic Application (VBA) we have created a "Refactored" Worksheet Button that will help seeing multiple stocks faster and more efficient. 

Purpose: To make looking at a Stock Data Set easier and more efficient to read. 

# Results: 

Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

**All Stocks Analysis (Orignial Code)**

```
Sub AllStocksAnalysis()

    yearValue = InputBox("What year would you like to run the analysis on?")

    Dim startTime As Single
    Dim endTime  As Single

       startTime = Timer
       
   '1) Format the output sheet on All Stocks Analysis worksheet
   Worksheets("All Stocks Analysis").Activate
   Range("A1").Value = "All Stocks (" + yearValue + ")"

   'Create a header row
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

   '2) Initialize array of all tickers
   
   Dim tickers(11) As String
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
   
   '3a) Initialize variables for starting price and ending price
   Dim startingPrice As Single
   Dim endingPrice As Single
   '3b) Activate data worksheet
   
    Worksheets(yearValue).Activate
        
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
            Worksheets(yearValue).Activate
        For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i
   
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

 'Formatting
   Worksheets("All Stocks Analysis").Activate
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

```

In the process of making my code more efficent, it was imparative that I changed the order of my loops. 

By creating a Ticker Index I created 3 output arrays:

* TickerVolumes

* TickerStartingPrices

* TickerEndingPrices

**Refactored Code**

```


```   
# Code Run Times:

**Original Code**

*2017*

![This is an image](https://github.com/ABorden23/stock-analysis/blob/main/All%20Stocks%20Analysis%202017%20.png)

*2018*

![This is an image](https://github.com/ABorden23/stock-analysis/blob/main/All%20Stocks%20Analysis%202017%20.png)

**Refactored Code**

*2017*

![This is an image](https://github.com/ABorden23/stock-analysis/blob/main/All%20Stocks%20Analysis%202017%20.png)

*2018*

![This is an image](https://github.com/ABorden23/stock-analysis/blob/main/All%20Stocks%20Analysis%202017%20.png)

Conclusion of Code: Using our Refactored Code, we are able to run our 2017 code 75% faster and our 2018 82% faster. 

# Summary: 

Using VBA is an extremely strong tool I have found... although vastly tedious. It will run some powerful analysis and is very useful. What I found throughout this project is that when writing the original code then backtracking to refactor. You must watch your steps and think carefully about each move you are making, unless you will end up with multiple errors. Without a strong backround in coding, this can be quite a learning curve and requires fundamentals to be learned to understand some of these commands.


