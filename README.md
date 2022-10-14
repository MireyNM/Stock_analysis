# Stock-Analysis

## Overview of Project
Steve's parents decided to invest all their money into DAQO New Energy Corporation. However, Steve believes that his parent's fund should be more diversified. In this project we are going to analyze several green energy stocks, in addition to DAQO, for the years 2017 and 2018. Moreover, we will refactor the code, written in Visual Basic for Applications (VBA), to make it more efficient for Steve when adding more dataset. 

### Purpose
The aim of this analyze is to compare the original code and the refactored one in order to find out if refactoring the code will successfully make the VBA script run faster. This will help us to determine the advantages and disadvantages of both original and refactored codes. 
 
## Analysis of Data and Results
  
### Analysis of Data Using the Original Code 
The below original code was written in order to visualize the total daily volume and the yearly return for each ticker in our data. 

We have created the first loop to go through all the rows (j=2 to 3013 (found by RowCount function)) in order to:
1. Check if the current row is for the selected stock ticker, then it increases the total Volume for that ticker.
2. Check if the current row is the first row with the selected ticker. If it is, then assign the current starting price for the value in Cells(j, 6)
3. Check if the current row is the last row with the selected ticker. If it is, then assign the current ending price for the value in Cells(j, 6).

In order to do the same for all the tickers (12 tickers), we have created an array ```tickers(11)``` and we have looped for each element in the array the previous loop. Hence, we have created nested loops.

**Original Code**
```
Sub yearValueAnalysis()

 Worksheets("All Stocks Analysis").Activate
 Dim startTime As Single
 Dim endTime  As Single
 yearValue = InputBox("What year would you like to run the analysis on?")
 startTime = Timer

'format the output worksheet
Range("A1").Value = "All Stocks (" + yearValue + ")"
Cells(3, 1).Value = "Ticker"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = "Return"

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
    
    Dim startingPrice As Single
    Dim endingPrice As Single
 
   'Activate data worksheet
    Worksheets(yearValue).Activate

 'Establish the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    

For i = 0 To 11
    
    ticker = tickers(i)
    totalVolume = 0
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
For j = 2 To RowCount
   'Should write ticker not "ticker" to take it as number
        If Cells(j, 1).Value = ticker Then

            'increase totalVolume by the value in the current row
            totalVolume = totalVolume + Cells(j, 8).Value

        End If

        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

            startingPrice = Cells(j, 6).Value

        End If

        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

            endingPrice = Cells(j, 6).Value

        End If
 
  Next j

'MsgBox (totalVolume)
 Worksheets("All Stocks Analysis").Activate
 Cells(4 + i, 1).Value = ticker
 Cells(4 + i, 2).Value = totalVolume
 Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
   Next i


   endTime = Timer
   MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)


End Sub
```
When we run the original code, we get the following stock analysis outputs for the year 2017 and 2018 (See Table 1 and 2). 




<p align="center">
  <img alt="Light" src="https://user-images.githubusercontent.com/109363759/188256387-42e04ec2-f30d-43b3-a578-3437dfc56b09.png" width="45%"> 
&nbsp; &nbsp; &nbsp; &nbsp;
  <img alt="Dark" src="https://user-images.githubusercontent.com/109363759/188255457-0dfbc5fd-e1f7-43b6-abf0-8c106c5df03d.png" width="45%">
</p>

<p align="center">
 Table 1: Stock analysis outputs for the year 2017         
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
 Table 2: Stock analysis outputs for the year 2018
</p>

The pop-up messages showing the running time for the year 2017 and 2018 are shown below (See Fig.1 and Fig.2). We can see that analyzing 2017 stocks takes 0.546875s while analyzing 2018 stocks takes 0.5546875s.

<p align="center">
  <img alt="Light" src="https://user-images.githubusercontent.com/109363759/188256507-09559e84-35e0-4bbc-a39f-1f6bb8837375.png" width="45%"> 
&nbsp; &nbsp; &nbsp; &nbsp;
  <img alt="Dark" src="https://user-images.githubusercontent.com/109363759/188256510-807b6c60-2146-4a73-b011-8839a2de06dd.png" width="45%">
</p>

<p align="center">
 Fig. 1: The running time for year 2017 stock analysis      
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
 Fig. 2: The running time for year 2018 stock analysis 
</p>


### Analysis of Data Using the Refactored Code
In this section we have edited the previous code to loop through all the data one time instead of nested loops. 

To do that, we have created a ```tickerIndex``` variable and set it equal to zero before iterating over all the rows. Then we have created 3 output arrays ```tickerVolumes```, ```tickerStartingPrices```, and ```tickerEndingPrices``` and we have used the ```tickerIndex``` variable 
to access across the different arrays. At the end of our loop we wrote a script to increase the ```tickerIndex``` if the next row’s ticker doesn’t match the previous row’s ticker. 

**Refactored Code** 
```
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    
    '1a) Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
        tickerVolumes(i) = 0
        Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        
    ''3b) Check if the current row is the first row with the selected tickerIndex.
        
        'If  Then
           If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
           tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
           End If
        'End If
        
    ''3c) Check if the current row is the last row with the selected ticker
         'If the next row's ticker and the current row's ticker doesn't match, increase the tickerIndex.
        
        'If  Then
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            End If
            
        'End If
    
     Next i
    
    
    ''4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    Next i
    
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
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```

When we run the refactored  code, we get the same stock analysis outputs, for the year 2017 and 2018, as when using the original code (Table 1 and 2). However, the running time for analysing 2017 stock is now 0.109375s and it is 0.125s for the year 2018 (See Fig.3 and Fig.4) 


<p align="center">
  <img alt="Light" src="https://user-images.githubusercontent.com/109363759/188256950-837cf0c9-d400-43db-999a-6427be9d6579.png" width="45%"> 
&nbsp; &nbsp; &nbsp; &nbsp;
  <img alt="Dark" src="https://user-images.githubusercontent.com/109363759/188256957-f8fc381d-b0a0-4d3d-87bf-88439d0ead71.png" width="45%">
</p>

<p align="center">
 Fig. 3: Running time for year 2017 using refactored code     
&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
 Fig. 4: Running time for year 2018 using refactored code  
</p>



## Summary


- What are the advantages or disadvantages of refactoring code?

According to *Martin Fowler*, who is considered the father of refactoring and Code Smell, [^1]
>"Refactoring is a disciplined technique for restructuring an existing body of code, altering its internal structure without changing its external behavior."

In his book, Martin Fowler (Fowler e al.,2000)[^1] summarised the advantages of refactoring a code as follow: 

1.	Refactoring improves the design of software.
2.	Refactoring makes the code easier to understand and read because it simplifies the code, reduces its complexities and makes it more efficient. 
3.	Refactoring helps finding and fixing bugs and vulnerabilities in the code. Hence it improves maintainability.
4.	Refactoring helps programming faster.

Even if refactoring a code can have a lot of beneficial outcome, yet there are many negative impacts which must be considered.  

1. One important issue is time. While a refactored code is faster than the original one, the process will take extra time doing something that does not add any new feature or functionality in a software system. 
2. Another important factor that should be considered is the cost. Refactoring the code will cost companies to pay money for programmers to go through a process that could be much more cost-effective than preparing an entirely new code structure.
3. Finally, refactoring could be so risky and complex that it can introduce new bugs in the update. 


- How do these pros and cons apply to refactoring the original VBA script? 

On one hand, refactoring our original VBA script has many advantages: 
1. As we can see in Fig.1 and Fig.2, the running time of the original code was around 0.5s for the year 2017 and 2018 stock analysis. After refactoring the code, this time droped to around 0.1s. Therefore, the refactored code is around 5 times faster than the original one.

2. By removing the nested loop and making the code looping through all the data one time, we made the code less complex, easier to read and more efficient. 

On the other hand, it took us more time to get the same stock analysis output (Table 1 and 2).It would not be necessary to go through this process if Steve wants to analyze 2017 and 2018 data only. However, it will save him time and money when adding thousands of stocks to the data.  

In conclusion, there is no definitive answer whether it is better or not to refactor a code. However, I believe every situation is different and one should consider the available cost and the deadlines of a project before starting refactoring a code. 

[^1]: Fowler, M., Beck, K., Brant, J., Opdyke, W., Roberts, D. Refactoring: Improving the Design of Existing Code, 1999

