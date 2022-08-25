# VBA Challenge - Stocks Analysis
## **Project Overview:**

In Module 2, we used stock information to determine which stocks Steve's parents should invest in. The data contains stock information on 12 different stocks across years 2017 and 2018 and includes date, opening price, highest & lowest prices, closing and adjusted closing prices. Using this data, we created a VBA code that helped us determine the total volume and return percent rate for each of the stocks in 2017 and 2018. 

In this specific exercise, we took the VBA code we developed in Module 2 and refactored the solution code to determine whether we could get the same information in a quicker way. 

***

## **Results:**
Below is the refactored code that was used

    Sub YearValueAnalysis()

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
        Dim tickers(12) As String
        
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
        
        '1a) Create a ticker Index and set to zero before looping rows.
        Dim tickerindex As Single
        tickerindex = 0
        
        '1b) Create three output arrays for tickervolumes, tickerstartingprices, and tickerendingprices.
        Dim tickervolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
        
        '2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
            tickervolumes(i) = 0
            tickerStartingPrices(i) = 0
            tickerEndingPrices(i) = 0

        Next i
            
        '2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount

            '3a) Increase volume for current ticker
            tickervolumes(tickerindex) = tickervolumes(tickerindex) + Cells(i, 8).Value
            
            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i, 1).Value = tickers(tickerindex) And Cells(i - 1, 1).Value <> tickers(tickerindex) Then
                tickerStartingPrices(tickerindex) = Cells(i, 6).Value
            End If
                
            '3c) check if the current row is the last row with the selected ticker. If the next row’s ticker doesn’t match, increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
                tickerEndingPrices(tickerindex) = Cells(i, 6).Value
            End If
        
            '3d Increase the tickerindex if the next row’s ticker doesn’t match the previous row’s ticker.
            If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
                tickerindex = tickerindex + 1
            End If

        Next i
        
        '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11
            
            Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickervolumes(i)
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






## **Analysis of Stock Performance:**

<img width="225" alt="Screen Shot 2022-08-24 at 8 27 04 PM" src="https://user-images.githubusercontent.com/110875578/186567911-dd0ed127-9cd3-401f-aec5-5a504e9a1062.png"> <img width="229" alt="Screen Shot 2022-08-24 at 8 26 43 PM" src="https://user-images.githubusercontent.com/110875578/186567900-d5f775c9-77d5-4734-9129-fab8e2ab2599.png">

Based on the data, the recommended stocks for Steve's parents to invest in would be RUN and ENPH. These two stocks were the only ones that maintained a positive return in 2018, while the remaining stocks took a drop. 
1) RUN had a 5.5% return in 2017 but jumped to an 84.0% return in 2018.
2) ENPH had a 129.5% return in 2017 and 81.9% return in 2018. 

***
## **Summary:**

**What are the advantages or disadvantages of refactoring code?** Refactoring allows for a cleaner and more efficient code by using less memory and taking less steps. Some of the advantages of refactoring code include improving a design of software/application, debugging, and making a program run faster. The disadvantages of refactoring code can include the amount of time it can take to refactor and imposing a higher risk of imprecise coding which could result in bugs and errors.

**How do these pros and cons apply to refactoring the original VBA script?** The benefit that resulted from refactoring the original VBA code is the decrease in run time of the macro. It can be seen that the time to run the code for both 2017 and 2018 decreased about 0.2 seconds with the new refactored code.

2017 VBA Code (Old Version)* - Run time 0.29 sec

<img width="225" alt="greenstocks2017" src="https://user-images.githubusercontent.com/110875578/186568928-6b5555b4-2066-4e75-af2e-b61276940cba.jpeg"> 

2017 VBA Code (Refactored Version)* - Run time 0.07 sec

<img width="225" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/110875578/186568949-fbf90c80-719f-4cdc-94a0-9c54ccdb9db4.jpeg">

2018 VBA Code (Old Version) - Run time 0.29 sec

<img width="225" alt="greenstocks2018" src="https://user-images.githubusercontent.com/110875578/186569230-c3813e6b-b4a8-4207-a20c-730c3c471679.jpeg"> 

2018 VBA Code (Refactored Version)* - Run time 0.08 sec

<img width="225" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/110875578/186569235-8d7da0a0-9db9-468b-a333-33cb61c413e4.jpeg">


***
## **Associated Links:**
Link to VBA_Challenge Excel File https://github.com/ireneshin26/stocks-analysis/blob/main/VBA_Challenge.xlsm

Link to VBA_Challenge Resources Folder https://github.com/ireneshin26/stocks-analysis/tree/main/Resources
