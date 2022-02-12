Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
    Dim tickerArraySise As Integer
    Dim yearValue As String
    
    ' assign the input value of Year to variable
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    'Logic to check if the entered sheet name exists in the workbook or not
    Dim ws As Worksheet
    Dim invalidSheet As Boolean
    invalidSheet = True
    
    ' Looping through each worksheet and checking the name of sheet against the entered value
    For Each ws In Worksheets
        If (ws.Name = yearValue And ws.Index < Worksheets.Count) Then
           invalidSheet = False
           Exit For
        End If
    Next ws
    
    If (invalidSheet = True) Then
            MsgBox "Could not find the sheet " & yearValue
            Exit Sub
    End If
    
    'Starting the timer to calculate the time taken to rin the analysis and update "AllStocksAnalysis" sheet
    startTime = Timer
    
    ' Get Count of distinct ticker values to set the array size of ticer
    tickerArraySise = CountOfUniqueTickers(yearValue)
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("AllStocksAnalysis").Activate
    
    Range("A1").Value = "AllStocks (" + yearValue + ")"
    
    'Create a header row
    Cells(2, 1).Value = "Ticker"
    Cells(2, 2).Value = "Total Daily Volume"
    Cells(2, 3).Value = "Return"

    'Initialize array of all tickers, using the varibale to get the count of unique tickers
    Dim tickers() As String
    
    ' ReDim to initialize the array with a varaiable
    ReDim tickers(tickerArraySise)
    
    
   ' NOT Planning to use this hard coded array
    ' tickers(0) = "AY"
    ' tickers(1) = "CSIQ"
    ' tickers(2) = "DQ"
    ' tickers(3) = "ENPH"
    ' tickers(4) = "FSLR"
    ' tickers(5) = "HASI"
    ' tickers(6) = "JKS"
    ' tickers(7) = "RUN"
    ' tickers(8) = "SEDG"
    ' tickers(9) = "SPWR"
    ' tickers(10) = "TERP"
    ' tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim tickerIndex  As Integer
    tickerIndex = 0

    '1b) Create three output arrays, using the varibale to get the count of unique tickers
    Dim tickerVolumes() As Long
    Dim tickerStartingPrices() As Single
    Dim tickerEndingPrices() As Single
    
    ' ReDim to initialize the array with a varaiable
    
    ReDim tickerVolumes(tickerArraySise)
    ReDim tickerStartingPrices(tickerArraySise)
    ReDim tickerEndingPrices(tickerArraySise)
    
    ' programatically Getting Ticker, Volume and Close Column Index
    tickerCol = Worksheets(yearValue).Cells.Find("Ticker", searchorder:=xlByColumns, searchdirection:=xlPrevious).Column
    volumeCol = Worksheets(yearValue).Cells.Find("Volume", searchorder:=xlByColumns, searchdirection:=xlPrevious).Column
    closingColumn = Worksheets(yearValue).Cells.Find("Close", searchorder:=xlByColumns, searchdirection:=xlNext).Column
    tickerColumnName = Worksheets(yearValue).Cells.Find("Ticker", searchorder:=xlByColumns, searchdirection:=xlPrevious)
    
     ' Populate array for Tickers (dynamically). I decided not to use the existing ticker array. Rather read from Worksheet and populate the ticker array
    For i = 2 To RowCount
        For tickerIndex = 0 To tickerArraySise
            If (Worksheets(yearValue).Cells(i, 1).Value <> Worksheets(yearValue).Cells(i - 1, 1)) Then
                If (tickers(tickerIndex) <> Worksheets(yearValue).Cells(i, 1).Value And tickers(tickerIndex) = "" And tickers(tickerIndex) <> tickerColumnName) Then
                    tickers(tickerIndex) = Worksheets(yearValue).Cells(i, 1).Value
                    Exit For
                End If
            End If
        Next tickerIndex
    Next
    
    ''2a) Create a for loop to initialize the arrays tickerVolumes, tickerStartingPrices and tickerEndingPrices to zero.
    For tickerIndex = 0 To UBound(tickers) - LBound(tickers)
        tickerVolumes(tickerIndex) = 0
        tickerStartingPrices(tickerIndex) = 0
        tickerEndingPrices(tickerIndex) = 0
    Next tickerIndex
    
    'After initializing the tickerVolumes to zero, need to set the ticker index back to zero becuase we rae using the same variable in the For loop below
    tickerIndex = 0
      
      
      ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
            '3a) Increase volume for current ticker
            If (Worksheets(yearValue).Cells(i, tickerCol).Value = tickers(tickerIndex)) Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Worksheets(yearValue).Cells(i, volumeCol).Value
            End If
            
            '3b) Check if the current row is the first row with the selected tickerIndex.
             ' assigning starting price to tickerStartingPrices for each ticker
            If (Worksheets(yearValue).Cells(i - 1, tickerCol).Value <> tickers(tickerIndex) And Worksheets(yearValue).Cells(i, tickerCol).Value = tickers(tickerIndex)) Then
                tickerStartingPrices(tickerIndex) = Worksheets(yearValue).Cells(i, closingColumn).Value
            End If
            
            ' assigning ending Price to tickerEndingPrices for each ticker
            If (Worksheets(yearValue).Cells(i + 1, tickerCol).Value <> tickers(tickerIndex) And Worksheets(yearValue).Cells(i, tickerCol).Value = tickers(tickerIndex)) Then
                tickerEndingPrices(tickerIndex) = Worksheets(yearValue).Cells(i, closingColumn).Value
            End If
            
            '3c) check if the current row is the last row with the selected ticker
               'If the next row ticker doesn't match, increase the tickerIndex.
              If (Worksheets(yearValue).Cells(i, tickerCol).Value <> Worksheets(yearValue).Cells(i + 1, tickerCol).Value And i < RowCount) Then
                tickerIndex = tickerIndex + 1
              End If
        Next i

    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
     Worksheets("AllStocksAnalysis").Activate
     
     'get  column index in sheet programatically
     
    tickerCol = Cells.Find("Ticker", searchorder:=xlByColumns, searchdirection:=xlPrevious).Column
    totalVolColumn = Cells.Find("Total Daily Volume", searchorder:=xlByColumns, searchdirection:=xlPrevious).Column
    returnColumn = Cells.Find("Return", searchorder:=xlByColumns, searchdirection:=xlNext).Column
     
     'Will NOT use i =0 to 11, i'll rather use UBound and Lboud to get the tickers array length and then loop through
     
       For tickerIndex = 0 To UBound(tickers) - LBound(tickers)
            
           Cells(tickerIndex + 3, tickerCol).Value = tickers(tickerIndex)
           Cells(tickerIndex + 3, totalVolColumn).Value = tickerVolumes(tickerIndex)
           Cells(tickerIndex + 3, returnColumn).Value = (tickerEndingPrices(tickerIndex) - tickerStartingPrices(tickerIndex)) / tickerStartingPrices(tickerIndex)
            
       Next tickerIndex
    
    'Formatting
    Worksheets("AllStocksAnalysis").Activate
    Range("A1:C2").Font.FontStyle = "Bold"
    Range("A1:C2").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B3:B14").NumberFormat = "#,##0"
    Range("C3:C14").NumberFormat = "0.0%"
    Columns("B").AutoFit
    Columns("C").AutoFit
    
    'Additional code to merge the Top column header, Center Align and Font Formatting
    Range("A1:C1").MergeCells = True
    Range("A1:C1").VerticalAlignment = xlCenter
    Range("A1:C1").HorizontalAlignment = xlCenter
    Range("A1:C1").Font.Size = 14
    Range("A1:C1").Interior.Color = vbYellow
    
    dataRowStart = 3
    dataRowEnd = 14

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
    'Ending the timer to calculate the time taken to rin the analysis and update "AllStocksAnalysis" sheet and displaying time taken in message box
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

' To clear contents of the "AllStocksAnalysis" worksheet
Sub ClearWorksheet()
    ' Clear Year from Top Header and Calculated data
    Range("A1:c1").Value = "AllStocks(xxxx)"
    Range("A3:C14").Clear

End Sub


Function CountOfUniqueTickers(yearValue As String) As Integer
    'Active the year sheet
    Worksheets(yearValue).Activate
    
    'Declare variable to hold row cound, range and list object
    Dim LstRw As Long
    Dim Rng As Range
    Dim List As Object
    
    'Getting row count in the sheet
    LstRw = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Creating a list object
    Set List = CreateObject("Scripting.Dictionary")
    
    ' Adding unique ticker to the list
    For Each Rng In Range("A2:A" & LstRw)
      If Not List.Exists(Rng.Value) Then List.Add Rng.Value, Nothing
    Next
    
    ' Assigning List.Count - 1 to the return function becuase array position start at zero.
    CountOfUniqueTickers = List.Count - 1
    
End Function


