Attribute VB_Name = "StockData"
Option Explicit

Sub StockMarketData02():

'********************************************************************
'**This subroutine steps through 3 worksheets containing           **
'**daily stock prices for numerous stocks. The code                **
'**calculates the annual volume, the annual change in price,       **
'**and the annual percentage change for each stock. The            **
'**code also determines the stock with the greatest annual change, **
'**the stock with the lowest annual change, and the stock with     **
'**the greatest trading volume.                                    **
'**                                                                **
'**Lasted edited by Justin Bein, 02/12/2024                        **
'********************************************************************


'********************************
'**variable declaration section**
'********************************


Dim Ticker As String                    'the stock's ticker symbol
Dim TickerDate As Integer               'the date for each ticker symbol's prices
Dim ResultsTableRow As Integer          'keeps track of the summary table row position
Dim DataStartRow As Integer             'identifies the first row of the stock data

Dim TickerOpen As Double                'the opening price for the date
Dim TickerHigh As Double                'the high price for the date
Dim TickerLow As Double                 'the low price for the date
Dim TickerClose As Double               'the closing price for the date
Dim YearlyChange As Double              'the dollar change in stock price from the first day of the the trading year to the last day of the trading year
Dim PercentChange As Double             'the percentage change in stock price from the first day of the the trading year to the last day of the trading year
Dim YearlyVolume As Double              'the volume of stock traded during the trading year
Dim GreatestIncreasePct As Double       'holds the greatest annual percentage increase
Dim GreatestDecreasePct As Double       'holds the greatest annual percentge decrease

Dim GreatestIncreaseTicker As String    'identifies the ticker with the greatest percentage increase
Dim GreatestDecreaseTicker As String    'identifies the ticker with the greatest percentage decrease
Dim GreatestVolumeTicker As String      'identifies the ticker with the greatest percentage increase

Dim GreatestVolume As LongLong          'holds the greatest annual volume
Dim LastRowOfData As LongLong           'used to find the last row of data in column A
Dim i As LongLong                       'a counter used to step through the stock ticker data
Dim j As LongLong                       'a counter used to step through the results table
Dim Volume As LongLong                  'the volume of the stocke traded on the date

Dim ws As Worksheet                     'used to step through each worksheet in the workbook

'********************************
'**Step through every worksheet**
'********************************


For Each ws In Worksheets
    
    'MsgBox ("worksheet name: " + ws.Name)   'I used this to make sure each sheet was being evaluated in turn.
    
    '***********************************
    '**variable initialization section**
    '**re-initialize with every sheet **
    '**to avoid unexpected results    **
    '***********************************
    
    Ticker = ""
    TickerDate = 0
    TickerOpen = 0#
    TickerHigh = 0#
    TickerLow = 0#
    TickerClose = 0#
    Volume = 0
    YearlyChange = 0#
    PercentChange = 0#
    YearlyVolume = 0
    
    DataStartRow = 2
    ResultsTableRow = 2
    
    i = 0
    j = 0
    
    
    '*************************
    '**create output section**
    '*************************
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 9).ColumnWidth = 10
    
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 10).ColumnWidth = 15
    
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 11).ColumnWidth = 15
    
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 12).ColumnWidth = 21
    
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 16).ColumnWidth = 10
    
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(1, 17).ColumnWidth = 15
    
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(2, 15).ColumnWidth = 25
    
        
    '*****************************************
    '**Loop through all stock ticker symbols**
    '*****************************************
    
    '**************************************************************
    '**Get the first stock's data - need to get the opening price**
    '**of the first stock.                                       **
    '**************************************************************
    
    Ticker = ws.Cells(DataStartRow, 1).Value
    TickerOpen = ws.Cells(DataStartRow, 3).Value
    
    '******************************************
    '**Find the last row of data in column A**'
    '******************************************
    
    LastRowOfData = ws.Range("A:A").SpecialCells(xlCellTypeLastCell).Row
    'MsgBox ("Last row of data in column A: " + Str(LastRowOfData))  I used this to confirm the last row of data in each worksheet.
    
    
    For i = DataStartRow To LastRowOfData
    
        '*******************************************************
        '**Check if we are still within the same stock ticker.**
        '**If the stock ticker is different, output totals.   **
        '*******************************************************
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            'set the stock ticker symbol
            Ticker = ws.Cells(i, 1).Value
             
            'obtain the closing price
            TickerClose = ws.Cells(i, 6).Value
              
            'obtain the volume
            Volume = Volume + ws.Cells(i, 7).Value
    
            'print the stock ticker symbol in the summary table
            ws.Cells(ResultsTableRow, 9).Value = Ticker
        
            'print the yearly change in the summary table
            ws.Cells(ResultsTableRow, 10).Value = TickerClose - TickerOpen
        
            'shade the cells depending on the annual change: red if negative, green if positive, none if no change
            
            If (TickerClose - TickerOpen) < 0 Then
                ws.Cells(ResultsTableRow, 10).Interior.ColorIndex = 3   'red
                ws.Cells(ResultsTableRow, 11).Interior.ColorIndex = 3   'red
            ElseIf (TickerClose - TickerOpen) > 0 Then
                ws.Cells(ResultsTableRow, 10).Interior.ColorIndex = 4    'green
                ws.Cells(ResultsTableRow, 11).Interior.ColorIndex = 4    'green
            End If
            
    
            'print the percent change in the summary table - this is to double-check results!
            ws.Cells(ResultsTableRow, 11).Value = Round(TickerClose / TickerOpen - 1, 4)
            ws.Cells(ResultsTableRow, 11).NumberFormat = "0.00%"
                
    
            'print the volume to the summary table
            ws.Cells(ResultsTableRow, 12).Value = Volume
            
            'ws.Cells(ResultsTableRow, 13).Value = TickerOpen       'use this to verify code is working as intended.
            'ws.Cells(ResultsTableRow, 14).Value = TickerClose      'use this to verify code is working as intended.
            
            'Add one to the summary table row
            ResultsTableRow = ResultsTableRow + 1
          
            'reset the volume, opening price, and closing price
            Volume = 0
            TickerOpen = ws.Cells(i + 1, 3).Value
            TickerClose = ws.Cells(i + 1, 6).Value
    
            '*************************************************************
            '**If the cell immediately following a row is the same stock**
            '**get the closing price and the volume.                    **
            '*************************************************************
        
        Else
    
            
            'add to stock's volume. don't need to change the opening price
            Volume = Volume + ws.Cells(i, 7).Value
            
        End If
    
      Next i
    
    
    '**************************************************************************
    '**Find the tickers with the greatest % increase and greatest % decrease.**
    '**                                                                      **
    '**get the start row and end row of the results table                    **
    '**we know the result row starts in row 2                                **
    '**we know the result end row from the loop above.                       **
    '**use loops to find max, min, and greatest volume                       **
    '**                                                                      **
    '**************************************************************************
    
        GreatestIncreaseTicker = ""
        GreatestDecreaseTicker = ""
        GreatestVolumeTicker = ""
        
        GreatestIncreasePct = 0#
        GreatestDecreasePct = 0#
        GreatestVolume = 0#
    
    For j = 2 To ResultsTableRow
    
        'check whether the percentage change of the next row's values is greater than or equal to the current value.
        
        If ws.Cells(j, 11).Value >= GreatestIncreasePct Then
            GreatestIncreasePct = ws.Cells(j, 11).Value
            GreatestIncreaseTicker = ws.Cells(j, 9).Value
        End If
        
        'check whether the percentage change of the next row's values is less than the current value.
        
        If ws.Cells(j, 11).Value < GreatestDecreasePct Then
            GreatestDecreasePct = ws.Cells(j, 11).Value
            GreatestDecreaseTicker = ws.Cells(j, 9).Value
        End If
        
        'check whether the volume of the next row's values is greaterh than or equal to the current value.
        
        If ws.Cells(j, 12).Value > GreatestVolume Then
            GreatestVolume = ws.Cells(j, 12).Value
            GreatestVolumeTicker = ws.Cells(j, 9).Value
        End If
    
    Next j
    
        'print the results
        ws.Cells(2, 16).Value = GreatestIncreaseTicker
        ws.Cells(2, 17).Value = GreatestIncreasePct
        ws.Cells(2, 17).NumberFormat = "0.00%"
         
        ws.Cells(3, 16).Value = GreatestDecreaseTicker
        ws.Cells(3, 17).Value = GreatestDecreasePct
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 16).Value = GreatestVolumeTicker
        ws.Cells(4, 17).Value = GreatestVolume
   
Next ws
End Sub

