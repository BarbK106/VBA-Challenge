VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Homework()
For Each ws In Worksheets

        'Create a Variable to Hold File Name and Last Row
        Dim WorksheetName As String
        Dim Ticker_Name As String
        Dim Summary_Table_Row As Long
        Dim Opening As Double
        Dim Change As Double
        Dim Percent_Change As Double
        Dim Max_Ticker As String
        Dim Max_Increase As Double
        Dim Min_Ticker As String
        Dim Min_Decrease As Double
        Dim Max_Volume As String
        Dim Max_Total As Double
        

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Grabbed the WorksheetName
        WorksheetName = ws.Name

        ' Add Column Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
   
        ' Set an initial variable for holding ticker symbol
        
        Summary_Table_Row = 2
        Opening = ws.Cells(2, 3).Value
        
        'Loop through all tickers
        For i = 2 To LastRow
    
            ' Check if we are still within the same ticker symbol, if we are not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker_Name = ws.Cells(i, 1).Value
                
                ' Add to the Ticker Volume Total
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
                'Calculate Yearly Change
                Change = ws.Cells(i, 6).Value - Opening
                Percent_Change = Change / Opening
                Opening = ws.Cells(i + 1, 3).Value
        
            ' Print Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
            ws.Range("J" & Summary_Table_Row).Value = Change
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            ws.Range("L" & Summary_Table_Row).Value = Ticker_Volume
    
       
       If Percent_Change > Max_Increase Then
       Max_Increase = Percent_Change
       Max_Ticker = Ticker_Name
       
       End If
       
        If Percent_Change < Min_Decrease Then
       Min_Decrease = Percent_Change
       Min_Ticker = Ticker_Name
       
       End If
       
       If Ticker_Volume > Max_Total Then
        Max_Total = Ticker_Volume
        Max_Volume = Ticker_Name
        
        
       End If
       
        'Update Variables
       Ticker_Volume = 0
       Summary_Table_Row = Summary_Table_Row + 1
       
    Else
    
        Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
     
    End If
    Next i
    
    
        'Print 2nd Table
            ws.Cells(1, 15).Value = "Ticker"
            ws.Cells(1, 16).Value = "Value"
            ws.Cells(2, 14).Value = "Greatest % Increase"
            ws.Cells(3, 14).Value = "Greatest % Decrease"
            ws.Cells(4, 14).Value = "Greatest Total Volume"
            ws.Cells(2, 15).Value = Max_Ticker
            ws.Cells(3, 15).Value = Min_Ticker
            ws.Cells(4, 15).Value = Max_Volume
            ws.Cells(2, 16).Value = Max_Increase
            ws.Cells(3, 16).Value = Min_Decrease
            ws.Cells(4, 16).Value = Max_Total
    
    Next
    
    End Sub


