Sub Run_All_Worksheets()

Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        ws.Select
        Call StockData
    Next
    Application.ScreenUpdating = True
    
End Sub

Sub StockData()


'Create headers for Summary Table_Moderate
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percent Change"
    Range("M1").Value = "Total Stock Volume"
    
     
'Set and define initial variables
    Dim Ticker As String
    Dim TotalVol As Double
        TotalVol = 0
    Dim open_price As Double
        open_price = Cells(2, 3)
    Dim close_price As Double
    Dim YearlyChange As Double
    
     
'Set the location of the Ticker in the Summary Table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

'Set and find the LastRow
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Loop through all stocks in the worksheet
    For i = 2 To LastRow

    'Use If statement to compare cells in the row
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
               
        'Set the Ticker in the Summary Table
        'Ticker Name
        Ticker = Cells(i, 1).Value
               
        'Add Stock Volume in the Summary Table
        'Total Stock Volume
        TotalVol = TotalVol + Cells(i, 7).Value
             
        'Set value for close_price and calculate Yearly change
        
        close_price = Cells(i, 6).Value
            
        YearlyChange = (close_price - open_price)
        
        'Set variables and calculate Percent Change
        Dim PercentChange As Double
        
        If (open_price = 0) Then
            PercentChange = 0
            
        Else
            PercentChange = YearlyChange / open_price
        
        End If
             
        'Put data of each Ticker and Volume in the Summary Table
        Range("J" & Summary_Table_Row).Value = Ticker
        Range("M" & Summary_Table_Row).Value = TotalVol
        Range("K" & Summary_Table_Row).Value = YearlyChange
        Range("L" & Summary_Table_Row).Value = PercentChange
        
        'Change Format of Percent Change Column to 0.00%
        Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
        
        
        'Add one to the Summary Table Row
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Resent the TotalVol and open_price
        TotalVol = 0
        open_price = Cells(i + 1, 3)
       
                       
    'If Statement for TotalVol if the cell is the same Ticker in previous row
        Else
        TotalVol = TotalVol + Cells(i, 7).Value
        
       
        End If
        
    Next i
    

'Format and Conditional Formatting of Summary Table

Range("J:M").Columns.AutoFit

Dim LastRow_Summary_Table As Long
LastRow_Summary_Table = Cells(Rows.Count, 10).End(xlUp).Row


For i = 2 To LastRow_Summary_Table
    If Cells(i, 11).Value > 0 Then
        Cells(i, 11).Interior.ColorIndex = 4
    Else
        Cells(i, 11).Interior.ColorIndex = 3
    End If

Next i
    

End Sub


