Attribute VB_Name = "Module1"
Sub Stocks()
    ' Set WS as a worksheet
        Dim WS As Worksheet
        
    ' Loop through all sheets
        For Each WS In Worksheets
    
        ' Set the variables for the ticker and requirements
            Dim Ticker_Name As String
            Ticker_Name = " "
            Dim ticker_volume As Double
            ticker_volume = 0
            Dim Open_Price As Double
            Open_Price = 0
            Dim Close_Price As Double
            Close_Price = 0
            Dim Yearly_Change As Double
            Yearly_Change = 0
            Dim Percent_Change As Double
            Percent_Change = 0
         
        ' Record the ticker location for each sheet and keep it in the summary table
            Dim Summary_Table As Long
            Summary_Table = 2
        
        ' Set the row count for the current worksheet and determine the last row.
            Dim Lastrow As Long
            Dim i As Long
            Lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        ' Set the titles for the requirments
            WS.Range("I1").Value = "Ticker"
            WS.Range("J1").Value = "Yearly Change"
            WS.Range("K1").Value = "Percent Change"
            WS.Range("L1").Value = "Total Stock Volume"
            
        ' Set initial open price from column C
        Open_Price = WS.Cells(2, 3).Value
        
        ' Loop through the ticker symbol
        For x = 2 To Lastrow
        
      
            ' Check if we are still within the same ticker name,if not save them to the summary table
            If WS.Cells(x + 1, 1).Value <> WS.Cells(x, 1).Value Then
            
                ' Set the ticker name from column A
                Ticker_Name = WS.Cells(x, 1).Value
                'Print ticket name
                WS.Range("I" & Summary_Table).Value = Ticker_Name
                
                ' Define close price from column F
                Close_Price = WS.Cells(x, 6).Value
                
                ' Calculate yearly Change
                Yearly_Change = Close_Price - Open_Price
                 'Print yearly_change
                WS.Range("J" & Summary_Table).Value = Yearly_Change
                
                ' Calculate percent change
                If Open_Price <> 0 Then
                    Percent_Change = (Yearly_Change / Open_Price) * 100
                End If
                 ' Print percent change
                WS.Range("K" & Summary_Table).Value = (CStr(Percent_Change) & "%")
               
                ' Add to the ticker name total volume from column G
                ticker_volume = ticker_volume + WS.Cells(x, 7).Value
                                            
                ' Fill Yearly change with colors depending on the value
                If (Yearly_Change > 0) Then
                    'Fill column with green color
                    WS.Range("J" & Summary_Table).Interior.ColorIndex = 4
                ElseIf (Yearly_Change <= 0) Then
                    'Fill column with red color
                    WS.Range("J" & Summary_Table).Interior.ColorIndex = 3
                End If
                
                
                ' Print the total stock volume
                WS.Range("L" & Summary_Table).Value = ticker_volume
                
               ' Add 1 to the summary table count
                Summary_Table = Summary_Table + 1
                ' Reset yearly_change and percent_change holders since we will be working with new ticker
                Yearly_Change = 0
                Close_Price = 0
                Percent_Change = 0
                ticker_volume = 0
                ' Capture next Ticker's Open_Price
                Open_Price = WS.Cells(x + 1, 3).Value
                 
            Else
                ' if cells are the same ticker, increase the total ticker volume
                ticker_volume = ticker_volume + WS.Cells(x, 7).Value
            End If
            
                  
        Next x

        Next WS
End Sub

