Attribute VB_Name = "Module1"

Sub VBAStocks()
    Dim ws As Worksheet
    
    For Each ws In Sheets
        Call analyzeStockMarket(ws)
    Next ws
End Sub

Sub analyzeStockMarket(ws As Worksheet)
'Sub analyzeStockMarket()
    ' activate the passed worksheet
    ws.Activate

   ' general variables
    Dim data_idx As Long
    Dim display_idx As Integer
    Dim cum_volume As LongLong
    Dim error_limit As LongLong
    
    ' variables for the current ticker
    Dim previous_ticker As String
    Dim previous_open As Double
    Dim previous_close As Double
    Dim yearly_change As Double
    Dim yearly_percent As Double
    
    ' variables for the current row
    Dim t_ticker As String
    Dim t_open As Double
    Dim t_volume As LongLong
             
    ' variables for greatest analysis
    Dim greatest_percent_increase As Double
    Dim greatest_percent_decrease As Double
    Dim greatest_total_volume As LongLong
    Dim greatest_percent_increase_idx As Integer
    Dim greatest_percent_decrease_idx As Integer
    Dim greatest_total_volume_idx As Integer
        
    ' initialize
    display_idx = 1
    error_limit = Rows.Count

    Call display_headers
    
    ' handle the starting boundary
    previous_ticker = Cells(2, 1)
    previous_open = CDbl(Cells(2, 3))
    display_idx = display_idx + 1
    
    For data_idx = 3 To (WorksheetFunction.CountA(Columns(1)) + 1)
            ' read row
            t_ticker = Cells(data_idx, 1)
            t_open = CDbl(Cells(data_idx, 3))
            t_volume = CLngLng(Cells(data_idx, 7))
            previous_ticker = Cells(data_idx - 1, 1)
            previous_close = CDbl(Cells(data_idx - 1, 6))
            
            If Cells(data_idx, 1) = previous_ticker Then
                    ' not a new ticker
                    cum_volume = cum_volume + t_volume
            Else
                    ' new ticker
                    
                   yearly_change = previous_close - previous_open
                    If previous_open <> 0 Then
                            yearly_percent = yearly_change / previous_open
                    Else
                            yearly_percent = 0
                    End If
                    Cells(display_idx, 9) = previous_ticker
                    Cells(display_idx, 10).Interior.ColorIndex = 4     ' green
                    If yearly_change < 0 Then
                        Cells(display_idx, 10).Interior.ColorIndex = 3     ' red
                    End If
                    Cells(display_idx, 10) = yearly_change
                    Cells(display_idx, 11) = yearly_percent
                    Cells(display_idx, 12) = cum_volume
                    previous_open = t_open
                    cum_volume = t_volume
                    display_idx = display_idx + 1
            End If
                
            If data_idx >= error_limit Then
                    '  data_idx erroneously reaches Excel limit
                    MsgBox ("Execution limit exceeded")
                    Exit For
            End If
    Next data_idx
    
    '
    '  Find greatest stocks
    '
    greatest_percent_increase = CDbl(0)
    greatest_percent_decrease = CDbl(1)
    greatest_total_volume = CLngLng(0)
    
    For display_idx = 2 To WorksheetFunction.CountA(Columns(9))
        yearly_percent = CDbl(Cells(display_idx, 11))
        cum_volume = CLngLng(Cells(display_idx, 12))
        If yearly_percent > greatest_percent_increase Then
            greatest_percent_increase = yearly_percent
            greatest_percent_increase_idx = display_idx
        End If
        If yearly_percent < greatest_percent_decrease Then
            greatest_percent_decrease = yearly_percent
            greatest_percent_decrease_idx = display_idx
        End If
        If cum_volume > greatest_total_volume Then
            greatest_total_volume = cum_volume
            greatest_total_volume_idx = display_idx
        End If
    Next display_idx

    ' display greatest
    Cells(2, 16) = Cells(greatest_percent_increase_idx, 9)
    Cells(2, 17) = Cells(greatest_percent_increase_idx, 11)
    Cells(3, 16) = Cells(greatest_percent_decrease_idx, 9)
    Cells(3, 17) = Cells(greatest_percent_decrease_idx, 11)
    Cells(4, 16) = Cells(greatest_total_volume_idx, 9)
    Cells(4, 17) = Cells(greatest_total_volume_idx, 12)
    
    ' format percentage columns
    Range("K:K").NumberFormat = "0.00%"
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).NumberFormat = "0.00%"
        
    ' AutoFit All Columns on Worksheet
    Cells.EntireColumn.AutoFit
    
End Sub

Sub display_headers()
    ' Place result headers
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
    Cells(2, 15) = "Greatest % Increase"
    Cells(3, 15) = "Greatest % Decrease"
    Cells(4, 15) = "Greatest Total Volume"
    Cells(1, 16) = "Ticker"
    Cells(1, 17) = "Value"
End Sub

