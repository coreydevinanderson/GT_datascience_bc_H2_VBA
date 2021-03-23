Attribute VB_Name = "Module11"

Sub Stocks():

'Set Worksheet object
Dim ws As Worksheet

'Iterate through each worksheet

For Each ws In ThisWorkbook.Worksheets

    ws.Activate 'Must activate the worksheet

    'Determine last row of data in sheet
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    '----------------
    'Get ticker names
    '----------------

    'Create blank array
    Dim ticker_names() As String

    'Set counter for first loop
    Dim Count As Integer
    Count = 0

    'Loop from row 2 to last row. Identify ticker names when the row names change.
    'Add procedure for handling the last row

    For i = 2 To lastrow
        If (i = lastrow) Then
            ReDim Preserve ticker_names(Count)
            ticker_names(Count) = Cells(i, 1).Value
        ElseIf (Cells(i, 1).Value <> Cells(i + 1, 1).Value) Then
            ReDim Preserve ticker_names(Count)
            ticker_names(Count) = Cells(i, 1).Value
            Count = Count + 1
        End If
    Next i
    
    'Determine how many ticker names were found
    ticker_length = UBound(ticker_names) - LBound(ticker_names) + 1

    'Add column labels for output
    Range("I1").Value = "Ticker"
    
    Range("J1").Value = "Yearly Change"
    Columns("J").ColumnWidth = 13
    
    Range("K1").Value = "Percent Change"
    Columns("K").ColumnWidth = 14
    
    Range("L1").Value = "Total Stock Volume"
    Columns("L").ColumnWidth = 18

    'Write ticker names to column I
    'Start by setting counter for indexing to extract each element from ticker_names array
    ticker_count = 0

    For j = 2 To (ticker_length + 1)
        Cells(j, 9).Value = ticker_names(ticker_count)
        ticker_count = ticker_count + 1
    Next j
    
    
    '--------------------------------------------------------
    'Calculate Yearly change, Percent change, and Total volue
    'Color Yearly change using conditional formatting
    '--------------------------------------------------------


    'Start by defing the object classes to be used in the calculations
    Dim opening As Double
    Dim closing As Double
    
    Dim volume As LongLong
    volume = 0 'Initialize volume to zero
    
    'Create group count variable to keep track of number of rows for each unique stock
    Dim group_count As Integer
    group_count = 0

    'Create counter to  position output into sheet, starting in row 2
    Dim row_out As Integer
    row_out = 2

    ' Set and initialize variables for storing max and min percent, max volume and the ticker ID that corresponds to each value
    Dim max_percent_ID As String
    Dim min_percent_ID As String
    Dim max_volume_ID As String
    Dim max_percent As Double
    Dim min_percent As Double
    Dim max_volume As LongLong

    min_percent = 0
    max_percent = 0
    max_volume = 0

    '----------

    'Add row and column names for output
    
    'row names
    Columns("O").ColumnWidth = 21 'Set column width based on longest character string for row name
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"

    'column names
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"


    '----------


    'For loop to do calculations and formating (including the bonus questions)
    
    'If the ticker names are the same for row k and k + 1, add one to group_count
    'If the ticker names are different, the calculation is triggered, using group_count to back track to the first...
    '...row for that ticker name.
    
    'Start with separate procedure for last row (first), since for this row there is no values in row k + 1
    

    For k = 2 To lastrow
    
        If (k = lastrow) Then
            closing = Cells(k, 6).Value
            opening = Cells(k - group_count, 3).Value
            Yearly_change = closing - opening
            Cells(row_out, 10).Value = Yearly_change
            
            'Conditional formatting
            If (Yearly_change > 0) Then
                Cells(row_out, 10).Interior.ColorIndex = 4
            ElseIf (Yearly_change < 0) Then
                Cells(row_out, 10).Interior.ColorIndex = 3
            ElseIf (Yearly_change = 0) Then
                Cells(row_out, 10).Interior.ColorIndex = 6
            End If
            
            ' Make sure to account for division by zero when calculating percent change
            If (opening <> 0) Then
                Percent_Change = (closing - opening) / opening
            ElseIf (opening = 0) Then
                Percent_Change = (closing - opening) / 1
            End If
            
            Percent_Change = (closing - opening) / opening
            Cells(row_out, 11).Value = Percent_Change
            Cells(row_out, 11).NumberFormat = "0.00%"
            
            volume = volume + Cells(k, 7).Value
            Cells(row_out, 12).Value = volume
        
            'For each calcuation, compare to previous max and min up to that point.
            'Replace if it is higher or lower than previous "winner."
            If (Percent_Change > max_percent) Then
                max_percent = Percent_Change
                max_percent_ID = Cells(k, 1).Value
            End If
        
            If (Percent_Change < min_percent) Then
                min_percent = Percent_Change
                min_percent_ID = Cells(k, 1).Value
            End If
        
            If (volume > max_volume) Then
                max_volume = volume
                max_volume_ID = Cells(k, 1).Value
            End If
        
        ElseIf (Cells(k, 1).Value = Cells(k + 1, 1).Value) Then
            group_count = group_count + 1
            volume = volume + Cells(k, 7).Value
        
        ElseIf (Cells(k, 1).Value <> Cells(k + 1, 1).Value) Then
            closing = Cells(k, 6).Value
            opening = Cells(k - group_count, 3).Value
            Yearly_change = closing - opening
            Cells(row_out, 10).Value = Yearly_change
        
            'Conditional formatting
            If (Yearly_change > 0) Then
                Cells(row_out, 10).Interior.ColorIndex = 4
            ElseIf (Yearly_change < 0) Then
                Cells(row_out, 10).Interior.ColorIndex = 3
            ElseIf (Yearly_change = 0) Then
                Cells(row_out, 10).Interior.ColorIndex = 6
            End If
            
            
            'Make sure to account for division by zero when calculating percent change
            If (opening <> 0) Then
                Percent_Change = (closing - opening) / opening
            ElseIf (opening = 0) Then
                Percent_Change = (closing - opening) / 1
            End If
            
            Cells(row_out, 11).Value = Percent_Change
            Cells(row_out, 11).NumberFormat = "0.00%"
            
            volume = volume + Cells(k, 7).Value
            Cells(row_out, 12).Value = volume
        
            'For each calcuation, compare to previous max and min up to that point.
            'Replace if it is higher or lower than previous "winner."
            If (Percent_Change > max_percent) Then
                max_percent = Percent_Change
                max_percent_ID = Cells(k, 1).Value
            End If
        
            If (Percent_Change < min_percent) Then
                min_percent = Percent_Change
                min_percent_ID = Cells(k, 1).Value
            End If
        
            If (volume > max_volume) Then
                max_volume = volume
                max_volume_ID = Cells(k, 1).Value
            End If
        
            row_out = row_out + 1
            
            'reset the counters for group and volume after doing calculation for each ticker name
            group_count = 0
            volume = 0
    
        End If
    Next k

    'Write summary stats to worksheet
    Cells(2, 16).Value = max_percent_ID
    Cells(2, 17).Value = max_percent
    Cells(2, 17).NumberFormat = "0.00%"

    Cells(3, 16).Value = min_percent_ID
    Cells(3, 17).Value = min_percent
    Cells(3, 17).NumberFormat = "0.00%"

    Cells(4, 16).Value = max_volume_ID
    Cells(4, 17).Value = max_volume

Next ws 'Move on to the next worksheet.

End Sub


