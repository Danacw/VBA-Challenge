Attribute VB_Name = "Module1"
Sub Stocks()

'Make sure to loop through all sheets
For Each ws In Worksheets

'Declare last row
Dim LastRow As Long

'Determine Last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Set initial variable for Ticker Name
    Dim Ticker As String
    
    'Set Yearly_Change as double variable
    Dim Yearly_Change As Double
    Yearly_Change = 0
    
    'Set percent change as double variable
    Dim Percent_Change As Double
    Percent_Change = 0
    
    'Set Total Stock Volume as Integer
    Dim Total_Stock As Double
    Total_Stock = 0
    
    'Set total yearly as double
    Dim Total_yearly As Double
    
    'Keep track in summary table
    Dim Summary_table_row As Integer
    Summary_table_row = 2
    
    'set headers for each worksheet
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    Year_Open = ws.Cells(2, 3).Value
    
    'Loop through all Ticker rows
    For i = 2 To LastRow
        
        'Check if we are still within same Ticker name if not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Set the ticker name
            Ticker = ws.Cells(i, 1).Value
            
                'Define year close value
                Year_Close = ws.Cells(i, 6).Value
                
                    'Calculate Yearly_Change
                    Yearly_Change = Year_Close - Year_Open
                    
                            ' Solve for 0 issue by setting percent change to 0
                            If Year_Open = 0 Then
                            Percent_Change = 0
                            Else
                    
                    'Calculate Percent_Change
                    Percent_Change = (Yearly_Change / Year_Open)
            
                Stock_Volume = ws.Cells(i, 7).Value
                    
                    'Calculate Total Stock Volume
                    Total_Stock = Total_Stock + Stock_Volume
                    
                    End If
        
                'Print the ticker name in summary table
                ws.Range("I" & Summary_table_row).Value = Ticker
                
                'Print the yearly change in summary table
                 ws.Range("J" & Summary_table_row).Value = Yearly_Change
                
                'Print Percent Change in summary table
                ws.Range("K" & Summary_table_row).Value = Percent_Change
                ws.Range("K" & Summary_table_row).NumberFormat = "0.00%"
                       
                'Print Total Stock in summary table
                ws.Range("L" & Summary_table_row).Value = Total_Stock
    
        'Add one to the summary table row
            Summary_table_row = Summary_table_row + 1
            
    'Assign new value to opening price
    Year_Open = ws.Range("C" & i + 1).Value
            
    'Reset the total stock
    Total_Stock = 0
    
    'Reset Yearly_Change
    Yearly_Change = 0
    
    'Reset percent change
    Percent_Change = 0
        
'if the cell following a row is the same ticker...
Else
    
    'Add to the total yearly change, percent change and total stock
        
        'Calculate Yearly_Change
        Yearly_Change = Year_Close - Year_Open
        
        ' Solve for 0 issue by setting percent change to 0
        If Year_Open = 0 Then
        Percent_Change = 0
        Else
        'Calculate Percent_Change
        Percent_Change = (Yearly_Change / Year_Open)
        
        Stock_Volume = ws.Cells(i, 7).Value
        
        'Calculate Total Stock Volume
        Total_Stock = Total_Stock + Stock_Volume
        
        End If
        
        End If
        
    Next i
    
 'Add loop for cell formatting
 For i = 2 To LastRow
 
    'Add conditional for cell formatting
    If ws.Cells(i, 10).Value >= 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
    
    End If

Next i
    
'Define variable for finding max and min values
Dim PercentLastRow As Long
PercentLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
Dim Percent_max As Double
Percent_max = 0
Dim Percent_min As Double
Percent_min = 0

    'Add loop to find max and min
    For i = 2 To PercentLastRow

        'Add conditionals to calculate max
        If Percent_max < ws.Cells(i, 11) Then
            Percent_max = ws.Cells(i, 11).Value
            'print values
            ws.Cells(2, 17).Value = Percent_max
            ws.Cells(2, 17).NumberFormat = "0.00%"
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        
        'Add conditionals to calculate min
        ElseIf Percent_min > ws.Cells(i, 11) Then
            Percent_min = ws.Cells(i, 11).Value
            'print values
            ws.Cells(3, 17).Value = Percent_min
            ws.Cells(3, 17).NumberFormat = "0.00%"
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    
        End If
    
    Next i

''Define variable for greatest volume
Dim VolumeLastRow As Long
VolumeLastRow = ws.Cells(Rows.Count, 12).End(xlUp).Row
Dim Greatest_Volume As Double
Greatest_Volume = 0

    'Add loop to find Greatest_Volume
    For i = 2 To VolumeLastRow
    
        'Add conditionals to calculate volume
        If Greatest_Volume < ws.Cells(i, 12) Then
            Greatest_Volume = ws.Cells(i, 12).Value
            'print values
            ws.Cells(4, 17).Value = Greatest_Volume
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    
        End If
    
    Next i
    
Next ws
    
End Sub

