Attribute VB_Name = "Module1"
Sub StockVolumes()

    'Declare variables
    Dim i As Double
    Dim LastRow As Double
    Dim StockTotal As Double
    Dim TableCount As Integer
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearChange As Double
    Dim PerChange As Double
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Volume As Double
    Dim Current_Max As Double
    Dim Current_Min As Double
    Dim Max_Ticker As String
    Dim Min_Ticker As String
    Dim Current_Greatest_Vol As Double
    Dim Greatest_Vol_Ticker As String
    
For Each ws In Worksheets

    'Set initial values
    TableCount = 2
    StockTotal = 0
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    OpenPrice = ws.Cells(2, 3).Value
    Current_Max = 0
    Current_Min = 0
    
    'Insert headers and labels
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("N2") = "Greatest % Increase"
    ws.Range("N3") = "Greatest % Decrease"
    ws.Range("N4") = "Greatest Total Volume"
    ws.Range("O1") = "Ticker"
    ws.Range("P1") = "Value"
    
    'Iterate through data by rows to find info for each ticker type
    For i = 2 To LastRow
            
        'Set Number Format to Percentage in Percent Change column
        ws.Cells(i, 11).NumberFormat = "0.00%"
            
        'Find each change in Ticker letter
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Insert value for Total Stock Volume
            ws.Cells(TableCount, 12).Value = StockTotal + ws.Cells(i, 7).Value
            StockTotal = 0
            
            'Insert value for Ticker letter
            ws.Cells(TableCount, 9).Value = ws.Cells(i, 1).Value
            
            'Find value for Closing Price at end of year
            ClosePrice = ws.Cells(i, 6).Value
                
            'Calculate the Yearly Change and set value in table
            YearChange = ClosePrice - OpenPrice
            ws.Cells(TableCount, 10).Value = YearChange
            
            'Set positive Yearly Change in green and negative Yearly Change in red
            If ws.Cells(TableCount, 10).Value > 0 Then
                ws.Cells(TableCount, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(TableCount, 10).Value < 0 Then
                ws.Cells(TableCount, 10).Interior.ColorIndex = 3
            End If
            
            'Calculate the Percent Change and set value in table
            If OpenPrice = 0 Then
                PerChange = 0
            Else
                PerChange = YearChange / OpenPrice
            End If
            
            ws.Cells(TableCount, 11).Value = PerChange
                
            'Set new opening price for next Ticker letter
            OpenPrice = ws.Cells(i + 1, 3).Value
            
            'Move to next row down in data output table
            TableCount = TableCount + 1
        
        Else
            
            'Increase Total Stock Volume
            StockTotal = StockTotal + ws.Cells(i, 7).Value
            
        End If
        
    Next i
    
    'Find greatest values
    LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    For i = 2 To LastRow2
    
        'Find greatest increase and decrease
        If ws.Cells(i, 11) > 0 And ws.Cells(i, 11) > Current_Max Then
            
            'Set new greatest increase with corresponding ticker type
            Current_Max = ws.Cells(i, 11)
            Max_Ticker = ws.Cells(i, 9)
        
        ElseIf ws.Cells(i, 11) < 0 And ws.Cells(i, 11) < Current_Min Then
            
            'Set new greatest decrease with corresponding ticker type
            Current_Min = ws.Cells(i, 11)
            Min_Ticker = ws.Cells(i, 9)
        
        End If
        
        'Find greatest total stock volume
        If ws.Cells(i, 12) > Current_Greatest_Vol Then
        
            'Set new greatest volume with corresponding ticker type
            Current_Greatest_Vol = ws.Cells(i, 12)
            Greatest_Vol_Ticker = ws.Cells(i, 9)
        
        End If
        
    Next i
    
    'Set greatest values into table
    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"
    ws.Range("O2") = Max_Ticker
    ws.Range("P2") = Current_Max
    ws.Range("O3") = Min_Ticker
    ws.Range("P3") = Current_Min
    ws.Range("O4") = Greatest_Vol_Ticker
    ws.Range("P4") = Current_Greatest_Vol
    
    Current_Greatest_Vol = 0
Next ws
End Sub
