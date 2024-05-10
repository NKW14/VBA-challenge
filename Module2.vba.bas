Attribute VB_Name = "Module1"
Sub Stocks():

    For Each ws In Worksheets
        Dim WorksheetName As String
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Dim i As Long
        
        
        
        WorksheetName = ws.Name
        
        ws.Cells(1, 9).Value = "Ticker"
        
        ws.Cells(1, 10).Value = "Quarterly Change"
        
        ws.Cells(1, 11).Value = "Percent Change"
        
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(1, 16).Value = "Ticker"
        
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest%Increase"
        
        ws.Cells(3, 15).Value = "Greatest%Decrease"
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        Dim Stock_Name As String
        Dim Stock_Value As Double
        Dim Stock_Value_1 As Double
        Dim Total_Stock_Volume As Double
        Dim Summary_1 As Integer
        Dim Stock_Range As Range
        Dim Summary_2 As Integer
        Dim lastrow_p As Long
        Dim lastrow_d As Long
        Dim lastrow_v As Long
        
        

        Stock_Value = 0
        Stock_Value_1 = 0
        Total_Stock_Volume = 0
        Summary_1 = 2
        
        For i = 2 To LastRow
            If ws.Cells(i, 2).Value = "1/2/2022" Or ws.Cells(i, 2).Value = "4/1/2022" Or ws.Cells(i, 2).Value = "7/1/2022" Or ws.Cells(i, 2).Value = "10/1/2022" Then
            ws.Range("W" & Summary_1) = ws.Cells(i, 3).Value
            ws.Range("V" & Summary_1) = ws.Cells(i, 1).Value
            Summary_1 = Summary_1 + 1
            
        Else
        End If
        Next i
        Summary_2 = 2
        
        Set Stock_Range = ws.Range("V:W")
                
        For i = 2 To LastRow
        
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    Stock_Name = ws.Cells(i, 1).Value
                    
                    Stock_Value_1 = Application.WorksheetFunction.VLookup(Stock_Name, Stock_Range, 2, False)
                    
                    Stock_Value = ws.Cells(i, 6).Value
                    
                    Stock_Quarterly_Change = Stock_Value - Stock_Value_1
                    
                    Stock_Percent_Change = (Stock_Quarterly_Change / Stock_Value_1)
                    
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                    
                    ws.Range("I" & Summary_2).Value = Stock_Name
                    
                    ws.Range("J" & Summary_2).Value = Stock_Quarterly_Change
                    
                    ws.Range("K" & Summary_2).Value = Stock_Percent_Change
                    
                    ws.Range("L" & Summary_2).Value = Total_Stock_Volume
                    
                    
                    Summary_2 = Summary_2 + 1
                    
                    Stock_Value = 0
                    Total_Stock_Volume = 0
                    Start_Value = 0
                    
                    Else
                    Stock_Value = Stock_Value + (ws.Cells(i, 3).Value - ws.Cells(i, 6).Value)
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                    End If
                
            Next i
    
        
        For i = 2 To LastRow
        
        ws.Range("K" & i).NumberFormat = "0.00%"
        
        Next i
        
        
            
        For i = 2 To LastRow
        If ws.Cells(i, 10).Value > 0 Then
        
        ws.Cells(i, 10).Interior.ColorIndex = 4
        
        ElseIf ws.Cells(i, 10).Value < 0 Then
        
        ws.Cells(i, 10).Interior.ColorIndex = 3
        
        Else
        
        End If
        Next i
        
        lastrow_p = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For i = 2 To lastrow_p
        If ws.Cells(i, 11).Value > ws.Cells(i + 1, 11).Value And ws.Cells(i, 11).Value > ws.Cells(2, 17).Value Then
        ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        
        
        
        
        Else
        
        End If
        
        Next i
        
        lastrow_d = ws.Cells(Rows.Count, 9).End(xlUp).Row - 1
        
        For i = 2 To lastrow_d
        If ws.Cells(i, 11).Value < ws.Cells(i + 1, 11).Value And ws.Cells(i, 11).Value < ws.Cells(3, 17).Value Then
        ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        
        Else
        End If
        
        Next i
        
        Volume_range = ws.Range("L2:L1501").Value
        max_Volume = WorksheetFunction.Max(Volume_range)
        
        ws.Cells(4, 17).Value = max_Volume
        
        For i = 2 To lastrow_p
        If ws.Cells(i, 12).Value = ws.Cells(4, 17).Value Then
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        
        Else
        End If
        
        Next i
        
        
        For i = 2 To 3
        
        ws.Range("Q" & i).NumberFormat = "0.00%"
        
        Next i
        
    ws.Range("W2:W1501").ClearContents
    ws.Range("V2:V1501").ClearContents
    
    Next ws
        
End Sub
