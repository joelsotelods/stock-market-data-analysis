Sub stock_market_data_analysis_greatests()

    Dim SheetNameStr As String
    Dim CodeNameStr As String

    Dim LastRow As Long

    Dim i As Long

    Dim index_for_metrics As Long
    Dim metric_x As Double

    metric_x = 0

    For Each currentsheet In ThisWorkbook.Worksheets

        SheetNameStr = currentsheet.Name
        CodeNameStr = currentsheet.CodeName

        currentsheet.Range("P1").Value = "Ticker"
        currentsheet.Range("Q1").Value = "Value"
    

        LastRow = currentsheet.Cells(currentsheet.Rows.Count, 9).End(xlUp).Row


        ' Greatest % Increase   
                 
        For i = 2 To LastRow
            'currentsheet.Cells(i, 7).Value
            if currentsheet.Cells(i, 11).Value <> "N/A" then 
                if currentsheet.Cells(i, 11).Value > metric_x then
                    index_for_metrics = i
                    
                    metric_x = currentsheet.Cells(i, 11).Value
                End if
            End if

        Next i

        currentsheet.Cells(2, 15).Value = "Greatest % Increase: "
        currentsheet.Cells(2, 16).Value = currentsheet.Cells(index_for_metrics, 9).Value

        currentsheet.Cells(2, 17).NumberFormat="0.00%"
        currentsheet.Cells(2, 17).Value = currentsheet.Cells(index_for_metrics, 11).Value


        ' Greatest % Decrease:
                     
        For i = 2 To LastRow
            'currentsheet.Cells(i, 7).Value
            if currentsheet.Cells(i, 11).Value <> "N/A" then 
                if currentsheet.Cells(i, 11).Value < metric_x then
                    index_for_metrics = i
                    
                    metric_x = currentsheet.Cells(i, 11).Value
                End if
            End if

        Next i

        currentsheet.Cells(3, 15).Value = "Greatest % Decrease: "
        currentsheet.Cells(3, 16).Value = currentsheet.Cells(index_for_metrics, 9).Value

        currentsheet.Cells(3, 17).NumberFormat="0.00%"
        currentsheet.Cells(3, 17).Value = currentsheet.Cells(index_for_metrics, 11).Value


        ' Greatest Total Volume:
        metric_x = 0

        For i = 2 To LastRow
            'currentsheet.Cells(i, 7).Value
            if currentsheet.Cells(i, 12).Value > metric_x then
                index_for_metrics = i
                metric_x = currentsheet.Cells(i, 12).Value

            End if
        Next i

        currentsheet.Cells(4, 15).Value = "Greatest Total Volume: "
        currentsheet.Cells(4, 16).Value = currentsheet.Cells(index_for_metrics, 9).Value

        currentsheet.Cells(4, 17).NumberFormat="0"
        currentsheet.Cells(4, 17).Value = currentsheet.Cells(index_for_metrics, 12).Value

        currentsheet.Columns("O").EntireColumn.AutoFit
        currentsheet.Columns("Q").EntireColumn.AutoFit

        MsgBox SheetNameStr & " was processed. Rows of the file: " & LastRow
        metric_x = 0

    Next currentsheet

End Sub
