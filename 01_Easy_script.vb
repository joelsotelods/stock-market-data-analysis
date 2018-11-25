Sub stock_market_data_analysis()

    Dim SheetNameStr As String
    Dim CodeNameStr As String

    Dim LastRow As Long
    Dim LastRow_tmp As Long

    Dim i As Long

    ticker_index = 2

    For Each currentsheet In ThisWorkbook.Worksheets

        SheetNameStr = currentsheet.Name
        CodeNameStr = currentsheet.CodeName

        ' To clean the cells that we are going to use in the case this script runs twice
        LastRow_tmp = currentsheet.Cells(currentsheet.Rows.Count, 9).End(xlUp).Row
        currentsheet.Range("H1:J" & LastRow_tmp).Clear

        currentsheet.Range("I1").Value = "Ticker"
        currentsheet.Range("J1").Value = "Total Stock Volume"

        LastRow = currentsheet.Cells(currentsheet.Rows.Count, 1).End(xlUp).Row
                                  
        For i = 2 To LastRow
            
            if IsEmpty(currentsheet.Cells(ticker_index, 9).Value)  then
                currentsheet.Cells(ticker_index, 9).Value = currentsheet.Cells(i, 1).Value
                currentsheet.Cells(ticker_index, 10).Value = currentsheet.Cells(i, 7).Value
            else
                if currentsheet.Cells(ticker_index, 9).Value = currentsheet.Cells(i, 1).Value then
                    currentsheet.Cells(ticker_index, 10).Value = currentsheet.Cells(ticker_index, 10).Value + currentsheet.Cells(i, 7).Value
                else
                    ticker_index = ticker_index + 1
                    i = i -1 'this is needed because the <Next i> will ommit the first value of the new Ticker if not added.
                        ' I figure this out by making a Pivot Table of the original data and compairing that to the info obtained by this script.
                end if

            end if

        Next i

        'Autofit the column size
        currentsheet.Columns("J").EntireColumn.AutoFit

        'MsgBox SheetNameStr & " was processed. Rows on the sheet: " & LastRow
        ticker_index = 2

    Next currentsheet

    MsgBox "The last Sheet (" & SheetNameStr & ") was processed."

End Sub
