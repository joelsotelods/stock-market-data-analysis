Sub stock_market_data_analysis_greatests()

    Dim SheetNameStr As String
    Dim CodeNameStr As String

    Dim LastRow As Long

    Dim i As Long

    ticker_index = 2

    'Varialbes to store the yearly change from open to close
    Dim open_val As Double
    Dim close_val As Double

    For Each currentsheet In ThisWorkbook.Worksheets

        SheetNameStr = currentsheet.Name
        CodeNameStr = currentsheet.CodeName

        currentsheet.Range("I1").Value = "Ticker"
        currentsheet.Range("J1").Value = "Yearly Change"
        currentsheet.Range("K1").Value = "Percent Change"
        currentsheet.Range("L1").Value = "Total Stock Volume"

        LastRow = currentsheet.Cells(currentsheet.Rows.Count, 1).End(xlUp).Row
                                  
        For i = 2 To LastRow
            'currentsheet.Cells(i, 7).Value
            
            if IsEmpty(currentsheet.Cells(ticker_index, 9).Value)  then
                currentsheet.Cells(ticker_index, 9).Value = currentsheet.Cells(i, 1).Value
                currentsheet.Cells(ticker_index, 12).Value = currentsheet.Cells(i, 7).Value
                
                'xx
                open_val = currentsheet.Cells(i, 3).Value
                'MsgBox "Open " & open_val
            else
                if currentsheet.Cells(ticker_index, 9).Value = currentsheet.Cells(i, 1).Value then
                    currentsheet.Cells(ticker_index, 12).Value = currentsheet.Cells(ticker_index, 12).Value + currentsheet.Cells(i, 7).Value
                else
                    'save the close val of the year (last row of the Ticker)
                    close_val = currentsheet.Cells(i-1, 6).Value

                    'MsgBox "Close " & close_val & " Open " & open_val
                    currentsheet.Cells(ticker_index, 10).Value = close_val - open_val

                    'percent change
                    if open_val <> 0 then
                        currentsheet.Cells(ticker_index, 11).Value = (close_val/open_val) -1
                    else
                        currentsheet.Cells(ticker_index, 11).Value = "N/A"
                    end if
                    


                    ticker_index = ticker_index + 1
                    i = i -1 'this is needed because the <Next i> will ommit the first value of the new Ticker if not added.
                        ' I figure this out by making a Pivot Table of the original data and compairing that to the info obtained by this script.
                end if

            end if


        Next i
        
        'save the close val of the year (last row of the Ticker) -- unique for last Ticker
        close_val = currentsheet.Cells(i-1, 6).Value

        'MsgBox "Close " & close_val & " Open " & open_val
        currentsheet.Cells(ticker_index, 10).Value = close_val - open_val

        'percent change
        ' Compound Annual Growth Rate = (ending Balance / Beginning Balance)^(1/#years) - 1
        if open_val <> 0 then
            currentsheet.Cells(ticker_index, 11).Value = (close_val/open_val) -1
        else
            currentsheet.Cells(ticker_index, 11).Value = "N/A"
        end if


        MsgBox SheetNameStr & " was processed. Rows of the file: " & LastRow
        ticker_index = 2

    Next currentsheet

End Sub
