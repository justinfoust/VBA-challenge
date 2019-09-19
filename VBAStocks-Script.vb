Sub stockFormat()

'----------  Declarations  ----------

    Dim rowNum As Single
    Dim resulteRow As Single
    Dim lastRow As Single
    Dim lastResultRow As Single
    
    Dim iStockValue As Single
    Dim fStockValue As Single

    Dim iVolumeRow As Single
    Dim fVolumeRow As Single
    Dim sumVolume As Single

    Dim wsCount As Integer
    Dim wsNum As Integer
    
    Dim resultsArr(1 To 3, 1 To 2) As Variant
    

'----------  Code  ----------

   ' Total number of worksheets
   wsCount = ActiveWorkbook.Worksheets.Count
    
'--- FOR LOOP - Cycle through each worksheet ----------
    For wsNum = 1 To wsCount
        
        ' Start recording results from row 2
        resultRow = 2
           
        With ActiveWorkbook.Worksheets(wsNum)
        
            ' Write column and row headers for results
            .Range("I1").Value = "Ticker"
            .Range("J1").Value = "Yearly Change"
            .Range("K1").Value = "Percent Change"
            .Range("L1").Value = "Total Stock Volume"
        
            .Range("P1").Value = "Ticker"
            .Range("Q1").Value = "Value"
        
            .Range("O2").Value = "Greatest % Increase"
            .Range("O3").Value = "Greatest % Decrease"
            .Range("O4").Value = "Greatest Total Volume"
        
            ' Record row number of final row with data
            lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
            
    '---------- FOR LOOP - Read through each row with data in worksheet ----------
            For rowNum = 1 To lastRow
            
        '---------- IF - At Each Ticker Difference ----------
            
                 If .Cells(rowNum, "A").Value <> .Cells(rowNum + 1, "A").Value Then
                    
            '---------- IF - Skip first ticker difference before recording results ----------
                    
                    If rowNum <> 1 Then
                        
                        ' Calculate and write Yearly Change
                        fStockValue = .Cells(rowNum, "F").Value
                        .Cells(resultRow - 1, "J").Value = Format(fStockValue - iStockValue, "#,###.00")
                        
                        ' Calculate and write Total Stock Volume
                        fVolumeRow = rowNum
                        .Cells(resultRow - 1, "L").Value _
                        = .Application.Sum(Range(Cells(iVolumeRow, "G"), Cells(fVolumeRow, "G")))
                        

                 '---------- IF - Only Calculate if Denomonator is not 0 to avoid undefined error ----------
                        
                        If iStockValue <> 0 Then
                            
                            ' Calculate and write Percent Change of stock value
                            .Cells(resultRow - 1, "K").Value = Format(fStockValue / iStockValue - 1, "Percent")
                            
                            ' IF - Color cell green for positive change
                            If .Cells(resultRow - 1, "J").Value > 0 Then
                            
                                .Cells(resultRow - 1, "J").Interior.ColorIndex = 4
                                
                            ' IF - Color cell red for negative change
                            Else
                                .Cells(resultRow - 1, "J").Interior.ColorIndex = 3
                            End If
                        
                        End If
                        
                 '---------- END IF - Don't Calculate if Denomonator is 0 ---------
                        
                    End If
                    
            '---------- END IF - Skip First ----------
            '---------- IF - Skip Last Blank Row ----------
                    
                    If VarType(.Cells(rowNum + 1, "A").Value) <> 0 Then
                        
                        ' Write ticker value in results column
                        .Cells(resultRow, "I").Value = .Cells(rowNum + 1, "A").Value
                        
                        ' Record initial stock value for next stock
                        iStockValue = .Cells(rowNum + 1, "C").Value
                        
                        ' Record initial row number to be used for summing stock volume for next stock
                        iVolumeRow = rowNum + 1
                        
                        ' increment one row for next set of results
                        resultRow = resultRow + 1
                    
                    End If
                    
             '---------- END IF - Skip Last Blank Row ----------
                    
                 End If
            
        '---------- END IF - At Each Ticker Difference ---------
            
            Next rowNum
            
            lastResultRow = .Cells(.Rows.Count, "I").End(xlUp).Row
            
            ' Record initial values in array to compare to
            resultsArr(1, 2) = 0
            resultsArr(2, 2) = 0
            resultsArr(3, 2) = 0
            
    '---------- FOR LOOP - Increment through results column ----------
            
            For resultRow = 2 To lastResultRow
            
           '---------- IF STATEMENTS - If values are greater/less than value saved in array,
           'write over them with current value and record corresponding ticker symbol ----------
           
                If .Cells(resultRow, "K").Value > resultsArr(1, 2) Then
                    resultsArr(1, 2) = .Cells(resultRow, "K").Value
                    resultsArr(1, 1) = .Cells(resultRow, "I").Value
                End If
                
                If .Cells(resultRow, "K").Value < resultsArr(2, 2) Then
                    resultsArr(2, 2) = .Cells(resultRow, "K").Value
                    resultsArr(2, 1) = .Cells(resultRow, "I").Value
                End If
                
                If .Cells(resultRow, "L").Value > resultsArr(3, 2) Then
                    resultsArr(3, 2) = .Cells(resultRow, "L").Value
                    resultsArr(3, 1) = .Cells(resultRow, "I").Value
                End If
            
            Next resultRow
        
            ' Write values stored in array to appropriate cells
            .Cells(2, "P").Value = resultsArr(1, 1)
            .Cells(2, "Q").Value = Format(resultsArr(1, 2), "Percent")
            .Cells(3, "P").Value = resultsArr(2, 1)
            .Cells(3, "Q").Value = Format(resultsArr(2, 2), "Percent")
            .Cells(4, "P").Value = resultsArr(3, 1)
            .Cells(4, "Q").Value = resultsArr(3, 2)
        
       End With

    Next wsNum
    

End Sub
