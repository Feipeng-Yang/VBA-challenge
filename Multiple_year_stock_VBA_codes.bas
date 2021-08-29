Attribute VB_Name = "Module1"
Sub Stock_Summary()

    ' ****************************
    ' Loop through all sheets
    ' ****************************
    For Each ws In Worksheets

        ' ***********************
        ' Summarize the Stocks
        ' ***********************
        
        ' Define varialbes
        Dim Row_Number As Double             ' row number of the data set
        Dim Summary_Row_Number As Double      ' row number in the summary list
        Summary_Row_Number = 2
        
        ' Set an variable for holding total stock volume for each stock
        Dim Total_Volume As Double
        Total_Volume = 0
        
        ' Set varialbes for holding opening and closing price
        Dim opening_price As Double
        Dim closing_price As Double
        Dim percent_change As Double
        Dim yearly_change As Double
        
        ' Set variables for greatest increase & decrease & volume
        Dim current_increase As Double
        Dim current_decrease As Double
        Dim current_volume As Double
        Dim increase_ticker As String
        Dim decrease_ticker As String
        Dim volume_ticker As String
        
        current_increase = 0
        current_decrease = 0
        current_volume = 0
        
        ' initializae opeing and closing price to the first entry
        opening_price = ws.Cells(2, 3).Value
        closing_price = ws.Cells(2, 6).Value
        
        
        ' Put ticker, yearly change, percent change and total stock volume on the first row
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' put greatest increase, decrease, volume to cells
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ' Count row number in this data set
        Row_Number = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Start at row 2, for each stock
        ' identify opening price at the beginning of the year
        ' identify closing price at the end of the year
        ' calulate total volume of the stock
        ' initialize opening price
        opening_price = ws.Cells(2, 3).Value
        
        For i = 2 To Row_Number
    
            'Check if it's still the same stock
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' if the next row is not the same stock
                ' add the values of last stock to the summary table
                ' Add ticker to "Ticker" column
                ws.Cells(Summary_Row_Number, 9).Value = ws.Cells(i, 1).Value
                
                ' Add closing price
                closing_price = ws.Cells(i, 6).Value
                
                ' Add to Total Volume
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                
                ' Add yearly price change to "Yearly Change" column
                yearly_change = closing_price - opening_price
                ws.Cells(Summary_Row_Number, 10).Value = yearly_change
                
                ' positive change in green and negative change in red
                If yearly_change < 0 Then
                    ws.Cells(Summary_Row_Number, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(Summary_Row_Number, 10).Interior.ColorIndex = 4
                End If
                    
                ' Add percent change to "Percent Change" column
                If opening_price <> 0 Then      'make sure the Denominator is not 0
                
                    percent_change = yearly_change / opening_price
                    
                Else                            'in case the Denominator is 0, make the percent_change 0
                    percent_change = 0
                    
                End If

                ws.Cells(Summary_Row_Number, 11).Value = percent_change
                
                ' Add total volume to "Total Stock Volume" column
                ws.Cells(Summary_Row_Number, 12).Value = Total_Volume
    
                ' move to the next row in summary table
                Summary_Row_Number = Summary_Row_Number + 1
                
                ' reset opening_price
                opening_price = ws.Cells(i + 1, 3).Value
                
                ' reset total stock volume
                Total_Volume = 0
            Else            ' if the next row is stil the same stock
                
                ' Add to Total Volume
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
        ws.Range("K2:K" & Summary_Row_Number).Style = "Percent"
        ws.Range("K2:K" & Summary_Row_Number).NumberFormat = "0.00%"
        
        ' **********************************
        ' Bunos
        '***********************************
        
        'ws.Cells(i, j).Style = "Currency"
        ' Find greatest increase & decrease & volume
        For j = 2 To Summary_Row_Number
        
            If ws.Cells(j, 11).Value > current_increase Then
            
                current_increase = ws.Cells(j, 11).Value
                increase_ticker = ws.Cells(j, 9).Value
                
            ElseIf ws.Cells(j, 11).Value < current_decrease Then
            
                current_decrease = ws.Cells(j, 11).Value
                decrease_ticker = ws.Cells(j, 9).Value
                
            End If
            
            If ws.Cells(j, 12) > current_volume Then
                current_volume = ws.Cells(j, 12).Value
                volume_ticker = ws.Cells(j, 9).Value
            End If
            
        Next j
        
        ' put value of greatest increase, decrease, volume to cells
        ws.Range("P2").Value = increase_ticker
        ws.Range("P3").Value = decrease_ticker
        ws.Range("P4").Value = volume_ticker
        
        ws.Range("Q2").Value = current_increase
        ws.Range("Q3").Value = current_decrease
        ws.Range("Q4").Value = current_volume
        
        ' Formating a bit
        ws.Range("Q2:Q3").Style = "Percent"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "0.0000E+00"
        'Autofit the summary table
        ws.Range("I1:Q1").EntireColumn.AutoFit
        
    Next ws
    
    MsgBox ("Complete!")
    
End Sub



