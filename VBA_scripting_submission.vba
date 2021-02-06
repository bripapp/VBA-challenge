Attribute VB_Name = "Module1"
 Sub main()
         
        ' set initial sheet variable
        Dim sheet As Worksheet
        
        ' Loop through all sheets containing data
        For Each sheet In Worksheets
        
            ' set variable for ticker
            Dim ticker As String
            ticker = " "
            
            ' set variables for tickerVolume, openingPrice, closingPrice, priceDifference, and percentDifference
            Dim tickerVolume, openingPrice, closingPrice, priceDifference, percentDifference As Double
            
            ' set all equal to 0
            tickerVolume = 0
            openingPrice = 0
            closingPrice = 0
            priceDifference = 0
            percentDifference = 0
            
            ' create row counter to keep track of new ticker location
            Dim rowCounter As Long
            rowCounter = 2
            
            ' get row count for the current worksheet
            Dim lastRow As Long
            Dim i As Long
            
            lastRow = sheet.Cells(Rows.Count, 1).End(xlUp).Row
            
            ' give columns in new analysis table individual headers
             sheet.Range("I1").Value = "Ticker"
             sheet.Range("J1").Value = "Yearly Change"
             sheet.Range("K1").Value = "Percent Change"
             sheet.Range("L1").Value = "Total Stock Volume"
            
            ' begin looping through rows, starting with 2 and ending with the last row of data
            For i = 2 To lastRow
    
                ' check if ticker symbol is the same from the last row and if not, write new data to analysis table
                If sheet.Cells(i + 1, 1).Value <> sheet.Cells(i, 1).Value Then
                
                    ' copy ticker symbol to table
                    ticker = sheet.Cells(i, 1).Value
                    
                    ' identify openingPrice and closingPrice locations
                    openingPrice = sheet.Cells(i, 3).Value
                    closingPrice = sheet.Cells(i, 6).Value
                    
                    ' calculate priceDifference and percentDifference
                    priceDifference = closingPrice - openingPrice
                    
                    ' make sure you can't divide by 0
                    If openingPrice <> 0 Then
                        percentDifference = (priceDifference / openingPrice) * 100
                        
                    End If
                    
                    ' add to ticker volume in designated cell
                    tickerVolume = tickerVolume + sheet.Cells(i, 7).Value
                    
                    ' print data to summary table
                    sheet.Range("I" & rowCounter).Value = ticker
                    sheet.Range("J" & rowCounter).Value = priceDifference
                    sheet.Range("K" & rowCounter).Value = (CStr(percentDifference) & "%")
                    sheet.Range("L" & rowCounter).Value = tickerVolume
                    
                    ' highlight increases in green, decreases in red
                     If (priceDifference > 0) Then
                        sheet.Range("J" & rowCounter).Interior.ColorIndex = 4
                        
                    ElseIf (priceDifference <= 0) Then
                        sheet.Range("J" & rowCounter).Interior.ColorIndex = 3
                        
                    End If
                    
                    ' add 1 to rowCounter
                    rowCounter = rowCounter + 1
                    
                    ' reset variables for next iteration
                    priceDifference = 0
                    closingPrice = 0
                    openingPrice = 0
                
                ' if the ticker is the same as the last row, simply add to the tickerVolume
                Else
                    tickerVolume = tickerVolume + sheet.Cells(i, 7).Value
                    
                End If
          
            Next i
             
         Next sheet
         
 End Sub


