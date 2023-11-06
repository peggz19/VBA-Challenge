Sub VBAChallenge()

' Ensure it does it on each worksheet

For Each ws In Worksheets


' Name my columns and format them to look nice

ws.Range("M1") = "Ticker"
ws.Range("N1") = "Yearly Change"
ws.Range("O1") = "Percent Change"
ws.Range("P1") = "Total Stock Volume"
ws.Range("M1:W1").Columns.AutoFit
ws.Range("M1:W1").Font.Bold = True
ws.Range("S2") = "Greatest % Increase"
ws.Range("S3") = "Greatest % Decrease"
ws.Range("S4") = "Greatest Total Volume"
ws.Range("T1") = "Ticker"
ws.Range("U1") = " Value"
ws.Range("S2:S4").Font.Bold = True

' Extract the last row number

    Dim last_row As Long

    
        last_row = ws.Range("A:A").End(xlDown).Row
        
    

' Extract each ticker and write them on the summary table

    Dim summary_row As Long
    Dim stock_volume As Variant
    Dim i As Long
    Dim open_price_counter As Integer
    
    summary_row = 2
    stock_volume = 0
    open_price_counter = -1
    
    For i = 2 To last_row
    
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        
            stock_volume = stock_volume + ws.Cells(i, 7).Value
            open_price_counter = open_price_counter + 1
            
            
    
        Else
        
            stock_volume = stock_volume + ws.Cells(i, 7).Value
            open_price_counter = open_price_counter + 1
        
            ws.Cells(summary_row, 13).Value = ws.Cells(i, 1).Value
            ws.Cells(summary_row, 16).Value = stock_volume
            ' to calculate the below, we are basically doing the last price - open price using a variable that brings "i" back to the row of the open price of the begining of the year
            ws.Cells(summary_row, 14).Value = ws.Cells(i, 6).Value - ws.Cells(i - open_price_counter, 3).Value
            ws.Cells(summary_row, 15).Value = (ws.Cells(i, 6).Value - ws.Cells(i - open_price_counter, 3).Value) / (ws.Cells(i - open_price_counter, 3).Value)
            
            
' Reset the values of my variables and ensure we enter the new ticker in a new row

            summary_row = summary_row + 1
            stock_volume = 0
            open_price_counter = -1
           
    
        End If
                

    Next i

' Change the format of the Percent Change column

ws.Range("O:O").NumberFormat = "0.00%"

' Change the color of each Yearly Change by using a new variable for our new generated table


Dim last_summary_row As Integer
Dim j As Integer

    last_summary_row = ws.Range("M:M").End(xlDown).Row

    For j = 2 To last_summary_row
    
    If ws.Cells(j, 14) < 0 Then
    ws.Cells(j, 14).Interior.ColorIndex = 3
    
    Else
    
    ws.Cells(j, 14).Interior.ColorIndex = 4

    End If


   
' Find the greatest % Increase, Greatest % Decrease and the greatest total Volume



    Set rngO = ws.Range("O:O")
    Set rngP = ws.Range("P:P")


'Find the Ticker with the expected greatest value and that value"


    If ws.Cells(j, 15) = Application.WorksheetFunction.Max(rngO) Then
    
        ws.Range("T2") = ws.Cells(j, 13).Value
        ws.Range("U2") = ws.Cells(j, 15).Value
        
    End If
    
    If ws.Cells(j, 15) = Application.WorksheetFunction.Min(rngO) Then
    
        ws.Range("T3") = ws.Cells(j, 13).Value
        ws.Range("U3") = ws.Cells(j, 15).Value
        
    End If
    
    If ws.Cells(j, 16) = Application.WorksheetFunction.Max(rngP) Then
    
        ws.Range("T4") = ws.Cells(j, 13).Value
        ws.Range("U4") = ws.Cells(j, 16).Value
        
    End If


  Next j
  
' Change the format of the "Greatest of" table

ws.Range("U2:U3").NumberFormat = "0.00%"
ws.Range("S:U").Columns.AutoFit

Next ws

End Sub




