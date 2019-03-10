Sub COMBINED_STOCK()

'To loop through all the worksheets in the workbook
'Declare worksheet variable
Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets

    ws.Activate
    
'Delclare variables

        Dim Ticker_Open_Price As Double
        Dim Ticker_Close_Price As Double
        Dim Ticker_Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Ticker_Percent_Change As Double
        Dim Ticker_Volume As Variant
        Dim Table_Row As Integer
        Dim Last_Row As Long
        Dim Change_Last_Row As Long
                
        
'Assign starting values
        
        Table_Row = 2
        Ticker_Volume = 0
        

'Create headers for summary
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        

'Open Price for the first Ticker Name in a spread sheet
        Ticker_Open_Price = ws.Range("C2").Value
        

'Determine the Last Row in a spreedsheet
        Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Determine the Last Row in the summary table
        Change_Last_Row = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Loop through all ticker tickers and compile totals for groups
       
        For i = 2 To Last_Row
        

'Check if we are still within the same ticker symbol, if not then,

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            
'Determine Ticker_name and print in the summary table

                Ticker_Name = ws.Cells(i, 1).Value
                ws.Range("I" & Table_Row).Value = Ticker_Name
               
'Closing Price
                Ticker_Close_Price = ws.Cells(i, 6).Value
                
                
                
'Calculate Yearly Change and add the value to the summary table
               
                Ticker_Yearly_Change = Ticker_Close_Price - Ticker_Open_Price
                ws.Range("J" & Table_Row).Value = Ticker_Yearly_Change
                ws.Range("J1").Columns.AutoFit
                ws.Range("J" & Table_Row).NumberFormat = "0.00000000#"
'Add Percent Change and the conditions

                If (Ticker_Open_Price = 0 And Ticker_Close_Price = 0) Then
                    Ticker_Percent_Change = 0
                
                
                ElseIf (Ticker_Open_Price = 0 And Ticker_Close_Price > 0) Then
                    Ticker_Percent_Change = 1
            
                
                ElseIf (Ticker_Open_Price = 0 And Ticker_Close_Price < 0) Then
                    Ticker_Percent_Change = -1
              
                
                Else
                    Ticker_Percent_Change = Ticker_Yearly_Change / Ticker_Open_Price
                    ws.Range("K" & Table_Row).Value = Ticker_Percent_Change
                    ws.Range("K" & Table_Row).NumberFormat = "0.00%"
                    ws.Range("K1").Columns.AutoFit
                End If
                
                
'Compile the Ticker Total volume and add to the summary table
 
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
                ws.Range("L" & Table_Row).Value = Ticker_Volume
                ws.Range("L1").Columns.AutoFit
'Add one row to the summary table
                Table_Row = Table_Row + 1
                
'New Ticker Open Price
                Ticker_Open_Price = ws.Cells(i + 1, 3).Value
                
'Reset Ticker_Total_Volume
                Ticker_Volume = 0
                
'if cells immediately following a row the same ticker...

            Else
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
                
            End If
        Next i
        
        
'conditional formatting for yearly change

    For i = 2 To Change_Last_Row
    
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 43
                
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
                
            End If
        Next i
       
        
'Create row titles for the Greatest values table and autofit contents
        ws.Range("O1:Q4").Columns.AutoFit
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        
'Get the greateast value and the ticker associated


    For i = 2 To Change_Last_Row
    
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & Change_Last_Row)) Then
            
                
                
                ws.Range("Q2").Value = ws.Cells(i, 11).Value
                ws.Range("P2").Value = ws.Cells(i, 9).Value
                ws.Range("Q2").NumberFormat = "0.00%"
                
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & Change_Last_Row)) Then
             
                ws.Range("Q3").Value = ws.Cells(i, 11).Value
                ws.Range("P3").Value = ws.Cells(i, 9).Value
                ws.Range("Q3").NumberFormat = "0.00%"
                
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & Change_Last_Row)) Then
            
            
                ws.Range("Q4").Value = ws.Cells(i, 12).Value
                ws.Range("P4").Value = ws.Cells(i, 9).Value
                ws.Range("Q4").NumberFormat = "#"
            End If
        Next i
        
    Next ws
        
End Sub




