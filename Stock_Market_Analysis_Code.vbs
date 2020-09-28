' Create a script that will loop through all the stocks for one year and output the following information.

' The ticker symbol.


' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.


' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

' The total stock volume of the stock.

' Conditional formatting that will highlight positive change in green and negative change in red.

Sub getticker()

MsgBox ("Bienvenido, iniciando el procedimiento")

For Each ws In Worksheets
    
    ws.Activate

' Creating headers and colors
  
    Cells(1, 10).Value = Cells(1, 1).Value
    Cells(1, 11).Value = "Yearly_change"
    Cells(1, 12).Value = "Percent_change"
    Cells(1, 13).Value = "Total_Stock_Volume"
    Range("J1:M1").Interior.ColorIndex = 25
    Range("J1:M1").Font.ColorIndex = 2
  
' Setting variables in original table
  
    Dim ticker As String
    Dim op, cl, percent_change As Double
    Dim total_stock As LongLong
      
' The summary table
  
    Dim row_summary As Integer
    row_summary = 2

' Loop for getting all tickers
  
    LowRow = Cells(Rows.Count, 1).End(xlUp).Row
    LowColumb = Cells(1, Columns.Count).End(xlToLeft).Column
    op = 0
    total_stock = 0
 
For i = 2 To LowRow
           
'Conditional to get open price
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            op = Cells(i, 3).Value
        End If
              
                   
' Check if we are still within the same ticker, if it is not oper the following:
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            cl = Cells(i, 6).Value
                                       
            ticker_Name = Cells(i, 1).Value
    
' Print the ticker in the Summary Table
        Range("J" & row_summary).Value = ticker_Name
        
            yearly_change = cl - op

' Print the Yearly_Change to the Summary Table
        Range("K" & row_summary).Value = yearly_change
        
        If yearly_change >= 0 Then
            Range("K" & row_summary).Interior.ColorIndex = 50
        Else
            Range("K" & row_summary).Interior.ColorIndex = 30
            Range("K" & row_summary).Font.ColorIndex = 2
        End If
                
' I detected a problem because of the zero values for cl and op
' Before calculate percent_change we need to implement an If so as to
' determine values equal to zero and not procedent when op=0
    
        If cl = 0 Then
            percent_change = 0
                        
            Cells(i, 12).Value = 0
            Cells(i, 12).NumberFormat = "#.00%"
        
        ElseIf op = 0 Then
            percent_change = 0
                    
        Else
        
            percent_change = ((cl / op) - 1)
        
        End If
             
' print the percent change to the summary table
        Range("L" & row_summary).Value = percent_change
        Range("L" & row_summary).NumberFormat = "#.0%"
            
            total_stock = total_stock + Cells(i, 7).Value

' Print the total stock Amount to the Summary Table
        Range("M" & row_summary).Value = total_stock

' Add one to the summary table row
            row_summary = row_summary + 1
      
' Reset the total
            total_stock = 0
            
        Else

        total_stock = total_stock + Cells(i, 7).Value

    End If
          
  Next i
         
    
'looking for the greatest % increses, greatest % decrease and greatest total volume
'tabla resumen
    
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Range("O1:Q1").Interior.ColorIndex = 25
    Range("O1:Q1").Font.ColorIndex = 2
    Range("Q2:Q3").NumberFormat = "#.00%"
    Range("O2:Q4").Interior.ColorIndex = 36
                  
' variables de tabla resumen
    
    Dim ticker2 As String
    Dim best, worst, total_volume As Double
    
          
    LowRowR = Cells(Rows.Count, 12).End(xlUp).Row
    
    For t = 2 To LowRowR
                
    If Cells(t, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & LowRowR)) Then
                Cells(2, 16).Value = Cells(t, 10).Value
                Cells(2, 17).Value = Cells(t, 12).Value
        
        ElseIf Cells(t, 12).Value = Application.WorksheetFunction.Min(Range("L2:L" & LowRowR)) Then
    
                Cells(3, 16).Value = Cells(t, 10).Value
                Cells(3, 17).Value = Cells(t, 12).Value
                
     ElseIf Cells(t, 13).Value = Application.WorksheetFunction.Max(Range("M2:M" & LowRowR)) Then
    
                Cells(4, 16).Value = Cells(t, 10).Value
                Cells(4, 17).Value = Cells(t, 13).Value
                          
        End If
                  
    Next t
      
Next ws

MsgBox ("Se ha concluido exitosamente el proceso")

End Sub