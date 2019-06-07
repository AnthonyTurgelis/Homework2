Sub stock2014()

Dim stock_name As String
Dim lastrow As Long
Dim ws As Worksheet
Dim stock_total As Double
Dim summary_table_row As Integer
summary_table_row = 2

Worksheets("2014").Activate

lastrow = Range("C" & Rows.Count).End(xlUp).Row

For i = 2 To lastrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        stock_name = Cells(i, 1).Value
        
        stock_total = stock_total + Cells(i, 7).Value
        
        Range("I" & summary_table_row).Value = stock_name
        
        Range("J" & summary_table_row).Value = stock_total
        
        summary_table_row = summary_table_row + 1
        
        stock_total = 0
        
    Else
    
        stock_total = stock_total + Cells(i, 7).Value
        
    End If
    
Next i

End Sub


Sub stock2015()

Dim stock_name As String
Dim lastrow As Long
Dim ws As Worksheet
Dim stock_total As Double
Dim summary_table_row As Integer
summary_table_row = 2

Worksheets("2015").Activate

lastrow = Range("C" & Rows.Count).End(xlUp).Row

For i = 2 To lastrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        stock_name = Cells(i, 1).Value
        
        stock_total = stock_total + Cells(i, 7).Value
        
        Range("I" & summary_table_row).Value = stock_name
        
        Range("J" & summary_table_row).Value = stock_total
        
        summary_table_row = summary_table_row + 1
        
        stock_total = 0
        
    Else
    
        stock_total = stock_total + Cells(i, 7).Value
        
    End If
    
Next i

End Sub


Sub stock2016()

Dim stock_name As String
Dim lastrow As Long
Dim ws As Worksheet
Dim stock_total As Double
Dim summary_table_row As Integer
summary_table_row = 2

Worksheets("2016").Activate

lastrow = Range("C" & Rows.Count).End(xlUp).Row

For i = 2 To lastrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        stock_name = Cells(i, 1).Value
        
        stock_total = stock_total + Cells(i, 7).Value
        
        Range("I" & summary_table_row).Value = stock_name
        
        Range("J" & summary_table_row).Value = stock_total
        
        summary_table_row = summary_table_row + 1
        
        stock_total = 0
        
    Else
    
        stock_total = stock_total + Cells(i, 7).Value
        
    End If
    
Next i

End Sub

Sub runallmultiple()

Call stock2014
Call stock2015
Call stock2016

End Sub
