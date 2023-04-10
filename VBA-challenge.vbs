Sub Stocks():

' Declaring Variables

Dim WorksheetName As String
Dim Ticker_Name As String
Dim i As Long
Dim j As Long
Dim k As Long
Dim Begin As Double
Dim Ending As Double
Dim Change As Double
Dim Percent As Double
Dim Vol_Total As LongLong
Dim Max As Double
Dim Max_Ticker As String
Dim Min As Double
Dim Min_Ticker As String
Dim Volume As LongLong
Dim Volume_Ticker As String

'For each worksheet
    
For Each ws In Worksheets

WorksheetName = ws.Name
Ticker_Name = " "
Summary_row = 2

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Begin = ws.Cells(2, 3).Value
Vol_Total = 0

'Put titles on corresponding columns

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Columns("I:O").AutoFit
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
    
' Looping for raw data
    
For i = 2 To LastRow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    ' Gets the ticker name
            
        Ticker_Name = ws.Cells(i, 1).Value
        ws.Cells(Summary_row, 9).Value = Ticker_Name
            
    ' Gets the Total Volume
            
        Vol_Total = Vol_Total + ws.Cells(i, 7).Value
        ws.Cells(Summary_row, 12).Value = Vol_Total
    
    ' Computes for the Yearly Change
    
        Ending = ws.Cells(i, 6).Value
        Change = Ending - Begin
        ws.Cells(Summary_row, 10).Value = Change
        ws.Cells(Summary_row, 10).NumberFormat = "0.00"
    
    ' Fills Cell Color Green when +, Red when - and Yellow when 0
    
        If ws.Cells(Summary_row, 10).Value < 0 Then
            ws.Cells(Summary_row, 10).Interior.Color = vbRed
        ElseIf ws.Cells(Summary_row, 10).Value > 0 Then
            ws.Cells(Summary_row, 10).Interior.Color = vbGreen
        Else: ws.Cells(Summary_row, 10).Interior.Color = vbYellow
        End If
        
    ' Computes for the Percent Change
    
        Percent = Change / Begin
        ws.Cells(Summary_row, 11).Value = Percent
        ws.Cells(Summary_row, 11).NumberFormat = "0.00%"
        
    ' Resets to different values
            
        Summary_row = Summary_row + 1
        Begin = ws.Cells(i + 1, 3).Value
        Ticker_Name = " "
        Vol_Total = 0

        Else
        
        Vol_Total = Vol_Total + ws.Cells(i, 7).Value
            
        End If
    Next i

 'Greatest % Increase & Greatest % Decrease
 
Max = ws.Cells(2, 11).Value
Max_Ticker = ws.Cells(2, 9).Value
Min = ws.Cells(2, 11).Value
Min_Ticker = ws.Cells(2, 9).Value

' Looping for Summary Data in % Change

LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    For j = 2 To LastRow
    
       ' Greatest % Increase
       
       If Max < ws.Cells(j + 1, 11).Value Then
            Max = ws.Cells(j + 1, 11).Value
            Max_Ticker = ws.Cells(j + 1, 9).Value
        End If
        
       ' Greatest % Decrease
       
        If Min > ws.Cells(j + 1, 11).Value Then
            Min = ws.Cells(j + 1, 11).Value
            Min_Ticker = ws.Cells(j + 1, 9).Value
        End If
        
        ws.Range("Q2").Value = Max
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P2").Value = Max_Ticker
        ws.Range("Q3").Value = Min
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P3").Value = Min_Ticker
    Next j
    
' Greatest Total Volume

Volume = ws.Cells(2, 12).Value
Volume_Ticker = ws.Cells(2, 9).Value

' Looping for Summary Data in Total Stock Volume

LastRow = ws.Cells(Rows.Count, 12).End(xlUp).Row
    For k = 2 To LastRow
    
        If Volume < ws.Cells(k + 1, 12).Value Then
            Volume = ws.Cells(k + 1, 12).Value
            Volume_Ticker = ws.Cells(k + 1, 9).Value
        End If

        ws.Range("Q4").Value = Volume
        ws.Range("P4").Value = Volume_Ticker
Next k
    
Next ws
MsgBox ("Computation Complete!")
End Sub
