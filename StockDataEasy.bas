Attribute VB_Name = "Module1"
Sub StockDataEasy()
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Add Heading for summary
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Total Stock Volume"
        'Create Variable to hold Value
        Dim Ticker_Name As String
        Dim Vol As Double
        Vol = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
        

         ' Loop through all ticker symbol
        
        For i = 2 To LastRow
         ' Check if we are still within the same ticker symbol, if it is not...
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                ' Set Ticker name
                Ticker_Name = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = Ticker_Name
             
                ' Add Total Volumn
                Vol = Vol + Cells(i, Column + 6).Value
                Cells(Row, Column + 9).Value = Vol
                ' Add one to the summary table row
                Row = Row + 1
                
                ' reset the Volumn Total
                Vol = 0
            'if cells are the same ticker
            Else
                Vol = Vol + Cells(i, Column + 6).Value
            End If
        Next i
        
    Next ws
        
End Sub

