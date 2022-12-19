Attribute VB_Name = "Module1"
Sub VBAchallenge()
'declaring variables
Dim ticker As String

Dim volume As Double

Dim y_open As Double

Dim y_close As Double

Dim y_change As Double

Dim percent_change As Double

'Establishing loop to go through all worksheets
For Each ws In Worksheets
'Setting Headers through all worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    'Integer to be used in loop
    combined_sheetvalue = 2
    
    For i = 2 To ws.UsedRange.Rows.Count
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        y_open = ws.Cells(i, 3).Value
        y_close = ws.Cells(i, 6).Value
        volume = ws.Cells(i, 7).Value
        ticker = ws.Cells(i, 1).Value
        y_change = y_close - y_open
        percent_change = (y_change / y_open) * 100
    
        
   
        
        ws.Cells(combined_sheetvalue, 9).Value = ticker
        ws.Cells(combined_sheetvalue, 10).Value = y_change
        ws.Cells(combined_sheetvalue, 11).Value = percent_change
        ws.Cells(combined_sheetvalue, 12).Value = volume
        combined_sheetvalue = combined_sheetvalue + 1
        
       
    
    End If
    Next i
   
    
   
  
    
Next ws


'Sorry about this challange. My Current job makes me unable to work on this for the time being. I plan on potenitally revisting this later in class.
'I would rather turn something in to show i actually worked on it then nothing at all.


End Sub






