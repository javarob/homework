Attribute VB_Name = "Module2"
Sub stock_ticker()


  Dim lRow As Long
  Dim lCol As Long
  Dim i, j As Long
  
  Dim Ticker_Name As String
  Dim Summary_Table_Row As Integer
  Dim Ticker_Total As Double
  
  Dim year_start_opening As Double
  Dim year_end_closing As Double
  
  Dim ws As Worksheet
  Dim starting_ws As Worksheet
  Set starting_ws = ActiveSheet
  

  For Each ws In Worksheets
    ws.Activate
    Ticker_Total = 0
    Summary_Table_Row = 2
    year_start_opening = 0
    year_end_closing = 0
    yearly_change = 0

      'Find the last non-blank cell in column

      lRow = Cells(Rows.Count, 1).End(xlUp).Row

      'Find the last non-blank cell in row 1
      'lCol = Cells(1, Columns.Count).End(xlToLeft).Column

      'MsgBox "Last Row: " & lRow & vbNewLine & _
      '    "Last Column: " & lCol
      
      j = 2 'counter for start price
      For i = 2 To lRow
      
      year_start_opening = Cells(j, 3)

      ' Check if we are still within the same ticker, if it is not...
      If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
  
        j = i + 1 ' set to next starting price
        
        ' Grab last closing price
        year_end_closing = Cells(i, 6).Value
        
        ' Set ticker
        Ticker_Name = Cells(i, 1).Value

        ' Add to the ticker Total
        Ticker_Total = Ticker_Total + Cells(i, 7).Value
        'MsgBox ("Ticker_Total is" + Str(Cells(i, 7).Value))

        ' Print the ticker symbol in the Summary Table
        Range("I1").Value = "Ticker Symbol"
        Range("I" & Summary_Table_Row).Value = Ticker_Name
  
        ' Print the Ticker vol to the Summary Table
        Range("L1").Value = "Ticker Yearly Volume"
        Range("L" & Summary_Table_Row).Value = Ticker_Total
        
        ' Print starting & closing columns
        'Range("L1").Value = "Opening Price"
        'Range("M1").Value = "Closing Price"
        'Range("L" & Summary_Table_Row).Value = year_start_opening
        'Range("M" & Summary_Table_Row).Value = year_end_closing
        
        Range("J1").Value = "Yearly Change"
        yearly_change = year_end_closing - year_start_opening
        If yearly_change > 0 Then
          Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        Else
          Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
        Range("J" & Summary_Table_Row).Value = yearly_change
        
        ' Calc % change over year
        Range("K1").Value = "Percent Change"
        If year_start_opening = 0 Then
           Range("K" & Summary_Table_Row).Value = "Zero Error"
        Else
           Range("K" & Summary_Table_Row).Value = Round(((year_end_closing - year_start_opening) / year_start_opening), 2)
        End If
        Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1

       ' Reset the ticker Total
        Ticker_Total = 0

      ' If the cell immediately following a row is the same ticker ...
      Else
 
        ' Add to the ticker Total
        Ticker_Total = Ticker_Total + Cells(i, 7).Value
 
      End If

    Next i 'end of ticket calc
    
  MsgBox ws.Name
  
  Next ws 'end WS
  
  starting_ws.Activate

End Sub
  
Sub clear()

Dim Ticker_Name As String
Dim lRow As Long
Dim lCol As Long
Dim Summary_Table_Row As Long

Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

  For Each ws In Worksheets
    ws.Activate

    Summary_Table_Row = 1

    'Find the last non-blank cell in column
    lRow = Cells(Rows.Count, 9).End(xlUp).Row

    'Find the last non-blank cell in row 1
    'lCol = Cells(1, Columns.Count).End(xlToLeft).Column

    'MsgBox "Last Row: " & lRow & vbNewLine & _
    '       "Last Column: " & lCol

    For i = 1 To lRow

      ' clear cells
      Range("I" & Summary_Table_Row).Value = ""
      Range("J" & Summary_Table_Row).Value = ""
      Range("K" & Summary_Table_Row).Value = ""
      Range("L" & Summary_Table_Row).Value = ""
      Range("M" & Summary_Table_Row).Value = ""
      Range("O" & Summary_Table_Row).Value = ""
      Range("J" & Summary_Table_Row).Interior.ColorIndex = 2
      
      Summary_Table_Row = Summary_Table_Row + 1

    Next i
    
    MsgBox (ws.Name)
  Next ws
  
  starting_ws.Activate

End Sub

  

 

