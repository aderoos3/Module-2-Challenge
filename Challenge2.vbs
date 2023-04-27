Sub stock():

'identify your variables- there is a lot - so i'm giving those to you
  Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim ws As Worksheet

    
    
For Each ws In Worksheets
        j = 0
        total = 0
        change = 0
        start = 2
        
        
        
'Set title row - columsn with new names
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'find the row number of the last row with data
rowCount = Cells(Rows.Count, "A").End(xlUp).Row

' go through the whole data set starting at row 2 until the last row
For i = 2 To rowCount

' If ticker changes then print results- like in the credit card example
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
'store results in variables - we had an example like this in class
total = total + ws.Cells(i, 7).Value
'handle zero total volume- you don't want to divide by 0 in your future code
If total = 0 Then
'print the results
ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
'need 3 more lines to take care of columns J K and L
ws.Range("J" & 2 + j).Value = 0
ws.Range("K" & 2 + j).Value = "%" & 0
ws.Range("L" & 2 + j).Value = 0


Else
 ' Find First non zero starting value
If ws.Cells(start, 3) = 0 Then

      For find_value = start To i
     If ws.Cells(find_value, 3).Value <> 0 Then
      start = find_value
       'exit the whole for loop all together - wasn't covered in detail
        Exit For
        
      End If
      
      
    Next find_value
    
    
     End If
     
     'Calculate change
   change = (ws.Cells(i, 6) - ws.Cells(start, 3))
    percentChange = change / ws.Cells(start, 3)
   ' start of the next stock ticker
    
   ' print the results
ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
ws.Range("J" & 2 + j).Value = change
ws.Range("J" & 2 + j).NumberFormat = "0.00"
ws.Range("K" & 2 + j).Value = percentChange
ws.Range("K" & 2 + j).NumberFormat = "0.00%"
ws.Range("L" & 2 + j).Value = total

            
                    ' colors positives green and negative numbers red
                    Select Case change
                        Case Is > 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                        Case Else
                           ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                    End Select
                End If
                ' reset variables for new stock ticker-  more values need to be 0
                total = 0
                start = i + 1
                j = j + 1
               change = 0
               percentChange = 0
                
                
            ' If ticker is still the same add results
            Else
                total = total + ws.Cells(i, 7).Value
            End If
        Next i
        ' take the max and min and place them in a separate part in the worksheet
        'examples of max function. you need a Min too, which works similarly
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount)) * 100
        ' returns one less because header row not a factor
        'Another function - Match
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
        ' final ticker symbol for  total, greatest % of increase and decrease, and average
        ws.Range("P2") = ws.Cells(increase_number + 1, 9).Value
        ws.Range("P3") = ws.Cells(decrease_number + 1, 9).Value
        ws.Range("P4") = ws.Cells(volume_number + 1, 9).Value
    Next ws
    
End Sub


