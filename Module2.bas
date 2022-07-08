Attribute VB_Name = "Module2"
Sub year2019():

'declare variables
'variable for ticker
Dim ticker As String

'variable for yearly change
Dim yearly_Change As Variant


'variable for percent change
Dim perecentChange As Integer

'variable for total volume of stocl
'Dim totalVolume As Long
totalVolume = 0

    ' variable to hold the summary table starter row

    summaryTableRow = 2

    ' use function to find the last row in the sheet

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    
 ' loop from row 2 in column A out to the last row

    For Row = 2 To lastRow

' check to see if the ticker symbol changes
' if the ticker changes, do ....

      If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then

                ' first set the ticker symbol

                ticker = Cells(Row, 1).Value
                Cells(summaryTableRow, 9).Value = ticker
' add the last charge from the row

                totalVolume = totalVolume + Cells(Row, 7).Value
              
                
' add the total charges to the H column in the summary table row

                Cells(summaryTableRow, 12).Value = totalVolume
                  
                summaryTableRow = summaryTableRow + 1
                
                       ' reset the total volume  to 0
 ' MsgBox (totalVolume)
              
                totalVolume = 0
               
                
                
                
    Else

                ' if the ticker symbols stays the same, do....

                ' add on to the total volume

                totalVolume = totalVolume + Cells(Row, 12).Value
              
                
        End If
        
    Next Row
  

End Sub

