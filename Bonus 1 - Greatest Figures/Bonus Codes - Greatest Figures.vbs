Attribute VB_Name = "Module1"
Sub StockMarket():

'- - - - - - - - - - - - - - -
'PART 1: SUMMARY TABLE
'- - - - - - - - - - - - - - -
'Add headers to the summary table
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Set a variable to hold the ticket symbol
Dim TickSymbol As String

'Set a variable to hold the Close Price
Dim ClosePrice As Double

'Set a variable to hold the Total Volume
Dim Total_Vol As Double
Total_Vol = 0

'Keep track of ticket symbol in the summary table
Dim Summary_Row As Double
Summary_Row = 2

'Set a variable to hold the last row of the sheet
Dim lastrow As Double
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Set loop row count for the same ticket symbol
Dim Loop_rowcount As Double
Loop_rowcount = 0

'Loop through all rows
For i = 2 To lastrow
    'Set Close Price
    ClosePrice = Cells(i, 6).Value
        
    'Check if the loop is still in the same ticket symbol
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'Set the ticket symbol
        Ticketsymbol = Cells(i, 1).Value
        
        'Print the Ticket Symbol in the summary table
        Range("I" & Summary_Row).Value = Ticketsymbol
        
        'Add to the Total Volume
        Total_Vol = Total_Vol + Cells(i, 7).Value
        
        'Print the total volume to the summary table
        Range("L" & Summary_Row).Value = Total_Vol
        
        'Add 1 to the loop row count
        Loop_rowcount = Loop_rowcount + 1
        
        'Set open price
        Dim OpenPrice As Double
        OpenPrice = Cells((i - Loop_rowcount + 1), 3).Value
        
        'Calculate the Yearly Change and print to the table
        Range("J" & Summary_Row).Value = ClosePrice - OpenPrice
        
        'Calculate the Percent Change and print to the table
        Range("K" & Summary_Row).Value = (ClosePrice - OpenPrice) / OpenPrice
        
        'Format the positive change in green and negative in red
        'If there is a positive change
        If Range("J" & Summary_Row).Value >= 0 Then
    
            'Highlight with Green
            Range("J" & Summary_Row).Interior.ColorIndex = 4
        
            'If there is no positive change (negative change)
            Else
        
            'Highlight with red
            Range("J" & Summary_Row).Interior.ColorIndex = 3
    
            'Close if statement
            End If

        
        'Add one to the summary table row
        Summary_Row = Summary_Row + 1
        
        'Reset the total volume
        Total_Vol = 0
        
        'Reset Loop_Rowcount
        Loop_rowcount = 0
        
    'If the next cell when running loop is the same ticket symbol
    Else
    
        'Add to the total volume
        Total_Vol = Total_Vol + Cells(i, 7).Value
        
        'Add 1 to Loop Row Count
        Loop_rowcount = Loop_rowcount + 1
        
    'Close if statement
    End If
    
'Iterate the loop
Next i

    'Change format of Percent Change to percentage
        Columns("K:K").NumberFormat = "0.00%"
        
    'Autofit
        Columns("I:L").AutoFit

'- - - - - - - - - - - - - - -
'PART 2: GREATEST FIGURES
'- - - - - - - - - - - - - - -

'Create headers for the new table if the greatest figures
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'Set a variable to hold the last row of the Summary Table
Dim lastrow_summary As Double
lastrow_summary = Cells(Rows.Count, 9).End(xlUp).Row

'Locate the biggest number of Percent Change
Dim PercentChange_Max As Double
PercentChange_Max = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary))
    
'Print Value
Range("Q2").Value = PercentChange_Max

'Print Ticker Symbol

    'Find the row of the ticketnumber
        Dim rowcount_max As Double
        rowcount_max = Application.WorksheetFunction.Match(PercentChange_Max, Range("K2:K" & lastrow_summary), 0)
        
    'Find the ticket symbol according to the rowcount
        Range("P2").Value = Range("I" & (1 + rowcount_max)).Value

     
'Locate the smallest number of Percent Change (Greatest Decrease)
Dim PercentChange_Min As Double
PercentChange_Min = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary))
    
'Print Ticker Symbol

   'Find the row of the ticketnumber
   Dim rowcount_min As Double
       rowcount_min = Application.WorksheetFunction.Match(PercentChange_Min, Range("K2:K" & lastrow_summary), 0)
       
   'Find the ticket symbol of according to the rowcount
       Range("P3").Value = Range("I" & (1 + rowcount_min)).Value

'Print Value
Range("Q3").Value = PercentChange_Min
        
   
'Locate the biggest total volumn summary
Dim Total_Vol_Sum As Double
Total_Vol_Sum = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary))
    
'Print Ticker Symbol

    'Find the row of the ticketnumber
    rowcount_tol = Application.WorksheetFunction.Match(Total_Vol_Sum, Range("L2:L" & lastrow_summary), 0)
            
    'Find the ticket symbol of according to the rowcount
    Range("P4").Value = Range("I" & (1 + rowcount_tol)).Value

'Print Value
Range("Q4").Value = Total_Vol_Sum
 
'Change format of Percent Change to percentage
Range("Q2:Q3").NumberFormat = "0.00%"

'Autofit
Columns("O:Q").AutoFit

End Sub
