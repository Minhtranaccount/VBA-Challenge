Attribute VB_Name = "Module1"
Sub StockMarket():

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
        
    'Check if the loop is still in the same ticket symbol...
    'If the loop is at the lastrow of a ticket symbol
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'Set the ticket symbol
        TicketSymbol = Cells(i, 1).Value
        
        'Print the Ticket Symbol in the summary table
        Range("I" & Summary_Row).Value = TicketSymbol
        
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

End Sub

