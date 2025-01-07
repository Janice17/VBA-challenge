Attribute VB_Name = "Module1"
Sub stock()
    
    Dim i As Long
    ' Create a variable to hold worksheet
    Dim ws As Worksheet
    
    ' Set worksheet
    
    
' --------------------------------------------
' LOOP THROUGH ALL SHEETS
' --------------------------------------------
For Each ws In Worksheets

    ' --------------------------------------------
    ' EXTRACT AND INSERT THE TICKER, QUARTERLY CHANGE, PERCENTAGE CHANGE, TOTAL STOCK VOLUME
    ' --------------------------------------------
    
    ' Creat variables to hold LastRow as a Long, and Quarter() as a String Array
    Dim LastRow As Long
    ' Create variables for summary quarterly calculations
    Dim Quarterlychange As Double
    Dim Percentchange As Double
    Dim Totalvolume As Double
    Dim opening As Double
    Dim closing As Double
    Dim greatestincrease As Double
    Dim greatestdecrease As Double
    Dim greatestvolume As Double
    Dim greatestincreaseticker As String
    Dim greatestdecreaseticker As String
    Dim greatestvolumeticker As String
    
    ' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    ' Set an initial variable for holding the ticker symbol
    Dim ticker As String
    
    ' Store first opening value and set a boolean
    'Dim openingcap As Boolean
    
    ' Keep track of the location for each ticker symbol in summary table
    Dim Summary_ticker As Long
    Summary_ticker = 2
    ' Set variables to zero
    greatestincrease = 0
    greatestdecrease = 0
    greatestvolume = 0
    Totalvolume = 0
    
    
    
    ' Loop through all the ticker symbol
    For i = 2 To LastRow
        
        If openingcap = False Then
            ' Quarterly opening price at the beginning of the quarter
            opening = ws.Cells(i, 3).Value
            ' ensures no future prices are captured until condition met.
            openingcap = True
        End If
        
        ' Check if we are still within the same ticker symbol, if we are not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        ' Set the ticker symbol
        ticker = Cells(i, 1).Value
        
        ' Quarterly closing price at the end of the quarter
        closing = ws.Cells(i, 6).Value
        
        ' Quarterly change
        Quarterlychange = closing - opening
        
        ' The total stock volume of the stock
        Totalvolume = Totalvolume + Cells(i, 7).Value
        
        

        ' The percent change from opening price at the beginning of a given quarter to the closing price at the end of that quarter
        If opening <> 0 Then
        Percentchange = ((closing - opening) / opening)
        Else
        Percentchange = 0
        End If
        
        ' Print the ticker symbol in the summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(Summary_ticker, 9).Value = ticker
        ' Print the quarterly change in the summary table
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(Summary_ticker, 10).Value = Quarterlychange
        ws.Cells(Summary_ticker, 10).NumberFormat = "0.00"
        ' Print the percent change in the summary table and format to percentage
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(Summary_ticker, 11).Value = Percentchange
        ws.Cells(Summary_ticker, 11).NumberFormat = "0.00%"
        ' Print total volume in the summary table
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(Summary_ticker, 12).Value = Totalvolume
        
        ' Conditional formatting for positive and negative yearly changes
        If Quarterlychange > 0 Then
        ws.Cells(Summary_ticker, 10).Interior.Color = RGB(0, 255, 0)
        ElseIf Quarterlychange < 0 Then
        ws.Cells(Summary_ticker, 10).Interior.Color = RGB(255, 0, 0)
        Else
        ws.Cells(Summary_ticker, 10).Interior.Color = RGB(255, 255, 255)
        End If
        
        If Percentchange > greatestincrease Then
                greatestincrease = Percentchange
                greatestincreaseticker = ticker
        End If
                
        If Percentchange < greatestdecrease Then
                greatestdecrease = Percentchange
                greatestdecreaseticker = ticker
        End If
        
        If Totalvolume > greatestvolume Then
            greatestvolume = Totalvolume
            greatestvolumeticker = ticker
        End If
        
        ' Reset the Total volume
        Totalvolume = 0
        ' Add summary ticker line
        Summary_ticker = Summary_ticker + 1
        ' Switch your condition back to capture the next price line.
        openingcap = False
        
        Else
        Totalvolume = Totalvolume + Cells(i, 7).Value
        
        End If
        

    Next i
        
        ' Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        ws.Cells(2, 15).Value = greatestincreaseticker
        ws.Cells(2, 16).Value = greatestincrease
        ws.Cells(2, 16).NumberFormat = "0.00%"
        
        ws.Cells(3, 15).Value = greatestdecreaseticker
        ws.Cells(3, 16).Value = greatestdecrease
        ws.Cells(3, 16).NumberFormat = "0.00%"
        
        ws.Cells(4, 15).Value = greatestvolumeticker
        ws.Cells(4, 16).Value = greatestvolume
        ws.Cells(4, 16).NumberFormat = "0.00E+00"
        
        ' Adjust column widith for figures
        ws.Range("I:L").EntireColumn.AutoFit
        ws.Range("N:P").EntireColumn.AutoFit
Next ws

    
End Sub
