Attribute VB_Name = "Module1"
Sub Worksheet_Loop()
    
    'Reference/credit for this sub only (apply macro to multiple sheets):
    'Please see https://github.com/davidjaimes/yearly-stock-market-analysis
        
    Dim xsheet As Worksheet
    For Each xsheet In ThisWorkbook.Worksheets
        xsheet.Select
        Call stock_data
        xsheet.Range("I:Q").Columns.AutoFit
    Next xsheet

End Sub




Sub stock_data():

    'Set an initial variable for the Ticker Symbol
    Dim Ticker_Symbol As String
        
    'Set an initial variable for Yearly Change
    Dim Yearly_Change As Double
    Yearly_Change = 0
    
    'Set an initial variable for Opening Price
    Dim Opening_Price As Double
    Opening_Price = 0
    
    'Set an initial variable for Closing Price
    Dim Closing_Price As Double
    Closing_Price = 0
    
    'Set an initial variable for Percent Change
    Dim Percent_Change As Double
    Percent_Change = 0
    
    'Set an initial variable for Total Stock Volume
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
    
    'Create Summary Table Column Headers
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    
    'Create Table for Greatest Inrease, Decrease, and Total Volume
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"
    
    'Set a variable for the Ticker symbol with the greatest percent increase
    Dim Max_Ticker_Symbol As String
    Max_Ticker_Symbol = " "
    
    'Set a variable for the Ticker symbol with the greatest percent decrease
    Dim Min_Ticker_Symbol As String
    Min_Ticker_Symbol = " "
    
    'Set a variable for the Max percent
    Dim Max_Percent As Double
    Max_Percent = 0
    
    'Set a variable for the Min percent
    Dim Min_Percent As Double
    Min_Percent = 0
    
    'Set a variable for the stock ticker symbol with the greatest total volume
    Dim Greatest_Volume_Ticker As String
    Greatest_Volume_Ticker = " "
    
    'Set a variale for the Max_Volume
    Dim Max_Volume As Double
    Max_Volume = 0
    
    'Keep track of the location for each Ticker Symbol in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'Set initial value of Opening Price for the first Ticker Symbol
    Opening_Price = Cells(2, 3).Value
    
    'Define last row
    Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
            
    'Loop through all Ticker Symbols
    For i = 2 To Lastrow
          
        'Check to see if we are still on the same Ticker Symbol, if we are not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Set the Ticker Symbol
            Ticker_Symbol = Cells(i, 1).Value
            
            'Calculate the Yearly Change
            Closing_Price = Cells(i, 6).Value
            Yearly_Change = Closing_Price - Opening_Price
            
            'Calculate Percent Change (include instructions for 0 values)
            If Opening_Price <> 0 Then
                Percent_Change = (Yearly_Change / Opening_Price) * 100
            
            End If
                        
            'Calculate Total Stock Volume
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                                                                                              
            'Print the Ticker Symbol to the Summary Table
            Range("I" & Summary_Table_Row).Value = Ticker_Symbol
                        
            'Print the Yearly Change to the Summary Table
            Range("J" & Summary_Table_Row).Value = Yearly_Change
            
            'Format Yearly Change Cell Color (Red for decrease, Green for increase)
            If (Yearly_Change > 0) Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
            ElseIf (Yearly_Change < 0) Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            
            End If
                                    
            'Print Percent Change to the Summary Table
            Range("K" & Summary_Table_Row).Value = Percent_Change
                                    
            'Format the Percent change to display as a percentage
            Range("K" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")
                        
            'Print Total Stock Volume to the Summary Table
            Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                                  
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
                        
            'Get next Opening Price
            Opening_Price = Cells(i + 1, 3).Value
            
            'Calculate Greatest Inrease, Decrease, and Total Volume
            'Credit/acknowledgement reference for this section'
            'https://splynters.com/stock-market-data
            
                If (Percent_Change > Max_Percent) Then
                Max_Percent = Percent_Change
                Max_Ticker_Symbol = Ticker_Symbol
                
            ElseIf (Percent_Change < Min_Percent) Then
                Min_Percent = Percent_Change
                Min_Ticker_Symbol = Ticker_Symbol
                
            End If
            
            If (Total_Stock_Volume > Max_Volume) Then
                Max_Volume = Total_Stock_Volume
                Greatest_Volume_Ticker = Ticker_Symbol
            
            End If
            
            'Print to Greatest Increase, Decrease, Total Volume Table
            Range("Q2").Value = (CStr(Max_Percent) & "%")
            Range("Q3").Value = (CStr(Min_Percent) & "%")
            Range("P2").Value = Max_Ticker_Symbol
            Range("P3").Value = Min_Ticker_Symbol
            Range("Q4").Value = Max_Volume
            Range("P4").Value = Greatest_Volume_Ticker
            
            'Reset values
            Percent_Change = 0
            Total_Stock_Volume = 0
                                            
        'if the next Ticker Symbol is the same as the Ticker Symbol in the previous row
        Else
        
            'Keep adding to Total Stock Volume
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        
        End If
                                                       
    Next i
    
End Sub




































