VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub vbastockticker()


    Dim CurrentWs As Worksheet
   
    Need_Summary_Table_Header = False
    COMMAND_SPREADSHEET = True
    
    ' Loop through all of the worksheets in the active workbook.
    For Each CurrentWs In Worksheets
    
        ' Set variable for ticker name
        Dim Ticker_Name As String
        Ticker_Name = " "
        
        ' Set variable for total per ticker name
        Dim Total_Ticker_Volume As Double
        Total_Ticker_Volume = 0
        
        ' Set new variables (* Delta as in ammount of change)
        
        Dim Open_Price As Double
        Open_Price = 0
        Dim Close_Price As Double
        Close_Price = 0
        Dim Delta_Price As Double
        Delta_Price = 0
        Dim Delta_Percent As Double
        Delta_Percent = 0
        ' Set new variables
        Dim MAX_TICKER_NAME As String
        MAX_TICKER_NAME = " "
        Dim MIN_TICKER_NAME As String
        MIN_TICKER_NAME = " "
        Dim MAX_PERCENT As Double
        MAX_PERCENT = 0
        Dim MIN_PERCENT As Double
        MIN_PERCENT = 0
        Dim MAX_VOLUME_TICKER As String
        MAX_VOLUME_TICKER = " "
        Dim MAX_VOLUME As Double
        MAX_VOLUME = 0
      
         
       
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
    
        Dim Lastrow As Long
        Dim i As Long
        
        Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

        
        If Need_Summary_Table_Header Then
           
            CurrentWs.Range("I1").Value = "Ticker"
            CurrentWs.Range("J1").Value = "Yearly Change"
            CurrentWs.Range("K1").Value = "Percent Change"
            CurrentWs.Range("L1").Value = "Total Stock Volume"
           
        Else
          
            Need_Summary_Table_Header = True
        End If
        
        ' Set initial value of Open Price for the first Ticker of CurrentWs,

        Open_Price = CurrentWs.Cells(2, 3).Value
        
        ' Loop from the beginning of the current worksheet to last row
        For i = 2 To Lastrow
        
      
            ' Check if we are still within the same ticker name,
            ' if not - write results to summary table
            If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
            
                Ticker_Name = CurrentWs.Cells(i, 1).Value
                
                ' Calculate Delta_Price and Delta_Percent
                Close_Price = CurrentWs.Cells(i, 6).Value
                Delta_Price = Close_Price - Open_Price
                 ' Check Division by 0 condition
                If Open_Price <> 0 Then
                    Delta_Percent = (Delta_Price / Open_Price) * 100
                Else
                    ' Unlikely, but check to avoid program crushing (per Google search lol)
                    MsgBox ("For " & Ticker_Name & ", Row " & CStr(i) & ": Open Price =" & Open_Price & ". Fix <open> field manually and save the spreadsheet.")
                End If
                
                ' Add to the Ticker name total volume
                Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
              
                
                ' Print the Ticker Name in the Summary Table, Column I
                
                CurrentWs.Range("I" & Summary_Table_Row).Value = Ticker_Name
                ' Print the Ticker Name in the Summary Table, Column I
                
                CurrentWs.Range("J" & Summary_Table_Row).Value = Delta_Price
                ' Fill Yearly Change
                
                If (Delta_Price > 0) Then
                    
                    'Fill column with GREEN color for positive
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
                ElseIf (Delta_Price <= 0) Then
                    
                    'Fill column with RED color for negative
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                 ' Print the Ticker Name in the Summary Table, Column I
                
                CurrentWs.Range("K" & Summary_Table_Row).Value = (CStr(Delta_Percent) & "%")
                
                ' Print the Ticker Name in the Summary Table, Column J
                
                CurrentWs.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                
                ' Add 1 to the summary table row count
                
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset Delta_Price and Delta_Percent holders
                Delta_Price = 0
                
                Close_Price = 0
                ' Next ticker's Open_Price
                Open_Price = CurrentWs.Cells(i + 1, 3).Value
              
                 ' Challenge : Populate new Summary table on right side of spreadsheet
            
                If (Delta_Percent > MAX_PERCENT) Then
                    MAX_PERCENT = Delta_Percent
                    MAX_TICKER_NAME = Ticker_Name
                ElseIf (Delta_Percent < MIN_PERCENT) Then
                    MIN_PERCENT = Delta_Percent
                    MIN_TICKER_NAME = Ticker_Name
                End If
                       
                If (Total_Ticker_Volume > MAX_VOLUME) Then
                    MAX_VOLUME = Total_Ticker_Volume
                    MAX_VOLUME_TICKER = Ticker_Name
                End If
                
                ' Reset counters
                Delta_Percent = 0
                Total_Ticker_Volume = 0
                
            
           
            Else
                ' increase the Total Ticker Volume
                Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
            End If
            
      
        Next i

        
            ' Challenge: Record all new counts to the new summary table on the right of the current spreadsheet
            If Not COMMAND_SPREADSHEET Then
            
                CurrentWs.Range("P2").Value = MAX_TICKER_NAME
                CurrentWs.Range("P3").Value = MIN_TICKER_NAME
                CurrentWs.Range("Q2").Value = (CStr(MAX_PERCENT) & "%")
                CurrentWs.Range("Q3").Value = (CStr(MIN_PERCENT) & "%")
                CurrentWs.Range("Q4").Value = MAX_VOLUME
                CurrentWs.Range("P4").Value = MAX_VOLUME_TICKER
                
            Else
                COMMAND_SPREADSHEET = False
            End If
               
        
     Next CurrentWs

End Sub
