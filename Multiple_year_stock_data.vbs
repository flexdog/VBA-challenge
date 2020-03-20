Sub stock()

    For Each ws In Worksheets
       Dim WorksheetName As String
       ' Set initial variable for holding the stock ticker
       Dim ticker_name As String
       WorksheetName = ws.Name
       ' Set variable for holding max and min values per ticker
       Dim ticker_min As Double
       Dim ticker_max As Double
       'ticker_max = 0
       'ticker_min = 100000000000000#
       
       ' Set ticker summary data
       Dim ticker_volume As Double
       ticker_volume = 0
       Dim ticker_initial As Double
       Dim ticker_final As Double
       Dim ticker_percent As Double
       Dim ticker_change As Double
       
       ' Set final data
       Dim p_increase As Double
       Dim p_decrease As Double
       Dim v_greatest As Double
       
       'Setup column titles
       Dim t_ticker As String
       t_ticker = "Ticker"
       ws.Range("J1").Value = t_ticker
       ws.Range("J1").Font.Bold = True
       Dim t_year_chg As String
       t_year_chg = "Yearly Change"
       ws.Range("K1").Value = t_year_chg
       ws.Range("K1").Font.Bold = True
       Dim t_percent_chg As String
       t_percent_chg = "Percent Change"
       ws.Range("L1").Value = t_percent_chg
       ws.Range("L1").Font.Bold = True
       Dim t_total_stk_vol As String
       t_total_stk_vol = "Total Stock Volume"
       ws.Range("M1").Value = t_total_stk_vol
       ws.Range("M1").Font.Bold = True
       Dim t_value As String
       t_value = "Value"
       ws.Range("Q1").Value = t_value
       ws.Range("Q1").Font.Bold = True
       ws.Range("P1").Value = t_ticker
       ws.Range("P1").Font.Bold = True
       Dim t_great_percent_in As String
       t_great_percent_in = "Greatest % Increase"
       ws.Range("O2").Value = t_great_percent_in
       ws.Range("O2").Font.Bold = True
       Dim t_great_percent_de As String
       t_great_percent_de = "Greatest % Decrease"
       ws.Range("O3").Value = t_great_percent_de
       ws.Range("O3").Font.Bold = True
       Dim t_great_vol As String
       t_great_vol = "Greatest Total Volume"
       ws.Range("O4").Value = t_great_vol
       ws.Range("O4").Font.Bold = True
       
       'Regular expression setup
       Dim regexp As Object
       Dim strInput As String
       Dim strPattern As String
       
       Set regexp = CreateObject("vbscript.regexp")
       
       'Set pattern to match 01Jan for start of year
       strPattern = "([0]{1}[1]{1}[0]{1}[1]{1}$)"
       
       With regexp
         .Global = False
         .MultiLine = False
         .ignoreCase = True
         .Pattern = strPattern
       End With
       
       
        
       ' Keep track of the location for each ticker in summary table
       Dim summary_table_row As Integer
       summary_table_row = 2
       ' Find max row number of raw data table
       LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
       ' Find max row number of summery table
       s_LastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
       ' Loop through all tickers
       For I = 2 To LastRow
           ' Conditional stop for debugging
           'If I = 549049 Then Stop
           
           ' Check if we are still in the same ticker range
           If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
               ' Set ticker
               ticker_name = ws.Cells(I, 1).Value
               ' Collect ticker volume
               ticker_volume = ticker_volume + ws.Cells(I, 7).Value
               ' Print ticker in summary table
               ws.Range("J" & summary_table_row).Value = ticker_name
               
               'Debug stop
               'If stringinput = "20150803" Then Stop
               
               
               ' Calculate Yearly Change
               
               'If ticker starts after 01Jan
               next_ticker_value = ws.Cells(I + 1, 3).Value
               If ticker_initial = 0 And next_ticker_value > 0 Then
                  ticker_initial = ws.Cells(I + 1, 3).Value
               End If
               
               ticker_final = ws.Cells(I, 6).Value
               ticker_change = ticker_final - ticker_initial
               ws.Range("K" & summary_table_row).Value = ticker_change
               ' Color cell backgroud based on negative or positive ticker change
               If ticker_change > 0 Then
                  ws.Range("K" & summary_table_row).Interior.ColorIndex = 4
               Else:
                  ws.Range("K" & summary_table_row).Interior.ColorIndex = 3
               End If
               ' Generate the anual percent change
               ticker_percent = ticker_change / ticker_initial
               ws.Range("L" & summary_table_row).Value = ticker_percent
               ws.Range("L" & summary_table_row).NumberFormat = "0.00%"
               ' Add one to the summary table row
               summary_table_row = summary_table_row + 1
               
               ' Reset ticker total
               ticker_total = 0
               
               ' If cell below is the same ticker
             Else
               ' Add to ticker total volume
               ticker_total = ticker_total + ws.Cells(I, 7).Value
               ws.Range("M" & summary_table_row).Value = ticker_total
               ' Get ticker opening year value
               strInput = ws.Cells(I, 2).Value
               If regexp.test(strInput) Then
                  ticker_initial = ws.Cells(I, 3).Value
                  
                  'Stop for debug when ticker_initial is zero
                  'If ticker_initial = 0 Then Stop
               
               End If
               
         End If
       
       Next I
       
       ' Find max percent increase
       ws.Range("Q2").Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & s_LastRow))
       ws.Range("Q2").NumberFormat = "0.00%"
         
       ' Find related ticker
       max_increase_val = Application.WorksheetFunction.Max(ws.Range("L2:L" & s_LastRow))
       ws.Range("P2").Value = WorksheetFunction.Index(ws.Range("j2:j" & s_LastRow), WorksheetFunction.Match(max_increase_val, ws.Range("L2:L" & s_LastRow), 0))
         
       ' Find min percent decrease
       ws.Range("Q3").Value = Application.WorksheetFunction.Min(ws.Range("L2:L" & s_LastRow))
       ws.Range("Q3").NumberFormat = "0.00%"
         
       ' Find related ticker
       max_decrease_val = Application.WorksheetFunction.Min(ws.Range("L2:L" & s_LastRow))
       ws.Range("P3").Value = WorksheetFunction.Index(ws.Range("j2:j" & s_LastRow), WorksheetFunction.Match(max_decrease_val, ws.Range("L2:L" & s_LastRow), 0))
         
       ' Find max total volume
       ws.Range("Q4").Value = Application.WorksheetFunction.Max(ws.Range("M2:M" & s_LastRow))
       
       ' Find related ticker
       max_volume = Application.WorksheetFunction.Max(ws.Range("M2:M" & s_LastRow))
       ws.Range("P4").Value = WorksheetFunction.Index(ws.Range("j2:j" & s_LastRow), WorksheetFunction.Match(max_volume, ws.Range("M2:M" & s_LastRow), 0))
    Next ws
      
End Sub
