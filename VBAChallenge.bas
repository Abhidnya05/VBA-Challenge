Attribute VB_Name = "Module3"
'Defining a function HW2 to extract required information from each worksheet
'Function will extract information on each worksheet and in combined worksheet

Sub HW2()

    For Each ws In Worksheets

        ' Create a variable to hold worksheet name
        Dim WorksheetName As String

        ' Determine the last row of the worksheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Copy and print the worksheet name
        WorksheetName = ws.Name
        MsgBox ("We are working on worksheet " + WorksheetName)
        
        ' Create a separate colums for extracted information on same worksheet
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 1
        
        Dim Ticker As String
        Dim Volume As Double
        Dim Total As Double
        Dim First As Double
        Dim Last As Double
        Dim Change As Double
        Dim Change1 As Double
        Dim Percent1 As Double
        Volume = 0
        Total = 0
        First = 0
        Last = 0
        Change = 0
        Change1 = 0
        Percent = 0
        
        For i = 2 To lastRow
          
            ' Assigned columnsfrom worksheet to defined variables
            
            Ticker = ws.Cells(i, 1).Value
            
            Volume = ws.Cells(i, 7).Value
            Total = Total + Volume
            
            First = ws.Cells(i, 3).Value
            'MsgBox (First)
            Last = ws.Cells(i, 6).Value
            'MsgBox (Last)
            Change = Last - First
            'MsgBox (Change)
            Change1 = Change1 + Change
      
            'Searches for when the value of the next cell is different than that of the current cell
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & Summary_Table_Row).Value = Ticker
    
                ws.Range("J" & Summary_Table_Row).Value = Change1
                If Change1 < 0 Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                End If
                
                If ws.Cells(i, 6).Value = 0 Then
                    Percent = "NA"
                Else
                    Percent = (Change1 / ws.Cells(i, 6).Value)
                End If
                
                ws.Range("K" & Summary_Table_Row).Value = Percent
                If Percent < 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                Else
                    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                End If
                
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                ws.Range("L" & Summary_Table_Row).Value = Total
                
                Total = 0
                Chnage1 = 0
                
                Summary_Table_Row = Summary_Table_Row + 1

            End If
        Next i
    
        'Ticker = ws.Cells(i, 1).Value
       ' ws.Range("I" & Summary_Table_Row).Value = Ticker
    
        'ws.Range("J" & Summary_Table_Row).Value = Change1
        'If Change1 < 0 Then
        '    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        'Else
         '   ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        'End If
                
        'If ws.Cells(i, 6).Value = 0 Then
        '    Percent = "NA"
        'Else
        '    Percent = (Change1 / ws.Cells(i, 6).Value) * 100
        'End If
                
        'ws.Range("K" & Summary_Table_Row).Value = Percent
        '    If Percent < 0 Then
         '       ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
         '   Else
         '       ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
         '   End If

        'ws.Range("L" & Summary_Table_Row).Value = Total
                
        'Summary_Table_Row = Summary_Table_Row + 1
        
        'lastRow1 = ws.Cells(Rows.Count, 10).End(xlUp).Row
        'MsgBox (lastRow1)
        
'        Dim r1 As Range
'        Dim r2 As Range
'        Set r1 = Range("K2:K" & Rows.Count)
'        Set r2 = Range("L2:L" & Rows.Count)
'        MinP = Application.WorksheetFunction.Min(r1)
        'MsgBox (MinP)
'        MaxP = Application.WorksheetFunction.Max(r1)
        'MsgBox (MaxP)
'        MaxV = Application.WorksheetFunction.Max(r2)
        'MsgBox (MaxV)
'        ws.Range("P2").Value = MinP
'        ws.Range("P3").Value = MaxP
'        ws.Range("P4").Value = MaxV
'        ws.Range("N2").Value = "Greatest % decrease"
'        ws.Range("N3").Value = "Greatest % increase"
'        ws.Range("N4").Value = "Greatest Total Volume"
'        ws.Range("P1").Value = "Value"
'        ws.Range("O1").Value = ws.Range("A1").Value
        
'        For i = 2 To lastRow
'            If ws.Cells(i, 11).Value = MinP Then
'                ws.Range("O2").Value = ws.Cells(i, 9).Value
'            End If
            
'            If ws.Cells(i, 11).Value = MaxP Then
'                ws.Range("O3").Value = ws.Cells(i, 9).Value
'            End If
            
'            If ws.Cells(i, 12).Value = MaxV Then
'                ws.Range("O4").Value = ws.Cells(i, 9).Value
'            End If

'        Next i
        
    Next ws
    
    For Each ws In Worksheets

        ' Determine the last row of the worksheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Copy and print the worksheet name
        WorksheetName = ws.Name
        
        Dim r1 As Range
        Dim r2 As Range
        Set r1 = Range("K2:K" & Rows.Count)
        Set r2 = Range("L2:L" & Rows.Count)
        MinP = Application.WorksheetFunction.Min(r1)
        'MsgBox (MinP)
        MaxP = Application.WorksheetFunction.Max(r1)
        'MsgBox (MaxP)
        MaxV = Application.WorksheetFunction.Max(r2)
        'MsgBox (MaxV)
        ws.Range("P2").Value = MinP
        ws.Range("P3").Value = MaxP
        ws.Range("P4").Value = MaxV
        ws.Range("N2").Value = "Greatest % decrease"
        ws.Range("N3").Value = "Greatest % increase"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Value"
        ws.Range("O1").Value = ws.Range("A1").Value
        
        For i = 2 To lastRow
            If ws.Cells(i, 11).Value = MinP Then
                ws.Range("O2").Value = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 11).Value = MaxP Then
                ws.Range("O3").Value = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 12).Value = MaxV Then
                ws.Range("O4").Value = ws.Cells(i, 9).Value
            End If

        Next i
        
      Next ws
      
    MsgBox ("Fixes Complete")
    
    ' Add a sheet named "Combined Data"
    Sheets.Add.Name = "Combined_Data"
    'Move created sheet to be first sheet
    Sheets("Combined_Data").Move Before:=Sheets(1)
    ' Specify the location of the combined sheet
    Set combined_sheet = Worksheets("Combined_Data")

    ' Loop through all sheets
    For Each ws In Worksheets

        ' Find the last row of the combined sheet after each paste
        ' Add 1 to get first empty row
        lastRow = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1

        ' Find the last row of each worksheet
        lastRowYear = ws.Cells(Rows.Count, "I").End(xlUp).Row - 1

        ' Copy the contents of each year sheet into the combined sheet
        combined_sheet.Range("A" & lastRow & ":D" & ((lastRowYear - 1) + lastRow)).Value = ws.Range("I1:L" & (lastRowYear + 2)).Value
         
        ' Autofit to display data
        combined_sheet.Columns("A:D").AutoFit
    
    Next ws

    ' Copy the headers from First sheet
    combined_sheet.Range("A1").Value = Sheets(2).Range("A1").Value
    combined_sheet.Range("B1").Value = "Yearly Change"
    combined_sheet.Range("C1").Value = "Percent Change"
    combined_sheet.Range("D1").Value = "Total Stock Volume"
    
    Dim s1 As Range
    Dim s2 As Range
    Set s1 = combined_sheet.Range("C2:C" & Rows.Count)
    Set s2 = combined_sheet.Range("D2:D" & Rows.Count)
    MinP = Application.WorksheetFunction.Min(s1)
    MsgBox (MinP)
    MaxP = Application.WorksheetFunction.Max(s1)
    MsgBox (MaxP)
    MaxV = Application.WorksheetFunction.Max(s2)
    MsgBox (MaxV)
    combined_sheet.Range("H2").Value = MinP
    combined_sheet.Range("H3").Value = MaxP
    combined_sheet.Range("H4").Value = MaxV
    combined_sheet.Range("F2").Value = "Greatest % decrease"
    combined_sheet.Range("F3").Value = "Greatest % increase"
    combined_sheet.Range("F4").Value = "Greatest Total Volume"
    combined_sheet.Range("H1").Value = "Value"
    combined_sheet.Range("G1").Value = combined_sheet.Range("A1").Value
       
    lastRow = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row

    For j = 2 To lastRow

      combined_sheet.Cells(j, 3).NumberFormat = "0.00%"
      
       If combined_sheet.Cells(j, 3).Value = MinP Then
            combined_sheet.Range("G2").Value = combined_sheet.Cells(j, 1).Value
        End If
        
        If combined_sheet.Cells(j, 3).Value = MaxP Then
            combined_sheet.Range("G3").Value = combined_sheet.Cells(j, 1).Value
        End If
        
        If combined_sheet.Cells(j, 4).Value = MaxV Then
            combined_sheet.Range("G4").Value = combined_sheet.Cells(j, 1).Value
        End If
     
        If combined_sheet.Cells(j, 2).Value > 0 Then
            combined_sheet.Cells(j, 2).Interior.ColorIndex = 4
        Else
            combined_sheet.Cells(j, 2).Interior.ColorIndex = 3
        End If
        
        If combined_sheet.Cells(j, 3).Value > 0 Then
            combined_sheet.Cells(j, 3).Interior.ColorIndex = 4
        Else
            combined_sheet.Cells(j, 3).Interior.ColorIndex = 3
        End If

    Next j
    
End Sub



