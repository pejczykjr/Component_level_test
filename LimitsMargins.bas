Attribute VB_Name = "LimitsMargins"
Option Explicit

'This Sub corrects frequency values.
Sub limitsAdd(measurementType As String, measurementFileName As String)

'   VARIABLES
    
    'Keeps range that is needed to have updated values of frequency (from Hz to MHz).
    Dim cell As Range
    
'   FUNCTIONAL PART
    
    'Assigns number of last used column.
    maxCols = ActiveSheet.UsedRange.Columns.Count
    
    'Divides frequency values to have them in MHz.
    For Each cell In Range("A2:A" + CStr(Cells(Rows.Count, 1).End(xlUp).Row))
        cell.Value = cell.Value / 1000000
    Next cell
    
    'Sets title in the new column for limit values.
    Cells(1, maxCols + 1).Value = "Limit [dB]"
    
    Call limit(measurementType, measurementFileName)
    
End Sub

'This sub sets configuration- test category and type of measurement.
Sub limit(measurementType As String, measurementFileName As String)

'   VARIABLES
    
    'These are formulas for il, next and rl.
    Dim ilFormula, nextFormula, rlFormula As String
    'Keeps address of first row where value appears.
    Dim firstValueCell As Range, fVCString As String
    
'   FUNCTIONAL PART
    
    'Sets address of cell where data begin (frequency).
    Set firstValueCell = Range("A2")
    fVCString = firstValueCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    'Checks what test category user chose from dialog box and based on that assigns formulas.
    If (testCategory = "CAT5E") Then
        'C5E
        ilFormula = "=(1.967*SQRT(" + fVCString + ")+0.023*" + fVCString + "+0.05/SQRT(" + fVCString + "))"
        nextFormula = "=-(35.3-15*LOG10(" + fVCString + "/100))"
        rlFormula = "=-IF(AND(" + fVCString + ">=1," + fVCString + "<10),20+5*LOG10(" + fVCString + "),IF(AND(" + fVCString + ">=10," + fVCString + "<20), 25, 25-7*LOG10(" + fVCString + "/20)))"

    ElseIf (testCategory = "CAT6") Then
        'C6
        ilFormula = "=(1.808*SQRT(" + fVCString + ")+0.017*" + fVCString + "+0.2/SQRT(" + fVCString + "))"
        nextFormula = "=-(44.3-15*LOG10(" + fVCString + "/100))"
        rlFormula = "=-IF(AND(" + fVCString + ">=1," + fVCString + "<10),20+5*LOG10(" + fVCString + "),IF(AND(" + fVCString + ">=10," + fVCString + "<20), 25, 25-7*LOG10(" + fVCString + "/20)))"
    ElseIf (testCategory = "CAT6A") Then
        'C6A
        ilFormula = "=(1.82*SQRT(" + fVCString + ")+0.0091*" + fVCString + "+0.25/SQRT(" + fVCString + "))"
        nextFormula = "=-(44.3-15*LOG10(" + fVCString + "/100))"
        rlFormula = "=-IF(AND(" + fVCString + ">=1," + fVCString + "<10),20+5*LOG10(" + fVCString + "),IF(AND(" + fVCString + ">=10," + fVCString + "<20), 25, 25-7*LOG10(" + fVCString + "/20)))"

    End If

    'Checks type of measurement and calls other subs.
    Select Case measurementType
        Case "il": Call ilLimit(ilFormula, measurementFileName)
        Case "next": Call nextLimit(nextFormula, measurementFileName)
        Case "rl": Call rlLimit(rlFormula, measurementFileName)
    End Select

End Sub

'Subs below add limits, calculate margins and show worst margins.
Sub ilLimit(ByVal ilFormula As String, ByVal measurementFileName As String)

'   Variables

    'These variables keep worst margin and related with it frequency.
    Dim worstMargin, worstFrequency As Double
    'Variables used in for loop, i is iterator, cell is range that goes through all the margins.
    Dim cell As Range, i As Integer
    
'   FUNCTIONAL PART

    'Adds limits - fills the whole range with this formula. Value changes automatically for the next row (no reference). maxCols+1 -> column C
    Range(Cells(2, maxCols + 1), Cells(Cells(Rows.Count, 2).End(xlUp).Row, maxCols + 1)).Formula = ilFormula
    
    'Sets title for margins. maxCols+2 -> column D
    Cells(1, maxCols + 2).Value = "Margin [dB]"
    
    'Calculates margins "Limit - Measurement". If value is >0, then pass. Else fail.
    Range(Cells(2, maxCols + 2), Cells(Cells(Rows.Count, 2).End(xlUp).Row, maxCols + 2)) = "=C2-B2"
    
    'Merges columns and adds title for worst margin. maxCols+4 -> column F
    Range(Cells(2, maxCols + 4), Cells(2, maxCols + 5)).Merge
    Cells(2, maxCols + 4).Value = "Insertion Loss WORST MARGIN"
    Cells(2, maxCols + 4).Characters(1, Len("Insertion Loss")).Font.Bold = True
    
    'Merges columns and adds title for pair which is associated with current file and measurement (e.g. orange -> pair 1,2).
    Range(Cells(3, maxCols + 4), Cells(3, maxCols + 5)).Merge
    If InStr(1, measurementFileName, "orange") <> 0 Then
        Cells(3, maxCols + 4).Value = "Pair 1,2 (Orange)"
            
    ElseIf InStr(1, measurementFileName, "brown") <> 0 Then
        Cells(3, maxCols + 4).Value = "Pair 7,8 (Brown)"
                        
    ElseIf InStr(1, measurementFileName, "green") <> 0 Then
        Cells(3, maxCols + 4).Value = "Pair 3,6 (Green)"
                        
    ElseIf InStr(1, measurementFileName, "blue") <> 0 Then
        Cells(3, maxCols + 4).Value = "Pair 4,5 (Blue)"
    End If
    
    'Adds titles for freqency and related value.
    Cells(4, maxCols + 4).Value = "Frequency [MHz]"
    Cells(4, maxCols + 5).Value = "Margin [dB]"

    'First assignment of temporary worst margin. Attributed value from first frequency.
    worstMargin = Cells(2, maxCols + 2)
    worstFrequency = Cells(2, 1)
    
    'This for loop checks if there occurs any worse value than current worstMargin and
    'reassigns this variable and corresponding frequency. It also checks if values are greater than 0 to mark cells
    'with green for pass and red for fail.
    For i = 2 To (Cells(Rows.Count, maxCols + 2).End(xlUp).Row)
    
        Set cell = Cells(i, maxCols + 2)

        If (cell.Value < worstMargin) Then
            worstMargin = cell.Value
            worstFrequency = Cells(i, 1).Value
        End If

        If (cell.Value > 0) Then
            cell.Style = "Good"
        Else
            cell.Style = "Bad"
        End If
        
    Next

    'Puts values of worst margin into cells. Formatted to 2 decimal places without rounding.
    Cells(5, maxCols + 4).Formula = "=TRUNC(" + Replace(worstFrequency, ",", ".") + ",2)"
    Cells(5, maxCols + 5).Formula = "=TRUNC(" + Replace(worstMargin, ",", ".") + ",2)"
    
    'Checks if value is greater than 0 to mark it with green for pass and red for fail.
    If (worstMargin < 0) Then
        Cells(5, maxCols + 5).Style = "Bad"
    Else
        Cells(5, maxCols + 5).Style = "Good"
    End If

    'Creates borders inside and around WORST MARGIN to make it look like a table.
    Range(Cells(2, maxCols + 4), Cells(5, maxCols + 5)).Borders.LineStyle = XlLineStyle.xlContinuous
       
End Sub

Sub nextLimit(ByVal nextFormula As String, ByVal measurementFileName As String)

'   VARIABLES

    'These variables keep worst margin and related with it frequency.
    Dim worstMargin, worstFrequency As Double
    'Variables used in for loop, i,j are iterators, cell is range that goes through all the margins.
    Dim cell As Range, i, j As Integer
    'It keeps colours of different pairs
    Dim colours() As Variant
    
'   FUNCTIONAL PART

    'Indexing from 0, that's why in colours(0) there is nothing. Indexes related with
    'colours, 1-Blue, 2-Orange...
    colours() = Array("", "Blue", "Orange", "Green", "Brown")

    'Adds limits - fills the whole range with this formula. Value changes automatically for the next row (no reference). maxCols+1 -> column E
    Range(Cells(2, maxCols + 1), Cells(Cells(Rows.Count, 2).End(xlUp).Row, maxCols + 1)).Formula = nextFormula
    
    'Merges columns and sets title for margins. maxCols+2 -> column F, maxCols+4 -> column H
    Range(Cells(1, maxCols + 2), Cells(1, maxCols + 4)).Merge
    Cells(1, maxCols + 2).Value = "Margin [dB]"
    
    'Merges columns and sets title for worst margins. maxCols+6 -> column J, maxCols+11 -> column O
    Range(Cells(2, maxCols + 6), Cells(2, maxCols + 11)).Merge
    If InStr(1, measurementFileName, "orange") <> 0 Then
        Cells(2, maxCols + 6).Value = "NEXT forward WORST MARGINS Pair 1,2 (Orange)"
            
    ElseIf InStr(1, measurementFileName, "brown") <> 0 Then
        Cells(2, maxCols + 6).Value = "NEXT forward WORST MARGINS Pair 7,8 (Brown)"
                        
    ElseIf InStr(1, measurementFileName, "green") <> 0 Then
        Cells(2, maxCols + 6).Value = "NEXT forward WORST MARGINS Pair 3,6 (Green)"
                        
    ElseIf InStr(1, measurementFileName, "blue") <> 0 Then
        Cells(2, maxCols + 6).Value = "NEXT forward WORST MARGINS Pair 4,5 (Blue)"
        
    End If
    Cells(2, maxCols + 6).Characters(1, Len("NEXT forward")).Font.Bold = True
    
    'It assigns colours (e.g. S12 = Blue to Orange) to headers in worst margins.
    Range(Cells(3, maxCols + 6), Cells(3, maxCols + 7)).Merge
    Cells(3, maxCols + 6).Value = colours(Mid(Cells(1, 2).Value, 2, 1)) + _
                                 " to " + colours(Mid(Cells(1, 2).Value, 3, 1))
    Range(Cells(3, maxCols + 8), Cells(3, maxCols + 9)).Merge
    Cells(3, maxCols + 8).Value = colours(Mid(Cells(1, 3).Value, 2, 1)) + _
                                  " to " + colours(Mid(Cells(1, 3).Value, 3, 1))
    Range(Cells(3, maxCols + 10), Cells(3, maxCols + 11)).Merge
    Cells(3, maxCols + 10).Value = colours(Mid(Cells(1, 4).Value, 2, 1)) + _
                                  " to " + colours(Mid(Cells(1, 4).Value, 3, 1))
    
    'Adds titles for freqency and related value.
    Cells(4, maxCols + 6).Value = "Frequency [MHz]"
    Cells(4, maxCols + 7).Value = "Margin [dB]"
    Cells(4, maxCols + 8).Value = "Frequency [MHz]"
    Cells(4, maxCols + 9).Value = "Margin [dB]"
    Cells(4, maxCols + 10).Value = "Frequency [MHz]"
    Cells(4, maxCols + 11).Value = "Margin [dB]"
    
    'Calculates margins "Limit - Measurement". If value is >0, then pass. Else fail.
    Range(Cells(2, maxCols + 2), Cells(Cells(Rows.Count, 1).End(xlUp).Row, maxCols + 2)) = "=E2-B2"
    Range(Cells(2, maxCols + 3), Cells(Cells(Rows.Count, 1).End(xlUp).Row, maxCols + 3)) = "=E2-C2"
    Range(Cells(2, maxCols + 4), Cells(Cells(Rows.Count, 1).End(xlUp).Row, maxCols + 4)) = "=E2-D2"
     
    'This for loop checks if there occurs any worse value than current worstMargin and
    'reassigns this variable and corresponding frequency. It also checks if values are greater than 0 to mark cells
    'with green for pass and red for fail. Values upto 10MHz are omitted (under last before 10MHz there is line for seperation).
    'This loop goes through columns.
    For i = maxCols + 2 To maxCols + 4
        
        'First assignment of temporary worst margin. Attributed value from last frequency.
        worstMargin = Cells(Cells(Rows.Count, i).End(xlUp).Row, i)
        worstFrequency = Cells(Cells(Rows.Count, 1).End(xlUp).Row, 1)
        
        'This loop goes through rows.
        For j = 2 To (Cells(Rows.Count, maxCols).End(xlUp).Row)
            
            Set cell = Cells(j, i)
                
            If (Cells(j, 1).Value >= 10) Then
                If (cell.Value < worstMargin) Then
                    worstMargin = cell.Value
                    worstFrequency = Cells(j, 1).Value
                End If
            ElseIf (Cells(j, 1).Value < 10 And Cells(j + 1, 1).Value >= 10) Then
                Cells(j, i).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
                Cells(j, i).Borders(xlEdgeBottom).Weight = xlThick
            End If
            
            If (cell.Value > 0) Then
                cell.Style = "Good"
            Else
                cell.Style = "Bad"
            End If
            
        Next
        
        'Still in the for i loop, but outside for j loop. These Ifs put values of worst margins into cells.
        'They are formatted to 2 decimal places without rounding.
        'Checks if value is greater than 0 to mark it with green for pass and red for fail.
        If (i = maxCols + 2) Then
            Cells(5, i + 4).Formula = "=TRUNC(" + Replace(worstFrequency, ",", ".") + ",2)"
            Cells(5, i + 5).Formula = "=TRUNC(" + Replace(worstMargin, ",", ".") + ",2)"
            
            If (worstMargin < 0) Then
                Cells(5, i + 5).Style = "Bad"
            Else
                Cells(5, i + 5).Style = "Good"
            End If
        
        ElseIf (i = maxCols + 3) Then
            Cells(5, i + 5).Formula = "=TRUNC(" + Replace(worstFrequency, ",", ".") + ",2)"
            Cells(5, i + 6).Formula = "=TRUNC(" + Replace(worstMargin, ",", ".") + ",2)"
        
            If (worstMargin < 0) Then
                Cells(5, i + 6).Style = "Bad"
            Else
                Cells(5, i + 6).Style = "Good"
            End If
         
        ElseIf (i = maxCols + 4) Then
            Cells(5, i + 6).Formula = "=TRUNC(" + Replace(worstFrequency, ",", ".") + ",2)"
            Cells(5, i + 7).Formula = "=TRUNC(" + Replace(worstMargin, ",", ".") + ",2)"
            
            If (worstMargin < 0) Then
                Cells(5, i + 7).Style = "Bad"
            Else
                Cells(5, i + 7).Style = "Good"
            End If
        End If
        
    Next
    
    'Creates borders inside and around WORST MARGINS to make it look like a table.
    Range(Cells(2, maxCols + 6), Cells(5, maxCols + 11)).Borders.LineStyle = XlLineStyle.xlContinuous
    
End Sub

Sub rlLimit(ByVal rlFormula As String, ByVal measurementFileName As String)

'   VARIABLES

    'These variables keep worst margin and related with it frequency.
    Dim worstMargin, worstFrequency As Double
    'Variables used in for loop, i,j are iterators, cell is range that goes through all the margins.
    Dim cell As Range, i, j As Integer

'   FUNCTIONAL PART
    
    'Adds limits - fills the whole range with this formula. Value changes automatically for the next row (no reference). maxCols+1 -> column F
    Range(Cells(2, maxCols + 1), Cells(Cells(Rows.Count, 2).End(xlUp).Row, maxCols + 1)) = rlFormula

    'Merges columns and sets title for margins. maxCols+2 -> column G, maxCols+5 -> column J
    Range(Cells(1, maxCols + 2), Cells(1, maxCols + 5)).Merge
    Cells(1, maxCols + 2).Value = "Margin [dB]"
    
    'Merges columns and adds title for pair which is associated measurement (e.g. rl-fw S11 -> pair 4,5 (Blue)).
    Range(Cells(3, maxCols + 7), Cells(3, maxCols + 8)).Merge
    Cells(3, maxCols + 7).Value = "Pair 4,5 (Blue)"
    Range(Cells(3, maxCols + 9), Cells(3, maxCols + 10)).Merge
    Range(Cells(3, maxCols + 11), Cells(3, maxCols + 12)).Merge
    Cells(3, maxCols + 11).Value = "Pair 3,6 (Green)"
    Range(Cells(3, maxCols + 13), Cells(3, maxCols + 14)).Merge
    
    'Merges columns and sets title for worst margins. maxCols+7 -> column L, maxCols+14 -> column S
    'Here are assigned "Pair 1,2" and "Pair 7,8" as names to not create Ifs above.
    Range(Cells(2, maxCols + 7), Cells(2, maxCols + 14)).Merge
    If InStr(1, measurementFileName, "fw") <> 0 Then
        Cells(2, maxCols + 7).Value = "Return Loss WORST MARGINS Forward"
        Cells(3, maxCols + 9).Value = "Pair 1,2 (Orange)"
        Cells(3, maxCols + 13).Value = "Pair 7,8 (Brown)"
        
    ElseIf InStr(1, measurementFileName, "rev") <> 0 Then
        Cells(2, maxCols + 7).Value = "Return Loss WORST MARGINS Reverse"
        Cells(3, maxCols + 9).Value = "Pair 7,8 (Brown)"
        Cells(3, maxCols + 13).Value = "Pair 1,2 (Orange)"
    
    End If
    Cells(2, maxCols + 7).Characters(1, Len("Return Loss")).Font.Bold = True
    
    'Adds titles for freqency and related value.
    Cells(4, maxCols + 7).Value = "Frequency [MHz]"
    Cells(4, maxCols + 8).Value = "Margin [dB]"
    Cells(4, maxCols + 9).Value = "Frequency [MHz]"
    Cells(4, maxCols + 10).Value = "Margin [dB]"
    Cells(4, maxCols + 11).Value = "Frequency [MHz]"
    Cells(4, maxCols + 12).Value = "Margin [dB]"
    Cells(4, maxCols + 13).Value = "Frequency [MHz]"
    Cells(4, maxCols + 14).Value = "Margin [dB]"
    
    'Calculates margins "Limit - Measurement". If value is >0, then pass. Else fail.
    Range(Cells(2, maxCols + 2), Cells(Cells(Rows.Count, 1).End(xlUp).Row, maxCols + 2)) = "=F2-B2"
    Range(Cells(2, maxCols + 3), Cells(Cells(Rows.Count, 1).End(xlUp).Row, maxCols + 3)) = "=F2-C2"
    Range(Cells(2, maxCols + 4), Cells(Cells(Rows.Count, 1).End(xlUp).Row, maxCols + 4)) = "=F2-D2"
    Range(Cells(2, maxCols + 5), Cells(Cells(Rows.Count, 1).End(xlUp).Row, maxCols + 5)) = "=F2-E2"
    
    'This for loop checks if there occurs any worse value than current worstMargin and
    'reassigns this variable and corresponding frequency. It also checks if values are greater than 0 to mark cells
    'with green for pass and red for fail. This loop goes through columns.
    For i = maxCols + 2 To maxCols + 5
        
        worstMargin = Cells(2, i).Value
        worstFrequency = Cells(2, 1).Value
        
        'This loop goes through rows.
        For j = 2 To (Cells(Rows.Count, maxCols).End(xlUp).Row)
        
            Set cell = Cells(j, i)
            
            If (cell.Value < worstMargin) Then
                worstMargin = cell.Value
                worstFrequency = Cells(j, 1).Value
            End If
            
            If (cell.Value > 0) Then
                cell.Style = "Good"
            Else
                cell.Style = "Bad"
            End If
        Next
        
        'Still in the for i loop, but outside for j loop. These Ifs put values of worst margins into cells.
        'They are formatted to 2 decimal places without rounding.
        'Checks if value is greater than 0 to mark it with green for pass and red for fail.
        If (i = maxCols + 2) Then
            Cells(5, i + 5).Formula = "=TRUNC(" + Replace(worstFrequency, ",", ".") + ",2)"
            Cells(5, i + 6).Formula = "=TRUNC(" + Replace(worstMargin, ",", ".") + ",2)"
            
            If (worstMargin < 0) Then
                Cells(5, i + 6).Style = "Bad"
            Else
                Cells(5, i + 6).Style = "Good"
            End If
        
        ElseIf (i = maxCols + 3) Then
            Cells(5, i + 6).Formula = "=TRUNC(" + Replace(worstFrequency, ",", ".") + ",2)"
            Cells(5, i + 7).Formula = "=TRUNC(" + Replace(worstMargin, ",", ".") + ",2)"
            
            If (worstMargin < 0) Then
                Cells(5, i + 7).Style = "Bad"
            Else
                Cells(5, i + 7).Style = "Good"
            End If
         
        ElseIf (i = maxCols + 4) Then
            Cells(5, i + 7).Formula = "=TRUNC(" + Replace(worstFrequency, ",", ".") + ",2)"
            Cells(5, i + 8).Formula = "=TRUNC(" + Replace(worstMargin, ",", ".") + ",2)"
            
            If (worstMargin < 0) Then
                Cells(5, i + 8).Style = "Bad"
            Else
                Cells(5, i + 8).Style = "Good"
            End If
            
        ElseIf (i = maxCols + 5) Then
            Cells(5, i + 8).Formula = "=TRUNC(" + Replace(worstFrequency, ",", ".") + ",2)"
            Cells(5, i + 9).Formula = "=TRUNC(" + Replace(worstMargin, ",", ".") + ",2)"
            
            If (worstMargin < 0) Then
                Cells(5, i + 9).Style = "Bad"
            Else
                Cells(5, i + 9).Style = "Good"
            End If
        End If
    Next
    
    'Creates borders inside and around WORST MARGINS to make it look like a table.
    Range(Cells(2, maxCols + 7), Cells(5, maxCols + 14)).Borders.LineStyle = XlLineStyle.xlContinuous
    
End Sub
