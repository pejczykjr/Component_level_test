Attribute VB_Name = "Module4"
'NA SZTYWNO USTAWIONE KOMORKI!!!!!!!!!!!!!!!!!!!

'TA FUNKCJA Z 1,1 DAJE "A1", ULATWIENIE DLA RANGE
'Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)

Option Explicit

'This Sub adds limits from formula to the spreadsheet
Sub limitsAdd()

'   VARIABLES
    
    Dim cell As Range
    'Sets range to update values of frequency
    
'   FUNCTIONAL PART
    
    Cells(7, 1).Value = "Frequency(MHz)"
    'Correct title from Frequency(Hz) to Frequency(MHz)
    
    For Each cell In Range("A8:A" + CStr(Cells(Rows.count, 1).End(xlUp).Row - 1))
        cell.Value = cell.Value / 1000000
    Next
    'Divide values to have MHz units
    
    Cells(7, maxCols + 1).Value = "Limit(DB)"
    'Set title in the new column
    
    Select Case measurementType
        Case "il": Call ilLimit
        Case "next": Call nextLimit
        Case "rl": Call rlLimit
    End Select
    'Assigned limit according to measurementType
    
    'Call limit(measurementType)
    
    
End Sub

Sub limit(limitTpype As String)

'   VARIABLES
    
    Dim limitFormula, ilFormula, nextFormula, rlFormula As String
    
    ilFormula = "=-(1.808*SQRT(A8)+0.017*A8+0.2/SQRT(A8))"
    nextFormula = "=-(44.3-15*LOG10(A8/100))"
    rlFormula = "=-IF(AND(A8>=1,A8<10),20+5*LOG10(A8),IF(AND(A8>=10,A8<20), 25, 25-7*LOG10(A8/20)))"
    
'   FUNCTIONAL PART
    
    Select Case measurementType
        Case "il": limitFormula = ilFormula
        Case "next": limitFormula = nextFormula
        Case "rl": limitFormula = rlFormula
    End Select
    'Assigned limit according to measurementType

End Sub




















Sub ilLimit()

'   Variables

    Dim ilFormula As String, cell As Range, worstMargin, worstFrequency, i As Integer
    
'   FUNCTIONAL PART

    ilFormula = "=(1.808*SQRT(A8)+0.017*A8+0.2/SQRT(A8))"
    Range(Cells(8, maxCols + 1), Cells(Cells(Rows.count, 1).End(xlUp).Row - 1, maxCols + 1)).formula = ilFormula
    'It fills the whole range with that formula, A8 value changes automatically for the next row, no reference
    
   
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'POPRAWIC: DO AUTOMATYCZNEGO WYKRWYANIA KOMOREK DLA KONKRETNYCH WARTOSCI
    
    Cells(6, maxCols + 2).Value = "MARGINS"
    Cells(7, maxCols + 2).Value = "S21"
    
    Range(Cells(6, maxCols + 3), Cells(6, maxCols + 4)).Merge
    Cells(6, maxCols + 3).Value = "WORST MARGIN"
    Range(Cells(7, maxCols + 3), Cells(7, maxCols + 4)).Merge
    Cells(7, maxCols + 3).Value = "S21"
    
    Cells(8, maxCols + 3).Value = "Frequency"
    Cells(8, maxCols + 4).Value = "Value"
    

    Range(Cells(8, maxCols + 2), Cells(Cells(Rows.count, 1).End(xlUp).Row - 1, maxCols + 2)) = "=B8-C8"
    
    worstMargin = Cells(8, maxCols + 2)
    worstFrequency = Cells(8, 1)
    'First assignment, temporary worst value
    
    For i = 8 To (Cells(Rows.count, 1).End(xlUp).Row - 1)
        Set cell = Cells(i, maxCols + 2)
        
        If (cell.Value > worstMargin) Then
            worstMargin = cell.Value
            worstFrequency = Cells(i, 1)
        End If
        
        
        If (cell.Value < 0) Then
            cell.Style = "Good"
        Else
            cell.Style = "Bad"
        End If
    Next
    
    Cells(9, maxCols + 3).Value = worstMargin
    Cells(9, maxCols + 4).Value = worstFrequency
    
    If (worstMargin > 0) Then
        Cells(9, maxCols + 4).Style = "Bad"
    Else
        Cells(9, maxCols + 4).Style = "Good"
    End If
       
       'TU MOZNA SKOPIOWAC PO PROSTU WARTOSCI
       
End Sub

Sub nextLimit()

'   VARIABLES

    Dim nextFormula As String, worstMargin, worstFrequency, i, j As Integer, cell As Range
    
'   FUNCTIONAL PART
    
    nextFormula = "=-(44.3-15*LOG10(A8/100))"
    Range(Cells(8, maxCols + 1), Cells(Cells(Rows.count, 1).End(xlUp).Row - 1, maxCols + 1)).formula = nextFormula
    

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'POPRAWIC: DO AUTOMATYCZNEGO WYKRWYANIA KOMOREK DLA KONKRETNYCH WARTOSCI

    Range(Cells(6, maxCols + 2), Cells(6, maxCols + 4)).Merge
    Cells(6, maxCols + 2).Value = "MARGINS"
    Cells(7, maxCols + 2).Value = "SXY"
    Cells(7, maxCols + 3).Value = "SXY"
    Cells(7, maxCols + 4).Value = "SXY"
    
    Range(Cells(6, maxCols + 5), Cells(6, maxCols + 10)).Merge
    Cells(6, maxCols + 5).Value = "WORST MARGINS"
    Range(Cells(7, maxCols + 5), Cells(7, maxCols + 6)).Merge
    Cells(7, maxCols + 5).Value = "SXY"
    Range(Cells(7, maxCols + 7), Cells(7, maxCols + 8)).Merge
    Cells(7, maxCols + 7).Value = "SXY"
    Range(Cells(7, maxCols + 9), Cells(7, maxCols + 10)).Merge
    Cells(7, maxCols + 9).Value = "SXY"
    
    Cells(8, maxCols + 5).Value = "Frequency"
    Cells(8, maxCols + 6).Value = "Value"
    Cells(8, maxCols + 7).Value = "Frequency"
    Cells(8, maxCols + 8).Value = "Value"
    Cells(8, maxCols + 9).Value = "Frequency"
    Cells(8, maxCols + 10).Value = "Value"
    
    Range(Cells(8, maxCols + 2), Cells(Cells(Rows.count, 1).End(xlUp).Row - 1, maxCols + 2)) = "=B8-E8"
    Range(Cells(8, maxCols + 3), Cells(Cells(Rows.count, 1).End(xlUp).Row - 1, maxCols + 3)) = "=C8-E8"
    Range(Cells(8, maxCols + 4), Cells(Cells(Rows.count, 1).End(xlUp).Row - 1, maxCols + 4)) = "=D8-E8"
    
    
    'upto 10MHZ OMMIT !!!!!!!!!!!!!!!!!!!!!
    For i = maxCols + 2 To maxCols + 4
        
        worstMargin = -1000
        worstFrequency = 0
        
        For j = 8 To (Cells(Rows.count, 1).End(xlUp).Row - 1)
            If (Cells(j, 1).Value >= 10) Then
                Set cell = Cells(j, i)
            
                If (cell.Value > worstMargin) Then
                    worstMargin = cell.Value
                    worstFrequency = Cells(j, 1).Value
                End If
                    
                If (cell.Value < 0) Then
                    cell.Style = "Good"
                Else
                    cell.Style = "Bad"
                End If
            Else
            End If
        Next
        
        If (i = maxCols + 2) Then
            Cells(9, i + 3).Value = worstFrequency
            Cells(9, i + 4).Value = worstMargin
            
            If (worstMargin > 0) Then
                Cells(9, i + 4).Style = "Bad"
            Else
                Cells(9, i + 4).Style = "Good"
            End If
        
        ElseIf (i = maxCols + 3) Then
            Cells(9, i + 4).Value = worstFrequency
            Cells(9, i + 5).Value = worstMargin
            
            If (worstMargin > 0) Then
                Cells(9, i + 5).Style = "Bad"
            Else
                Cells(9, i + 5).Style = "Good"
            End If
         
        ElseIf (i = maxCols + 4) Then
            Cells(9, i + 5).Value = worstFrequency
            Cells(9, i + 6).Value = worstMargin
            
            If (worstMargin > 0) Then
                Cells(9, i + 6).Style = "Bad"
            Else
                Cells(9, i + 6).Style = "Good"
            End If
        End If
    Next

End Sub

Sub rlLimit()

'   VARIABLES

    Dim rlFormula As String, worstMargin, worstFrequency, i, j As Integer, cell As Range
    
'   FUNCTIONAL PART
    
    rlFormula = "=-IF(AND(A8>=1,A8<10),20+5*LOG10(A8),IF(AND(A8>=10,A8<20), 25, 25-7*LOG10(A8/20)))"
    Range(Cells(8, maxCols + 1), Cells(Cells(Rows.count, 1).End(xlUp).Row - 1, maxCols + 1)) = rlFormula
    
    
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'POPRAWIC: DO AUTOMATYCZNEGO WYKRWYANIA KOMOREK DLA KONKRETNYCH WARTOSCI

    Range(Cells(6, maxCols + 2), Cells(6, maxCols + 5)).Merge
    Cells(6, maxCols + 2).Value = "MARGINS"
    
    Cells(7, maxCols + 2).Value = "S11"
    Cells(7, maxCols + 3).Value = "S22"
    Cells(7, maxCols + 4).Value = "S33"
    Cells(7, maxCols + 5).Value = "S44"
    
    Range(Cells(6, maxCols + 6), Cells(6, maxCols + 13)).Merge
    Cells(6, maxCols + 6).Value = "WORST MARGINS"
    Range(Cells(7, maxCols + 6), Cells(7, maxCols + 7)).Merge
    Cells(7, maxCols + 6).Value = "S11"
    Range(Cells(7, maxCols + 8), Cells(7, maxCols + 9)).Merge
    Cells(7, maxCols + 8).Value = "S22"
    Range(Cells(7, maxCols + 10), Cells(7, maxCols + 11)).Merge
    Cells(7, maxCols + 10).Value = "S33"
    Range(Cells(7, maxCols + 12), Cells(7, maxCols + 13)).Merge
    Cells(7, maxCols + 12).Value = "S44"
    
    Cells(8, maxCols + 6).Value = "Frequency"
    Cells(8, maxCols + 7).Value = "Value"
    Cells(8, maxCols + 8).Value = "Frequency"
    Cells(8, maxCols + 9).Value = "Value"
    Cells(8, maxCols + 10).Value = "Frequency"
    Cells(8, maxCols + 11).Value = "Value"
    Cells(8, maxCols + 12).Value = "Frequency"
    Cells(8, maxCols + 13).Value = "Value"
        
    Range(Cells(8, maxCols + 2), Cells(Cells(Rows.count, 1).End(xlUp).Row - 1, maxCols + 2)) = "=B8-F8"
    Range(Cells(8, maxCols + 3), Cells(Cells(Rows.count, 1).End(xlUp).Row - 1, maxCols + 3)) = "=C8-F8"
    Range(Cells(8, maxCols + 4), Cells(Cells(Rows.count, 1).End(xlUp).Row - 1, maxCols + 4)) = "=D8-F8"
    Range(Cells(8, maxCols + 5), Cells(Cells(Rows.count, 1).End(xlUp).Row - 1, maxCols + 5)) = "=E8-F8"
    
    
    For i = maxCols + 2 To maxCols + 5
        
        worstMargin = Cells(8, i).Value
        worstFrequency = Cells(8, 1).Value
        
        For j = 8 To (Cells(Rows.count, 1).End(xlUp).Row - 1)
            Set cell = Cells(j, i)
            
            If (cell.Value > worstMargin) Then
                worstMargin = cell.Value
                worstFrequency = Cells(j, 1).Value
            End If
            
            If (cell.Value < 0) Then
                cell.Style = "Good"
            Else
                cell.Style = "Bad"
            End If
        Next
        
        If (i = maxCols + 2) Then
            Cells(9, i + 4).Value = worstFrequency
            Cells(9, i + 5).Value = worstMargin
            
            If (worstMargin > 0) Then
                Cells(9, i + 5).Style = "Bad"
            Else
                Cells(9, i + 5).Style = "Good"
            End If
        
        ElseIf (i = maxCols + 3) Then
            Cells(9, i + 5).Value = worstFrequency
            Cells(9, i + 6).Value = worstMargin
            
            If (worstMargin > 0) Then
                Cells(9, i + 6).Style = "Bad"
            Else
                Cells(9, i + 6).Style = "Good"
            End If
         
        ElseIf (i = maxCols + 4) Then
            Cells(9, i + 6).Value = worstFrequency
            Cells(9, i + 7).Value = worstMargin
            
            If (worstMargin > 0) Then
                Cells(9, i + 7).Style = "Bad"
            Else
                Cells(9, i + 7).Style = "Good"
            End If
            
        ElseIf (i = maxCols + 5) Then
            Cells(9, i + 7).Value = worstFrequency
            Cells(9, i + 8).Value = worstMargin
            
            If (worstMargin > 0) Then
                Cells(9, i + 8).Style = "Bad"
            Else
                Cells(9, i + 8).Style = "Good"
            End If
        End If
    Next
    
End Sub
