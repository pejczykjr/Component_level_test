Attribute VB_Name = "Module4"
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
    
End Sub

Sub ilLimit()

'   Variables

    Dim ilFormula As String
    
'   FUNCTIONAL PART
    
    ilFormula = "=-(1.808*SQRT(A8)+0.017*A8+0.2/SQRT(A8))"
    
    With Range(Cells(8, maxCols + 1), Cells(Cells(Rows.count, 1).End(xlUp).Row - 1, maxCols + 1))
        .formula = ilFormula
        .FillDown
    End With
    'It fills the whole range with that formula, A8 value changes automatically for the next row, that's not absolute reference
        
End Sub

Sub nextLimit()

'   VARIABLES

    Dim nextFormula As String
    
'   FUNCTIONAL PART
    
    nextFormula = "=-(44.3-15*LOG10(A8/100))"
    
    With Range(Cells(8, maxCols + 1), Cells(Cells(Rows.count, 1).End(xlUp).Row - 1, maxCols + 1))
        .formula = nextFormula
        .FillDown
    End With
    
End Sub

Sub rlLimit()

'   VARIABLES

    Dim rlFormula As String
    
'   FUNCTIONAL PART
    
    rlFormula = "=-IF(AND(A8>=1,A8<10),20+5*LOG10(A8),IF(AND(A8>=10,A8<20), 25, 25-7*LOG10(A8/20)))"
    
    With Range(Cells(8, maxCols + 1), Cells(Cells(Rows.count, 1).End(xlUp).Row - 1, maxCols + 1))
        .formula = rlFormula
        .FillDown
    End With
    
End Sub
