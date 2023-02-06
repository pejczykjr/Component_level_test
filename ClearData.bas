Attribute VB_Name = "ClearData"
Option Explicit

'This Sub deletes excess of unneeded data from files with measurements.
Sub deleteRedundantData(measurementFileName As String, measurementType As String)

'   FUNCTIONAL PART
    
    'Keeps number of last used column.
    maxCols = ActiveSheet.UsedRange.Columns.Count
    
    'Checks what type of measurement it is (checks what name contains) and then calls other subs.
    Select Case measurementType
        Case "il": Call ilAll
        Case "rl": Call rlAll
        Case "next":
        
        If InStr(1, measurementFileName, "orange") <> 0 Then
            Call nextOrange
            
        ElseIf InStr(1, measurementFileName, "brown") <> 0 Then
            Call nextBrown
                        
        ElseIf InStr(1, measurementFileName, "green") <> 0 Then
            Call nextGreen
                        
        ElseIf InStr(1, measurementFileName, "blue") <> 0 Then
            Call nextBlue
                        
        End If
    End Select
    
End Sub

'Subs below declare what data is redundant and delete it.
Sub ilAll()

'   VARIABLES
    
    'Iterator in for loop.
    Dim i As Integer
    'Range where values' sign needs to be changed. For il measurement only.
    Dim cell As Range
    
'   FUNCTIONAL PART
    
    'Loop starts from last column and goes to second one (in first is "Frequency [MHz]").
    'If header is not what is specified, it deletes whole column and moves rest of sheet left into that column.
    'Else it keeps column and corrects header from "(db)" to " [dB]".
    For i = maxCols To 2 Step -1
        If ActiveSheet.Cells(1, i).Value <> "S21(DB)" Then
            ActiveSheet.Cells(1, i).EntireColumn.delete Shift:=xlToLeft
        Else
            ActiveSheet.Cells(1, i).Value = Replace(ActiveSheet.Cells(1, i).Value, "(DB)", " [dB]", 1, 1, vbTextCompare)
        End If
            
    Next

    'Changes S21 values to positive (test was conducted based on positive limit).
    For Each cell In Range("B2:B" + CStr(Cells(Rows.Count, 2).End(xlUp).Row))
        cell.Value = cell.Value * -1
    Next

End Sub

Sub nextOrange()

'   VARIABLES

    'Iterator in for loop.
    Dim i As Integer
    
'   FUNCTIONAL PART

    'Loop starts from last column and goes to second one (in first is "Frequency [MHz]").
    'If header is not what is specified, it deletes whole column and moves rest of sheet left into that column.
    'Else it keeps column and corrects header from "(db)" to " [dB]".
    For i = maxCols To 2 Step -1
        If (ActiveSheet.Cells(1, i).Value <> "S21(DB)" And ActiveSheet.Cells(1, i).Value <> "S23(DB)" _
        And ActiveSheet.Cells(1, i).Value <> "S24(DB)") Then
            ActiveSheet.Cells(1, i).EntireColumn.delete Shift:=xlToLeft
        Else
            ActiveSheet.Cells(1, i).Value = Replace(ActiveSheet.Cells(1, i).Value, "(DB)", " [dB]", 1, 1, vbTextCompare)
        End If
    Next

End Sub

Sub nextBlue()

'   VARIABLES
    
    'Iterator in for loop.
    Dim i As Integer
    
'   FUNCTIONAL PART

    'Loop starts from last column and goes to second one (in first is "Frequency [MHz]").
    'If header is not what is specified, it deletes whole column and moves rest of sheet left into that column.
    'Else it keeps column and corrects header from "(db)" to " [dB]".
    For i = maxCols To 2 Step -1
        If (ActiveSheet.Cells(1, i).Value <> "S12(DB)" And ActiveSheet.Cells(1, i).Value <> "S13(DB)" _
        And ActiveSheet.Cells(1, i).Value <> "S14(DB)") Then
            ActiveSheet.Cells(1, i).EntireColumn.delete Shift:=xlToLeft
        Else
            ActiveSheet.Cells(1, i).Value = Replace(ActiveSheet.Cells(1, i).Value, "(DB)", " [dB]", 1, 1, vbTextCompare)
        End If
    Next

End Sub

Sub nextGreen()

'   VARIABLES

    'Iterator in for loop.
    Dim i As Integer
    
'   FUNCTIONAL PART
    
    'Loop starts from last column and goes to second one (in first is "Frequency [MHz]").
    'If header is not what is specified, it deletes whole column and moves rest of sheet left into that column.
    'Else it keeps column and corrects header from "(db)" to " [dB]".
    For i = maxCols To 2 Step -1
        If (ActiveSheet.Cells(1, i).Value <> "S31(DB)" And ActiveSheet.Cells(1, i).Value <> "S32(DB)" _
        And ActiveSheet.Cells(1, i).Value <> "S34(DB)") Then
            ActiveSheet.Cells(1, i).EntireColumn.delete Shift:=xlToLeft
        Else
            ActiveSheet.Cells(1, i).Value = Replace(ActiveSheet.Cells(1, i).Value, "(DB)", " [dB]", 1, 1, vbTextCompare)
        End If
    Next

End Sub

Sub nextBrown()

'   VARIABLES

    'Iterator in for loop.
    Dim i As Integer
    
'   FUNCTIONAL PART
    
    'Loop starts from last column and goes to second one (in first is "Frequency [MHz]").
    'If header is not what is specified, it deletes whole column and moves rest of sheet left into that column.
    'Else it keeps column and corrects header from "(db)" to " [dB]".
    For i = maxCols To 2 Step -1
        If (ActiveSheet.Cells(1, i).Value <> "S41(DB)" And ActiveSheet.Cells(1, i).Value <> "S42(DB)" _
        And ActiveSheet.Cells(1, i).Value <> "S43(DB)") Then
            ActiveSheet.Cells(1, i).EntireColumn.delete Shift:=xlToLeft
        Else
            ActiveSheet.Cells(1, i).Value = Replace(ActiveSheet.Cells(1, i).Value, "(DB)", " [dB]", 1, 1, vbTextCompare)
        End If
    Next

End Sub

Sub rlAll()
    
'   VARIABLES

    'Iterator in for loop.
    Dim i As Integer
    
'   FUNCTIONAL PART
    
    'Loop starts from last column and goes to second one (in first is "Frequency [MHz]").
    'If header is not what is specified, it deletes whole column and moves rest of sheet left into that column.
    'Else it keeps column and corrects header from "(db)" to " [dB]".
    For i = maxCols To 2 Step -1
        If (ActiveSheet.Cells(1, i).Value <> "S11(DB)" And ActiveSheet.Cells(1, i).Value <> "S22(DB)" _
        And ActiveSheet.Cells(1, i).Value <> "S33(DB)" And ActiveSheet.Cells(1, i).Value <> "S44(DB)") Then
            ActiveSheet.Cells(1, i).EntireColumn.delete Shift:=xlToLeft
        Else
            ActiveSheet.Cells(1, i).Value = Replace(ActiveSheet.Cells(1, i).Value, "(DB)", " [dB]", 1, 1, vbTextCompare)
        End If
    Next

End Sub


