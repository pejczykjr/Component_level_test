Attribute VB_Name = "Module3"
Option Explicit

'This Sub deletes excess of unneeded data from files with measurements
Sub deleteRedundantData()

'   VARIABLES
    
    Dim i As Integer
    'Iterator in Subs under deleteRedundantData

'   FUNCTIONAL PART
    
    maxCols = ActiveSheet.UsedRange.Columns.count
    
    If (InStr(1, measurementFileName, "il")) <> 0 Then
    
        Call Module3.ilAll(i)
        
    ElseIf (InStr(1, measurementFileName, "next")) <> 0 Then
    
        If InStr(1, measurementFileName, "orange") <> 0 Then
            Call Module3.nextOrange(i)
        ElseIf InStr(1, measurementFileName, "brown") <> 0 Then
            Call Module3.nextBrown(i)
        ElseIf InStr(1, measurementFileName, "green") <> 0 Then
            Call Module3.nextGreen(i)
        ElseIf InStr(1, measurementFileName, "blue") <> 0 Then
            Call Module3.nextBlue(i)
        End If
        
    ElseIf (InStr(1, measurementFileName, "rl")) <> 0 Then
    
        Call Module3.rlAll(i)
        
    End If
    'Function where it chooses type of measurement and erases specified redundant data
    'Subs under are checking if column title is what it needs and if not, it deletes the whole column and moves to left side

    maxCols = ActiveSheet.UsedRange.Columns.count
    
End Sub

Sub ilAll(i As Integer)
    
'   FUNCTIONAL PART
    
    For i = maxCols To 2 Step -1
        If ActiveSheet.Cells(7, i).Value <> "S21(DB)" Then ActiveSheet.Cells(7, i).EntireColumn.Delete Shift:=xlToLeft
    Next
    'From last column go to second and check conditions

End Sub

Sub nextOrange(i As Integer)

'   FUNCTIONAL PART

    For i = maxCols To 2 Step -1
        If (ActiveSheet.Cells(7, i).Value <> "S21(DB)" And ActiveSheet.Cells(7, i).Value <> "S23(DB)" _
        And ActiveSheet.Cells(7, i).Value <> "S24(DB)") Then ActiveSheet.Cells(7, i).EntireColumn.Delete Shift:=xlToLeft
    Next

End Sub

Sub nextBlue(i As Integer)
    
'   FUNCTIONAL PART

    For i = maxCols To 2 Step -1
        If (ActiveSheet.Cells(7, i).Value <> "S12(DB)" And ActiveSheet.Cells(7, i).Value <> "S13(DB)" _
        And ActiveSheet.Cells(7, i).Value <> "S14(DB)") Then ActiveSheet.Cells(7, i).EntireColumn.Delete Shift:=xlToLeft
    Next

End Sub

Sub nextGreen(i As Integer)

'   FUNCTIONAL PART
    
    For i = maxCols To 2 Step -1
        If (ActiveSheet.Cells(7, i).Value <> "S31(DB)" And ActiveSheet.Cells(7, i).Value <> "S32(DB)" _
        And ActiveSheet.Cells(7, i).Value <> "S34(DB)") Then ActiveSheet.Cells(7, i).EntireColumn.Delete Shift:=xlToLeft
    Next

End Sub

Sub nextBrown(i As Integer)

'   FUNCTIONAL PART
    
    For i = maxCols To 2 Step -1
        If (ActiveSheet.Cells(7, i).Value <> "S41(DB)" And ActiveSheet.Cells(7, i).Value <> "S42(DB)" _
        And ActiveSheet.Cells(7, i).Value <> "S43(DB)") Then ActiveSheet.Cells(7, i).EntireColumn.Delete Shift:=xlToLeft
    Next

End Sub

Sub rlAll(i As Integer)
    
'   FUNCTIONAL PART
    
    For i = maxCols To 2 Step -1
        If (ActiveSheet.Cells(7, i).Value <> "S11(DB)" And ActiveSheet.Cells(7, i).Value <> "S22(DB)" _
        And ActiveSheet.Cells(7, i).Value <> "S33(DB)" And ActiveSheet.Cells(7, i).Value <> "S44(DB)") _
        Then ActiveSheet.Cells(7, i).EntireColumn.Delete Shift:=xlToLeft
    Next

End Sub
