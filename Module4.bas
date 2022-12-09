Attribute VB_Name = "Module4"
Sub limitsAdd(limitNumber As Integer, measurementFileName As String, sameLimit As Boolean, count As Integer, ByRef limitSrc As Workbook)

' limitsAdd Macro

    Dim rlLimitPath As String
    Dim ilLimitPath As String
    Dim nextLimitPath As String
    'Paths where limits are saved
    
    Dim limitFileName As String
    'Limit file name without extension
    
    Dim currLimitPath As String
    'Assigned one of above limits depending on measurements
    
    Dim maxCols As Integer
    'Checks how many columns are used
    
    
    
    If (sameLimit = False) Then
        limitSrc.Close False
        Set limitSrc = Nothing
    End If
    'It closes file with limits and doesn't save it when file with measurements changed to different type
        
    maxCols = ActiveSheet.UsedRange.Columns.count
    
    ilLimitPath = "C:\Users\mateup3\OneDrive - kochind.com\Documents\3. 100m test limits\C6\Insertion Loss\Insertion Loss Limit C6.xlsx"
    nextLimitPath = "C:\Users\mateup3\OneDrive - kochind.com\Documents\3. 100m test limits\C6\NEXT\NEXT_LIMIT_C6.xlsx"
    rlLimitPath = "C:\Users\mateup3\OneDrive - kochind.com\Documents\3. 100m test limits\C6\Return Loss\Return Loss Limit C6.xlsx"
    'Change only when limits files' directory changes
    
    If limitNumber = 1 Then
        currLimitPath = ilLimitPath
    ElseIf limitNumber = 2 Then
        currLimitPath = nextLimitPath
    ElseIf limitNumber = 3 Then
        currLimitPath = rlLimitPath
    End If
    'Assigned limit according to measurement, limitNumber is given from Module2.openAllWorkbooks

    limitFileName = Replace(Right(currLimitPath, Len(currLimitPath) - InStrRev(currLimitPath, "\")), ".xlsx", "")
    'It cuts string from right side till meets "\" sign,
    'counts from left side so it is required to deduct length of currLimitPath
    'and then replaces extension with nothing to have clear file name
    
     On Error GoTo ErrHandler
    
        If (sameLimit = False Or count = 1) Then
            Set limitSrc = Workbooks.Open(currLimitPath, True, True)
        End If
        'It opens the source excel workbook in "read only mode" with limits when there is a file with new type of measurement
        
        Workbooks(limitFileName).Worksheets(limitFileName).Range("A1:E14").Copy _
        Workbooks(measurementFileName).Worksheets(measurementFileName).Cells(5, maxCols + 2)
        'Copy data from source to the destination workbook
        'Opens workbook with limits and copy values to workbook with measurements to the
        'second empty column
        
        If (count = 10) Then
            limitSrc.Close False
            Set limitSrc = Nothing
        End If
        'If it is last file with measurements (count = 10) then close workbook with limits and don't save it
    
ErrHandler:
'Shows if any error occurs

End Sub







    

