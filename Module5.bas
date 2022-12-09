Attribute VB_Name = "Module5"
Sub dataFormat(ByRef src As Workbook)

' dataFormat Macro

    src.Activate
    ActiveSheet.UsedRange.EntireRow.AutoFit
    ActiveSheet.UsedRange.EntireColumn.AutoFit
    ActiveSheet.UsedRange.HorizontalAlignment = xlCenter
    'Adjusts height and weight of all active cells in worksheet

    src.Close True
    'Closes the source file
    'True - saves the source file
    Set src = Nothing

End Sub
