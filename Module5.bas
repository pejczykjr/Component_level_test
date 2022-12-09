Attribute VB_Name = "Module5"
Option Explicit

'This Sub formats data of finished file and makes it look better
Sub dataFormat(ByRef src As Workbook)

'   FUNCTIONAL PART

    With ActiveSheet.UsedRange
        .EntireRow.AutoFit
        .EntireColumn.AutoFit
        .HorizontalAlignment = xlCenter
    End With
    'Adjusts height and weight of all active cells in worksheet

    src.Close True
    'Closes the source file
    'True - saves the source file
    Set src = Nothing

End Sub
