Attribute VB_Name = "DataMoveToA1"
Option Explicit

'This sub moves data to the beginning of the sheet (A1 cell).
Sub moveData()

'   VARIABLES
    
    'Keeps address of "Freq(Hz)" header.
    Dim frequencyCellAddress As String
    'Keeps address of last used cell (e.g. last value/text in C30 cell).
    Dim lastCellAddress As String
    
'   FUNCTIONAL PART
    
    'Looks for "Freq(Hz)" title and assigns its address to variable.
    'Corrects title from "Freq(Hz)" to "Frequency [MHz]".
    With ActiveSheet.UsedRange.Find(What:="Freq(Hz)", LookIn:=xlValues, LookAt:=xlPart, _
    SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
            
        frequencyCellAddress = .Address(RowAbsolute:=False, ColumnAbsolute:=False)
        Range(frequencyCellAddress).Value = "Frequency [MHz]"
    
    End With

    'Assigns address of last used cell.
    lastCellAddress = ActiveSheet.Range("A1").SpecialCells(xlCellTypeLastCell).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    'Moves given range to A1 cell.
    Range(frequencyCellAddress + ":" + lastCellAddress).Cut Range("A1")
    
    'Looks for "END" title and deletes it.
    ActiveSheet.UsedRange.Find(What:="END", LookIn:=xlValues, LookAt:=xlPart, _
    SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, _
    SearchFormat:=False).ClearContents

End Sub
