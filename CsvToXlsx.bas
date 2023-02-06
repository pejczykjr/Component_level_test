Attribute VB_Name = "CsvToXlsx"
Option Explicit

'This Sub converts file from csv to xlsx.
Sub conversion(ByRef measurementFileName As String, folderPathXLSX As String)

'   VARIABLES

    'Contains path and file name with xlsx extension.
    Dim filePathXLSX As String
    'Name of current file with extension.
    Dim fileNameExt As String
    
'   FUNCTIONAL PART
    
    'Gets name of active workbook.
    fileNameExt = LCase(ActiveWorkbook.Name)
    
    'Sets new path file path.
    filePathXLSX = folderPathXLSX + Replace(fileNameExt, "csv", "xlsx")
    
    'Bare file name without extension.
    measurementFileName = Replace(fileNameExt, ".csv", "")
    
    'It saves workbook as xlsx.
    ActiveWorkbook.SaveAs fileName:=filePathXLSX, _
    FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

End Sub
