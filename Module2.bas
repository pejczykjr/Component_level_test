Attribute VB_Name = "Module2"
Sub conversion(filePathCSV As String, ByRef measurementFileName As String)

' conversion Macro

    Dim folderPathXLSX As String
    'Path where output files will be saved
    
    Dim filePathXLSX As String
    'Name of file path saved as xlsx

    Dim fileNameExt As String
    'Name of current file with extension
    
    
    
    folderPathXLSX = "C:\Users\mateup3\OneDrive - kochind.com\Documents\2. FLUKE TESTS\ORIENT DOUBLE JACKET C6 U_UTP\100m test\!MADE EXCEL WORKSHEETS\"
    'Change folder path depending where you want to save xlsx output files
    
    fileNameExt = ActiveWorkbook.Name
    filePathXLSX = folderPathXLSX + Replace(fileNameExt, "csv", "xlsx")
    
    ActiveWorkbook.SaveAs fileName:=filePathXLSX, _
    FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    'It saves workbook as xlsx

    measurementFileName = Replace(fileNameExt, ".csv", "")

End Sub


