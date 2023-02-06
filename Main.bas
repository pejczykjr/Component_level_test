Attribute VB_Name = "Main"
'It requires to always define variable.
Option Explicit

'Allows to choose proper limits.
Public testCategory As String

'Indicates if user chose to close program.
Public exitTrue As Boolean

'Keeps number of columns in use (index of last).
Public maxCols As Integer

'This is the main Sub, it searches for specified files and opens them.
'Later on it calls other Subs from different Modules.
Sub openAllWorkbooks()

'   VARIABLES
    
    'Variables that store temporary names of files that are searched.
    Dim vFiles, vFile As Variant
    'Folder path where are keysight measurements.
    Dim testDirectory As String
    'Folder path where output files will be saved.
    Dim folderPathXLSX As String
    'Gets value from CsvToXlsx.conversion().
    Dim measurementFileName As String
    'Keeps type of measurement (e.g. NEXT).
    Dim measurementType As String
    'Keeps active workbook as a variable.
    Dim src As Workbook
    
'   FUNCTIONAL PART

    'It does background staff.
    ApplicationOptimization (True)
    
    'Shows form with options to choose test category. If user closes form,
    'quits program.
    TestCategoryForm.Show
    If exitTrue = True Then: Exit Sub
    
    'Those assignments are needed for FoldersDialog.displayFolderDialog().
    testDirectory = "testDirectory"
    folderPathXLSX = "folderPathXLSX"
    
    'Calls FoldersDialog.displayFolderDialog() to choose folders' paths.
    Call FoldersDialog.displayFolderDialog(testDirectory)
    If exitTrue = True Then: Exit Sub
    Call FoldersDialog.displayFolderDialog(folderPathXLSX)
    If exitTrue = True Then: Exit Sub
    
    'It assigns a function that goes through all subfolders to files with csv extension.
    vFiles = enumerateFiles(testDirectory, "csv")
        
'If any error appears, go to errorMsg.
'On Error GoTo errorMsg
    
    'Loops through all of csv files.
    For Each vFile In vFiles
        
        'Ifs check if csv files contain il/next/rl in name.
        If (InStr(Len(testDirectory), LCase(vFile), "il")) <> 0 Then
            measurementType = "il"
        
        ElseIf (InStr(Len(testDirectory), LCase(vFile), "next")) <> 0 Then
            measurementType = "next"
        
        ElseIf (InStr(Len(testDirectory), LCase(vFile), "rl")) <> 0 Then
            measurementType = "rl"
        End If
        
        'Calling Subs if found proper file with measurements.
        If (measurementType = "il" Or measurementType = "next" Or measurementType = "rl") Then
        
            'Opens the source excel workbook in "read only mode" ("it works faster than opening workbooks").
            Set src = Workbooks.Open(vFile, True, True)
            
            Call CsvToXlsx.conversion(measurementFileName, folderPathXLSX)
            Call DataMoveToA1.moveData
            Call ClearData.deleteRedundantData(measurementFileName, measurementType)
            Call LimitsMargins.limitsAdd(measurementType, measurementFileName)
            Call Charts.dataFormat(src, measurementFileName)
        
        End If
        
    Next vFile
    
    'Turn off optimization.
    ApplicationOptimization (False)
    
    'Indicates that program finished its job and there was no error, then exits.
    MsgBox "Program finished with success.", vbOKOnly
    Exit Sub
    
errorMsg:
    'Indicates that program finished with error and displays message.
    MsgBox "Program finished with Error!" & vbNewLine & Err.Description

End Sub

'Function that searches through folders and subfolders.
Public Function enumerateFiles(sDirectory As String, _
            Optional sFileSpec As String = "*", _
            Optional InclSubFolders As Boolean = True) As Variant

    enumerateFiles = Filter(Split(CreateObject("WScript.Shell").Exec _
        ("CMD /C DIR """ & sDirectory & "*." & sFileSpec & """ " & _
        IIf(InclSubFolders, "/S ", "") & "/B /A:-D").StdOut.ReadAll, vbCrLf), ".")
End Function

'Function that speeds up code.
Public Function ApplicationOptimization(BOOLEAN_TRIGGER As Boolean)
    With Application
        Select Case BOOLEAN_TRIGGER
            Case True
                .ScreenUpdating = False
                .DisplayStatusBar = False
                .EnableEvents = False
                .Calculation = xlCalculationManual
            Case False
                .ScreenUpdating = True
                .DisplayStatusBar = True
                .EnableEvents = True
                .Calculation = xlCalculationAutomatic
        End Select
    End With
    If (BOOLEAN_TRIGGER = True) Then
        ActiveSheet.DisplayPageBreaks = False
    ElseIf (BOOLEAN_TRIGGER = False) Then
        ActiveSheet.DisplayPageBreaks = True
    End If
End Function
