Attribute VB_Name = "Module1"
Option Explicit                         'It requires to always define variable

'   PUBLIC VARIABLES

Public measurementType As String        'Keeps type of measurement, assigned and used in here and used in Module4.limitsAdd
Public measurementFileName As String    'Gets value from Module2.conversion and is used in Module3.deleteRedundantData
Public maxCols As Integer                  'Counts used columns to find first empty, values assigned and used in Module3.deleteRedundantData
                                        'and used in Module4.limitsAdd

'This is the main Sub, it searches for specified files and opens them, later on it calls other Subs from different Modules
Sub openAllWorkbooks()

    ApplicationOptimization (True)

'   VARIABLES
    
    Dim vFiles, vFile As Variant        'Variables that store temporary names of files that are searched
    Dim testDirectory As String         'Folder path where are keysight measurements
    Dim src As Workbook
    

'   FUNCTIONAL PART
    
    testDirectory = "C:\Users\mateup3\OneDrive - kochind.com\Documents\2. FLUKE TESTS\ORIENT DOUBLE JACKET C6 U_UTP\100m test\"
    'Change directory for different cables tests
    
    vFiles = enumerateFiles(testDirectory, "csv")
    'It calls a  function required to go through all subfolders to files with csv extension
        
'If any error occurs go to errorMsg
On Error GoTo errorMsg
    
    For Each vFile In vFiles
        
        If (InStr(Len(testDirectory), vFile, "il")) <> 0 Then

            measurementType = "il"
            Set src = Workbooks.Open(vFile, True, True)
            'Opens the source excel workbook in "read only mode"
            'It works faster than opening workbooks
        
        ElseIf (InStr(Len(testDirectory), vFile, "next")) <> 0 Then

            measurementType = "next"
            Set src = Workbooks.Open(vFile, True, True)
        
        ElseIf (InStr(Len(testDirectory), vFile, "rl")) <> 0 Then

            measurementType = "rl"
            Set src = Workbooks.Open(vFile, True, True)
            
        End If
        'Checks if csv files are ones we want to actually open and operate on them
        
        Call Module2.conversion
        
        If (measurementType = "il" Or measurementType = "next" Or measurementType = "rl") Then
            Call Module3.deleteRedundantData
            Call Module4.limitsAdd
            Call Module5.dataFormat(src)
        End If
        'Calling Subs if found proper file with measurements
        
    Next vFile
    'Loops through all of csv files, assigning numbers to corresponding measurements
    
    ApplicationOptimization (False)
    
    MsgBox "Program finished with success.", vbOKOnly
    'Indicates that program finished its job and there was no error
    Exit Sub
    
errorMsg:
    MsgBox "Program finished with Error!" & vbNewLine & Err.Description
    'Indicates that program finished with error and displays message

End Sub

' Function that searches through folders and subfolders
Public Function enumerateFiles(sDirectory As String, _
            Optional sFileSpec As String = "*", _
            Optional InclSubFolders As Boolean = True) As Variant

    enumerateFiles = Filter(Split(CreateObject("WScript.Shell").Exec _
        ("CMD /C DIR """ & sDirectory & "*." & sFileSpec & """ " & _
        IIf(InclSubFolders, "/S ", "") & "/B /A:-D").StdOut.ReadAll, vbCrLf), ".")

End Function

' Function that speeds up code
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
