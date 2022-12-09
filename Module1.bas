Attribute VB_Name = "Module1"
'   MAIN FUNCTION
' -------------------
'
' First you need to save csv limits to xlsx
'
'   LINES TO CHANGE
'
'       Module1:
'   - Sub openAllWorkbooks
'       -> line 57 change specific cable testlog's directory to your own
'
'       Module2:
'   - Sub conversion
'       -> line 16 change output xlsx files folder
'
'       Module4:
'   - Sub limitsAdd
'       -> lines 29-31 change limit files' paths to your own
'
'       Module5:
'   - Sub dataFormat
'       -> line
'

Sub openAllWorkbooks()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False

' openAllWorkbooks Macro
    
    Dim vFiles As Variant
    Dim vFile As Variant
    'Variables that store temporary names of files
    
    Dim testDirectory As String
    'Folder path where are keysight measurements
    
    Dim measurementFileName As String
    'Name of converted filename from conversion Sub
    
    Dim measurementNumber As Integer
    'Argument for limitsAdd Sub
    
    Dim src As Workbook
    'For workbooks creation without opening them
    
    Dim limitSrc As Workbook
    'For limit workbooks creation without opening them
    
    Dim buff As Integer
    Dim sameLimit As Boolean
    Dim count As Integer
    'Assigns current file with specified type of measurement, has information if it is file with same type of measurement, counts files
    
        
        
    testDirectory = "C:\Users\mateup3\OneDrive - kochind.com\Documents\2. FLUKE TESTS\ORIENT DOUBLE JACKET C6 U_UTP\100m test\"
    'Change directory for different cables tests
    
    vFiles = enumerateFiles(testDirectory, "csv")
    'It calls a public function required to go through all subfolders to files with csv extension
    

    buff = 0
    sameLimit = True
    count = 0
    'Start values
    
    For Each vFile In vFiles
        
        If (InStr(Len(testDirectory), vFile, "il")) <> 0 Then

            measurementNumber = 1
            Set src = Workbooks.Open(vFile, True, True)
            'Opens the source excel workbook in "read only mode"
            'It works faster than opening workbooks
        
        ElseIf (InStr(Len(testDirectory), vFile, "next")) <> 0 Then

            measurementNumber = 2
            Set src = Workbooks.Open(vFile, True, True)
        
        ElseIf (InStr(Len(testDirectory), vFile, "rl")) <> 0 Then

            measurementNumber = 3
            Set src = Workbooks.Open(vFile, True, True)
            
        End If
        'Checks if csv files are ones we want to actually open and operate on them
        
        If (count <> 0) Then
            If (buff <> measurementNumber) Then
                sameLimit = False
            Else
                buff = measurementNumber
                sameLimit = True
            End If
        End If
        'If it is first file, ommit it, if not check if previous file had same type of measurement
        
        Call Module2.conversion(CStr(vFile), measurementFileName)
        'Converts file from csv to xlsx
        
        If (measurementNumber = 1 Or measurementNumber = 2 Or measurementNumber = 3) Then
            
            count = count + 1
            'Counts how many files where opened
            
            Call Module3.deleteRedundantData(src, measurementFileName)
            Call Module4.limitsAdd(measurementNumber, measurementFileName, sameLimit, count, limitSrc)
            Call Module5.dataFormat(src)
        
        End If
        'Calling Subs if found proper csv with measurements
        
    Next vFile
    'Loops through all of csv files, assigning numbers to corresponding measurements
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub

' Don't change anything here

Public Function enumerateFiles(sDirectory As String, _
            Optional sFileSpec As String = "*", _
            Optional InclSubFolders As Boolean = True) As Variant

    enumerateFiles = Filter(Split(CreateObject("WScript.Shell").Exec _
        ("CMD /C DIR """ & sDirectory & "*." & sFileSpec & """ " & _
        IIf(InclSubFolders, "/S ", "") & "/B /A:-D").StdOut.ReadAll, vbCrLf), ".")

End Function




