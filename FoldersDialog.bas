Attribute VB_Name = "FoldersDialog"
Option Explicit

'This Sub opens folder dialog and allows user to choose folders manually.
Sub displayFolderDialog(ByRef directory As String)

'   VARIABLES

    'Indicates if user chose folder.
    Dim notChosen As Boolean
    'It keeps user's response from MsgBox.
    Dim button As Integer

'   FUNCTIONAL PART
    
    'Declares that the user hasn't chosen any folders yet.
    notChosen = True
    
    'This loop checks user's response. When all the required directories are picked, program starts.
    Do While notChosen = True
    
        'Opens dialog for user to choose folder.
        With Application.FileDialog(msoFileDialogFolderPicker)
            
            'Setting title of dialog according to current directory(for user to know what he sees).
            If directory = "testDirectory" Then
                .Title = "Select a folder containing keysight measurements"
            ElseIf directory = "folderPathXLSX" Then
                .Title = "Select a folder to save your output files"
            End If
        
            'If OK is pressed, assigns path to testDirectory.
            If .Show = -1 Then
                notChosen = False
                directory = .SelectedItems(1)
            Else
                button = MsgBox("You didn't choose a folder path!" & vbNewLine & "Do you want to try again?", vbYesNo Or vbExclamation, "Wrong path")
                'Allows user to choose folder again.
                If button = vbYes Then
                    notChosen = True
                'Exits this function and program.
                Else
                    exitTrue = True
                    notChosen = False
                    Exit Sub
                End If
            
            End If
        End With
    Loop
    
    'Adds slash at the end of a directory in case user doesn't. Searching function requires it.
    If Right(directory, 1) <> "\" Then
        directory = directory + "\"
    End If
    
End Sub
