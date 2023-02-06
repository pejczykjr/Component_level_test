VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TestCategoryForm 
   Caption         =   "TEST CATEGORY"
   ClientHeight    =   1420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5025
   OleObjectBlob   =   "TestCategoryForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TestCategoryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cat5eButton_Click()
    Dim iResponse As Integer
    
    iResponse = MsgBox("Do you want to set test limits to CAT5E?", vbQuestion + vbYesNo)
    
    Select Case iResponse
        Case vbYes
            testCategory = "CAT5E"
            exitTrue = False
            Unload Me
        Case vbNo: Exit Sub
    End Select

End Sub

Private Sub cat6Button_Click()
    Dim iResponse As Integer
    
    iResponse = MsgBox("Do you want to set test limits to CAT6?", vbQuestion + vbYesNo)
    
    Select Case iResponse
        Case vbYes
            testCategory = "CAT6"
            exitTrue = False
            Unload Me
        Case vbNo: Exit Sub
    End Select
    
End Sub

Private Sub cat6aButton_Click()
    Dim iResponse As Integer
    
    iResponse = MsgBox("Do you want to set test limits to CAT6A", vbQuestion + vbYesNo)
    
    Select Case iResponse
        Case vbYes
            testCategory = "CAT6A"
            exitTrue = False
            Unload Me
        Case vbNo: Exit Sub
    End Select
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        If MsgBox("Are you sure you want to exit program?", vbQuestion + vbYesNo, "Ready to Exit?") = vbNo Then
            Cancel = True
        Else
            exitTrue = True
            Unload Me
            Exit Sub
        End If
    End If
    
End Sub

