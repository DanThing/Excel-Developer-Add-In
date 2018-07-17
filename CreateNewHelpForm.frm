VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateNewHelpForm 
   Caption         =   "Create / Edit New Help Option"
   ClientHeight    =   3900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6540
   OleObjectBlob   =   "CreateNewHelpForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateNewHelpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-----------------------------------
'- CreateNewHelpForm
'   - btn_Delete_Click
'   - btn_Save_Click
'   - Edit
'   - UserForm_QueryClose
'-----------------------------------
Option Explicit

Private indexID As Long

Private Sub btn_Delete_Click()
    If MsgBox("Do you want to delete this Help Option?", vbYesNo, Company) = vbNo Then
        Exit Sub
    End If
    If indexID = 1 Then
        MsgBox "Unable to Delete this option.", vbOKOnly + vbInformation, comapny
        Exit Sub
    End If
    With ThisWorkbook.Worksheets("HelpContents")
        .Cells(indexID + 1, 1).EntireRow.Delete
    End With
End Sub

Private Sub btn_Save_Click()
Dim lr As Long
    If IsEmpty(Me.TextBox1.Text) Or IsEmpty(Me.TextBox2.Text) Then
        MsgBox "Unable to save blank values.", vbOKOnly + vbExclamation, Company
        Exit Sub
    End If
    lr = Lastrow(ThisWorkbook.Worksheets("HelpContents")) + 1
    With ThisWorkbook.Worksheets("HelpContents")
        .Cells(indexID, 1).value = Me.TextBox1.Text
        .Cells(indexID, 2).value = Me.TextBox2.Text
    End With
    ThisWorkbook.Save
    Unload Me
End Sub

Public Sub Edit(listIndex As Long)
    indexID = listIndex
    If indexID > 1 Then
        Me.TextBox1.Text = .Cells(indexID, 1).value
        Me.TextBox2.Text = .Cells(indexID, 2).value
    Else
        Me.TextBox1.Text = vbNullString
        Me.TextBox2.Text = vbNullString
        indexID = Lastrow(ThisWorkbook.Worksheets("HelpContents")) + 1
    End If
    Me.Show modal
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Unload Me
End Sub

