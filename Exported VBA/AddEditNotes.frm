VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddEditNotes 
   Caption         =   "Add/Edit Notes"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4200
   OleObjectBlob   =   "AddEditNotes.frx":0000
End
Attribute VB_Name = "AddEditNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Form: AddEditNotes                                     '
'                                                             '
' Notes are used to communicate additional information to     '
' drafters and engineer about certain items in the BOM. Notes '
' in this application manifest themselves as comments on the  '
' Long Description cell of the particular item. This form is  '
' the avenue to add and edit those notes.                     '
'                                                             '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private gbl_mark_num As Integer
Private vbCell As Range

Public Sub LoadForm(ByVal mark_num As Integer)
'Subroutine: LoadForm - Controls the .Show procedure for this form.
'                       Sets item description and form title for prescribed item.
'Arguments: mark_num - The mark number for the item

    If get_row(mark_num) = -1 Then
        Exit Sub
    End If
    
    gbl_mark_num = mark_num
    Set vbCell = VB_MASTER.Cells(get_row(mark_num), get_col_num("Long Description"))
    
    ItemDesc.Caption = get_property(mark_num, "Long Description")
    
    If HasNote(mark_num) Then
        AddEditNotes.Caption = "Edit Notes"
        NotesTextBox.Value = vbCell.Comment.text
    Else
        AddEditNotes.Caption = "Add Notes"
    End If
    
    NotesTextBox.SetFocus
    
    'set position on screen
    Me.Top = Application.Top + Application.height / 2 - Me.height / 2
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    Me.Show
    
    Set vbCell = Nothing
End Sub

Private Sub OKButton_Click()
'Subroutine: OKButton_Click - UserForm event handler. Called when OKButton is clicked. Sets Comment
    
    On Error Resume Next
    vbCell.Comment.Delete
    Err.Clear
    On Error GoTo 0
    
    Dim cmt As Comment
    If Len(NotesTextBox.Value) > 0 Then
        Set cmt = vbCell.AddComment(NotesTextBox.Value)
        cmt.Shape.TextFrame.AutoSize = True
    End If
    Set cmt = Nothing
    
    Unload Me
End Sub

Private Sub CancelButton_Click()
'Subroutine: CancelButton_Click - UserForm event handler. Called when CancelButton is clicked.

    Unload Me
End Sub
