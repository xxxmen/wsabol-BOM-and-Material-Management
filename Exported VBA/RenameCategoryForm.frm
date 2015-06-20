VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RenameCategoryForm 
   Caption         =   "Rename Category..."
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4545
   OleObjectBlob   =   "RenameCategoryForm.frx":0000
End
Attribute VB_Name = "RenameCategoryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Form: RenameCategoryForm                                '
'                                                              '
' On a Rename Category action for the EditItemWindow, this     '
' form provides a means for the user to rename categories that '
' may be misspelled or poorly named.                           '
'                                                              '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Dim old_cat_name As String
Dim cat_row As Integer

Public Sub LoadForm(ByRef category_name As String, ByRef new_cat_name As String)
'Suroutine: LoadForm - Controls the .Show routine for this form. Sets focus, etc.
'Arguments: category_name - String containing existing category name
'           new_cat_name - String returning the new name for the category.


    cat_row = get_cat_row(category_name)
    old_cat_name = category_name
    
    'set focus
    NewName_TxtBx.Value = category_name
    NewName_TxtBx.SetFocus
    
    'sets screen position
    Me.Top = Application.Top + Application.height / 2 - Me.height / 2
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    Me.Show
    
    If cat_row <> -1 Then
        new_cat_name = get_category(cat_row)
    Else
        new_cat_name = vbNullString
    End If
End Sub

Private Sub AcceptButton_Click()
'Subroutine: AcceptButton_Click - UserForm event handler. Called when user clicks AcceptButton.

    Dim new_cat_name As String
    
    NewName_TxtBx.Value = TrimWhiteSpace(UCase(NewName_TxtBx.Value))
    new_cat_name = NewName_TxtBx.Value
    
    'check for completed form
    If Len(new_cat_name) = 0 Then
        MsgBox "Please complete the form.", vbExclamation
        NewName_TxtBx.SetFocus
        Exit Sub
    End If
    
    'check for preceding "'"
    If Left(new_cat_name, 1) = "'" Then
        new_cat_name = Right(new_cat_name, Len(new_cat_name) - 1)
    End If
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    'Rename Category
    If old_cat_name <> new_cat_name Then
        VB_CATEGORY.RenameCategory old_cat_name, new_cat_name
    Else
        MsgBox "This is the same name. Try again.", vbExclamation
        NewName_TxtBx.Value = ""
        NewName_TxtBx.SetFocus
        If ui_change Then RenderUI True
        Exit Sub
    End If
    
    'setup returns
    cat_row = get_cat_row(new_cat_name)

    If ui_change Then RenderUI True
    Unload Me
End Sub

Private Sub CancelButton_Click()
'Subroutine: CancelButton_Click - UserForm event handler. Called when user clicks CancelButton.

    cat_row = -1
    Unload Me
End Sub

Private Sub NewName_TxtBx_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Subroutine: NewName_TxtBx_KeyPress - UserForm event handler. Called when user pressing a key down in NewName_TxtBx.
'Arguments: KeyAscii - Integer containing the key code from the key pressed.


    Dim acceptable_char As Boolean
    acceptable_char = CheckKeyCode("Name", KeyAscii)
    
    If Not acceptable_char Then
        MsgBox "You entered an illegal character. Site names cannot contain " & vbCrLf & vbCrLf & "/ ? < > \ : * | " & Chr(34), vbExclamation
        
        'cancel key pressed
        KeyAscii = 0
    ElseIf Len(NewName_TxtBx.Value) = 0 And KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Subroutine: UserForm_QueryClose - UserForm event handler. Called when form is closed.
'Arguments: Cancel - Integer returning a Cancel action.
'           CloseMode - Integer containing the close mode.

    VB_MASTER.CalculateQuantityFormat
    If CloseMode <> vbFormCode Then ' with an Unload statement
        cat_row = -1
    Else
        ' just close
    End If
End Sub
