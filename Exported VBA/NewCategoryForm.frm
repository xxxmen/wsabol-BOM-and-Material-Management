VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewCategoryForm 
   Caption         =   "New Category..."
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4545
   OleObjectBlob   =   "NewCategoryForm.frx":0000
End
Attribute VB_Name = "NewCategoryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Form: NewCategoryForm                                 '
'                                                            '
' On Add Category action, this form presents the user with   '
' text boxes to specify the name of the new category and the '
' unit.                                                      '
'                                                            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private new_cat_name As String
Private new_cat_row As Integer
Private new_cat_unit As String

Public Sub LoadForm(ByRef cat_name As String, ByVal cat_row As Integer)
'Subroutine: LoadForm - Controls the .Show routine for this form. Adds item categories to the list box, sets position on the screen, etc.
'Arguments: cat_name - String reference returning the new category name
'           cat_row - Integer containing the category insertion row


    new_cat_name = ""
    new_cat_unit = ""
    new_cat_row = cat_row
    
    NewCategory_TxtBx.SetFocus
    
    Me.Top = Application.Top + Application.height / 2 - Me.height / 2
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    Me.Show
    
    If get_cat_row(new_cat_name) = cat_row Then
        cat_name = new_cat_name
    Else
        cat_name = vbNullString
    End If
End Sub

Private Sub AddCategory_Button_Click()
'Subroutine: NewCategory_TxtBx_KeyPress - UserForm event handler. Called when user clicks AddCategory_Button.

    NewCategory_TxtBx.Value = TrimWhiteSpace(UCase(NewCategory_TxtBx.Value))
    Unit_TxtBx.Value = TrimWhiteSpace(UCase(Unit_TxtBx.Value))

    new_cat_name = NewCategory_TxtBx.Value
    new_cat_unit = Unit_TxtBx.Value
    
    'check if form is complete
    If Len(new_cat_unit) = 0 Or Len(new_cat_name) = 0 Then
        MsgBox "Please Complete the form.", vbExclamation
        Exit Sub
    End If
    
    ' check for preceding apostrophe in the name
    If Left(new_cat_name, 1) = "'" Then
        new_cat_name = Right(new_cat_name, Len(new_cat_name) - 1)
    End If
    
    ' check for preceding apostrophe in the unit
    If Left(new_cat_unit, 1) = "'" Then
        new_cat_unit = Right(new_cat_unit, Len(new_cat_unit) - 1)
    End If
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    'add category
    Call NewCategory(new_cat_name, new_cat_unit, new_cat_row)
    
    If ui_change Then RenderUI True
    
    Unload Me
    
    new_cat_name = get_category(new_cat_row)
End Sub

Private Sub CancelButton_Click()
'Subroutine: CancelButton_Click - UserForm event handler. Called when user clicks CancelButton.

    new_cat_name = vbNullString
    Unload Me
End Sub

Private Sub NewCategory_TxtBx_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Subroutine: NewCategory_TxtBx_KeyPress - UserForm event handler. Called when user types a charactor in NewCategory_TxtBx.
'Arguments: KeyAscii - Integer containing the key code from the key pressed.


    Dim acceptable_char As Boolean
    acceptable_char = CheckKeyCode("Name", KeyAscii)
    
    If Not acceptable_char Then
        MsgBox "You entered an illegal character. Category titles cannot contain " & vbCrLf & vbCrLf & "/ ? < > \ : * | " & Chr(34), vbCritical
        
        'cancel key pressed
        KeyAscii = 0
    ElseIf Len(NewCategory_TxtBx.Value) = 0 And KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Unit_TxtBx_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Subroutine: Unit_TxtBx_KeyPress - UserForm event handler. Called when user pressing a key down in Unit_TxtBx.
'Arguments: KeyAscii - Integer containing the key code from the key pressed.


    Dim acceptable_char As Boolean
    acceptable_char = CheckKeyCode("Unit", KeyAscii)
    
    If Not acceptable_char Then
        MsgBox "You entered an illegal character. Units cannot contain " & vbCrLf & vbCrLf & "/ ? < > \ : * | " & Chr(34) & " _ [ ] ^", vbCritical
        
        'cancel key pressed
        KeyAscii = 0
    ElseIf Len(Unit_TxtBx.Value) = 0 And KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Subroutine: UserForm_QueryClose - UserForm event handler. Called when form is closed.
'Arguments: Cancel - Integer returning a Cancel action.
'           CloseMode - Integer containing the close mode.


    If CloseMode <> vbFormCode Then ' with an Unload statement
        new_cat_name = vbNullString
    Else
        ' just close
    End If
End Sub
