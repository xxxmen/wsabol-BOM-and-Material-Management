VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChooseCategory 
   Caption         =   "Choose Category"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3240
   OleObjectBlob   =   "ChooseCategory.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChooseCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Form: ChooseCategory                           '
'                                                     '
' Lets the user choose a category to from the list of '
' categories in the project. Returns the selection in '
' LoadForm.                                           '
'                                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private successful As Boolean

Public Sub LoadForm(ByRef category As String)
'Subroutine: LoadForm - Controls the .Show routine for this form. Adds item categories to the list box, sets position on the screen, etc.
'Arguments: category - String returning the selection.

    
    Dim index As Integer
    Dim list_index As Integer
    Dim cat_row As Integer
    Dim cat_col As Integer
    cat_row = VB_CATEGORY.FirstRow()
    cat_col = VB_CATEGORY.CategoryColumn()
    successful = False
    
    'add categories to form
    Do While Not VB_CATEGORY.EOF(cat_row)
        CategoryBox.AddItem VB_CATEGORY.Cells(cat_row, cat_col).Value2
        
        cat_row = cat_row + 1
    Loop
    CategoryBox.ListIndex = 0
    VB_VAR_STORE.ClearChosenCategory
    
    Me.Top = Application.Top + Application.height / 2 - Me.height / 2
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    Me.Show
    
    category = VB_VAR_STORE.GetChosenCategory()
    VB_VAR_STORE.ClearChosenCategory
End Sub

Private Sub OKButton_Click()
'Subroutine: OKButton_Click - UserForm event handler. Called when user clicks OKButton.
    
    
    If Len(CategoryBox.Value) > 0 Then
        VB_VAR_STORE.SetChosenCategory CategoryBox.Value
        
        Unload Me
    Else
        MsgBox "Please choose a category before continuing.", vbExclamation
        'do nothing
    End If
End Sub

Private Sub CancelButton_Click()
'Subroutine: CancelButton_Click - UserForm event handler. Called when user clicks CancelButton.

    Unload Me
End Sub

