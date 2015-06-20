VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PurgeCategoryForm 
   Caption         =   "Purge Category..."
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3240
   OleObjectBlob   =   "PurgeCategoryForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PurgeCategoryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Form: PurgeCategory                                  '
'                                                           '
' Lets the user choose a category to purge from the list of '
' categories in the project, and carries out the            '
' DeleteCategory procedure.                                 '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private successful As Boolean

Public Function LoadForm(Optional ByVal master_row As Integer = -1) As Boolean
'Function: LoadForm - Controls the .Show routine for this form. Adds item categories to the list box, sets position on the screen, etc.
'Arguments: master_row - Integer containing a row in master defining the default selection.
'Returns: Boolean. True - purge was successful. False - purge was not successful.

    
    Dim index As Integer
    Dim list_index As Integer
    Dim cat_row As Integer
    Dim category As String
    Dim cat_col As Integer
    index = 0
    list_index = index
    cat_row = VB_CATEGORY.FirstRow()
    cat_col = VB_CATEGORY.CategoryColumn()
    category = get_category(master_row)
    successful = False
    
    'add categories to form
    Do While Not VB_CATEGORY.EOF(cat_row)
        CategoryBox.AddItem VB_CATEGORY.Cells(cat_row, cat_col).Value2
        If category = VB_CATEGORY.Cells(cat_row, cat_col).Value2 Then
            list_index = index
        End If
        
        index = index + 1
        cat_row = cat_row + 1
    Loop
    
    CategoryBox.ListIndex = list_index
    
    Me.Top = Application.Top + Application.height / 2 - Me.height / 2
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    Me.Show
    
    LoadForm = successful
    
End Function

Private Sub PurgeButton_Click()
'Subroutine: PurgeButton_Click - UserForm event handler. Called when user clicks PurgeButton.
'                                If form is complete, purges category, otherwise prompts user
'                                to complete the form.
    
    
    If Len(CategoryBox.Value) > 0 Then
        Dim ui_change As Boolean
        
        ui_change = RenderUI(False)
        DeleteCategory CategoryBox.Value
        
        successful = (get_cat_row(CategoryBox.Value) = -1)
        
        VB_MASTER.CalculateQuantityFormat
        
        If ui_change Then RenderUI True
        
        If successful Then
            Unload Me
        End If
    Else
        MsgBox "Please choose a category before continuing.", vbExclamation
        'do nothing
    End If
End Sub

Private Sub CancelButton_Click()
'Subroutine: CancelButton_Click - UserForm event handler. Called when user clicks CancelButton.

    Unload Me
End Sub
