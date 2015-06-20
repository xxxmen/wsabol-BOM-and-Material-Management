VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReorderCategoriesForm 
   Caption         =   "Reorder Categories..."
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3240
   OleObjectBlob   =   "ReorderCategoriesForm.frx":0000
End
Attribute VB_Name = "ReorderCategoriesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Form: ReorderCategoriesForm                            '
'                                                             '
' When categories are added they are appended to the end of   '
' the BOM. For proper viewing/browsing, this provides a means '
' to reorder the categories in MASTER however the user sees   '
' fit. The order affects how item are sorted in VB_MASTER,    '
' VB_ORDER_TMPLT, and VB_RFP_REPORT                           '
'                                                             '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Public Sub LoadForm()
'Subroutine: LoadForm - Controls the .Show routine for this form.
'                       Adds item categories to the list, sets position on the screen, etc.


    'load Categoies to List
    Dim row As Integer
    row = VB_CATEGORY.FirstRow()
    
    Dim translate As Integer
    translate = Me.CategoryListBox.Font.Size + 2
    
    'ADD CATEGORIES TO LIST IN THE ORDER THEY APPEAR
    Do While Not VB_CATEGORY.EOF(row)
        CategoryListBox.AddItem VB_CATEGORY.Cells(row, VB_CATEGORY.CategoryColumn()).Value2
        
        If row > VB_CATEGORY.FirstRow() Then
            Me.height = Me.height + translate
            OKButton.Top = OKButton.Top + translate
            CancelButton.Top = CancelButton.Top + translate
            MoveUp.Top = MoveUp.Top + translate / 2
            MoveDown.Top = MoveDown.Top + translate / 2
            CategoryListBox.height = CategoryListBox.height + translate
        End If
        
        row = row + 1
    Loop
    CategoryListBox.ListIndex = 0
    CategoryListBox.SetFocus
    
    'set position on screen
    Me.Top = Application.Top + Application.height / 2 - Me.height / 2
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    Me.Show
End Sub

Private Sub MoveUp_Click()
'Subroutine: MoveUp_Click - UserForm event handler. Called when user clicks MoveUp.
'                           If item is not at the top of the list, move it up one position.

    Dim list_index As Integer
    list_index = CategoryListBox.ListIndex
    
    'check for error and if at top position
    If list_index = -1 Or list_index = 0 Then
        CategoryListBox.SetFocus
        Exit Sub
    End If
    
    Dim category As String
    category = CategoryListBox.Value
    CategoryListBox.RemoveItem list_index
    
    'decrease index
    list_index = list_index - 1
    
    'add item to list at new position
    CategoryListBox.AddItem category, list_index
    CategoryListBox.ListIndex = list_index
    CategoryListBox.SetFocus
End Sub

Private Sub MoveDown_Click()
'Subroutine: MoveUp_Click - UserForm event handler. Called when user clicks MoveDown.
'                           If item is not at the bottom of the list, move it down one position.


    Dim list_index As Integer
    list_index = CategoryListBox.ListIndex
    
    'check for error and if at bottom position
    If list_index = -1 Or list_index = CategoryListBox.ListCount - 1 Then
        CategoryListBox.SetFocus
        Exit Sub
    End If
    
    Dim category As String
    category = CategoryListBox.Value
    CategoryListBox.RemoveItem list_index
    
    'up index
    list_index = list_index + 1
    
    'add to list at new position
    CategoryListBox.AddItem category, list_index
    CategoryListBox.ListIndex = list_index
    CategoryListBox.SetFocus
End Sub

Private Sub OKButton_Click()
'Subroutine: OKButton_Click - UserForm event handler. Called when user clicks OKButton.
'                             Match order of categories in list to actual order of categories.
'                             If out of order is detected, create TEMP category, move all items to that,
'                             delete original, rename TEMP to original name.


    Dim temp_cat As String
    temp_cat = "TEMP"
    
    Dim current_category As String
    Dim next_category As String
    Dim category_unit As String
    Dim cat_row As Integer
    Dim cat_start As Integer
    Dim cat_end As Integer
    Dim change As Boolean
    change = False
    
    Dim index As Integer
    index = 0
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    Do While index < CategoryListBox.ListCount
        CategoryListBox.ListIndex = index
        
        next_category = CategoryListBox.Value
        current_category = VB_CATEGORY.Cells(VB_CATEGORY.FirstRow() + index, VB_CATEGORY.CategoryColumn()).Value2
        
        If current_category <> next_category Then
            cat_row = get_cat_row(current_category)
            category_unit = VB_CATEGORY.GetCategoryUnit(next_category)
            
            change = False
            If VB_CHANGE_LOG.TrackChanges() Then
                change = True
                VB_CHANGE_LOG.SetTrackChanges False
            End If
            
            'create new category in correct spot.
            Call NewCategory(temp_cat, category_unit, cat_row)
            
            'move items
            cat_start = get_cat_row(next_category)
            cat_end = cat_start + VB_MASTER.Cells(cat_start, VB_MASTER.CategoryColumn()).MergeArea.count - 1
            Do While cat_start <> -1
                AutoSortItem get_mark_num(cat_start), temp_cat
                
                cat_start = get_cat_row(next_category)
            Loop
            
            'delete original
            If get_cat_row(next_category) <> -1 Then
                DeleteCategory next_category
            End If
            
            'rename new
            VB_CATEGORY.RenameCategory temp_cat, next_category
            
            If change Then VB_CHANGE_LOG.SetTrackChanges True
        End If
        
        index = index + 1
    Loop
    
    If ui_change Then RenderUI True
    
    Unload Me
End Sub

Private Sub CancelButton_Click()
'Subroutine: CancelButton_Click - UserForm event handler. Called when user clicks CancelButton.

    Unload Me
End Sub
