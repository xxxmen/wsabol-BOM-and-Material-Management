VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddItemForm 
   Caption         =   "Choose Item Category"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4020
   OleObjectBlob   =   "AddItemForm.frx":0000
End
Attribute VB_Name = "AddItemForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Form: AddItemForm                                      '
'                                                             '
' To add an item to the BOM, the user must specify the item   '
' category. This form lets the user choose the category for a '
' pending item, and add the add in that category.             '
'                                                             '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private new_mark_num As Integer

Public Function LoadForm(ByVal in_desc As String) As Integer
'Function: LoadForm - Controls the .Show routine for this form. Adds item categories to the list box, sets position on the screen, etc.
'Arguments: in_desc - String containing description of new item.
'Returns: Integer containing the mark number for the new item == [-1] U [1, Inf)


    'set item description
    Item_Desc.Caption = in_desc
    
    Dim row As Integer
    Dim cat_col As Integer
    Dim index As Integer
    Dim list_index As Integer
    row = VB_CATEGORY.FirstRow()
    cat_col = VB_CATEGORY.CategoryColumn()
    list_index = 0
    index = 0
    
    'add categories
    Do While Not VB_CATEGORY.EOF(row)
        ItemCategories_List.AddItem VB_CATEGORY.Cells(row, cat_col).Value2
        If VB_CATEGORY.FindCategory(FirstPhrase(in_desc)) = VB_CATEGORY.Cells(row, cat_col).Value2 Then
            list_index = index
        End If
        
        index = index + 1
        row = row + 1
    Loop

    'set best guess
    ItemCategories_List.ListIndex = list_index
    
    'set position on screen
    Me.Top = Application.Top + Application.height / 2 - Me.height / 2
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    Me.Show
    
    LoadForm = new_mark_num
End Function

Private Sub InsertButton_Click()
'Subroutine: InsertButton_Click - UserForm event handler. Called when user clicks InsertButton.
'                                 If form is complete, add new item to BOM and unloads form, otherwise
'                                 prompts user to complete the form.


    If Len(ItemCategories_List.Value) > 0 Then
        new_mark_num = InsertItemBOM(ItemCategories_List.Value, Item_Desc.Caption)
        VB_CATEGORY.StoreFirstPhraseKey ItemCategories_List.Value, FirstPhrase(Item_Desc.Caption)
        Unload Me
    Else
        MsgBox "Please choose a category before continuing.", vbExclamation
        'do nothing
    End If
End Sub

Private Sub CancelButton_Click()
'Subroutine: CancelButton_Click - UserForm event handler. Called when user clicks CancelButton.
'                                 Unloads form; sets new_mark_num to -1

    new_mark_num = -1
    Unload Me
End Sub

Private Sub NewCategoryButton_Click()
'Subroutine: NewCategoryButton_Click - UserForm event handler. Called when user clicks NewCategoryButton.
'                                      Runs routine to create new category, and refreshes category list for user to choose from.


    Dim new_cat As String
    new_cat = NewCategoryHandler()
    
    Do While ItemCategories_List.ListCount > 0
        ItemCategories_List.RemoveItem ItemCategories_List.ListCount - 1
    Loop
    
    Dim row As Integer
    Dim cat_col As Integer
    Dim index As Integer
    Dim list_index As Integer
    row = VB_CATEGORY.FirstRow()
    list_index = 0
    index = 0
    cat_col = VB_CATEGORY.CategoryColumn()
    
    Do While Not VB_CATEGORY.EOF(row)
        ItemCategories_List.AddItem VB_CATEGORY.Cells(row, cat_col).Value2
        If CStr(VB_CATEGORY.Cells(row, cat_col).Value2) = new_cat Then
            list_index = index
        End If
        
        index = index + 1
        row = row + 1
    Loop
    
    'set best guess
    ItemCategories_List.ListIndex = list_index
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Subroutine: UserForm_QueryClose - UserForm event handler. Called when form is closed.
'Arguments: Cancel - Integer returning a Cancel action.
'           CloseMode - Integer containing the close mode.

    
    If CloseMode <> vbFormCode Then ' with an Unload statement
        CancelButton_Click
        Cancel = True
    Else
        ' just close
    End If
End Sub
