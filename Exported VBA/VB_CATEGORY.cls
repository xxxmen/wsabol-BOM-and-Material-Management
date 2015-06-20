VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VB_CATEGORY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Row/Column location constants
Private Const first_row = 2
Private Const category_col = 1
Private Const unit_col = 2
Private Const list_count_col = 3
Private Const first_words_col = 4

Public Function FindCategory(ByVal first_phrase As String) As String
'Function: FindCategory - Searches the CategoryMatrix for first_phrase and returns the category in which its belongs
'Arguments: find_phrase - String contains the 'first phrase' value to locate
'Returns: String containing the matched category; if first_phrase doesn't yet belong to a category, returns vbNullString


    FindCategory = vbNullString

    Dim count As Integer
    Dim row As Integer
    Dim i As Integer
    row = first_row

    Do While Not VB_CATEGORY.EOF(row)
        count = VB_CATEGORY.Cells(row, list_count_col).Value2
        
        For i = first_words_col To first_words_col + count - 1
            If CStr(VB_CATEGORY.Cells(row, i).Value2) = first_phrase Then
                FindCategory = VB_CATEGORY.Cells(row, category_col).Value2
                Exit Function
            End If
        Next i
        
        row = row + 1
    Loop
End Function

Public Sub StoreFirstPhraseKey(ByVal category As String, ByVal first_phrase As String)
'Subroutine: StoreFirstPhraseKey - Stores first_phrase in the designated category. If the category doesn't exist in the
'                                CategoryMatrix, it will add it to the matrix. If first_phrase already exists in category, it will do nothing.
'Arguments: category - String containing the category in which to store the first_phrase
'           first_phrase - String containing the first_phrase value to store.


    'check if first_phrase already exists in any category.
    If Len(FindCategory(first_phrase)) > 0 Then
        'first_phrase already exists
        Exit Sub
    End If

    'check if category exists in MASTER
    Dim mst_cat_row As Integer
    mst_cat_row = get_cat_row(category)
    
    If mst_cat_row = -1 Then
        'category doesn't exists. Period.
        Exit Sub
    End If

    'get category row
    Dim row As Integer
    row = GetCategoryRow(category)
    
    If row = -1 Then
        ' category doesn't exist, add it
        AddCategoryKey category, get_property(get_mark_num(mst_cat_row), "Unit")
        row = GetCategoryRow(category)
    End If
    
    Dim count As Integer
    
    ' set first_phrase
    count = VB_CATEGORY.Cells(row, list_count_col).Value2
    VB_CATEGORY.Cells(row, first_words_col + count).Value = first_phrase
    
    ' set count value
    VB_CATEGORY.Cells(row, list_count_col).Value = count + 1
End Sub

Public Sub ClearPhrases(ByVal category As String)
'Subroutine: ClearPhrases - Removes all phrases from a category in this matrix; resetting phrase count to 0
'Arguments: category - the category to perform the action on

    
    Dim row As Integer
    row = GetCategoryRow(category)
    
    If row = -1 Then ' category doesn't exist
        Exit Sub
    End If
    
    Dim column As Integer
    column = first_words_col
    
    Do While Not IsEmpty(VB_CATEGORY.Cells(row, column))
        VB_CATEGORY.Cells(row, column).ClearContents
        
        column = column + 1
    Loop
    
    VB_CATEGORY.Cells(row, list_count_col).Value = 0
End Sub

Public Sub AddCategoryKey(ByVal category As String, ByVal unit As String, Optional ByVal cat_row As Integer = -1)
'Subroutine: AddCategoryKey - Appends new category to CategoryMatrix with zero first_phrase values.
'                             If the category already exists, it will do nothing.
'Arguments: category - String containing the new category to add
'           unit - String containing the category's unit
'           cat_row - Row number in MASTER corresponding to the new category. Required to maintain matching order.
    
    
    If GetCategoryRow(category) > 0 Then
        'category already exists; exit
        Exit Sub
    End If
    
    If cat_row = -1 Then
        cat_row = VB_CATEGORY.UsedRange.Rows.count + 1
    Else
        Dim row As Integer
        row = first_row
        Do While Not EOF(row)
            If cat_row < get_cat_row(VB_CATEGORY.Cells(row, category_col).Value2) Then
                VB_CATEGORY.Rows(row).Insert
                Exit Do
            End If
            row = row + 1
        Loop
    End If
    
    ' set category value
    VB_CATEGORY.Cells(row, category_col).Value = category
    VB_CATEGORY.Rows(row).AutoFit
    
    'set unit value
    VB_CATEGORY.Cells(row, unit_col).Value = unit
    
    ' set count value
    VB_CATEGORY.Cells(row, list_count_col).Value = 0
End Sub

Public Sub DeleteCategory(ByVal category As String)
'Subroutine: DeleteCategory - Deletes category from CategoryMatrix. This subroutine is called
'                             everytime there is a category deletion in Master. If the category does not
'                             exist in the matrix, it does nothing.
'Arguments: category - String containing the category name to delete.


    Dim cat_row As Integer
    cat_row = GetCategoryRow(category)
    
    If cat_row <> -1 Then
        VB_CATEGORY.Rows(cat_row).Delete
        VB_MASTER.ResetScrollArea
    End If
End Sub

Public Sub RenameCategory(ByVal current_category As String, ByVal new_name As String)
'Subroutine: RenameCategory - Renames the category 'current_category' in this project to 'new_name'.
'                             If the category doesn't exist, it does nothing.
'Arguments: current_category - String containing the category's current name
'           new_name - String containing the category's new name


    Dim cat_row As Integer
    cat_row = get_cat_row(current_category)
    
    'check if category exists
    If cat_row = -1 Then
        Exit Sub
    End If
    
    'rewrite name in MASTER
    VB_MASTER.Cells(cat_row, VB_MASTER.CategoryColumn()).MergeArea.Cells(1, 1).Value = UCase(new_name)
    
    'rewrite name in VB_CATEGORY
    cat_row = GetCategoryRow(current_category)
    VB_CATEGORY.Cells(cat_row, category_col).Value = UCase(new_name)
    
    'autofit
    VB_CATEGORY.Rows(cat_row).AutoFit
    
    'LOG CHANGE
    If VB_CHANGE_LOG.TrackChanges() Then
        If VB_CHANGE_LOG.LogChange(VB_MASTER.Name & "!$" & ColLet(VB_MASTER.CategoryColumn()) & "$" & cat_row, "", "Renamed Category", current_category, new_name) <> 0 Then
            'throw error
            ErrorHandling "LogOrder", 0, "LogChange() operation failed:" & vbCrLf & _
                "LogChange(" & VB_MASTER.Name & "!$" & ColLet(VB_MASTER.CategoryColumn()) & "$" & cat_row & ", , Renamed Category, " & current_category & ", " & new_name & ")", 1
        End If
    End If
    
    CategoryAutoFit new_name
End Sub

Public Sub CategoryAutoFit(ByVal category As String)
'Subroutine: CategoryAutoFit - Controls category row height so that the category title is always visible
'Arguments: category - String containing the category to autofit

    
    Dim cat_start As Integer
    cat_start = get_cat_row(category)
    
    'check if category exists
    If cat_start = -1 Then
        Exit Sub
    End If
    
    'finds category's end
    Dim cat_end As Integer
    cat_end = cat_start + VB_MASTER.Cells(cat_start, VB_MASTER.CategoryColumn()).MergeArea.Cells.count - 1
    
    Dim row As Integer
    Dim height As Integer
    height = 0
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    'add total height for items in category
    For row = cat_start To cat_end
        VB_MASTER.Rows(row).AutoFit
        height = height + VB_MASTER.Rows(row).RowHeight
    Next row
    
    Dim cat_row As Integer
    cat_row = GetCategoryRow(category)
    
    'compare to min height seen in CATEGORYMATRIX
    If height - 1.5 < VB_CATEGORY.Rows(cat_row).RowHeight Then
        VB_MASTER.Rows(cat_end).RowHeight = VB_MASTER.Rows(cat_end).RowHeight + VB_CATEGORY.Rows(cat_row).RowHeight + 1.5 - height
    End If
    
    If ui_change Then RenderUI True
End Sub

Private Function GetCategoryRow(ByVal category As String) As Integer
'Function: GetCategoryRow - Returns the row where this application can find the
'                           specified category in CategoryMatrix. If category doesn't exist, returns -1.
'Arguments: category - String containing the category in question.
'Returns: Integer containing the category row, or -1 if the category doesn't exist.

    
    Dim row As Integer
    GetCategoryRow = first_row
    
    Do While Not VB_CATEGORY.EOF(GetCategoryRow)
        If CStr(VB_CATEGORY.Cells(GetCategoryRow, category_col).Value2) = category Then
            Exit Function
        End If
        GetCategoryRow = GetCategoryRow + 1
    Loop
    
    ' category was not found, return -1
    GetCategoryRow = -1
End Function

Public Function GetCategoryUnit(ByVal category As String) As String
'Function: GetCategoryUnit - Returns the respective Unit for the specified category.
'Arguments: category - String containing the category.
'Returns: String containing the category's Unit. If category doesn't exist, vbNullString is returned.


    GetCategoryUnit = vbNullString

    Dim cat_row As Integer
    cat_row = GetCategoryRow(category)
    
    If cat_row <> -1 Then
        GetCategoryUnit = VB_CATEGORY.Cells(cat_row, unit_col).Value2
    End If
End Function

Public Function CategoryExists(ByVal category As String) As Boolean
'Function: CategoryExists - Tests if category exists in this BOM.
'Arguments: category - String containing the category.
'Returns: Boolean. True - category exists; False - category does not exist


    CategoryExists = (GetCategoryRow(category) <> -1)
End Function

Public Function EOF(ByVal row As Integer) As Boolean
'Function: EOF - Tests if row is beyond the scope of this sheet
'Arguments: row - Integer containing the row number in this sheet
'Returns: Boolean. True - row is beyond the scope; False - row is within the meaningful scop


    EOF = IsEmpty(VB_CATEGORY.Cells(row, category_col))
End Function

Public Function FirstRow() As Integer
'Function: FirstRow - 'Get' method for the first row property for this sheet
'Returns: Integer containing the first row of meaningful data in the sheet.


    FirstRow = first_row
End Function

Public Function CategoryColumn() As Integer
'Function: CategoryColumn - 'Get' method for the category_column property for this sheet
'Returns: Integer containing the category column number.


    CategoryColumn = category_col
End Function

Public Function UnitColumn() As Integer
'Function: UnitColumn - 'Get' method for the unit_column property for this sheet
'Returns: Integer containing the unit column number.


    UnitColumn = unit_col
End Function

Public Function PhrasesColumn() As Integer
'Function: PhrasesColumn - 'Get' method for the first_words_col property for this sheet
'Returns: Integer containing the first phrases column number.
    
    
    PhrasesColumn = first_words_col
End Function
