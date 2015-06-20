Attribute VB_Name = "MasterBOMMembers"
Option Explicit

Public Function get_mark_num(ByVal master_row As Integer) As Integer
'Function: get_mark_num - Retrieves the "Mark No." from MASTER in the given row.
'Arguments: master_row - Integer containing the row in MASTER
'Returns: Integer containing the "Mark No."
    
    
    Dim markRng As Range
    
    On Error GoTo NONPOSITIVEROW
    Set markRng = VB_MASTER.Cells(master_row, get_col_num("Mark No."))

    If IsEmpty(markRng) Then
        get_mark_num = -1
    Else
        get_mark_num = CInt(markRng.Value2)
    End If
    
    Err.Clear
    On Error GoTo 0
    Set markRng = Nothing
    Exit Function
    
NONPOSITIVEROW:
    get_mark_num = -1
End Function

Public Function getMark4UniqueID(ByVal uniqueID As String) As Integer
'Function: getMark4UniqueID - For a given uniqueID, this function returns the corresponding mark number
'Arguments: uniqueID - String containing the Unique ID
'Returns: Integer containing the mark number


    Dim unIDcol As Integer
    unIDcol = get_col_num("Unique ID")
    
    Dim first As Integer
    Dim last As Integer
    first = VB_MASTER.FirstRow()
    last = VB_MASTER.LastRow()
    Dim row As Integer
    For row = first To last
        If CStr(VB_MASTER.Cells(row, unIDcol).Value2) = uniqueID Then
            getMark4UniqueID = CInt(VB_MASTER.Cells(row, get_col_num("Mark No.")).Value2)
            Exit Function
        End If
    Next row
    
    getMark4UniqueID = -1
End Function

Public Function get_ColTitle(ByVal col_num As Integer) As String
'Function: get_ColTitle - Retrieves the column title from MASTER for the given cell location.
'                         It can handle column titles that are part of merged cells
'Arguments: locationStr - String containing the cell location
'Returns: String containing the column title


    'Define the title cell for this column
    Dim title_cell As Range
    Set title_cell = VB_MASTER.Cells(VB_MASTER.TitleRow(), col_num)
    
    'Test to see if its in a merged cell and handle it
    If title_cell.MergeCells Then
        get_ColTitle = CStr(title_cell.MergeArea.Cells(1, 1).Value2)
    Else
        get_ColTitle = CStr(title_cell.Value2)
    End If
    
    Set title_cell = Nothing
End Function

Public Function get_first_row(Optional srcSheet As Worksheet) As Integer
'Function: get_first_row - Retrieve row number in MASTER for a given Mark No
'Arguments: mark_num - Integer containing the mark number
'Returns: Integer containing row number
    
    
    If srcSheet Is Nothing Then
        Set srcSheet = VB_MASTER
    End If
    
    Dim col As Integer
    'Find the column number which has the "Mark No." title
    col = get_col_num("Mark No.", srcSheet)

    get_first_row = -1

    'Go to row where Mark No. = mark_num
    Dim row As Integer
    For row = 1 To srcSheet.UsedRange.Rows.count
        If IsNumeric(srcSheet.Cells(row, col).Value2) And Not IsEmpty(srcSheet.Cells(row, col)) Then
            get_first_row = row
            Exit Function
        End If
    Next row
    
End Function

Public Function get_row(ByVal mark_num As Integer, Optional srcSheet As Worksheet) As Integer
'Function: get_row - Retrieve row number in MASTER for a given Mark No
'Arguments: mark_num - Integer containing the mark number
'Returns: Integer containing row number
    
    get_row = -1
    
    If mark_num < 1 Then
        Exit Function
    End If
    
    If srcSheet Is Nothing Then
        Set srcSheet = VB_MASTER
    End If
    
    Dim col As Integer
    'Find the column number which has the "Mark No." title
    col = get_col_num("Mark No.", srcSheet)
    
    'Go to row where Mark No. = mark_num
    Dim row As Integer
    For row = 1 To srcSheet.UsedRange.Rows.count
        If IsNumeric(srcSheet.Cells(row, col).Value2) And Not IsEmpty(srcSheet.Cells(row, col)) Then
            If CInt(srcSheet.Cells(row, col).Value2) = mark_num Then
                get_row = row
                Exit Function
            End If
        End If
    Next row
End Function

Public Function LastMark() As Integer
'Function: LastMark - gets the mark number from the last row item in VB_MASTER.
'Returns: Integer containing the last mark number.


    Dim row As Integer
    row = VB_MASTER.FirstRow()

    Dim mark_col As Integer
    mark_col = get_col_num("Mark No.")
    
    Do While Not IsEmpty(VB_MASTER.Cells(row + 2, mark_col))
        row = row + 2
    Loop
    
    If Not IsEmpty(VB_MASTER.Cells(row + 1, mark_col)) Then
        row = row + 1
    End If
    
    LastMark = CInt(VB_MASTER.Cells(row, mark_col).Value2)
End Function

Public Function MaxMark() As Integer
'Function: MaxMark - gets the highest mark number from VB_MASTER.
'Returns: Integer containing the highest mark number.


    Dim first As Integer
    Dim last As Integer
    first = VB_MASTER.FirstRow()
    last = VB_MASTER.LastRow()
    
    Dim mark_col As String
    mark_col = ColLet(get_col_num("Mark No."))
    
    MaxMark = VBRangeMax(VB_MASTER.Name & "!$" & mark_col & "$" & first & ":$" & mark_col & "$" & last)
End Function

Public Function get_cat_row(ByVal category As String) As Integer
'Function: get_cat_row - Return the row number for the first row of the specified category.
'                        If category doesn't exist, returns -1
'Arguments: category - String containing the category to search for
'Returns: Integer containing the relavent row number


    get_cat_row = -1
    Dim row As Integer
    row = VB_MASTER.FirstRow()
    
    'assumes item categories are continuous. There are no skipped rows between categories.
    Do While Not IsEmpty(VB_MASTER.Cells(row, VB_MASTER.CategoryColumn()).MergeArea)
        If get_category(row) = category Then
            get_cat_row = row
            Exit Function
        End If
        row = row + VB_MASTER.Cells(row, VB_MASTER.CategoryColumn()).MergeArea.Cells.count
    Loop

End Function

Public Function IsDeleted(ByVal mark_num As Integer) As Boolean
'Function: IsDeleted - Returns whether an item has been ClientDeleted.
'Arguments: mark_num - mark number for the item to check
'Returns: Boolean. True - is deleted. False - is not deleted

    IsDeleted = (Len(get_property(mark_num, "Delete?")) > 0)
End Function

Public Function HasNote(ByVal mark_num As Integer) As Boolean
'Function: HasNote - Returns whether an item has a note added to the description.
'Arguments: mark_num - mark number for the item to check
'Returns: Boolean. True - has note. False - does not have note

    HasNote = Not VB_MASTER.Cells(get_row(mark_num), get_col_num("Long Description")).Comment Is Nothing
End Function

Public Function InsertItemBOM(ByVal category As String, ByVal DESC As String) As Integer
'Function: InsertItemBOM - Inserts a newly specified item into the BOM at a given category
'                          Based on the description, it will determine what row it is inserted to.
'                          Sets the unit, and sets the description based on description argument
'Arguments: category - String containing the category in which the new item will go
'           desc - String containing the description of the new item
'Returns: Integer containing the mark number for the new item


    Dim mark_num As Integer
    
    'Will be the row where the item is inserted, but because we have to search
    'for the right spot, start at the beginning of the category
    Dim row As Integer
    row = get_cat_row(category) ' category start
    
    If row = -1 Then
        InsertItemBOM = -1
        Exit Function
    End If
    
    Dim desc_col As Integer
    desc_col = get_col_num("Long Description")
    
    Dim cat_end As Integer
    cat_end = row + VB_MASTER.Cells(row, VB_MASTER.CategoryColumn()).MergeArea.count - 1
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    If Not IsEmpty(VB_MASTER.Cells(row, desc_col)) Then
        
        ' determine where in the category this item should go
        Do While row <= cat_end
            
            If DescriptionCompare(DESC, VB_MASTER.Cells(row, desc_col).Value2) = 1 Then
                Exit Do
            End If
        
            row = row + 1
        Loop
        
        'Insert the row
        InsertMasterRow category, row, mark_num
        
        'set the description property
        Call set_property(mark_num, get_col_num("Long Description"), DESC)
    Else
        'no items in category yet.
        Dim unit_col As Integer
        unit_col = get_col_num("Unit")
        VB_MASTER.Cells(row, desc_col).Value = DESC
        
        mark_num = get_mark_num(row)
        'If mark_num = -1 Then
    End If
    
    'set the unit property
    Call set_property(mark_num, get_col_num("Unit"), VB_CATEGORY.GetCategoryUnit(category))
    
    'init Client Inventory property
    Call set_property(mark_num, get_col_num("Client Inventory"), "0")
    
    'init Long Lead property
    Call set_property(mark_num, get_col_num("Long Lead"), False)
    
    VB_CATEGORY.CategoryAutoFit get_category(row)
    
    If ui_change Then RenderUI True
    
    'LOG CHANGE
    If VB_CHANGE_LOG.TrackChanges() Then
        If VB_CHANGE_LOG.LogChange(VB_MASTER.Name & "!$" & ColLet(get_col_num("Mark No.")) & "$" & row, mark_num, "Added Material Item", category, DESC) <> 0 Then
            'throw error
            ErrorHandling "InsertItemBOM", 10, "LogChange(" & VB_MASTER.Name & "!$" & ColLet(get_col_num("Mark No.")) & "$" & row & ", " & mark_num & _
                ", Added Material Item, " & category & ", " & DESC & ")", 1
        End If
    End If
    
    'Store first phrase inincase it doesn't exist
    VB_CATEGORY.StoreFirstPhraseKey category, FirstPhrase(DESC)
    
    'return the new item's mark number
    InsertItemBOM = mark_num

End Function

Public Function AutoSortItem(ByVal mark_num As Integer, Optional ByVal category As String) As Integer
'Function: AutoSortItem - On an Item Description change, we must check sorting to ensure the right order for RFP reports.
'                         Order is measured with DescriptionCompare. If out of order: Add a duplicate item, copy properties,
'                         delete original item, set original mark number.
'Arguments: mark_num - mark number for item to check sort
'           category - category scope. If going to a new category, specify. If category not specified, original category is used
'Returns: Integer containing the new row number for the item.


    'check if mark num exists
    AutoSortItem = get_row(mark_num)
    If AutoSortItem = -1 Then
        Exit Function
    End If
    
    'get description
    Dim in_desc As String
    in_desc = get_property(mark_num, "Long Description")
    
    Dim original_category As String
    original_category = get_category(AutoSortItem)
    
    If Len(category) = 0 Then
        category = original_category
    End If
    
    Dim cat_start As Integer
    Dim cat_end As Integer
    cat_start = get_cat_row(original_category)
    cat_end = cat_start + VB_MASTER.Cells(AutoSortItem, VB_MASTER.CategoryColumn()).MergeArea.count - 1
    
    'CHECK SORT
    Dim correctly_sorted As Boolean
    If cat_start = cat_end Then
        correctly_sorted = (category = original_category)
        
    ElseIf AutoSortItem = cat_start Then
        correctly_sorted = (DescriptionCompare(in_desc, get_property(get_mark_num(AutoSortItem + 1), "Long Description")) = 1) And _
            (category = original_category)
        
    ElseIf AutoSortItem = cat_end Then
        correctly_sorted = (DescriptionCompare(in_desc, get_property(get_mark_num(AutoSortItem - 1), "Long Description")) = -1) And _
            (category = original_category)
        
    Else
        correctly_sorted = (DescriptionCompare(in_desc, get_property(get_mark_num(AutoSortItem - 1), "Long Description")) = -1) And _
            (DescriptionCompare(in_desc, get_property(get_mark_num(AutoSortItem + 1), "Long Description")) = 1) And _
            (category = original_category)
    End If
    
    
    If Not correctly_sorted Then
        'ITEM NOT CORRECTLY SORTED
        
        Dim tc As Boolean
        tc = VB_CHANGE_LOG.TrackChanges()
        VB_CHANGE_LOG.SetTrackChanges False
        
        Dim ui_change As Boolean
        ui_change = RenderUI(False)
        
        'ADD ITEM, GET CURRENT ROW
        Dim new_row As Integer
        Dim temp_mark As Integer
        temp_mark = InsertItemBOM(category, in_desc)
        new_row = get_row(temp_mark)
        AutoSortItem = get_row(mark_num)
        
        'COPY ITEM PROPERTIES, except mark num
        Dim cmt As Comment
        Dim col As Integer
        Dim first_col As Integer
        Dim last_col As Integer
        Dim mark_col As Integer
        first_col = get_col_num("Orders")
        last_col = LastColumn()
        mark_col = get_col_num("Mark No.")
        
        For col = get_col_num("Orders") To LastColumn()
            If Not (VB_MASTER.Cells(new_row, col).HasFormula Or col = mark_col) Then
                VB_MASTER.Cells(new_row, col).Value = VB_MASTER.Cells(AutoSortItem, col).Value2
                Set cmt = VB_MASTER.Cells(AutoSortItem, col).Comment
                If Not cmt Is Nothing Then
                    If Len(cmt.text) > 0 Then
                        VB_MASTER.Cells(new_row, col).AddComment cmt.text
                        VB_MASTER.Cells(new_row, col).Comment.Shape.TextFrame.AutoSize = True
                    End If
                End If
            End If
        Next col
        Set cmt = Nothing
        
        'DELETE ORIGINAL
        DeleteMasterRow AutoSortItem
        
        If cat_start = cat_end Then 'was only item in category, so delete category
            DeleteCategory original_category
        End If
        
        If new_row > AutoSortItem Then
            new_row = new_row - 1
        End If
        
        AutoSortItem = new_row
        
        'COPY MARK NUMBER
        VB_MASTER.Cells(AutoSortItem, mark_col).Value = mark_num
        VB_MASTER.CalculateQuantityFormat AutoSortItem
        
        'REFRESH CELLS LOCATIONS IN VB_CHANGE_LOG
        Dim cl_row As Integer
        Dim clCell As Range
        cl_row = VB_CHANGE_LOG.FirstRow()
        
        With VB_CHANGE_LOG
            Do While Not .EOF(cl_row)
                If CInt(.Cells(cl_row, .ColID("Mark")).Value2) = mark_num Then
                    If IsCellLocation(.Cells(cl_row, .ColID("Cell")).Value2, clCell) Then
                        .Cells(cl_row, .ColID("Cell")).Value = VB_MASTER.Name & "!" & VB_MASTER.Cells(AutoSortItem, clCell.column).Address
                    End If
                End If
                
                cl_row = cl_row + 1
            Loop
        End With
        
ErrorExit:
        If ui_change Then RenderUI True
        VB_CHANGE_LOG.SetTrackChanges tc
        
    End If
    
End Function

Public Function get_property(ByVal mark_num As Integer, ByVal col_desc As String)
'Function: get_property - Retrieve any row property from MASTER given a Mark No.
'Arguments: mark_num - Integer containing the mark number
'           col_desc - String containing the column title for the property one is looking for
'Returns: String containing the retrieved property. If mark number doesn't exist, returns vbNullString
    
    Dim row As Integer
    row = get_row(mark_num)
    
    If row = -1 Then
        get_property = vbNullString
        Exit Function
    End If
    
    Dim col As Integer
    col = get_col_num(col_desc)

    get_property = CStr(VB_MASTER.Cells(row, col).Value2)
End Function

Public Function MatchDesc_like(ByVal DESC As String) As Integer
'Function: MatchDesc_like - Matches the given description to a description in the BOM using the Like operator.
'                           Used to find a best guess estimate.
'                           If it finds a match, return the corresponding mark number, else return -1
'Arguments: desc - String containing the description with Like operators to search for
'Returns: Integer containing the corresponding mark number, or if the description doesn't exist, -1


    Dim temp_row As Integer
    Dim last_row As Integer
    Dim temp_mark As Integer
    Dim test_desc As String
    temp_row = VB_MASTER.FirstRow()
    last_row = VB_MASTER.LastRow()
    
    Do While temp_row <= last_row
        temp_mark = get_mark_num(temp_row)
        test_desc = get_property(temp_mark, "Long Description")
        If test_desc Like DESC Then
            MatchDesc_like = temp_mark
            Exit Function
        End If
        
        temp_row = temp_row + 1
    Loop

    MatchDesc_like = -1
End Function

Public Function MatchDesc(ByVal DESC As String, Optional ByVal estimate As Boolean = False) As Integer
'Function: MatchDesc - Matches the given description to a description in the BOM or finds best guess.
'                      If it finds a match, return the corresponding mark number, else return -1
'Arguments: desc - String containing the description to search for
'           estimate - Boolean determining whether to estimate a best guess or not.
'Returns: Integer containing the corresponding mark number, or if the description doesn't exist, -1


    Dim follow As Boolean
    Dim comp_track As Integer
    Dim comp_track_temp As Integer
    Dim temp_row As Integer
    Dim last_row As Integer
    Dim temp_mark As Integer
    Dim test_desc As String
    
    temp_row = VB_MASTER.FirstRow()
    last_row = VB_MASTER.LastRow()
    follow = False
    comp_track_temp = 0
    comp_track = comp_track_temp
    
    Do While temp_row <= last_row
        temp_mark = get_mark_num(temp_row)
        test_desc = get_property(temp_mark, "Long Description")
        follow = (StrComp(FirstPhrase(test_desc), FirstPhrase(DESC), vbTextCompare) = 0)
        
        comp_track_temp = DescriptionCompare(DESC, test_desc)
        If comp_track_temp = 0 Then
            MatchDesc = temp_mark
            Exit Function
        ElseIf follow And estimate Then
            'comp_track_temp = DescriptionCompare(desc, test_desc)
            If comp_track_temp = 1 And comp_track = -1 Then
                MatchDesc = get_mark_num(temp_row - 1)
                Exit Function
            Else
                comp_track = comp_track_temp
            End If
        End If
        
        temp_row = temp_row + 1
    Loop

    If estimate Then
        temp_row = get_row(MatchDesc_like(FirstPhrase(DESC) & "*"))
        
        If temp_row = -1 Then
            GoTo NO_MATCH
        End If
        
        If comp_track = -1 Then ' desc < test_desc - send to bottom of category
            MatchDesc = get_mark_num(temp_row - 1 + VB_MASTER.Cells(temp_row, VB_MASTER.CategoryColumn()).MergeArea.count)
        ElseIf comp_track = 1 Then ' desc > test_desc - send to top of category
            MatchDesc = get_mark_num(temp_row)
        Else
            MatchDesc = -1
        End If
    Else
NO_MATCH:
        MatchDesc = -1
    End If
End Function

Public Function get_col_num(ByVal DESC As String, Optional srcSheet As Worksheet) As Integer
'Function: get_col_num - Retrieves the column number for the given column title/description.
'Arguments: desc - String contains the column title
'Returns: Integer containing the column number


    If srcSheet Is Nothing Then
        Set srcSheet = VB_MASTER
    End If
    
    get_col_num = -1

    If Len(DESC) = 0 Then Exit Function

    Dim col As Integer
    'Find the column number which has the 'desc' title
    For col = 1 To srcSheet.UsedRange.Columns.count
        If CStr(srcSheet.Cells(2, col).Value2) = DESC Then
            get_col_num = col
            Exit Function
        End If
    Next col
End Function

Public Sub set_property(ByVal mark_num As Integer, ByVal col_num As Integer, ByVal property_value)
'Function: set_property - Modifies a line item in MASTER with the specified mark number, under
'                         the specified column number to the specified value.
'Arguments: mark_num - Integer specifying the unique mark number
'           col_num - Integer specifying the column number for the correct column title. get_col_num() is useful here.
'           property_value - The property value.


    Dim row As Integer
    row = get_row(mark_num)
    
    If row > 0 Then
        VB_MASTER.Cells(row, col_num).Value = property_value
    End If
End Sub

Public Sub set_formula(ByVal row_num As Integer, ByVal col_num As Integer, ByVal in_formula As String)
'Function: set_property - Modifies a line item in MASTER with the specified mark number, under
'                         the specified column number to the specified value.
'Arguments: row_num - Integer specifying the row number; see get_row()
'           col_num - Integer specifying the column number for the correct column title. get_col_num() is useful here.
'           property_value - The property value.
    
    VB_MASTER.Cells(row_num, col_num).Formula = in_formula
End Sub

Public Function get_category(ByVal row As Integer, Optional srcSheet As Worksheet) As String
'Function: get_category - Based on the row number, will return the corresponding category section.
'                         WARNING: For row insertion, this is not a good judge of category, the category
'                                  must be specified by other means. this function can not distinguish
'                                  between row insertion before the first row of a category or after the last row of a category.
'Arguments: row - Integer containing the row number in question
'Returns: String containing the category name. If there is no category, returns vbNullString


    If srcSheet Is Nothing Then
        Set srcSheet = VB_MASTER
    End If

    get_category = vbNullString
    
    On Error GoTo F_END
    If Not IsEmpty(srcSheet.Cells(row, VB_MASTER.CategoryColumn()).MergeArea.Cells(1, 1)) Then
        get_category = srcSheet.Cells(row, VB_MASTER.CategoryColumn()).MergeArea.Cells(1, 1).Value2
    End If

F_END:
End Function

Public Sub ClientDeleteItem(ByVal Target As Range, Optional ByVal reason As String = vbNullString, _
    Optional ByVal write_change As Boolean = True, Optional ByVal inUser As String, Optional ByVal inDate As String)
    
'Subroutine: ClientDeleteItem - Handles a client delete operation. If material has already been ordered for an item that is now removed from
'                               the model, use this Delete rather than purge. The item and RFP records stay intact, but the quantities are removed from the model space.
'Arguments: Target - Range reference to cell in "Delete?" column where the Delete operation applies.
'           reason - String containing the delete reason. If this has 0 length, this procedure will load a user form to get a reason from the user.
'           write_change - Boolean deciding whether or not this procedure will call WriteChange. If instance from WriteChange, write_change should be False.


    Dim ui_change As Boolean

    If get_ColTitle(Target.column) = "Delete?" And get_mark_num(Target.row) > 0 Then
        ui_change = RenderUI(False)
        
        If IsEmpty(Target) Then ' Then proceed with the Delete action
            Target.Value = "x"
            
            'add comment
            On Error Resume Next
            Target.Comment.Delete
            Err.Clear
            On Error GoTo 0
            
            If Len(inUser) = 0 Then
                inUser = Environ("USERNAME")
            End If
            If Len(inDate) = 0 Then
                inDate = Date
            End If
            Target.AddComment "Deleted by " & inUser & " on " & inDate & "."
            Target.Comment.Shape.TextFrame.AutoSize = True
            
            'Handle formatting..
            '' conditional formatting; grey text, light grey background, strikethrough.
            '' add user comment in RED and NOT strikethrough
            With VB_MASTER.Range(Target.Address & ":$" & _
                ColLet(get_col_num("Checked Quantities") + GetSitesRange().count - 1) & "$" & Target.row).Font
            
                .Strikethrough = True
            End With
            
            ' load form the get user comment.
            
            If DeleteItem_Explanation.SetReason(Target.row, reason) Then
                ' successful
                If write_change Then VB_MASTER.WriteChange Target, "", True
            Else
                ' unsuccessful; rollback.
                Target.ClearContents
                On Error Resume Next
                Target.Comment.Delete
                Err.Clear
                On Error GoTo 0
                
                With VB_MASTER.Range(Target.Address & ":$" & _
                    ColLet(get_col_num("Checked Quantities") + GetSitesRange().count - 1) & "$" & Target.row).Font
                
                    .Strikethrough = False
                End With
            End If
            
        End If
        
        If ui_change Then RenderUI True
    End If
End Sub

Public Sub ClientUndeleteItem(ByVal Target As Range, Optional ByVal write_change As Boolean = True)
'Subroutine: ClientUndeleteItem - Handles an undo client delete operation. If material has already been ordered for an item that is now removed from
'                                 the model, use this Delete rather than purge. The item and RFP records stay intact, but the quantities are removed from the model space.
'                                 This procedure undoes the delete and removes the delete reason.
'Arguments: Target - Range reference to cell in "Delete?" column where the Delete operation applies.
'           write_change - Boolean deciding whether or not this procedure will call WriteChange. If instance from WriteChange, write_change should be False.


    Dim ui_change As Boolean
    
    If get_ColTitle(Target.column) = "Delete?" And get_mark_num(Target.row) > 0 Then
        
        If Not IsEmpty(Target) Then ' Then proceed with the Undelete Action
            ui_change = RenderUI(False)
        
            Target.ClearContents
            If write_change Then VB_MASTER.WriteChange Target, "x", True
            
            On Error Resume Next
            Target.Comment.Delete
            Err.Clear
            On Error GoTo 0
            
            'handle formatting
            Dim desc_col As Integer
            desc_col = get_col_num("Long Description")
            
            ' remove strikethrough
            With VB_MASTER.Range(Target.Address & ":$" & _
                ColLet(get_col_num("Checked Quantities") + GetSitesRange().count - 1) & "$" & Target.row).Font
            
                .Strikethrough = False
            End With
            
            Dim mark_num As Integer
            mark_num = get_mark_num(Target.row)
            
            Dim DESC As String
            DESC = get_property(mark_num, "Long Description")
            
            Dim delete_reason As String
            delete_reason = GetDeleteReason(mark_num)
            
            ' remove user comment from description.
            DESC = Left(DESC, Len(DESC) - Len(delete_reason) - 1)
            Call set_property(mark_num, desc_col, DESC)
        End If
        
        If ui_change Then RenderUI True
    End If
End Sub

Public Function GetDeleteReason(ByVal mark_num As Integer) As String
'Function: GetDeleteReason - Gets the delete reason from a item description. If not deleted, returns ""
'Arguments: mark_num - mark number for item to get delete reason.
'Returns: String containing the delete reason. If not deleted, returns ""


    GetDeleteReason = ""
    
    Dim DESC As String
    DESC = get_property(mark_num, "Long Description")
    
    Dim desc_col As Integer
    desc_col = get_col_num("Long Description")
    
    Dim master_row As Integer
    master_row = get_row(mark_num)
    
    If master_row = -1 Then
        Exit Function
    End If
    
    Dim c As Integer
    ' find start position of user comment (in RED)
    For c = Len(DESC) To 1 Step -1
        With VB_MASTER.Cells(master_row, desc_col).Characters(start:=c, length:=1).Font
            If .ColorIndex <> 3 Then
                Exit For
            End If
        End With
    Next c
    
    If c < Len(DESC) Then
        ' remove user comment from description.
        GetDeleteReason = Right(DESC, Len(DESC) - c - 1)
    End If

End Function

Public Sub SetDeleteReason(ByVal mark_num As Integer, ByVal reason As String)
'Subroutine: SetDeleteReason - Set a delete reason in description. Should only be called if item is being deleted.
'Arguments: mark_num - mark number for item being deleted
'           reason - Delete Reason to set. If reason has 0 length. This procedure does nothing

    
    If Len(reason) = 0 Then
        Exit Sub
    End If
    
    Dim desc_col As Integer
    desc_col = get_col_num("Long Description")
    
    Dim DESC As String
    DESC = get_property(mark_num, "Long Description")
    
    If IsDeleted(mark_num) Then ' remove user comment from description.
        Dim org_reason As String
        org_reason = GetDeleteReason(mark_num)
        If Len(org_reason) > 0 Then
            DESC = Left(DESC, Len(DESC) - Len(org_reason) - 1)
        End If
    End If

    'set reason
    Call set_property(mark_num, desc_col, DESC & " " & reason)
    
    Dim mst_row As Integer
    mst_row = get_row(mark_num)
    
    ' set reason formatting to strikethrough=False and color=red
    With VB_MASTER.Cells(mst_row, desc_col).Characters(start:=Len(VB_MASTER.Cells(mst_row, desc_col).Value2) - Len(reason), length:=Len(reason) + 1).Font
        .ColorIndex = 3
        .Strikethrough = False
    End With
End Sub

Public Sub ApproveDescription(Optional ByVal Target As Range, Optional ByVal MarkNum As Integer, _
    Optional ByVal write_change As Boolean = True, Optional ByVal inComment As String)
    
'Subroutine: ApproveDescription - Handles description approval operations. Initiated by double clicking the "Description check" column, making
'                                 a cell in the Description check column not empty, or in the EditItemWindow.
'Arguments: Target - Range reference to cell in "Description check" column where the operation applies. Cannot be used in conjunction with MarkNum
'           MarkNum - Integer containing the mark number for the item to approve. Cannot be used in conjunction with Target
'           write_change - Boolean deciding whether or not this procedure will call WriteChange. If this instance from WriteChange, write_change should be False.


    Dim ui_change As Boolean
    
    If MarkNum < 1 And Not IsNull(Target) Then
        If get_ColTitle(Target.column) = "Description Check" And get_mark_num(Target.row) > 0 Then
            ui_change = RenderUI(False)
            
            'write check mark
            If IsEmpty(Target) Then
                Target.Font.Name = "Marlett"
                Target.Value = "a"
            End If
            
            'add comment
            On Error Resume Next
            Target.Comment.Delete
            Err.Clear
            On Error GoTo 0
            
            If Len(inComment) = 0 Then
                inComment = "Description Approved by " & Environ("USERNAME") & " on " & Date & "."
            End If
            Target.AddComment inComment
            Target.Comment.Shape.TextFrame.AutoSize = True
            
            If write_change Then VB_MASTER.WriteChange Target, ""
            
            If ui_change Then RenderUI True
        End If
    
    ElseIf get_row(MarkNum) > 0 And Target Is Nothing Then
        ui_change = RenderUI(False)
        
        Set Target = VB_MASTER.Cells(get_row(MarkNum), get_col_num("Description Check"))
        
        'write check mark
        If IsEmpty(Target) Then
            Target.Font.Name = "Marlett"
            Target.Value = "a"
            If write_change Then VB_MASTER.WriteChange Target, ""
        End If
        
        'add comment
        On Error Resume Next
        Target.Comment.Delete
        Err.Clear
        On Error GoTo 0
        
        If Len(inComment) = 0 Then
            inComment = "Description Approved by " & Environ("USERNAME") & " on " & Date & "."
        End If
        Target.AddComment inComment
        Target.Comment.Shape.TextFrame.AutoSize = True
        
        If ui_change Then RenderUI True
    
    End If
End Sub

Public Sub UnapproveDescription(Optional ByVal Target As Range, Optional ByVal MarkNum As Integer, Optional ByVal write_change As Boolean = True)
'Subroutine: UnapproveDescription - Handles undo description approval operations. Initiated by double clicking the "Description check" column, making
'                                   a cell in the Description check column empty, or in the EditItemWindow.
'Arguments: Target - Range reference to cell in "Description check" column where the operation applies. Cannot be used in conjunction with MarkNum
'           MarkNum - Integer containing the mark number for the item to un-approve. Cannot be used in conjunction with Target
'           write_change - Boolean deciding whether or not this procedure will call WriteChange. If instance from WriteChange, write_change should be False.


    Dim ui_change As Boolean
    Dim row As Integer
    
    If MarkNum < 1 And Not IsNull(Target) Then
        If get_ColTitle(Target.column) = "Description Check" And get_mark_num(Target.row) > 0 Then
            ui_change = RenderUI(False)
            
            'empty cell
            If Not IsEmpty(Target) Then
                Target.ClearContents
                If write_change Then VB_MASTER.WriteChange Target, "a"
            End If
            
            'add comment
            On Error Resume Next
            Target.Comment.Delete
            Err.Clear
            On Error GoTo 0
            
            If ui_change Then RenderUI True
        End If
        
    ElseIf get_row(MarkNum) > 0 And Target Is Nothing Then
        
        ui_change = RenderUI(False)
        
        Set Target = VB_MASTER.Cells(get_row(MarkNum), get_col_num("Description Check"))
        
        'empty cell
        If Not IsEmpty(Target) Then
            Target.ClearContents
            If write_change Then VB_MASTER.WriteChange Target, "a"
        End If
        
        'add comment
        On Error Resume Next
        Target.Comment.Delete
        Err.Clear
        On Error GoTo 0
        
        If ui_change Then RenderUI True
    
    End If
End Sub

Public Function NewOrderColumn(ByVal order_col As Integer, ByVal order_num As String, ByVal order_site As String, ByVal order_date As String) As Integer
'Function: NewOrderColumn - Handles inserting and formatting a column for a new order.
'Arguments: order_col - Integer containing the column number where a new order could go.
'           order_num - String containing the new order number
'           order_site - String containing the new order site
'           order_date - String containing the new order date
'Returns: Integer containing the column for the new order


    Dim ui_change As Boolean
    ui_change = RenderUI(False)

    If IsEmpty(VB_MASTER.Cells(VB_MASTER.SubtitleRow(), order_col - 1)) And IsEmpty(VB_MASTER.Cells(VB_MASTER.SubtitleRow2(), order_col - 1)) Then
        'do not insert because we are already dealing with an empty column
        order_col = order_col - 1
        'Call InsertMasterColumn(new_order_col, True)
    Else
        InsertMasterColumn order_col
    End If
    
    'Title/Format Column
    VB_MASTER.Cells(VB_MASTER.SubtitleRow(), order_col).Value = order_num  'Order Number
    If order_site = "PROJECT" Then
        VB_MASTER.Cells(VB_MASTER.SubtitleRow2(), order_col).Value = order_date 'Order Date
    Else
        VB_MASTER.Cells(VB_MASTER.SubtitleRow2(), order_col).Value = order_site
    End If
    
    VB_MASTER.Columns(order_col).AutoFit
    If VB_MASTER.Cells(VB_MASTER.SubtitleRow(), order_col).ColumnWidth < 8.86 Then
        VB_MASTER.Cells(VB_MASTER.SubtitleRow(), order_col).ColumnWidth = 8.86
    End If
    
    With VB_MASTER.Range(ColLet(order_col) & VB_MASTER.SubtitleRow() & ":" & ColLet(order_col) & VB_MASTER.SubtitleRow2()).Font
        .Bold = False
        .Italic = True
    End With
    
    'Set "Total Ordered"
    VB_ORDER_LOG.SetTotalOrderedFormula

    If ui_change Then RenderUI True

    NewOrderColumn = order_col

End Function

Public Sub DeleteMasterRow(ByVal delete_row As Integer)
'Function: DeleteMasterRow - Handles row deletion on the Master BOM Sheet
'Arguments: delete_row - Integer containing the row to delete


    Dim removed_row As Boolean
    removed_row = False
    
    Dim category As String
    category = get_category(delete_row)

    Dim ui_change As Boolean
    ui_change = RenderUI(False)

    If category = vbNullString Then
        VB_MASTER.Rows(delete_row).Delete
    Else
        Dim cat_start As Integer
        Dim cat_end As Integer
        cat_start = get_cat_row(category)
        cat_end = cat_start + VB_MASTER.Cells(delete_row, VB_MASTER.CategoryColumn()).MergeArea.count - 1
        
        If cat_start = cat_end Then 'only item in category
            Dim mark As Integer
            mark = get_mark_num(delete_row)
            
            If IsEmpty(VB_MASTER.Cells(delete_row, get_col_num("Long Description"))) Then
                'deleteing empty category - this should be handled elsewhere, no double dipping.
            Else
                ' clear row
                Dim clear_col As Integer
                For clear_col = get_col_num("Orders") To LastColumn()
                    If Not (get_ColTitle(clear_col) = "Unit" Or get_ColTitle(clear_col) = "Mark No." Or get_ColTitle(clear_col) = "Unique ID") _
                        And Not VB_MASTER.Cells(delete_row, clear_col).HasFormula Then
                        
                        VB_MASTER.Cells(delete_row, clear_col).ClearContents
                    End If
                Next clear_col
                
                'delete comments if any
                On Error Resume Next
                VB_MASTER.Cells(delete_row, get_col_num("Delete?")).Comment.Delete
                VB_MASTER.Cells(delete_row, get_col_num("Description Check")).Comment.Delete
                Err.Clear
                On Error GoTo 0
                
                VB_MASTER.Cells(delete_row, get_col_num("Mark No.")).Value = mark
            End If
        ElseIf delete_row = cat_start Then
            VB_MASTER.Rows(delete_row).Delete
            removed_row = True
            
            VB_MASTER.Cells(delete_row, VB_MASTER.CategoryColumn()).Value = category
            
            With VB_MASTER.Rows(cat_start & ":" & cat_end).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThick
                .ColorIndex = 1 'black
            End With
        ElseIf delete_row = cat_end Then
            VB_MASTER.Rows(delete_row).Delete
            removed_row = True
            
            With VB_MASTER.Rows(cat_start & ":" & cat_end - 1).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThick
                .ColorIndex = 1 'black
            End With
        Else
            VB_MASTER.Rows(delete_row).Delete
            removed_row = True
            
            With VB_MASTER.Rows(delete_row).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = 1 'black
            End With
            
        End If
        
    End If
    
    VB_CATEGORY.CategoryAutoFit category
    
    VB_MASTER.ResetScrollArea
    VB_MASTER.UnlockColumns
    
    If removed_row Then
        Call VB_CHANGE_LOG.offset(VB_MASTER.Name, delete_row, get_col_num("Mark No."), -1, 0)
        ResetItemDescRange
    End If
    
    If ui_change Then RenderUI True
End Sub

Public Function SetNewMarkNumber(ByVal new_row As Integer) As Integer
'Function: SetNewMarkNumber - Sets unique mark number and Unique ID for the new_row.
'Arguments: new_row - Integer containing the row number for the new mark number and Unique ID
'Returns: Integer containing the mark number

    
    'Set Mark Number/Unit
    '''''''''''''''''''''''''''''''''
    Dim mark_num As Integer
    SetNewMarkNumber = get_mark_num(new_row - 1)

    Do
        SetNewMarkNumber = SetNewMarkNumber + 1
    Loop While (get_row(SetNewMarkNumber) <> -1 Or SetNewMarkNumber <= 0)
    
    VB_MASTER.Cells(new_row, get_col_num("Mark No.")).Value = SetNewMarkNumber

    
    'Set Unique ID
    '''''''''''''''''''''''''''''''''
    Dim uIDcol As Integer
    uIDcol = get_col_num("Unique ID")
    
    Dim uniqueID As Integer
    If new_row <> VB_MASTER.FirstRow() Then
        If Len(VB_MASTER.Cells(new_row - 1, uIDcol).Value2) > 0 Then
            uniqueID = CInt(Right(VB_MASTER.Cells(new_row - 1, uIDcol).Value2, 6))
        Else
            uniqueID = 1
        End If
    Else
        uniqueID = 1
    End If

    Do
        uniqueID = uniqueID + 1
    Loop While VB_UNIQUEID.UniqueIDExists("A" & Format(uniqueID, "000000"))
    
    VB_UNIQUEID.AddUniqueID "A" & Format(uniqueID, "000000")
    Call set_property(SetNewMarkNumber, uIDcol, "A" & Format(uniqueID, "000000"))
    
End Function

Public Sub InsertMasterRow(ByVal category As String, ByRef ins_row_num As Integer, Optional ByRef out_mark_num As Integer)
'Function: InsertMasterRow - Handles row insertion on the Master BOM Sheet
'Arguments: category - String containing the category in which the row should go
'           ins_row_num - the row number on which to apply the insert command
'           out_mark_number - Return argument for the mark number on the new row, no valid relevant mark number, returns -1

    
    If Len(get_category(ins_row_num)) = 0 Then
        If Len(get_category(ins_row_num - 1)) = 0 Then 'no category, empty row/Out of range
            VB_MASTER.Rows(ins_row_num).Insert CopyOrigin:=xlFormatFromRightOrBelow
            
            With VB_MASTER.Cells(ins_row_num, VB_MASTER.CategoryColumn()).EntireRow.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = 1 'black
            End With
            
            With VB_MASTER.Cells(ins_row_num, VB_MASTER.CategoryColumn()).EntireRow.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = 1 'black
            End With
            
            out_mark_num = -1
            Exit Sub
        Else
            'Insert after last row in MASTER.
            category = get_category(ins_row_num - 1)
        End If
    End If
    
    Dim mark_col As Integer
    mark_col = get_col_num("Mark No.")
    
    Dim org_cat_row As Integer
    org_cat_row = get_cat_row(category)
    
    'if empty category, do nothing
    If Len(get_property(get_mark_num(org_cat_row), "Long Description")) = 0 Then
        Exit Sub
    End If
    
    'the unit will be the the same for all items in the category, so get it now
    Dim unit As String
    unit = get_property(get_mark_num(org_cat_row), "Unit")
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    'Insert Row
    If ins_row_num = VB_MASTER.FirstRow() Then
        VB_MASTER.Rows(ins_row_num).Insert CopyOrigin:=xlFormatFromRightOrBelow
    Else
        VB_MASTER.Rows(ins_row_num).Insert CopyOrigin:=xlFormatFromLeftOrAbove
    End If
    Call VB_CHANGE_LOG.offset(VB_MASTER.Name, ins_row_num, mark_col, 1, 0)
    
    VB_MASTER.Rows(ins_row_num).Font.Strikethrough = False
    
    Dim c As Integer
    For c = 1 To LastColumn()
        With VB_MASTER.Cells(ins_row_num, c).Borders(xlEdgeRight)
            If .LineStyle = xlLineStyleNone Then
                .LineStyle = xlContinuous
                .Weight = xlThin
            End If
        End With
    Next c
    
    VB_MASTER.Unprotect
    
    If ins_row_num = org_cat_row Then 'top item in category
        VB_MASTER.Range("$" & ColLet(VB_MASTER.CategoryColumn()) & "$" & ins_row_num & ":$" & _
            ColLet(VB_MASTER.CategoryColumn()) & "$" & ins_row_num + VB_MASTER.Cells(ins_row_num + 1, VB_MASTER.CategoryColumn()).MergeArea.count).Merge
    
        With VB_MASTER.Rows(ins_row_num).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThick
        End With
        
        With VB_MASTER.Rows(ins_row_num).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        
    ElseIf ins_row_num = org_cat_row + VB_MASTER.Cells(org_cat_row, VB_MASTER.CategoryColumn()).MergeArea.count Then 'bottom Item in category
        
        VB_MASTER.Range(ColLet(VB_MASTER.CategoryColumn()) & org_cat_row & ":" & ColLet(VB_MASTER.CategoryColumn()) & ins_row_num).Merge
    
        With VB_MASTER.Rows(ins_row_num).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 1 'black
        End With
        
        With VB_MASTER.Rows(ins_row_num).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = 1 'black
        End With
        
    Else ' normal case
        With VB_MASTER.Rows(ins_row_num).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 1 'black
        End With
        
        With VB_MASTER.Rows(ins_row_num).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 1 'black
        End With
    End If
    
    VB_MASTER.SetProtection
    
    With VB_MASTER.Cells(get_cat_row(category), VB_MASTER.CategoryColumn()).MergeArea.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    With VB_MASTER.Cells(get_cat_row(category), VB_MASTER.CategoryColumn()).MergeArea.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    
    VB_CATEGORY.CategoryAutoFit category
    
    'Set Mark Number/Unique ID/Unit
    '''''''''''''''''''''''''''''''''
    Dim mark_num As Integer
    mark_num = SetNewMarkNumber(ins_row_num)
    
    Call set_property(mark_num, get_col_num("Unit"), unit)
    '''''''''''''''''''''''''''''''''
    
    ResetItemDescRange
    VB_MASTER.ResetScrollArea
    VB_MASTER.UnlockColumns
    
    SetRowFormulas mark_num
    
    If ui_change Then RenderUI True
    
    out_mark_num = mark_num
End Sub

Public Sub SetRowFormulas(ByVal mark_num As Integer)
'Subroutine: SetRowFormulas - Sets row formulas for get_row(mark_num)
'Arguments: mark_num - mark number for row to set formulas.


    If mark_num < 1 Then
        Exit Sub
    End If

    'Set row formulas
    Dim this_row As Integer
    this_row = get_row(mark_num)
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    'Set "Total Ordered"
    VB_ORDER_LOG.SetTotalOrderedFormula this_row
    
    'Set "Model Total" and "Total Extras"
    Dim mt_col As Integer
    Dim te_col As Integer
    
    mt_col = get_col_num("Model Total")
    te_col = get_col_num("Total Extras")
    
    Dim te_form As String
    Dim mt_form As String
    
    te_form = "="
    mt_form = "="
    
    Dim rSites As Range
    Set rSites = GetSitesRange()
    Dim vbCell
    Dim count As Integer
    count = 0
    
    For Each vbCell In rSites
        If Not IsFabPackage(vbCell.Value2) And Not IsEmpty(vbCell.Value2) Then
            If count = 0 Then
                mt_form = mt_form & ColLet(vbCell.column) & this_row
                te_form = te_form & ColLet(vbCell.column + rSites.count) & this_row
            Else
                mt_form = mt_form & "+" & ColLet(vbCell.column) & this_row
                te_form = te_form & "+" & ColLet(vbCell.column + rSites.count) & this_row
            End If
            count = count + 1
        End If
    Next vbCell
    Set rSites = Nothing
    
    If count = 0 Then
        mt_form = mt_form & "0"
        te_form = te_form & "0"
    ElseIf VB_MASTER.Columns(get_col_num("Total Extras")).Hidden Then
        te_form = "=0"
    End If
    
    Call set_formula(this_row, mt_col, mt_form)
    Call set_formula(this_row, te_col, te_form)
    
    'Set "Quantity Needed"
    Dim qn_col As Integer
    qn_col = get_col_num("Quantity Needed")
    Dim sur_col As Integer
    sur_col = get_col_num("Client Inventory")
    Dim to_col As Integer
    to_col = get_col_num("Total Ordered")
    
    Dim sp_form As String
    sp_form = "=IF(ISBLANK("
    sp_form = sp_form & ColLet(get_col_num("Delete?")) & this_row & "),"
    sp_form = sp_form & "(" & ColLet(mt_col) & this_row & "+" & ColLet(te_col) & this_row & ")-(" & ColLet(to_col) & this_row & "+" & ColLet(sur_col) & this_row & "),"
    sp_form = sp_form & "0-(" & ColLet(to_col) & this_row & "+" & ColLet(sur_col) & this_row & "))"
    
    Call set_formula(this_row, qn_col, sp_form)
    
    If ui_change Then RenderUI True
End Sub

Public Sub DeleteMasterColumn(ByVal delete_col As Integer)
'Subroutine: DeleteMasterColumn - Handles column deletion in VB_MASTER
'Arguments: delete_col - the column number in VB_MASTER to delete


    Dim col_title As String
    col_title = get_ColTitle(delete_col)
    If Not (col_title = "Orders" _
            Or col_title = "Current Model Quantities" _
            Or col_title = "Model Extras" _
            Or col_title = "Checked Quantities") Then
            
        MsgBox "Can't delete this column here.", vbCritical
        Exit Sub
    End If
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    If get_col_num(get_ColTitle(delete_col)) = delete_col Then
        ' Case where the delete column is the first of the lot
        
        Dim section As String
        Dim title As String
        section = VB_MASTER.Cells(VB_MASTER.SectionRow(), delete_col).Value2
        title = VB_MASTER.Cells(VB_MASTER.TitleRow(), delete_col).Value2
        
        VB_MASTER.Columns(delete_col).Delete
        
        VB_MASTER.Cells(VB_MASTER.SectionRow(), delete_col).Value = section
        VB_MASTER.Cells(VB_MASTER.TitleRow(), delete_col).Value = title
        
        With VB_MASTER.Cells(VB_MASTER.SectionRow(), delete_col).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .Color = RGB(150, 54, 52)
        End With
        With VB_MASTER.Cells(VB_MASTER.TitleRow(), delete_col).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .Color = RGB(150, 54, 52)
        End With
        With VB_MASTER.Columns(delete_col).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .Color = RGB(150, 54, 52)
        End With
        
    Else
        VB_MASTER.Columns(delete_col).Delete
    End If
    
    Dim first As Integer
    first = VB_MASTER.FirstRow()
    
    Call VB_CHANGE_LOG.offset(VB_MASTER.Name, first, delete_col, 0, -1)
    VB_MASTER.ResetScrollArea
    ResetItemDescRange
    VB_MASTER.UnlockColumns
    
    Dim last As Integer
    last = VB_MASTER.LastRow()
    Dim row As Integer
    For row = first To last
        SetRowFormulas get_mark_num(row)
    Next row
    
    If ui_change Then RenderUI True
End Sub

Public Sub InsertMasterColumn(ByVal insert_col As Integer, Optional ByVal formatOnly = False)
'Subroutine: InsertMasterColumn - Handles column insertion
'Arguments: insert_col - insert column number
'           formatOnly - in the case where new column already exists, refresh formatting by setting format_only True


    If Not (insert_col = get_col_num("Total Ordered") Or _
        insert_col = get_col_num("Model Extras") Or _
        insert_col = get_col_num("Total Extras") Or _
        insert_col = LastColumn() + 1) And Not formatOnly Then
        
        MsgBox "Can't insert columns here.", vbCritical
        Exit Sub
    End If
    
    Dim first As Integer
    Dim last As Integer
    first = VB_MASTER.FirstRow()
    last = VB_MASTER.LastRow()
    
    Dim position As Integer
    ' 1 - Orders insert
    ' 2 - Current Model Quantities
    ' 3 - Checked Quantities
    If insert_col = get_col_num("Total Ordered") Then
        position = 1
    ElseIf insert_col = get_col_num("Model Extras") Then
        position = 2
    ElseIf insert_col = get_col_num("Total Extras") Then
        position = 3
    ElseIf insert_col = LastColumn() + 1 Then
        position = 4
    End If
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    VB_MASTER.Unprotect
    If Not formatOnly Then
        VB_MASTER.Columns(insert_col).Insert
        Call VB_CHANGE_LOG.offset(VB_MASTER.Name, first, insert_col, 0, 1)
        VB_MASTER.ResetScrollArea
        ResetItemDescRange
        VB_MASTER.UnlockColumns
        
        With VB_MASTER.Columns(insert_col).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    
        VB_MASTER.Range(ColLet(get_col_num(get_ColTitle(insert_col - 1))) & VB_MASTER.SectionRow() & ":" & ColLet(insert_col) & VB_MASTER.SectionRow()).Merge
        VB_MASTER.Range(ColLet(get_col_num(get_ColTitle(insert_col - 1))) & VB_MASTER.TitleRow() & ":" & ColLet(insert_col) & VB_MASTER.TitleRow()).Merge
    End If
    VB_MASTER.SetProtection
    
    With VB_MASTER.Columns(insert_col).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    VB_MASTER.Range(ColLet(insert_col) & VB_MASTER.SubtitleRow() & ":" & ColLet(insert_col) & VB_MASTER.SubtitleRow2()).BorderAround _
        LineStyle:=xlContinuous, Weight:=xlMedium
        
    With VB_MASTER.Range(ColLet(insert_col) & VB_MASTER.SubtitleRow() & ":" & ColLet(insert_col) & VB_MASTER.SubtitleRow2()).Font
        .Bold = True
        .Italic = False
    End With
        
    With VB_MASTER.Cells(VB_MASTER.SectionRow(), insert_col).MergeArea.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With VB_MASTER.Cells(VB_MASTER.TitleRow(), insert_col).MergeArea.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With

    With VB_MASTER.Columns(insert_col + 1).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        
        If position = 1 Then 'Total Ordered
            .LineStyle = xlDouble
            
        ElseIf position = 2 Then ' Current Model Quantities
            .Weight = xlThick
            .Color = RGB(150, 54, 52)
            
        ElseIf position = 3 Then ' Model Extras
            .LineStyle = xlDouble
            
        ElseIf position = 4 Then ' Checked Quantities
            .Weight = xlThick
            .Color = RGB(150, 54, 52)
            
        End If
    End With

    'add bold lines between categories
    Dim row As Integer
    row = first
    Dim MergeInfo As Range
    Do While Not IsEmpty(VB_MASTER.Cells(row, VB_MASTER.CategoryColumn()).MergeArea) Or Not IsEmpty(VB_MASTER.Cells(row, VB_MASTER.CategoryColumn()))
        Set MergeInfo = VB_MASTER.Cells(row, VB_MASTER.CategoryColumn()).MergeArea
        
        With MergeInfo.EntireRow.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = 1 'black
        End With

        row = row + MergeInfo.Rows.count
    Loop
    Set MergeInfo = Nothing
    
    With VB_MASTER.Rows(row).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = 1 'black
    End With
    With VB_MASTER.Rows(row + 1).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = 1 'black
    End With
    
    For row = first To last
        SetRowFormulas get_mark_num(row)
    Next row
    
    ' check if collapsed,
    If VB_MASTER.Columns(insert_col).Hidden <> VB_MASTER.Columns(get_col_num(get_ColTitle(insert_col))).Hidden Then
        VB_MASTER.Columns(insert_col).Hidden = VB_MASTER.Columns(get_col_num(get_ColTitle(insert_col))).Hidden
    End If
    
    If ui_change Then RenderUI True
    
End Sub

Public Sub DeleteCategory(ByVal category As String)
'Subroutine: DeleteCategory - Handles category deletion. Will not delete category if valid items are in it.
'Arguments: category - name of category to delete


    If category = vbNullString Then
        Exit Sub
    End If
    
    If VB_MASTER.FirstRow() = VB_MASTER.LastRow() Then
        'there is only one category - do not delete
        Exit Sub
    End If
    
    Dim cat_start As Integer
    Dim cat_end As Integer
    cat_start = get_cat_row(category)
    cat_end = cat_start + VB_MASTER.Cells(cat_start, VB_MASTER.CategoryColumn()).MergeArea.count - 1
    
    If cat_start = cat_end And IsEmpty(VB_MASTER.Cells(cat_start, get_col_num("Long Description"))) Then
        'empty category, delete
        
        Dim ui_change As Boolean
        ui_change = RenderUI(False)
        
        VB_MASTER.Rows(cat_start).Delete
        
        VB_MASTER.ResetScrollArea
        ResetItemDescRange
        
        Call VB_CHANGE_LOG.offset(VB_MASTER.Name, cat_start, VB_MASTER.CategoryColumn(), -1, 0)
        
        With VB_MASTER.Rows(cat_start).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = 1 'black
        End With
        
        With VB_MASTER.Rows(VB_MASTER.LastRow()).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThick
        End With

        Dim unit As String
        unit = VB_CATEGORY.GetCategoryUnit(category)
        
        'Delete from Category Key Table
        VB_CATEGORY.DeleteCategory category
        
        If ui_change Then RenderUI True
        
        If VB_CHANGE_LOG.TrackChanges() Then
            If VB_CHANGE_LOG.LogChange(VB_MASTER.Name & "!$" & ColLet(VB_MASTER.CategoryColumn()) & "$" & cat_start, "", "Purged Category: " & _
                category & " (" & unit & ")", "", "") <> 0 Then
                'throw error
                ErrorHandling "DeleteCategory", 10, "LogChange(" & VB_MASTER.Name & "!$" & ColLet(VB_MASTER.CategoryColumn()) & "$" & cat_start & ", , Purged Category: " & _
                    category & " (" & unit & "), , )", 1
            End If
        End If
    Else
        MsgBox "This category contains material. Will not delete.", vbExclamation
    End If
    
End Sub

Public Function NewCategoryHandler() As String
'Function: NewCategoryHandler - Handles category addition.
'Returns: String containing the new category name


    Dim new_cat_row As Integer
    new_cat_row = VB_MASTER.LastRow() + 1
    
    If new_cat_row = -1 Then
        Exit Function
    End If
    
    Dim new_cat_name As String
    NewCategoryForm.LoadForm new_cat_name, new_cat_row
    
    NewCategoryHandler = new_cat_name
End Function

Public Function NewCategory(ByVal new_cat_name As String, ByVal cat_unit As String, ByVal new_cat_row As Integer) As Integer
'Function: NewCategory - Add new category to VB_MASTER.
'Arguments: new_cat_name - name of new category. If category already exists, does nothing.
'           cat_unit - the category's unit
'           new_cat_row - the insertion row for the new category.
'Returns: Integer containing error code. 0 - no error.


    If new_cat_row <= 0 Or Len(new_cat_name) = 0 Or Len(cat_unit) = 0 Then
        ErrorHandling "NewCategory", 10, "Invalid Arguments:: new_cat_name:=" & Chr(34) & new_cat_name & Chr(34) & _
            ", cat_unit:=" & Chr(34) & cat_unit & Chr(34) & ", new_cat_row:=" & new_cat_row, 1
        NewCategory = -1
        Exit Function
    End If
    
    If get_cat_row(new_cat_name) <> -1 Then
        MsgBox "Category (" & new_cat_name & ") already exists.", vbExclamation
        NewCategory = -1
        Exit Function
    End If
    
    If new_cat_row <> VB_MASTER.LastRow() + 1 Then
        new_cat_row = get_cat_row(get_category(new_cat_row))
    End If
    If new_cat_row = -1 Then
        Exit Function
    End If
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    If new_cat_row = VB_MASTER.FirstRow() Then
        VB_MASTER.Rows(new_cat_row).Insert CopyOrigin:=xlFormatFromRightOrBelow
    Else
        VB_MASTER.Rows(new_cat_row).Insert CopyOrigin:=xlFormatFromLeftOrAbove
    End If
    
    With VB_MASTER.Rows(new_cat_row).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = 1 'black
    End With
    
    With VB_MASTER.Rows(new_cat_row).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = 1 'black
    End With
    
    Dim c As Integer
    For c = 1 To LastColumn()
        With VB_MASTER.Cells(new_cat_row, c).Borders(xlEdgeRight)
            If .LineStyle = xlLineStyleNone Then
                .LineStyle = xlContinuous
                .Weight = xlThin
            End If
        End With
    Next c
    
    Call VB_CHANGE_LOG.offset(VB_MASTER.Name, new_cat_row, get_col_num("Mark No."), 1, 0)
    
    VB_MASTER.Cells(new_cat_row, VB_MASTER.CategoryColumn()).Value = new_cat_name
    
    'Set Mark Number/Unique ID/Unit
    '''''''''''''''''''''''''''''''''
    Dim mark_num As Integer
    mark_num = SetNewMarkNumber(new_cat_row)
    
    Call set_property(mark_num, get_col_num("Unit"), cat_unit)
    '''''''''''''''''''''''''''''''''
    
    ResetItemDescRange
    VB_MASTER.ResetScrollArea
    
    'reformat
    VB_MASTER.CalculateQuantityFormat
    
    SetRowFormulas mark_num
    
    'add category to the matrix
    VB_CATEGORY.AddCategoryKey new_cat_name, cat_unit, new_cat_row
    
    If ui_change Then RenderUI True
    
    'LOG CHANGE
    If VB_CHANGE_LOG.TrackChanges() Then
        If VB_CHANGE_LOG.LogChange(VB_MASTER.Name & "!$" & ColLet(get_col_num("Long Description")) & "$" & new_cat_row, "", "Added Category: " & new_cat_name & " (" & cat_unit & ")", "", "") <> 0 Then
            'throw error
            ErrorHandling "NewCategory", 0, "LogChange() operation failed:" & vbCrLf & _
                "LogChange(" & VB_MASTER.Name & "!$" & ColLet(get_col_num("Long Description")) & "$" & new_cat_row & ", , Added Category: " & new_cat_name & ", , )" & vbCrLf & _
                "Continue [OK] or Cancel procedure?", 1
        End If
    End If
    
    NewCategory = 0
End Function

Public Sub ResetItemDescRange()
'Subroutine: ResetItemDescRange - For item row insertion/deletion, this procedure refreshes the ranges of descriptions for
'                                 data valaidation elsewhere in the application.

    Dim rList As String
    Dim descs As Name
    Set descs = ThisWorkbook.Names.Item("item_descriptions")
    
    rList = VB_MASTER.Name & "!$" & ColLet(get_col_num("Long Description")) & "$" & VB_MASTER.FirstRow() & _
        ":$" & ColLet(get_col_num("Long Description")) & "$" & VB_MASTER.LastRow()
        
    descs.RefersTo = "=" & rList
    
    Set descs = Nothing
End Sub

Public Function LastColumn() As Integer
'Function: LastColumn - Calculates the last column in the Master BOM.
'Returns: Integer containing the last column number


    LastColumn = get_col_num("Checked Quantities")
    LastColumn = LastColumn + VB_MASTER.Cells(VB_MASTER.TitleRow(), LastColumn).MergeArea.count - 1
End Function
