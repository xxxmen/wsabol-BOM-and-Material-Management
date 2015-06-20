VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VB_MASTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'multi is True/False identifier defining whether or not multiple cells have been selected
Private multi As Boolean
                            
Dim vOldValR() As Range  'Vector of all selected cell ranges
Dim vOldVal()            'Vector of all all selected cell values

Private Const category_col = 1
Private Const section_row = 1
Private Const title_row = 2
Private Const subtitle_row = 3
Private Const subtitle2_row = 4
Private Const first_row = 5

Private Sub Assign_Range_Values(ByVal Target As Range)
'Subroutine: Assign_Range_Values - In a multiple selection case, this subroutine assigns the vOldVal(), vOldValR() values from Target
'Arguments: target - Range of selected values from Excel's event handling
    
    
    Dim lRange As Range
    Dim i As Integer
    i = 0
    'Assigns the vOldVal(), vOldValR() values from Target
    For Each lRange In Target.Cells
        Set vOldValR(i) = lRange
        vOldVal(i) = lRange.Value2
        i = i + 1
    Next lRange
    
    Set lRange = Nothing
    Set Target = Nothing
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'Subroutine: Worksheet_SelectionChange - Excel event handler. This Subroutine is called everytime the active selection
'                                        in MASTER changes. Determines if there is one or multiple cells selected,
'                                        and defines vOldVal() and vOldValR. Assumes >1000 cells is a whole column
'                                        and ignores that selection.
'Arguments: Target - Range object containing all selected cells


    Dim Action As Integer
    
    'Check the Target is a cell reference
    If Not IsCellLocation(Target.Address) Then
        ErrorHandling "VB_MASTER.Worksheet_SelectionChange", 10, "Target Range is not a valid cell reference." & vbCrLf & _
            "target = " & Target.Address & "Terminating Process.", Action
        Exit Sub
    End If
    
    Dim last_mark As Integer
    last_mark = LastMark()
    
    '' HIDE COMMENTS
    Dim cmt As Comment
    For Each cmt In VB_MASTER.Comments
        cmt.Visible = False
    Next cmt
    
    'The following code could be used to change the active selection. Disable events so that this sub is not called again and again.
    '''''''''''''''''''''''''''''''''''''''''''''''
    Dim event_change As Boolean
    event_change = Application.EnableEvents
    If event_change Then Application.EnableEvents = False
    
    Dim max_row As Integer
    Dim min_row As Integer
    Dim max_column As Integer
    Dim min_column As Integer
    max_row = MaxRow(Target)
    min_row = MinRow(Target)
    max_column = MaxColumn(Target)
    min_column = MinColumn(Target)
    
    With Target
        If min_row < first_row Then
            If max_row < first_row Then
                If .Cells(1, 1).Interior.ColorIndex = 1 Then ' black
                    VB_MASTER.Range("$" & ColLet(min_column) & first_row & ":$" & ColLet(max_column) & first_row).Select
                Else
                    VB_MASTER.Range("$" & ColLet(min_column) & subtitle_row & ":$" & ColLet(min_column) & subtitle2_row).Select
                End If
            Else
                VB_MASTER.Range("$" & ColLet(min_column) & first_row & ":$" & ColLet(max_column) & max_row).Select
            End If
        End If
    End With
    If Target.Address <> Selection.Address Then
        Set Target = Selection
    End If
   
    If event_change Then Application.EnableEvents = True
    
    '' SHOW COMMENTS BASED ON SELECTION
    If Target.Cells.count = 1 Then
        If Not Target.Comment Is Nothing Then
            Target.Comment.Visible = True
        End If
    End If
    
    'Define Selection Limit: the number of cells in the Scroll Area
    Dim MaxSelection As Long
    MaxSelection = VB_MASTER.Range(VB_MASTER.ScrollArea).Cells.CountLarge 'ScrollArea = LastColumn() * (get_row(last_mark) - first_row + 1)
    
    ' Excel forbids you from assigning a whole column
    ' or row to an array for performance reasons so
    ' keep it simple and set a hard limit: 3000
    '
    ' http://support.microsoft.com/kb/166342
    MaxSelection = VBMin(2999, MaxSelection)
    
    If Target.Cells.CountLarge > MaxSelection Then
        Exit Sub

    ElseIf Target.Cells.CountLarge <= MaxSelection Then 'Redefines vOldVal() and vOalValR() with an unambiguous length
        ReDim vOldValR(0 To Target.Cells.CountLarge - 1) As Range
        ReDim vOldVal(0 To Target.Cells.CountLarge - 1)
    End If
    
    'handle selection
    Select Case True
        Case Target.Cells.CountLarge > 1:
            ' multiple cells case
            multi = True ' More than one cell is selected
            Assign_Range_Values Target
        
        Case IsNull(Target.text):
            ' single Null cell case
            multi = False
            Set vOldValR(0) = Target
            vOldVal(0) = vbNullString
        
        Case Else:
            ' single cell case
            multi = False
            Set vOldValR(0) = Target
            vOldVal(0) = Target.Value2
    End Select

    Set Target = Nothing
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
'Subroutine: Worksheet_Change - Excel event handler. This Sub is called everytime a cell in VB_MASTER is changed.
'                               Worksheet_Change calls Write_Change for each selected cell where vOldVal
'                               differs from its current value. Since adding massive amount of protection, this
'                               only handles manual entry in VB_MASTER, ie "Checked Quantities".
'Arguments: target - Range object pointing to the cell or cells that have changed.


    Dim Action As Integer
    
    'Save event status
    Dim eventsOn As Boolean
    eventsOn = Application.EnableEvents
    
    On Error GoTo ResetApplication 'if Edit-Undo or some other method of changing cells is called, things could get hairy..
    
    ''''''' FROM THIS POINT ON do NOT use 'Exit Sub' instead use 'GoTo ResetApplication'
    
    If Target.Address = Target.EntireColumn.Address Then
        'do not handle entire columns here
        Application.EnabledEvents = False
        Application.Undo
        Application.EnabledEvents = eventsOn
        GoTo ResetApplication
    End If
    
    If Target.Address = Target.EntireRow.Address Then
        'do not handle entire rows here
        Application.EnabledEvents = False
        Application.Undo
        Application.EnabledEvents = eventsOn
        GoTo ResetApplication
    End If
    
    Dim wc_returns As Integer ' return value for Write_Change

    If multi Then 'multiple cell case -> For Loop; Call Write_Change for each item in array
        Dim lArr As Variant
        Dim i As Integer
        i = 0
        
        For Each lArr In vOldValR()
            If lArr.Value2 <> vOldVal(i) Or IsEmpty(lArr.Value2) = Not IsEmpty(vOldVal(i)) Then 'if the cell value has changed
                wc_returns = WriteChange(lArr, vOldVal(i))
            
                If wc_returns <> 0 Then
                    ErrorHandling "VB_MASTER.Worksheet_Change", 10 - wc_returns, "Write_Change failed: Err #(" & wc_returns & ")" & _
                        "Continue? [OK] End? [Cancel]", Action
                    If Action = Err_Exit Then
                        GoTo ResetApplication
                    End If
                End If
                vOldVal(i) = vbNullString
            End If
            
            i = i + 1
        Next lArr
        
    Else 'Single cell case
    
        If Target.Value2 <> vOldVal(0) Or IsEmpty(Target.Value2) = Not IsEmpty(vOldVal(0)) Then 'if the cell value has changed
            wc_returns = WriteChange(Target, vOldVal(0))
            
            If wc_returns <> 0 Then
                ErrorHandling "VB_MASTER.Worksheet_Change", 10 - wc_returns, "Write_Change failed: Err #(" & wc_returns & ")" & _
                    "Continue? [OK] End? [Cancel]", Action
                If Action = Err_Exit Then
                    GoTo ResetApplication
                End If
            End If
            
            vOldVal(0) = vbNullString
        End If
    End If
    
ResetApplication:
    
    Set Target = Nothing
End Sub

Private Sub Worksheet_Calculate()
'Subroutine: Worksheet_Calculate - Excel event handler. This Sub is called everytime a calculation ensues on the sheet.
'                                  Calls the CalculateQualitityFormat sub


    CalculateQuantityFormat
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
'Subroutine: Worksheet_BeforeRightClick - Excel event handler. This Subroutine is called everytime the user right-clicks
'                                         in a cell on this sheet. It suppresses the rick click menu.
'Arguments: Target - Range object containing cell in which the event happened
'           Cancel - Boolean returning the cancel operation

    Cancel = True
    CreateDisplayPopUpMenu Target
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
'Subroutine: Worksheet_BeforeDoubleClick - Excel event handler. The Sub is called everytime a double click event occurs.
'                                          Double clicking in the "Description Check" or "Delete?" columns requires
'                                          specific action; anywhere else, initialize the EditItem dialogue
'Arguments: Target - Range object for the cell in which the double click event occured.
'           Cancel - Boolean that returns the cancel operation.

    
    If IsEmpty(VB_MASTER.Cells(Target.row, get_col_num("Long Description"))) Then GoTo CancelDblClick
    
    ' handle doubling clicking in "Description Check"
    If get_ColTitle(Target.column) = "Description Check" And get_mark_num(Target.row) > 0 Then
        
        If IsEmpty(Target) Then 'originally unchecked
            ApproveDescription Target:=Target
            
        Else ' originally checked
            UnapproveDescription Target:=Target
        End If
    
    'handle double clicking in the "Delete?" column
    ElseIf get_ColTitle(Target.column) = "Delete?" And get_mark_num(Target.row) > 0 Then
        
        If IsEmpty(Target) Then ' delete
            ClientDeleteItem Target
        
        Else ' undelete
            ClientUndeleteItem Target
        End If
    
    ' handle other case: open EditItemWindow
    ElseIf Target.column <= LastColumn() Then
        
        Dim mark_num As Integer
        mark_num = get_mark_num(Target.row)
        
        If Target.column < 2 Or mark_num = -1 Then
            GoTo CancelDblClick
        End If
        
        If mark_num > 0 Then
            EditItemWindow.LoadForm Target
        End If
        
    End If
    
CancelDblClick:
    ' cancel double click action.
    Cancel = True
    
    Set Target = Nothing
End Sub

Public Sub CalculateQuantityFormat(Optional ByVal row As Integer = -1, Optional ByVal column As Integer = -1)
'Subroutine: CalculateQuantityFormat - Conditional Formatting fails to handle add/remove site actions, so
'                                      we handle site depend column formatting here. If the Checked Quantity is
'                                      equal to the Model Quantity, the cell is themed GREEN. If they are not equal,
'                                      the cell is themed RED.
'Arguments: row - If row is specified, this function only adjusts formatting for this row.
'           column - If column is specified, this function only adjusts formatting for this column.


    Dim chk_quanA As Integer
    Dim chk_quanB As Integer
    Dim first As Integer
    Dim last As Integer
    Dim diff As Integer
    
    'initialize position variables
    chk_quanA = get_col_num("Checked Quantities")
    chk_quanB = get_col_num("Checked Quantities") + VB_MASTER.Cells(title_row, chk_quanA).MergeArea.count - 1
    first = first_row
    last = VB_MASTER.LastRow()
    diff = chk_quanA - get_col_num("Current Model Quantities")
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    If row >= first And row <= last _
        And column >= chk_quanA And column <= chk_quanB Then
        'change this specifc cell
    
        If VB_MASTER.Cells(row, column).Value2 <> VB_MASTER.Cells(row, column - diff).Value2 Then
        'NONMATCH
            If VB_MASTER.Cells(row, column).Interior.Color <> RGB(218, 150, 148) Then
                VB_MASTER.Cells(row, column).Interior.Color = RGB(218, 150, 148)
                VB_MASTER.Cells(row, column).Font.Color = RGB(255, 0, 0) ' RED
            End If
        Else
        'MATCH
            If VB_MASTER.Cells(row, column).Interior.Color <> RGB(196, 215, 155) Then
                VB_MASTER.Cells(row, column).Interior.Color = RGB(196, 215, 155)
                VB_MASTER.Cells(row, column).Font.Color = RGB(79, 98, 40) ' GREEN
            End If
        End If
    
    Else
    
        'loop through all rows and columns matching "Current Model Quantities with "Checked Quantities"
        Dim i As Integer
        Dim j As Integer
        For i = chk_quanA To chk_quanB ' COLUMNS
            If i = column Or column = -1 Then
            
                For j = first To last ' ROWS
                
                    If j = row Or row = -1 Then
                        If VB_MASTER.Cells(j, i).Value2 <> VB_MASTER.Cells(j, i - diff).Value2 Then
                        'NONMATCH
                            If VB_MASTER.Cells(j, i).Interior.Color <> RGB(218, 150, 148) Then
                                VB_MASTER.Cells(j, i).Interior.Color = RGB(218, 150, 148)
                                VB_MASTER.Cells(j, i).Font.Color = RGB(255, 0, 0) ' RED
                            End If
                        Else
                        'MATCH
                            If VB_MASTER.Cells(j, i).Interior.Color <> RGB(196, 215, 155) Then
                                VB_MASTER.Cells(j, i).Interior.Color = RGB(196, 215, 155)
                                VB_MASTER.Cells(j, i).Font.Color = RGB(79, 98, 40) ' GREEN
                            End If
                        End If
                    End If
                    
                Next j
                
            End If
        Next i
        
    End If
    
    If ui_change Then RenderUI True
End Sub

Public Sub ResetScrollArea()
'Subroutine: ResetScrollArea - Resets ScrollArea Property for VB_MASTER. Controls where the user can scroll and select cells

    VB_MASTER.ScrollArea = "$" & ColLet(category_col) & "$" & subtitle_row & ":$" & ColLet(LastColumn()) & "$" & VB_MASTER.LastRow()
End Sub

Public Function WriteChange(ByVal Target As Range, ByVal old_value, Optional retroactive As Boolean = False) As Integer
'Function: WriteChange - Writes appropriate changes to LogChange(). If Target is out of bounds, it does nothing.
'Arguments: Target - Range object containing the current changed/selected cell being processed
'           old_value - The previous value held by Target
'           retroactive - Some methods called in WriteChange also call WriteChange, and vise versa. This optional
'                         argument is used by those methods to distingush between what's been called and what hasn't, and who is calling who.
'                         True - Call is retroactive, don't perpetuate; False - Call is not retroactive, continue normally
'Returns: Integer containing an error code. 0 - no error.

    
    ' correct sheet name so it goes into Change log correctly
    If InStr(VB_MASTER.Name, " ") <> 0 Then
        Dim corrected_name As String
        Dim i As Integer
        corrected_name = ""
        
        For i = 1 To Len(VB_MASTER.Name)
            If Mid(VB_MASTER.Name, i, 1) <> " " Then
                corrected_name = corrected_name & Mid(VB_MASTER.Name, i, 1)
            End If
        Next i
        
        VB_MASTER.Name = corrected_name
    End If
    
    'sanity check that Target is actually a cell reference
    If Not IsCellLocation(Target.Address) Then
        ErrorHandling "VB_MASTER.Worksheet_Change", 10, "Target Range is not a valid cell reference." & vbCrLf & _
            "target = " & Target.Address & "Terminating Process.", 1
        Exit Function
    End If
    
    Dim ValidChange As Boolean
    ValidChange = True
    
    Dim Action As Integer
    WriteChange = 0

    Dim mark_col As Integer
    mark_col = get_col_num("Mark No.")

    Dim first As Integer
    Dim last As Integer
    first = first_row
    last = VB_MASTER.LastRow()
    
    Dim mark_num As Integer
    mark_num = get_mark_num(Target.row)
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    'Check Target Bounds
    If Target.row < first Then
        Target.Value = old_value
        'MsgBox "Don't change this cell.", vbCritical
    ElseIf mark_num = -1 Then
        Target.Value = old_value
    ElseIf Target.column = mark_col And ThisWorkbook.MarkNumLock() Then
        Target.Value = old_value
        MsgBox "Mark numbers are locked.", vbCritical
    ElseIf Target.column = mark_col And Len(Target.Value2) = 0 Then
        Target.Value = old_value
    Else
        ' Target is in bounds: Check validity of mark num and deal with column specific instructions.
        
        Dim new_value
        new_value = Target.Value2
        
        Dim col_title As String
        col_title = get_ColTitle(Target.column)
        
        'check col_title for specific instructions
        If col_title = "Mark No." Then
            VB_CHANGE_LOG.reset_mark_num old_value
            
        'ElseIf col_title = "Long Description" Then
            'UnapproveDescription MarkNum:=mark_num <- This is taken care of in EditItemWindow now
        
        ElseIf col_title = "Description Check" Then
            'Approval command outside of EditItemWindow.
            'Revert cell to original value, continue normally, and prevent retroactive calling
            
            If IsEmpty(Target) Then
                new_value = ""
                old_value = Target.Comment.text
                If Not retroactive Then
                    Target.Value = old_value
                    
                    UnapproveDescription Target:=Target, write_change:=False
                End If
            Else
                If Not retroactive Then
                    Target.ClearContents
                    
                    ApproveDescription Target:=Target, write_change:=False
                End If
                new_value = Target.Comment.text
                old_value = ""
            End If
            
        ElseIf col_title = "Delete?" Then
            'ClientDelete command outside of EditItemWindow.
            'Revert cell to original value, continue normally, and prevent retroactive calling
            If IsEmpty(Target) Then
                new_value = ""
                old_value = GetDeleteReason(mark_num)
                If Not retroactive Then
                    Target.Value = old_value
                    
                    ClientUndeleteItem Target, write_change:=False
                End If
                col_title = "Undelete Item"
            Else
                If Not retroactive Then
                    Target.ClearContents
                    
                    ClientDeleteItem Target, write_change:=False
                End If
                new_value = GetDeleteReason(mark_num)
                old_value = ""
                col_title = "Delete Item"
            End If
        
        ElseIf col_title = "Orders" Then
            Target.Value = old_value
            If VB_ORDER_LOG.NumOrders() > 0 Then
                MsgBox "Published orders are set in stone. Changes can only be made by unpublishing then republishing the RFP with the corrected information.", vbExclamation
            End If
            GoTo ResetApplication
            
        ElseIf col_title = "Current Model Quantities" Or col_title = "Checked Quantities" Then
            If NumSites() = 0 Then
                Target.Value = old_value
                MsgBox "These columns become functional after site(s)/area(s) have been added. Use ""Add Site/Area"" under the Site Management group on the Ribbon to accomplish this.", vbInformation
                ValidChange = False
            Else
                If col_title = "Current Model Quantities" Then
                    CalculateQuantityFormat Target.row, Target.column + GetSitesRange().count
                
                ElseIf col_title = "Checked Quantities" Then
                    On Error Resume Next
                    Target.Comment.Delete
                    Err.Clear
                    On Error GoTo ResetApplication
                    
                    If IsEmpty(Target) Then
                        ' do nothing
                    Else
                        If Not IsNumeric(Target.Value2) Then
                            MsgBox "Quantities must be numeric.", vbExclamation
                            Target.Value = old_value
                            ValidChange = False
                        ElseIf Target.Value2 < 0 Then
                            MsgBox "Quantities must be positive integers.", vbExclamation
                            Target.Value = old_value
                            ValidChange = False
                        Else
                            Target.AddComment Environ("USERNAME") & vbCrLf & Date
                            Target.Comment.Shape.TextFrame.AutoSize = True
                        End If
                    End If
                    
                    CalculateQuantityFormat Target.row, Target.column
                End If
            End If
            
        End If
        
        VB_CATEGORY.CategoryAutoFit get_category(Target.row)
        Dim longlead As Boolean
        ShowLongLead_GetPressed Nothing, longlead
        If longlead Then Ribbon_ShowLongLead Nothing, True
        
        'Log Change in VB_CHANGE_LOG
        If VB_CHANGE_LOG.TrackChanges() And ValidChange Then
            If VB_CHANGE_LOG.LogChange(VB_MASTER.Name & "!" & Target.Address, mark_num, col_title, old_value, new_value) <> 0 Then
                WriteChange = -2
                GoTo ResetApplication
            End If
        End If
    End If
    
ResetApplication:
    If ui_change Then RenderUI True

    Set Target = Nothing
End Function

Public Sub UnlockColumns()
'Subroutine: UnlockColumns - The only direct user input allowed is in the "Checked Quantities" columns. The rest of
'                            VB_MASTER is locked and protected. On add/delete row/column events, this subroutine makes sure
'                            the right columns and rows are unlocked for editing.


    'set position variables
    Dim first_col As Integer
    first_col = get_col_num("Checked Quantities")
    
    Dim last_col As Integer
    last_col = first_col + GetSitesRange().count - 1
    
    Dim last_row As Integer
    last_row = LastRow()
    
    '''''''''''''''''''''''''''''''''''''''''''
    'Lock entire sheet, then unlock the appropriate rows/columns
    
    VB_MASTER.UsedRange.Locked = True
    
    VB_MASTER.Range("$" & ColLet(first_col) & "$" & first_row & ":$" & _
        ColLet(last_col) & "$" & last_row).Locked = False
        
    Dim delete_col As Integer
    delete_col = get_col_num("Delete?")
    VB_MASTER.Range("$" & ColLet(delete_col) & "$" & first_row & ":$" & _
        ColLet(delete_col) & "$" & last_row).Locked = False
    
    delete_col = get_col_num("Description Check")
    VB_MASTER.Range("$" & ColLet(delete_col) & "$" & first_row & ":$" & _
        ColLet(delete_col) & "$" & last_row).Locked = False
End Sub

Public Sub SetProtection()
'Subroutine: SetProtection - Sets UI protection.

    VB_MASTER.EnableOutlining = True
    VB_MASTER.Unprotect
    VB_MASTER.Protect DrawingObjects:=False, _
                        Contents:=True, _
                        Scenarios:=False, _
                        UserInterfaceOnly:=True, _
                        AllowFormattingCells:=True, _
                        AllowFormattingColumns:=False, _
                        AllowFormattingRows:=False, _
                        AllowInsertingColumns:=False, _
                        AllowInsertingRows:=False, _
                        AllowInsertingHyperlinks:=False, _
                        AllowDeletingColumns:=False, _
                        AllowDeletingRows:=False, _
                        AllowSorting:=False, _
                        AllowFiltering:=False, _
                        AllowUsingPivotTables:=False
    
    VB_MASTER.UnlockColumns
End Sub

Public Function LastRow() As Integer
'Function: LastRow - gets the row number from the last row item in VB_MASTER.
'Returns: Integer containing the last row number.


    LastRow = first_row

    Dim mark_col As Integer
    mark_col = get_col_num("Mark No.")
    
    Do While Not IsEmpty(VB_MASTER.Cells(LastRow + 2, mark_col))
        LastRow = LastRow + 2
    Loop
    
    If Not IsEmpty(VB_MASTER.Cells(LastRow + 1, mark_col)) Then
        LastRow = LastRow + 1
    End If
End Function

Public Function CategoryColumn() As Integer
'Function: TitleRow - 'Get' method for category_col property in VB_MASTER
'Returns: Integer containing the column number of the category titles in VB_MASTER

    CategoryColumn = category_col
End Function

Public Function SectionRow() As Integer
'Function: SectionRow - 'Get' method for section_row property in VB_MASTER
'Returns: Integer containing the row of the section titles in VB_MASTER

    SectionRow = section_row
End Function

Public Function TitleRow() As Integer
'Function: TitleRow - 'Get' method for title_row property in VB_MASTER
'Returns: Integer containing the row of the column titles in VB_MASTER

    TitleRow = title_row
End Function

Public Function SubtitleRow() As Integer
'Function: SubtitleRow - 'Get' method for subtitle_row property in VB_MASTER
'Returns: Integer containing the row of the first subtitle row in VB_MASTER

    SubtitleRow = subtitle_row
End Function

Public Function SubtitleRow2() As Integer
'Function: FirstRow - 'Get' method for subtitle2_row property in VB_MASTER
'Returns: Integer containing the row of the second subtitle row in VB_MASTER

    SubtitleRow2 = subtitle2_row
End Function

Public Function FirstRow() As Integer
'Function: FirstRow - 'Get' method for first_row property in VB_MASTER
'Returns: Integer containing the row of the first material item in VB_MASTER

    FirstRow = first_row
End Function
