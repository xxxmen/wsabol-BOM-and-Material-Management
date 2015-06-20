Attribute VB_Name = "RightClickMenu"
Option Explicit

Private Const Mname As String = "HMMMasterBOMPopUpMenu"
Private clickTarget As Range

Private Sub DeletePopUpMenu()
'Subroutine: DeletePopUpMenu - Removes the Pop Up menu from the Application's Command Bars

    ' Delete the popup menu if it already exists.
    On Error Resume Next
    Application.CommandBars(Mname).Delete
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub CreateDisplayPopUpMenu(ByVal Target As Range)
'Subroutine: CreateDisplayPopUpMenu - Creates and Display the proper pop-up menu. There are different popups depending
'                                     on what cell(s) are selected.
'Arguments: Target - Range object containing the Active Selection.


    ' Delete any existing popup menu.
    Call DeletePopUpMenu
    
    ' Sanity Check
    If Target Is Nothing Then
        Exit Sub
    End If
    If Target.Parent.Name <> VB_MASTER.Name Or Not VB_MASTER Is ThisWorkbook.ActiveSheet Then
        Exit Sub
    End If
    
    ' In VB_MASTER_SelectionChange(), the selection could change so,
    Set clickTarget = Selection
    
    ' Create the popup menu.
    If (clickTarget.Address Like "$*$" & VB_MASTER.SubtitleRow() & ":$*$" & VB_MASTER.SubtitleRow2() _
        And MinColumn(clickTarget) > VB_MASTER.CategoryColumn()) _
        Or (MinRow(Target) >= VB_MASTER.FirstRow() And _
            (clickTarget.Cells.count = 1 Or MinColumn(clickTarget) = VB_MASTER.CategoryColumn())) Then
        
        ' single item dialogue
        SingleSelectPopUpMenu
    ElseIf MinRow(Target) >= VB_MASTER.FirstRow() _
        And clickTarget.Cells.count > 1 _
        And MinColumn(clickTarget) > VB_MASTER.CategoryColumn() Then
        
        ' multi item dialogue
        MultiSelectPopUpMenu
    End If

    ' Display the popup menu.
    On Error Resume Next
    Application.CommandBars(Mname).ShowPopup
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub SingleSelectPopUpMenu()
'Subroutine: SingleSelectPopUpMenu - Create the menu for the single item case.
'                                    Single items include: categoies, items, sites, orders.

    Dim order_num As String
    Dim site_name As String
    Dim category As String
    Dim mark_num As Integer
    mark_num = get_mark_num(clickTarget.row)
    
    If mark_num > 0 And clickTarget.column > VB_MASTER.CategoryColumn() Then
        
        ' SINGLE MATERIAL ITEM
        With Application.CommandBars.Add(Name:=Mname, position:=msoBarPopup, _
             MenuBar:=False, Temporary:=True)
            
            ' First, add two buttons to the menu.
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Master BOM Item #" & mark_num
                .enabled = False
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Edit Material Info"
                .FaceId = 71
                .OnAction = "'" & ThisWorkbook.Name & "'!GetInfo_Click"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                If HasNote(mark_num) Then
                    .Caption = "Edit Notes"
                    .OnAction = "'" & ThisWorkbook.Name & "'!" & "AddEditNotes_Click"
                Else
                    .Caption = "Add Notes"
                    .OnAction = "'" & ThisWorkbook.Name & "'!" & "AddEditNotes_Click"
                End If
                .FaceId = 72
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                If IsEmpty(VB_MASTER.Cells(clickTarget.row, get_col_num("Description Check"))) Then
                    .Caption = "Approve Description"
                    .OnAction = "'" & ThisWorkbook.Name & "'!" & "ApproveDescription_Click"
                Else
                    .Caption = "Unapprove Description"
                    .OnAction = "'" & ThisWorkbook.Name & "'!" & "UnapproveDescription_Click"
                End If
                .FaceId = 73
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                If CBool(VB_MASTER.Cells(clickTarget.row, get_col_num("Long Lead")).Value2) Then
                    .Caption = "Mark as NOT Long Lead"
                    .OnAction = "'" & ThisWorkbook.Name & "'!" & "LongLead_Click"
                Else
                    .Caption = "Mark as Long Lead"
                    .OnAction = "'" & ThisWorkbook.Name & "'!" & "LongLead_Click"
                End If
                .FaceId = 74
            End With
        End With
        
    ElseIf mark_num > 0 And clickTarget.column = VB_MASTER.CategoryColumn() Then
        category = get_category(get_row(mark_num))
        
        ' SINGLE CATEGORY
        With Application.CommandBars.Add(Name:=Mname, position:=msoBarPopup, _
             MenuBar:=False, Temporary:=True)
            
           ' First, add two buttons to the menu.
            With .Controls.Add(Type:=msoControlButton)
                .Caption = category
                .enabled = False
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Rename Category"
                .FaceId = 71
                .OnAction = "'" & ThisWorkbook.Name & "'!RenameCategory_Click"
            End With
            
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Reorder Categories"
                .FaceId = 72
                .OnAction = "'" & ThisWorkbook.Name & "'!ReorderCategories_Click"
            End With
        End With
    Else
    
        Dim col_title As String
        col_title = get_ColTitle(clickTarget.column)
        
        If col_title = "Orders" Then
            order_num = VB_MASTER.Cells(VB_MASTER.SubtitleRow(), clickTarget.column).Value2
            
            ' SINGLE ORDER/RFP
            With Application.CommandBars.Add(Name:=Mname, position:=msoBarPopup, _
                 MenuBar:=False, Temporary:=True)
                
               ' First, add two buttons to the menu.
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = order_num
                    .enabled = False
                End With
                
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "Open PDF"
                    .FaceId = 71
                    .OnAction = "'" & ThisWorkbook.Name & "'!OpenRFPPDF_Click"
                End With
                
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "Go to RFP Directory"
                    .FaceId = 72
                    .OnAction = "'" & ThisWorkbook.Name & "'!GotoRFPDir_Click"
                End With
                
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "Order Manager"
                    .FaceId = 73
                    .OnAction = "'" & ThisWorkbook.Name & "'!OrderManager_Click"
                End With
            End With
            
        Else
            site_name = VB_MASTER.Cells(VB_MASTER.SubtitleRow(), clickTarget.column).Value2
            
            ' SINGLE SITE
            With Application.CommandBars.Add(Name:=Mname, position:=msoBarPopup, _
                 MenuBar:=False, Temporary:=True)
                
               ' First, add two buttons to the menu.
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = site_name
                    .enabled = False
                End With
                
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "Edit Site Info"
                    .FaceId = 71
                    .OnAction = "'" & ThisWorkbook.Name & "'!EditSite_Click"
                End With
                
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "Clear Model Quantities"
                    .FaceId = 72
                    .OnAction = "'" & ThisWorkbook.Name & "'!ClearModelQuantities_Click"
                End With
                
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "Go to Site BOM Directory"
                    .FaceId = 73
                    .OnAction = "'" & ThisWorkbook.Name & "'!GotoSiteBOMDir_Click"
                End With
                
            End With
        End If
    End If
End Sub

Private Sub MultiSelectPopUpMenu()
'Subroutine: MultiSelectPopUpMenu - Create the menu for the multi-item case. Valid for
'                                   a selection of multiple material items.


    With Application.CommandBars.Add(Name:=Mname, position:=msoBarPopup, _
         MenuBar:=False, Temporary:=True)
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Selected Items"
            .enabled = False
        End With
        
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Move to Category"
            .FaceId = 71
            .OnAction = "'" & ThisWorkbook.Name & "'!MovetoCategory_Click"
        End With
    End With
End Sub

' Callback for "Get Info" onAction
Sub GetInfo_Click()
    EditItemWindow.LoadForm clickTarget
End Sub

' Callback for "Add/Edit Notes" onAction
Sub AddEditNotes_Click()
    AddEditNotes.LoadForm get_mark_num(clickTarget.row)
End Sub

' Callback for "ApproveDescription" onAction
Sub ApproveDescription_Click()
    ApproveDescription MarkNum:=get_mark_num(clickTarget.row)
End Sub

' Callback for "UnapproveDescription" onAction
Sub UnapproveDescription_Click()
    UnapproveDescription MarkNum:=get_mark_num(clickTarget.row)
End Sub

' Callback for "Mark as (NOT) Long Lead" onAction
Sub LongLead_Click()
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    VB_MASTER.Cells(clickTarget.row, get_col_num("Long Lead")).Value = Not CBool(VB_MASTER.Cells(clickTarget.row, get_col_num("Long Lead")).Value2)
    VB_MASTER.WriteChange VB_MASTER.Cells(clickTarget.row, get_col_num("Long Lead")), Not CBool(VB_MASTER.Cells(clickTarget.row, get_col_num("Long Lead")).Value2)
    
    If ui_change Then RenderUI True
End Sub

' Callback for "Order Manager" onAction
Sub OrderManager_Click()
    OrderTracking.LoadForm
End Sub

' Callback for "Got to RFP Directory" onAction
Sub GotoRFPDir_Click()
    Shell "explorer.exe " & ThisWorkbook.Path & VB_VAR_STORE.GetRFPDirectory(), vbNormalFocus
End Sub

' Callback for "Open PDF" onAction
Sub OpenRFPPDF_Click()
    Dim RFPresult As String
    
    Dim order_num As String
    order_num = VB_MASTER.Cells(VB_MASTER.SubtitleRow(), clickTarget.column).Value2
    
    RFPresult = Dir(ThisWorkbook.Path & VB_VAR_STORE.GetRFPDirectory() & order_num & "*.pdf")
    
    If Len(RFPresult) > 0 Then
        ActiveWorkbook.FollowHyperlink ThisWorkbook.Path & VB_VAR_STORE.GetRFPDirectory() & RFPresult
    Else
        MsgBox "File not found in the specified RFP directory. Check your folder settings and try again.", vbCritical
    End If
End Sub

' Callback for "Rename Category" onAction
Sub RenameCategory_Click()
    Dim category As String
    category = get_category(clickTarget.row)

    Dim new_category As String
    RenameCategoryForm.LoadForm category, new_category

    If get_cat_row(new_category) <> -1 Then
        MsgBox "Successfully renamed " & category & " as " & new_category & "."
    End If
End Sub

' Callback for "Reorder Categories" onAction
Sub ReorderCategories_Click()
    ReorderCategoriesForm.LoadForm
End Sub

' Callback for "Edit Site Info" onAction
Sub EditSite_Click()
    Dim site_name As String
    site_name = VB_MASTER.Cells(VB_MASTER.SubtitleRow(), clickTarget.column).Value2

    Dim new_site_name As String
    NewSiteForm.LoadForm new_site_name, site_name

    If SiteExists(new_site_name) Then
        MsgBox "Successfully renamed " & site_name & " as " & new_site_name & "."
    End If
End Sub

' Callback for "Clear Model Quantities" onAction
Sub ClearModelQuantities_Click()
    Dim site_name As String
    site_name = VB_MASTER.Cells(VB_MASTER.SubtitleRow(), clickTarget.column).Value2

    Dim site_col As Integer
    If SiteExists(site_name, site_col) Then
        ' not crazy
    End If
    
    Dim quantity As Integer
    Dim row As Integer
    Dim first As Integer
    Dim last As Integer
    first = VB_MASTER.FirstRow()
    last = VB_MASTER.LastRow()
    
    Dim msgResult As Integer
    msgResult = MsgBox("You are about to clear all the model quantity information in this column. This can be undone/redone. Continue?", vbOKCancel)
    
    If msgResult = vbOK Then
        ' continue...
        
        Dim ui_change As Boolean
        ui_change = RenderUI(False)
        
        'LOG CHANGE
        VB_CHANGE_LOG.LogChange VB_MASTER.Name & "!" & VB_MASTER.Cells(VB_MASTER.SubtitleRow(), site_col).Address, "", "BEGIN Clear Model Quantities", "", ""
        
        For row = first To last
            SetStatusBar "Clearing " & site_name & " Quantities...", (row - first), (last - first)
        
            If Len(VB_MASTER.Cells(row, site_col).Value2) > 0 Then
                quantity = CInt(VB_MASTER.Cells(row, site_col).Value2)
                
                VB_MASTER.Cells(row, site_col).ClearContents
                VB_MASTER.WriteChange VB_MASTER.Cells(row, site_col), quantity
                
            End If
        Next row
        
        ResetStatusBar
        
        'LOG CHANGE
        VB_CHANGE_LOG.LogChange VB_MASTER.Name & "!" & VB_MASTER.Cells(VB_MASTER.SubtitleRow(), site_col).Address, "", "END Clear Model Quantities", "", ""
        
        If ui_change Then RenderUI True
    
    End If
    
End Sub

' Callback for "Go to Site BOM Directory" onAction
Sub GotoSiteBOMDir_Click()
    Shell "explorer.exe " & ThisWorkbook.Path & VB_VAR_STORE.GetSiteBOMDirectory(), vbNormalFocus
End Sub

' Callback for "Move to Category" onAction
Sub MovetoCategory_Click()
    Dim category As String
    ChooseCategory.LoadForm category
    
    If Len(category) = 0 Then
        Exit Sub
    End If
    
    If get_cat_row(category) <> -1 Then
        Dim vbCell As Range
        Dim MarkNums As Collection
        Set MarkNums = New Collection
        
        Dim temp_mark As Integer
        
        For Each vbCell In clickTarget
            temp_mark = get_mark_num(vbCell.row)
            
            If temp_mark > 0 Then
                If Not CollectionContains(MarkNums, temp_mark) Then
                    MarkNums.Add temp_mark
                End If
            End If
        Next vbCell
        
        Dim org_category As String
        Dim ui_change As Boolean
        ui_change = RenderUI(False)
        
        Dim mark_num
        For Each mark_num In MarkNums
            org_category = get_category(get_row(mark_num))
            If org_category <> category Then
                AutoSortItem mark_num, category
                 Call VB_CHANGE_LOG.LogChange(VB_MASTER.Name & "!$" & VB_MASTER.CategoryColumn() & "$" & get_row(mark_num), _
                    mark_num, "Category Change", org_category, category)
            End If
        Next mark_num
        
        If ui_change Then RenderUI True
    End If
End Sub
