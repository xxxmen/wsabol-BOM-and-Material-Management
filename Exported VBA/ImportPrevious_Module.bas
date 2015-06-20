Attribute VB_Name = "ImportPrevious_Module"
Option Explicit

Public Function GetPreviousBOMFileName()
'Function: GetPreviousBOMFileName - Handles 'Open' dialogue and returns the file path for the file choosen.
'Returns: String containing the file path. If dialogue was cancelled, returns vbNullString


    On Error Resume Next
    ChDir ThisWorkbook.Path
    Err.Clear
    On Error GoTo 0

    Dim filename As Variant
    filename = Application.GetOpenFilename(FileFilter:="Microsoft Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", _
        FilterIndex:=1, _
        title:="Select Previous BOM...", _
        MultiSelect:=False)
    
    If filename <> False Then
        GetPreviousBOMFileName = CStr(filename)
    Else
        GetPreviousBOMFileName = vbNullString
    End If
End Function

Public Function ImportPreviousBOM(ByVal BOM_filename As String) As Integer
'Function: ImportPreviousBOM - Imports previous version or another BOM into ThisWorkbook. It accomplishes this by using the procedures already set
'                              in place for normal operation. No need to reinvent the wheel. First start with opening to Previous Workbook. Then transfer
'                              the categories, transfer sites, add columns for the orders, transfer material items copying all quantites and descriptions.
'
'                                 **NOTE** This function does not copy files and folder associated with application. The process requires that ThisWorkbook
'                                 is in the same directory as the original to continue operation as usual.
'
'Arguments: BOM_filename - filename for previous workbook.
'Returns: Integer error code. 0: no error. -1: revert to blank BOM.

    
    VB_CHANGE_LOG.SetTrackChanges False

    Dim SrcWrbk As Workbook
    
    Dim CategorySource As Worksheet
    Dim ChangeLogSource As Worksheet
    Dim MasterBOMSource As Worksheet
    Dim OrderLogSource As Worksheet
    Dim CoverSheetSource As Worksheet
    Dim SiteDBSource As Worksheet
    Dim VarStoreSource As Worksheet
    Dim cadSheetSource As Worksheet
    
    Dim tempSheet As Variant
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    SetStatusBar "Matching Sheets to the Previous Version..."
    
    Application.DisplayAlerts = False
    Set SrcWrbk = Application.Workbooks.Open(filename:=BOM_filename, ReadOnly:=True)
    Application.DisplayAlerts = True
    
    'set source sheet object. loop through SrcWrkbk and match codenames
    For Each tempSheet In SrcWrbk.Worksheets
        If tempSheet.CodeName = "VB_MASTER" Then
            Set MasterBOMSource = tempSheet
            
        ElseIf tempSheet.CodeName = "VB_CATEGORY" Then
            Set CategorySource = tempSheet
            
        ElseIf tempSheet.CodeName = "VB_CHANGE_LOG" Then
            Set ChangeLogSource = tempSheet
            
        ElseIf tempSheet.CodeName = "VB_ORDER_LOG" Then
            Set OrderLogSource = tempSheet
            
        ElseIf tempSheet.CodeName = "VB_COVERSHEET" Then
            Set CoverSheetSource = tempSheet
            
        ElseIf tempSheet.CodeName = "VB_SITEDB" Then
            Set SiteDBSource = tempSheet
            
        ElseIf tempSheet.CodeName = "VB_VAR_STORE" Then
            Set VarStoreSource = tempSheet
        
        ElseIf tempSheet.Name = "CADWORX_IMPORT_OLD" Then
            Set cadSheetSource = tempSheet
        
        End If
        
        Set tempSheet = Nothing
    Next tempSheet
    
    'Get focus
    ThisWorkbook.Activate
    
    ''''''''''''''''''''''''''''
    ' transfer var_store stuff
    With VB_VAR_STORE
        If Len(.HMMFileTagFormat(sheet:=VarStoreSource)) > 0 Then .HMMFileTagFormat .HMMFileTagFormat(sheet:=VarStoreSource)
        If Len(CStr(.GetClientSummaryRev(VarStoreSource))) > 0 Then .SetClientSummaryRev .GetClientSummaryRev(VarStoreSource)
        If Len(.GetRFPDirectory(VarStoreSource)) > 0 Then .SetRFPDirectory .GetRFPDirectory(VarStoreSource)
        If Len(.GetClientSummaryDirectory(VarStoreSource)) > 0 Then .SetClientSummaryDirectory .GetClientSummaryDirectory(VarStoreSource)
        If Len(.GetSiteBOMDirectory(VarStoreSource)) > 0 Then .SetSiteBOMDirectory .GetSiteBOMDirectory(VarStoreSource)
        If Len(.GetMDRDirectory(VarStoreSource)) > 0 Then .SetMDRDirectory .GetMDRDirectory(VarStoreSource)
        If Len(.GetDeletedRFPDirectory(VarStoreSource)) > 0 Then .SetDeletedRFPDirectory .GetDeletedRFPDirectory(VarStoreSource)
    End With
    
    Dim row As Integer
    
    ''''''''''''''''''''''''''''
    'transfer categories
    SetStatusBar "Copying Categories..."
   
    Dim key_count As Integer
    Dim temp_cat As String
    Dim temp_unit As String
    
    ' erase all but one category
    row = VB_CATEGORY.FirstRow() + 1
    Do While Not IsEmpty(VB_CATEGORY.Cells(row, VB_CATEGORY.CategoryColumn()))
        DeleteCategory VB_CATEGORY.Cells(row, VB_CATEGORY.CategoryColumn()).Value2
    Loop
    
    'rename first category to match source
    temp_cat = CategorySource.Cells(VB_CATEGORY.FirstRow(), 1).Value2
    VB_CATEGORY.RenameCategory VB_CATEGORY.Cells(VB_CATEGORY.FirstRow(), VB_CATEGORY.CategoryColumn()).Value2, temp_cat
    
    temp_unit = CategorySource.Cells(VB_CATEGORY.FirstRow(), 2).Value2
    VB_CATEGORY.Cells(VB_CATEGORY.FirstRow(), VB_CATEGORY.UnitColumn()).Value = temp_unit
    VB_MASTER.Cells(get_cat_row(temp_cat), get_col_num("Unit")).Value = temp_unit
    
    VB_CATEGORY.ClearPhrases temp_cat
    Dim phrases_col As Integer
    phrases_col = VB_CATEGORY.PhrasesColumn()
    
    Do While Not IsEmpty(CategorySource.Cells(VB_CATEGORY.FirstRow(), phrases_col))
        VB_CATEGORY.StoreFirstPhraseKey temp_cat, CategorySource.Cells(VB_CATEGORY.FirstRow(), phrases_col).Value2
        
        phrases_col = phrases_col + 1
    Loop
    
    row = VB_CATEGORY.FirstRow() + 1
    'loop through source, copy category, copy keys
    Do While Not IsEmpty(CategorySource.Cells(row, 1))
        temp_cat = CategorySource.Cells(row, 1).Value2
        temp_unit = CategorySource.Cells(row, 2).Value2
        
        'add category to MASTER
        If Not VB_CATEGORY.CategoryExists(temp_cat) Then
            Call NewCategory(temp_cat, temp_unit, VB_MASTER.LastRow() + 1)
        End If
        
        VB_CATEGORY.ClearPhrases temp_cat
        phrases_col = VB_CATEGORY.PhrasesColumn()
        
        Do While Not IsEmpty(CategorySource.Cells(row, phrases_col))
            VB_CATEGORY.StoreFirstPhraseKey temp_cat, CategorySource.Cells(row, phrases_col).Value2
            
            phrases_col = phrases_col + 1
        Loop
        
        row = row + 1
    Loop
    
    '''''''''''''''''''''''''''''
    'transfer sites
    SetStatusBar "Copying Sites..."
    
    Dim num_site_cols As Integer
    Dim has_sites As Boolean
    Dim site_col As Integer
    Dim site_end As Integer
    has_sites = False
    site_col = get_col_num("Current Model Quantities", MasterBOMSource)
    site_end = site_col + MasterBOMSource.Cells(VB_MASTER.TitleRow(), site_col).MergeArea.count - 1
    num_site_cols = site_end - site_col + 1
    
    'loop through source's site columns, add each site
    Do While site_col <= site_end
        If Not IsEmpty(MasterBOMSource.Cells(VB_MASTER.SubtitleRow(), site_col)) Then
            has_sites = True
            AddSite MasterBOMSource.Cells(VB_MASTER.SubtitleRow(), site_col).Value2, MasterBOMSource.Cells(VB_MASTER.SubtitleRow(), site_col).Font.Italic
        End If
        site_col = site_col + 1
    Loop
    
    ' transfer SiteBOM rev# and Client Summary info
    If has_sites And Not SiteDBSource Is Nothing Then
        VB_SITEDB.UsedRange.Value = SiteDBSource.UsedRange.Value2
    End If

    '''''''''''''''''''''''''''''''
    'transfer RFP/Orders
    SetStatusBar "Copying RFP/Order columns..."
    
    Dim new_order_col As Integer
    Dim o_index As Integer
    Dim order_num As String
    Dim order_site As String
    Dim order_date As String
    Dim SrcNumOrders As Integer
    SrcNumOrders = VB_ORDER_LOG.NumOrders(OrderLogSource)
    
    For o_index = 0 To SrcNumOrders - 1
        order_num = OrderLogSource.Cells(o_index + VB_ORDER_LOG.FirstRow(), VB_ORDER_LOG.ColID("Order")).Value2
        order_date = OrderLogSource.Cells(o_index + VB_ORDER_LOG.FirstRow(), VB_ORDER_LOG.ColID("Date")).Value
        order_site = OrderLogSource.Cells(o_index + VB_ORDER_LOG.FirstRow(), VB_ORDER_LOG.ColID("Site")).Value2
        
        new_order_col = get_col_num("Total Ordered")
        NewOrderColumn new_order_col, order_num, order_site, order_date
        VB_ORDER_LOG.LogOrder order_num, order_date, order_site
        
        VB_ORDER_LOG.SetReceipt order_num, _
            (OrderLogSource.Cells(o_index + VB_ORDER_LOG.FirstRow(), VB_ORDER_LOG.ColID("Receipt")).Value2 = "1"), _
            OrderLogSource.Cells(o_index + VB_ORDER_LOG.FirstRow(), VB_ORDER_LOG.ColID("Receipt Date")).Value
    Next
    
    ''''''''''''''''''''''''''''''''
    'transfer items
    SetStatusBar "Copying Material Properties and Quantities Columns..."
    
    Dim item_row As Integer
    Dim src_desc_col As Integer
    Dim src_ID_col As Integer
    Dim first_col As Integer
    Dim last_col As Integer
    Dim new_mark As Integer
    Dim new_item_row As Integer
    Dim mark_col As Integer
    
    item_row = get_first_row(MasterBOMSource)
    src_desc_col = get_col_num("Long Description", MasterBOMSource)
    src_ID_col = get_col_num("Unique ID", MasterBOMSource)
    first_col = get_col_num("Orders")
    last_col = LastColumn()
    mark_col = get_col_num("Mark No.")
    
    'loop through items and add each one, copy site qualitities too.
    Do While Not IsEmpty(MasterBOMSource.Cells(item_row, src_ID_col))
        temp_cat = get_category(item_row, MasterBOMSource)
        
        'check if category exists
        If get_cat_row(temp_cat) = -1 Then
            If Not VB_CATEGORY.CategoryExists(temp_cat) Then
                NewCategory temp_cat, temp_unit, VB_MASTER.LastRow() + 1
            Else
                GoTo SKIP_ITEM
            End If
        End If
        
        If Len(MasterBOMSource.Cells(item_row, src_desc_col).Value2) = 0 Then
            GoTo SKIP_ITEM
        End If
        
        'Add item to this BOM
        new_mark = InsertItemBOM(temp_cat, MasterBOMSource.Cells(item_row, src_desc_col).Value2)
        new_item_row = get_row(new_mark, VB_MASTER)
        
        'copy column values; not mark nums
        Dim i As Integer
        Dim i_src As Integer
        Dim c As Integer
        Dim this_cmt As Comment
        Dim prev_cmt As Comment
        
        'this loop accounts for different column orders between versions
        'If cell is not empty and doesn't have a formula, copy value to this BOM
        For i = first_col To last_col
            i_src = get_col_num(get_ColTitle(i), MasterBOMSource)
            If i_src = -1 Then
                GoTo GOTO_NEXT
            End If
            i_src = i_src + i - get_col_num(get_ColTitle(i))
            
            If Not (MasterBOMSource.Cells(item_row, i_src).HasFormula Or i = mark_col) Then
                VB_MASTER.Cells(new_item_row, i).Value = MasterBOMSource.Cells(item_row, i_src).Value2
                
                'for ClientDELETED items, font color is important
                For c = Len(VB_MASTER.Cells(new_item_row, i).Value2) To 1 Step -1
                    With VB_MASTER.Cells(new_item_row, i).Characters(start:=c, length:=1).Font
                        .Color = MasterBOMSource.Cells(item_row, i_src).Characters(start:=c, length:=1).Font.Color
                        .Strikethrough = MasterBOMSource.Cells(item_row, i_src).Characters(start:=c, length:=1).Font.Strikethrough
                    End With
                Next c
                
                ' copy comments
                If Not MasterBOMSource.Cells(item_row, i_src).Comment Is Nothing Then
                    Set prev_cmt = MasterBOMSource.Cells(item_row, i_src).Comment
                    If Len(prev_cmt.text) > 0 Then
                        On Error Resume Next
                        VB_MASTER.Cells(new_item_row, i).Comment.Delete
                        Err.Clear
                        On Error GoTo 0
                        
                        'add comment to this BOM
                        Set this_cmt = VB_MASTER.Cells(new_item_row, i).AddComment(TrimWhiteSpace(prev_cmt.text))
                        
                        'size and shape
                        this_cmt.Shape.TextFrame.AutoSize = True
                    End If
                End If
                
            End If
            
GOTO_NEXT:
        Next i
SKIP_ITEM:
        item_row = item_row + 1
    Loop
    Err.Clear
    On Error GoTo 0
    
    '''''''''''''''''''''''''''''''''''''''''
    'copy mark_nums
    new_item_row = VB_MASTER.FirstRow()
    
    Do While Not IsEmpty(VB_MASTER.Cells(new_item_row, mark_col))
        item_row = get_first_row(MasterBOMSource)
        new_mark = 0
        
        Do While Not IsEmpty(MasterBOMSource.Cells(item_row, mark_col))
            If MasterBOMSource.Cells(item_row, src_desc_col).Value2 = VB_MASTER.Cells(new_item_row, src_desc_col).Value2 _
                And get_category(item_row, MasterBOMSource) = get_category(new_item_row) Then
                
                new_mark = MasterBOMSource.Cells(item_row, mark_col).Value2
                Exit Do
            End If
            item_row = item_row + 1
        Loop
        If new_mark <> 0 Then
            VB_MASTER.Cells(new_item_row, mark_col).Value = new_mark
        End If
        new_item_row = new_item_row + 1
    Loop
    
    ''''''''''''''''''''''''''''''''''''''''
    ' match hidden columns
    
    On Error GoTo COPY_SHEETS
    If VB_MASTER.Columns(get_col_num("SAP#")).Hidden <> MasterBOMSource.Columns(get_col_num("SAP#", MasterBOMSource)).Hidden Then
        VB_MASTER.Columns(get_col_num("SAP#")).Hidden = MasterBOMSource.Columns(get_col_num("SAP#", MasterBOMSource)).Hidden
        GetRibbon().Invalidate
    End If
    If VB_MASTER.Columns(get_col_num("Total Extras")).Hidden <> MasterBOMSource.Columns(get_col_num("Total Extras", MasterBOMSource)).Hidden Then
        VB_MASTER.Columns(get_col_num("Total Extras")).Hidden = MasterBOMSource.Columns(get_col_num("Total Extras", MasterBOMSource)).Hidden
        GetRibbon().Invalidate
    End If
    If VB_MASTER.Columns(get_col_num("Mark No.")).Hidden <> MasterBOMSource.Columns(get_col_num("Mark No.", MasterBOMSource)).Hidden Then
        VB_MASTER.Columns(get_col_num("Mark No.")).Hidden = MasterBOMSource.Columns(get_col_num("Mark No.", MasterBOMSource)).Hidden
        GetRibbon().Invalidate
    End If
    Err.Clear
    On Error GoTo 0
    
    ''''''''''''''''''''''''''''''''''''''''
    'copy RFP Report Current
    
COPY_SHEETS:
    SetStatusBar "Copying Sheets..."

    Dim current_rfp_report As Worksheet
    Dim new_rfp_report As Worksheet
    If Not SheetExists("RFP Report Current", SrcWrbk) Then
        GoTo DONEWITHRFPREPORT
    End If
    On Error GoTo DONEWITHRFPREPORT
    Set current_rfp_report = SrcWrbk.Sheets("RFP Report Current")
    
    Dim curr_visibility As Integer
    curr_visibility = current_rfp_report.Visible
    current_rfp_report.Visible = xlSheetVisible
    
    Application.DisplayAlerts = False
    current_rfp_report.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    Set new_rfp_report = ThisWorkbook.ActiveSheet
    Application.DisplayAlerts = True
    
    current_rfp_report.Visible = curr_visibility
    new_rfp_report.Visible = curr_visibility
    
    On Error GoTo ClientSummaryError
    VB_SITEDB.SetClientSummaryAliasToPrevious
    Err.Clear
    On Error GoTo 0
    GoTo DONEWITHRFPREPORT
    
ClientSummaryError:
    On Error GoTo DONEWITHRFPREPORT
    Set new_rfp_report = ThisWorkbook.Sheets("RFP Report Current")
    new_rfp_report.Visible = xlSheetVisible
    Application.DisplayAlerts = False
    new_rfp_report.Delete
    Application.DisplayAlerts = True
    Err.Clear
    On Error GoTo 0
    
DONEWITHRFPREPORT:

    ''''''''''''''''''''''''''''''''''''''''
    ' copy cad import sheet
    
    Dim cadSheetDest As Worksheet
    
    If Not cadSheetSource Is Nothing Then
        On Error GoTo DONEWITHCADIMPORT
        cadSheetSource.Visible = xlSheetVisible
        
        Application.DisplayAlerts = False
        cadSheetSource.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
        Set cadSheetDest = ThisWorkbook.ActiveSheet
        Application.DisplayAlerts = True
        
        cadSheetSource.Visible = xlSheetVeryHidden
        cadSheetDest.Visible = xlSheetVeryHidden
        
        cadSheetDest.Name = "CADWORX_IMPORT_OLD"
        cadSheetDest.UsedRange.Validation.Delete
    End If
    
DONEWITHCADIMPORT:

    ''''''''''''''''''''''''''''''''''''''''
    ' copy Site BOMs
    
    Application.DisplayAlerts = False
    On Error GoTo DONEWITHSITEBOMS

    Dim site_name As String
    Dim site_sheet_name As String
    site_col = get_col_num("Current Model Quantities", MasterBOMSource)
    site_end = site_col + MasterBOMSource.Cells(VB_MASTER.TitleRow(), site_col).MergeArea.count - 1
    
    'loop through source's site columns, add each site
    Do While site_col <= site_end
        If Not IsEmpty(MasterBOMSource.Cells(VB_MASTER.SubtitleRow(), site_col)) Then
            site_name = MasterBOMSource.Cells(VB_MASTER.SubtitleRow(), site_col).Value2
            
            If Len(site_name) > 25 Then
                site_sheet_name = Left(site_name, 25) & " - BOM"
            Else
                site_sheet_name = site_name & " - BOM"
            End If
            
            If SheetExists(site_sheet_name, SrcWrbk) Then
                CreateSiteSpecificBOM site_name, False
            End If
            
        End If
        site_col = site_col + 1
    Loop
    
    Err.Clear
    On Error GoTo 0
    Application.DisplayAlerts = True

DONEWITHSITEBOMS:
    
    'transfer change_log
    SetStatusBar "Copying Change Log..."
    With ChangeLogSource.UsedRange.Rows("2:" & ChangeLogSource.UsedRange.Rows.CountLarge)
        VB_CHANGE_LOG.Range(.Address).Value = .Value2
    End With
    
    Dim cl_row As Long
    For cl_row = 1 To ChangeLogSource.UsedRange.Rows.CountLarge
        If ChangeLogSource.UsedRange.Rows(cl_row).Hidden Then
            VB_CHANGE_LOG.UsedRange.Rows(cl_row).Hidden = True
        End If
    Next cl_row
    
    SetStatusBar "Copying Cover Sheet Information..."
    
    On Error GoTo NUMLOCK_HERE
    'COVERSHEET
    VB_COVERSHEET.Unprotect

    If Not CoverSheetSource Is Nothing Then
        Application.EnableEvents = True
        VB_COVERSHEET.Range("$F$14").MergeArea.Value = CoverSheetSource.Range("$F$14").MergeArea.Cells(1, 1).Value2
        VB_COVERSHEET.Range("$F$16").MergeArea.Value = CoverSheetSource.Range("$F$16").MergeArea.Cells(1, 1).Value2
        VB_COVERSHEET.Range("$F$20").MergeArea.Value = CoverSheetSource.Range("$F$20").MergeArea.Cells(1, 1).Value2
        VB_COVERSHEET.Range("$F$22").MergeArea.Value = CoverSheetSource.Range("$F$22").MergeArea.Cells(1, 1).Value2
        VB_COVERSHEET.Range("$K$22").MergeArea.Value = CoverSheetSource.Range("$K$22").MergeArea.Cells(1, 1).Value2
        VB_COVERSHEET.Range("$F$25").MergeArea.Value = CoverSheetSource.Range("$F$25").MergeArea.Cells(1, 1).Value2
        VB_COVERSHEET.Range("$I$25").MergeArea.Value = CoverSheetSource.Range("$I$25").MergeArea.Cells(1, 1).Value2
        Application.EnableEvents = False
    End If
    
    VB_COVERSHEET.SetProtection
    On Error GoTo 0
    Err.Clear
    
NUMLOCK_HERE:
    On Error GoTo NamesElse
    ThisWorkbook.SetMarkNumLock StrComp(SrcWrbk.Names("NumLockState"), "=TRUE", vbTextCompare) = 0
    GoTo NamesSuccess
NamesElse:
    Err.Clear
    ThisWorkbook.SetMarkNumLock False
    VB_CHANGE_LOG.LogChange VB_CHANGE_LOG.Name & "!$A$1", "UNLOCKED Mark Numbers", "", "", "FALSE"
NamesSuccess:

    SetStatusBar "Finishing..."
    
    Set new_rfp_report = Nothing
    Set cadSheetSource = Nothing
    Set cadSheetDest = Nothing
    Set current_rfp_report = Nothing
    Set new_rfp_report = Nothing
    
    'close previous BOM
    Application.DisplayAlerts = False
    SrcWrbk.Saved = True
    SrcWrbk.Close
    Set SrcWrbk = Nothing
    Application.DisplayAlerts = True

    VB_MASTER.CalculateQuantityFormat
    VB_CHANGE_LOG.SetTrackChanges True
    GetRibbon().Invalidate
    VB_COVERSHEET.Activate
    ResetStatusBar
    If ui_change Then RenderUI True

    Set CategorySource = Nothing
    Set ChangeLogSource = Nothing
    Set MasterBOMSource = Nothing
    Set OrderLogSource = Nothing
    Set CoverSheetSource = Nothing
    ImportPreviousBOM = 0
End Function
