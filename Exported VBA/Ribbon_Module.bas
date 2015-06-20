Attribute VB_Name = "Ribbon_Module"
Option Explicit

Private ThisRibbon As IRibbonUI

Private comparison_site As String

Private Old_CurrentCS_Text As String
Private Old_PrevCS_Text As String
Private Old_HideZeros As Boolean
Private CurrentCS_Text As String
Private PrevCS_Text As String
Private hideZeros As Boolean

Public Declare Sub CopyMemory Lib "kernel32" Alias _
    "RtlMoveMemory" (destination As Any, source As Any, _
    ByVal length As Long)

Public Sub RibbonSaver(ribbon As IRibbonUI)
    ' Store pointer to IRibbonUI
    Dim lngRibPtr As Long
    
    ' Store the custom ribbonUI ID in a static variable.
    ' This is done once during load of UI. i.e. during workbook open.
    lngRibPtr = ObjPtr(ribbon)
    'MsgBox lngRibPtr
    
    ' Write pointer to worksheet for safe keeping
    VB_VAR_STORE.SetRibbonID lngRibPtr
End Sub

Public Function GetRibbon() As Object
'Function: GetRibbon - Returns IRibbonUI object pointing to the Excel Ribbon
'Returns: Object. Equivalent to ThisRibbon.


    ' copy ThisRibbon if good
    If Not ThisRibbon Is Nothing Then
        Set GetRibbon = ThisRibbon
        Exit Function
    End If
    
    ' copy pointer
    Dim lngRibPtr As Long
    lngRibPtr = VB_VAR_STORE.GetRibbonID()

    ' copy memory
    Dim objRibbon As Object
    CopyMemory objRibbon, lngRibPtr, 4
    
    ' set returns
    Set GetRibbon = objRibbon
    Set objRibbon = Nothing
End Function

Public Sub RefreshGlobalVariables()
'Subroutine: RefreshGlobalVaribles - When Excel's VB environment breaks for runtime errors, it purges its memory stack, so every
'                                    variable outside of procedure scope is set to 0. This function is an attempt to refresh those variables
'                                    and bring the application back to life.

    If ThisRibbon Is Nothing Then
        Set ThisRibbon = GetRibbon()
    End If
    If Not ThisRibbon Is Nothing Then
        ThisRibbon.Invalidate
    End If
    
    comparison_site = VB_VAR_STORE.GetComparisonSite()
    VB_CHANGE_LOG.SetTrackChanges True
    ThisWorkbook.SetWorkbookClosed False
End Sub

'Callback for customUI.onLoad
Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    Set ThisRibbon = ribbon
    RibbonSaver ribbon
    
    ribbon.ActivateTab "TabBOM"
    
    comparison_site = ""

    Old_CurrentCS_Text = ""
    Old_PrevCS_Text = ""
    Old_HideZeros = True
    CurrentCS_Text = Old_CurrentCS_Text
    PrevCS_Text = Old_PrevCS_Text
    hideZeros = Old_HideZeros
End Sub

'Callback for TabBOM getVisible
Sub BlankOrInventory_GetVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Not (ThisWorkbook.IsBlankBOM() Or VB_CLIENTINVENTORY.Visible = xlSheetVisible)
End Sub

'Callback for add_item_rbutton onAction
Sub Ribbon_InsertRow(control As IRibbonControl)
    Dim newMark As Integer
    
    AddItemWindow.LoadForm newMark
End Sub

'Callback for delete_item_rbutton onAction
Sub Ribbon_DeleteRow(control As IRibbonControl)
    If Not TypeOf Application.Selection Is Range Then Exit Sub
    
    Dim index As Integer
    Dim i_row() As Integer
    ReDim i_row(1 To Application.Selection.Rows.count)
    
    'If BOM is blank, do nothing
    If ThisWorkbook.IsBlankBOM() Then
        Exit Sub
    End If
    
    'adds all selected rows marked for deletion to integer array
    For index = LBound(i_row) To UBound(i_row)
        i_row(index) = Application.Selection.Rows(index).row
    Next index
    
    ' sort
    QuickSort i_row, LBound(i_row), UBound(i_row)
    
    Dim ui_change As Boolean
    Dim result As Integer
    Dim category As String
    
    For index = UBound(i_row) To LBound(i_row) Step -1
        
        If ThisWorkbook.ActiveSheet.Name = VB_MASTER.Name Then
            Dim mark As Integer
            mark = get_mark_num(i_row(index))
            If mark <> -1 Then
                Dim DESC As String
                DESC = get_property(mark, "Long Description")
                
                'check if description exists, or is this a blank category
                If Len(DESC) > 0 Then
                
                    ' check if material has been ordered; if so, cancel
                    If VB_ORDER_LOG.TotalRequested(get_mark_num(i_row(index))) > 0 Then
                        MsgBox "This project contains requested/purchased material for this item, so it cannot be removed." & vbCrLf & vbCrLf & _
                            "If you think you are receiving this message in error, unpublishing the RFPs that contain this ordered material will allow you to remove this item completely from the project. " & vbCrLf & vbCrLf & _
                            "In the meantime, double-click the cell in the ""Delete?"" column corresponding to this item. This operation will remove the model quantities for this item while retaining the RFP records.", vbInformation
                            
                            Exit Sub
                    End If
                
                    result = MsgBox("Are you sure you want to delete " & get_category(i_row(index)) & " Item# " & mark & "? This cannot be undone.", vbYesNo)
                    If result = vbYes Then
                        ui_change = RenderUI(False)
                        
                        category = get_category(i_row(index))
                        DeleteMasterRow i_row(index)
                        
                        If VB_CHANGE_LOG.LogChange(VB_MASTER.Name & "!" & Selection.Cells(index).Address, mark, "Purged Material Item", DESC, category) <> 0 Then
                            'throw error
                            ErrorHandling "Ribbon_DeleteItem", 10, "LogChange(" & VB_MASTER.Name & "!" & VB_MASTER.Cells(i_row(index), get_col_num("Mark No.")).Address & ", " & mark & ", Delete Material Item, " & DESC & ", )", 1
                        End If
                        
                        VB_MASTER.CalculateQuantityFormat
                    End If
                Else
                    'blank item description or blank category
                    Dim cat_start As Integer
                    Dim cat_end As Integer
                    cat_start = get_cat_row(get_category(i_row(index)))
                    cat_end = cat_start + VB_MASTER.Cells(i_row(index), VB_MASTER.CategoryColumn()).MergeArea.count - 1
                    
                    If cat_start = cat_end Then 'only item in category
                        result = MsgBox("There is no item here. Do you wish to delete the entire category?", vbYesNo)
                        
                        If result = vbYes Then 'DELETE WHOLE CATEGORY
                            ui_change = RenderUI(False)
                            DeleteCategory get_category(cat_start)
                            
                            VB_MASTER.CalculateQuantityFormat
                        End If
                    Else
                        'just delete the row. This shouldn't ever happen.
                        ui_change = RenderUI(False)
                        DeleteMasterRow i_row(index)
                    End If
                    
                End If
            Else
                'do nothing
            End If
        Else
            'do nothing
        End If
        
        If ui_change Then RenderUI True
    Next index
End Sub

'Callback for populate_rbutton getEnabled
Sub PopulateNumbers_GetEnabled(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Not (ThisWorkbook.MarkNumLock() Or VB_MASTER.Columns(get_col_num("Mark No.")).Hidden)
End Sub

'Callback for populate_rbutton onAction
Sub Ribbon_PopulateNums(control As IRibbonControl)
    Dim result As Integer
    result = MsgBox("This operation could permanently change every mark number in this BOM? Are you sure you want to continue?", vbYesNo)
    
    If result = vbYes Then
        Dim ui_change As Boolean
        ui_change = RenderUI(False)
        
        Dim init_row As Integer
        Dim prev_mark As Integer
        Dim mark_col As Integer
        Dim row As Integer
        Dim last As Integer
        Dim mark As Integer
        Dim current_category As String
        last = VB_MASTER.LastRow()
        init_row = VB_MASTER.FirstRow()
        mark_col = get_col_num("Mark No.")
        
        Call VB_CHANGE_LOG.LogChange(VB_MASTER.Name & "!" & VB_MASTER.Cells(VB_MASTER.TitleRow(), mark_col).Address, "", _
            "BEGIN Populate Mark Numbers", "", "")
        
        mark = 0
        current_category = get_category(init_row)
        
        'loop through rows, incrementing mark numbers
        For row = init_row To last
            SetStatusBar "Populating Mark Numbers - Row# " & row, (row - init_row), (last - init_row)
            
            prev_mark = get_mark_num(row)
            
            'new mark number
            mark = mark + 1
            If get_category(row) <> current_category Then
                mark = mark + 4
                current_category = get_category(row)
            End If
            
            ' if mark number doesn't match, WriteChange
            If prev_mark <> mark Then
                VB_MASTER.Cells(row, mark_col).Value = mark
                VB_MASTER.WriteChange VB_MASTER.Cells(row, mark_col), prev_mark
            End If
        Next row
        
        Call VB_CHANGE_LOG.LogChange(VB_MASTER.Name & "!" & VB_MASTER.Cells(VB_MASTER.TitleRow(), mark_col).Address, "", _
            "END Populate Mark Numbers", "", "")
        
        ResetStatusBar
        
        If ui_change Then RenderUI True
    Else
        ' do nothing
    End If
End Sub

'Callback for new_cat_rbutton onAction
Sub Ribbon_NewCategory(control As IRibbonControl)
    Dim new_category As String
    
    new_category = NewCategoryHandler()
    
    VB_MASTER.Activate
    If Len(new_category) > 0 Then
        MsgBox "Successfully added category: " & new_category & "."
    End If
End Sub

'Callback for purge_cat_rbutton onAction
Sub Ribbon_PurgeCategory(control As IRibbonControl)
    Dim successful As Boolean
    successful = False
    
    If ThisWorkbook.ActiveSheet.Name = VB_MASTER.Name Then
        successful = PurgeCategoryForm.LoadForm(Selection.row)
    Else
        successful = PurgeCategoryForm.LoadForm
    End If
    
    If successful Then
        MsgBox "Successfully Purged Category."
    End If
End Sub

'Callback for reorder_cat_rbutton onAction
Sub Ribbon_ReorderCategories(control As IRibbonControl)
    ReorderCategoriesForm.LoadForm
End Sub

'Callback for longlead_rbutton getPressed
Sub ShowLongLead_GetPressed(control As IRibbonControl, ByRef returnedVal)
    Dim first As Integer
    Dim last As Integer
    first = VB_MASTER.FirstRow()
    last = VB_MASTER.LastRow()
    
    Dim longlead As Integer
    longlead = get_col_num("Long Lead")
    
    Dim row As Integer
    
    returnedVal = False
    
    ' find hidden rows that aren't long lead
    For row = first To last
        If VB_MASTER.Rows(row).Hidden And Not CBool(VB_MASTER.Cells(row, longlead).Value2) Then
            returnedVal = True
            Exit Sub
        End If
    Next row
End Sub

'Callback for longlead_rbutton onAction
Sub Ribbon_ShowLongLead(control As IRibbonControl, pressed As Boolean)
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    VB_MASTER.Activate
    
    Dim first As Integer
    Dim last As Integer
    first = VB_MASTER.FirstRow()
    last = VB_MASTER.LastRow()
    
    Dim longlead As Integer
    longlead = get_col_num("Long Lead")
    
    Dim row As Integer
    
    ' hide rows that aren't long lead
    For row = first To last
        If pressed Then
            VB_MASTER.Rows(row).Hidden = Not CBool(VB_MASTER.Cells(row, longlead))
        Else
            VB_MASTER.Rows(row).Hidden = False
        End If
    Next row
    
    If ui_change Then RenderUI True
End Sub

'Callback for numlock_rbutton getLabel
Sub NumLock_GetLabel(control As IRibbonControl, ByRef returnedVal)
    If ThisWorkbook.MarkNumLock() Then
        returnedVal = "Item Numbers are Locked"
    Else
        returnedVal = "Item Numbers are Unlocked"
    End If
End Sub

'Callback for numlock_rbutton getPressed
Sub NumLock_GetPressed(control As IRibbonControl, ByRef returnedVal)
    If returnedVal Then
        ThisWorkbook.SetMarkNumLock True
    End If
    
    returnedVal = ThisWorkbook.MarkNumLock()
End Sub

'Callback for numlock_rbutton getEnabled
Sub NumLock_GetEnabled(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Not VB_MASTER.Columns(get_col_num("Mark No.")).Hidden
End Sub

'Callback for numlock_rbutton onAction
Sub Ribbon_NumLock(control As IRibbonControl, pressed As Boolean)
    ThisWorkbook.SetMarkNumLock pressed
    
    Dim response As String
    If pressed Then
        response = "LOCKED"
    Else
        response = "UNLOCKED"
    End If
    
    If VB_CHANGE_LOG.TrackChanges() Then
        If VB_CHANGE_LOG.LogChange(VB_CHANGE_LOG.Name & "!$A$1", "", response & " Mark Numbers", CStr(Not pressed), CStr(pressed)) <> 0 Then
            'throw error
            ErrorHandling "Ribbon_NumLock", 0, "LogChange() operation failed:" & vbCrLf & _
                "LogChange(" & VB_CHANGE_LOG.Name & "!$A$1, , " & response & " Mark Numbers, " & CStr(Not pressed) & ", " & CStr(pressed) & ")" & vbCrLf & _
                "Continue [OK] or Cancel procedure?", 1
        End If
    End If
    
    If ThisRibbon Is Nothing Then
        RefreshGlobalVariables
    End If
    ThisRibbon.Invalidate
End Sub

'Callback for addsite_rbutton onAction
Sub Ribbon_AddSite(control As IRibbonControl)
    Dim new_site_name As String
    NewSiteForm.LoadForm new_site_name
    
    If SiteExists(new_site_name) Then
        MsgBox "Successfully added " & new_site_name & "."
    End If
End Sub

'Callback for removesite_rbutton onAction
Sub Ribbon_RemoveSite(control As IRibbonControl)
    Dim rSites As Range
    Set rSites = GetSitesRange()
    
    If Not IsEmpty(rSites.Cells(1, 1)) Then
        Dim site As String
        
        ChooseSpecificSite.LoadForm site
        
        If Len(site) > 0 Then
            Dim result As Integer
            result = MsgBox("Are you sure you want to permanently remove " & site & " from this project?", vbOKCancel)
            
            If result = vbOK Then
                Dim order As Integer
                order = 0
                Do While Not VB_ORDER_LOG.EOF(VB_ORDER_LOG.FirstRow() + order)
                    If VB_ORDER_LOG.get_order_site(VB_ORDER_LOG.OrderNumber(order)) = site Then
                        result = MsgBox("This project contains an RFP specifically pretaining to this site. Once purged, material ordered on this RFP will be considered " & _
                            "general project supply, and all model quantity information for this site will be lost. This loss cannot be recovered. " & _
                            "Do you still wish to purge " & site & " from this project?", vbYesNo)
                            
                        If result = vbNo Then
                            Exit Sub
                        Else
                            Exit Do
                        End If
                    End If
                    
                    order = order + 1
                Loop
                
                Dim ui_change As Boolean
                ui_change = RenderUI(False)
                
                Dim site_col As Integer
                SiteExists site, site_col
                
                RemoveSite site
                
                If ui_change Then RenderUI True
            End If
        End If
    Else
        MsgBox "No sites in this project.", vbInformation
    End If
    
    Set rSites = Nothing
End Sub

'Callback for renamesite_rbutton onAction
Sub Ribbon_RenameSite(control As IRibbonControl)
    Dim rSites As Range
    Set rSites = GetSitesRange()
    
    If Not IsEmpty(rSites.Cells(1, 1)) Then
        Dim old_site_name As String
        ChooseSpecificSite.LoadForm old_site_name
        
        If Len(old_site_name) > 0 Then 'user didn't close the popup
            Dim new_site_name As String
            NewSiteForm.LoadForm new_site_name, old_site_name
        End If
        If SiteExists(new_site_name) Then
            MsgBox "Successfully renamed " & old_site_name & " as " & new_site_name & "."
        End If
    Else
        MsgBox "No sites in this project.", vbInformation
        Exit Sub
    End If

    Set rSites = Nothing
End Sub

'Callback for GrpDRFT getVisible
Sub BlankGroup_GetVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Not ThisWorkbook.IsBlankBOM()
End Sub

'Callback for cad_rbutton getPressed
Sub CADImport_GetPressed(control As IRibbonControl, ByRef returnedVal)
    Dim cadSheet As Worksheet
    On Error Resume Next
    Set cadSheet = ThisWorkbook.Sheets("CADWORX_IMPORT")
    Err.Clear
    On Error GoTo 0
    
    returnedVal = Not (cadSheet Is Nothing)
    
    Set cadSheet = Nothing
End Sub

'Callback for cad_rbutton onAction
Sub Ribbon_CADImport(control As IRibbonControl, pressed As Boolean)
    Dim rSites As Range
    Dim cadSheet As Worksheet
    On Error Resume Next
    Set cadSheet = ThisWorkbook.Sheets("CADWORX_IMPORT")
    Err.Clear
    On Error GoTo 0
    
    If pressed And cadSheet Is Nothing Then
        ' get CAD filename and import sheet
        
        Dim filename As String
        filename = GetCADFileName()
        
        If Len(filename) = 0 Then
            pressed = False
        Else
            'import
            If ImportCADSheet(filename) <> 0 Then
                pressed = False
                GoTo INVALIDATETHIS
            End If
        End If
        
    ElseIf Not pressed And Not cadSheet Is Nothing Then
        Dim result As Integer
        result = CADWorxImportDialogue.LoadForm
                    
        If result = vbYes Then
            'Continue with CADWorx Import
            'CAD sheet exists, import to MASTER BOM
            Set rSites = GetSitesRange()
            
            Dim site_col As Integer
            Dim site_name As String
            site_name = vbNullString
            site_col = -1
            
            If Not IsEmpty(rSites.Cells(1, 1)) Then
                ChooseSpecificSite.LoadForm site_name
                
                If Len(site_name) > 0 Then 'user didn't close the popup
                    SiteExists site_name, site_col
                Else
                    pressed = True
                    GoTo INVALIDATETHIS
                End If
            Else
                MsgBox "Create Site first please.", vbExclamation
                pressed = True
                GoTo INVALIDATETHIS
            End If
            
            'finalize
            FinalizeImport site_col
            
        ElseIf result = vbNo Then
            'Back out of CADWorx Import completely
            Dim ui_change As Boolean
            ui_change = RenderUI(False)
            
            Application.DisplayAlerts = False
            cadSheet.Delete
            Application.DisplayAlerts = True
            
            VB_MASTER.Activate
            pressed = Not (cadSheet Is Nothing)
            If ui_change Then RenderUI True
            
            GoTo INVALIDATETHIS
        Else
            'Close this dialogue.
            pressed = Not (cadSheet Is Nothing)
            GoTo INVALIDATETHIS
        End If
        
    Else
        pressed = Not (cadSheet Is Nothing)
        GoTo INVALIDATETHIS
    End If
    
INVALIDATETHIS:
    Set rSites = Nothing
    Set cadSheet = Nothing
    
    If ThisRibbon Is Nothing Then
        RefreshGlobalVariables
    End If
    ThisRibbon.Invalidate
End Sub

'Callback for sitebom_rbutton getEnabled
Sub SiteBOMs_GetEnabled(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (Len(comparison_site) = 0)
End Sub

'Callback for sitebom_rbutton onAction
Sub Ribbon_SiteBOMs(control As IRibbonControl)
    ChooseSiteBOMs.LoadForm comparison_site
    
    If ThisRibbon Is Nothing Then
        RefreshGlobalVariables
    End If
    ThisRibbon.Invalidate
End Sub

'Callback for marknums_rbox getPressed
Sub SuppressMarkNums_GetPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = VB_MASTER.Columns(get_col_num("Mark No.")).Hidden
    
    If VB_ORDER_TMPLT.Columns(VB_ORDER_TMPLT.Form_col_num("Item #")).Hidden <> returnedVal Then
        Ribbon_SuppressMarkNums control, CBool(returnedVal)
    End If
End Sub

'Callback for marknums_rbox onAction
Sub Ribbon_SuppressMarkNums(control As IRibbonControl, pressed As Boolean)
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    VB_MASTER.Columns(get_col_num("Mark No.")).Hidden = pressed
    VB_ORDER_TMPLT.Columns(VB_ORDER_TMPLT.Form_col_num("Item #")).Hidden = pressed
    
    If ThisRibbon Is Nothing Then
        RefreshGlobalVariables
    End If
    ThisRibbon.Invalidate
    
    If ui_change Then RenderUI True
End Sub

'Callback for punchlist_rbutton getEnabled
Sub DraftingPunchlist_GetEnabled(control As IRibbonControl, ByRef returnedVal)
    Dim cadSheet As Worksheet
    On Error Resume Next
    Set cadSheet = ThisWorkbook.Sheets("CADWORX_IMPORT_OLD")
    Err.Clear
    On Error GoTo 0

    returnedVal = Not (cadSheet Is Nothing)
End Sub

'Callback for punchlist_rbutton onAction
Sub Ribbon_DraftingPunchlist(control As IRibbonControl)
    DraftingPunchList
End Sub

'Callback for punchlist_rbutton onAction
Sub Ribbon_DraftingCheckCopy(control As IRibbonControl)
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    ChooseCheckCopySites.LoadForm
    
    VB_DFTGCHECK.PrepareDraftingCheckCopy
    
    VB_DFTGCHECK.Visible = xlSheetVisible
    Print2PDF VB_DFTGCHECK.Name, "\Check Copies\", VB_DFTGCHECK.Name & HMMFileTag()
    VB_DFTGCHECK.Visible = xlSheetVeryHidden
    
    VB_DFTGCHECK.EmptyCheckCopy
    
    If ui_change Then RenderUI True
End Sub

'Callback for GrpRFP getVisible
Sub InventoryMgmt_GetEnabled(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (VB_CLIENTINVENTORY.Visible = xlSheetVeryHidden)
End Sub

'Callback for lltotal_rbutton onAction
Sub Ribbon_LongLead_Total(control As IRibbonControl)
    AutoCompleteRFP False, False, True
End Sub

'Callback for lldeficit_rbutton onAction
Sub Ribbon_LongLead_Deficit(control As IRibbonControl)
    AutoCompleteRFP False, True, True
End Sub

'Callback for llsitetotal_rbutton onAction
Sub Ribbon_LongLead_SiteTotal(control As IRibbonControl)
    AutoCompleteRFP True, False, True
End Sub

'Callback for llsitedeficit_rbutton onAction
Sub Ribbon_LongLead_SiteDeficit(control As IRibbonControl)
    AutoCompleteRFP True, True, True
End Sub

'Callback for projecttotal_rbutton onAction
Sub Ribbon_All_Total(control As IRibbonControl)
    AutoCompleteRFP False, False, False
End Sub

'Callback for projectdeficit_rbutton onAction
Sub Ribbon_All_Deficit(control As IRibbonControl)
    AutoCompleteRFP False, True, False
End Sub

'Callback for sitetotal_rbutton onAction
Sub Ribbon_All_SiteTotal(control As IRibbonControl)
    AutoCompleteRFP True, False, False
End Sub

'Callback for sitedeficit_rbutton onAction
Sub Ribbon_All_SiteDeficit(control As IRibbonControl)
    AutoCompleteRFP True, True, False
End Sub

Private Sub AutoCompleteRFP(ByVal site_specific As Boolean, ByVal deficit As Boolean, ByVal longlead As Boolean)
'Subroutine: AutoCompleteRFP - This is a procedure for controlling autocompletion of the RPF Form. It handles
'                              site specific RFPs, deficit only RFPS, and long lead item RFPs.
'Arguments: site_specific - Boolean controlling the site specific nature of RFPs.
'           deficit - Boolean telling the procedure whether or not to calculate deficit quantities for items.
'           longlead - Boolean telling the procedure whether or not to include only long lead items or all items.


    Dim ui_change As Boolean
    ui_change = False
    
    If Not VB_ORDER_TMPLT.IsFormEmpty() Then
        VB_ORDER_TMPLT.Activate
        Dim result As Integer
        result = MsgBox("Excel has detected that this form is not empty. The following operation will permanently discard the information in this form. Would you like to continue?", vbYesNo)
        
        If result = vbYes Then
            'continue
            ui_change = RenderUI(False)
    
            VB_ORDER_TMPLT.ClearForm
        Else
            Exit Sub
        End If
    End If
    
    If Not ui_change Then ui_change = RenderUI(False)
    
    Dim rSites As Range
    Dim RFPsite As String
    RFPsite = vbNullString
    
    ' remember original sheet in case of an error
    Dim active_sheet As String
    active_sheet = ActiveSheet.Name
    
    ' change sheet to VB_ORDER_TMPLT
    If ActiveSheet.Name <> VB_ORDER_TMPLT.Name Then
        VB_ORDER_TMPLT.Activate
    End If
    
    VB_ORDER_TMPLT.Cells(7, 5).ClearContents
    VB_ORDER_TMPLT.Rows(7).Hidden = Not site_specific
    
    If site_specific Then
        ' grab site from user
        
        Set rSites = GetSitesRange()
    
        If Not IsEmpty(rSites.Cells(1, 1)) Then
            Dim temp_site As String
            ChooseSpecificSite.LoadForm temp_site
            
            If Len(temp_site) > 0 Then
                RFPsite = temp_site
            Else
                ' form was exited to end the process..
                ThisWorkbook.Sheets(active_sheet).Activate
                GoTo FINISH_AUTORFP
            End If
        Else
            MsgBox "No sites in this project. Cannot create site specific RFP.", vbExclamation
            ThisWorkbook.Sheets(active_sheet).Activate
            GoTo FINISH_AUTORFP
        End If
    End If
    
    ' write items
    Dim items_written As Integer
    items_written = VB_ORDER_TMPLT.AutoCompleteForm(RFPsite, deficit, longlead)
    
    If items_written = 0 Then
        VB_ORDER_TMPLT.ClearForm
        MsgBox "No items to report.", vbInformation
    ElseIf items_written = -1 Then 'AutoCompleteForm Error
        ErrorHandling "AutoCompleteRFP", -1, "AutoCompleteForm() failure:" & vbCrLf & "AutoCompleteForm(" & RFPsite & ", " & deficit & ")", 1
        VB_ORDER_TMPLT.ClearForm
    End If
    
FINISH_AUTORFP:
    Set rSites = Nothing
    
    If ui_change Then RenderUI True
End Sub

'Callback for clearform_rbutton onAction
Sub Ribbon_ClearForm(control As IRibbonControl)
    Dim ui_change As Boolean
    ui_change = False

    VB_ORDER_TMPLT.Activate
    If Not VB_ORDER_TMPLT.IsFormEmpty() Then
        Dim result As Integer
        result = MsgBox("This operation will permanently discard the information currently entered in this form. Are you sure you would like to continue?", vbYesNo)
        
        If result = vbYes Then
            'continue
            ui_change = RenderUI(False)
            
            VB_ORDER_TMPLT.ClearForm
        Else
            Exit Sub
        End If
    Else
        If VB_ORDER_TMPLT.Rows(7).Hidden Then
            VB_ORDER_TMPLT.Rows(7).Hidden = False
        End If
    End If
    
    If ui_change Then RenderUI True
End Sub

'Callback for ordermgr_rbutton onAction
Sub Ribbon_OrderManager(control As IRibbonControl)
    OrderTracking.LoadForm
End Sub

'Callback for publishRFPform_rbutton onAction
Sub Ribbon_RFPForm(control As IRibbonControl)
    VB_ORDER_TMPLT.Activate
    VB_ORDER_TMPLT.OrderFormImport
End Sub

'Callback for delete_order_rbutton onAction
Sub Ribbon_DeleteOrder(control As IRibbonControl)
    If VB_ORDER_LOG.NumOrders() = 0 Then
        MsgBox "There are no RFPs to unpublish at this time.", vbInformation
        Exit Sub
    End If
    
    'CHOOSE ORDER
    Dim order_num As String
    DeleteOrderForm.LoadForm order_num
    
    If Len(order_num) = 0 Then 'operation canceled
        Exit Sub
    End If
    
    'CHECK WITH USER
    Dim result As Integer
    result = MsgBox("Are you sure you want to permanently delete Order #" & order_num & "?", vbOKCancel)
    If result <> vbOK Then
        Exit Sub
    End If
    
    'UNPUBLISH
    On Error Resume Next
    VB_ORDER_LOG.DeleteOrder (order_num)
    Err.Clear
    On Error GoTo 0
End Sub

'Callback for inventorymgmt_rbutton onAction
Sub Ribbon_InventoryMgmt(control As IRibbonControl)
    VB_CLIENTINVENTORY.LoadForm
    
    If ThisRibbon Is Nothing Then
        RefreshGlobalVariables
    End If
    ThisRibbon.Invalidate
End Sub

'Callback for sap_rbox getPressed
Sub IncludeSAP_GetPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Not VB_MASTER.Columns(get_col_num("SAP#")).Hidden
    
    If VB_SITEBOM.Columns(1).Hidden <> Not returnedVal _
        Or VB_RFP_REPORT.Columns(VB_RFP_REPORT.Form_col_num(VB_RFP_REPORT, "Mark No.") - 1).Hidden <> Not returnedVal Then
         
        Ribbon_IncludeSAP control, CBool(returnedVal)
    End If
End Sub

'Callback for sap_rbox onAction
Sub Ribbon_IncludeSAP(control As IRibbonControl, pressed As Boolean)
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    ' SAP# on MASTER
    VB_MASTER.Columns(get_col_num("SAP#")).Hidden = Not pressed
    
    ' SAP# on Site BOMs
    Dim wrkSheet
    For Each wrkSheet In ThisWorkbook.Sheets
        If wrkSheet.CodeName Like "VB_SITEBOM" Then
            wrkSheet.Columns(1).Hidden = Not pressed
        End If
    Next wrkSheet
    
    ' SAP# on Client Summary
    VB_RFP_REPORT.Columns(VB_RFP_REPORT.Form_col_num(VB_RFP_REPORT, "Mark No.") - 1).Hidden = Not pressed
    
    If ui_change Then RenderUI True
End Sub

'Callback for consumables_rbox getPressed
Sub TrackConsumables_GetPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Not VB_MASTER.Columns(get_col_num("Total Extras")).Hidden
    
    If VB_RFP_REPORT.Columns(VB_RFP_REPORT.Form_col_num(VB_RFP_REPORT, "Consumables")).Hidden <> Not returnedVal Then
        Ribbon_TrackConsumables control, CBool(returnedVal)
    End If
End Sub

'Callback for consumables_rbox onAction
Sub Ribbon_TrackConsumables(control As IRibbonControl, pressed As Boolean)
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    ' Note: DO NOT CLEAR CONTENTS OF CONSUMABLES SECTION. These values are saved
    '   if for some reason this action was performed by mistake.
    '''''''''''''''''''''''''''
    
    ' Consumables on MASTER
    VB_MASTER.Cells(VB_MASTER.SectionRow(), get_col_num("Model Extras")).MergeArea.Columns.Hidden = Not pressed
    With VB_MASTER.Columns(get_col_num("Checked Quantities")).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(150, 54, 52)
    End With
    
    Dim first As Integer
    Dim last As Integer
    first = VB_MASTER.FirstRow()
    last = VB_MASTER.LastRow()
    Dim row As Integer
    
    'Change formulas: Total Extras = 0
    For row = first To last
        SetRowFormulas get_mark_num(row)
    Next row
    
    ' Consumables on Client Summary
    VB_RFP_REPORT.Columns(VB_RFP_REPORT.Form_col_num(VB_RFP_REPORT, "Consumables")).Hidden = Not pressed
    
    If ui_change Then RenderUI True
End Sub

'Callback for issueRFPreport_rbutton onAction
Sub Ribbon_RFPReport(control As IRibbonControl)
    
    If Len(VB_COVERSHEET.HMMContact) = 0 Then
        Dim vbResult As Integer
        vbResult = MsgBox("There is no information given on the Coversheet about an HMM Contact. " & _
            "This information should appear on the Client Summary, although you may exclude it and continue if you like.", vbOKCancel)
            
        If vbResult = vbCancel Then
            Exit Sub
        End If
    End If
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    Dim current_sheet As String
    current_sheet = ThisWorkbook.ActiveSheet.Name
    
    On Error GoTo EmptyThisReport
    VB_RFP_REPORT.Prepare4Publish
    VB_RFP_REPORT.Visible = xlSheetVisible
    VB_RFP_REPORT.Activate
    SetStatusBar "Waiting for User Approval..."
    
    Old_CurrentCS_Text = Str(VB_VAR_STORE.GetClientSummaryRev())
    Old_PrevCS_Text = Str(VB_VAR_STORE.GetClientSummaryRev() - 1)
    Old_HideZeros = hideZeros
    CurrentCS_Text = Old_CurrentCS_Text
    PrevCS_Text = Old_PrevCS_Text
    hideZeros = Old_HideZeros
    
    If ui_change Then RenderUI True
    'HOLD
    'Wait for user Approval
    
    If ThisRibbon Is Nothing Then
        RefreshGlobalVariables
    End If
    ThisRibbon.Invalidate
    
    Exit Sub
    
EmptyThisReport:
    'Empty report to restart for next time.
    VB_RFP_REPORT.EmptyReport
    ThisWorkbook.Sheets(current_sheet).Activate
    
    If ui_change Then RenderUI True
    ResetStatusBar
End Sub

'Callback for GrpInventory getVisible
Sub GrpInventory_GetVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (VB_CLIENTINVENTORY.Visible = xlSheetVisible)
End Sub

'Callback for inventorysave_rbutton onAction
Sub Ribbon_InventorySave(control As IRibbonControl)
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    VB_CLIENTINVENTORY.SaveInventoryForm
    VB_CLIENTINVENTORY.CloseForm
    
    If ui_change Then RenderUI True
    
    If ThisRibbon Is Nothing Then
        RefreshGlobalVariables
    End If
    ThisRibbon.Invalidate
End Sub

'Callback for inventorydiscard_rbutton onAction
Sub Ribbon_InventoryDiscard(control As IRibbonControl)
    VB_CLIENTINVENTORY.CloseForm
    
    If ThisRibbon Is Nothing Then
        RefreshGlobalVariables
    End If
    ThisRibbon.Invalidate
End Sub

'Callback for comparesite_rbutton getPressed
Sub SiteCompare_GetPressed(control As IRibbonControl, ByRef returnedVal)
    If returnedVal Then
        comparison_site = VB_VAR_STORE.GetComparisonSite()
    End If
    
    VB_VAR_STORE.SetComparisonSite comparison_site
    returnedVal = (Len(comparison_site) > 0)
End Sub

'Callback for comparesite_rbutton onAction
Sub Ribbon_SiteCompare(control As IRibbonControl, pressed As Boolean)
    Dim ui_change As Boolean
    
    If pressed And Len(comparison_site) = 0 Then
    
        ChooseSpecificSite.LoadForm comparison_site
        VB_VAR_STORE.SetComparisonSite comparison_site
        
        If Not SiteExists(comparison_site) Then
            pressed = False
            comparison_site = vbNullString
            GoTo INVALIDATETHIS
        End If
        
        If Not SheetExists(comparison_site & " - BOM") Then ' sheet doesn't exist
            MsgBox "Specific BOM's for this site don't exist yet. Please take advantage of the ""Create Site/Area BOM"" feature first.", vbInformation
        End If
        
        'user choose Previous Site BOM
        Dim filename As String
        filename = GetPreviousSiteBOMFileName()
        
        'cancelled?
        If Len(filename) = 0 Then
            'cancel
            pressed = False
            comparison_site = vbNullString
            GoTo INVALIDATETHIS
        ElseIf Not filename Like "* - BOM*" Then
            MsgBox "This filename suggests that this file does not relate the site specific BOMs generated by this application. " & vbCrLf & _
                "The filename should look like [site name] - BOM[...]_rev#.xlsx. This opeation cannot continue.", vbInformation
            
            pressed = False
            comparison_site = vbNullString
            GoTo INVALIDATETHIS
        End If
        
        'Open workbook
        Dim bom_wkbk As Workbook
        Dim prevBOM As Worksheet
        
        Set bom_wkbk = Application.Workbooks.Open(filename:=filename, ReadOnly:=True)
        Set prevBOM = bom_wkbk.Sheets(1)
        
        ui_change = RenderUI(False)
        
        SiteBOMCompare comparison_site, prevBOM
    
        Set prevBOM = Nothing
        bom_wkbk.Close False
        Set bom_wkbk = Nothing

    ElseIf Not pressed And SiteExists(comparison_site) Then 'unpressed
        
        Dim site_sheet_name As String
        If Len(comparison_site) > 25 Then
            site_sheet_name = Left(comparison_site, 25) & " - BOM"
        Else
            site_sheet_name = comparison_site & " - BOM"
        End If
        
        Dim sheet As Worksheet
        Set sheet = ThisWorkbook.Sheets(site_sheet_name)
       
        ui_change = RenderUI(False)
        
        ' unhighlight/unhide everything
        sheet.UsedRange.Columns.Hidden = False
        sheet.UsedRange.Interior.Color = RGB(255, 255, 255)
        
        ' delete (OBSOLETE) side
        sheet.Columns(ColLet(sheet.UsedRange.Columns.count / 2 + 1) & ":" & ColLet(sheet.UsedRange.Columns.count)).Delete
        
        ' delete empty rows from arranging
        Dim row As Integer
        For row = VB_SITEBOM.FirstRow() To sheet.UsedRange.Rows.CountLarge
            If IsEmpty(sheet.Cells(row, 2)) And Not IsEmpty(sheet.Cells(row, sheet.UsedRange.Columns.count / 2 + 2)) Then
                sheet.Rows(row).Delete
                row = row - 1
            End If
        Next row

        comparison_site = vbNullString
        VB_VAR_STORE.ClearComparisonSite
        
        sheet.Cells(2, 1).MergeArea.Select
    Else
        pressed = (Len(comparison_site) > 0)
        GoTo INVALIDATETHIS
    End If
    
INVALIDATETHIS:
    If ui_change Then RenderUI True
    Set bom_wkbk = Nothing
    Set prevBOM = Nothing
    
    If ThisRibbon Is Nothing Then
        RefreshGlobalVariables
    End If
    ThisRibbon.Invalidate
End Sub

'Callback for reset_check_rbutton onAction
Sub Ribbon_ResetDescCheck(control As IRibbonControl)
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    Dim mark As Integer
    Dim last As Integer
    last = MaxMark()
    
    Call VB_CHANGE_LOG.LogChange(VB_MASTER.Name & "!" & VB_MASTER.Cells(VB_MASTER.TitleRow(), get_col_num("Description Check")).Address, "", _
        "BEGIN Description Check Reset", "", "")
    
    For mark = 1 To last
        If get_row(mark) <> -1 Then
            UnapproveDescription MarkNum:=mark
        End If
    Next mark
    
    Call VB_CHANGE_LOG.LogChange(VB_MASTER.Name & "!" & VB_MASTER.Cells(VB_MASTER.TitleRow(), get_col_num("Description Check")).Address, "", _
        "END Description Check Reset", "", "")
    
    If ui_change Then RenderUI True
End Sub

'Callback for reset_quantity_rbutton onAction
Sub Ribbon_ResetQuanCheck(control As IRibbonControl)
    Dim rSites As Range
    Set rSites = GetSitesRange()
    
    If Not IsEmpty(rSites.Cells(1, 1)) Then
        Dim chk_site As String
        ChooseSpecificSite.LoadForm chk_site
        
        Dim site_col As Integer
        If SiteExists(chk_site, site_col) Then
            Dim ui_change As Boolean
            ui_change = RenderUI(False)
            
            site_col = site_col + 2 * rSites.Cells.count + 1
            
            Call VB_CHANGE_LOG.LogChange(VB_MASTER.Name & "!" & VB_MASTER.Cells(VB_MASTER.SubtitleRow(), site_col).Address, "", _
                "BEGIN Quantity Check Reset for " & VB_MASTER.Cells(VB_MASTER.SubtitleRow(), site_col).Value2, "", "")
            
            Dim row As Integer
            Dim first As Integer
            Dim last As Integer
            Dim oldVal
            first = VB_MASTER.FirstRow()
            last = VB_MASTER.LastRow()
            
            For row = first To last
                If Not IsEmpty(VB_MASTER.Cells(row, site_col)) Then
                    oldVal = VB_MASTER.Cells(row, site_col).Value2
                    VB_MASTER.Cells(row, site_col).ClearContents
                    VB_MASTER.WriteChange VB_MASTER.Cells(row, site_col), oldVal
                End If
            Next row
            
            Call VB_CHANGE_LOG.LogChange(VB_MASTER.Name & "!" & VB_MASTER.Cells(VB_MASTER.SubtitleRow(), site_col).Address, "", _
                "END Quantity Check Reset for " & VB_MASTER.Cells(VB_MASTER.SubtitleRow(), site_col).Value2, "", "")
            
            VB_MASTER.CalculateQuantityFormat
            
            If ui_change Then RenderUI True
        End If
    Else
        MsgBox "No sites in this project.", vbInformation
    End If
End Sub

'Callback for undo_rbutton getEnabled
Sub Undo_GetEnabled(control As IRibbonControl, ByRef returnedVal)
    Dim changelog As Worksheet
    
    On Error Resume Next
    Set changelog = ThisWorkbook.Worksheets("Change Log")
    Err.Clear
    On Error GoTo 0
    
    returnedVal = changelog Is Nothing And _
        (VB_CHANGE_LOG.LastChangeID() <> -1)
End Sub

'Callback for undo_rbutton onAction
Sub Ribbon_Undo(control As IRibbonControl)
    VB_CHANGE_LOG.Set_inProgress False
    VB_CHANGE_LOG.HMMUndo
End Sub

'Callback for redo_rbutton getEnabled
Sub Redo_GetEnabled(control As IRibbonControl, ByRef returnedVal)
    Dim changelog As Worksheet
    
    On Error Resume Next
    Set changelog = ThisWorkbook.Worksheets("Change Log")
    Err.Clear
    On Error GoTo 0
    
    returnedVal = changelog Is Nothing And _
        Not VB_CHANGE_LOG.EOF(VB_CHANGE_LOG.LastChangeID() + 1) And VB_CHANGE_LOG.LastChangeID() <> -1
End Sub

'Callback for redo_rbutton onAction
Sub Ribbon_Redo(control As IRibbonControl)
    VB_CHANGE_LOG.Set_inProgress False
    VB_CHANGE_LOG.HMMRedo
End Sub

'Callback for changelog_rbutton getPressed
Sub ChangeLog_GetPressed(control As IRibbonControl, ByRef returnedVal)
    Dim changelog As Worksheet
    
    On Error Resume Next
    Set changelog = ThisWorkbook.Worksheets("Change Log")
    Err.Clear
    On Error GoTo 0
    
    returnedVal = Not (changelog Is Nothing)
End Sub

'Callback for changelog_rbutton onAction
Sub Ribbon_ChangeLog(control As IRibbonControl, pressed As Boolean)
    Dim changelog As Worksheet
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    If pressed Then
        If Not SheetExists("Change Log") Then
            VB_CHANGE_LOG.Visible = xlSheetVisible
            VB_CHANGE_LOG.Copy After:=VB_CHANGE_LOG
            Set changelog = ThisWorkbook.ActiveSheet
            VB_CHANGE_LOG.Visible = xlSheetVeryHidden
            changelog.Name = "Change Log"
            changelog.UsedRange.Rows(1).AutoFilter
            Application.ScreenUpdating = True
            changelog.UsedRange.Rows(VB_CHANGE_LOG.LastChangeID()).Select
            Application.ScreenUpdating = False
        End If
    Else
        Dim current_sheet_name As String
        If ThisWorkbook.ActiveSheet.Name <> "Change Log" Then
            current_sheet_name = ThisWorkbook.ActiveSheet.Name
        Else
            current_sheet_name = VB_MASTER.Name
        End If
        
        If SheetExists("Change Log") Then
            Application.DisplayAlerts = False
            ThisWorkbook.Sheets("Change Log").Delete
            Application.DisplayAlerts = True
        End If
        
        ThisWorkbook.Sheets(current_sheet_name).Activate
    End If
    
    If ui_change Then RenderUI True
    
    If ThisRibbon Is Nothing Then
        RefreshGlobalVariables
    End If
    ThisRibbon.Invalidate

End Sub

'Callback for filetag_rbutton onAction
Sub Ribbon_FileTagFormat(control As IRibbonControl)
    FilenameTagForm.LoadForm
End Sub

'Callback for folderoptions_rbutton onAction
Sub Ribbon_FolderOptions(control As IRibbonControl)
    FolderOptions.LoadForm
End Sub

'Callback for import_old_version_rbutton getVisible
Sub ImportOldVersion_GetVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = ThisWorkbook.IsBlankBOM()
End Sub

'Callback for import_old_version_rbutton onAction
Sub Ribbon_ImportOldVersion(control As IRibbonControl)
    If Not ThisWorkbook.IsBlankBOM() Then
        Exit Sub
    End If
    
    'user choose BOM
    Dim filename As String
    filename = GetPreviousBOMFileName()
    
    'cancelled?
    If Len(filename) = 0 Then
        'cancel
        Exit Sub
    End If
    
    ' ImportPreviousBOM doesn't copy file or folders. To ensure correct continued operation,
    ' ThisWorkbook should be in the same directory as the original
    If Not (filename Like ThisWorkbook.Path & "\*") Then
        MsgBox "New version must be in the same directory as the Previous Version. Please move the BLANK version to the appropriate BOM folder.", vbExclamation
        Exit Sub
    End If
    
    If filename = ThisWorkbook.Path & "\" & ThisWorkbook.Name Then
        MsgBox "The BOM you selected to import is the BOM you now have open. You cannot import a BOM into itself.", vbExclamation
        Exit Sub
    End If
    
    'Open workbook
    If ImportPreviousBOM(filename) <> 0 Then
        ' rollback, empty BOM
        'BlankBOM
    End If
    
    VB_COVERSHEET.Activate
End Sub

'Callback for blank_bom_rbutton getVisible
Sub BlankBOM_GetVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = Not ThisWorkbook.IsBlankBOM()
End Sub

'Callback for blank_bom_rbutton onAction
Sub Ribbon_BlankBOM(control As IRibbonControl)
    Dim result As Integer
    result = MsgBox("This is a premanent action. This document will lose the data that in now has. There will be no sites, no RFPs, no materials items, etc." & _
        vbCrLf & vbCrLf & "This action is reserved for preparing a document for future use on a new project. Before you continue, make sure you have " & _
        "this document saved elsewhere as this process will permanently discard your previous work in the document." & _
        vbCrLf & vbCrLf & "Are you sure you wish to continue?", vbYesNo)
        
    If result = vbYes Then
        ThisWorkbook.BlankBOM
    End If
End Sub

'Callback for refresh_rbutton onAction
Sub Ribbon_Refresh(control As IRibbonControl)
    RenderUI False
    
    ThisWorkbook.OpenWorkbook
    ResetStatusBar
    RefreshGlobalVariables

    RenderUI True
End Sub

'Callback for TabClientSummary getVisible
Sub TabClientSummary_GetVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (VB_RFP_REPORT.Visible = xlSheetVisible)
    
    If returnedVal Then
        ThisRibbon.ActivateTab control.ID
    End If
End Sub

'Callback for applychanges_rbutton getVisible
Sub ApplyChangesCS_GetVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (CurrentCS_Text <> Old_CurrentCS_Text Or PrevCS_Text <> Old_PrevCS_Text Or Old_HideZeros <> hideZeros)
End Sub

'Callback for applychanges_rbutton onAction
Sub Ribbon_ApplyChanges(control As IRibbonControl)
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    Dim rev As Integer
    If Old_CurrentCS_Text <> CurrentCS_Text Then
        rev = CInt(CurrentCS_Text)
        
        VB_VAR_STORE.SetClientSummaryRev rev
        Old_CurrentCS_Text = CurrentCS_Text
    End If
    
    Dim refillReport As Boolean
    refillReport = False
    
    If Old_PrevCS_Text <> PrevCS_Text Then
        refillReport = True
        
        Dim delRev As Worksheet
        Dim prevRevFile As String
        Dim msgResult As String
        
        Dim prev_rev As Integer
        prev_rev = CInt(PrevCS_Text)
        
        ' Find previous revision
        ' bring it in as RFP Current
        prevRevFile = Dir(ThisWorkbook.Path & VB_VAR_STORE.GetClientSummaryDirectory() & VB_RFP_REPORT.ClientSummaryTitle & "*_rev" & prev_rev & ".xlsx")
                    
RESTART_UNDO:
        If Len(prevRevFile) > 0 Then ' bring in prev_rev - 1 as RFP Report Current
            'Delete "RFP Report Current", if it exists, which it should.
            Set delRev = ThisWorkbook.Sheets("RFP Report Current")
            delRev.Visible = xlSheetVisible
            Application.DisplayAlerts = False
            On Error Resume Next
            delRev.Delete
            Err.Clear
            On Error GoTo 0
            
            Set delRev = Nothing
            
            ' restore RFP Report Current
            Dim newCurrBk As Workbook
            Dim newCurrSht As Worksheet
            Set newCurrBk = Application.Workbooks.Open(ThisWorkbook.Path & VB_VAR_STORE.GetClientSummaryDirectory() & prevRevFile)
            newCurrBk.Sheets(1).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
            Set newCurrSht = ThisWorkbook.ActiveSheet
            newCurrSht.Name = "RFP Report Current"
            newCurrSht.Visible = xlSheetHidden
            Application.DisplayAlerts = True
            
            ' refresh DB
            VB_SITEDB.SetClientSummaryAliasToPrevious
            
            Set newCurrSht = Nothing
            newCurrBk.Close False
            Set newCurrBk = Nothing
            
            Old_PrevCS_Text = PrevCS_Text
        Else
            ' prevRevFile was not found..
            msgResult = MsgBox("Unable to locate the Client Summary that sets the comparison highlights; it may have been moved or renamed. Browse for the file?", vbYesNo)
            If msgResult = vbYes Then
                prevRevFile = FindMissingFile("Find Client Summary _rev" & prev_rev)
                If prevRevFile <> vbNullString Then
                    GoTo RESTART_UNDO
                Else
                    ' findMissingFile was cancelled or failed
                    MsgBox "OK. You can leave the highlights as is, or do them manually. Just go to the ""Format Cells"" tab and color the appropriate cells like any other Excel application.", vbInformation
                    
                    PrevCS_Text = Old_PrevCS_Text
                End If
            Else
                PrevCS_Text = Old_PrevCS_Text
            End If
        End If
        
    End If
    
    If Old_HideZeros <> hideZeros Then
        refillReport = True
        Old_HideZeros = hideZeros
    End If
    
    If refillReport Then
        SetStatusBar "Emptying Report..."
        VB_RFP_REPORT.EmptyReport
        VB_RFP_REPORT.Prepare4Publish
        VB_RFP_REPORT.Visible = xlSheetVisible
        VB_RFP_REPORT.Activate
        SetStatusBar "Waiting for User Approval..."
    End If
    
    If ui_change Then RenderUI True
    
    If ThisRibbon Is Nothing Then
        RefreshGlobalVariables
    End If
    ThisRibbon.Invalidate
End Sub

'Callback for approve_rbutton onAction
Sub Ribbon_ApproveCS(control As IRibbonControl)
    If ThisRibbon Is Nothing Then
        RefreshGlobalVariables
    End If
    ThisRibbon.Invalidate
    
    Dim unapproved As Boolean
    ApplyChangesCS_GetVisible Nothing, unapproved
    If unapproved Then
        Dim vbResult As Integer
        vbResult = MsgBox("Changes have made to the Report Parameters have not been approved and saved. Continuing with the issue process with not account for these changes. Are you sure you wish to continue?", vbYesNo)
        If vbResult = vbYes Then
            'GoTo ContinueIssue
        Else
            Exit Sub
        End If
    End If
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    SetStatusBar "Publishing and Exporting Client Summary..."
    If VB_RFP_REPORT.PublishAndExport() = -1 Then
        ' do something
    End If
    SetStatusBar "Finishing..."
    
    With VB_VAR_STORE
        .SetClientSummaryRev (.GetClientSummaryRev() + 1)
    End With

    'Empty report to restart for next time.
    VB_RFP_REPORT.EmptyReport
    VB_RFP_REPORT.Visible = xlSheetVeryHidden
    
    ResetStatusBar
    VB_MASTER.Activate
    
    If ui_change Then RenderUI True
    
    If ThisRibbon Is Nothing Then
        RefreshGlobalVariables
    End If
    ThisRibbon.Invalidate
End Sub

'Callback for cancel_rbutton onAction
Sub Ribbon_CancelCS(control As IRibbonControl)
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    SetStatusBar "Emptying Report..."
    VB_RFP_REPORT.EmptyReport
    VB_RFP_REPORT.Visible = xlSheetVeryHidden
    
    ResetStatusBar
    VB_MASTER.Activate
    
    If ui_change Then RenderUI True
    
    Old_CurrentCS_Text = ""
    Old_PrevCS_Text = ""
    Old_HideZeros = hideZeros
    CurrentCS_Text = Old_CurrentCS_Text
    PrevCS_Text = Old_PrevCS_Text
    hideZeros = Old_HideZeros
    
    If ThisRibbon Is Nothing Then
        RefreshGlobalVariables
    End If
    ThisRibbon.Invalidate
End Sub

'Callback for current_rev_box getText
Sub CurrentCS_GetText(control As IRibbonControl, ByRef returnedVal)
    returnedVal = CurrentCS_Text
End Sub

'Callback for current_rev_box onChange
Sub CurrentCS_OnChange(control As IRibbonControl, text As String)
    Dim rev_num As Integer
    
    If IsNumeric(text) Then
        rev_num = Abs(Fix(text))
    Else
        MsgBox "A revision number need to non negative integer.", vbExclamation
        If ThisRibbon Is Nothing Then
            RefreshGlobalVariables
        End If
        ThisRibbon.InvalidateControl control.ID
        Exit Sub
    End If

    Dim prev_rev_num As Integer
    prev_rev_num = CInt(PrevCS_Text)
    
    CurrentCS_Text = Str(rev_num)
    If rev_num = 0 Then
        PrevCS_Text = vbNullString
    ElseIf rev_num <= prev_rev_num Then
        PrevCS_Text = Str(rev_num - 1)
    End If
    
    If ThisRibbon Is Nothing Then
        RefreshGlobalVariables
    End If
    ThisRibbon.Invalidate
End Sub

'Callback for prev_rev_box getText
Sub PrevCS_GetText(control As IRibbonControl, ByRef returnedVal)
    returnedVal = PrevCS_Text
End Sub

'Callback for prev_rev_box onChange
Sub PrevCS_GetEnabled(control As IRibbonControl, ByRef enabled)
    Dim str_rev_num As String
    CurrentCS_GetText Nothing, str_rev_num
    
    Dim rev_num As Integer
    rev_num = Str(str_rev_num)
    
    enabled = (rev_num > 0)
End Sub

'Callback for prev_rev_box onChange
Sub PrevCS_OnChange(control As IRibbonControl, text As String)
    Dim prev_rev_num As Integer
    
    If IsNumeric(text) Then
        prev_rev_num = Fix(text)
        If prev_rev_num < -1 Then
            GoTo ErrorMsg
        End If
    Else
ErrorMsg:
        MsgBox "A revision number need to non negative integer, or -1 for no comparison.", vbExclamation
        If ThisRibbon Is Nothing Then
            RefreshGlobalVariables
        End If
        ThisRibbon.InvalidateControl control.ID
        Exit Sub
    End If

    Dim rev_num As Integer
    rev_num = CInt(CurrentCS_Text)
    
    If prev_rev_num < rev_num Then
        PrevCS_Text = Str(prev_rev_num)
    Else
        MsgBox "We are comparing this report to a past revision, so the previous revision number should be less than the current.", vbExclamation
    End If
    
    If ThisRibbon Is Nothing Then
        RefreshGlobalVariables
    End If
    ThisRibbon.Invalidate
End Sub

'Callback for hidezeros_box getPressed
Sub HideZeros_GetPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = hideZeros
End Sub

'Callback for hidezeros_box onAction
Sub Ribbon_HideZeros(control As IRibbonControl, pressed As Boolean)
    hideZeros = pressed
    
    If ThisRibbon Is Nothing Then
        RefreshGlobalVariables
    End If
    ThisRibbon.Invalidate
End Sub

