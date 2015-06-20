VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'On Worksheet_Deactivate, save sheet name to handle in Worksheet_Activate. If leaving certain sheet, action may be required
Private last_sheet_name As String

'gbl_num_lock: local variable to ThisWorkbook. Hold the state of the mark number lock property
Private gbl_num_lock As Boolean

'WORKBOOK_CLOSED: During "AfterClose" operation, this workbook is reopened. This property is used to differentiate
'between ThisWorkbook actually closing, and it closing the second time after being reopened in AfterClose.
Private WORKBOOK_CLOSED As Boolean

Private Sub Workbook_BeforeClose(Cancel As Boolean)
'Subroutine: Workbook_BeforeClose - Excel event handler. This subroutine is called before ThisWorkbook closes.
'                                   If workbook is not saved, prompt user. Save the workbook's sheet visibility settings
'                                   and hide all sheets except the introduction sheet. This Workbook_BeforeClose
'                                   will run only if macros are enabled. If Workbook is opened as Read-Only (during an AutoCAD import),
'                                   don't save changes, just close.
'Arguments: Cancel - Boolean returning a cancel command. True - Cancel close; False - Continue close


    Cancel = False
    
    If ThisWorkbook.ReadOnly Or WORKBOOK_CLOSED Then
        'Workbook is opened as Read-Only.
        ThisWorkbook.Saved = True
    Else
        ' if in Inventory Manager, force user to discard/close or cancel
        If VB_CLIENTINVENTORY.Visible = xlSheetVisible Then
            VB_MASTER.Activate
            
            ' if Inventory Manager is still open, user chose Cancel
            If VB_CLIENTINVENTORY.Visible = xlSheetVisible Then
                GoTo CANCEL_CLOSE
            Else
                ThisWorkbook.Save
            End If
        End If
        
        ' if in Client Summary preview, force cancel
        If VB_RFP_REPORT.Visible = xlSheetVisible Then
            Ribbon_CancelCS Nothing
        End If
        
        ' if in CADWORX IMPORT, cancel.
        If SheetExists("CADWORX_IMPORT") Then
            'Back out of CADWorx Import completely
            Application.DisplayAlerts = False
            ThisWorkbook.Sheets("CADWORX_IMPORT").Delete
            Application.DisplayAlerts = True
            VB_MASTER.Activate
        End If
        
        'USER SAVE DIALOGUE
        If Not ThisWorkbook.Saved Then
            ' Workbook is not saved. Prompt user with save dialouge
            
            Dim result As Integer
            result = MsgBox("Do you want to save the " & _
                "changes you made to " & ThisWorkbook.Name & "?", vbYesNoCancel)
            
            Select Case result
                Case vbYes:
                    'Save workbook outright
                    ThisWorkbook.Save
                    GoTo WORKBOOK_SAVED
                    
                Case vbNo:
                    'Close workbook, then reopen to revert changes beck to last saved position.
                    ThisWorkbook.Saved = True
                    
                    'If WORKBOOK_CLOSED = True, this has already been reopened and is now closing again.
                    If Not WORKBOOK_CLOSED Then
                        'Revert to last saved state
                        Cancel = True
                        
                        ' Application.OnTime Now, ... adds "AfterClose" command to the end of the stack: after the application
                        ' leaves this subroutine. Therefore, "AfterClose" will execute after the workbook closes.
                        Application.OnTime Now, "AfterClose"
                        
                        Application.EnableEvents = False
                        ThisWorkbook.Close SaveChanges:=False
                        Application.EnableEvents = True
                    End If
                    
                Case vbCancel:
                    'Cancel
CANCEL_CLOSE:
                    Cancel = True
                    Exit Sub
                    
            End Select
            
        Else
            ' Workbook is Saved. Continue to close
WORKBOOK_SAVED:
            If Not WORKBOOK_CLOSED Then
                'Clear ThisRibbon Pointer
                VB_VAR_STORE.ClearRibbonID
                
                ' SAVE VISIBILITY STATE AND HIDE ALL SHEETS BUT THE INTRO_SHEET
                SaveStateAndHide
            End If
        End If
    End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'Subroutine: Workbook_BeforeSave - Excel event handler. This subroutine is called before ThisWorkbook is saved.
'                                  Sets the 'Last Saved By' cells on the Coversheet
'Arguments: SaveAsUI - Boolean. True - Save As dialogue will appear; False - Save As dialogue won't appear
'           Cancel - Boolean returning a cancel command. True - Cancel close; False - Continue close.
    
    
    If Not WORKBOOK_CLOSED Then
        
        Dim ui_change As Boolean
        ui_change = RenderUI(False)

        If Not IsBlankBOM() Then
            VB_COVERSHEET.SetLastSavedBy
        Else
            VB_COVERSHEET.ClearLastSavedBy
        End If
    
        If ui_change Then RenderUI True
    End If
    
End Sub

Public Sub OpenWorkbook()
'Subroutine: OpenWorkbook - Used to call the ThisWorkbook's _Open Event from outside this module.

    Workbook_Open
End Sub

Private Sub Workbook_Open()
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Unhide the sheets. This Workbook_Open event
    ' will run only if macros are enabled. If macros
    ' are not enabled, this code will not run and
    ' only the introduction sheet will be visible.
    ''''''''''''''''''''''''''''''''''''''''''''''''
    
    If Not IsUniqueInstance() Then
        MsgBox "Only ONE(1) HMM BOM Application may be open at once. Close your other version(s) before attempting to open any others. " & _
            "If the still doesn't work. Completely exit all Excel workbooks and try again.", vbCritical
            
        WORKBOOK_CLOSED = True
        ThisWorkbook.Close False
    End If
    Application.ScreenUpdating = False
    
    WORKBOOK_CLOSED = False
    
    UnHideSheets
    
    'Establish Protection/Security
    VB_COVERSHEET.SetProtection
    VB_MASTER.SetProtection
    VB_CHANGE_LOG.SetProtection
    VB_MASTER.ResetScrollArea
    VB_ORDER_TMPLT.SetProtection
    
    VB_CHANGE_LOG.SetTrackChanges True
    
    If IsBlankBOM() Then
        BlankBOM
    Else
        If VB_COVERSHEET.Visible = xlSheetVisible Then
            VB_COVERSHEET.Activate
        End If
        
        If VB_CLIENTINVENTORY.Visible = xlSheetVisible Then
            VB_CLIENTINVENTORY.Activate
        End If
    End If
    
    If Not Application.ScreenUpdating Then
        Application.ScreenUpdating = True
    End If
    If Not Application.EnableEvents Then
        Application.EnableEvents = True
    End If
End Sub

Public Function MarkNumLock() As Boolean
'Function: MarkNumLock - 'Get' method for the BOM's number_lock state.
'Returns: Boolean - True: mark numbers are locked; False: mark numbers are not locked

    MarkNumLock = gbl_num_lock
End Function

Public Sub SetMarkNumLock(ByVal vbVal As Boolean)
'Subroutine: SetMarkNumLock - 'Set' method for the BOM's number_lock state.
'Arguments: vbVal - Boolean contains the future state of gbl_num_lock

    gbl_num_lock = vbVal
End Sub

Public Function IsWorkbookClosed() As Boolean
'Function: IsWorkbookClosed - 'Get' method for the BOM's WORKBOOK_CLOSED.
'Returns: The state of WORKBOOK_CLOSED

    IsWorkbookClosed = WORKBOOK_CLOSED
End Function

Public Sub SetWorkbookClosed(ByVal vbVal As Boolean)
'Subroutine: SetMarkNumLock - 'Set' method for the WORKBOOK_CLOSED state.
'Arguments: vbVal - Boolean containing the new state of WORKBOOK_CLOSED

    WORKBOOK_CLOSED = vbVal
End Sub

Public Function IsBlankBOM() As Boolean
'Function: IsBlankBOM - Determines if this BOM is blank or not. A Blank BOM has no sites, no orders, no items, and no info in the coversheet.
'Returns: Boolean - True: BOM is blank; False: BOM is not blank.


    IsBlankBOM = (NumSites() = 0)
    IsBlankBOM = IsBlankBOM And VB_ORDER_LOG.NumOrders() = 0
    'IsBlankBOM = IsBlankBOM And MaxMark() = 1
    IsBlankBOM = IsBlankBOM And IsEmpty(VB_MASTER.Cells(VB_MASTER.FirstRow(), get_col_num("Long Description")))
    IsBlankBOM = IsBlankBOM And (IsEmpty(VB_COVERSHEET.Range("$F$14").MergeArea.Cells(1, 1)) _
            Or IsEmpty(VB_COVERSHEET.Range("$F$16").MergeArea.Cells(1, 1)))
    
    Dim wrksht As Worksheet
    On Error Resume Next
    Set wrksht = ThisWorkbook.Sheets("RFP Report Current")
    Err.Clear
    On Error GoTo 0
    
    IsBlankBOM = IsBlankBOM And (wrksht Is Nothing)
    
    Set wrksht = Nothing
    On Error Resume Next
    Set wrksht = ThisWorkbook.Sheets("CADWORX_IMPORT_OLD")
    Err.Clear
    On Error GoTo 0
    
    IsBlankBOM = IsBlankBOM And (wrksht Is Nothing)
    
    Set wrksht = Nothing
End Function

Public Sub BlankBOM()
'Subroutine: BlankBOM - Empties BOM for future projects.
    
    SetStatusBar "Prepping..."
    
    'Turn off track changes
    VB_CHANGE_LOG.SetTrackChanges False
    SetMarkNumLock False
    RenderUI False
    
    'Unsuppress mark numbers
    VB_MASTER.Columns(get_col_num("Mark No.")).Hidden = False
    VB_ORDER_TMPLT.Columns(VB_ORDER_TMPLT.Form_col_num("Item #")).Hidden = False
    
    'Make sure all items are visible
    Dim first As Integer
    Dim last As Integer
    first = VB_MASTER.FirstRow()
    last = VB_MASTER.LastRow()
    Dim row As Integer
    For row = first To last
        VB_MASTER.Rows(row).Hidden = False
    Next row
    
    SetStatusBar "Clearing RFP Form..."
    
    'Empty RFP Form
    VB_ORDER_TMPLT.ClearForm
    
    'Remove orders
    Dim num_ords As Integer
    num_ords = VB_ORDER_LOG.NumOrders()
    Do While VB_ORDER_LOG.NumOrders() > 0
        SetStatusBar "Removing Orders...", (num_ords - VB_ORDER_LOG.NumOrders()), num_ords
        VB_ORDER_LOG.DeleteOrder VB_ORDER_LOG.OrderNumber(0)
    Loop
    
    'Remove sites
    Dim num_sites As Integer
    num_sites = NumSites()
    Do While Not IsEmpty(VB_MASTER.Cells(VB_MASTER.SubtitleRow(), get_col_num("Current Model Quantities")))
        SetStatusBar "Removing Sites...", (num_sites - NumSites()), num_sites
        RemoveSite CStr(VB_MASTER.Cells(VB_MASTER.SubtitleRow(), get_col_num("Current Model Quantities")).Value2)
    Loop
    If VB_MASTER.Columns(get_col_num("Total Extras")).Hidden Then
        VB_MASTER.Columns(get_col_num("Total Extras")).Hidden = False
        GetRibbon().Invalidate
    End If
    
    'Remove Items
    last = VB_MASTER.LastRow()
    For row = VB_MASTER.LastRow() To first Step -1
        SetStatusBar "Removing Material Items...", (last - row), (last - first)
        DeleteMasterRow row
    Next row
    If Not VB_MASTER.Columns(get_col_num("SAP#")).Hidden Then
        VB_MASTER.Columns(get_col_num("SAP#")).Hidden = True
        GetRibbon().Invalidate
    End If
    
    SetStatusBar "Deleting Categories...", 1, 1
    'remove categories
    row = VB_CATEGORY.FirstRow() + 1
    Do While Not IsEmpty(VB_CATEGORY.Cells(row, VB_CATEGORY.CategoryColumn()))
        DeleteCategory VB_CATEGORY.Cells(row, VB_CATEGORY.CategoryColumn()).Value2
    Loop
    
    SetStatusBar "Loading Preset Categories..."
    'force category to LINE PIPE
    row = VB_CATEGORY.FirstRow()
    If VB_CATEGORY.Cells(row, VB_CATEGORY.CategoryColumn()).Value2 <> "LINE PIPE" Then
        VB_CATEGORY.RenameCategory VB_CATEGORY.Cells(row, VB_CATEGORY.CategoryColumn()).Value2, "LINE PIPE"
    End If
    VB_CATEGORY.ClearPhrases "LINE PIPE"
    VB_CATEGORY.StoreFirstPhraseKey "LINE PIPE", "PIPE"
    
    'set Mark# and uniqueID
    VB_MASTER.Cells(get_cat_row("LINE PIPE"), get_col_num("Mark No.")).Value = 1
    Call set_property(1, get_col_num("Unique ID"), "A000001")
    
    'clear all Unique IDs, add A000001
    VB_UNIQUEID.ClearAll
    VB_UNIQUEID.AddUniqueID "A000001"
    
    'add category other preincluded categories
    NewCategory "FLANGES", "EA", VB_MASTER.LastRow() + 1
    VB_CATEGORY.StoreFirstPhraseKey "FLANGES", "FLANGE"
    
    NewCategory "TEES", "EA", VB_MASTER.LastRow() + 1
    VB_CATEGORY.StoreFirstPhraseKey "TEES", "TEE"
    
    NewCategory "REDUCERS", "EA", VB_MASTER.LastRow() + 1
    VB_CATEGORY.StoreFirstPhraseKey "REDUCERS", "REDUCER"
    
    NewCategory "ELLS", "EA", VB_MASTER.LastRow() + 1
    VB_CATEGORY.StoreFirstPhraseKey "ELLS", "ELL"
    
    NewCategory "VALVES", "EA", VB_MASTER.LastRow() + 1
    VB_CATEGORY.StoreFirstPhraseKey "VALVES", "VALVE"
    
    SetStatusBar "Resetting Auxilary Sheets...", 0, 1
    'Delete RFP Report Current
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("RFP Report Current").Visible = xlSheetVisible
    ThisWorkbook.Sheets("RFP Report Current").Delete
    Err.Clear
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    SetStatusBar "Resetting Auxilary Sheets...", 0.2, 1
    'Delete CADWORX_IMPORT_OLD
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("CADWORX_IMPORT_OLD").Visible = xlSheetVisible
    ThisWorkbook.Sheets("CADWORX_IMPORT_OLD").Delete
    Err.Clear
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    SetStatusBar "Resetting Auxilary Sheets...", 0.4, 1
    'clear change log
    VB_CHANGE_LOG.Rows("1:" & VB_CHANGE_LOG.UsedRange.Rows.CountLarge).Hidden = False
    If VB_CHANGE_LOG.UsedRange.Rows.CountLarge > 2 Then
        VB_CHANGE_LOG.UsedRange.Rows("2:" & VB_CHANGE_LOG.UsedRange.Rows.CountLarge).ClearContents
    End If
    
    SetStatusBar "Resetting Auxilary Sheets...", 0.6, 1
    ' reset VB_VAR_STORE info
    With VB_VAR_STORE
        .HMMFileTagFormat "_""~ProjectName""_yyyy-mm-dd"
        .SetClientSummaryRev 0
        
        ' Current Directory is \Master_Material_List\
        .SetRFPDirectory "\..\RFP\"
        .SetClientSummaryDirectory "\..\Client_List\"
        .SetSiteBOMDirectory "\Site BOM Archive\"
        .SetMDRDirectory "\Model Discrepancy Reports\"
        .SetDeletedRFPDirectory "\..\RFP\Unpublished RFP\"
    End With
    
    SetStatusBar "Resetting Coversheet...", 0.8, 1
    
    VB_COVERSHEET.Unprotect
    'Clear coversheet information
    VB_COVERSHEET.SetProjectName vbNullString
    VB_COVERSHEET.SetProjectClientName vbNullString
    VB_COVERSHEET.SetProjectNumber vbNullString
    VB_COVERSHEET.ClearHMMContact
    VB_COVERSHEET.ClearLastSavedBy
    
    VB_COVERSHEET.SetProtection
    VB_CHANGE_LOG.SetTrackChanges True
    
    SetStatusBar "Finishing...", 1, 1
    VB_ORDER_TMPLT.Visible = xlSheetVisible
    VB_ORDER_TMPLT.Activate
    VB_ORDER_TMPLT.Cells(18, 2).Select
    
    VB_CHANGE_LOG.Visible = xlSheetVisible
    VB_CHANGE_LOG.Activate
    VB_CHANGE_LOG.Cells(2, 1).Select
    VB_CHANGE_LOG.Visible = xlSheetVeryHidden
    
    VB_MASTER.Visible = xlSheetVisible
    VB_MASTER.Activate
    VB_MASTER.Columns(get_col_num("Mark No.")).Hidden = True
    VB_MASTER.Columns(get_col_num("Model Extras")).Hidden = False
    VB_MASTER.Columns(get_col_num("Total Extras")).Hidden = False
    VB_MASTER.Cells(VB_MASTER.FirstRow(), get_col_num("Long Description")).Select
    
    VB_COVERSHEET.Visible = xlSheetVisible
    VB_COVERSHEET.Activate
    VB_COVERSHEET.Cells(14, 6).MergeArea.Select
    
    'Reset Application
    RefreshGlobalVariables
    ResetStatusBar
    RenderUI True
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
'Subroutine: Workbook_SheetDeactivate - Excel event handler. This subroutine is called right after a sheet is deactivated.
'                                       Sets the last_sheet_name so it can be used in Workbook_SheetActivate.
'Arguments: Sh - Generalized Worksheet object for the sheet that is deactivating
    
    If Sh.Name = "Change Log" Then
        Application.DisplayAlerts = False
        Sh.Delete
        Application.DisplayAlerts = True
        
        GetRibbon().Invalidate
    Else
        last_sheet_name = Sh.Name
    End If
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
'Subroutine: Workbook_SheetActivate - Excel event handler. This subroutine is called right before a sheet is activated.
'                                     If user is leaving VB_CLIENTINVENTORY, send them back, and make them Save or Discard their progress on that sheet.
'Arguments: Sh - Generalized Worksheet object for the sheet that is activating


    If last_sheet_name = VB_CLIENTINVENTORY.Name And Sh.Name <> VB_CLIENTINVENTORY.Name Then
        VB_CLIENTINVENTORY.Activate
    
        Dim vbResult As Integer
        vbResult = MsgBox("You should not leave without saving or discarding your changes. Leaving this page will discard your changes. " & _
            "Are you sure you want to leave this page and discard your changes?", vbYesNo)
        
        If vbResult = vbYes Then
            VB_CLIENTINVENTORY.CloseForm
            GetRibbon().Invalidate
            
            Sh.Activate
        Else
            VB_CLIENTINVENTORY.Activate
        End If
    End If
    
    If Sh.CodeName = "VB_MASTER" And InStr(Sh.Name, " ") <> 0 Then
        Dim corrected_name As String
        Dim i As Integer
        corrected_name = ""
        
        For i = 1 To Len(Sh.Name)
            If Mid(Sh.Name, i, 1) <> " " Then
                corrected_name = corrected_name & Mid(Sh.Name, i, 1)
            End If
        Next i
        
        Sh.Name = corrected_name
    End If
End Sub
