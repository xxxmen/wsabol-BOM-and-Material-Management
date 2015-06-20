Attribute VB_Name = "EnableMacroCheck"
Option Explicit

Private Const C_NUMLOCK_NAME = "NumLockState"
Private Const C_INTRO_SHEETNAME = "ENABLE MACROS"
Private Const C_WORKBOOK_PASSWORD = ""

Public Sub SaveStateAndHide()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SaveStateAndHide
' This is called from Workbook_BeforeClose.
' This procedure saves the Visible properties of all worksheets
' in the workbook. This will run only if macros are enabled. It
' saves the Visible properties as a colon-delimited string, each
' element of which is the Visible property of a sheet. In the
' property string, C_INTRO_SHEETNAME is set to xlSheetVeryHidden
' so that if the workbook is opened with macros enabled, that
' sheet will not be visible. If macros are not enabled, only
' that sheet will be visible.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim S As String
    Dim ws As Object
    Dim n As Long
    ''''''''''''''''''''''''''''''''''''''''''''
    ' Protection settings. We must be
    ' able to unprotect the workbook in
    ' order to modify the sheet visibility
    ' properties. We will restore the
    ' protection at the end of this procedure.
    ''''''''''''''''''''''''''''''''''''''''''''
    Dim HasProtectWindows As Boolean
    Dim HasProtectStructure As Boolean
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    ' Save the workbook's protection settings and
    ' attempt to unprotect the workbook.
    '''''''''''''''''''''''''''''''''''''''''''''''
    HasProtectWindows = ThisWorkbook.ProtectWindows
    HasProtectStructure = ThisWorkbook.ProtectStructure
    
    ThisWorkbook.Unprotect Password:=C_WORKBOOK_PASSWORD
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    ' Make the introduction sheet visible
    '''''''''''''''''''''''''''''''''''''''''''''''
    If ThisWorkbook.Worksheets(C_INTRO_SHEETNAME).Visible = xlSheetVisible Then
        GoTo PROTECT_SAVE
    Else
        ThisWorkbook.Worksheets(C_INTRO_SHEETNAME).Visible = xlSheetVisible
    End If
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    For Each ws In ThisWorkbook.Sheets
        '''''''''''''''''''''''''''''''''''''''''''''''
        ' If WS is the intro sheet, make it visible,
        ' otherwise make it VeryHidden. This sets all
        ' sheets except C_INTRO_SHEETNAME to very
        ' hidden.
        ''''''''''''''''''''''''''''''''''''''''''''''''
        If StrComp(ws.Name, C_INTRO_SHEETNAME, vbTextCompare) = 0 Then
            ws.Visible = xlSheetVisible
            'MsgBox ws.Name & " Visible"
        Else
            ws.Visible = xlSheetVeryHidden
            'MsgBox ws.Name & " VeryHidden"
        End If
    Next ws
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Save state of "Toggle Number Lock" button (gbl_num_lock)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error Resume Next
    ThisWorkbook.Names(C_NUMLOCK_NAME).Delete
    Err.Clear
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=C_NUMLOCK_NAME, RefersTo:=CStr(ThisWorkbook.MarkNumLock()), Visible:=False
    
    
PROTECT_SAVE:
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set the workbook protection back to what it was.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ThisWorkbook.Protect C_WORKBOOK_PASSWORD, _
        structure:=HasProtectStructure, Windows:=HasProtectWindows

    
    If ui_change Then RenderUI True
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Save Workbook with sheets hidden
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ThisWorkbook.Save
End Sub

Public Sub UnHideSheets()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' UnHideSheets
' This is called by Workbook_Open to hide the introduction sheet
' and set all the other worksheets to their visible state that
' was stored when the workbook was last closed. The introduction
' sheet is set to xlSheetVeryHidden. This maro is executed only
' if macros are enabled. If the workbook is opened without
' macros enabled, only the introduction sheet will be visible.
' If an error occurs, make the intro sheet visible and get out.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim S As String
    Dim n As Long
    Dim VisibleArr As Variant
    Dim HasProtectWindows As Boolean
    Dim HasProtectStructure As Boolean
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    ' Save the workbook's protection settings and
    ' attempt to unprotect the workbook.
    '''''''''''''''''''''''''''''''''''''''''''''''
    HasProtectWindows = ThisWorkbook.ProtectWindows
    HasProtectStructure = ThisWorkbook.ProtectStructure
    
    ThisWorkbook.Unprotect Password:=C_WORKBOOK_PASSWORD
    
    On Error GoTo ErrHandler:
    Err.Clear
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    Dim ws As Worksheet
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Loop through the array of Worksheets and set the Visible propety
    ' for each sheet. If we're processing the C_INTRO_SHEETNAME
    ' sheet, make it Visible (since it may be the only
    ' visbile sheet). We'll hide it later after the loop.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    INTRO_MACROCHECK.Visible = xlSheetVisible
    For Each ws In ThisWorkbook.Worksheets
        If Not ws Is INTRO_MACROCHECK Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    Set ws = Nothing
    
    VB_COVERSHEET.Visible = xlSheetVisible
    VB_MASTER.Visible = xlSheetVisible
    VB_ORDER_TMPLT.Visible = xlSheetVisible
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "* - BOM" Then
            ws.Visible = xlSheetVisible
        ElseIf ws.Name = "RFP Report Current" Then
            ws.Visible = xlSheetHidden
        End If
    Next ws
    
    ''''''''''''''''''''''''''''''''
    ' Hide the INTRO sheet.
    ''''''''''''''''''''''''''''''''
    ThisWorkbook.Sheets(C_INTRO_SHEETNAME).Visible = xlSheetVeryHidden
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set state of "Toggle Number Lock" button (gbl_num_lock)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ThisWorkbook.SetMarkNumLock (StrComp(UCase(ThisWorkbook.Names(C_NUMLOCK_NAME).Value), "=TRUE", vbTextCompare) = 0)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Set the workbook protection back to what it was.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ThisWorkbook.Protect Password:=C_WORKBOOK_PASSWORD, _
        structure:=HasProtectStructure, Windows:=HasProtectWindows
    
    If ui_change Then RenderUI True
    
    Exit Sub
    
ErrHandler:
    ThisWorkbook.Worksheets(C_INTRO_SHEETNAME).Visible = xlSheetVisible
    
End Sub

Private Sub AfterClose()
'Subroutine: AfterClose - Pseudo event handler that runs after ThisWorkbook has closed. In the case where the
'                         user chooses not to save changes before close, ThisWorkbook must revert back to its previously saved
'                         state by closing and reopening so that the application can continue with the SaveStateAndHide routine.
'                         This procedure runs after ThisWorkbook has been closed - and in its reverted state.


    ' Get Focus
    ThisWorkbook.Activate
    
    ' So Workbook_BeforeClose doesn't run completely and recalls this method, set WORKBOOK_CLOSED = True
    ThisWorkbook.SetWorkbookClosed True
    
    'Clear ThisRibbon Pointer
    VB_VAR_STORE.ClearRibbonID
    
    ' SAVE VISIBILITY STATE AND HIDE ALL SHEETS BUT THE INTRO_SHEET
    SaveStateAndHide
    
    'If this is the only worbook open, quit application.
    If Application.Workbooks.count > 1 Then
        ThisWorkbook.Close SaveChanges:=True
    Else
        Application.Quit
    End If
End Sub
