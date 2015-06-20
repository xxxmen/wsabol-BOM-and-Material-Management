VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VB_CHANGE_LOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' hold_tracking: Property dictating whether this BOM is currently holding track changes - Not track_changes.
' When errors occur, global variables reset to false. Therefore, changes will still be tracked after errors
Private hold_tracking As Boolean

' Row/Column location properties
Private Const first_row = 2
Private Const vb_date_col = 1
Private Const vb_user_col = 2
Private Const vb_cell_col = 3
Private Const vb_unID_col = 4
Private Const vb_mark_col = 5
Private Const vb_desc_col = 6
Private Const vb_from_col = 7
Private Const vb_to_col = 8

' inProgress: keeps the User from initiating multiple HMMUndo and HMMRedo actions at once.
Private inProgress As Boolean

Public Function LogChange(ByVal locationStr As String, ByVal mark_num As String, ByVal description As String, ByVal fromStr, ByVal toStr) As Integer
'Function: LogChange -  Appends a row to the "Change Log" sheet cataloging activity in this BOM.
'                       Will work for single and multiple cell changes up to ~3000 cells at a time. This sheet
'                       acts as the Undo/Redo stack for this application.
'Arguments: locationStr - String containing sheet/cell location of where the change occured.
'                         i.e. "MASTER!$A$1"
'           mark_num - String containing the "Mark No." of the item if applicable
'           description - String containing a description of the change. For a generic change to MASTER,
'                         this shall be the column title where the change took place.
'           fromStr - String containing the original state of the cell/sheet/workbook before the change
'           toStr -   String containing the final state of the cell/sheet/workbook after the change
'Return: Integer containing an error code. 0: no error. -1: Error.


    LogChange = 0
    Dim row As Integer
    
    With VB_CHANGE_LOG
    
        'Find next blank row
        row = first_row
        Do While Not EOF(row)
            row = row + 1
        Loop
        
        'if row is bad
        If row <= 0 Then
            LogChange = -1
            ErrorHandling "LogChange", LogChange, "Change Log Sheet either doesn't exist or complete empty." & _
                vbCrLf & "Cannot log change: " & description, 1
            Exit Function
        End If
        
        ' remove hidden rows that have resulted from HMMUndo
        If row <> first_row Then
            Dim lastID As Integer
            lastID = LastChangeID()
            
            If row <> lastID + 1 Then
                'delete Undo Stack
                VB_CHANGE_LOG.Rows(lastID + 1 & ":" & row - 1).Delete
                GetRibbon().Invalidate
            End If
            
            row = LastChangeID() + 1
        End If
        
        Dim ui_change As Boolean
        ui_change = RenderUI(False)
        
        ''''''''''''''''''''''''''''''
        'Write the change to the log
        
        'if an error, clear row values so we can start fresh next time and undo won't crash
        On Error GoTo UNDO_ROW
        .Cells(row, vb_date_col).Value = Format(Now(), "yyyy/mmm/dd H:NN AM/PM")
        .Cells(row, vb_user_col).Value = Environ("USERNAME")
        .Cells(row, vb_cell_col).Value = locationStr
        If Len(mark_num) > 0 Then
            .Cells(row, vb_unID_col).Value = get_property(CInt(mark_num), "Unique ID")
        End If
        .Cells(row, vb_mark_col).Value = mark_num
        .Cells(row, vb_desc_col).Value = description
        .Cells(row, vb_from_col).Value = fromStr
        .Cells(row, vb_to_col).Value = toStr
        Err.Clear
        On Error GoTo 0
        
        If ui_change Then RenderUI True
        
    End With

    'if row = first_row, invalidate so Undo is enabled
    If row = first_row Then
        GetRibbon().Invalidate
    End If
    
    Exit Function
    
UNDO_ROW:
    VB_CHANGE_LOG.Rows(row).ClearContents
    LogChange = -1
    
End Function

Public Sub offset(ByVal sheet As String, ByVal in_row As Integer, ByVal in_col As Integer, ByVal offset_r As Integer, ByVal offset_c As Integer)
'Subroutine: Offset - The Change Log has a column for the cell location. If rows/columns are inserted/deleted
'                     on any sheet, the written cell locations in VB_CHANGE_LOG will be invalid. This function
'                     adjusts the cell locations in VB_CHANGE_LOG according to the arguments.
'Arguments: sheet - the sheet name that the change happened on
'           in_row - the initial row to start offsetting - the row you would have right clicked to say Insert/Delete
'           in_col - the initial col to start offsetting - the row you would have right clicked to say Insert/Delete
'           offset_r - the offset value for each row. offset_r < 0 - deleted row(s); offset_r > 0 - inserted rows(s)
'           offset_c - the offset value for each column. offset_c < 0 - deleted columns(s); offset_c > 0 - inserted columns(s)

    
    Dim read_addr As Range
    Dim row As Integer
    row = first_row
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
On Error GoTo Continue
    Do While Not IsEmpty(VB_CHANGE_LOG.Cells(row, vb_date_col))
        
        If IsCellLocation(VB_CHANGE_LOG.Cells(row, vb_cell_col).Value2, read_addr) Then
            'MsgBox VB_CHANGE_LOG.Cells(row, vb_cell_col).Text
            
            'check sheet name in the read address
            If read_addr.Worksheet.Name = sheet Then
                
                ' adjust row
                If offset_r > 0 Then ' row insertion
                    If read_addr.row >= in_row Then
                        Set read_addr = read_addr.offset(offset_r, 0)
                    End If
                ElseIf offset_r < 0 Then ' row deletion
                    If read_addr.row > in_row Then
                        Set read_addr = read_addr.offset(offset_r, 0)
                    End If
                End If
                
                ' adjust column
                If offset_c > 0 Then ' column insertion
                    If read_addr.column >= in_col Then
                        Set read_addr = read_addr.offset(0, offset_c)
                    End If
                ElseIf offset_c < 0 Then ' column deletion
                    If read_addr.column > in_col Then
                        Set read_addr = read_addr.offset(0, offset_c)
                    End If
                End If
                
                ' rewrite cell location
                VB_CHANGE_LOG.Cells(row, vb_cell_col).Value = VB_MASTER.Name & "!" & read_addr.Address
            End If
        End If
        
Continue:
        Set read_addr = Nothing
        row = row + 1
    Loop
    Err.Clear
    On Error GoTo 0
    
    If ui_change Then RenderUI True
    
End Sub

Public Function ColID(ByVal vbID As String) As Integer
'Function: ColID - 'Get' method for VB_CHANGE_LOG column numbers. Returns the column number for the
'                  specified column ID. If the column ID doesn't exist, returns -1.
'Arguments: vbID - String containing the column ID
'Returns: Integer containing the corresponding column number


    Select Case vbID
        Case "Date":
            ColID = vb_date_col
        Case "User":
            ColID = vb_user_col
        Case "Cell":
            ColID = vb_cell_col
        Case "Mark":
            ColID = vb_mark_col
        Case "Unique ID":
            ColID = vb_unID_col
        Case "Description":
            ColID = vb_desc_col
        Case "From":
            ColID = vb_from_col
        Case "To":
            ColID = vb_to_col
        Case Else:
            ColID = -1
    End Select
End Function

Public Sub reset_mark_num(ByVal old_mark As Integer)
'Subroutine: reset_mark_num - Just as we update cell location as rows/columns are added/deleted,
'                             with a mark number change, we update all changes that involve that past mark number
'                             with the new mark number located on MASTER. We correlate mark numbers with Unique ID
'Arguments: old_mark - Integer containing the old mark number. The new mark number will be grabbed from
'                      VB_MASTER using the cell information in VB_CHANGE_LOG


    Dim last_row As Integer
    last_row = LastChangeID()

    Dim row As Integer
    row = first_row
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    Do
        If CInt(VB_CHANGE_LOG.Cells(row, vb_mark_col).Value2) = old_mark Then
            VB_CHANGE_LOG.Cells(row, vb_mark_col).Value = getMark4UniqueID(VB_CHANGE_LOG.Cells(row, vb_unID_col).Value2)
            If CInt(VB_CHANGE_LOG.Cells(row, vb_mark_col).Value2) = -1 Then
                VB_CHANGE_LOG.Cells(row, vb_mark_col).Value = old_mark
            End If
        End If
        row = row + 1
    Loop While row <= last_row
    
    If ui_change Then RenderUI True
End Sub

Public Function HMMUndo(Optional ByVal recursive As Boolean = False, Optional ByVal getPermission As Boolean = True) As Boolean
'Subroutine: HMMUndo - Handles undo commands. The Undo/Redo stack is defined in VB_CHANGE_LOG by hiding and unhiding rows.
'                      Potentially large actions such as Description Check Resets, and CADWorx BOM Imports are preceded by a
'                      prompt seeking the user's permission to continue. Global varibale inProgess prevents the User from
'                      initiating multiple Undo's at once.
'Arguments: recursive - This routine utilizes recursion to facilitate potentially large actions such as Description Check
'                       Resets, and CADWorx BOM Imports. Setting recursive to True prevents this routine from performing certain
'                       aesthetic actions, but the actual undo process runs normally.
'           getPermission
'Returns: Boolean - success.
    
    HMMUndo = False

    'GET ROW ID FOR LAST CHANGE
    Dim row As Integer
    row = LastChangeID()
    
    'sanity check
    If row = -1 Then
        Exit Function
    End If

    'check/set inProgress
    If inProgress Then
        Exit Function
    Else
        inProgress = True
    End If
    
    On Error GoTo ResetApp
    
    'HOLD TRACK CHANGES
    Dim change As Boolean
    change = False
    If TrackChanges Then
        change = True
        SetTrackChanges False
    End If
    
    'SAVE ACTIVESHEET
    Dim shtName As String
    If Not recursive Then
        shtName = ThisWorkbook.ActiveSheet.Name
    End If
    
    'prepare for success
    Dim success As Boolean
    success = True
    
    Dim ui_change As Boolean
    
    Dim trgCell As Range
    Dim vbDate As String
    Dim vbMark As Integer
    Dim vbUser As String
    Dim vbCellAddress As String
    Dim description As String
    Dim vbFrom As String
    Dim vbTo As String
    Dim TempRibbon As IRibbonUI
    Dim change_col As Integer
    Dim site_name As String
    Dim order_num As String
    Dim order_site As String
    Dim msgResult As Integer
    Dim finish_row As Integer
    Dim cat_name As String
    Dim prev_rev As Integer
    Dim OldFileName As String
    Dim NewFileName As String
    Dim delRev As Worksheet
    Dim prevRevFile As String
    Dim cat_unit As String
    Dim old_site_name As String
    
    'INITIALIZE VARIABLES
    vbDate = VB_CHANGE_LOG.Cells(row, vb_date_col).Value
    vbUser = VB_CHANGE_LOG.Cells(row, vb_user_col).Value2
    vbMark = VB_CHANGE_LOG.Cells(row, vb_mark_col).Value2
    vbCellAddress = VB_CHANGE_LOG.Cells(row, vb_cell_col).Value2
    description = VB_CHANGE_LOG.Cells(row, vb_desc_col).Value2
    vbFrom = VB_CHANGE_LOG.Cells(row, vb_from_col).Value2
    vbTo = VB_CHANGE_LOG.Cells(row, vb_to_col).Value2
    
    ui_change = RenderUI(False)
    
    'GET CONFIRMATION FROM USER
    If getPermission Then
        If get_col_num(description) = -1 Then
            msgResult = MsgBox("Undo the following action?" & vbCrLf & vbCrLf & _
                description & "." & IIf(Len(vbFrom) > 0, vbCrLf & "From: " & vbFrom, "") & IIf(Len(vbTo) > 0, vbCrLf & "To: " & vbTo, ""), vbYesNo)
        Else
            msgResult = MsgBox("Revert the following property?" & vbCrLf & vbCrLf & _
                description & IIf(Len(vbFrom) > 0, vbCrLf & "From: " & vbFrom, "") & IIf(Len(vbTo) > 0, vbCrLf & "To: " & vbTo, ""), vbYesNo)
        End If
    Else
        msgResult = vbYes
    End If
    
    If msgResult <> vbYes Then
        success = False
        GoTo ResetApp
    End If

    Select Case True
    
        ' Like "Added Material Item"
        Case description Like "Added Material Item":
            If get_row(vbMark) <> -1 Then
                DeleteMasterRow get_row(vbMark)
            Else
                success = False
            End If
            If get_row(vbMark) <> -1 Then
                'ERROR
                success = False
            End If
            
        ' Like "Purged Material Item"
        Case description Like "Purged Material Item":
            Dim new_mark As Integer
            new_mark = InsertItemBOM(vbTo, vbFrom)
            If new_mark <> vbMark And get_row(vbMark) = -1 Then
                VB_MASTER.Cells(get_row(new_mark), get_col_num("Mark No.")).Value = vbMark
                new_mark = vbMark
            End If
            If get_property(new_mark, "Long Description") <> vbTo Then
                success = False
            End If
        
        ' Like "Delete Item"
        Case description Like "Delete Item":
            If IsCellLocation(vbCellAddress, trgCell) Then
                ClientUndeleteItem trgCell, False
            End If
            Set trgCell = Nothing
            If IsDeleted(vbMark) Then
                success = False
            End If
        
        ' Like "Undelete Item"
        Case description Like "Undelete Item":
            If IsCellLocation(vbCellAddress, trgCell) Then
                ClientDeleteItem trgCell, vbFrom, False, vbUser, vbDate
            End If
            Set trgCell = Nothing
            If Not IsDeleted(vbMark) Then
                success = False
            End If
        
        ' Like "BEGIN Populate Mark Numbers"
        Case description Like "BEGIN Populate Mark Numbers":
            success = Not recursive
            
        ' Like "END Populate Mark Numbers"
        Case description Like "END Populate Mark Numbers":
            finish_row = row
            Do 'find a stopping point
                finish_row = finish_row - 1
            Loop While VB_CHANGE_LOG.Cells(finish_row, vb_desc_col).Value2 <> "BEGIN Populate Mark Numbers"
            
            VB_CHANGE_LOG.Rows(row).Hidden = True ' hide "END Populate Mark Numbers"
            inProgress = False ' prepare for recursion
            Do
                SetStatusBar "Undoing Populate Mark Numbers...", (LastChangeID() - row), (finish_row - row)
                
                'Undo the individual mark number change
                If Not HMMUndo(True, False) Then
                    MsgBox "Something went wrong. Check the Change Log and to see where the processed left off.", vbCritical
                    success = False
                    Exit Do
                End If
            Loop While VB_CHANGE_LOG.Cells(LastChangeID(), vb_desc_col).Value2 <> "BEGIN Populate Mark Numbers"
            inProgress = True ' end recursion
            
            VB_CHANGE_LOG.Rows(LastChangeID()).Hidden = True ' hide "BEGIN Populate Mark Numbers"
            ResetStatusBar
        
        ' Like "LOCKED Mark Numbers"
        Case description Like "LOCKED Mark Numbers":
            ThisWorkbook.SetMarkNumLock False
            Set TempRibbon = GetRibbon()
            TempRibbon.Invalidate
            Set TempRibbon = Nothing
        
        ' Like "UNLOCKED Mark Numbers"
        Case description Like "UNLOCKED Mark Numbers":
            ThisWorkbook.SetMarkNumLock True
            Set TempRibbon = GetRibbon()
            TempRibbon.Invalidate
            Set TempRibbon = Nothing
        
        ' Like "Added Category: *"
        Case description Like "Added Category: *":
            description = Mid(description, Len("Added Catergory: ")) ' "TEST (EA)"
            cat_name = Left(description, InStr(1, description, " (") - 1)
            DeleteCategory cat_name
        
        ' Like "Purged Category: *"
        Case description Like "Purged Category: *":
            description = Mid(description, Len("Purged Catergory: ")) ' "TEST (EA)"
            cat_name = Left(description, InStr(1, description, " (") - 1)
            cat_unit = Mid(description, InStr(1, description, " (") + 2, Len(description) - InStr(1, description, " (") - 2)
            
            If IsCellLocation(vbCellAddress, trgCell) Then
                If trgCell.row > VB_MASTER.LastRow() + 1 Then
                    Set trgCell = trgCell.offset(RowOffset:=(VB_MASTER.LastRow() + 1 - trgCell.row))
                    VB_CHANGE_LOG.Cells(row, vb_cell_col).Value = VB_MASTER.Name & "!" & trgCell.Address
                End If
                NewCategory cat_name, cat_unit, trgCell.row
            End If
            Set trgCell = Nothing
        
        ' Like "Renamed Category"
        Case description Like "Renamed Category":
            VB_CATEGORY.RenameCategory vbTo, vbFrom
        
        ' Like "Description Check"
        Case description Like "Description Check":
            If IsCellLocation(vbCellAddress, trgCell) Then
                If Len(vbTo) > 0 Then ' undo approval
                    UnapproveDescription Target:=trgCell
                ElseIf Len(vbFrom) > 0 Then ' undo disapproval
                    ApproveDescription Target:=trgCell, inComment:=vbFrom, write_change:=False
                End If
            End If
            Set trgCell = Nothing
        
        ' Like "Added Site: *"
        Case description Like "Added Site: *":
            If SiteExists(Mid(vbTo, 4)) Then
                RemoveSite Mid(vbTo, 4)
            Else
                success = False
            End If
        
        ' Like "Purged Site: *"
        Case description Like "Purged Site: *":
            AddSite Mid(vbFrom, 4), IIf(Mid(vbFrom, 2, 1) = "1", True, False)
            If Not SiteExists(Mid(vbFrom, 4)) Then
                success = False
            Else
                MsgBox Mid(vbFrom, 4) & " has been successfully added back into this project, but all model quantities and RFP records for this site that may have existed " & _
                    "prior to the original purge are now gone.", vbInformation
            End If
        
        ' Like "Renamed *: *"
        Case description Like "Renamed *: *":
            old_site_name = Mid(vbFrom, 4)
            site_name = Mid(vbTo, 4)
            If (SiteExists(site_name) And Not SiteExists(old_site_name)) Or (SiteExists(site_name) And IsFabPackage(site_name) <> IIf(Mid(vbFrom, 2, 1) = "1", True, False)) Then
                RenameSite site_name, old_site_name, IIf(Mid(vbFrom, 2, 1) = "1", True, False)
            Else
                success = False
            End If
        
        ' Like "BEGIN CADWorx Import"
        Case description Like "BEGIN CADWorx Import":
            success = Not recursive
            
        ' Like "END CADWorx Import"
        Case description Like "END CADWorx Import":
            finish_row = row
            Do ' find a stopping point
                finish_row = finish_row - 1
            Loop While VB_CHANGE_LOG.Cells(finish_row, vb_desc_col).Value2 <> "BEGIN CADWorx Import"
        
            VB_CHANGE_LOG.Rows(row).Hidden = True ' hide "END CADWorx Import"
            inProgress = False ' prepare for recursion
            Do
                SetStatusBar "Undoing CADWorx Import...", (LastChangeID() - row), (finish_row - row)
                
                'Undo the individual site quantity change
                If Not HMMUndo(True, False) Then
                    MsgBox "Something went wrong. Check the Change Log and to see where the processed left off.", vbCritical
                    success = False
                    Exit Do
                End If
            Loop While VB_CHANGE_LOG.Cells(LastChangeID(), vb_desc_col).Value2 <> "BEGIN CADWorx Import"
            inProgress = True ' end recursion
            
            VB_CHANGE_LOG.Rows(LastChangeID()).Hidden = True ' hide "BEGIN CADWorx Import"
            ResetStatusBar
        
        ' Like "Published *"
        Case description Like "Published *":
            order_num = Mid(description, Len("Published ") + 1, Len(description) - Len("Published ") - Len(Mid(description, InStr(Len("Published "), description, " dated "))))
            msgResult = MsgBox("Are you sure you want to unpublish Order " & order_num & "? RFP/Orders cannot be republished using 'Undo/Redo'.", vbYesNo)
            If msgResult = vbYes Then
                VB_ORDER_LOG.DeleteOrder order_num
            Else
                success = False
            End If
            
        ' Like "Unpublished Order *"
        Case description Like "Unpublished Order *":
            order_num = Mid(description, Len("Unpublished Order ") + 1, Len(description) - Len("Unpublished Order ") - Len(Mid(description, InStr(Len("Unpublished Order "), description, " for "))))
            order_site = Right(description, Len(description) - InStrRev(description, " for ") - Len(" for ") + 1)
            MsgBox "Cannot republish Order " & order_num & " for " & order_site & ".", vbExclamation
            success = False
        
        ' Like "Issue RFP Client Report"
        Case description Like "Issue RFP Client Report":
            'restore previous revision to "RFP Report Current"
            prev_rev = VB_VAR_STORE.GetClientSummaryRev() - 1
            
            ' request permission from User
            msgResult = MsgBox("Revert RFP Report to previous revision: " & prev_rev & "?", vbYesNo)
            If msgResult = vbYes Then
                'make current files OBSOLETE
                OldFileName = ThisWorkbook.Path & VB_VAR_STORE.GetClientSummaryDirectory() & vbTo
                NewFileName = ThisWorkbook.Path & VB_VAR_STORE.GetClientSummaryDirectory() & "OBSOLETE_" & vbTo
                On Error Resume Next
                If Len(Dir(NewFileName & ".xlsx")) > 0 Then
                    Kill NewFileName & ".xlsx"
                End If
                If Len(Dir(NewFileName & ".pdf")) > 0 Then
                    Kill NewFileName & ".pdf"
                End If
                Name OldFileName & ".pdf" As NewFileName & ".pdf"
                Name OldFileName & ".xlsx" As NewFileName & ".xlsx"
                Err.Clear
                On Error GoTo 0
                
                'check prev_rev#
                If prev_rev = 0 Then
                    'Just Delete "RFP Report Current", if it exists, which it should.
                    Set delRev = ThisWorkbook.Sheets("RFP Report Current")
                    delRev.Visible = xlSheetVisible
                    Application.DisplayAlerts = False
                    On Error Resume Next
                    delRev.Delete
                    Err.Clear
                    On Error GoTo 0
                    Application.DisplayAlerts = True
                    Set delRev = Nothing
                    
                    VB_VAR_STORE.SetClientSummaryRev prev_rev
                
                ElseIf prev_rev > 0 Then ' we have to look at prev-1 and restore "RFP Report Current"
                    prevRevFile = Dir(ThisWorkbook.Path & VB_VAR_STORE.GetClientSummaryDirectory() & VB_RFP_REPORT.ClientSummaryTitle & "*_rev" & prev_rev - 1 & ".xlsx")
                    
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
                        Application.DisplayAlerts = True
                        Set delRev = Nothing
                        
                        ' restore RFP Report Current
                        Dim newCurrBk As Workbook
                        Dim newCurrSht As Worksheet
                        Set newCurrBk = Application.Workbooks.Open(ThisWorkbook.Path & VB_VAR_STORE.GetClientSummaryDirectory() & prevRevFile)
                        newCurrBk.Sheets(1).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
                        Set newCurrSht = ThisWorkbook.ActiveSheet
                        newCurrSht.Name = "RFP Report Current"
                        newCurrSht.Visible = xlSheetHidden

                        ' refresh DB
                        VB_SITEDB.SetClientSummaryAliasToPrevious
                        VB_VAR_STORE.SetClientSummaryRev prev_rev
                        
                        Set newCurrSht = Nothing
                        newCurrBk.Close False
                        Set newCurrBk = Nothing
                    Else
                        ' prevRevFile was not found..
                        msgResult = MsgBox("Unable to locate the right file; it may have been moved or renamed. Browse for the file?", vbYesNo)
                        If msgResult = vbYes Then
                            prevRevFile = FindMissingFile("Find Client Summary _rev" & prev_rev - 1)
                            If prevRevFile <> vbNullString Then
                                GoTo RESTART_UNDO
                            Else
                                ' findMissingFile was cancelled or failed
                                success = False
                            End If
                        Else
                            success = False
                        End If
                    End If
                Else
                    'do nothing: already at rev0..
                End If
            Else
                success = False
            End If
        
        ' Like "BEGIN Description Check Reset"
        Case description Like "BEGIN Description Check Reset":
            success = Not recursive
            
        ' Like "END Description Check Reset"
        Case description Like "END Description Check Reset":
            VB_CHANGE_LOG.Rows(row).Hidden = True ' hide "END Description Check Reset"
            inProgress = False ' prepare for recursion
            Do
                'Undo the individual Description check changes
                If Not HMMUndo(True, False) Then
                    MsgBox "Something went wrong. Check the Change Log and to see where the processed left off.", vbCritical
                    success = False
                    Exit Do
                End If
            Loop While VB_CHANGE_LOG.Cells(LastChangeID(), vb_desc_col).Value2 <> "BEGIN Description Check Reset"
            inProgress = True ' end recursion
            
            VB_CHANGE_LOG.Rows(LastChangeID()).Hidden = True ' hide "BEGIN Description Check Reset"
        
        ' Like "BEGIN Quantity Check Reset*"
        Case description Like "BEGIN Quantity Check Reset*":
            success = Not recursive
            
        ' Like "END Quantity Check Reset*"
        Case description Like "END Quantity Check Reset*":
            VB_CHANGE_LOG.Rows(row).Hidden = True ' hide "END Quantity Check Reset"
            inProgress = False ' prepare for recursion
            Do
                'Undo the individual Site Quantity changes
                If Not HMMUndo(True, False) Then
                    MsgBox "Something went wrong. Check the Change Log and to see where the processed left off.", vbCritical
                    success = False
                    Exit Do
                End If
            Loop While Not VB_CHANGE_LOG.Cells(LastChangeID(), vb_desc_col).Value2 Like "BEGIN Quantity Check Reset*"
            inProgress = True ' end recursion
            
            VB_CHANGE_LOG.Rows(LastChangeID()).Hidden = True ' hide "BEGIN Quantity Check Reset
        
        ' Like "BEGIN Clear Model Quantities"
        Case description Like "BEGIN Clear Model Quantities":
            success = Not recursive
            
        ' Like "END Clear Model Quantities"
        Case description Like "END Clear Model Quantities":
            VB_CHANGE_LOG.Rows(row).Hidden = True ' hide "END Clear Model Quantities"
            inProgress = False ' prepare for recursion
            Do
                'Undo the individual Site Quantity changes
                If Not HMMUndo(True, False) Then
                    MsgBox "Something went wrong. Check the Change Log and to see where the processed left off.", vbCritical
                    success = False
                    Exit Do
                End If
            Loop While Not VB_CHANGE_LOG.Cells(LastChangeID(), vb_desc_col).Value2 Like "BEGIN Clear Model Quantities"
            inProgress = True ' end recursion
            
            VB_CHANGE_LOG.Rows(LastChangeID()).Hidden = True ' hide "BEGIN Clear Model Quantities"
        
        ' Like "Category Change"
        Case description Like "Category Change":
            If vbTo = get_category(get_row(vbMark)) And get_cat_row(vbFrom) <> -1 Then
                Dim new_row As Integer
                new_row = AutoSortItem(vbMark, vbFrom)
                
                If get_category(new_row) = vbFrom Then
                    success = True
                Else
                    success = False
                End If
            Else
                MsgBox "Undo was unsuccessful.", vbExclamation
                success = False
            End If
        
        ' get_col_num(description) <> -1
        Case get_col_num(description) <> -1:
            If IsCellLocation(vbCellAddress, trgCell) Then
                If description = "Current Model Quantities" _
                    Or description = "Model Extras" _
                    Or description = "Checked Quantities" Then
                    
                    change_col = trgCell.column
                Else
                    change_col = get_col_num(description)
                End If
                
                If description = "Mark No." Then
                    VB_MASTER.Cells(trgCell.row, change_col).Value = vbFrom
                Else
                    VB_MASTER.Cells(get_row(vbMark), change_col).Value = vbFrom
                End If
                
                If description = "Checked Quantities" Then
                    On Error Resume Next
                    VB_MASTER.Cells(get_row(vbMark), change_col).Comment.Delete
                    Err.Clear
                    On Error GoTo ResetApp
                End If
                
                If description = "Long Description" Then
                    AutoSortItem vbMark
                ElseIf description = "Mark No." Then
                    reset_mark_num vbTo
                End If
            Else
                success = False
            End If
            Set trgCell = Nothing
            
        Case Else:
            success = False
            
    End Select
    
ResetApp:
    ' if success, hide row to add it to the undo stack, and invalidate the Ribbon
    If success Then
        VB_CHANGE_LOG.Rows(row).Hidden = True
        If Not recursive And (VB_CHANGE_LOG.LastChangeID() = -1 Or Not VB_CHANGE_LOG.Rows(row + 1).Hidden) Then
            Set TempRibbon = GetRibbon()
            If Not TempRibbon Is Nothing Then
                On Error Resume Next
                TempRibbon.Invalidate
                Err.Clear
                On Error GoTo 0
                Set TempRibbon = Nothing
            End If
        End If
    End If
    
    If Not recursive Then
        If ThisWorkbook.ActiveSheet.Name <> shtName Then
            ThisWorkbook.Sheets(shtName).Activate
        End If
        VB_MASTER.CalculateQuantityFormat
    End If
    
    If ui_change Then RenderUI True
    inProgress = False
    
    'RESET TRACK CHANGES
    If change Then
        SetTrackChanges True
    End If
    
    HMMUndo = success
End Function

Public Function HMMRedo(Optional ByVal recursive As Boolean = False, Optional ByVal getPermission As Boolean = True) As Boolean
'Subroutine: HMMRedo - Handles redo commands. The Undo/Redo stack is defined in VB_CHANGE_LOG by hiding and unhiding rows.
'                      Potentially large actions such as Description Check Resets, and CADWorx BOM Imports are preceded by a
'                      prompt seeking the user's permission to continue. Global variable inProgess prevents the user from
'                      initiating multiple Redo's at once.
'Arguments: recursive - This routine utilizes recursion to facilitate potentially large actions such as Description Check
'                       Resets, and CADWorx BOM Imports. Setting recursive to True prevents this routine from performing certain
'                       aesthetic actions, but the actual redo process runs normally.
'           getPermission
'Returns: Boolean - success.

    HMMRedo = False

    'GET ROW ID FOR LAST CHANGE
    Dim row As Integer
    row = LastChangeID()

    ' sanity check
    If row = -1 Then
        If VB_CHANGE_LOG.Rows(first_row).Hidden Then
            row = first_row - 1
        Else
            Exit Function
        End If
    End If
    
    'check if this is a completely undone app
    row = row + 1
    If Not VB_CHANGE_LOG.Rows(row).Hidden Then
        Exit Function
    End If

    'check/set inProgress
    If inProgress Then
        Exit Function
    Else
        inProgress = True
    End If

    'HOLD TRACK CHANGES
    Dim change As Boolean
    change = False
    If TrackChanges Then
        change = True
        SetTrackChanges False
    End If
    
    'SAVE ACTIVESHEET
    Dim shtName As String
    If Not recursive Then
        shtName = ThisWorkbook.ActiveSheet.Name
    End If
    
    'prepare for success
    Dim success As Boolean
    success = True
    
    Dim ui_change As Boolean
    
    Dim trgCell As Range
    Dim vbDate As String
    Dim vbMark As Integer
    Dim vbUser As String
    Dim vbCellAddress As String
    Dim description As String
    Dim vbFrom As String
    Dim vbTo As String
    Dim TempRibbon As IRibbonUI
    Dim change_col As Integer
    Dim site_name As String
    Dim order_num As String
    Dim order_site As String
    Dim msgResult As Integer
    Dim finish_row As Integer
    Dim cat_name As String
    Dim prev_rev As Integer
    Dim OldFileName As String
    Dim NewFileName As String
    Dim delRev As Worksheet
    Dim prevRevFile As String
    Dim cat_unit As String
    Dim old_site_name As String
    
    'INITIALIZE VARIABLES
    vbDate = VB_CHANGE_LOG.Cells(row, vb_date_col).Value
    vbUser = VB_CHANGE_LOG.Cells(row, vb_user_col).Value2
    vbMark = VB_CHANGE_LOG.Cells(row, vb_mark_col).Value2
    vbCellAddress = VB_CHANGE_LOG.Cells(row, vb_cell_col).Value2
    description = VB_CHANGE_LOG.Cells(row, vb_desc_col).Value2
    vbFrom = VB_CHANGE_LOG.Cells(row, vb_from_col).Value2
    vbTo = VB_CHANGE_LOG.Cells(row, vb_to_col).Value2
    
    ui_change = RenderUI(False)
    
    If getPermission Then
        If get_col_num(description) = -1 Then
            msgResult = MsgBox("Redo the following action?" & vbCrLf & vbCrLf & _
                description & "." & IIf(Len(vbFrom) > 0, vbCrLf & "From: " & vbFrom, "") & IIf(Len(vbTo) > 0, vbCrLf & "To: " & vbTo, ""), vbYesNo)
        Else
            msgResult = MsgBox("Set the following property?" & vbCrLf & vbCrLf & _
                description & IIf(Len(vbFrom) > 0, vbCrLf & "From: " & vbFrom, "") & IIf(Len(vbTo) > 0, vbCrLf & "To: " & vbTo, ""), vbYesNo)
        End If
    Else
        msgResult = vbYes
    End If
    
    If msgResult <> vbYes Then
        success = False
        GoTo ResetApp
    End If
    
    Select Case True
    
        ' Like "Added Material Item"
        Case description Like "Added Material Item":
            Dim new_mark As Integer
            new_mark = InsertItemBOM(vbFrom, vbTo)
            If new_mark <> vbMark And get_row(vbMark) = -1 Then
                VB_MASTER.Cells(get_row(new_mark), get_col_num("Mark No.")).Value = vbMark
                new_mark = vbMark
            End If
            If get_property(new_mark, "Long Description") <> vbTo Then
                success = False
            End If
            
        ' Like "Purged Material Item"
        Case description Like "Purged Material Item":
            If get_row(vbMark) <> -1 Then
                DeleteMasterRow get_row(vbMark)
            End If
            If get_row(vbMark) <> -1 Then
                success = False
            End If
        
        ' Like "Delete Item"
        Case description Like "Delete Item":
            If IsCellLocation(vbCellAddress, trgCell) Then
                ClientDeleteItem trgCell, vbTo, False, vbUser, vbDate
            End If
            Set trgCell = Nothing
            If Not IsDeleted(vbMark) Then
                success = False
            End If
        
        ' Like "Undelete Item"
        Case description Like "Undelete Item":
            If IsCellLocation(vbCellAddress, trgCell) Then
                ClientUndeleteItem trgCell, False
            End If
            Set trgCell = Nothing
            If IsDeleted(vbMark) Then
                success = False
            End If
        
        ' Like "END Populate Mark Numbers"
        Case description Like "END Populate Mark Numbers":
            success = Not recursive
        
        ' Like "BEGIN Populate Mark Numbers"
        Case description Like "BEGIN Populate Mark Numbers":
            finish_row = row
            Do ' find a stopping point
                finish_row = finish_row + 1
            Loop While VB_CHANGE_LOG.Cells(finish_row, vb_desc_col).Value2 <> "END Populate Mark Numbers"
                
            VB_CHANGE_LOG.Rows(row).Hidden = False ' unhide "BEGIN Populate Mark Numbers"
            inProgress = False ' prepare for recursion
            Do
                SetStatusBar "Redoing Populate Mark Numbers...", (LastChangeID() - row), (finish_row - row)
                
                'Redo the individual mark number change
                If Not HMMRedo(True, False) Then
                    MsgBox "Something went wrong. Check the Change Log and to see where the processed left off.", vbCritical
                    success = False
                    Exit Do
                End If
            Loop While VB_CHANGE_LOG.Cells(LastChangeID() + 1, vb_desc_col).Value2 <> "END Populate Mark Numbers"
            inProgress = True ' end recursion
            
            VB_CHANGE_LOG.Rows(LastChangeID() + 1).Hidden = False ' unhide "END Populate Mark Numbers"
            ResetStatusBar
            
        ' Like "LOCKED Mark Numbers"
        Case description Like "LOCKED Mark Numbers":
            ThisWorkbook.SetMarkNumLock True
            Set TempRibbon = GetRibbon()
            TempRibbon.Invalidate
            Set TempRibbon = Nothing
        
        ' Like "UNLOCKED Mark Numbers"
        Case description Like "UNLOCKED Mark Numbers":
            ThisWorkbook.SetMarkNumLock False
            Set TempRibbon = GetRibbon()
            TempRibbon.Invalidate
            Set TempRibbon = Nothing
        
        ' Like "Added Category: *"
        Case description Like "Added Category: *":
            description = Mid(description, Len("Added Catergory: ")) ' "TEST (EA)"
            cat_name = Left(description, InStr(1, description, " (") - 1)
            cat_unit = Mid(description, InStr(1, description, " (") + 2, Len(description) - InStr(1, description, " (") - 2)
            
            If IsCellLocation(vbCellAddress, trgCell) Then
                If trgCell.row > VB_MASTER.LastRow() + 1 Then
                    Set trgCell = trgCell.offset(RowOffset:=(VB_MASTER.LastRow() + 1 - trgCell.row))
                    VB_CHANGE_LOG.Cells(row, vb_cell_col).Value = VB_MASTER.Name & "!" & trgCell.Address
                End If
                NewCategory cat_name, cat_unit, trgCell.row
            End If
            Set trgCell = Nothing
        
        ' Like "Purged Category: *"
        Case description Like "Purged Category: *":
            description = Mid(description, Len("Purged Catergory: ")) ' "TEST (EA)"
            cat_name = Left(description, InStr(1, description, " (") - 1)
            DeleteCategory cat_name
        
        ' Like "Renamed Category"
        Case description Like "Renamed Category":
            VB_CATEGORY.RenameCategory vbFrom, vbTo
        
        ' Like "Description Check"
        Case description Like "Description Check":
            If IsCellLocation(vbCellAddress, trgCell) Then
                If Len(vbTo) > 0 Then ' redo approval
                    ApproveDescription Target:=trgCell, inComment:=vbTo, write_change:=False
                ElseIf Len(vbFrom) > 0 Then ' redo disapproval
                    UnapproveDescription Target:=trgCell
                End If
            End If
            Set trgCell = Nothing
        
        ' Like "Added Site: *"
        Case description Like "Added Site: *":
            AddSite Mid(vbTo, 4), IIf(Mid(vbTo, 2, 1) = "1", True, False)
            If Not SiteExists(Mid(vbTo, 4)) Then
                success = False
            End If
        
        ' Like "Purged Site: *"
        Case description Like "Purged Site: *":
            If SiteExists(Mid(vbFrom, 4)) Then
                RemoveSite Mid(vbFrom, 4)
            Else
                success = False
            End If
        
        ' Like "Renamed *: *"
        Case description Like "Renamed *: *":
            old_site_name = Mid(vbFrom, 4)
            site_name = Mid(vbTo, 4)
            If (SiteExists(old_site_name) And Not SiteExists(site_name)) Or (SiteExists(site_name) And IsFabPackage(site_name) <> IIf(Mid(vbTo, 2, 1) = "1", True, False)) Then
                RenameSite old_site_name, site_name, IIf(Mid(vbTo, 2, 1) = "1", True, False)
            Else
                success = False
            End If
        
        ' Like "END CADWorx Import"
        Case description Like "END CADWorx Import":
            success = Not recursive
            
        ' Like "BEGIN CADWorx Import"
        Case description Like "BEGIN CADWorx Import":
            finish_row = row
            Do ' find a stopping point
                finish_row = finish_row + 1
            Loop While VB_CHANGE_LOG.Cells(finish_row, vb_desc_col).Value2 <> "END CADWorx Import"
            
            VB_CHANGE_LOG.Rows(row).Hidden = False ' unhide "BEGIN CADWorx Import"
            inProgress = False ' prepare for recursion
            Do
                SetStatusBar "Redoing CADWorxImport...", (LastChangeID() - row), (finish_row - row)
                
                ' redo the individual site quantity change
                If Not HMMRedo(True, False) Then
                    MsgBox "Something went wrong. Check the Change Log and to see where the processed left off.", vbCritical
                    success = False
                    Exit Do
                End If
            Loop While VB_CHANGE_LOG.Cells(LastChangeID() + 1, vb_desc_col).Value2 <> "END CADWorx Import"
            inProgress = True ' end recursion
            
            VB_CHANGE_LOG.Rows(LastChangeID() + 1).Hidden = False ' unhide "END CADWorx Import"
            ResetStatusBar
        
        ' Like "Published *"
        Case description Like "Published *":
            order_num = Mid(description, Len("Published ") + 1, Len(description) - Len("Published ") - Len(Mid(description, InStr(Len("Published "), description, " dated "))))
            order_site = Right(description, Len(description) - InStrRev(description, " for ") - Len(" for ") + 1)
            MsgBox "Cannot republish Order " & order_num & " for " & order_site & ".", vbExclamation
            success = False
        
        ' Like "Unpublished Order *"
        Case description Like "Unpublished Order *":
            order_num = Mid(description, Len("Unpublished Order ") + 1, Len(description) - Len("Unpublished Order ") - Len(Mid(description, InStr(Len("Unpublished Order "), description, " for "))))
            msgResult = MsgBox("Are you sure you want to unpublish Order " & order_num & "? RFP/Orders cannot be republished using 'Undo/Redo'.", vbYesNo)
            If msgResult = vbYes Then
                VB_ORDER_LOG.DeleteOrder order_num
            Else
                success = False
            End If
        
        ' Like "Issue RFP Client Report"
        Case description Like "Issue RFP Client Report":
            'revive OBSOLETE files
            OldFileName = ThisWorkbook.Path & VB_VAR_STORE.GetClientSummaryDirectory() & "OBSOLETE_" & vbTo
            NewFileName = ThisWorkbook.Path & VB_VAR_STORE.GetClientSummaryDirectory() & vbTo
            
            msgResult = MsgBox("Restore Current RFP Report to Revision: " & vbTo & "?", vbYesNo)
            If msgResult = vbYes Then
                On Error Resume Next
                Name OldFileName & ".pdf" As NewFileName & ".pdf"
                Name OldFileName & ".xlsx" As NewFileName & ".xlsx"
                Err.Clear
                On Error GoTo 0
                
                ' we have to restore "RFP Report Current"
                prevRevFile = Dir(NewFileName & ".xlsx")
                
RESTART_REDO:
                If Len(prevRevFile) > 0 Then ' file found
                    'Delete "RFP Report Current", if it exists, which it should.
                    On Error GoTo CURRENTDOESNTEXIST
                    Set delRev = ThisWorkbook.Sheets("RFP Report Current")
                    delRev.Visible = xlSheetVisible
                    Application.DisplayAlerts = False
                    delRev.Delete
                    Err.Clear
                    On Error GoTo 0
                    Application.DisplayAlerts = True
                    Set delRev = Nothing
CURRENTDOESNTEXIST:
                    'restore next RFP Report Current
                    Dim newCurrBk As Workbook
                    Dim newCurrSht As Worksheet
                    Set newCurrBk = Application.Workbooks.Open(ThisWorkbook.Path & VB_VAR_STORE.GetClientSummaryDirectory() & prevRevFile)
                    newCurrBk.Sheets(1).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
                    Set newCurrSht = ThisWorkbook.ActiveSheet
                    newCurrSht.Name = "RFP Report Current"
                    newCurrSht.Visible = xlSheetHidden
                    
                    'set previous version
                    Dim prev_version As Integer
                    prev_version = CInt(Mid(prevRevFile, _
                        InStr(prevRevFile, "_rev") + 4, _
                        InStr(prevRevFile, ".xlsx") - (InStr(prevRevFile, "_rev") + 4)))
                    'refresh DB
                    VB_VAR_STORE.SetClientSummaryRev (prev_version + 1)
                    VB_SITEDB.SetClientSummaryAliasToMatch
                    
                    Set newCurrSht = Nothing
                    newCurrBk.Close False
                    Set newCurrBk = Nothing
                Else
                    ' prevRevFile was not found..
                    msgResult = MsgBox("Unable to locate the right file; it may have been moved or renamed. Browse for the file?", vbYesNo)
                    If msgResult = vbYes Then
                        prevRevFile = FindMissingFile("Find Client Summary _rev" & prev_version)
                        If prevRevFile <> vbNullString Then
                            GoTo RESTART_REDO
                        Else
                            ' findMissingFile was cancelled or failed
                            success = False
                        End If
                    Else
                        success = False
                    End If
                End If
            Else
                success = False
            End If
        
        ' Like "END Description Check Reset"
        Case description Like "END Description Check Reset":
            success = Not recursive
            
        ' Like "BEGIN Description Check Reset"
        Case description Like "BEGIN Description Check Reset":
            VB_CHANGE_LOG.Rows(row).Hidden = False ' unhide "BEGIN Description Check Reset"
            inProgress = False ' prepare for recrusion
            Do
                'Redo the individual Description check changes
                If Not HMMRedo(True, False) Then
                    MsgBox "Something went wrong. Check the Change Log and to see where the processed left off.", vbCritical
                    success = False
                    Exit Do
                End If
            Loop While VB_CHANGE_LOG.Cells(LastChangeID() + 1, vb_desc_col).Value2 <> "END Description Check Reset"
            inProgress = True ' end recursion
            
            VB_CHANGE_LOG.Rows(LastChangeID() + 1).Hidden = False ' unhide "END Description Check Reset"
            
        ' Like "END Quantity Check Reset*"
        Case description Like "END Quantity Check Reset*":
            success = Not recursive
            
        ' Like "BEGIN Quantity Check Reset*"
        Case description Like "BEGIN Quantity Check Reset*":
            VB_CHANGE_LOG.Rows(row).Hidden = False ' unhide "BEGIN Quantity Check Reset"
            inProgress = False ' prepare for recursion
            Do
                'Redo the individual Site Quantity change
                If Not HMMRedo(True, False) Then
                    MsgBox "Something went wrong. Check the Change Log and to see where the processed left off.", vbCritical
                    success = False
                    Exit Do
                End If
            Loop While Not VB_CHANGE_LOG.Cells(LastChangeID() + 1, vb_desc_col).Value2 Like "END Quantity Check Reset*"
            inProgress = True ' end recursion
            
            VB_CHANGE_LOG.Rows(LastChangeID() + 1).Hidden = False ' unhide "END Quantity Check Reset"
        
        ' Like "END Clear Model Quantities"
        Case description Like "END Clear Model Quantities":
            success = Not recursive
            
        ' Like "BEGIN Clear Model Quantities"
        Case description Like "BEGIN Clear Model Quantities":
            VB_CHANGE_LOG.Rows(row).Hidden = False ' hide "BEGIN Clear Model Quantities"
            inProgress = False ' prepare for recursion
            Do
                'Redo the individual Site Quantity changes
                If Not HMMRedo(True, False) Then
                    MsgBox "Something went wrong. Check the Change Log and to see where the processed left off.", vbCritical
                    success = False
                    Exit Do
                End If
            Loop While Not VB_CHANGE_LOG.Cells(LastChangeID() + 1, vb_desc_col).Value2 Like "END Clear Model Quantities"
            inProgress = True ' end recursion
            
            VB_CHANGE_LOG.Rows(LastChangeID() + 1).Hidden = False ' hide "END Clear Model Quantities"
        
        ' Like "Category Change"
        Case description Like "Category Change":
            If vbFrom = get_category(get_row(vbMark)) And get_cat_row(vbTo) <> -1 Then
                Dim new_row As Integer
                new_row = AutoSortItem(vbMark, vbTo)
                
                If get_category(new_row) = vbTo Then
                    success = True
                Else
                    success = False
                End If
            Else
                MsgBox "Redo was unsuccessful.", vbExclamation
                success = False
            End If
        
        ' get_col_num(description) <> -1
        Case get_col_num(description) <> -1:
            If IsCellLocation(vbCellAddress, trgCell) Then
                If description = "Current Model Quantities" _
                    Or description = "Model Extras" _
                    Or description = "Checked Quantities" Then
                    
                    change_col = trgCell.column
                Else
                    change_col = get_col_num(description)
                End If
                
                If description = "Mark No." Then
                    VB_MASTER.Cells(trgCell.row, change_col).Value = vbTo
                Else
                    VB_MASTER.Cells(get_row(vbMark), change_col).Value = vbTo
                End If
                
                If description = "Checked Quantities" Then
                    On Error Resume Next
                    VB_MASTER.Cells(get_row(vbMark), change_col).Comment.Delete
                    Err.Clear
                    On Error GoTo ResetApp
                    
                    If Not IsEmpty(VB_MASTER.Cells(get_row(vbMark), change_col)) Then
                        VB_MASTER.Cells(get_row(vbMark), change_col).AddComment vbUser & " on " & Format(DateValue(vbDate), "mm/dd/yyyy")
                        VB_MASTER.Cells(get_row(vbMark), change_col).Comment.Shape.TextFrame.AutoSize = True
                    End If
                End If
                
                If description = "Long Description" Then
                    AutoSortItem vbMark
                ElseIf description = "Mark No." Then
                    reset_mark_num vbFrom
                End If
            Else
                success = False
            End If
            Set trgCell = Nothing
        
        Case Else:
            success = False
            
    End Select
    
ResetApp:
    ' if success, hide row to add it to the undo stack, and invalidate the Ribbon
    If success Then
        VB_CHANGE_LOG.Rows(row).Hidden = False
        If Not recursive And (VB_CHANGE_LOG.LastChangeID() = first_row Or VB_CHANGE_LOG.EOF(VB_CHANGE_LOG.LastChangeID() + 1)) Then
            Set TempRibbon = GetRibbon()
            If Not TempRibbon Is Nothing Then
                On Error Resume Next
                TempRibbon.Invalidate
                Err.Clear
                On Error GoTo 0
                Set TempRibbon = Nothing
            End If
        End If
    End If

    If Not recursive Then
        If ThisWorkbook.ActiveSheet.Name <> shtName Then
            ThisWorkbook.Sheets(shtName).Activate
        End If
        VB_MASTER.CalculateQuantityFormat
    End If

    If ui_change Then RenderUI True
    inProgress = False
    
    'RESET TRACK CHANGES
    If change Then
        SetTrackChanges True
    End If
    
    HMMRedo = success
End Function

Public Sub SetProtection()
'Subroutine: SetProtection - Defines protection scheme for the VB_CHANGE_LOG sheet. Users should not edit any of this manually

    VB_CHANGE_LOG.Unprotect
    VB_CHANGE_LOG.Protect DrawingObjects:=False, _
                        Contents:=True, _
                        Scenarios:=False, _
                        UserInterfaceOnly:=True, _
                        AllowFormattingCells:=False, _
                        AllowFormattingColumns:=False, _
                        AllowFormattingRows:=False, _
                        AllowInsertingColumns:=False, _
                        AllowInsertingRows:=False, _
                        AllowInsertingHyperlinks:=False, _
                        AllowDeletingColumns:=False, _
                        AllowDeletingRows:=True, _
                        AllowSorting:=True, _
                        AllowFiltering:=False, _
                        AllowUsingPivotTables:=False
End Sub

Public Function TrackChanges() As Boolean
'Function: TrackChanges - 'Get' method for the track_changes property
'Returns: Boolean - True: changes are being tracked; False: changes are not being tracked


    TrackChanges = Not hold_tracking
End Function

Public Sub SetTrackChanges(ByVal vbVal As Boolean)
'Subroutine: SetTrackChanges - 'Set' method for the track_changes property
'Arguments: vbVal - Boolean containing the new value for track_changes


    hold_tracking = Not vbVal
End Sub

Public Function LastChangeID() As Integer
'Function: LastChangeID - Gets row number of the last change item recorded
'Returns: Integer containing the row number for the last change recorded. If no
'         changes are recorded, returns -1


    LastChangeID = first_row
    
    If IsEmpty(VB_CHANGE_LOG.Cells(LastChangeID, vb_date_col)) Or VB_CHANGE_LOG.Rows(LastChangeID).Hidden Then
        LastChangeID = -1
    Else
        Do While Not IsEmpty(VB_CHANGE_LOG.Cells(LastChangeID + 1, vb_date_col)) And Not VB_CHANGE_LOG.Rows(LastChangeID + 1).Hidden
            LastChangeID = LastChangeID + 1
        Loop
    End If
End Function

Public Function Get_inProgress() As Boolean
'Subroutine: Get_inProgress - 'Get' method for the inProgress state of this sheet.
'Returns: The Boolean state of inProgress
    
    Get_inProgress = inProgress
End Function

Public Sub Set_inProgress(ByVal in_state As Boolean)
'Subroutine: Set_inProgress - 'Set' method for the inProgress state of this sheet.
'Arguments: in_state - The new Boolean state to set
    
    inProgress = in_state
End Sub

Public Function FirstRow() As Integer
'Function: FirstRow - 'Get' method for first_row property in VB_CHANGE_LOG
'Returns: Integer containing the row of the first material item in VB_CHANGE_LOG

    FirstRow = first_row
End Function

Public Function EOF(ByVal row As Integer) As Boolean
'Function: EOF - Tests if row is beyond the scope of this sheet
'Arguments: row - Integer containing the row number in this sheet
'Returns: Boolean. True - row is beyond the scope; False - row is within the meaningful scope
    
    EOF = False
    On Error Resume Next
    EOF = IsEmpty(VB_CHANGE_LOG.Cells(row, vb_date_col))
    Err.Clear
    On Error GoTo 0
End Function
