VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VB_COVERSHEET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const ProjectNameCell = "$F$14"
Private Const ClientNameCell = "$F$16"
Private Const ProjectNumberCell = "$F$20"
Private Const HMMContactCell = "$F$22"
Private Const HMMContactNumberCell = "$K$22"
Private Const SavedByUSERCell = "$F$25"
Private Const SavedByDATECell = "$I$25"

Private Sub Worksheet_Change(ByVal Target As Range)
'Subroutine: Worksheet_Change - Excel event handler. This Sub is called everytime a cell in COVERSHEET is changed.
'                               Updates Project Name, Project number, and Client information throughout the Workbook
'Arguments: target - Range object containing all changed cells
    
    
    VB_ORDER_TMPLT.Calculate
    GetRibbon().Invalidate
End Sub

Private Sub Worksheet_Deactivate()
'Subroutine: Worksheet_Deactivate - Excel event handler. This Sub is called everytime COVERSHEET is deactivated or loses focus.
'                                   Makes sure Project Name, Project number, and Client information are filled out before
'                                   allowing the user to leave the page.


    If (VB_COVERSHEET.Visible <> xlSheetVeryHidden Or Not ThisWorkbook.IsBlankBOM()) _
        And (IsEmpty(VB_COVERSHEET.Range(ProjectNameCell).MergeArea.Cells(1, 1)) _
            Or IsEmpty(VB_COVERSHEET.Range(ClientNameCell).MergeArea.Cells(1, 1))) Then
        
        VB_COVERSHEET.Activate
        MsgBox "Please make sure the relavant project information. Then you may proceed.", vbInformation
    End If
End Sub

Public Sub SetProtection()
'Subroutine: SetProtection - Sets the sheet protection and scroll range for VB_COVERSHEET. The user should only
'                            be allowed to edit 3 cells: Project Name, Project Number, and Client Name
    
    
    VB_COVERSHEET.Unprotect
    VB_COVERSHEET.UsedRange.Locked = True
    VB_COVERSHEET.ScrollArea = VB_COVERSHEET.UsedRange.Address
    
    VB_COVERSHEET.Protect DrawingObjects:=True, _
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
                        
    VB_COVERSHEET.Range(ProjectNameCell).MergeArea.Locked = False
    VB_COVERSHEET.Range(ClientNameCell).MergeArea.Locked = False
    VB_COVERSHEET.Range(ProjectNumberCell).MergeArea.Locked = False
    VB_COVERSHEET.Range(HMMContactCell).MergeArea.Locked = False
    VB_COVERSHEET.Range(HMMContactNumberCell).MergeArea.Locked = False
End Sub

Public Function ProjectName() As String
'Function: ProjectName - Returns whatever the user has filled out for project name on the coversheet
'Returns: String containing the project name

    ProjectName = VB_COVERSHEET.Range(ProjectNameCell).MergeArea.Cells(1, 1).Value2
End Function

Public Function ProjectClientName() As String
'Function: ProjectClientName - Returns whatever the user has filled out for client name on the coversheet
'Returns: String containing the client's name

    ProjectClientName = VB_COVERSHEET.Range(ClientNameCell).MergeArea.Cells(1, 1).Value2
End Function

Public Function ProjectNumber() As String
'Function: ProjectNumber - Returns whatever the user has filled out for project number on the coversheet
'Returns: String containing the project number

    ProjectNumber = VB_COVERSHEET.Range(ProjectNumberCell).MergeArea.Cells(1, 1).Value2
End Function

Public Sub SetProjectName(ByVal project_name As String)
'Subroutine: SetProjectName - Set Project Name on the Coversheet
'Arguments: project_name - String containing the project name

    VB_COVERSHEET.Range(ProjectNameCell).MergeArea.Value = project_name
End Sub

Public Sub SetProjectClientName(ByVal project_client As String)
'Subroutine: SetProjectClientName - Set Project's Client Name on the Coversheet
'Arguments: project_client - String containing the project's client name

    VB_COVERSHEET.Range(ClientNameCell).MergeArea.Value = project_client
End Sub

Public Sub SetProjectNumber(ByVal project_number As String)
'Subroutine: SetProjectNumber - Set Project Number on the Coversheet
'Arguments: project_number - String containing the Project Number

    VB_COVERSHEET.Range(ProjectNumberCell).MergeArea.Value = project_number
End Sub

Public Function HMMContact() As String
'Function: HMMContact - Returns the name and number for the HMM Contact on the Cover page
'Returns: String containing the HMM Contact string

    HMMContact = ""

    If Not IsEmpty(VB_COVERSHEET.Range(HMMContactCell).MergeArea) Then
        HMMContact = VB_COVERSHEET.Range(HMMContactCell).MergeArea.Cells(1, 1).Value2
    End If
    
    If Not IsEmpty(VB_COVERSHEET.Range(HMMContactNumberCell).MergeArea) Then
        HMMContact = HMMContact & " - " & VB_COVERSHEET.Range(HMMContactNumberCell).MergeArea.Cells(1, 1).Value2
    End If
End Function

Public Sub ClearHMMContact()
'Subroutine: ClearHMMContact - Clear the HMMContact Cells on the Coversheet

    VB_COVERSHEET.Range(HMMContactCell).MergeArea.ClearContents
    VB_COVERSHEET.Range(HMMContactNumberCell).MergeArea.ClearContents
End Sub

Public Sub SetLastSavedBy()
'Subroutine: SetLastSavedBy - Sets the Last Saved By values on the Coversheet with the current user and date/time.

    VB_COVERSHEET.Range(SavedByUSERCell).MergeArea.Cells(1, 1).Value = Environ("USERNAME")
    VB_COVERSHEET.Range(SavedByDATECell).MergeArea.Cells(1, 1).Value = Date & " - " & TimeValue(Time)
End Sub

Public Sub ClearLastSavedBy()
'Subroutine: ClearLastSavedBy - Clears the Last Saved By data on the Coversheet

    VB_COVERSHEET.Range(SavedByUSERCell).MergeArea.ClearContents
    VB_COVERSHEET.Range(SavedByDATECell).MergeArea.ClearContents
End Sub
