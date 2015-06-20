VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VB_VAR_STORE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const RibbonCell = "$A$1"
Private Const CompSiteCell = "$A$2"
Private Const AutoDescriptorCell = "$AD$1"
Private Const FileTagFormatCell = "$F$1"
Private Const ClientSummaryRevCell = "$R$1"
Private Const NewSiteNameCell = "$NS$1"
Private Const ChosenCategoryCell = "$C$4"

Private Const RFPDirectoryCell = "$F$2"
Private Const ClientSummaryDirectoryCell = "$F$3"
Private Const SiteBOMDirectoryCell = "$F$4"
Private Const MDRDirectoryCell = "$F$5"
Private Const DeletedRFPDirectoryCell = "$F$6"

Public Function GetRibbonID() As Long
'Function: GetRibbonID - 'Read' method for the IRibbonUI id cell
'Returns: Long ID

    GetRibbonID = CLng(VB_VAR_STORE.Range(RibbonCell).Value2)
End Function

Public Sub SetRibbonID(ByVal vbID As Long)
'Subroutine: SetRibbonID - 'Write' method for the IRibbonUI id cell
'Arguments: vbID - Long ID

    VB_VAR_STORE.Range(RibbonCell).Value = vbID
End Sub

Public Sub ClearRibbonID()
'Subroutine: ClearRibbonID - Clear method for the IRibbonUI id cell

    VB_VAR_STORE.Range(RibbonCell).ClearContents
End Sub

Public Function GetComparisonSite() As String
'Function: GetComparisonSite - 'Read' method for the ComparisonSite cell
'Returns: String containing the comparison site

    GetComparisonSite = CStr(VB_VAR_STORE.Range(CompSiteCell).text)
End Function

Public Sub SetComparisonSite(ByVal vbCS As String)
'Subroutine: SetComparisonSite - 'Write' method for the ComparisonSite cell
'Arguments: vbCS - String containing the comparison site

    VB_VAR_STORE.Range(CompSiteCell).Value = vbCS
End Sub

Public Sub ClearComparisonSite()
'Subroutine: ClearComparisonSite - Clear method for the ComparisonSite cell

    VB_VAR_STORE.Range(CompSiteCell).ClearContents
End Sub

Public Function GetAutoDescription() As String
'Function: GetAutoDescription - 'Read' method for the AutoDescription cell
'Returns: String containing the AutoDescriptor result

    GetAutoDescription = CStr(VB_VAR_STORE.Range(AutoDescriptorCell).text)
End Function

Public Sub SetAutoDescription(ByVal vbAD As String)
'Function: SetAutoDescription - 'Write' method for the AutoDescription cell
'Arguments: vbAD - String containing the AutoDescriptor result

    VB_VAR_STORE.Range(AutoDescriptorCell).Value = vbAD
End Sub

Public Sub ClearAutoDescription()
'Subroutine: ClearAutoDescription - Clear method for the AutoDescription cell

    VB_VAR_STORE.Range(AutoDescriptorCell).ClearContents
End Sub

Public Function GetNewSiteName() As String
'Function: GetNewSiteName - 'Read' method for the NewSiteName cell
'Returns: String containing the NewSiteForm result

    GetNewSiteName = CStr(VB_VAR_STORE.Range(NewSiteNameCell).text)
End Function

Public Sub SetNewSiteName(ByVal vbNSN As String)
'Function: SetNewSiteName - 'Write' method for the NewSiteName cell
'Arguments: vbNSN - String containing the NewSiteForm result

    VB_VAR_STORE.Range(NewSiteNameCell).Value = vbNSN
End Sub

Public Sub ClearNewSiteName()
'Subroutine: ClearNewSiteName - Clear method for the NewSiteName cell

    VB_VAR_STORE.Range(NewSiteNameCell).ClearContents
End Sub

Public Function GetChosenCategory() As String
'Function: GetChosenCategory - 'Read' method for the ChosenCategory cell
'Returns: String containing the ChooseCategory result

    GetChosenCategory = CStr(VB_VAR_STORE.Range(ChosenCategoryCell).Value2)
End Function

Public Sub SetChosenCategory(ByVal category As String)
'Function: SetChosenCategory - 'Write' method for the ChosenCategory cell
'Arguments: category - String containing the ChooseCategory result

    VB_VAR_STORE.Range(ChosenCategoryCell).Value = category
End Sub

Public Sub ClearChosenCategory()
'Subroutine: ClearChosenCategory - Clear method for the ChosenCategory cell

    VB_VAR_STORE.Range(ChosenCategoryCell).ClearContents
End Sub

Public Function HMMFileTagFormat(Optional ByVal set_format As String = vbNullString, Optional ByVal sheet As Worksheet) As String
'Function: HMMFileTagFormat - Gets/Sets the HMMFileTag format in VB_VAR_STORE
'Agruments: set_format - String containing the format to set. If not specified, this simply retrieves the current format.
'Returns: The current/new file format
    
    If sheet Is Nothing Then
        Set sheet = VB_VAR_STORE
    End If
    
    If set_format <> vbNullString Then
        sheet.Range(FileTagFormatCell).Value = set_format
    End If
    
    HMMFileTagFormat = CStr(sheet.Range(FileTagFormatCell).Value2)
    
     Set sheet = Nothing
End Function

Public Function GetClientSummaryRev(Optional ByVal sheet As Worksheet) As Integer
'Function: GetClientSummaryRev - Retrieves the Client Summary rev#
'Returns: Integer containing the rev#

    If sheet Is Nothing Then
        Set sheet = VB_VAR_STORE
    End If
    
    GetClientSummaryRev = CInt(sheet.Range(ClientSummaryRevCell).Value2)
    
    Set sheet = Nothing
End Function

Public Sub SetClientSummaryRev(ByVal rev As Integer)
'Subroutine: SetClientSummaryRev - set the Client Summary rev# in VB_VAR_STORE
'Arguments: rev - The next rev number

    VB_VAR_STORE.Range(ClientSummaryRevCell).Value = rev
End Sub

Public Function GetRFPDirectory(Optional ByVal sheet As Worksheet) As String
'Function: GetRFPDirectory - 'Read' method for the RFPDirectory cell
'Returns: String containing the RFP directory

    If sheet Is Nothing Then
        Set sheet = VB_VAR_STORE
    End If

    GetRFPDirectory = CStr(sheet.Range(RFPDirectoryCell).text)
    
    Set sheet = Nothing
End Function

Public Sub SetRFPDirectory(ByVal vbRFPDir As String)
'Function: SetRFPDirectory - 'Write' method for the RFPDirectory cell
'Arguments: String containing the RFP directory

    VB_VAR_STORE.Range(RFPDirectoryCell).Value = vbRFPDir
End Sub

Public Function GetClientSummaryDirectory(Optional ByVal sheet As Worksheet) As String
'Function: GetClientSummaryDirectory - 'Read' method for the ClientSummaryDirectory cell
'Returns: String containing the ClientSummary directory
    
    If sheet Is Nothing Then
        Set sheet = VB_VAR_STORE
    End If
    
    GetClientSummaryDirectory = CStr(sheet.Range(ClientSummaryDirectoryCell).text)
    
    Set sheet = Nothing
End Function

Public Sub SetClientSummaryDirectory(ByVal vbCSDir As String)
'Function: SetRFPDirectory - 'Write' method for the ClientSummaryDirectory cell
'Arguments: String containing the ClientSummary directory

    VB_VAR_STORE.Range(ClientSummaryDirectoryCell).Value = vbCSDir
End Sub

Public Function GetSiteBOMDirectory(Optional ByVal sheet As Worksheet) As String
'Function: GetSiteBOMDirectory - 'Read' method for the SiteBOMDirectory cell
'Returns: String containing the SiteBOM directory
    
    If sheet Is Nothing Then
        Set sheet = VB_VAR_STORE
    End If

    GetSiteBOMDirectory = CStr(sheet.Range(SiteBOMDirectoryCell).text)
    
    Set sheet = Nothing
End Function

Public Sub SetSiteBOMDirectory(ByVal vbSiteBOMDir As String)
'Function: SetSiteBOMDirectory - 'Write' method for the SiteBOMDirectory cell
'Arguments: String containing the SiteBOM directory

    VB_VAR_STORE.Range(SiteBOMDirectoryCell).Value = vbSiteBOMDir
End Sub

Public Function GetMDRDirectory(Optional ByVal sheet As Worksheet) As String
'Function: GetMDRDirectory - 'Read' method for the MDRDirectory cell
'Returns: String containing the MDR directory
    
    If sheet Is Nothing Then
        Set sheet = VB_VAR_STORE
    End If
    
    GetMDRDirectory = CStr(sheet.Range(MDRDirectoryCell).text)
    
    Set sheet = Nothing
End Function

Public Sub SetMDRDirectory(ByVal vbMDRDir As String)
'Function: SetMDRDirectory - 'Write' method for the MDRDirectory cell
'Arguments: String containing the MDR directory

    VB_VAR_STORE.Range(MDRDirectoryCell).Value = vbMDRDir
End Sub

Public Function GetDeletedRFPDirectory(Optional ByVal sheet As Worksheet) As String
'Function: GetDeletedRFPDirectory - 'Read' method for the DeletedRFPDirectory cell
'Returns: String containing the DeletedRFP directory

    If sheet Is Nothing Then
        Set sheet = VB_VAR_STORE
    End If

    GetDeletedRFPDirectory = CStr(sheet.Range(DeletedRFPDirectoryCell).text)
    
    Set sheet = Nothing
End Function

Public Sub SetDeletedRFPDirectory(ByVal vbDeletedRFPDir As String)
'Function: SetDeletedRFPDirectory - 'Write' method for the DeletedRFPDirectory cell
'Arguments: String containing the DeletedRFP directory

    VB_VAR_STORE.Range(DeletedRFPDirectoryCell).Value = vbDeletedRFPDir
End Sub
