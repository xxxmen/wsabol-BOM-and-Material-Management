VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VB_SITEDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const first_row = 2

Private Const site_name_col = 1
Private Const bom_rev_col = 2
Private Const alias_col = 3

Public Sub AddSite(ByVal site_name As String)
'Subroutine: AddSite - Adds site to SiteMatrix. This is called whenever AddSite in SiteBOMMembers
'                      in called, after the site is added to MASTER
'Agruments: site_name - name of site that has just been added.


    If Not SiteExists(site_name) Then
        Exit Sub
    End If
    
    Dim row As Integer
    row = first_row
    Do While Not EOF(row)
        row = row + 1
    Loop
    
    VB_SITEDB.Cells(row, site_name_col).Value = site_name
    VB_SITEDB.Cells(row, bom_rev_col).Value = "0"
End Sub

Public Sub RenameSite(ByVal old_site_name As String, site_name As String)
'Subroutine: RenameSite - Renames site in SiteMatrix. This is called whenever RenameSite in SiteBOMMembers
'                      is called, after the site is renamed in MASTER
'Agruments: old_site_name - original name of the site
'           site_name - new name of the site


    If Not SiteExists(site_name) Then
        Exit Sub
    End If
    
    Dim row As Integer
    row = first_row
    Do While Not EOF(row)
        If CStr(VB_SITEDB.Cells(row, site_name_col).Value2) = old_site_name Then
            VB_SITEDB.Cells(row, site_name_col).Value = site_name
            Exit Sub
        End If
        
        row = row + 1
    Loop

End Sub

Public Sub RemoveSite(ByVal site_name As String)
'Subroutine: RemoveSite - Removes site in SiteMatrix. This is called whenever RemoveSite in SiteBOMMembers
'                         in called, after the site is removed from MASTER
'Arguments: site_name - new name of the site

    
    Dim row As Integer
    row = first_row
    Do While Not EOF(row)
        If CStr(VB_SITEDB.Cells(row, site_name_col).Value2) = site_name Then
            VB_SITEDB.Rows(row).Delete
            Exit Sub
        End If
        
        row = row + 1
    Loop
End Sub

Public Function GetSiteBOMRev(ByVal site_name As String) As Integer
'Function: GetSiteBOMRev - Retrieves the SiteBOM rev property from SiteDB for the given site name
'Arguments: site_name - name of site
'Returns: Integer containing the Site's BOM rev#
    
    Dim row As Integer
    row = first_row
    Do While Not EOF(row)
        If CStr(VB_SITEDB.Cells(row, site_name_col).Value2) = site_name Then
            GetSiteBOMRev = VB_SITEDB.Cells(row, bom_rev_col).Value2
            Exit Function
        End If
        
        row = row + 1
    Loop
    
    GetSiteBOMRev = -1
End Function

Public Sub SetSiteBOMRev(ByVal site_name As String, ByVal bom_rev As Integer)
'Subroutine: SetSiteBOMRev - Set the SiteBOM rev# in SiteDB for the given site
'Arguments: site_name - The name of the site to update
'           bom_rev - the new rev number

    Dim row As Integer
    row = first_row
    Do While Not EOF(row)
        If CStr(VB_SITEDB.Cells(row, site_name_col).Value2) = site_name Then
            VB_SITEDB.Cells(row, bom_rev_col).Value = bom_rev
            Exit Sub
        End If
        
        row = row + 1
    Loop
End Sub

Public Function GetClientSummaryAlias(ByVal org_site_name As String) As String
'Function: GetClientSummaryAlias - Retrieves the current name for site in the previous Client Report
'Arguments: org_site_name - the name of site as shown in "RFP Report Current"
'Returns: String containing the most current site name


    Dim row As Integer
    row = first_row
    Do While Not EOF(row)
        If CStr(VB_SITEDB.Cells(row, alias_col).Value2) = org_site_name Then
            GetClientSummaryAlias = CStr(VB_SITEDB.Cells(row, site_name_col).Value2)
            Exit Function
        End If
        
        row = row + 1
    Loop
    
    GetClientSummaryAlias = vbNullString
End Function

Public Sub SetClientSummaryAliasToMatch()
'Subroutine: SetClientSummaryAliasToMatch - When a Client Summary is Issued, site aliases should match the
'                                   current site names. This procedure sets all the aliases to match the current site name.


    Dim row As Integer
    row = first_row
    Do While Not EOF(row)
        If Not IsFabPackage(VB_SITEDB.Cells(row, site_name_col).Value2) Then
            VB_SITEDB.Cells(row, alias_col).Value = VB_SITEDB.Cells(row, site_name_col).Value2
        Else
            VB_SITEDB.Cells(row, alias_col).Value = ""
        End If
        
        row = row + 1
    Loop
End Sub

Public Sub SetClientSummaryAliasToPrevious()
'Subroutine: SetClientSummaryAliasToPrevious - When a Client Summary Issue is undone, site aliases should match the
'                                   site names on "RFP Report Current". This procedure sets all the aliases to
'                                   match "RFP Report Current".


    Dim row As Integer
    row = first_row
    
    Dim prev_rev As Worksheet
    
    If Not SheetExists("RFP Report Current") Then
        Exit Sub
    Else
        Set prev_rev = ThisWorkbook.Sheets("RFP Report Current")
    End If
    
    Do While Not EOF(row)
        If Not IsFabPackage(VB_SITEDB.Cells(row, site_name_col).Value2) Then
            Dim index As Integer
            index = SiteIndex(VB_SITEDB.Cells(row, site_name_col).Value2)
            
            Dim prev_site_col As Integer
            prev_site_col = VB_RFP_REPORT.Form_col_num(prev_rev, "Description") + index + 1
            
            If prev_site_col > 0 Then
                VB_SITEDB.Cells(row, alias_col).Value = prev_rev.Cells(VB_RFP_REPORT.TitleRow(), prev_site_col).Value2
            End If
        Else
            VB_SITEDB.Cells(row, alias_col).Value = ""
        End If
        
        row = row + 1
    Loop
End Sub

Private Function EOF(ByVal row_num As Integer) As Boolean
'Function: EOF - End-of-file checker for SiteDB
'Arguments: row_num - Row number in SiteDB to check

    EOF = IsEmpty(VB_SITEDB.Cells(row_num, site_name_col)) Or row_num < first_row
End Function

