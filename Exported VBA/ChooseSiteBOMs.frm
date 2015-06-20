VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChooseSiteBOMs 
   Caption         =   "Site BOMs"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2910
   OleObjectBlob   =   "ChooseSiteBOMs.frx":0000
End
Attribute VB_Name = "ChooseSiteBOMs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''
' User Form: ChooseSiteBOMs                '
'                                          '
' Called to let the user choose which site '
' specific BOMs to create.                 '
'                                          '
''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Public Sub LoadForm(Optional ByVal null_site As String = vbNullString)
'Subroutine: LoadForm - Controls the .Show routine for this form. Loads all sites as check
'                       boxes on the form; Sets position on the screen; etc.


    Dim index As Integer
    Dim rSites As Range
    Set rSites = GetSitesRange()
    
    Dim total_sites As Integer
    total_sites = 0
    
    Dim cb_top As Integer
    Dim cb_left As Integer
    Dim cb_height As Integer
    Dim cb_width As Integer
    cb_left = All_CheckBox.Left
    cb_height = All_CheckBox.height
    cb_width = All_CheckBox.Width
    cb_top = All_CheckBox.Top - cb_height
    
    'add sites
    For index = 1 To rSites.count
        If Not IsEmpty(rSites.Cells(1, index)) Then
            total_sites = total_sites + 1
            Dim chbx As control
            'add check box
            Set chbx = Me.Controls.Add("Forms.CheckBox.1", "Site" & total_sites & "_CheckBox", True)
            With chbx
                .Caption = rSites.Cells(1, index).Value2
                .Left = cb_left
                .Width = rSites.Cells(1, index).ColumnWidth * 4 + 24
                
                If .Left + .Width > Me.Width - 6 Then
                    Me.Width = .Left + .Width + 6
                End If
                
                .Top = cb_top + cb_height
                .Value = True
            End With
            
            If null_site = chbx.Caption Then
                chbx.Value = False
                chbx.enabled = False
            End If
            
            'heighten form window
            If total_sites > 1 Then
                Me.height = Me.height + cb_height
                AcceptSitesButton.Top = AcceptSitesButton.Top + cb_height
                CancelButton.Top = CancelButton.Top + cb_height
            End If
            Set chbx = Nothing
            cb_top = cb_top + cb_height
        End If
    Next index
    
    'set screen position
    Me.Top = Application.Top + Application.height / 2 - Me.height / 2
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    
    If total_sites > 0 Then
        Me.Show
    Else
        MsgBox "There are no sites in this project.", vbInformation
    End If
    
    Set rSites = Nothing
End Sub

Private Sub AcceptSitesButton_Click()
'Subroutine: AcceptSitesButton_Click - UserForm event handler. Called when the user clicks AcceptSitesButton.
'                                      If site check box is checked, call CreateSiteSpecificBOM for that site.


    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    Dim chbx As control
    
    For Each chbx In Me.Controls
        If TypeName(chbx) = "CheckBox" And chbx.Caption <> "All Site(s)" Then
            If chbx.Value Then
                CreateSiteSpecificBOM chbx.Caption
            End If
        End If
    Next chbx
    
    Set chbx = Nothing
    If ui_change Then RenderUI True
    
    Unload Me
End Sub

Private Sub CancelButton_Click()
'Subroutine: CancelButton_Click - UserForm event handler. Called when user clicks CancelButton.

    Unload Me
End Sub

