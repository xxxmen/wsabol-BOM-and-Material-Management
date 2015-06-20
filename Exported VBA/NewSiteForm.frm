VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewSiteForm 
   Caption         =   "New Site..."
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4545
   OleObjectBlob   =   "NewSiteForm.frx":0000
End
Attribute VB_Name = "NewSiteForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Form: NewSiteForm                                      '
'                                                             '
' On a New Site or Rename Site action, this form gives the    '
' user the means to name the new site or provide the new name '
' for an existing site.                                       '
'                                                             '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private new_site_name As String
Private new_site_col As Integer
Private add_column As Boolean
Private gbl_old_name As String

Public Sub LoadForm(ByRef site_name As String, Optional ByVal old_name As String = vbNullString)
'Subroutine: LoadForm - Controls the .Show routine for this form. Sets position on the screen, etc.
'Arguments: site_name - String reference returning the new site name.
'           old_name - Optional (Used only for Rename action). String containing the old site site,
'                      used to find the site's column in VB_MASTER, and distinquishing between
'                      add and rename actions.

    
    VB_VAR_STORE.ClearNewSiteName
    new_site_name = ""
    gbl_old_name = old_name
    
    'ensures a column is not inserted in an "Rename Site" case
    add_column = (Len(old_name) = 0)
    
    'set focus
    If Not add_column Then NewSite_TxtBx.Value = old_name
    NewSite_TxtBx.SetFocus
    
    If Not add_column Then FabCheckBox.Value = IsFabPackage(old_name)
    
    'sets screen position
    Me.Top = Application.Top + Application.height / 2 - Me.height / 2
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    Me.Show
    
    new_site_name = VB_VAR_STORE.GetNewSiteName()
    
    If SiteExists(new_site_name) Then
        site_name = new_site_name
    Else
        site_name = vbNullString
    End If
    
    VB_VAR_STORE.ClearNewSiteName
End Sub

Private Sub AcceptButton_Click()
'Subroutine: AcceptButton_Click - UserForm event handler. Called when user clicks AcceptButton.

    NewSite_TxtBx.Value = TrimWhiteSpace(NewSite_TxtBx.Value)
    new_site_name = NewSite_TxtBx.Value
    
    'sanity check
    If Len(new_site_name) = 0 Then
        MsgBox "Please complete the form.", vbExclamation
        NewSite_TxtBx.SetFocus
        Exit Sub
    End If
    If IsNumeric(new_site_name) Then
        MsgBox "Site name cannot be numeric", vbExclamation
        NewSite_TxtBx.SetFocus
        Exit Sub
    End If
    
    'prefix character check
    If Left(new_site_name, 1) = "'" Then
        new_site_name = Right(new_site_name, Len(new_site_name) - 1)
    End If
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    'Add/Rename Site
    If add_column Then
        AddSite new_site_name, FabCheckBox.Value
    Else
        RenameSite gbl_old_name, new_site_name, FabCheckBox.Value
    End If
    
    ' if site already exists, let user try again.
    If new_site_name <> "ERR" Then
        SiteExists new_site_name, new_site_col
        VB_VAR_STORE.SetNewSiteName new_site_name
    Else
        VB_VAR_STORE.SetNewSiteName ""
        If ui_change Then RenderUI True
        Exit Sub
    End If

    Unload Me
    
    If ui_change Then RenderUI True
End Sub

Private Sub CancelButton_Click()
'Subroutine: CancelButton_Click - UserForm event handler. Called when user clicks CancelButton.

    new_site_name = vbNullString
    Unload Me
End Sub

Private Sub NewSite_TxtBx_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Subroutine: NewSite_TxtBx_KeyPress - UserForm event handler. Called when user pressing a key down in NewSite_TxtBx.
'Arguments: KeyAscii - Integer containing the key code from the key pressed.


    Dim acceptable_char As Boolean
    acceptable_char = CheckKeyCode("Name", KeyAscii)
    
    If Not acceptable_char Then
        MsgBox "You entered an illegal character. Site names cannot contain " & vbCrLf & vbCrLf & "/ ? < > \ : * | " & Chr(34), vbExclamation
        
        'cancel key pressed
        KeyAscii = 0
    ElseIf Len(NewSite_TxtBx.Value) = 0 And KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Subroutine: UserForm_QueryClose - UserForm event handler. Called when form is closed.
'Arguments: Cancel - Integer returning a Cancel action.
'           CloseMode - Integer containing the close mode.

    VB_MASTER.CalculateQuantityFormat
    If CloseMode <> vbFormCode Then ' with an Unload statement
        new_site_name = vbNullString
    Else
        ' just close
    End If
End Sub
