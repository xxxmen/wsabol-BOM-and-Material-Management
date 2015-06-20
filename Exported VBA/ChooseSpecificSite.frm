VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChooseSpecificSite 
   Caption         =   "Choose Site/Area ..."
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2535
   OleObjectBlob   =   "ChooseSpecificSite.frx":0000
End
Attribute VB_Name = "ChooseSpecificSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Form: ChooseSpecificSite                               '
'                                                             '
' Form is used during multiple procedures. This form lets the '
' user choose a single site form a list box to specify a site '
' for the upcoming action.                                    '
'                                                             '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private selection_row As Integer

Public Sub LoadForm(ByRef site_select As String)
'Subroutine: LoadForm - Controls the .Show routine for this form. Adds sites to the list box, sets position on the screen, etc.
'Arguments: site_select - String reference returning the name of the site selected.


    Dim height As Integer
    height = Sites_ListBox.height
    
    Dim index As Integer
    Dim rSites As Range
    Set rSites = GetSitesRange()
    
    Dim total_sites As Integer
    total_sites = 0
    
    'load list box
    For index = 1 To rSites.count
        If Not IsEmpty(rSites.Cells(1, index)) Then
            total_sites = total_sites + 1
            Sites_ListBox.AddItem rSites.Cells(1, index).Value2
    
            If total_sites > 1 Then
                Me.height = Me.height + height
                AcceptSite_Button.Top = AcceptSite_Button.Top + height
                CancelButton.Top = CancelButton.Top + height
                Sites_ListBox.height = Sites_ListBox.height + height
            End If
        End If
    Next index
    
    ' set screen position
    Me.Top = Application.Top + Application.height / 2 - Me.height / 2
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    
    If Sites_ListBox.ListCount > 0 Then
        Sites_ListBox.Selected(0) = True
        Me.Show
        
        ' set returns.
        If selection_row >= 0 Then
            site_select = VB_MASTER.Cells(VB_MASTER.SubtitleRow(), get_col_num("Current Model Quantities") + selection_row).Value2
        Else
            site_select = vbNullString
        End If
    
    Else
        MsgBox "No Sites/Areas to choose from.", vbInformation
    End If
    
    Set rSites = Nothing
End Sub

Private Sub AcceptSite_Button_Click()
'Subroutine: AcceptSite_Button_Click - UserForm event handler. Called when user clicks AcceptSite_Button.

    selection_row = Sites_ListBox.ListIndex
    Unload Me
End Sub

Private Sub CancelButton_Click()
'Subroutine: CancelButton_Click - UserForm event handler. Called when user clicks CancelButton.

    selection_row = -1
    Unload Me
End Sub

Private Sub Sites_ListBox_Click()
'Subroutine: Sites_ListBox_Click - UserForm event handler. Called when user clicks inside Sites_ListBox;
'                                  (i.e. when the user chooses a site in the list box). Sets selection_row.

    selection_row = Sites_ListBox.ListIndex
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Subroutine: UserForm_QueryClose - UserForm event handler. Called when form is closed.
'Arguments: Cancel - Integer returning a Cancel action.
'           CloseMode - Integer containing the close mode.


    If CloseMode <> vbFormCode Then ' with an Unload statement
        selection_row = -1
    Else
        ' just close
    End If
End Sub
