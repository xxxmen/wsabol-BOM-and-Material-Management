VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CADWorxImportDialogue 
   Caption         =   "CADWorx Import"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4920
   OleObjectBlob   =   "CADWorxImportDialogue.frx":0000
End
Attribute VB_Name = "CADWorxImportDialogue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Form: CADWorxImportDialogue                            '
'                                                             '
' Presents the uses with options when the CADWorx Import      '
' The user can either continue and finalize the import,       '
' completely back out the import, or just close the dialogue. '
' Ribbon button is unpressed.                                 '
'                                                             '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private vbClick As Integer

Public Function LoadForm() As Integer
'Function: LoadForm - Controls the .Show routine for this form.
'Returns: Integer containing a return code of which button was pressed. Uses standard MsgBox response codes.

    'set position on screen
    Me.Top = Application.Top + Application.height / 2 - Me.height / 2
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    Me.Show

    LoadForm = vbClick
End Function

Private Sub ContinueButton_Click()
'Subroutine: ContinueButton_Click - UserForm event handler. Called when user clicks ContinueButton.

    vbClick = vbYes
    Unload Me
End Sub

Private Sub BackOutButton_Click()
'Subroutine: BackOutButton_Click - UserForm event handler. Called when user clicks BackOutButton.

    vbClick = vbNo
    Unload Me
End Sub

Private Sub CloseButton_Click()
'Subroutine: CloseButton_Click - UserForm event handler. Called when user clicks CloseButton.

    vbClick = vbCancel
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Subroutine: UserForm_QueryClose - UserForm event handler. Called when form is closed.
'Arguments: Cancel - Integer returning a Cancel action.
'           CloseMode - Integer containing the close mode.

    
    If CloseMode <> vbFormCode Then ' with an Unload statement
        vbClick = vbCancel
    Else
        ' just close
    End If
End Sub
