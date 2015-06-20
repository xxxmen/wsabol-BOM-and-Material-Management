VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteItem_Explanation 
   Caption         =   "Explain Yourself..."
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3330
   OleObjectBlob   =   "DeleteItem_Explanation.frx":0000
End
Attribute VB_Name = "DeleteItem_Explanation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Form: DeleteItem_Explaination                        '
'                                                           '
' During Client deletion, a reason is required. If a reason '
' does not already exist, this form is used to get a reason '
' from the user.                                            '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private this_reason As String
Private return_val As Boolean
Private this_master_row As Integer

Public Function SetReason(ByVal master_row As Integer, Optional ByVal in_reason As String = vbNullString) As Boolean
'Subroutine: SetReason - Loads form to get reason from user so it can be set to description.
'Arguments: master_row - row number for item whose deletion needs explaining
'           in_reason - Prescribed reason. Don't load form just set reason.
'Returns: Boolean. True - successful. False - unsuccessful.


    this_master_row = master_row
    
    If in_reason = vbNullString Then
        Me.Top = Application.Top + Application.height / 2 - Me.height / 2
        Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
        Me.Show
    Else
        Reason_TxtBx.Value = in_reason
        SaveReason_Button_Click
    End If
    
    SetReason = return_val
End Function

Private Sub SaveReason_Button_Click()
'Subroutine: SaveReason_Button_Click - User form event handler. Called when the user clicks SaveReason_Button.
'                                      Sets reason to description.

    this_reason = UCase(Reason_TxtBx.text)
    this_reason = TrimWhiteSpace(this_reason)
    
    If Len(this_reason) > 0 Then
        'RenderUI already False
        SetDeleteReason get_mark_num(this_master_row), this_reason
        
        return_val = True
    Else
        return_val = False
    End If
    
    Unload Me
End Sub

Private Sub CancelButton_Click()
'Subroutine: CancelButton_Click - UserForm event handler. Called when user clicks CancelButton.

    this_reason = vbNullString
    Unload Me
End Sub

Private Sub Reason_TxtBx_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Subroutine: Reason_TxtBx_KeyPress - UserForm event handler. Called when user pressing a key down in Reason_TxtBx.
'Arguments: KeyAscii - Integer containing the key code from the key pressed.


    Dim acceptable_char As Boolean
    acceptable_char = CheckKeyCode("Description", KeyAscii)
    
    If Not acceptable_char Then
        MsgBox "You entered an illegal character. Do not use characters " & vbCrLf & vbCrLf & "* _ [ ] ^", vbExclamation
        
        'cancel key pressed
        KeyAscii = 0
    ElseIf Len(Reason_TxtBx.Value) = 0 And KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Subroutine: UserForm_QueryClose - UserForm event handler. Called when form is closed.
'Arguments: Cancel - Integer returning a Cancel action.
'           CloseMode - Integer containing the close mode.

    
    If CloseMode <> vbFormCode Then ' with an Unload statement
        this_reason = vbNullString
    Else
        ' just close
    End If
End Sub
