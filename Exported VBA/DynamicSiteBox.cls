VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DynamicSiteBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''
' Special Class of MSForm.TextBox with _Change() event
' prescribed for the EidtItemWindow
'''''''''''''''''''''''''''''''''''''''

Option Explicit

' Basis: MSForms.TextBox
Private WithEvents ThisSiteBox As MSForms.TextBox
Attribute ThisSiteBox.VB_VarHelpID = -1

'Set property to define MSForms.TextBox object
Public Property Set SetTextBox(msBox As MSForms.TextBox)
    Set ThisSiteBox = msBox
End Property

'Assigns value to MSForms.TextBox object
Public Sub AssignValue(val)
    ThisSiteBox.Value = val
End Sub

'Sets Focus to this MSForms.TextBox object
Public Sub SetFocus()
    ThisSiteBox.SetFocus
End Sub

'Get .Value from MSForms.TextBox object
Public Function Value()
    Value = ThisSiteBox.Value
End Function

'_Change() event
Private Sub ThisSiteBox_Change()
    EditItemWindow.UpdateSaveButtonColor
End Sub

