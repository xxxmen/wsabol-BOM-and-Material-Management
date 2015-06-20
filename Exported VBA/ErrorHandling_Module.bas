Attribute VB_Name = "ErrorHandling_Module"
Option Explicit

Public Const Err_Exit = 0
Public Const Err_Resume = 1

Public Sub ErrorHandling(ByVal Parent As String, ByVal ErrValue As Integer, ByVal ErrMessage As String, ReturnValue As Integer)
'Subroutine: ErrorHandling - Avenue for error handling. Handles message, message box, successive action.
'Arguments: Parent - String containing the name of the function that called ErrorHandling()
'           ErrValue - Integer containing the error value.
'           ErrMessage - String containing the error message.
'           ReturnValue - Integer reference returning the successive action code the user chose in the MsgBox.


    Dim MsgBoxResult As Integer
    Dim Choices As Integer
    Dim MsgBoxMessage As String
    
    'Application generated runtime error.
    If ErrMessage = Error(ErrValue) Then
        Select Case ErrValue
            
            Case Else: ' shenanigans
            
                MsgBoxMessage = "i call shenanigans on " & Parent & "() !!..." & vbCrLf & vbCrLf & _
                    "Error " & ErrValue & ": " & ErrMessage & vbCrLf & _
                    "What do you want to do? There is no telling what will happen :)" & vbCrLf & _
                    "OK - resume. CANCEL - end procedure."
                    
                Choices = vbOKCancel
        End Select
        
        ' Display the error message.
        MsgBoxResult = MsgBox(MsgBoxMessage, Choices)
        
        ' Determine the ReturnValue based on the user's choice from MsgBox.
        If MsgBoxResult = vbOK Then
           ReturnValue = Err_Resume
        Else
           ReturnValue = Err_Exit
        End If
        
    Else  'programmer handled error
        
        Select Case ErrValue:
            Case 0: 'logChange error
                MsgBoxMessage = ErrMessage
                Choices = vbOK
            
            Case 10:
                ReturnValue = Err_Resume
                Exit Sub
                
            Case Else:
                Select Case Parent
                    
                    Case Else: ' double shenanigans
                        MsgBoxMessage = ErrMessage
                        Choices = vbOKCancel
                
                End Select
            End Select
        
        ' Display the error message.
        MsgBoxResult = MsgBox("Error in " & Parent & "():" & vbCrLf & MsgBoxMessage, Choices)
        
        ' Determine the ReturnValue based on the user's choice from MsgBox.
        If MsgBoxResult = vbOK Then
           ReturnValue = Err_Resume
        Else
           ReturnValue = Err_Exit
        End If
    End If
End Sub
