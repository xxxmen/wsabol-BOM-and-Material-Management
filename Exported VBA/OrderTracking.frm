VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OrderTracking 
   Caption         =   "Order Tracking"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6495
   OleObjectBlob   =   "OrderTracking.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OrderTracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Form: OrderTracking                                  '
'                                                           '
' This form presents the user with a table representing the '
' information in VB_ORDER_LOG. The user can change receipt  '
' and receipt date.                                         '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Public Sub LoadForm()
'Subroutine: LoadForm - Controls the .Show routine for this form. Adds orders to the form, sets position on the screen, etc.

    Dim num_orders As Integer
    num_orders = VB_ORDER_LOG.NumOrders()
    
    If num_orders = 0 Then
        MsgBox "No orders to report.", vbInformation
        Exit Sub
    End If

    Dim order_index As Integer
    order_index = 0
    
    ' Set first Order
    With VB_ORDER_LOG
        OrderLabel0.Caption = .Cells(.FirstRow(), .ColID("Order")).Value2
        OrderDateLabel0.Caption = .Cells(.FirstRow(), .ColID("Date")).Value
        ReceivedCheckBox0.Value = (.Cells(.FirstRow(), .ColID("Receipt")).Value2 = 1)
        ReceiptDateBox0.Value = .Cells(.FirstRow(), .ColID("Receipt Date")).Value
    End With
    
    Dim NewOrderLabel As MSForms.Label
    Dim NewOrderDateLabel As MSForms.Label
    Dim NewReceivedCheckBox As MSForms.CheckBox
    Dim NewReceiptDateBox As MSForms.TextBox
    
    ' Add additional orders
    For order_index = 1 To num_orders - 1
    
        Me.height = Me.height + 18
        Me.OKButton.Top = Me.OKButton.Top + 18
        Me.CancelButton.Top = Me.CancelButton.Top + 18
        
        Set NewOrderLabel = Me.Controls.Add("Forms.Label.1", "OrderLabel" & order_index, True)
        Set NewOrderDateLabel = Me.Controls.Add("Forms.Label.1", "OrderDateLabel" & order_index, True)
        Set NewReceivedCheckBox = Me.Controls.Add("Forms.CheckBox.1", "ReceivedCheckBox & order_index" & order_index, True)
        Set NewReceiptDateBox = Me.Controls.Add("Forms.TextBox.1", "ReceiptDateBox" & order_index, True)
        
        With VB_ORDER_LOG
            NewOrderLabel.Caption = .Cells(.FirstRow() + order_index, .ColID("Order")).Value2
            NewOrderDateLabel.Caption = .Cells(.FirstRow() + order_index, .ColID("Date")).Value
            NewReceivedCheckBox.Value = (.Cells(.FirstRow() + order_index, .ColID("Receipt")).Value2 = 1)
            NewReceivedCheckBox.Caption = "Received?"
            NewReceiptDateBox.Value = .Cells(.FirstRow() + order_index, .ColID("Receipt Date")).Value
        End With
        
        NewOrderLabel.Top = OrderTop(order_index)
        NewOrderDateLabel.Top = OrderTop(order_index)
        NewReceivedCheckBox.Top = OrderTop(order_index) - 4
        NewReceiptDateBox.Top = OrderTop(order_index) - 4
        
        NewOrderLabel.Left = OrderLabel0.Left
        NewOrderDateLabel.Left = OrderDateLabel0.Left
        NewReceivedCheckBox.Left = ReceivedCheckBox0.Left
        NewReceiptDateBox.Left = ReceiptDateBox0.Left
        
    Next order_index
    
    'set position on screen
    Me.Top = Application.Top + Application.height / 2 - Me.height / 2
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    Me.Show
End Sub

Private Sub OKButton_Click()
'Subroutine: OKButton_Click - User Form event handler. Called when user clicks OKButton

    Dim order_index As Integer
    Dim chbx As control
    Dim txbx As control
    
    For Each chbx In Me.Controls
        order_index = OrderNumber(chbx.Top + 4)
    
        If TypeName(chbx) = "CheckBox" Then
            'Received
            Set txbx = Me.Controls("ReceiptDateBox" & order_index)
            
            'if not received, clear receipt date
            If Not chbx.Value Then
                txbx.Value = ""
            End If
            
            With VB_ORDER_LOG
            ' check for receipt state change
                If chbx.Value <> (.Cells(VB_ORDER_LOG.FirstRow() + order_index, .ColID("Receipt")).Value2 = 1) Then
                    .Cells(.FirstRow() + order_index, .ColID("Receipt")).Value = IIf(chbx.Value, 1, 0)
                    .SetReceipt .OrderNumber(order_index), chbx.Value, txbx.Value
                    
                ElseIf txbx.Value <> .Cells(.FirstRow() + order_index, .ColID("Receipt Date")).Value Then ' check for receipt date change
                    .Cells(.FirstRow() + order_index, .ColID("Receipt Date")).Value = txbx.Value
                End If
            End With
        End If
    Next chbx
    
    Set chbx = Nothing
    Set txbx = Nothing

    Unload Me
End Sub

Private Sub CancelButton_Click()
'Subroutine: CancelButton_Click - User Form event handler. Called when user clicks CancelButton

    Unload Me
End Sub

Private Function OrderNumber(ByVal topVal As Integer) As Integer
'Function: OrderNumber - Translates the .Top value for an order control to an order index.
'Arguments: topVal - Integer containing the .Top value to a control on this form
'Returns: Integer containing the order index

    OrderNumber = (topVal - 30) / 18
End Function

Private Function OrderTop(ByVal order_index As Integer) As Integer
'Function: OrderTop - Translates the order_index for an order to the .Top value for the corresponding control.
'Arguments: order_index - Integer containing the order index
'Returns: Integer containing the .Top value to a control on this form

    OrderTop = order_index * 18 + 30
End Function
