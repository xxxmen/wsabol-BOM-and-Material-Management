VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteOrderForm 
   Caption         =   "Delete Order ..."
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4665
   OleObjectBlob   =   "DeleteOrderForm.frx":0000
End
Attribute VB_Name = "DeleteOrderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Form: DeleteOrderForm                          '
'                                                     '
' Presents the list of the orders to the user. One is '
' choosen, double-clicked, and DeleteOrder is called. '
'                                                     '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private selection_row As Integer

Public Sub LoadForm(ByRef delete_order As String)
'Subroutine: LoadForm - Controls the .Show routine for this form. Adds item categories to the list box, sets position on the screen, etc.
'Arguments: String reference returning the order number for the selected order.


    Dim total_orders As Integer
    total_orders = 0
    Dim height As Integer
    height = OrdersListBox.height
    Dim ord_num As String
    
    OrdersListBox.ColumnWidths = OrdersListBox.Width / 3
    
    'add orders to list
    Do While total_orders < VB_ORDER_LOG.NumOrders()
        ord_num = VB_ORDER_LOG.OrderNumber(total_orders)
        
        OrdersListBox.AddItem
        OrdersListBox.List(total_orders, 0) = ord_num
        OrdersListBox.List(total_orders, 1) = VB_ORDER_LOG.get_order_site(ord_num)
        OrdersListBox.List(total_orders, 2) = VB_ORDER_LOG.get_order_date(ord_num)
        
        total_orders = total_orders + 1
        
        If total_orders > 1 Then
            OrdersListBox.height = OrdersListBox.height + height
            CancelButton.Top = CancelButton.Top + height
            UnpublishButton.Top = UnpublishButton.Top + height
            Me.height = Me.height + height
        End If
    Loop
    
    'set screen position
    Me.Top = Application.Top + Application.height / 2 - Me.height / 2
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    
    If OrdersListBox.ListCount > 0 Then
        OrdersListBox.Selected(0) = True
        Me.Show
    Else
        selection_row = -1
    End If
    
    If selection_row > 0 Then
        delete_order = VB_ORDER_LOG.OrderNumber(selection_row - 1)
    Else
        delete_order = vbNullString
    End If
    
End Sub

Private Sub CancelButton_Click()
'Subroutine: CancelButton_Click - UserForm event handler. Called when user clicks CancelButton.

    selection_row = -1
    Unload Me
End Sub

Private Sub UnpublishButton_Click()
'Subroutine: UnpublishButton_Click - UserForm event handler. Called when user clicks UnpublishButton.

    selection_row = OrdersListBox.ListIndex + 1
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Subroutine: UserForm_QueryClose - UserForm event handler. Called when form is closed.
'Arguments: Cancel - Integer returning a Cancel action.
'           CloseMode - Integer containing the close mode.


    If CloseMode <> vbFormCode Then 'with an Unload statement
        selection_row = -1
    Else
        ' just close
    End If
End Sub
