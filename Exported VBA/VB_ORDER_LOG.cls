VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VB_ORDER_LOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const first_row As Integer = 2 'row number for first order in the log

Private Const vb_order_col = 1
Private Const vb_date_col = 2
Private Const vb_site_col = 3
Private Const vb_receipt_col = 4
Private Const vb_receipt_date_col = 5

Public Function LogOrder(ByVal order_num As String, ByVal order_date As String, ByVal order_site As String) As Integer
'Function: LogOrder - Records orders in table on sheet VB_ORDER_LOG for reference. Order name, Order Date, and Order site are recorded.
'Arguments: order_num - String containing the Order number
'           order_date - String containing the Order date
'           order_site - String containing the Order site name
'Returns: Integer containing error code. 0: no error; -1: Invalid Aruments; -2: LogChange() failure


    LogOrder = 0
    
    'check validity of arguments
    If Len(order_num) = 0 Or Len(order_date) = 0 Or Len(order_site) = 0 Then
        'throw error
        ErrorHandling "LogOrder", 0, "LogOrder() operation failed: Invalid Arguments", 1
        LogOrder = -1
        Exit Function
    End If

    'Append order to log
    Dim row As Integer
    row = NumOrders() + first_row
    VB_ORDER_LOG.Cells(row, vb_order_col).Value = order_num
    VB_ORDER_LOG.Cells(row, vb_date_col).Value = order_date
    VB_ORDER_LOG.Cells(row, vb_site_col).Value = order_site
    VB_ORDER_LOG.Cells(row, vb_receipt_col).Value = "0"
    VB_ORDER_LOG.Columns("A:C").AutoFit

    'LOG CHANGE
    If VB_CHANGE_LOG.TrackChanges() Then
        If VB_CHANGE_LOG.LogChange(VB_ORDER_LOG.Name & "!$A$" & row, "", "Published " & order_num & " dated " & order_date & " for " & order_site, "", "") <> 0 Then
            'throw error
            LogOrder = -2
            ErrorHandling "LogOrder", 0, "LogChange() operation failed:" & vbCrLf & _
                "LogChange(" & VB_ORDER_LOG.Name & "!$A$" & row & ", , " & order_num & " dated " & order_date & " for " & order_site & ", , )", 1
        End If
    End If
End Function

Public Sub DeleteOrder(ByVal order_num As String)
'Subroutine: DeleteOrder - Removes order from BOM. Deletes record in VB_ORDER_LOG; Removes column
'                          on VB_MASTER. Order item quantities are absorbed by the overall surplus of the item.
'                          Moves the order PDF to "\_DELETED ORDERS". LogChange is called if TrackChanges = True.
'Arguments: order_num - String containing the Order number to delete


    'remove column from MASTER
    Dim mst_col As Integer
    mst_col = get_order_col(order_num)
    
    Dim delete_path As String
    delete_path = ThisWorkbook.Path & VB_VAR_STORE.GetDeletedRFPDirectory()
    
    If Not FileFolderExists(delete_path) Then
        MakeDirs delete_path
        If Not FileFolderExists(delete_path) Then
            MsgBox "For some reason, this application cannot automatically create this folder directory:" & vbCrLf & _
                delete_path & vbCrLf & _
                "Please create this folder manually, try again. Sorry for the inconvenience.", vbCritical
                
            Exit Sub
        End If
    End If
    
    If mst_col = -1 Then
        If OrderIndex(order_num) = -1 Then
            Exit Sub
        Else
            GoTo REMOVEFROMORDERLOG
        End If
    End If

    Dim last As Integer
    last = VB_MASTER.LastRow()
    
    'GET ORDER SITE
    Dim order_site As String
    order_site = get_order_site(order_num)
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    If NumOrders() > 1 Then
        DeleteMasterColumn mst_col  'Already calls Changelog_offset
    Else
        'clear column
        VB_MASTER.Range(ColLet(mst_col) & VB_MASTER.SubtitleRow() & ":" & ColLet(mst_col) & last).ClearContents
        With VB_MASTER.Range(ColLet(mst_col) & VB_MASTER.SubtitleRow() & ":" & ColLet(mst_col) & VB_MASTER.SubtitleRow2()).Font
            .Bold = False
            .Italic = True
        End With
        On Error Resume Next
        VB_MASTER.Range(ColLet(mst_col) & VB_MASTER.SubtitleRow()).Comment.Delete
        Err.Clear
        On Error GoTo 0
    End If
    
REMOVEFROMORDERLOG:
    'remove row from VB_ORDER_LOG
    Dim ol_row As Integer
    ol_row = first_row
    
    Do While Not VB_ORDER_LOG.EOF(ol_row)
        If CStr(VB_ORDER_LOG.Cells(ol_row, vb_order_col).Value2) = order_num Then
            VB_ORDER_LOG.Rows(ol_row).EntireRow.Delete
            Exit Do
        End If
        ol_row = ol_row + 1
    Loop
    
    'Move files to \Unpublished RFP
    
    Dim FSO As Object
    Dim filenameXLSX As String
    Dim filenamePDF As String
    Dim search_result As String
    Dim searchXLSX As String
    Dim searchPDF As String
    Dim temp_version As Integer
    Dim version As Integer
    Dim extension As String
    version = 1
    
    'get filenames
    searchXLSX = Dir(ThisWorkbook.Path & VB_VAR_STORE.GetRFPDirectory() & order_num & "*.xlsx")
    searchPDF = Dir(ThisWorkbook.Path & VB_VAR_STORE.GetRFPDirectory() & order_num & "*.pdf")
    
    'Determine next xlsx versions #
    If Len(searchXLSX) > 0 Then
        search_result = Dir(delete_path & searchXLSX)
        If Len(search_result) > 0 Then
            'previously deleted RFP with this order number and tag already exist. Rev up.
            search_result = Dir(delete_path & Left(searchXLSX, Len(searchXLSX) - 5) & " (*).xlsx")
            Do While Len(search_result) > 0
                temp_version = CInt(Mid(search_result, _
                                    InStr(search_result, " (") + 2, _
                                    InStr(search_result, ".xlsx") - (InStr(search_result, " (") + 2)))
                
                If temp_version > version Then
                    version = temp_version
                End If
                
                search_result = Dir
            Loop
            filenameXLSX = Left(searchXLSX, Len(searchXLSX) - 5) & " (" & (version + 1) & ").xlsx"
        Else
            'no deleted RFPs with this order number
            filenameXLSX = searchXLSX
        End If
    
        'move .xlsx file to _DELETED
        Set FSO = CreateObject("Scripting.FileSystemObject")
        On Error Resume Next
        FSO.movefile ThisWorkbook.Path & VB_VAR_STORE.GetRFPDirectory() & searchXLSX, _
            ThisWorkbook.Path & VB_VAR_STORE.GetDeletedRFPDirectory() & filenameXLSX
        Err.Clear
        On Error GoTo 0
    
    End If
    Set FSO = Nothing
    
    'Determine next pdf versions #
    If Len(searchPDF) > 0 Then
        search_result = Dir(delete_path & searchPDF)
        If Len(search_result) > 0 Then
            'previously deleted RFP with this order number and tag already exist. Rev up.
            search_result = Dir(delete_path & Left(searchPDF, Len(searchPDF) - 4) & " (*).pdf")
            Do While Len(search_result) > 0
                temp_version = CInt(Mid(search_result, _
                                    InStr(search_result, " (") + 2, _
                                    InStr(search_result, ".pdf") - (InStr(search_result, " (") + 2)))
                
                If temp_version > version Then
                    version = temp_version
                End If
                
                search_result = Dir
            Loop
            filenamePDF = Left(searchPDF, Len(searchPDF) - 4) & " (" & (version + 1) & ").pdf"
        Else
            'no deleted RFPs with this order number
            filenamePDF = searchPDF
        End If
    
        'move .xlsx file to _DELETED
        Set FSO = CreateObject("Scripting.FileSystemObject")
        On Error Resume Next
        FSO.movefile ThisWorkbook.Path & VB_VAR_STORE.GetRFPDirectory() & searchPDF, _
            ThisWorkbook.Path & VB_VAR_STORE.GetDeletedRFPDirectory() & filenamePDF
        Err.Clear
        On Error GoTo 0
    
    End If
    Set FSO = Nothing
    
    If VB_CHANGE_LOG.TrackChanges() Then
        If VB_CHANGE_LOG.LogChange(VB_ORDER_LOG.Name & "!$A$" & ol_row, "", "Unpublished Order " & order_num & " for " & order_site, "", "") <> 0 Then
            ErrorHandling "DeleteOrder", 0, "LogChange() operation failed: " & vbCrLf & _
                "LogChange(" & VB_ORDER_LOG.Name & "$A$" & ol_row & ", , Order #" & order_num & " DELETION, , )", 1
        End If
    End If
    
    If ui_change Then RenderUI True
End Sub

Public Sub SetReceipt(ByVal order_number As String, ByVal receipt As Boolean, Optional ByVal receipt_date As String)
'Subroutine: SetReceipt - Modifies the receipt property of an order and handles all subprocedures required
'                         with that: visual indication, receipt date, change log, total ordered and client inventory quantities.
'Arguments: order_number - String containing the Order number
'           receipt - String containing the Order date
'           receipt_date - String containing the Order site name

    
    ' get order index
    Dim order_index As Integer
    order_index = VB_ORDER_LOG.OrderIndex(order_number)
    
    If order_index = -1 Then
        Exit Sub
    End If
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    ' set property in VB_ORDER_LOG
    VB_ORDER_LOG.Cells(first_row + order_index, vb_receipt_col).Value = IIf(receipt, "1", "0")
    
    ' handle receipt date
    If receipt Then
        If Len(receipt_date) > 0 Then
            VB_ORDER_LOG.Cells(first_row + order_index, vb_receipt_date_col).Value = receipt_date
        Else
            VB_ORDER_LOG.Cells(first_row + order_index, vb_receipt_date_col).Value = Format(Now, "m/d/yyyy")
        End If
    Else
        VB_ORDER_LOG.Cells(first_row + order_index, vb_receipt_date_col).ClearContents
    End If
    
    ' Handle visual indication in VB_MASTER
    Dim order_col As Integer
    order_col = get_order_col(order_number)
    
    With VB_MASTER.Range(ColLet(order_col) & VB_MASTER.SubtitleRow() & ":" & ColLet(order_col) & VB_MASTER.SubtitleRow2())
        With .Font
            .Bold = receipt
            .Italic = Not receipt
        End With
        
        .Interior.Color = IIf(receipt, RGB(255, 255, 255), RGB(217, 217, 217))
    End With
    
    If receipt Then
        On Error Resume Next
        VB_MASTER.Cells(VB_MASTER.SubtitleRow(), order_col).Comment.Delete
        Err.Clear
        On Error GoTo 0
    Else
        On Error Resume Next
        VB_MASTER.Cells(VB_MASTER.SubtitleRow(), order_col).AddComment "This order is shown as not received by the client. To change this, use the Order Manager in the Procurement Tab."
        Err.Clear
        On Error GoTo 0
    End If
    
    ' modify "formulas"
    SetTotalOrderedFormula
    
    If ui_change Then RenderUI True
End Sub

Public Function ColID(ByVal vbID As String) As Integer
'Function: ColID - 'Get' method for VB_ORDER_LOG column numbers. Returns the column number for the
'                  specified column ID. If the column ID doesn't exist, returns -1.
'Arguments: vbID - String containing the column ID
'Returns: Integer containing the corresponding column number


    Select Case vbID
        Case "Order":
            ColID = vb_order_col
        Case "Date":
            ColID = vb_date_col
        Case "Site":
            ColID = vb_site_col
        Case "Receipt":
            ColID = vb_receipt_col
        Case "Receipt Date":
            ColID = vb_receipt_date_col
        Case Else:
            ColID = -1
    End Select
End Function

Public Sub SetTotalOrderedFormula(Optional ByVal row_num As Integer = -1)
'Subroutine: SetTotalOrderedFormula - Sets formula for "Total Ordered" column. Order quantities
'                                     are only added if RFP have been received.
'Arguments: row_num - If row_num is specified, formula is set for only that row.


    Dim order_index As Integer
    
    Dim to_form As String
    to_form = "="
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    ' row number is specified
    If row_num <> -1 And get_mark_num(row_num) <> -1 Then
        Do While Not EOF(first_row + order_index)
            'If is received,
            If OrderIsReceived(order_index) Then
                to_form = to_form & IIf(to_form = "=", "", "+") & ColLet(get_order_col(OrderNumber(order_index))) & row_num
            End If
            order_index = order_index + 1
        Loop
        
        If to_form = "=" Then
            to_form = to_form & "0"
        End If
        
        VB_MASTER.Cells(row_num, get_col_num("Total Ordered")).Formula = to_form
        
    Else ' all rows
        Dim first As Integer
        Dim last As Integer
        first = VB_MASTER.FirstRow()
        last = VB_MASTER.LastRow()
        
        For row_num = first To last
            SetTotalOrderedFormula row_num
        Next row_num
    End If
    
    If ui_change Then RenderUI True
End Sub

Public Function TotalRequested(ByVal mark_num As Integer) As Integer
'Function: TotalRequested - Get total material requests for a material item. Sums order quantities regardless
'                           of receipt or order site.
'Arguments: mark_num - Integer containing the mark number for the item in question
'Returns: Integer containing the total requests for the specified item

    Dim item_row As Integer
    item_row = get_row(mark_num)
    
    If mark_num = -1 Then
        TotalRequested = -1
        Exit Function
    End If
    
    Dim col As Integer
    Dim start As Integer
    Dim last As Integer
    start = get_col_num("Orders")
    last = start + NumOrders() - 1
    
    TotalRequested = 0
    
    For col = start To last
        TotalRequested = TotalRequested + CInt(VB_MASTER.Cells(item_row, col).Value2)
    Next col
End Function

Public Function get_order_col(ByVal order_num As String) As Integer
'Function: get_order_col - Retrieves the column number on VB_MASTER, under title "Orders", where the target order is located
'Arguments: order_num - String containing the Order number to find
'Returns: Integer containing the column number

    get_order_col = -1

    Dim col As Integer
    col = get_col_num("Orders")
    
    Do While get_ColTitle(col) = "Orders"
        If CStr(VB_MASTER.Cells(VB_MASTER.SubtitleRow(), col).Value2) = order_num Then
            get_order_col = col
            Exit Do
        End If
        col = col + 1
    Loop
End Function

Public Function get_order_site(ByVal order_num As String) As String
'Function: get_order_site - Retrieves the site name for a given order number
'Arguments: order_num - String containing the Order number
'Returns: String containing the site name


    Dim row As Integer
    row = first_row
    
    get_order_site = vbNullString
    
    Do While Not VB_ORDER_LOG.EOF(row)
        If CStr(VB_ORDER_LOG.Cells(row, vb_order_col).Value2) = order_num Then
            get_order_site = CStr(VB_ORDER_LOG.Cells(row, vb_site_col).Value2)
            Exit Function
        End If
        row = row + 1
    Loop
End Function

Public Function get_order_date(ByVal order_num As String) As String
'Function: get_order_date - Retrieves the order date for a given order number
'Arguments: order_num - String containing the Order number
'Returns: String containing the order date


    Dim row As Integer
    row = first_row
    
    get_order_date = vbNullString
    
    Do While Not VB_ORDER_LOG.EOF(row)
        If CStr(VB_ORDER_LOG.Cells(row, vb_order_col).Value2) = order_num Then
            get_order_date = VB_ORDER_LOG.Cells(row, vb_date_col).text
            Exit Function
        End If
        row = row + 1
    Loop
End Function

Public Function NumOrders(Optional srcSheet As Worksheet) As Integer
'Function: NumOrders - Retrieves the number of orders currently recorded in the BOM, on VB_ORDER_LOG
'Returns: Integer containing the number of orders

    
    If srcSheet Is Nothing Then
        Set srcSheet = VB_ORDER_LOG
    End If
    
    Dim row As Integer
    row = first_row
    
    NumOrders = 0
    
    Do While Not IsEmpty(srcSheet.Cells(row, vb_order_col))
        row = row + 1
        NumOrders = NumOrders + 1
    Loop
End Function

Public Function OrderExists(ByVal order_num As String) As Boolean
'Function: OrderExists - Determines whether a published order number already exists
'Arguments: order_num - String containing the Order number to test
'Returns: Boolean. True - order exists; False - order does not exist

    OrderExists = (get_order_col(order_num) <> -1)
End Function

Public Function OrderIsReceived(ByVal order_index As Integer) As Boolean
'Function: OrderIsReceived - Determines whether a published order has been marked as received by the client
'Arguments: order_index - String containing the Order Index [0, NumOrders)
'Returns: Boolean. True - order has been received; False - order has not been received

    If OrderExists(OrderNumber(order_index)) Then
        OrderIsReceived = (VB_ORDER_LOG.Cells(first_row + order_index, vb_receipt_col).Value2 = "1")
    Else
        OrderIsReceived = False
    End If
End Function

Public Function FirstRow() As Integer
'Function: FirstRow - Retrieves the private attribute of VB_ORDER_LOG: first_row
'Returns: Integer containing the row number for the first order in the log

    FirstRow = first_row
End Function

Public Function OrderNumber(ByVal index As Integer) As String
'Function: OrderNumber - Treats Orders as a collection, and grabs the order number for a certain index
'Arguments: index - Integer containing the Order index  [0, NumOrders() - 1]
'Returns: String containing the order number or vbNullstring if index is out of bounds
    
    OrderNumber = vbNullString
    If index >= 0 And Not EOF(first_row + index) Then OrderNumber = CStr(VB_ORDER_LOG.Cells(first_row + index, vb_order_col).Value2)
End Function

Public Function OrderIndex(ByVal order_number As String) As Integer
'Function: OrderIndex - Treats Orders as a collection, and returns the order index for a certain order number
'Arguments: order_number - String containing the Order number
'Returns: Integer containing the order index or -1 if order doesn't exist

    
    OrderIndex = 0
    Do While Not EOF(OrderIndex + first_row)
        If CStr(VB_ORDER_LOG.Cells(OrderIndex + first_row, vb_order_col).Value2) = order_number Then
            Exit Function
        End If
        
        OrderIndex = OrderIndex + 1
    Loop

    OrderIndex = -1
End Function

Public Function EOF(ByVal row As Integer) As Boolean
'Function: EOF - Evaluates 'End-of-File' state at the specified row
'Arguments: row - row number in sheet
'Returns: Boolean. True - reached/beyond EOF; False - not yet EOF


    EOF = IsEmpty(VB_ORDER_LOG.Cells(row, vb_order_col))
End Function
