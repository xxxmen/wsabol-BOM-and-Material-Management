VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VB_ORDER_TMPLT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const first_row As Integer = 18  'row number for first item in order form

Public Sub OrderFormImport()
'Subroutine: OrderFormImport - This subroutine imports all the data in the template into VB_MASTER and VB_ORDER_LOG
'                              Saves the form and all data as a PDF and a separate workbook named as the order number in
'                              folder "\RFP". Order name is required. If no order date, today's date is used; if no site,
'                              order will go to site name "PROJECT". If an item description in the order doesn't exist,
'                              the PM can Cancel (Exit Sub), Add the Item to VB_MASTER and continue, or
'                              Go Rouge (just continue without adding the item to VB_MASTER).


    Dim Action As Integer
    
    Dim going_rouge As Boolean
    going_rouge = False

    '~~~~~~~~~~~~~~~~~DEFINE CELL LOCATIONS HERE~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Dim date_cell As Range
    Dim onum_cell As Range
    Dim site_cell As Range
    
    Set date_cell = VB_ORDER_TMPLT.Cells(5, 6) 'row/col for origination date
    Set onum_cell = VB_ORDER_TMPLT.Cells(3, 6) 'row/col for order number
    Set site_cell = VB_ORDER_TMPLT.Cells(7, 5) 'row/col for order site
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    

    If IsFormEmpty() Then
        ErrorHandling "OrderImportButton_Click", 0, "Form is Empty.", 1
        Exit Sub 'don't import empty forms
    End If
    
    ' get order date
    If IsEmpty(onum_cell) Then
        MsgBox "Enter an ORDER NUMBER.", vbExclamation
        Exit Sub
    End If
    Dim o_num As String
    o_num = UCase(onum_cell.Value2)
    onum_cell.Value = o_num
    
    'check if order exists, if it does, cancel.
    If VB_ORDER_LOG.OrderExists(o_num) Then
        MsgBox "RFP : " & o_num & vbCrLf & "Already Exists. Use a different order num or unpublish the existing one to continue.", vbCritical
        Exit Sub
    End If
    
    'get order site
    Dim o_site As String
    o_site = site_cell.Value2
    If IsEmpty(site_cell) Then
        'PROJECT WIDE ORDERS ARE STORED UNDER SITENAME "PROJECT"
        o_site = "PROJECT"
    ElseIf Not SiteExists(o_site) Then
        'SITE NAME DOESN'T EXIST
        site_cell.ClearContents
        MsgBox "Enter an existing Site/Area.", vbExclamation
        Exit Sub
    End If
    
    'check order date
    Dim o_date As String
    o_date = date_cell.text
    If Len(o_date) = 0 Then
        o_date = Date
    End If
    If Not IsDate(o_date) Then
        date_cell.ClearContents
        MsgBox "Enter a valid ORIGINATION DATE.", vbExclamation
        Exit Sub
    End If
    date_cell.Value = o_date
    
    'Continue with Publication
On Error GoTo ResetApplication
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    ''''''' FROM THIS POINT ON do NOT use 'Exit Sub' instead use 'GoTo ResetApplication'
    
    Dim tc As Boolean
    tc = VB_CHANGE_LOG.TrackChanges()
    VB_CHANGE_LOG.SetTrackChanges False
        
    'set column numbers
    Dim quantity_col As Integer
    Dim mark_num_col As Integer
    Dim unit_col As Integer
    Dim desc_col As Integer
    Dim vendor_col As Integer
    quantity_col = Form_col_num("Quantity")
    mark_num_col = Form_col_num("Item #")
    unit_col = Form_col_num("Unit")
    desc_col = Form_col_num("Item Description")
    vendor_col = Form_col_num("Unit Price")

    'check validity of form entries.
    Dim row As Integer
    row = first_row
    
    Do While Not EOF(row)
    
        If IsNumeric(VB_ORDER_TMPLT.Cells(row, quantity_col).Value2) And Not IsEmpty(VB_ORDER_TMPLT.Cells(row, quantity_col)) And _
            Not IsEmpty(VB_ORDER_TMPLT.Cells(row, desc_col)) Then
            
            'Quantity and Description are valid,
            
            Dim poss_item_num As Integer
            Dim result As Integer
            Dim vbResult As Integer
            
            If IsNumeric(VB_ORDER_TMPLT.Cells(row, mark_num_col).Value2) And Not IsEmpty(VB_ORDER_TMPLT.Cells(row, mark_num_col)) Then
                'item number doesn't match description
                
                If CStr(get_property(VB_ORDER_TMPLT.Cells(row, mark_num_col).Value2, "Long Description")) <> VB_ORDER_TMPLT.Cells(row, desc_col).Value2 Then
                    'handle: ROUGE PM or ADD ITEM TO BOM
                    
                    poss_item_num = MatchDesc(VB_ORDER_TMPLT.Cells(row, desc_col).Value2)
                    If poss_item_num = -1 Then ' no description match
                        Call NewItemRFPForm.LoadForm(VB_ORDER_TMPLT.Cells(row, desc_col).Value2 & vbCrLf & vbCrLf & "This item is not found in the Master BOM. What would you like to do?", _
                            VB_ORDER_TMPLT.Cells(row, desc_col).Value2, row, result)
                            
                        If result = 2 Then
                            going_rouge = True
                        ElseIf result = 3 Then
                            GoTo ResetApplication
                        End If
                    Else
                        'in this case, the description already exists, but no item number is given on the order form, prompt user if this item should match the existing or not
                        vbResult = MsgBox(VB_ORDER_TMPLT.Cells(row, desc_col).Value2 & vbCrLf & vbCrLf & "This item description is already found in the " & _
                            "Master BOM. It is Item Number " & poss_item_num & "." & vbCrLf & vbCrLf & _
                            "Would you like to use this item from the Master BOM? [Retry]" & vbCrLf & _
                            "Would you like to Cancel so you can edit the description? [Cancel]", vbRetryCancel)
                            
                        If vbResult = vbRetry Then
                            'add Item Number to Form and continue as usual
                            VB_ORDER_TMPLT.Cells(row, mark_num_col).Value = poss_item_num
                        Else
                            GoTo ResetApplication
                        End If
                    End If
                End If
                
            ElseIf IsEmpty(VB_ORDER_TMPLT.Cells(row, mark_num_col)) Then
                'this line has no item number, but what if the description already exists
                
                poss_item_num = MatchDesc(VB_ORDER_TMPLT.Cells(row, desc_col).Value2)
                If poss_item_num = -1 Then 'no description match
                    'handle: ROUGE PM or ADD ITEM TO BOM
                    Call NewItemRFPForm.LoadForm(VB_ORDER_TMPLT.Cells(row, desc_col).Value2 & vbCrLf & vbCrLf & "This item is not found in the Master BOM. What would you like to do?", _
                        VB_ORDER_TMPLT.Cells(row, desc_col).Value2, row, result)
                        
                    If result = 2 Then
                        going_rouge = True
                    ElseIf result = 3 Then
                        GoTo ResetApplication
                    End If
                Else
                    'in this case, the description already exists, but no item number is given on the order form, prompt user if this item should match the existing or not
                    vbResult = MsgBox(VB_ORDER_TMPLT.Cells(row, desc_col).Value2 & vbCrLf & vbCrLf & "This item description is already found in the " & _
                        "Master BOM. It is Item Number " & poss_item_num & "." & vbCrLf & vbCrLf & _
                        "Would you like to use this item from the Master BOM? [Retry]" & vbCrLf & _
                        "Would you like to Cancel so you can edit the description? [Cancel]", vbRetryCancel)
                        
                    If vbResult = vbRetry Then
                        'add Item Number to Form and continue as usual
                        VB_ORDER_TMPLT.Cells(row, mark_num_col).Value = poss_item_num
                    Else
                        GoTo ResetApplication
                    End If
                End If
                
            Else 'item number is not empty and value is not a number
                ErrorHandling "OrderImportButton_Click", 0, "Item" & row & " is not a number.", Action
                GoTo ResetApplication
            End If
        
        Else ' Quantity is empty or not a number, or description is empty
        
            If IsEmpty(VB_ORDER_TMPLT.Cells(row, desc_col)) Then 'if description is empty
                ' do nothing
                
            Else 'quantity is empty or not a number
                If IsEmpty(VB_ORDER_TMPLT.Cells(row, quantity_col)) Then
                    ErrorHandling "OrderImportButton_Click", 0, "Quantity on row " & row & " is empty. Please enter a quantity for this item.", Action
                    GoTo ResetApplication
                ElseIf Not IsNumeric(VB_ORDER_TMPLT.Cells(row, quantity_col).Value2) Then
                    ErrorHandling "OrderImportButton_Click", 0, "Quantity on row " & row & " is not a number.", Action
                    GoTo ResetApplication
                End If
            End If
        End If
        
        row = row + 1
    Loop
    
    'GOING ROUGE means we are ordering material for an item that isn't in MASTER,
    'we aren't adding that item or accounting for the quantity
    
    'insert column in MASTER with column title from Order from
    Dim new_order_col As Integer
    new_order_col = get_col_num("Total Ordered")

    'ADD COLUMN TO MASTER
    new_order_col = NewOrderColumn(new_order_col, o_num, o_site, o_date)
    
    'Set order quantities for the correct mark no's
    row = first_row
    Do While Not EOF(row)
        If Not IsEmpty(VB_ORDER_TMPLT.Cells(row, mark_num_col)) Then
            ' IF GOING ROUGE, get_row with return -1, and set_property will do nothing.
            Call set_property(VB_ORDER_TMPLT.Cells(row, mark_num_col).Value2, new_order_col, VB_ORDER_TMPLT.Cells(row, quantity_col).Value2)
        End If
        row = row + 1
    Loop
    
    'Log order in ORDER_LOG
    If VB_ORDER_LOG.LogOrder(o_num, o_date, o_site) <> 0 Then
        GoTo ResetApplication
    End If
    VB_ORDER_LOG.SetReceipt o_num, False
    
    'Copy completed form to newly dedicated RFP workbook
    Dim wb As Workbook
    VB_ORDER_TMPLT.Visible = xlSheetVisible
    VB_ORDER_TMPLT.Copy
    Set wb = ActiveWorkbook
    wb.Sheets(wb.Worksheets.count).Name = o_num
    
    'SaveAs
    Application.DisplayAlerts = False
    
    Dim filename As String
    filename = o_num & HMMFileTag()
    
    If Not FileFolderExists(ThisWorkbook.Path & VB_VAR_STORE.GetRFPDirectory()) Then
        MakeDirs ThisWorkbook.Path & VB_VAR_STORE.GetRFPDirectory()
        If Not FileFolderExists(ThisWorkbook.Path & VB_VAR_STORE.GetRFPDirectory()) Then
            MsgBox "For some reason, this application cannot automatically create this folder directory:" & vbCrLf & _
                ThisWorkbook.Path & VB_VAR_STORE.GetRFPDirectory() & vbCrLf & _
                "Please create this folder manually, try again. Sorry for the inconvenience.", vbCritical
                
            GoTo ResetApplication
        End If
    End If
    
    On Error Resume Next
    wb.SaveAs ThisWorkbook.Path & VB_VAR_STORE.GetRFPDirectory() & filename & ".xlsx"  'Save file
    wb.Close 'Close file
    On Error GoTo ResetApplication
    Set wb = Nothing
    
    Application.DisplayAlerts = True
    
    'Reactivate BOM
    ThisWorkbook.Activate
    
    'Print sheet to PDF
    Call Print2PDF(VB_ORDER_TMPLT.Name, VB_VAR_STORE.GetRFPDirectory(), filename)

    'Clear template for future use
    date_cell.ClearContents
    onum_cell.ClearContents
    site_cell.ClearContents
    ClearForm
    VB_ORDER_TMPLT.Range("$E$12:$E$15").ClearContents
    
    'move user to MASTER sheet
    VB_MASTER.Activate
    
    VB_CHANGE_LOG.SetTrackChanges tc
    If ui_change Then RenderUI True
    
    Set date_cell = Nothing
    Set onum_cell = Nothing
    Set site_cell = Nothing
    Exit Sub
    
ResetApplication:
    
    VB_ORDER_TMPLT.Activate
    
    RenderUI True
End Sub

Public Sub ClearForm()
'Subroutine: ClearForm - Clears/Empties RFP Order form. Removes all material entires. Removes site name,
'                        order number, and order date.


    Dim quantity_col As Integer
    Dim mark_num_col As Integer
    Dim unit_col As Integer
    Dim desc_col As Integer
    Dim site_col As Integer
    Dim vendor_col As Integer
    Dim row As Integer
    quantity_col = Form_col_num("Quantity")
    mark_num_col = Form_col_num("Item #")
    unit_col = Form_col_num("Unit")
    desc_col = Form_col_num("Item Description")
    vendor_col = Form_col_num("Unit Price")

    Dim ui_change As Boolean
    ui_change = RenderUI(False)

    'Unhide site row
    If VB_ORDER_TMPLT.Rows(7).Hidden Then
        VB_ORDER_TMPLT.Rows(7).Hidden = False
    End If
    'Clear site contents
    VB_ORDER_TMPLT.Cells(7, 5).ClearContents
    
    'Clear date and order number
    VB_ORDER_TMPLT.Cells(5, 7).ClearContents
    VB_ORDER_TMPLT.Cells(3, 7).ClearContents

    'Clear material items
    row = first_row
    Dim USED As Range
    
    With VB_ORDER_TMPLT
        Set USED = .Range(ColLet(.UsedRange.Columns(1).column) & first_row & ":" & _
            ColLet(.UsedRange.Columns(.UsedRange.Columns.count).column) & .UsedRange.Rows(.UsedRange.Rows.count).row)
    End With
    
    USED.ClearContents
    USED.Rows.AutoFit
    
    If ui_change Then RenderUI True
    
    Set USED = Nothing
End Sub

Public Sub SetProtection()
'Subroutine: SetProtection - Sets the sheet proection for VB_ORDER_TMPLT.

    VB_ORDER_TMPLT.Unprotect
    VB_ORDER_TMPLT.Protect DrawingObjects:=True, _
                            Contents:=True, _
                            Scenarios:=True, _
                            UserInterfaceOnly:=True, _
                            AllowFormattingCells:=True, _
                            AllowFormattingColumns:=False, _
                            AllowFormattingRows:=False, _
                            AllowInsertingColumns:=False, _
                            AllowInsertingRows:=False, _
                            AllowInsertingHyperlinks:=False, _
                            AllowDeletingColumns:=False, _
                            AllowDeletingRows:=False, _
                            AllowSorting:=False, _
                            AllowFiltering:=False, _
                            AllowUsingPivotTables:=False
    
    VB_ORDER_TMPLT.UsedRange.Locked = False
    VB_ORDER_TMPLT.Range("$E$2:$E$4").Locked = True
    VB_ORDER_TMPLT.Range("$E$6").Locked = True
    VB_ORDER_TMPLT.Range("$B$17:$G$17").Locked = True
    VB_ORDER_TMPLT.Range("$B$6:$D$15").Locked = True
End Sub

Public Function Form_col_num(ByVal col_title As String) As Integer
'Function: Form_col_num - Retrieves the column number for the specified form heading: "Quantity", "Item #",
'                         "Unit", "Item Description", "Unit Price". If form heading doesn't exist, returns -1.
'Arguments: col_title: String containing the col_title. If not one listed above, heading does not exist.
'Returns: Integer containing the column number of the col_title, or -1.


    Dim col_num As Integer
    
    For col_num = 1 To VB_ORDER_TMPLT.UsedRange.Columns.count
        If UCase(VB_ORDER_TMPLT.Cells(first_row - 1, col_num).Value2) = UCase(col_title) Then
            Form_col_num = col_num
            Exit Function
        End If
    Next col_num
    
    Form_col_num = -1
End Function

Public Function IsFormEmpty() As Boolean
'Function: IsFormEmpty - Determines whether or not the Order Form Template is empty. Only evaluates
'                        rows >= first_row. If all cells are clear, form is empty.
'Returns: Boolean. True - form is empty. False - form is not empty.


    Dim i As Integer
    Dim j As Integer
    i = first_row
    
    Do While Not EOF(i)
        For j = Form_col_num("Quantity") To Form_col_num("Unit Price")
            If Len(VB_ORDER_TMPLT.Cells(i, j).Value2) > 0 Then
                IsFormEmpty = False
                Exit Function
            End If
        Next j
        
        i = i + 1
    Loop
    
    IsFormEmpty = True
End Function

Private Sub Worksheet_Change(ByVal Target As Range)
'Subroutine: Worksheet_Change - Excel event handler. This subroutine is triggered when any selection of
'                               cells is modified on VB_ORDER_TMPLT. When entering item numbers one at a
'                               time by hand, this subroutine autocompletes the description and maintains that
'                               Item # and Quanitity are in fact numbers. If rows or columns are inserted or deleted,
'                               the action is undone.
'Arguments: target - Range object containing the changed cells


    If Not IsCellLocation(Target.Address) Then
        ErrorHandling "VB_ORDER_TMPLT.Worksheet_SelectionChange", 10, "Target Range is not a valid cell reference." & vbCrLf & _
            "target = " & Target.Address & "Terminating Process.", 1
        Exit Sub
    End If
    
    Dim quantity_col As Integer
    Dim mark_num_col As Integer
    Dim unit_col As Integer
    Dim desc_col As Integer
    quantity_col = Form_col_num("Quantity")
    mark_num_col = Form_col_num("Item #")
    unit_col = Form_col_num("Unit")
    desc_col = Form_col_num("Item Description")
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)

    If Target.count = 1 And Target.row >= first_row Then 'single cell change in the form table.
        If Target.Cells.column = mark_num_col Then
            If IsEmpty(Target) Then
                'do nothing
            Else
                If Not IsNumeric(Target.Value2) Then
                    ErrorHandling "VB_ORDER_TMPLT.Worksheet_SelectionChange", 0, "Item number on row " & Target.row & " is not a number. " & _
                        "Please enter a positive natural numbers in this column.", 1
                    
                    Target.ClearContents
                    Target.Select
                ElseIf Fix(Target.Value2) <> CDbl(Target.Value2) Then
                    ErrorHandling "VB_ORDER_TMPLT.Worksheet_SelectionChange", 0, "Item number on row " & Target.row & " is not a natural number. " & _
                        "Please enter a positive natural number in this column.", 1
                    
                    Target.ClearContents
                    Target.Select
                ElseIf get_row(Target.Value2) = -1 Then
                    ErrorHandling "VB_ORDER_TMPLT.Worksheet_SelectionChange", 0, "Item number on row " & Target.row & " does not exist.", 1
                    
                    Target.ClearContents
                    Target.Select
                Else
                    VB_ORDER_TMPLT.Cells(Target.Cells.row, unit_col).Value = get_property(Target.Value2, "Unit")
                    VB_ORDER_TMPLT.Cells(Target.Cells.row, desc_col).Value = get_property(Target.Value2, "Long Description")
                End If
            End If
            
        ElseIf Target.column = quantity_col Then
            If IsEmpty(Target) Then
                'do nothing
            ElseIf Not IsNumeric(Target.Value2) Then
                ErrorHandling "VB_ORDER_TMPLT.Worksheet_SelectionChange", 0, "Quantity on row " & Target.row & " is not a number.", 1
                Target.ClearContents
                Target.Select
            ElseIf CInt(Target.Value) <= 0 Then
                ErrorHandling "VB_ORDER_TMPLT.Worksheet_SelectionChange", 0, "Quantity on row " & Target.row & " is non-positive. " & _
                    "Please enter a positive number for the quantity.", 1
                Target.ClearContents
                Target.Select
            ElseIf Fix(Target.Value) <> CDbl(Target.Value) Then
                ErrorHandling "VB_ORDER_TMPLT.Worksheet_SelectionChange", 0, "Quantity on row " & Target.row & " is not a natural number. " & _
                    "Please enter a natural number for the quantity.", 1
                Target.ClearContents
                Target.Select
            End If
            
        ElseIf Target.column = desc_col Then
            If IsEmpty(Target) Then
                VB_ORDER_TMPLT.Cells(Target.Cells.row, mark_num_col).ClearContents
                VB_ORDER_TMPLT.Cells(Target.Cells.row, unit_col).ClearContents
            Else
                Dim mark_num As Integer
                mark_num = MatchDesc(Target.Value2)
                If mark_num > 0 Then
                    VB_ORDER_TMPLT.Cells(Target.Cells.row, mark_num_col).Value = mark_num
                    VB_ORDER_TMPLT.Cells(Target.Cells.row, unit_col).Value = get_property(mark_num, "Unit")
                Else
                    VB_ORDER_TMPLT.Cells(Target.Cells.row, mark_num_col).ClearContents
                End If
            End If
        End If
    End If
    
    If ui_change Then RenderUI True
    
    Set Target = Nothing
End Sub

Public Function AutoCompleteForm(ByVal site_name As String, ByVal surplus As Boolean, ByVal longlead As Boolean) As Integer
'Function: AutoCompleteForm - Autocompletes VB_ORDER_TMPLT based on the specified site and surplus arguments.
'Arguments: site_name - String containing a site name. If site_name = "", continues using all sites.
'           surplus - Boolean signify whether to take quantity from "Model Total" or "Quantity Needed".
'                     True: pull from "Quantity Needed"; False: total order, pull from "Model Total"
'           longlead - Boolean signify whether to to only include long lead items.
'Returns: Integer containing the number of items written to the form. Returns -1 if error.
    
    
    If Not IsFormEmpty() Then
        MsgBox "This form is not empty. Will not Autocomplete. Please empty the RFP form first before trying again.", vbCritical
        AutoCompleteForm = -1
        Exit Function
    End If
    
    'start as empty form.
    AutoCompleteForm = 0
    
    ' Initialize constants
    Dim order_row As Integer
    Dim master_row As Integer
    'Dim surplus_col As Integer
    Dim quantity_col As Integer
    Dim mark_col As Integer
    Dim site_col As Integer
    Dim unit_col As Integer
    Dim desc_col As Integer
    Dim longlead_col As Integer

    unit_col = get_col_num("Unit")
    desc_col = get_col_num("Long Description")
    master_row = VB_MASTER.FirstRow()
    order_row = first_row
    mark_col = get_col_num("Mark No.")
    longlead_col = get_col_num("Long Lead")
    'surplus_col = get_col_num("Client Inventory")
    
    Dim f_quantity As Integer
    Dim f_mark As Integer
    Dim f_unit As Integer
    Dim f_desc As Integer
    f_quantity = Form_col_num("Quantity")
    f_mark = Form_col_num("Item #")
    f_unit = Form_col_num("Unit")
    f_desc = Form_col_num("Item Description")
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    If SiteExists(site_name, site_col) Then
        'site exists
        
        VB_ORDER_TMPLT.Cells(7, 5).Value = site_name
        If surplus Then
            quantity_col = -1
        Else
            quantity_col = site_col
        End If
    ElseIf Len(site_name) = 0 Then
        ' all sites
        
        If surplus Then
            quantity_col = get_col_num("Quantity Needed")
        Else
            quantity_col = get_col_num("Model Total")
        End If
    Else
        ErrorHandling "AutoCompleteForm", 10, "Site: " & site_name & " doesn't exist.", 1
        AutoCompleteForm = -1
        GoTo ResetApp
    End If
    
    Dim quantity As Integer
    
    ' loop and append items and quantities to the form
    Do While Not IsEmpty(VB_MASTER.Cells(master_row, mark_col))
        
        If Not longlead Or CBool(VB_MASTER.Cells(master_row, longlead_col).Value2) Then
            If quantity_col <> -1 Then
                quantity = CInt(VB_MASTER.Cells(master_row, quantity_col).Value2)
            Else
                'this case is site/surplus; use TotalOrder4SiteItem to define surplus quantity
                quantity = SiteDeficit(site_name, get_mark_num(master_row))
            End If
            
            If quantity > 0 And Not IsDeleted(get_mark_num(master_row)) Then
                VB_ORDER_TMPLT.Cells(order_row, f_quantity).Value = quantity
                VB_ORDER_TMPLT.Cells(order_row, f_mark).Value = VB_MASTER.Cells(master_row, mark_col).Value2
                VB_ORDER_TMPLT.Cells(order_row, f_unit).Value = VB_MASTER.Cells(master_row, unit_col).Value2
                VB_ORDER_TMPLT.Cells(order_row, f_desc).Value = VB_MASTER.Cells(master_row, desc_col).Value2
                order_row = order_row + 1
                AutoCompleteForm = AutoCompleteForm + 1
            End If
    
            VB_ORDER_TMPLT.Rows(order_row - 1).AutoFit
            
            VB_ORDER_TMPLT.Rows(order_row - 1).RowHeight = _
                VBRoundDown(VB_ORDER_TMPLT.Rows(order_row - 1).RowHeight) + 3
        End If

        master_row = master_row + 1
    Loop
    
ResetApp:
    If ui_change Then RenderUI True
    
End Function

Private Function EOF(ByVal row As Integer) As Boolean
'Function: EOF - Evaluates 'End-of-File' state at the specified row
'Arguments: row - row number in sheet
'Returns: Boolean. True - reached/beyond EOF; False - not yet EOF


    Dim empty_count As Integer
    empty_count = 0
    
    Dim col As Integer
    col = Form_col_num("Item Description")
    
    Do While IsEmpty(VB_ORDER_TMPLT.Cells(row, col)) And empty_count < 36
        empty_count = empty_count + 1
        row = row + 1
    Loop
    
    EOF = (empty_count = 36)
End Function

Public Function FirstRow() As Integer
'Function: FirstRow - Retrieves the private attribute of VB_ORDER_TMPLT: first_row
'Returns: Integer containing the row number for the first item in the order form

    FirstRow = first_row
End Function
