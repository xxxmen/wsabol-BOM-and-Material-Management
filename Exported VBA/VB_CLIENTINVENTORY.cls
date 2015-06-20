VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VB_CLIENTINVENTORY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const first_row = 3

Private Const orders_col = 4
Private Const inventory_col = 5
Private Const mark_col = 2
Private Const desc_col = 3

Private Sub Worksheet_Activate()
'Subroutine: Worksheet_Activate - Excel event handler. Called everytime this sheet is actiavted. Sets protection.
    
    SetProtection
End Sub

Public Sub LoadForm()
'Subroutine: LoadForm - Loads this sheet with item descriptions, mark numbers, and supply quantities. This sheet
'                       should be like a form: it has specific fields to fill out (client inventory), and you can't
'                       go into the rest of the workbook until you've properly closed this sheet. Protection is
'                       set for delineating fields, and ThisWorkbook events are set up to handle proper exit strategy.


    'initialize columns
    Dim mst_surplus_col As Integer
    Dim mst_orders_col As Integer
    Dim mst_mark_col As Integer
    Dim mst_desc_col As Integer
    mst_surplus_col = get_col_num("Client Inventory")
    mst_orders_col = get_col_num("Total Ordered")
    mst_mark_col = get_col_num("Mark No.")
    mst_desc_col = get_col_num("Long Description")
    
    Dim first As Integer
    Dim last As Integer
    first = VB_MASTER.FirstRow()
    last = VB_MASTER.LastRow()
    
    Dim row As Integer
    Dim mst_row As Integer
    row = first_row
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    With VB_CLIENTINVENTORY
        .Unprotect
    
        For mst_row = first To last
            
            'Mark No
            .Cells(row, mark_col).Value = VB_MASTER.Cells(mst_row, mst_mark_col).Value2
            .Range(ColLet(mark_col) & row).HorizontalAlignment = xlCenter
            .Range(ColLet(mark_col) & row).VerticalAlignment = xlTop
            .Range(ColLet(mark_col) & row).Font.Italic = True
            .Range(ColLet(mark_col) & row).Font.ColorIndex = 31
            
            'Long Description
            .Cells(row, desc_col).Value = VB_MASTER.Cells(mst_row, mst_desc_col).Value2
            .Range(ColLet(desc_col) & row).HorizontalAlignment = xlLeft
            .Range(ColLet(desc_col) & row).VerticalAlignment = xlTop
            .Range(ColLet(desc_col) & row).Font.Italic = True
            .Range(ColLet(desc_col) & row).Font.ColorIndex = 31
            .Range(ColLet(desc_col) & row).WrapText = True
            .Rows(row).AutoFit
            
            'Total Ordered
            .Cells(row, orders_col).Value = VB_MASTER.Cells(mst_row, mst_orders_col).Value2
            .Range(ColLet(orders_col) & row).HorizontalAlignment = xlCenter
            .Range(ColLet(orders_col) & row).VerticalAlignment = xlTop
            .Range(ColLet(orders_col) & row).Font.Italic = True
            .Range(ColLet(orders_col) & row).Font.ColorIndex = 31
            
            'Client Inventory
            .Cells(row, inventory_col).Value = CInt(.Cells(row, orders_col).Value2) + CInt(VB_MASTER.Cells(mst_row, mst_surplus_col).Value2)
            .Range(ColLet(inventory_col) & row).HorizontalAlignment = xlCenter
            .Range(ColLet(inventory_col) & row).VerticalAlignment = xlTop
            
            row = row + 1
        Next mst_row
        row = row - 1
        
        'Set border formatting
        With .Range(ColLet(mark_col) & first_row & ":" & ColLet(inventory_col) & row).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With .Range(ColLet(mark_col) & first_row & ":" & ColLet(inventory_col) & row).Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        .Range(ColLet(mark_col) & first_row & ":" & ColLet(inventory_col) & row).BorderAround _
            LineStyle:=xlContinuous, Weight:=xlMedium
        
        With .Range(ColLet(orders_col) & first_row & ":" & ColLet(orders_col) & row).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        
        With .Range(ColLet(inventory_col) & first_row & ":" & ColLet(inventory_col) & row).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        
        SetProtection
        
        .Visible = xlSheetVisible
        .Activate
        .Range(ColLet(inventory_col) & first_row & ":" & ColLet(inventory_col) & row).Select
    End With
    
    If ui_change Then RenderUI True
End Sub

Private Sub SetProtection()
'Subroutine: SetProtection - Sets protection scheme for this sheet

    With VB_CLIENTINVENTORY
        .Unprotect
        .UsedRange.Locked = True
        .Protect DrawingObjects:=True, _
                    Contents:=True, _
                    Scenarios:=True, _
                    UserInterfaceOnly:=True, _
                    AllowFormattingCells:=False, _
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
                
        'Set ScrollArea
        .Range(ColLet(inventory_col) & first_row & ":" & ColLet(inventory_col) & .UsedRange.Rows(.UsedRange.Rows.CountLarge).row).Locked = False
        .ScrollArea = ColLet(.UsedRange.Columns(1).column) & .UsedRange.Rows(1).row & ":" & _
            ColLet(.UsedRange.Columns(.UsedRange.Columns.count).column) & .UsedRange.Rows(.UsedRange.Rows.CountLarge).row
    End With
End Sub

Public Sub SaveInventoryForm()
'Subroutine: SaveInventoryForm - When the user presses the "Save & Close" ribbon button, this is called. It saves
'                                the Client Inventory data the user has entered and copies it over to MASTER,
'                                logging all the changes that are made. CloseForm() is called separating in
'                                the RibbonButton handler


    Dim surplus As Integer
    Dim mst_inventory_col As Integer
    Dim mst_mark_col As Integer
    mst_inventory_col = get_col_num("Client Inventory")
    mst_mark_col = get_col_num("Mark No.")

    Dim row As Integer
    row = first_row
    
    Dim old_value As Integer
    Dim mst_row As Integer
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    With VB_CLIENTINVENTORY
    
        'loop and find changes
        For row = first_row To .UsedRange.Rows.CountLarge
            If Not IsEmpty(.Cells(row, mark_col)) Then
                If IsNumeric(.Cells(row, inventory_col).Value2) Then
                    mst_row = get_row(.Cells(row, mark_col).Value2)
                    
                    surplus = .Cells(row, inventory_col).Value2 - .Cells(row, orders_col).Value2
                    
                    If surplus <> VB_MASTER.Cells(mst_row, mst_inventory_col).Value2 Then
                        'save old value
                        old_value = VB_MASTER.Cells(mst_row, mst_inventory_col)
                        
                        'set property and log change
                        Call set_property(.Cells(row, mark_col).Value2, mst_inventory_col, surplus)
                        Call VB_MASTER.WriteChange(VB_MASTER.Cells(mst_row, mst_inventory_col), old_value)
                    End If
                End If
            End If
        Next row
    End With
    
    If ui_change Then RenderUI True
End Sub

Public Sub CloseForm()
'Subroutine: CloseForm - When the user presses the "Save & Close" or "Discard & Close" ribbon button, this is called.
'                        It clears and reformats the sheet so LoadForm starts seamlessly whenever it is called again.


    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    With VB_CLIENTINVENTORY
        .Visible = xlSheetVeryHidden ' hide from user
        
        'unprotect
        .Unprotect
        .ScrollArea = ""
        
        'clear all data
        If .UsedRange.Rows.CountLarge >= first_row Then
            .UsedRange.Rows(first_row & ":" & .UsedRange.Rows.CountLarge).Delete
        End If
    End With
    
    'send user back to MASTER
    VB_MASTER.Activate
    
    If ui_change Then RenderUI True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
'Subroutine: Worksheet_Change - Excel event handler. This subroutine is called when changes are made on VB_CLIENTINVENTORY.
'                               This procedure acts as a forced "Paste Values". Because of the protection, the user can only
'                               make changes to cells on inventory_col, in the table. Set the Borders and Font to match
'                               existing for all changes.
'Arguments: Target - Range object pointing to the cell or cells that have changed.


    VB_CLIENTINVENTORY.Unprotect
    
    Dim vbCell
    
    Dim event_state As Boolean
    event_state = Application.EnableEvents
    Application.EnableEvents = False
    
    'loop for all cells in Target
    For Each vbCell In Target
        
        'sanity check
        If vbCell.column = inventory_col And Not vbCell.Locked Then
        
            'set borders
            With vbCell.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            
            With vbCell.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            
            With vbCell.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = IIf(vbCell.row = first_row, xlMedium, xlThin) ' checking if this is first_row
            End With
            
            With vbCell.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = IIf(vbCell.offset(1, 0).Locked, xlMedium, xlThin) ' checking if this is last row
            End With
            
            'set font
            vbCell.HorizontalAlignment = xlCenter
            vbCell.VerticalAlignment = xlTop
            vbCell.Font.Italic = False
            vbCell.Font.ColorIndex = 1
            vbCell.WrapText = False
            
            If Not IsNumeric(vbCell.Value2) Then
                MsgBox "Inventory values should be numeric.", vbExclamation
                vbCell.ClearContents
            ElseIf vbCell.Value2 <= 0 Or Fix(vbCell.Value2) <> CDbl(vbCell.Value2) Then
                MsgBox "Inventory values should be positive integers.", vbExclamation
                vbCell.ClearContents
            End If
        End If
        
    Next vbCell
    
    SetProtection
    
    Application.EnableEvents = event_state
End Sub
