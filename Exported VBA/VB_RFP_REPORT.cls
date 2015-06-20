VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VB_RFP_REPORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Set position constants
Private Const first_row As Integer = 11
Private Const title_row As Integer = 9
Private Const sap_col As Integer = 2
Private Const mark_col As Integer = 3
Private Const unit_col As Integer = 4
Private Const desc_col As Integer = 5
Private Const site1_col As Integer = 6

Private Const document_title = "CLIENT_SUMMARY"

Private Const min_orders_visible = 3

Public Function RFPReportRevision() As Integer
'Function: RFPReportRevision - Determined the current RFP Report revision based on the existing RFP Reports
'                              backed up in "\RFP CLIENT REPORTS".
'Returns: Integer containing the current revision number


    Dim search_result As String
    Dim temp_version As Integer
    Dim extension As String
    Dim RFPPath As String
    RFPPath = ThisWorkbook.Path & VB_VAR_STORE.GetClientSummaryDirectory()
    RFPReportRevision = 0
    
    'Search for relevant documents
    search_result = Dir(RFPPath & document_title & "*_rev*." & "*")
    If Len(search_result) > 0 Then
        'search yielded results. Extract rev# from file name for each result. Track the highest one.
        Do While Len(search_result) > 0
            If Right(search_result, 4) = ".pdf" Then
                extension = ".pdf"
            Else
                extension = ".xlsx"
            End If
            
            temp_version = CInt(Mid(search_result, _
                                InStr(search_result, "_rev") + 4, _
                                InStr(search_result, extension) - (InStr(search_result, "_rev") + 4)))
            
            If temp_version > RFPReportRevision Then
                RFPReportRevision = temp_version
            End If
            
            search_result = Dir
        Loop
        
        'Rev Up
        RFPReportRevision = RFPReportRevision + 1
    End If
    
End Function

Public Function FindRFPReportRevision(ByVal revision As Integer) As String
'Function: FindRFPReportRevision - Searches a specific revision in the Client Summary directory.
'Arguments: revision - Integer containing the revision to search for
'Returns: If revision exists, String containing the filename corresponding to the revision, otherwise return vbNullString.

    Dim search_result As String
    Dim extension As String
    Dim RFPPath As String
    RFPPath = ThisWorkbook.Path & VB_VAR_STORE.GetClientSummaryDirectory()
    
    'Search for relevant documents
    search_result = Dir(RFPPath & document_title & "*_rev" & revision & "." & "*")
    If Len(search_result) > 0 Then
        'search yielded results. Extract rev# from file name for each result. Track the highest one.
        If Right(search_result, 4) = ".pdf" Then
            extension = ".pdf"
        Else
            extension = ".xlsx"
        End If
        
        FindRFPReportRevision = Left(search_result, Len(search_result) - Len(extension))
    Else
        FindRFPReportRevision = vbNullString
    End If
    
End Function

Private Function LastRFPRow() As Integer
'Function: LastRFPRow - Finds the last row in this report: The last row is after the last row of data on the sheet.
'Returns: Integer containing the last row of this report.

    
    'Search and return.
    For LastRFPRow = VB_RFP_REPORT.UsedRange.Rows.count + 2 To first_row Step -1
        If Not IsEmpty(VB_RFP_REPORT.Cells(LastRFPRow, mark_col)) Then
            Exit For
        End If
    Next LastRFPRow
    
    LastRFPRow = LastRFPRow + 1
End Function

Public Function Form_col_num(work_sheet As Worksheet, ByVal title As String)
'Function: Form_col_num - Retrieves the column number for the specified form heading: If form heading
'                         doesn't exist, returns -1.
'Arguments: col_title: String containing the col_title.
'Returns: Integer containing the column number of the col_title, or -1.


    Form_col_num = -1

    If title = "Surplus" Then
        Form_col_num = Form_col_num(work_sheet, "Total Qty.") + 1
    ElseIf title = "Inventory" Then
        Form_col_num = Form_col_num(work_sheet, "Total Qty.") + 3
    ElseIf title = "Existing" Then
        Form_col_num = Form_col_num(work_sheet, "Total Qty.") + 3
    Else
        Dim col As Integer
        For col = 2 To work_sheet.UsedRange.Columns.count
            If title = "Previous" Then
                If work_sheet.Cells(title_row, col).Value2 Like "Previous* RFP(s)" Then
                    Form_col_num = col
                    Exit Function
                End If
            Else
                If work_sheet.Cells(title_row, col).MergeCells Then
                    If CStr(work_sheet.Cells(title_row, col).MergeArea.Cells(1, 1).Value2) = title Then
                        Form_col_num = col
                        Exit Function
                    End If
                Else
                    If CStr(work_sheet.Cells(title_row, col).Value2) = title Then
                        Form_col_num = col
                        Exit Function
                    End If
                End If
            End If
        Next col
    End If
End Function

Private Sub EvalPageSetup()
'Subroutine: EvalPageSetup - Procedure to evaluate the Report page size/layout and determine
'                            the best configuration for printing. Called during RFP Report Issue


    Dim ui_change As Boolean
    ui_change = RenderUI(False)

    'Unabsorb all orders
    ' All orders should start out visibile (not absorbed in the "Previous" column)
    Do While get_num_absorbed(VB_RFP_REPORT) > 0
        UnabsorbOrder VB_RFP_REPORT
    Loop
    
    VB_RFP_REPORT.PageSetup.TopMargin = Application.InchesToPoints(0.5)
    VB_RFP_REPORT.PageSetup.BottomMargin = Application.InchesToPoints(0.5)
    VB_RFP_REPORT.PageSetup.LeftMargin = Application.InchesToPoints(0.5)
    VB_RFP_REPORT.PageSetup.RightMargin = Application.InchesToPoints(0.5)
    VB_RFP_REPORT.PageSetup.HeaderMargin = Application.InchesToPoints(0.25)
    VB_RFP_REPORT.PageSetup.FooterMargin = Application.InchesToPoints(0.25)
    
    'Initial Defualt: Portrait 8.5x11
    VB_RFP_REPORT.PageSetup.FitToPagesWide = False
    VB_RFP_REPORT.PageSetup.FitToPagesTall = False
    VB_RFP_REPORT.PageSetup.Zoom = 100
    VB_RFP_REPORT.PageSetup.PaperSize = xlPaperLetter
    VB_RFP_REPORT.PageSetup.Orientation = xlPortrait
    UpdatePageBreaksOnSheet VB_RFP_REPORT
    
    Dim successful_absorb As Boolean
    
    'Absorb orders until no more page break or till maximum is reached
    Do While VB_RFP_REPORT.VPageBreaks.count > 0 And NumOrdersShown() > min_orders_visible
        AbsorbOrder successful_absorb
        UpdatePageBreaksOnSheet VB_RFP_REPORT
        
        If Not successful_absorb Then Exit Do
    Loop
    
    'If page breaks still exist, before changed page size try to zoom out to fix.
    If VB_RFP_REPORT.VPageBreaks.count > 0 Then
        With VB_RFP_REPORT.PageSetup
            '.FitToPagesWide = 1
            Do While VB_RFP_REPORT.VPageBreaks.count > 0 And .Zoom > 10
                .Zoom = .Zoom - 1
                UpdatePageBreaksOnSheet VB_RFP_REPORT
            Loop
        End With
        
        '75% Zoom is minimum for readability
        If VB_RFP_REPORT.PageSetup.Zoom < 75 Then
            VB_RFP_REPORT.PageSetup.Zoom = 75
            UpdatePageBreaksOnSheet VB_RFP_REPORT
            
            ' if still has vertical page breaks, ...
            If VB_RFP_REPORT.VPageBreaks.count > 0 Then
                'Reset absorption
                Do While get_num_absorbed(VB_RFP_REPORT) > 0
                    UnabsorbOrder VB_RFP_REPORT
                Loop
                
                'Change page size to landscape and try to fit data
                VB_RFP_REPORT.PageSetup.Orientation = xlLandscape
                VB_RFP_REPORT.PageSetup.Zoom = 100
                UpdatePageBreaksOnSheet VB_RFP_REPORT
                
                Do While VB_RFP_REPORT.VPageBreaks.count > 0 And NumOrdersShown() > min_orders_visible
                    AbsorbOrder successful_absorb
                    UpdatePageBreaksOnSheet VB_RFP_REPORT
                    
                    If Not successful_absorb Then Exit Do
                Loop
                
                'If page breaks still exist, zoom out to fix.
                If VB_RFP_REPORT.VPageBreaks.count > 0 Then
                    With VB_RFP_REPORT.PageSetup
                        '.FitToPagesWide = 1
                        Do While VB_RFP_REPORT.VPageBreaks.count > 0 And .Zoom > 0
                            .Zoom = .Zoom - 1
                            UpdatePageBreaksOnSheet VB_RFP_REPORT
                        Loop
                    End With
                    
                    '63% Zoom is minimum for readability
                    If VB_RFP_REPORT.PageSetup.Zoom < 63 Then
                        VB_RFP_REPORT.PageSetup.Zoom = 63
                        UpdatePageBreaksOnSheet VB_RFP_REPORT
                        
                        ' if still has vertical page breaks, ...
                        If VB_RFP_REPORT.VPageBreaks.count > 0 Then
                            'Reset absorption
                            Do While get_num_absorbed(VB_RFP_REPORT) > 0
                                UnabsorbOrder VB_RFP_REPORT
                            Loop
                            
                            'Change page size to 11x17 and try to fit data
                            VB_RFP_REPORT.PageSetup.Orientation = xlLandscape
                            VB_RFP_REPORT.PageSetup.Zoom = 100
                            On Error GoTo HideWithCurrentPageSize
                            VB_RFP_REPORT.PageSetup.PaperSize = xlPaper11x17
                            Err.Clear
                            On Error GoTo 0
                            UpdatePageBreaksOnSheet VB_RFP_REPORT
                            
                            Do While VB_RFP_REPORT.VPageBreaks.count > 0 And NumOrdersShown() > min_orders_visible
                                AbsorbOrder successful_absorb
                                UpdatePageBreaksOnSheet VB_RFP_REPORT
                                
                                If Not successful_absorb Then Exit Do
                            Loop
                            
HideWithCurrentPageSize:
                            'If page breaks still exist, zoom out to fix.
                            If VB_RFP_REPORT.VPageBreaks.count > 0 Then
                                With VB_RFP_REPORT.PageSetup
                                    '.FitToPagesWide = 1
                                    Do While VB_RFP_REPORT.VPageBreaks.count > 0 And .Zoom > 0
                                        .Zoom = .Zoom - 1
                                        UpdatePageBreaksOnSheet VB_RFP_REPORT
                                    Loop
                                End With
                                
                                '50% Zoom is minimum for readability
                                If VB_RFP_REPORT.PageSetup.Zoom < 50 Then
                                    VB_RFP_REPORT.PageSetup.Zoom = 50
                                    UpdatePageBreaksOnSheet VB_RFP_REPORT
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    If ui_change Then RenderUI False
    
End Sub

Private Function get_num_absorbed(work_sheet As Worksheet) As Integer
'Function: get_num_absorbed - Calculates the number orders that have been absorbed in the "Previous" column
'Returns: Integer containing the number of absorbed orders


    Dim prev_col As Integer
    prev_col = Form_col_num(work_sheet, "Previous")
    
    get_num_absorbed = CInt(Left(Right(work_sheet.Cells(title_row, prev_col).Value2, Len(work_sheet.Cells(title_row, prev_col).Value2) - 9), _
        Len(Right(work_sheet.Cells(title_row, prev_col).Value2, Len(work_sheet.Cells(title_row, prev_col).Value2) - 9)) - 7))
End Function

Private Sub AbsorbOrder(Optional ByRef success As Boolean)
'Function: AbsorbOrder - Absorbs the last order still visible on the RFP Report (First In - Last Out)

    'Checks minimum
    If NumOrdersShown() = min_orders_visible Then
        success = False
        Exit Sub
    End If

    Dim prev_col As Integer
    prev_col = Form_col_num(VB_RFP_REPORT, "Previous")
    
    Dim total_absorbed As Integer
    total_absorbed = get_num_absorbed(VB_RFP_REPORT)
    
    If VB_RFP_REPORT.Cells(title_row, prev_col - (total_absorbed + 1)).Interior.Color = RGB(141, 180, 226) Then
        ' RFP is new - don't absorb
        success = False
        Exit Sub
    End If
    
    Dim last_row As Integer
    last_row = LastRFPRow()
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    'Absorb order Quantities
    Dim row As Integer
    For row = first_row To last_row - 1
        If Not IsEmpty(VB_RFP_REPORT.Cells(row, mark_col)) Then
            VB_RFP_REPORT.Cells(row, prev_col).Value = CInt(VB_RFP_REPORT.Cells(row, prev_col).Value2) + CInt(VB_RFP_REPORT.Cells(row, prev_col - (total_absorbed + 1)).Value2)
        End If
    Next row
    
    total_absorbed = total_absorbed + 1
    
    'Set # absorbed
    VB_RFP_REPORT.Cells(title_row, prev_col).Value = "Previous " & total_absorbed & " RFP(s)"
    If total_absorbed = 0 Then
        VB_RFP_REPORT.Cells(title_row + 1, prev_col).ClearContents
    Else
        VB_RFP_REPORT.Cells(title_row + 1, prev_col).Value = VB_RFP_REPORT.Cells(title_row + 1, prev_col - total_absorbed).Value2
    End If
    
    'Hide order col
    VB_RFP_REPORT.Columns(prev_col - total_absorbed).Hidden = True
    VB_RFP_REPORT.Columns(prev_col).Hidden = (total_absorbed = 0)
    success = True
    
    If ui_change Then RenderUI True
End Sub

Private Sub UnabsorbOrder(work_sheet As Worksheet)
'Function: UnabsorbOrder - Undoes absorption of last absorbed order (First In - Last Out)


    Dim total_absorbed As Integer
    total_absorbed = get_num_absorbed(work_sheet)
    
    'if no orders are absorbed, exit
    If total_absorbed = 0 Then
        Exit Sub
    End If

    Dim prev_col As Integer
    prev_col = Form_col_num(work_sheet, "Previous")
    
    'Unhide Previous order col
    work_sheet.Columns(prev_col - total_absorbed).Hidden = False

    Dim last_row As Integer
    last_row = LastRFPRow()

    Dim ui_change As Boolean
    ui_change = RenderUI(False)

    'Subtract order quatities from Previous
    Dim row As Integer
    For row = first_row To last_row - 1
        If Not IsEmpty(work_sheet.Cells(row, mark_col)) Then
            work_sheet.Cells(row, prev_col).Value = CInt(work_sheet.Cells(row, prev_col).Value2) - CInt(work_sheet.Cells(row, prev_col - total_absorbed).Value2)
        End If
    Next row
    
    total_absorbed = total_absorbed - 1
    
    'Set # absorbed
    work_sheet.Cells(title_row, prev_col).Value = "Previous " & total_absorbed & " RFP(s)"
    work_sheet.Cells(title_row + 1, prev_col).Value = work_sheet.Cells(title_row + 1, prev_col - total_absorbed).Value2
    
    'Hide column
    work_sheet.Columns(prev_col).Hidden = (total_absorbed = 0)
    
    If ui_change Then RenderUI True
End Sub

Private Sub ReportOrder(ByVal order_num As String)
'Subroutine: ReportOrder - Reports order to RFP Report. Is only called during Prepare4Publish.
'Arguments: order_num - String containing order number to report.


    Dim order_col As Integer
    order_col = VB_ORDER_LOG.get_order_col(order_num)
    Dim ord_date As String
    ord_date = VB_ORDER_LOG.get_order_date(order_num)
    
    Dim new_col As Integer
    new_col = Form_col_num(VB_RFP_REPORT, "Inventory") + 1
    
    Dim last_row As Integer
    last_row = LastRFPRow()
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    'unabsorb all orders
    Do While get_num_absorbed(VB_RFP_REPORT) > 0
        UnabsorbOrder VB_RFP_REPORT
    Loop
    
    'if 0 orders in RFP report, don't insert column - use empty column for first order
    If NumOrdersShown() > 0 Then
        VB_RFP_REPORT.Columns(new_col).Insert
        
        With VB_RFP_REPORT.Range("$" & ColLet(new_col) & "$" & title_row & ":$" & _
            ColLet(new_col) & "$" & last_row).Borders(xlEdgeRight)
            
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With VB_RFP_REPORT.Cells(title_row, new_col).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        VB_RFP_REPORT.Cells(title_row + 1, new_col).Interior.Color = RGB(197, 217, 241)
    End If
    
    'set title/date
    VB_RFP_REPORT.Cells(title_row, new_col).Value = order_num
    VB_RFP_REPORT.Cells(title_row + 1, new_col).Value = ord_date
    VB_RFP_REPORT.Columns(new_col).AutoFit
    If VB_RFP_REPORT.Columns(new_col).ColumnWidth < 8.43 Then
        VB_RFP_REPORT.Columns(new_col).ColumnWidth = 8.43
    End If
    
    Dim temp As Integer
    Dim RFProw As Integer
    Dim site_col As Integer
    Dim item_order As Integer
    Dim first As Integer
    Dim last As Integer
    first = VB_MASTER.FirstRow()
    last = VB_MASTER.LastRow()
    Dim row As Integer
    
    'Set Order Quantities
    For row = first To last
        If CInt(VB_MASTER.Cells(row, order_col).Value2) > 0 Then
            item_order = CInt(VB_MASTER.Cells(row, order_col).Value2)
            
            RFProw = get_RFProw(VB_RFP_REPORT, get_mark_num(row))
            If RFProw = -1 Then
                'add item if it doesn't already exist in RFP report
                RFProw = AddItem(get_mark_num(row))
            End If
            
            VB_RFP_REPORT.Cells(RFProw, new_col).Value = item_order
        End If
    Next row
    
    If ui_change Then RenderUI True
End Sub

Private Sub DeleteOrder(Optional ByVal order_num As String)
'Subroutine: DeleteOrder - Deletes order from RFP Report. Is only called after RFP Report is issued.
'                          If order_num is not specfied, deletes the oldest order.
'Arguments: order_num - String containing order number to report.


    Dim num_orders As Integer
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    'unabsorb all orders
    Do While get_num_absorbed(VB_RFP_REPORT) > 0
        UnabsorbOrder VB_RFP_REPORT
    Loop
    
    num_orders = NumOrdersShown()
    
    Dim last_row As Integer
    last_row = LastRFPRow()
    
    Dim order_col As Integer
    
    'if no orders, exit; if 1 order, clear (don't delete) column; else, delete column.
    If num_orders = 0 Then
        GoTo ResetApp
    ElseIf num_orders = 1 Then
        order_col = Form_col_num(VB_RFP_REPORT, "Inventory") + 1
        
        If CStr(VB_RFP_REPORT.Cells(title_row, order_col).Value2) = order_num Or Len(order_num) = 0 Then
            VB_RFP_REPORT.Range("$" & ColLet(order_col) & "$" & title_row & ":$" & _
                ColLet(order_col) & "$" & last_row).ClearContents
                
            VB_RFP_REPORT.Cells(title_row, order_col).Value = "N/A"
            
            VB_RFP_REPORT.Cells(title_row, order_col).Interior.Color = RGB(141, 180, 226)
            VB_RFP_REPORT.Cells(title_row + 1, order_col).Interior.Color = RGB(197, 217, 241)
            VB_RFP_REPORT.Range("$" & ColLet(order_col) & "$" & first_row & ":$" & _
                ColLet(order_col) & "$" & last_row).Interior.Color = RGB(255, 255, 255)
        End If
    Else
        order_col = Form_col_num(VB_RFP_REPORT, "Inventory") + 1
        Dim o As Integer
        For o = 0 To num_orders - 1
            If CStr(VB_RFP_REPORT.Cells(title_row, order_col + o).Value2) = order_num Then
                Exit For
            End If
        Next o
        
        If Len(order_num) = 0 Then
            order_col = order_col + o - 1
        ElseIf o < num_orders Then
            order_col = order_col + o
        Else
            GoTo ResetApp
        End If
        
        'delete column
        VB_RFP_REPORT.Columns(order_col).Delete
    End If
    
ResetApp:
    If ui_change Then RenderUI True
    
End Sub

Public Sub DumpSites()
'Subroutine: DumpSites - Dumps site items and quantities into RFP report. Dumps those items that have non zero model or order quantities.
'                        If item does not have a quantity in a site, it is not copied here; it just stays in MasterBOM.


    Dim RFProw As Integer
    
    Dim i As Integer
    Dim j As Integer
    Dim first As Integer
    Dim last As Integer
    first = VB_MASTER.FirstRow()
    last = VB_MASTER.LastRow()
    Dim start As Integer
    Dim finish As Integer
    start = GetSitesRange().Cells(1, 1).column
    finish = start + GetSitesRange().count - 1
    
    Dim hideZeros As Boolean
    HideZeros_GetPressed Nothing, hideZeros
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    For i = first To last
        SetStatusBar "Reporting Material and Site Quantities", (i - first), (last - first)
        
        For j = start To finish
            If (Not hideZeros Or CInt(VB_MASTER.Cells(i, j).Value2) > 0) And Not IsFabPackage(VB_MASTER.Cells(VB_MASTER.SubtitleRow(), j)) Then 'material quantity > 0
                Exit For
            End If
        Next j
        
        If j <= finish Then
            'Nonzero site quantity found, add item
            RFProw = get_RFProw(VB_RFP_REPORT, get_mark_num(i))
            If RFProw = -1 Then
                RFProw = AddItem(get_mark_num(i))
            End If
        End If
    Next i
    ResetStatusBar
    
    If ui_change Then RenderUI True
End Sub

Private Function AddItem(ByVal in_mark As Integer) As Integer
'Function: AddItem - Adds material item in order of description (row #) and inputs item's site quantities to RFP Report
'Arguments: in_mark - Integer containing the mark number for the item to add
'Returns: Integer containing the row number in RFP report where the item was added


    Dim row As Integer
    row = first_row
    
    If Len(get_property(in_mark, "Long Description")) = 0 Then
        AddItem = 0
        Exit Function
    End If
    
    'iterate through RFP report, find the spot for the new add item based on VB_MASTER row #
    Do While Not IsEmpty(VB_RFP_REPORT.Cells(row, mark_col))
        If get_row(VB_RFP_REPORT.Cells(row, mark_col)) > get_row(in_mark) Then
            Exit Do
        End If
        row = row + 1
    Loop
    AddItem = row
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    'insert/format row
    VB_RFP_REPORT.Rows(row).Insert CopyOrigin:=xlFormatFromRightOrBelow
    With VB_RFP_REPORT.Range("$" & ColLet(mark_col) & "$" & row & ":$" & ColLet(Form_col_num(VB_RFP_REPORT, "Surplus")) & "$" & row)
        .VerticalAlignment = xlTop
        
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    End With
    With VB_RFP_REPORT.Range("$" & ColLet(Form_col_num(VB_RFP_REPORT, "Inventory")) & "$" & row & ":$" & _
        ColLet(Form_col_num(VB_RFP_REPORT, "Previous")) & "$" & row)
        
        .VerticalAlignment = xlTop
        
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    End With
    With VB_RFP_REPORT.Range("$" & ColLet(mark_col) & "$" & row & ":$" & ColLet(Form_col_num(VB_RFP_REPORT, "Inventory")) & "$" & row)
        .Interior.Color = RGB(255, 255, 255)
    End With
    
    ' copy mark number to Report
    VB_RFP_REPORT.Cells(row, mark_col).Value = in_mark
    
    ' copy SAP# to Report
    VB_RFP_REPORT.Cells(row, sap_col).Value = get_property(in_mark, "SAP#")
    
    Dim mst_desc_col As Integer
    Dim mst_row As Integer
    mst_desc_col = get_col_num("Long Description")
    mst_row = get_row(in_mark)
    
    Dim removed As Boolean
    removed = False
    
    'Set Description
    Dim i As Integer
    With VB_RFP_REPORT.Cells(row, desc_col)
        .Value = get_property(in_mark, "Long Description")
        
        For i = 1 To VB_MASTER.Cells(mst_row, mst_desc_col).Characters.count
            .Characters(i, 1).Font.ColorIndex = VB_MASTER.Cells(mst_row, mst_desc_col).Characters(i, 1).Font.ColorIndex
            '.Characters(i, 1).Font.Strikethrough = VB_MASTER.Cells(mst_row, mst_desc_col).Characters(i, 1).Font.Strikethrough
        Next i
    End With
    
    'Set Unit
    VB_RFP_REPORT.Cells(row, unit_col).Value = get_property(in_mark, "Unit")
    
    'Define Site Quantities
    Dim rSites As Range
    Dim vbCell
    Set rSites = GetSitesRange()
    Dim rfp_site_col As Integer
    
    For Each vbCell In rSites
        rfp_site_col = Form_col_num(VB_RFP_REPORT, vbCell.Value2)
        
        'if site is a fab_package, rfp_site_col will be -1.. no need to handle this further
        If rfp_site_col > 0 Then
            VB_RFP_REPORT.Cells(row, rfp_site_col).Value = vbCell.offset(get_row(in_mark) - VB_MASTER.SubtitleRow()).Value2
            VB_RFP_REPORT.Cells(row, rfp_site_col).Font.Strikethrough = vbCell.offset(get_row(in_mark) - VB_MASTER.SubtitleRow()).Font.Strikethrough
            
            If VB_RFP_REPORT.Cells(row, rfp_site_col).Font.Strikethrough Then
                removed = True
            End If
        End If
    Next vbCell
    
    'Define Consumables
    VB_RFP_REPORT.Cells(row, Form_col_num(VB_RFP_REPORT, "Consumables")).Value = get_property(in_mark, "Total Extras")
    
    'Define Total Qty.
    If removed Then
        VB_RFP_REPORT.Cells(row, Form_col_num(VB_RFP_REPORT, "Total Qty.")).Value = "0"
    Else
        VB_RFP_REPORT.Cells(row, Form_col_num(VB_RFP_REPORT, "Total Qty.")).Value = CInt(get_property(in_mark, "Model Total")) + CInt(get_property(in_mark, "Total Extras"))
    End If
    
    'Define Total RFP
    VB_RFP_REPORT.Cells(row, Form_col_num(VB_RFP_REPORT, "Inventory")).Value = CInt(get_property(in_mark, "Client Inventory")) + CInt(get_property(in_mark, "Total Ordered"))
    
    'Define Surplus/Deficit
    VB_RFP_REPORT.Cells(row, Form_col_num(VB_RFP_REPORT, "Surplus")).Formula = CInt(get_property(in_mark, "Quantity Needed"))
    
    VB_RFP_REPORT.Rows(row).AutoFit
    
    If ui_change Then RenderUI True
    
End Function

Private Sub RemoveItem(ByVal in_mark As Integer)
'Subroutine: RemoveItem - Removes item from RFP Report. This is only used afer RFP Report is issued to empty report.
'Arguments: in_mark - Mark number of item to remove from RFP Report


    Dim row As Integer
    row = first_row
    
    'get row number for item
    row = get_RFProw(VB_RFP_REPORT, in_mark)
    If row = -1 Then
        Exit Sub
    End If
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    'remove
    VB_RFP_REPORT.Rows(row).Delete
    
    If ui_change Then RenderUI True
End Sub

Private Function get_RFProw(work_sheet As Worksheet, ByVal in_mark As Integer) As Integer
'Function: get_RFProw - Gets the row number in work_sheet for item with mark number in_mark
'Arguments: work_sheet - During Prepare4Publish, two RFP Reports may be open at once, this function
'                        can apply to either by giving the worksheet object
'           in_mark - the mark number for the item to search for
'Returns: Integer containing the row number


    Dim last_row As Integer
    last_row = LastRFPRow()
    
    Dim i As Integer
    For i = first_row To last_row - 1
        If CInt(work_sheet.Cells(i, mark_col).Value2) = in_mark Then
            get_RFProw = i
            Exit Function
        End If
    Next i
    
    'unmatched mark number
    get_RFProw = -1
End Function

Private Function get_RFProw_DESC(work_sheet As Worksheet, ByVal DESC As String) As Integer
'Function: get_RFProw_DESC - Gets the row number in work_sheet for item with description DESC
'Arguments: work_sheet - During Prepare4Publish, two RFP Reports may be open at once, this function
'                        can apply to either by giving the worksheet object
'           DESC - the item description to search for
'Returns: Integer containing the row number


    Dim last_row As Integer
    last_row = LastRFPRow()
    
    Dim i As Integer
    For i = first_row To last_row - 1
        If DescriptionCompare(work_sheet.Cells(i, desc_col).Value2, DESC) = 0 Then
            get_RFProw_DESC = i
            Exit Function
        End If
    Next i
    
    'unmatched mark number
    get_RFProw_DESC = -1
End Function

Public Sub AddSite(ByVal site_name As String)
'Subroutine: AddSite - Adds site to RFP Report. This sub works parallel with AddSite to VB_MASTER
'Arguments: site_name - the site name to add


    If SiteIndex(site_name) = -1 Then
        Exit Sub
    End If
    
    Dim insert_index As Integer
    Dim insert_col As Integer
    insert_index = SiteIndex(site_name)
    
    Dim last_row As Integer
    last_row = LastRFPRow()
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    'if NumSitesRFP = 0, don't add column
    If NumSitesRFP(VB_RFP_REPORT) > 0 Then
        For insert_col = site1_col To Form_col_num(VB_RFP_REPORT, "Consumables") - 1
            If SiteIndex(VB_RFP_REPORT.Cells(title_row, insert_col).Value2) > insert_index Then
                Exit For
            End If
        Next insert_col
    
        VB_RFP_REPORT.Columns(insert_col).EntireColumn.Insert CopyOrigin:=xlFormatFromLeftOrAbove
        VB_RFP_REPORT.Cells(title_row, insert_col).Orientation = 90
        
        If insert_col = site1_col Then
            With VB_RFP_REPORT.Range("$" & ColLet(insert_col) & "$" & title_row & ":$" & _
                ColLet(insert_col) & "$" & last_row).Borders(xlEdgeLeft)
                
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
            With VB_RFP_REPORT.Range("$" & ColLet(insert_col) & "$" & title_row & ":$" & _
                ColLet(insert_col) & "$" & last_row).Borders(xlEdgeRight)
                
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        Else
            With VB_RFP_REPORT.Range("$" & ColLet(insert_col) & "$" & title_row & ":$" & _
                ColLet(insert_col) & "$" & last_row).Borders(xlEdgeLeft)
                
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With VB_RFP_REPORT.Range("$" & ColLet(insert_col) & "$" & title_row & ":$" & _
                ColLet(insert_col) & "$" & last_row).Borders(xlEdgeRight)
                
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        End If
        VB_RFP_REPORT.Range("$" & ColLet(insert_col) & "$" & title_row & ":$" & _
            ColLet(insert_col) & "$" & title_row + 1).Merge
        
        With VB_RFP_REPORT.Range("$" & ColLet(Form_col_num(VB_RFP_REPORT, "Consumables")) & "$" & title_row & ":$" & _
            ColLet(Form_col_num(VB_RFP_REPORT, "Consumables")) & "$" & last_row).Borders(xlEdgeLeft)
            
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
    Else
        'NumSitesRFP = 0
        insert_col = site1_col
    End If
    
    'Set site_name in title block
    VB_RFP_REPORT.Cells(title_row, insert_col).Value = site_name
    
    If ui_change Then RenderUI True
End Sub

Public Sub RenameSite(ByVal old_site_name As String, ByVal site_name As String)
'Subroutine: RenameSite - Renames site in congruence with RenameSite in VB_MASTER
'Arguments: old_site_name - name of site to change
'           site_name - the new site name


    Dim site_col As Integer
    site_col = Form_col_num(VB_RFP_REPORT, old_site_name)
    
    'change site name if site_col is valid
    If site_col <> -1 Then
        Dim ui_change As Boolean
        ui_change = RenderUI(False)
        
        VB_RFP_REPORT.Cells(title_row, site_col).Value = site_name
        
        If ui_change Then RenderUI True
    End If

End Sub

Public Sub RemoveSite(ByVal site_name As String)
'Subroutine: RemoveSite - Removes site in congruence with RemoveSite in VB_MASTER
'Arguments: site_name - name of site to remove


    Dim site_col As Integer
    site_col = Form_col_num(VB_RFP_REPORT, site_name)
    
    If site_col = -1 Then
        'site doesn't exist, exit
        Exit Sub
    End If
    
    Dim last_row As Integer
    last_row = LastRFPRow()
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    'if only one site, don't delete column
    If NumSitesRFP(VB_RFP_REPORT) = 1 Then ' only one site
        VB_RFP_REPORT.Cells(title_row, site_col).MergeArea.ClearContents
    Else
        VB_RFP_REPORT.Columns(site_col).EntireColumn.Delete
        
        With VB_RFP_REPORT.Range("$" & ColLet(site_col) & "$" & title_row & ":$" & _
            ColLet(site_col) & "$" & last_row).Borders(xlEdgeLeft)
            
            If site_col = site1_col Then
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End If
        End With
    End If
    
    If ui_change Then RenderUI True
End Sub

Private Function NumSitesRFP(work_sheet As Worksheet) As Integer
'Function: NumSitesRFP - Returns number of sites in RFP Report.
'Arguments: work_sheet - During Prepare4Publish, two RFP Reports may be open at once, this function
'                        can apply to either by giving the worksheet object
'Returns: Integer containing the number sites reported in the RFP report.


    Dim total_col As Integer
    total_col = Form_col_num(work_sheet, "Consumables")
    
    NumSitesRFP = total_col - site1_col
    If NumSitesRFP = 1 Then
        If IsEmpty(work_sheet.Cells(title_row, total_col - 1)) Then
            NumSitesRFP = 0
        End If
    End If
End Function

Private Function NumOrdersShown() As Integer
'Function: NumOrdersShown - Returns number of orders explicitly shown (unhidden/unabsorbed) in RFP Report.
'Returns: Integer containing the number sites reported in the RFP report.


    NumOrdersShown = 0
    Dim exst_col As Integer
    Dim prev_col As Integer
    exst_col = Form_col_num(VB_RFP_REPORT, "Inventory")
    prev_col = Form_col_num(VB_RFP_REPORT, "Previous")
    Dim i As Integer
    
    'loop columns and count ones that are not hidden
    For i = 1 To prev_col - exst_col - 1
        If Not VB_RFP_REPORT.Columns(exst_col + i).Hidden Then
            NumOrdersShown = NumOrdersShown + 1
            'If "N/A", NumOrders = 0, don't count
            If VB_RFP_REPORT.Cells(title_row, exst_col + i).Value2 = "N/A" Then
                NumOrdersShown = NumOrdersShown - 1
            End If
        End If
    Next i
End Function

Public Sub Prepare4Publish()
'Subroutine: Prepare4Publish - Prepares RFP Report to issuing
'                               1 - Dumps all sites/items into Report
'                               2 - Reports Project orders
'                               3 - If 'RFP Report Current' exists (if rev > 0), compare with previous version.
'                               4 - Saves New RFP Report as current version after deleted 'RFP Report Current'
'                               5 - Evaluate Page Setup


    Dim ui_change As Boolean
    ui_change = RenderUI(False)

    'Add all valid material items to report.
    DumpSites
    
    SetStatusBar "Reporting RFPs..."
    
    'Report orders
    Dim o As Integer
    For o = 0 To VB_ORDER_LOG.NumOrders() - 1
        If VB_ORDER_LOG.OrderIsReceived(o) Then
            ReportOrder VB_ORDER_LOG.OrderNumber(o)
        End If
    Next o
    
    'Update Header
    VB_RFP_REPORT.Cells(2, 5).Value = VB_COVERSHEET.ProjectClientName
    VB_RFP_REPORT.Cells(3, 5).Value = VB_COVERSHEET.ProjectName
    VB_RFP_REPORT.Cells(4, 5).Value = Format(Now(), "yyyy-mm-dd - HH:NN:SS AMPM")
    VB_RFP_REPORT.Cells(5, 5).Value = VB_COVERSHEET.HMMContact

    'COMPARE WITH RFP REPORT CURRENT
    Dim past_rev As Worksheet
    
    If SheetExists("RFP Report Current") Then
        Set past_rev = ThisWorkbook.Sheets("RFP Report Current")
    Else
        Set past_rev = Nothing
        GoTo EndComparison
    End If
    
    SetStatusBar "Comparing to Previous Issue..."
    
    Dim consum_col As Integer
    Dim surp_col As Integer
    Dim tot_qty_col As Integer
    Dim tot_rfp_col As Integer
    Dim last_col As Integer
    
    consum_col = Form_col_num(VB_RFP_REPORT, "Consumables")
    surp_col = Form_col_num(VB_RFP_REPORT, "Surplus")
    tot_qty_col = Form_col_num(VB_RFP_REPORT, "Total Qty.")
    tot_rfp_col = Form_col_num(VB_RFP_REPORT, "Inventory")
    last_col = Form_col_num(VB_RFP_REPORT, "Previous") - 1
    
    Dim mark_num As Integer
    
    Dim prev_consum_col As Integer
    Dim prev_surp_col As Integer
    Dim prev_tot_qty_col As Integer
    Dim prev_tot_rfp_col As Integer
    Dim prev_last_col As Integer
    
    prev_consum_col = Form_col_num(past_rev, "Consumables")
    prev_surp_col = Form_col_num(past_rev, "Surplus")
    prev_tot_qty_col = Form_col_num(past_rev, "Total Qty.")
    prev_tot_rfp_col = Form_col_num(past_rev, "Inventory")
    prev_last_col = Form_col_num(past_rev, "Previous") - 1
    
    Dim row As Integer
    Dim prev_row As Integer
    
    Dim prev_col As Integer
    Dim col As Integer
    
    Dim last_row As Integer
    last_row = LastRFPRow()
    
    'compare supply side - orders
    For prev_col = prev_tot_rfp_col + 1 To prev_last_col
        If past_rev.Cells(title_row, prev_col).Value2 = "N/A" Then
            GoTo CompareDemand
        Else
            For col = tot_rfp_col + 1 To last_col
                'MsgBox VB_RFP_REPORT.Cells(title_row, col).Value2 & " = " & past_rev.Cells(title_row, prev_col).Value2
                
                'go next when reach "Previous"
                'if order isn't new (contained in previous revision), grey out.
                'else, do nothing, keep it blue
                If VB_RFP_REPORT.Cells(title_row, col).Value2 = past_rev.Cells(title_row, prev_col).Value2 Then
                    ' order already exists
                    VB_RFP_REPORT.Cells(title_row, col).Interior.Color = RGB(166, 166, 166)
                    VB_RFP_REPORT.Cells(title_row + 1, col).Interior.Color = RGB(191, 191, 191)
                    VB_RFP_REPORT.Range("$" & ColLet(col) & "$" & first_row & ":$" & _
                        ColLet(col) & "$" & last_row).Interior.Color = RGB(217, 217, 217)
                        
                    GoTo NextPastRevCol
                Else
                    ' do nothing
                End If
            Next col
        End If
        
NextPastRevCol:
    Next prev_col
    
CompareDemand:
    Dim site_name As String
    Dim site_alias As String
    Dim site_num As Integer
    
    'compare demand side - mark num through site quantities
    For row = first_row To last_row
        SetStatusBar "Comparing to Previous Issue...", (row - first_row), (last_row - first_row)
        
        site_name = ""
        site_alias = ""
        site_name = 0
        
        If IsEmpty(VB_RFP_REPORT.Cells(row, desc_col)) Then
            GoTo EndComparison
        Else
            VB_RFP_REPORT.Rows(row).AutoFit
        End If
        
        'match description to one in Previous
        mark_num = VB_RFP_REPORT.Cells(row, mark_col).Value2
        prev_row = get_RFProw_DESC(past_rev, VB_RFP_REPORT.Cells(row, desc_col).Value2)
        
        If prev_row = -1 Then
            Call HighlightChange(row, desc_col)
            GoTo ContinueNext
        End If
        
        'Check for mark_num
        If mark_num <> past_rev.Cells(prev_row, mark_col).Value2 Then
            Call HighlightChange(row, mark_col)
        End If
        
        'Check for quantity changes
        
        site_num = 1
        For site_num = 1 To NumSitesRFP(past_rev)
            prev_col = desc_col + site_num
            site_name = past_rev.Cells(title_row, prev_col)
            site_alias = VB_SITEDB.GetClientSummaryAlias(site_name)
            
            'Sites are correlated to their alias in VB_SITEDB
            If Len(site_alias) > 0 Then
                col = Form_col_num(VB_RFP_REPORT, site_alias)
            
                If col <> -1 Then
                    If VB_RFP_REPORT.Cells(row, col).Value2 <> past_rev.Cells(prev_row, prev_col).Value2 Then
                        Call HighlightChange(row, col)
                    End If
                End If
            Else
                'new sites are not highlighted: do nothing here
            End If
        Next site_num
        
        'Check for 'Deleted' items
        For col = desc_col To desc_col + VBMin(NumSitesRFP(VB_RFP_REPORT), NumSitesRFP(past_rev))
            If VB_RFP_REPORT.Cells(row, col).Font.Strikethrough <> past_rev.Cells(prev_row, col).Font.Strikethrough _
                And Not IsEmpty(VB_RFP_REPORT.Cells(row, col)) Then
                
                Call HighlightChange(row, col)
            End If
        Next col
        
        'Check for Surplus change
        If VB_RFP_REPORT.Cells(row, surp_col).Value2 <> _
            past_rev.Cells(prev_row, prev_surp_col).Value2 Then
        
            Call HighlightChange(row, surp_col)
        End If
        
        'Check for Consumables change
        If VB_RFP_REPORT.Cells(row, consum_col).Value2 <> _
            past_rev.Cells(prev_row, prev_consum_col).Value2 Then
        
            Call HighlightChange(row, consum_col)
        End If
        
        'Check for Total Qty change
        If VB_RFP_REPORT.Cells(row, tot_qty_col).Value2 <> _
            past_rev.Cells(prev_row, prev_tot_qty_col).Value2 Then
        
            Call HighlightChange(row, tot_qty_col)
        End If
        
        'Check for Unit change
        If VB_RFP_REPORT.Cells(row, unit_col).Value2 <> _
            past_rev.Cells(prev_row, unit_col).Value2 Then
            
            Call HighlightChange(row, unit_col)
        End If
        
        'Check for Total RFP change
        If VB_RFP_REPORT.Cells(row, tot_rfp_col).Value2 <> _
            past_rev.Cells(prev_row, prev_tot_rfp_col).Value2 Then
            
            Call HighlightChange(row, tot_rfp_col)
        End If
        
ContinueNext:
    Next row
    
EndComparison:
    VB_SITEDB.SetClientSummaryAliasToMatch
    
    If Not past_rev Is Nothing Then
        Do While get_num_absorbed(past_rev) > 0
            UnabsorbOrder past_rev
        Loop
        past_rev.Visible = xlSheetHidden
    End If
    
    VB_RFP_REPORT.Visible = xlSheetVeryHidden

    ResetStatusBar
    Set past_rev = Nothing

    'PAGE SETUP FOR PRINTING
    EvalPageSetup
    
    If ui_change Then RenderUI True
End Sub

Public Function PublishAndExport() As Integer
'Function: PublishAndExport - Publishes VB_RFP_REPORT as PDF, and exports the sheet to a separate workbook.
'                             Handles rev numbers by marking duplicates as "OBSOLETE_". Only one(1) obsolete
'                             back up is kept.
'Returns: Integer error code. 0: no error, -1: file folder doesn't exist.

    Dim past_rev As Worksheet
    
    'Delete past_rev sheet is it exists
    If SheetExists("RFP Report Current") Then
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets("RFP Report Current").Visible = xlSheetVisible
        ThisWorkbook.Sheets("RFP Report Current").Delete
        Application.DisplayAlerts = True
    End If
    
    SetStatusBar "Creating Client Summary Records..."

    'save new revision as the current RFP Report for next time.
    VB_RFP_REPORT.Visible = xlSheetVisible
    VB_RFP_REPORT.Activate
    VB_RFP_REPORT.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    Set past_rev = ThisWorkbook.ActiveSheet
    past_rev.Name = "RFP Report Current"
    Set past_rev = Nothing
    
    Dim rev As Integer
    rev = VB_VAR_STORE.GetClientSummaryRev()
    
    ' find if revision exists
    Dim filename As String
    Dim OldFileName As String
    Dim NewFileName As String
    filename = FindRFPReportRevision(rev)
    
    If filename <> vbNullString Then
        ' mark it obsolete
        OldFileName = ThisWorkbook.Path & VB_VAR_STORE.GetClientSummaryDirectory() & filename
        NewFileName = ThisWorkbook.Path & VB_VAR_STORE.GetClientSummaryDirectory() & "OBSOLETE_" & filename
        On Error Resume Next
        If Len(Dir(NewFileName & ".xlsx")) > 0 Then
            Kill NewFileName & ".xlsx"
        End If
        If Len(Dir(NewFileName & ".pdf")) > 0 Then
            Kill NewFileName & ".pdf"
        End If
        Name OldFileName & ".pdf" As NewFileName & ".pdf"
        Name OldFileName & ".xlsx" As NewFileName & ".xlsx"
        Err.Clear
        On Error GoTo 0
    End If
    
    filename = document_title & HMMFileTag() & "_rev" & rev
    
    'Copy sheet to new workbook
    Dim wb As Workbook
    Dim CurrentReport As Worksheet
    Set CurrentReport = ThisWorkbook.Worksheets("RFP Report Current")
    
    CurrentReport.Visible = xlSheetVisible
    CurrentReport.Copy
    Set wb = ActiveWorkbook
    wb.ActiveSheet.Name = VB_RFP_REPORT.ClientSummaryTitle()
    ActiveWindow.View = xlPageBreakPreview
    ActiveWindow.Zoom = 100
    
    'Save As
    Application.DisplayAlerts = False
    
    If Not FileFolderExists(ThisWorkbook.Path & VB_VAR_STORE.GetClientSummaryDirectory()) Then
        MakeDirs ThisWorkbook.Path & VB_VAR_STORE.GetClientSummaryDirectory()
        If Not FileFolderExists(ThisWorkbook.Path & VB_VAR_STORE.GetClientSummaryDirectory()) Then
            MsgBox "For some reason, this application cannot automatically create this folder directory:" & vbCrLf & _
                ThisWorkbook.Path & VB_VAR_STORE.GetClientSummaryDirectory() & vbCrLf & _
                "Please create this folder manually, try again. Sorry for the inconvenience.", vbCritical
                
            GoTo ErrExit
        End If
    End If
    
    On Error Resume Next
    wb.SaveAs ThisWorkbook.Path & VB_VAR_STORE.GetClientSummaryDirectory() & filename & ".xlsx"  'Save file
    
    'Export file as PDF
    Call Print2PDF(wb.ActiveSheet.Name, "", filename)
    
    wb.Close 'Close file
    Err.Clear
    On Error GoTo 0
    
    CurrentReport.Visible = xlSheetHidden
    Application.DisplayAlerts = True
    
    VB_CHANGE_LOG.LogChange "RFP CLIENT REPORTS", "", "Issue RFP Client Report", "", filename
    
    Set wb = Nothing
    Set CurrentReport = Nothing
    PublishAndExport = 0
    Exit Function
    
ErrExit:
    Set wb = Nothing
    Set CurrentReport = Nothing
    PublishAndExport = -1
End Function

Private Sub HighlightChange(ByVal row As Integer, Optional ByVal column As Integer = 0)
'Subroutine: HighlightChange - Highlights a cell orange and the rest of the row light orange
'                              to signify a change from the past revision.
'Arguments: row - the row number in RFP Report to highlight change
'           column - the column number in RFP Report to highlight change
'                    If not specified, all non-empty cells in row are highlighted

    
    If Not (row > 0 And column >= 0) Then
        Exit Sub
    End If
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    If column > 0 Then ' column is specified
        'If VB_RFP_REPORT.Cells(row, column).Interior.Color = RGB(255, 255, 255) Then
        '    'this is the first change on the row, highlight row light orange
        '
        '    With VB_RFP_REPORT.Range("$" & ColLet(mark_col) & "$" & row & ":$" & _
        '        ColLet(Form_col_num(VB_RFP_REPORT, "Surplus")) & "$" & row)
        '
        '        .Interior.Color = RGB(255, 255, 153)
        '    End With
        '
        '    Dim temp_col As Integer
        '    For temp_col = Form_col_num(VB_RFP_REPORT, "Inventory") To Form_col_num(VB_RFP_REPORT, "Previous")
        '        With VB_RFP_REPORT.Cells(row, temp_col)
        '            If .Interior.Color = RGB(255, 255, 255) Then
        '                .Interior.Color = RGB(255, 255, 153)
        '            End If
        '        End With
        '    Next temp_col
        'End If
        
        If column = 13 Then
            MsgBox "!"
        End If
        
        'highlight specified cell yellow
        With VB_RFP_REPORT.Cells(row, column)
            '.Interior.Color = RGB(255, 192, 0)
            .Interior.Color = RGB(255, 255, 153)
        End With
        
    Else ' column is not specified
        'Highlight all non empty cells in row
        For column = mark_col To Form_col_num(VB_RFP_REPORT, "Existing")
            If Not IsEmpty(VB_RFP_REPORT.Cells(row, column)) Then
                Call HighlightChange(row, column)
            End If
        Next column
    End If
    
    If ui_change Then RenderUI True
End Sub

Private Sub UnHighlightChanges(ws As Worksheet)
'Subroutine: UnHighlightChanges - Clears highlighted changes to the specified RFP Report
'Arguments: ws - During Prepare4Publish, two RFP Reports may be open at once, this function
'                can apply to either by giving the worksheet object


    Dim i As Integer
    Dim j As Integer
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    'loop through all used rows and columns and change orange and light orange cells to white
    For i = title_row To ws.UsedRange.Rows.count
        For j = mark_col To ws.UsedRange.Columns.count
            With ws.Cells(i, j).Interior
                
                If .Color = RGB(255, 192, 0) Or .Color = RGB(255, 255, 153) Or .Color = RGB(217, 217, 217) Then
                    .Color = RGB(255, 255, 255)
                    
                ElseIf .Color = RGB(166, 166, 166) Then
                    .Color = RGB(141, 180, 226)
                    
                ElseIf .Color = RGB(191, 191, 191) Then
                    .Color = RGB(197, 217, 241)
                    
                End If
                
            End With
        Next j
    Next i
    
    If ui_change Then RenderUI True
End Sub

Public Sub EmptyReport()
'Subroutine: EmptyReport - Empties RFP report after issue in preparation for next issue.
'                           1 - Removes all items from report
'                           2 - Removes all orders from report
'                           3 - Evaluates page setup


    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    UnHighlightChanges VB_RFP_REPORT
    
    'remove all items
    Do While Not IsEmpty(VB_RFP_REPORT.Cells(first_row, mark_col))
        RemoveItem VB_RFP_REPORT.Cells(first_row, mark_col).Value2
    Loop
    
    'unabsorb all orders
    Do While get_num_absorbed(VB_RFP_REPORT) > 0
        UnabsorbOrder VB_RFP_REPORT
    Loop
    
    'remove orders
    Do While NumOrdersShown() > 0
        DeleteOrder
    Loop
    
    'reset page setup
    EvalPageSetup
    
    If ui_change Then RenderUI True
    
End Sub

Public Function ClientSummaryTitle() As String
'Function: ClientSummaryTitle - Get method for document_title property
'Returns: String containing the document_title client summaries

    ClientSummaryTitle = document_title
End Function

Public Function TitleRow() As Integer
'Function: TitleRow - Get method for title_row property
'Returns: Integer containing the title row

    TitleRow = title_row
End Function


