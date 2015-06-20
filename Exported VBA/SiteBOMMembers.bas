Attribute VB_Name = "SiteBOMMembers"
Option Explicit

Public Sub CreateSiteSpecificBOM(ByVal site_name As String, Optional ByVal create_file As Boolean = True)
'Subroutine: CreateSiteSpecificBOM - Create Site Specific BOM for specified site. Creates/update
'                                    site BOM sheet and archives BOM in "\Site BOM Archive"
'Arguments: site_name - name of site to handle site BOMs.
'           create_file - Boolean. True - create the separate workbook.
'           False - do not make the separate workbook, just the worksheet.


    Dim site_quan As Integer
    If Not SiteExists(site_name, site_quan) Then
        ErrorHandling "CreateSiteSpecificBOM", 10, "Site: " & site_name & " doesn't exist.", 1
        Exit Sub
    End If
    
    'Copy BlankSiteBOM sheet; rename and title it - if the sheet doesn't already exists, if it does, empty it
    Dim site_sheet As Worksheet
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    Dim site_sheet_name As String
    If Len(site_name) > 25 Then
        site_sheet_name = Left(site_name, 25) & " - BOM"
    Else
        site_sheet_name = site_name & " - BOM"
    End If
    
    If SheetExists(site_sheet_name) Then
        Set site_sheet = ThisWorkbook.Sheets(site_sheet_name)
        'EmptySiteBOM site_name
        Application.DisplayAlerts = False
        site_sheet.Delete
        Application.DisplayAlerts = True
        Set site_sheet = Nothing
    End If
    'Else
    
    'Copy BlankSiteBOM sheet
    VB_SITEBOM.Visible = xlSheetVisible
    VB_SITEBOM.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    Set site_sheet = ThisWorkbook.ActiveSheet

    site_sheet.Name = site_sheet_name
    site_sheet.Cells(2, 1).MergeArea.Value = site_name
    'End If
    
    VB_SITEBOM.Visible = xlSheetVeryHidden
    site_sheet.Visible = xlSheetVisible
    
    'iterate Master BOM, find items that belong to site_name; copy line to site specific sheet
    Dim master_mark As Integer
    Dim master_row As Integer
    Dim last_mark As Integer
    Dim site_row As Integer
    Dim mark_col As Integer
    site_row = VB_SITEBOM.FirstRow()
    last_mark = MaxMark()
    mark_col = get_col_num("Mark No.")
    
    For master_mark = 1 To last_mark
        master_row = get_row(master_mark)
        
        If master_row >= VB_MASTER.FirstRow() Then
            If CInt(VB_MASTER.Cells(master_row, site_quan).Value2) > 0 And Not IsDeleted(master_mark) Then
                
                ' 1 - SAP Number
                ' 2 - Mark No
                ' 3 - Quantity
                ' 4 - Unit
                ' 5 - Item Description
                
                'insert in Site BOM
                site_sheet.Cells(site_row, 1).Value = get_property(master_mark, "SAP#")
                site_sheet.Cells(site_row, 2).Value = master_mark
                site_sheet.Cells(site_row, 3).Value = VB_MASTER.Cells(master_row, site_quan).Value2
                site_sheet.Cells(site_row, 4).Value = get_property(master_mark, "Unit")
                site_sheet.Cells(site_row, 5).Value = get_property(master_mark, "Long Description")
                'figure site "Quantity Needed"
                
                site_sheet.Rows(site_row).AutoFit
                
                site_row = site_row + 1
            End If
        End If
        
    Next master_mark
    
    If site_sheet.Columns(1).Hidden Then
        Dim row1 As String
        Dim row2 As String
        row1 = site_sheet.Cells(1, 1).Value2
        row2 = site_sheet.Cells(2, 1).Value2
        site_sheet.Columns(1).Delete
        site_sheet.Cells(1, 1).Value = row1
        site_sheet.Cells(2, 1).Value = row2
    End If
    
    Application.ScreenUpdating = True
    site_sheet.Cells(VB_SITEBOM.FirstRow(), 1).Select
    Application.ScreenUpdating = False
    
    If create_file Then
        'Copy to new workbook and archive
        Dim wb As Workbook
        Dim auxSheet As Worksheet
        Dim filename As String
        filename = site_name & " - BOM" & HMMFileTag() & "_rev" & VB_SITEDB.GetSiteBOMRev(site_name)
        
        site_sheet.Copy
        Set wb = ActiveWorkbook
        Set auxSheet = ActiveWorkbook.ActiveSheet
        auxSheet.Name = site_sheet_name
        If ActiveWindow.ActiveSheet Is auxSheet Then
            ActiveWindow.DisplayHeadings = True
        End If
        
        'If auxSheet.Columns(1).Hidden Then
        '    Dim row1 As String
        '    Dim row2 As String
        '    row1 = auxSheet.Cells(1, 1).Value2
        '    row2 = auxSheet.Cells(2, 1).Value2
        '    auxSheet.Columns(1).Delete
        '    auxSheet.Cells(1, 1).Value = row1
        '    auxSheet.Cells(2, 1).Value = row2
        'End If
        
        'SaveAs
        Application.DisplayAlerts = False
        
        If Not FileFolderExists(ThisWorkbook.Path & VB_VAR_STORE.GetSiteBOMDirectory()) Then
            MakeDirs ThisWorkbook.Path & VB_VAR_STORE.GetSiteBOMDirectory()
            If Not FileFolderExists(ThisWorkbook.Path & VB_VAR_STORE.GetSiteBOMDirectory()) Then
                MsgBox "For some reason, this application cannot automatically create this folder directory:" & vbCrLf & _
                    ThisWorkbook.Path & VB_VAR_STORE.GetSiteBOMDirectory() & vbCrLf & _
                    "Please create this folder manually, and try again. Sorry for the inconvenience.", vbCritical
                    
                GoTo ResetApplication
            End If
        End If
        
        On Error Resume Next
        wb.SaveAs ThisWorkbook.Path & VB_VAR_STORE.GetSiteBOMDirectory() & filename & ".xlsx"  'Save file
        wb.Close 'Close file
        Err.Clear
        On Error GoTo 0
        Application.DisplayAlerts = True
        
        Set wb = Nothing
        Set auxSheet = Nothing
        
        VB_SITEDB.SetSiteBOMRev site_name, VB_SITEDB.GetSiteBOMRev(site_name) + 1
    End If
    
    ThisWorkbook.Save
    
ResetApplication:
    If ui_change Then RenderUI True

    Set site_sheet = Nothing
End Sub

Public Sub SiteBOMCompare(ByVal site_name As String, ByVal prevBOM As Worksheet)
'Subroutine: SiteBOMCompare - Handles a site BOM comparison. Places the two side by side, and highlights changes.
'Arguments: site_name - name of site in comparison.
'           prevBOM - worksheet object for prev BOM Sheet.

    
    Dim first_row As Integer
    first_row = VB_SITEBOM.FirstRow()
    
    Dim site_sheet_name As String
    If Len(site_name) > 25 Then
        site_sheet_name = Left(site_name, 25) & " - BOM"
    Else
        site_sheet_name = site_name & " - BOM"
    End If
    
    Dim thisBOM As Worksheet
    Set thisBOM = ThisWorkbook.Sheets(site_sheet_name)
    
    Dim those_columns As Integer
    Dim this_columns As Integer
    those_columns = prevBOM.UsedRange.Columns.count
    this_columns = thisBOM.UsedRange.Columns.count
    
    Dim insert_col As Integer
    insert_col = thisBOM.UsedRange.Columns.count + 1
    
    'set Destination Range for past BOM
    Dim destin As Range
    If those_columns = this_columns - 1 Then
        Set destin = thisBOM.Range(ColLet(insert_col + 1) & prevBOM.UsedRange.Rows(1).row & ":" & _
            ColLet(insert_col + prevBOM.UsedRange.Columns.count) & prevBOM.UsedRange.Rows.count)
    Else
        Set destin = thisBOM.Range(ColLet(insert_col) & prevBOM.UsedRange.Rows(1).row & ":" & _
            ColLet(insert_col + prevBOM.UsedRange.Columns.count - 1) & prevBOM.UsedRange.Rows.count)
    End If
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    'Copy/Paste old BOM side by side with current BOM.
    Application.DisplayAlerts = False
    If those_columns = this_columns + 1 Then
        prevBOM.Range(ColLet(2) & prevBOM.UsedRange.Rows(1).row & ":" & _
            ColLet(prevBOM.UsedRange.Columns(prevBOM.UsedRange.Columns.count).column) & prevBOM.UsedRange.Rows.count).Copy
    Else
        prevBOM.UsedRange.Copy
    End If
    destin.PasteSpecial xlPasteColumnWidths
    thisBOM.Paste destination:=destin
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    
    'mark past BOM title as OBSOLETE
    If those_columns = this_columns - 1 Then
        thisBOM.Cells(2, insert_col + 1).MergeArea.Cells(1, 1).Value = site_name & " (OBSOLETE)"
        thisBOM.Columns(insert_col).Hidden = True
    ElseIf those_columns = this_columns + 1 Then
        thisBOM.Range(ColLet(insert_col + 1) & "1:" & _
            ColLet(thisBOM.UsedRange.Columns(thisBOM.UsedRange.Columns.count).column) & "1").Merge
        thisBOM.Range(ColLet(insert_col + 1) & "2:" & _
            ColLet(thisBOM.UsedRange.Columns(thisBOM.UsedRange.Columns.count).column) & "2").Merge
            
        thisBOM.Cells(1, insert_col + 1).MergeArea.Cells(1, 1).Value = thisBOM.Cells(1, 1).MergeArea.Cells(1, 1).Value
        thisBOM.Cells(2, insert_col + 1).MergeArea.Cells(1, 1).Value = site_name & " (OBSOLETE)"
    Else
        thisBOM.Cells(2, insert_col).MergeArea.Cells(1, 1).Value = site_name & " (OBSOLETE)"
    End If
    
    'Sort Sides
    Dim this_mark_col As Integer
    Dim that_mark_col As Integer
    If those_columns = this_columns - 1 Then
        this_mark_col = 2
        that_mark_col = insert_col
    ElseIf those_columns = this_columns + 1 Then
        this_mark_col = 1
        that_mark_col = insert_col
    ElseIf this_columns = 5 Then
        this_mark_col = 2
        that_mark_col = insert_col + 1
    Else
        this_mark_col = 1
        that_mark_col = insert_col
    End If
    
    ' this side
    thisBOM.Range("A" & (first_row - 1) & ":" & ColLet(thisBOM.UsedRange.Columns.count / 2) & thisBOM.UsedRange.Rows.CountLarge).Sort _
        Key1:=thisBOM.Range(ColLet(this_mark_col) & first_row), Order1:=xlAscending, Header:=xlYes
    ' prev side
    thisBOM.Range(ColLet(insert_col) & (first_row - 1) & ":" & ColLet(thisBOM.UsedRange.Columns.count) & thisBOM.UsedRange.Rows.CountLarge).Sort _
        Key1:=thisBOM.Range(ColLet(that_mark_col) & first_row), Order1:=xlAscending, Header:=xlYes
    
    Dim r_this As Integer
    Dim c As Integer
    
    Dim this_mark As Integer
    Dim prev_mark As Integer
    
    With thisBOM
    
        'line rows up to matching mark numbers.
        For r_this = first_row To .UsedRange.Rows.count
            this_mark = .Cells(r_this, 2).Value2
            prev_mark = .Cells(r_this, insert_col + 1).Value2
            
            Do While prev_mark > this_mark And prev_mark <> 0 And this_mark <> 0
                With thisBOM.Range(ColLet(thisBOM.UsedRange.Columns.count / 2 + 1) & r_this & ":" & ColLet(thisBOM.UsedRange.Columns.count) & r_this)
                    .Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                End With
                With thisBOM.Range(ColLet(thisBOM.UsedRange.Columns.count / 2 + 1) & r_this + 1 & ":" & ColLet(thisBOM.UsedRange.Columns.count) & r_this + 1).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                
                r_this = r_this + 1
                
                this_mark = .Cells(r_this, 2).Value2
                prev_mark = .Cells(r_this, insert_col + 1).Value2
            Loop
            
            Do While this_mark > prev_mark And prev_mark <> 0 And this_mark <> 0
                With thisBOM.Range("A" & r_this & ":" & ColLet(thisBOM.UsedRange.Columns.count / 2) & r_this)
                    .Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
                End With
                With thisBOM.Range("A" & r_this + 1 & ":" & ColLet(thisBOM.UsedRange.Columns.count / 2) & r_this + 1).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                r_this = r_this + 1
                
                this_mark = .Cells(r_this, 2).Value2
                prev_mark = .Cells(r_this, insert_col + 1).Value2
            Loop
            
            If this_mark <> prev_mark And prev_mark <> 0 And this_mark <> 0 Then
                r_this = r_this - 1
            End If
            
        Next r_this
        
        'Compare...
        For r_this = first_row To .UsedRange.Rows.count
            this_mark = .Cells(r_this, 2).Value2
            prev_mark = .Cells(r_this, insert_col + 1).Value2
            
            If this_mark = prev_mark Then
               For c = 1 To insert_col - 1
                    If DescriptionCompare(.Cells(r_this, c).Value2, .Cells(r_this, c + insert_col - 1).Value2) <> 0 Then
                        .Cells(r_this, c).Interior.Color = RGB(255, 255, 0)
                        .Cells(r_this, c + insert_col - 1).Interior.Color = RGB(255, 255, 0)
                    End If
                Next c
            End If
        Next r_this
        
    End With
    
    'extend print area
    thisBOM.PageSetup.PrintArea = thisBOM.UsedRange.Address
    If thisBOM.Columns(insert_col).Hidden Then
        Set thisBOM.VPageBreaks(1).Location = thisBOM.Range("$" & ColLet(insert_col + 1) & "$" & first_row)
    Else
        Set thisBOM.VPageBreaks(1).Location = thisBOM.Range("$" & ColLet(insert_col) & "$" & first_row)
    End If
    
    If ui_change Then RenderUI True
    
    Set thisBOM = Nothing
End Sub

Public Function GetPreviousSiteBOMFileName() As String
'Function: GetPreviousSiteBOMFileName - Handles 'Open' dialogue and return the file path for the file choosen.
'Returns: String containing the file path. If dialogue was cancelled, returns vbNullString

    
    On Error Resume Next
    ChDir ThisWorkbook.Path & VB_VAR_STORE.GetSiteBOMDirectory
    Err.Clear
    On Error GoTo 0
    
    Dim filename As Variant
    filename = Application.GetOpenFilename(FileFilter:="Microsoft Excel Workbook (*.xlsx; *.xls), *.xlsx; *.xls", _
        FilterIndex:=1, _
        title:="Select Previous Site BOM...", _
        MultiSelect:=False)
    
    If filename <> False Then
        GetPreviousSiteBOMFileName = CStr(filename)
    Else
        GetPreviousSiteBOMFileName = vbNullString
    End If
End Function

Private Sub EmptySiteBOM(ByVal site_name As String)
'Subroutine: EmptySiteBOM - Empties Site BOM before refilling it with the most updated material data
'Arguments: site_name - name of site to perform operation for.


    Dim row As Integer
    Dim item_col As Integer
    Dim site_sheet As Worksheet
    
    
    Dim site_sheet_name As String
    If Len(site_name) > 25 Then
        site_sheet_name = Left(site_name, 25) & " - BOM"
    Else
        site_sheet_name = site_name & " - BOM"
    End If
    Set site_sheet = ThisWorkbook.Sheets(site_sheet_name)
    
    row = VB_SITEBOM.FirstRow()
    item_col = ColNum("B")
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    Do While row <= site_sheet.UsedRange.Rows.count
        site_sheet.UsedRange.Rows(row).ClearContents
        
        row = row + 1
    Loop

    If ui_change Then RenderUI True

    Set site_sheet = Nothing
End Sub

Public Function SiteExists(ByVal site_name As String, Optional ByRef return_site_col As Integer) As Boolean
'Function: SiteExists - Determines whether a site exists in this project.
'Arguments: site_name - name of site to test
'           return_site_col - Reference to Integer. If site exists, Integer will contain the retrieved site column
'Returns: Boolean. True - Site exists. False - Site doesn't exist.


    SiteExists = False

    Dim index As Integer
    Dim rSites As Range
    Set rSites = GetSitesRange()
    
    'loop through site range looking for site_name
    For index = 1 To rSites.count
        If CStr(rSites.Cells(1, index).Value2) = site_name Then
            SiteExists = True
            Exit For
        End If
    Next index
    
    'set returns
    If SiteExists Then
        return_site_col = index + get_col_num("Current Model Quantities") - 1
    Else
        return_site_col = -1
    End If
    
    Set rSites = Nothing
End Function

Public Function TotalOrder4SiteItem(ByVal site_name As String, ByVal item_num As Integer) As Integer
'Function: TotalOrder4SiteItem - Gets the total quantity of a specific material item ordered for a specific site.
'Arguments: site_name - name of specific site
'           item_num - mark number for item in question
'Returns: Integer containing the total quantity ordered for a certain item at a certain site.


    'check if site exists
    If Not (SiteExists(site_name) Or site_name = "PROJECT") Then
        ErrorHandling "TotalOrder4SiteItem", 10, "Site: " & site_name & " doesn't exist.", 1
        Exit Function
    End If
    
    Dim item_row As Integer
    item_row = get_row(item_num)
    
    'check if item exists
    If item_row = -1 Then
        ErrorHandling "TotalOrder4SiteItem", 10, "Item: " & item_num & " doesn't exist.", 1
        Exit Function
    End If
    
    TotalOrder4SiteItem = 0
    
    Dim order_row As Integer
    order_row = VB_ORDER_LOG.FirstRow()
    
    Dim oname As String
    Dim osite As String
    Dim ocol As Integer
    
    'loop through orders. if order site is site_name, and order is received, add to totel
    Do While Not VB_ORDER_LOG.EOF(order_row)
        oname = VB_ORDER_LOG.OrderNumber(order_row - VB_ORDER_LOG.FirstRow())
        osite = VB_ORDER_LOG.get_order_site(oname)
        
        If osite = site_name And VB_ORDER_LOG.OrderIsReceived(VB_ORDER_LOG.OrderIndex(oname)) Then
            'get quantity from MASTER
            ocol = VB_ORDER_LOG.get_order_col(oname)
            If Not IsEmpty(VB_MASTER.Cells(item_row, ocol)) Then
                TotalOrder4SiteItem = TotalOrder4SiteItem + CInt(VB_MASTER.Cells(item_row, ocol).Value2)
            End If
        End If
        
        order_row = order_row + 1
    Loop
    
End Function

Public Function SiteDeficit(ByVal site_name As String, ByVal item_num As Integer) As Integer
'Function: SiteDeficit - Gets the site deficit for specific item.
'Arguments: site_name - name of specific site
'           item_num - mark number for item in question
'Returns: Integer containing the order deficit for a certain item at a certain site.


    Dim site_col As Integer
    Dim running_surplus As Integer
    running_surplus = CInt(get_property(item_num, "Client Inventory")) + TotalOrder4SiteItem("PROJECT", item_num)
    Dim item_row As Integer
    item_row = get_row(item_num)
    
    Dim num_sites As Integer
    num_sites = NumSites()
    
    Dim col_diff As Integer
    col_diff = get_col_num("Model Extras") - get_col_num("Current Model Quantities")
    
    Dim i As Integer
    Dim supply As Integer
    Dim demand As Integer
    Dim quantity_needed As Integer
    SiteDeficit = 0
    
    'loop through sites
    For i = 0 To num_sites - 1
        'deduct from running surplus the quantity required for site: orders and extras
        'if not there isn't enough material in running surplus to accomodate the
        'demand at a site, quantity needed is the difference.
        
        If Not SiteExists(SiteNamebyIndex(i), site_col) Then
            GoTo NEXT_ITEM
        End If
        
        If IsFabPackage(VB_MASTER.Cells(VB_MASTER.SubtitleRow(), site_col)) Then
            i = i - 1
            GoTo NEXT_ITEM
        End If
        
        demand = VB_MASTER.Cells(item_row, site_col).Value2 + VB_MASTER.Cells(item_row, site_col + col_diff).Value2
        supply = TotalOrder4SiteItem(VB_MASTER.Cells(VB_MASTER.SubtitleRow(), site_col).Value2, item_num)
        
        If running_surplus >= (demand - supply) And i < num_sites - 1 Then
            quantity_needed = 0
            running_surplus = running_surplus - (demand - supply)
        Else
            quantity_needed = (demand - supply) - running_surplus
            running_surplus = 0
        End If
        
        'if site = site_name, return
        If SiteNamebyIndex(i) = site_name Then
            SiteDeficit = quantity_needed
            Exit Function
        End If
        
NEXT_ITEM:
    Next i

End Function

Public Sub RemoveSite(ByVal site_name As String)
'Subroutine: RemoveSite - Removes site from project
'Arguments: site_name - name of site to remove

    
    'Site column variables for columns in Current Model Quantities, Model Extras, and Check Quantities
    Dim site_col As Integer
    Dim site_col_2 As Integer
    Dim site_col_3 As Integer
    
    If Not SiteExists(site_name, site_col) Then
        ErrorHandling "RemoveSite", 10, "Site: " & site_name & " doesn't exist.", 1
        Exit Sub
    End If
    
    site_col_2 = get_col_num("Model Extras") + site_col - get_col_num("Current Model Quantities")
    site_col_3 = get_col_num("Checked Quantities") + site_col - get_col_num("Current Model Quantities")
    
    Dim fab_package As Boolean
    fab_package = IsFabPackage(site_name)
    
    Dim num_sites As Range
    Set num_sites = GetSitesRange()
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    Dim first As Integer
    Dim last As Integer
    first = VB_MASTER.FirstRow()
    last = VB_MASTER.LastRow()
    
    If num_sites.count > 1 Then
        DeleteMasterColumn site_col_3
        DeleteMasterColumn site_col_2
        DeleteMasterColumn site_col
    Else
        VB_MASTER.Range(ColLet(site_col) & VB_MASTER.SubtitleRow() & ":" & ColLet(site_col) & last).ClearContents
        VB_MASTER.Range(ColLet(site_col_2) & VB_MASTER.SubtitleRow() & ":" & ColLet(site_col_2) & last).ClearContents
        VB_MASTER.Range(ColLet(site_col_3) & VB_MASTER.SubtitleRow() & ":" & ColLet(site_col_3) & last).ClearContents
        
        With VB_MASTER.Range(ColLet(site_col) & VB_MASTER.SubtitleRow() & ":" & ColLet(site_col) & VB_MASTER.SubtitleRow2()).Font
            .Bold = True
            .Italic = False
        End With
        With VB_MASTER.Range(ColLet(site_col_2) & VB_MASTER.SubtitleRow() & ":" & ColLet(site_col_2) & VB_MASTER.SubtitleRow2()).Font
            .Bold = True
            .Italic = False
        End With
        With VB_MASTER.Range(ColLet(site_col_3) & VB_MASTER.SubtitleRow() & ":" & ColLet(site_col_3) & VB_MASTER.SubtitleRow2()).Font
            .Bold = True
            .Italic = False
        End With
    End If
    
    AutoFitCells "Current Model Quantities"
    AutoFitCells "Model Extras"
    AutoFitCells "Checked Quantities"
    
    'if order site = site_name, order becomes part of general surplus (site: PROJECT)
    Dim num_orders As Integer
    num_orders = VB_ORDER_LOG.NumOrders()
    If num_orders > 0 Then
        Dim ord_col As Integer
        Dim o As Integer
        Dim orow As Integer
        orow = VB_ORDER_LOG.FirstRow()
        For o = 1 To num_orders
            If VB_ORDER_LOG.Cells(orow, VB_ORDER_LOG.ColID("Site")).Value2 = site_name Then
                VB_ORDER_LOG.Cells(orow, VB_ORDER_LOG.ColID("Site")).Value = "PROJECT"
                ord_col = VB_ORDER_LOG.get_order_col(VB_ORDER_LOG.OrderNumber(o - 1))
                VB_MASTER.Cells(VB_MASTER.SubtitleRow2(), ord_col).Value = VB_ORDER_LOG.Cells(orow, VB_ORDER_LOG.ColID("Date")).Value2
            End If
            orow = orow + 1
        Next o
    End If
    
    Dim site_sheet_name As String
    If Len(site_name) > 25 Then
        site_sheet_name = Left(site_name, 25) & " - BOM"
    Else
        site_sheet_name = site_name & " - BOM"
    End If
    
    If SheetExists(site_sheet_name) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(site_sheet_name).Delete
        Application.DisplayAlerts = True
    End If
    
    'remove site from DATABASE
    VB_SITEDB.RemoveSite site_name
    
    'remove site from RFP_REPORT
    VB_RFP_REPORT.RemoveSite site_name
    
    'remove site in DRAFTING CHECK COPY
    VB_DFTGCHECK.RemoveSite site_name
    
    ResetSiteNamesRange
    
    Dim row As Integer
    For row = first To last
        SetRowFormulas get_mark_num(row)
    Next row
    
    VB_MASTER.CalculateQuantityFormat
    
    If ui_change Then RenderUI True
    
    'LOG CHANGE
    If VB_CHANGE_LOG.TrackChanges() Then
        If VB_CHANGE_LOG.LogChange(VB_MASTER.Name & "!$" & ColLet(site_col) & "$" & VB_MASTER.SubtitleRow(), "", "Purged Site: " & site_name, "(" & IIf(fab_package, "1", "0") & ")" & site_name, "") <> 0 Then
            'throw error
            ErrorHandling "Ribbon_RemoveSite", 0, "LogChange() operation failed:" & vbCrLf & _
                "LogChange(" & VB_MASTER.Name & "!$" & ColLet(site_col) & "$" & VB_MASTER.SubtitleRow() & ", , Purged Site: " & site_name & ", " & site_name & ", )", 1
        End If
    End If
    
    Set num_sites = Nothing
End Sub

Public Sub AddSite(ByRef site_name As String, ByVal fab_package As Boolean)
'Subroutine: AddSite - Adds site to Project. This column signifies a fab-package, the font is changed to designate that.
'Arguments: site_name - new site name

    'check if site exists
    If SiteExists(site_name) Then
        ErrorHandling "AddSite", 0, "Site: " & site_name & " already exists.", 1
        site_name = "ERR"
        Exit Sub
    End If
    
    Dim num_sites As Integer
    Dim curr_sites As Range
    Set curr_sites = GetSitesRange()
    num_sites = curr_sites.count
    
    Dim last_col As Integer
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    'add columns if necessary
    If IsEmpty(curr_sites.Cells(1, 1)) Then
        last_col = get_col_num("Current Model Quantities")
    Else
        last_col = get_col_num("Current Model Quantities") + num_sites
        
        InsertMasterColumn last_col + 2 * num_sites + 1
        InsertMasterColumn last_col + num_sites
        InsertMasterColumn last_col
        
        num_sites = num_sites + 1
    End If

    ' HEADING FONT
    With VB_MASTER.Range(ColLet(last_col) & VB_MASTER.SubtitleRow() & ":" & ColLet(last_col) & VB_MASTER.SubtitleRow2()).Font
        .Bold = Not fab_package
        .Italic = fab_package
    End With
    
    With VB_MASTER.Range(ColLet(last_col + num_sites) & VB_MASTER.SubtitleRow() & ":" & ColLet(last_col + num_sites) & VB_MASTER.SubtitleRow2()).Font
        .Bold = Not fab_package
        .Italic = fab_package
    End With
    
    With VB_MASTER.Range(ColLet(last_col + 2 * num_sites + 1) & VB_MASTER.SubtitleRow() & ":" & ColLet(last_col + 2 * num_sites + 1) & VB_MASTER.SubtitleRow2()).Font
        .Bold = Not fab_package
        .Italic = fab_package
    End With
    
    'autofit columns
    VB_MASTER.Cells(VB_MASTER.SubtitleRow(), last_col + 2 * num_sites + 1).Value = site_name
    AutoFitCells "Checked Quantities"
    
    VB_MASTER.Cells(VB_MASTER.SubtitleRow(), last_col + num_sites).Value = site_name
    AutoFitCells "Model Extras"
    
    VB_MASTER.Cells(VB_MASTER.SubtitleRow(), last_col).Value = site_name
    AutoFitCells "Current Model Quantities"
    
    'LOG CHANGE
    If VB_CHANGE_LOG.TrackChanges() Then
        If VB_CHANGE_LOG.LogChange(VB_MASTER.Name & "!$" & ColLet(last_col) & "$3", "", "Added Site: " & site_name, "", "(" & IIf(fab_package, "1", "0") & ")" & site_name) <> 0 Then
            'throw error
            ErrorHandling "AddSite", 0, "LogChange() operation failed:" & vbCrLf & _
                "LogChange(" & VB_MASTER.Name & "!$" & ColLet(last_col) & "$3, , Added Site: " & site_name & ", , (" & IIf(fab_package, "1", "0") & ")" & site_name & ")", 1
        End If
    End If
    
    ResetSiteNamesRange
    
    'adjust formulas
    Dim first As Integer
    Dim last As Integer
    first = VB_MASTER.FirstRow()
    last = VB_MASTER.LastRow()
    Dim row As Integer
    For row = first To last
        SetRowFormulas get_mark_num(row)
    Next row
    
    VB_MASTER.CalculateQuantityFormat
    
    'ADD SITE TO SITEDB
    VB_SITEDB.AddSite site_name
    
    'ADD SITE TO RFP REPORT
    If VB_RFP_REPORT.Form_col_num(VB_RFP_REPORT, site_name) = -1 And Not fab_package Then
        VB_RFP_REPORT.AddSite site_name
    End If
    
    'ADD SITE TO DRAFTING CHECK COPY
    VB_DFTGCHECK.AddSite site_name
    
    If ui_change Then RenderUI True
    
    Set curr_sites = Nothing
End Sub

Public Sub RenameSite(ByVal old_site_name, ByVal site_name As String, ByVal fab_package As Boolean)
'Subroutine: RenameSite - Renames and changes state of site in Project
'Arguments: old_site_name - original name of site
'           site_name - new site name


    'check if new site name exists
    If SiteExists(site_name) And old_site_name <> site_name Then
        ErrorHandling "RenameSite", 10, "Site: " & site_name & " already exists.", 1
        Exit Sub
    End If
    
    'check if old site name exists, get site column
    Dim site_col As Integer
    SiteExists old_site_name, site_col
    
    Dim curr_sites As Range
    Set curr_sites = GetSitesRange()
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    'change name
    'autofit columns
    VB_MASTER.Cells(VB_MASTER.SubtitleRow(), site_col + 2 * curr_sites.count + 1).Value = site_name
    AutoFitCells "Checked Quantities"
    
    VB_MASTER.Cells(VB_MASTER.SubtitleRow(), site_col + curr_sites.count).Value = site_name
    AutoFitCells "Model Extras"
    
    VB_MASTER.Cells(VB_MASTER.SubtitleRow(), site_col).Value = site_name
    AutoFitCells "Current Model Quantities"
    
    'check state
    Dim state_change As Boolean
    state_change = False
    If IsFabPackage(site_name) <> fab_package Then
        state_change = True
        With VB_MASTER.Range(ColLet(site_col) & VB_MASTER.SubtitleRow() & ":" & ColLet(site_col) & VB_MASTER.SubtitleRow2()).Font
            .Bold = Not fab_package
            .Italic = fab_package
        End With
        
        With VB_MASTER.Range(ColLet(site_col + curr_sites.count) & VB_MASTER.SubtitleRow() & ":" & ColLet(site_col + curr_sites.count) & VB_MASTER.SubtitleRow2()).Font
            .Bold = Not fab_package
            .Italic = fab_package
        End With
        
        Dim first As Integer
        Dim last As Integer
        first = VB_MASTER.FirstRow()
        last = VB_MASTER.LastRow()
        Dim row As Integer
        For row = first To last
            SetRowFormulas get_mark_num(row)
        Next row
    End If
    
    'rename in SITEDB
    VB_SITEDB.RenameSite old_site_name, site_name
    
    'rename site in RFP REPORT
    If fab_package Then
        VB_RFP_REPORT.RemoveSite old_site_name
    Else
        If VB_RFP_REPORT.Form_col_num(VB_RFP_REPORT, old_site_name) > 0 Then
            VB_RFP_REPORT.RenameSite old_site_name, site_name
        Else
            VB_RFP_REPORT.AddSite site_name
        End If
    End If
    
    'rename site in DRAFTING CHECK COPY
    VB_DFTGCHECK.RenameSite old_site_name, site_name
    
    'rename site BOM if it exists
    Dim old_site_sheet_name As String
    If Len(old_site_name) > 25 Then
        old_site_sheet_name = Left(old_site_name, 25) & " - BOM"
    Else
        old_site_sheet_name = old_site_name & " - BOM"
    End If
    
    Dim new_site_sheet_name As String
    If Len(site_name) > 25 Then
        new_site_sheet_name = Left(site_name, 25) & " - BOM"
    Else
        new_site_sheet_name = site_name & " - BOM"
    End If
    
    If SheetExists(old_site_sheet_name) Then
        With ThisWorkbook.Sheets(old_site_sheet_name)
            .Name = new_site_sheet_name
            .Cells(2, 1).Value = site_name
            .Visible = xlSheetVisible
        End With
    End If
    
    'rename order sites
    Dim num_orders As Integer
    num_orders = VB_ORDER_LOG.NumOrders()
    If num_orders > 0 Then
        Dim ord_col As Integer
        Dim o As Integer
        Dim orow As Integer
        orow = VB_ORDER_LOG.FirstRow()
        For o = 1 To num_orders
            If CStr(VB_ORDER_LOG.Cells(orow, VB_ORDER_LOG.ColID("Site")).Value2) = old_site_name Then
                VB_ORDER_LOG.Cells(orow, VB_ORDER_LOG.ColID("Site")).Value = site_name
                ord_col = VB_ORDER_LOG.get_order_col(VB_ORDER_LOG.OrderNumber(o - 1))
                VB_MASTER.Cells(VB_MASTER.SubtitleRow2(), ord_col).Value = site_name
                
                VB_MASTER.Columns(ord_col).AutoFit
                If VB_MASTER.Cells(VB_MASTER.SubtitleRow2(), ord_col).ColumnWidth < 8.86 Then
                    VB_MASTER.Cells(VB_MASTER.SubtitleRow2(), ord_col).ColumnWidth = 8.86
                End If
                
            End If
            orow = orow + 1
        Next o
    End If
    
    'LOG CHANGE
    If VB_CHANGE_LOG.TrackChanges() Then
        If VB_CHANGE_LOG.LogChange(VB_MASTER.Name & "!$" & ColLet(site_col) & "$3", "", "Renamed " & old_site_name & ": " & site_name, "(" & IIf(state_change, IIf(fab_package, "0", "1"), IIf(fab_package, "1", "0")) & ")" & old_site_name, "(" & IIf(fab_package, "1", "0") & ")" & site_name) <> 0 Then
            'throw error
            ErrorHandling "RenameSite", 0, "LogChange() operation failed:" & vbCrLf & _
                "LogChange(" & VB_MASTER.Name & "!$" & ColLet(site_col) & "$3, , Renamed " & old_site_name & ": " & site_name & ", " & old_site_name & ", " & site_name & ")", 1
        End If
    End If
    
    If ui_change Then RenderUI True

    Set curr_sites = Nothing
End Sub

Public Function SiteIndex(ByVal site_name As String) As Integer
'Function: SiteIndex - Treats SitesRange like a collection, EXCLUDING fab packages, and returns the
'                      index number for a given site name. Order is the same order shown in MASTER.
'Arguments: site_name - Site name to search for
'Returns: Integer containing the index number [0-#]


    Dim site_col As Integer
    
    If Not SiteExists(site_name, site_col) Or IsFabPackage(site_name) Then
        SiteIndex = -1
        Exit Function
    End If
    
    Dim end_col As Integer
    end_col = get_col_num("Current Model Quantities")
    SiteIndex = site_col - end_col
    
    site_col = site_col - 1
    Do While site_col >= end_col
        If IsFabPackage(VB_MASTER.Cells(VB_MASTER.SubtitleRow(), site_col)) Then
            SiteIndex = SiteIndex - 1
        End If
        
        site_col = site_col - 1
    Loop
End Function

Public Function SiteNamebyIndex(ByVal site_index As Integer) As String
'Function: SiteIndex - Treats SitesRange like a collection, EXCLUDING fab packages, and returns the
'                      site name for a given index. Order is the same order shown in MASTER.
'Arguments: site_index - Integer containing the index number [0-#]
'Returns: String containing the site name


    If site_index >= NumSites() Then
        SiteNamebyIndex = vbNullString
        Exit Function
    End If
    
    Dim rSites As Range
    Set rSites = GetSitesRange()
    Dim vbCell
    
    For Each vbCell In rSites
        If SiteIndex(vbCell.Value2) = site_index Then
            SiteNamebyIndex = vbCell.Value2
            Exit Function
        End If
    Next vbCell

End Function

Public Function IsFabPackage(ByVal site_name As String) As Boolean
'Function: IsFabPackage - Determines whether a 'site' name corresponds to a fab package, or an actual site.
'Arguments: site_name - String containing a site name. If site doesn't exists, returns False.
'Returns: Boolean. True - It is a fab package. False - it is NOT a fab package.


    Dim site_col As Integer
    
    If Not SiteExists(site_name, site_col) Then
        IsFabPackage = False
        Exit Function
    End If
    
    IsFabPackage = VB_MASTER.Cells(VB_MASTER.SubtitleRow(), site_col).Font.Italic
End Function

Public Sub ResetSiteNamesRange()
'Subroutine: ResetSiteNamesRange - Refreshes sites Range


    Dim rList As String
    Dim sites As Name
    Set sites = ThisWorkbook.Names.Item("site_names")
    
    Dim site_col As Integer
    site_col = get_col_num("Current Model Quantities")
    
    Dim end_col As Integer
    end_col = site_col
    
    Do While get_ColTitle(end_col + 1) = "Current Model Quantities"
        end_col = end_col + 1
    Loop
    
    'set new Range
    sites.RefersTo = "=" & VB_MASTER.Name & "!$" & ColLet(site_col) & "$3:$" & ColLet(end_col) & "$3"
    Set sites = Nothing
End Sub

Public Function GetSitesRange() As Range
'Function: GetSitesRange - Retrieves Range in VB_MASTER where sites are named
'Returns: Range object pointing to the site names under Current Model Quantities


    Dim sites As Name
    Set sites = ThisWorkbook.Names.Item("site_names")

    Set GetSitesRange = Range(sites.RefersTo)
    Set sites = Nothing
End Function

Public Function NumSites() As Integer
'Function: NumSites - Calculates number of sites
'Returns: Integer containing the number of sites


    Dim sites As Range
    Set sites = GetSitesRange()
    Dim vbCell
    
    NumSites = 0
    
    For Each vbCell In sites
        If IsEmpty(vbCell) Or IsFabPackage(vbCell.Value2) Then
            ' don't count
        Else
            NumSites = NumSites + 1
        End If
    Next vbCell
    
    Set sites = Nothing
End Function
