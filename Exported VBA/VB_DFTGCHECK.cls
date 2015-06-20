VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VB_DFTGCHECK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const document_title = 2
Private Const title_row = 4
Private Const sites_row = 5
Private Const first_row = 6

Private Const mark_col = 2
Private Const model_col = 3
Private Const unit_col = 4
Private Const desc_col = 5
Private Const quantities_col = 6

Public Sub PrepareDraftingCheckCopy()
'Subroutine: PrepareDraftingCheckCopy - Prepares the Drafting Check Copy for PDFing.
'            Adds appropraite material items to the table and formatts it properly.

    Dim row As Integer
    Dim mst_row As Integer
    Dim mark_num As Integer
    
    Dim first As Integer
    Dim last As Integer
    first = VB_MASTER.FirstRow()
    last = VB_MASTER.LastRow()
    row = first_row
    
    Dim col As Integer
    Dim num_sites As Integer
    num_sites = GetSitesRange().Cells.count
    
    Dim current_category As String
    current_category = vbNullString
    
    ' Add material item if is not Deleted
    For mst_row = first To last
        mark_num = get_mark_num(mst_row)
        If Not IsDeleted(mark_num) Then
            
            'Add row to table
            FormatRow row
            
            If get_category(mst_row) <> current_category Then
                With VB_DFTGCHECK.Range(ColLet(mark_col) & row & ":" & ColLet(quantities_col + 2 * num_sites - 1) & row).Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = 1
                    .Weight = xlThick
                End With
                
                current_category = get_category(mst_row)
            End If
            
            ' copy information for item
            With VB_DFTGCHECK
                .Cells(row, mark_col).Value = get_property(mark_num, "Mark No.")
                .Cells(row, model_col).Value = get_property(mark_num, "Model Total")
                .Cells(row, unit_col).Value = get_property(mark_num, "Unit")
                .Cells(row, desc_col).Value = get_property(mark_num, "Long Description")
            
            
                For col = quantities_col To quantities_col + num_sites - 1
                    .Cells(row, col).Value = VB_MASTER.Cells(mst_row, get_col_num("Current Model Quantities") + col - quantities_col).Value2
                Next col
                
                For col = quantities_col + num_sites To quantities_col + 2 * num_sites - 1
                    .Cells(row, col).Value = VB_MASTER.Cells(mst_row, get_col_num("Checked Quantities") + col - (quantities_col + num_sites)).Value2
                    
                    If .Cells(row, col).Value <> .Cells(row, col - num_sites).Value Then
                        .Cells(row, col).Interior.Color = RGB(255, 255, 0)
                    End If
                Next col
                
                .Rows(row).AutoFit
                
            End With
            
            ' next row
            row = row + 1
        End If
    Next mst_row
    
    AutoFitSiteColumns "Current Model Quantities"
    AutoFitSiteColumns "Checked Quantities"
    
    With VB_DFTGCHECK.Range(ColLet(mark_col) & row & ":" & ColLet(quantities_col + 2 * num_sites - 1) & row).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 1
        .Weight = xlThick
    End With

    
    With VB_DFTGCHECK.PageSetup
        .PrintArea = "$" & ColLet(mark_col) & "$" & document_title _
            & ":$" & ColLet(quantities_col + 2 * num_sites - 1) & "$" & row
        
        .LeftFooter = "&L&D"
        .CenterFooter = "&C" & VB_DFTGCHECK.Name
        .RightFooter = "&RPage &P of &N"
        
        .PaperSize = xlPaperLetter
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        
        .PrintTitleRows = VB_DFTGCHECK.Rows(title_row & ":" & sites_row).Address
        .PrintTitleColumns = ""
    End With
End Sub

Public Sub EmptyCheckCopy()
'Subroutine: EmptyCheckCopy - Clears information in the table for next time.

    With VB_DFTGCHECK
        If Not IsEmpty(.Cells(first_row, mark_col)) Then
            .Rows(first_row & ":" & .UsedRange.Rows(.UsedRange.Rows.CountLarge).row).Delete
        End If
        
        .UsedRange.Columns.Hidden = False
    End With
    
    AutoFitSiteColumns "Current Model Quantities"
    AutoFitSiteColumns "Checked Quantities"
End Sub

Public Sub AddSite(ByVal site_name As String)
'Subroutine: AddSite - Works in parallel with adding sites to the Master BOM. Adds a site to this table.
'Arguments: site_name - Name of the new site.

    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    Dim model_quan_col As Integer
    Dim check_col As Integer
    Dim this_num_sites As Integer
    this_num_sites = VB_DFTGCHECK.Cells(title_row, quantities_col).MergeArea.Cells.count
    
    If this_num_sites = 1 And IsEmpty(VB_DFTGCHECK.Cells(sites_row, quantities_col)) Then
        'no sites
        model_quan_col = quantities_col
        check_col = quantities_col + 1
        GoTo AddNames
    End If

    model_quan_col = quantities_col + this_num_sites
    check_col = quantities_col + 2 * this_num_sites
    
    ' ADD COLUMN TO CHECKED QUANTITIES
    
    VB_DFTGCHECK.Columns(check_col).Insert CopyOrigin:=xlFormatFromLeftOrAbove
    VB_DFTGCHECK.Range(ColLet(quantities_col + this_num_sites) & title_row & ":" & ColLet(check_col) & title_row).Merge
    
    With VB_DFTGCHECK
        With .Range(ColLet(check_col) & title_row & ":" & ColLet(check_col) & sites_row)
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Color = RGB(150, 54, 52)
                .Weight = xlThick
            End With
        End With
        With .Cells(title_row, check_col).MergeArea
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 1
                .Weight = xlMedium
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 1
                .Weight = xlMedium
            End With
        End With
        With .Cells(sites_row, check_col)
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 1
                .Weight = xlThin
            End With
            .HorizontalAlignment = xlCenter
        End With
    End With
    
    ' ADD COLUMN TO CURRENT MODEL QUANTITIES
    
    VB_DFTGCHECK.Columns(model_quan_col).Insert CopyOrigin:=xlFormatFromLeftOrAbove
    check_col = check_col + 1
    VB_DFTGCHECK.Range(ColLet(quantities_col) & title_row & ":" & ColLet(model_quan_col) & title_row).Merge
    
    With VB_DFTGCHECK
        With .Range(ColLet(model_quan_col) & title_row & ":" & ColLet(model_quan_col) & sites_row)
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Color = RGB(150, 54, 52)
                .Weight = xlThick
            End With
        End With
        With .Cells(title_row, model_quan_col).MergeArea
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 1
                .Weight = xlMedium
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 1
                .Weight = xlMedium
            End With
        End With
        With .Cells(sites_row, model_quan_col)
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 1
                .Weight = xlThin
            End With
            .HorizontalAlignment = xlCenter
        End With
    End With
    this_num_sites = this_num_sites + 1
    
AddNames:
    ' Add site name to the new column headings
    VB_DFTGCHECK.Cells(sites_row, check_col).Value = site_name
    VB_DFTGCHECK.Cells(sites_row, model_quan_col).Value = site_name
    
    Dim fab_package As Boolean
    fab_package = IsFabPackage(site_name)
    
    With VB_DFTGCHECK.Cells(sites_row, check_col).Font
        .Bold = Not fab_package
        .Italic = fab_package
    End With
    With VB_DFTGCHECK.Cells(sites_row, model_quan_col).Font
        .Bold = Not fab_package
        .Italic = fab_package
    End With
    
    VB_DFTGCHECK.Range("$" & ColLet(mark_col) & "$" & document_title & ":$" & _
        ColLet(quantities_col + 2 * this_num_sites - 1) & "$" & document_title).Merge
    
    Dim col As Integer
    
    VB_DFTGCHECK.Rows(sites_row).RowHeight = 15
    VB_DFTGCHECK.Rows(sites_row).WrapText = False
    VB_DFTGCHECK.Rows(sites_row).VerticalAlignment = xlCenter
    
    For col = quantities_col To quantities_col + this_num_sites - 1
        VB_DFTGCHECK.Columns(col).ColumnWidth = VB_MASTER.Columns(get_col_num("Current Model Quantities") + col - quantities_col).ColumnWidth
        VB_DFTGCHECK.Columns(col + this_num_sites).ColumnWidth = VB_MASTER.Columns(get_col_num("Checked Quantities") + col - quantities_col).ColumnWidth
        
        With VB_DFTGCHECK
            If .Columns(col).ColumnWidth > 24 Then
                .Rows(sites_row).RowHeight = 30
                .Columns(col).ColumnWidth = 0.6 * .Columns(col).ColumnWidth
                .Columns(col + this_num_sites).ColumnWidth = .Columns(col).ColumnWidth
                .Cells(sites_row, col).WrapText = True
                .Cells(sites_row, col + this_num_sites).WrapText = True
            End If
        End With
        
    Next col
    
    If ui_change Then RenderUI True
End Sub

Public Sub RenameSite(ByVal old_site_name As String, ByVal new_site_name As String)
'Subroutine: RenameSite - Works in parallel with renaming sites in the Master BOM.
'Arguments: old_site_name - Original name of the site.
'           new_site_name - New name of the site.

    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    Dim mst_site_col As Integer
    
    'test to if new_site_name exists in Master
    If Not SiteExists(new_site_name, mst_site_col) Then
        Exit Sub
    End If
    
    Dim model_quan_col As Integer
    Dim check_col As Integer
    Dim this_num_sites As Integer
    this_num_sites = VB_DFTGCHECK.Cells(title_row, quantities_col).MergeArea.Cells.count

    model_quan_col = quantities_col + mst_site_col - get_col_num("Current Model Quantities")
    check_col = model_quan_col + this_num_sites
    
    VB_DFTGCHECK.Cells(sites_row, check_col).Value = new_site_name
    VB_DFTGCHECK.Cells(sites_row, model_quan_col).Value = new_site_name
    
    Dim fab_package As Boolean
    fab_package = IsFabPackage(new_site_name)
    
    With VB_DFTGCHECK.Cells(sites_row, check_col).Font
        .Bold = Not fab_package
        .Italic = fab_package
    End With
    With VB_DFTGCHECK.Cells(sites_row, model_quan_col).Font
        .Bold = Not fab_package
        .Italic = fab_package
    End With
    
    'Match column widths to Master
    Dim col As Integer
    
    VB_DFTGCHECK.Rows(sites_row).RowHeight = 15
    VB_DFTGCHECK.Rows(sites_row).WrapText = False
    VB_DFTGCHECK.Rows(sites_row).VerticalAlignment = xlCenter
    
    For col = quantities_col To quantities_col + this_num_sites - 1
        VB_DFTGCHECK.Columns(col).ColumnWidth = VB_MASTER.Columns(get_col_num("Current Model Quantities") + col - quantities_col).ColumnWidth
        VB_DFTGCHECK.Columns(col + this_num_sites).ColumnWidth = VB_MASTER.Columns(get_col_num("Checked Quantities") + col - quantities_col).ColumnWidth
        
        With VB_DFTGCHECK
            If .Columns(col).ColumnWidth > 24 Then
                .Rows(sites_row).RowHeight = 30
                .Columns(col).ColumnWidth = 0.6 * .Columns(col).ColumnWidth
                .Columns(col + this_num_sites).ColumnWidth = .Columns(col).ColumnWidth
                .Cells(sites_row, col).WrapText = True
                .Cells(sites_row, col + this_num_sites).WrapText = True
            End If
        End With
        
    Next col
    
    If ui_change Then RenderUI True
End Sub

Public Sub RemoveSite(ByVal site_name As String)
'Subroutine: RenameSite - Works in parallel with removing sites from the Master BOM.
'Arguments: site_name - Name of the site to remove from this table.

    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    Dim model_quan_col As Integer
    Dim check_col As Integer
    Dim this_num_sites As Integer
    this_num_sites = VB_DFTGCHECK.Cells(title_row, quantities_col).MergeArea.Cells.count
    
    If this_num_sites = 0 Then
        Exit Sub
    End If
    
    model_quan_col = -1
    Dim col As Integer
    For col = quantities_col To quantities_col + this_num_sites - 1
        If VB_DFTGCHECK.Cells(sites_row, col).Value = site_name Then
            model_quan_col = col
            Exit For
        End If
    Next col
    
    If model_quan_col = -1 Then
        Exit Sub
    Else
        check_col = model_quan_col + this_num_sites
    End If
    
    ' REMOVE COLUMN FROM CHECKED QUANTITIES
    
    If this_num_sites = 1 Then
        VB_DFTGCHECK.Cells(sites_row, check_col).ClearContents
    Else
        VB_DFTGCHECK.Columns(check_col).Delete
        If check_col = quantities_col + this_num_sites Then
            VB_DFTGCHECK.Cells(title_row, check_col).MergeArea.Cells(1, 1).Value = "Checked Quantities"
        End If
    End If
    
    With VB_DFTGCHECK.Range(ColLet(check_col - 1) & title_row & ":" & ColLet(check_col - 1) & sites_row)
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Color = RGB(150, 54, 52)
            .Weight = xlThick
        End With
    End With
    
    ' REMOVE COLUMN FROM CURRENT MODEL QUANTITIES
    
    If this_num_sites = 1 Then
        VB_DFTGCHECK.Cells(sites_row, model_quan_col).ClearContents
    Else
        VB_DFTGCHECK.Columns(model_quan_col).Delete
        If model_quan_col = quantities_col Then
            VB_DFTGCHECK.Cells(title_row, model_quan_col).MergeArea.Cells(1, 1).Value = "Current Model Quantities"
        End If
        this_num_sites = this_num_sites - 1
    End If

    With VB_DFTGCHECK.Range(ColLet(model_quan_col - 1) & title_row & ":" & ColLet(model_quan_col - 1) & sites_row)
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Color = RGB(150, 54, 52)
            .Weight = xlThick
        End With
    End With
    
    VB_DFTGCHECK.Rows(sites_row).RowHeight = 15
    VB_DFTGCHECK.Rows(sites_row).WrapText = False
    VB_DFTGCHECK.Rows(sites_row).VerticalAlignment = xlCenter
    
    For col = quantities_col To quantities_col + this_num_sites - 1
        VB_DFTGCHECK.Columns(col).ColumnWidth = VB_MASTER.Columns(get_col_num("Current Model Quantities") + col - quantities_col).ColumnWidth
        VB_DFTGCHECK.Columns(col + this_num_sites).ColumnWidth = VB_MASTER.Columns(get_col_num("Checked Quantities") + col - quantities_col).ColumnWidth
        
        With VB_DFTGCHECK
            If .Columns(col).ColumnWidth > 24 Then
                .Rows(sites_row).RowHeight = 30
                .Columns(col).ColumnWidth = 0.6 * .Columns(col).ColumnWidth
                .Columns(col + this_num_sites).ColumnWidth = .Columns(col).ColumnWidth
                .Cells(sites_row, col).WrapText = True
                .Cells(sites_row, col + this_num_sites).WrapText = True
            End If
        End With
        
    Next col
    
    If ui_change Then RenderUI True
End Sub

Private Sub FormatRow(ByVal row_num As Integer)
'Subroutine: FormatRow - Appends a correctly formatted row to the table.

    Dim num_sites As Integer
    num_sites = GetSitesRange().Cells.count

    With VB_DFTGCHECK.Range(ColLet(mark_col) & row_num & ":" & ColLet(desc_col) & row_num)
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Color = RGB(150, 54, 52)
            .Weight = xlThick
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 1
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Color = RGB(150, 54, 52)
            .Weight = xlThick
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 1
            .Weight = xlThin
        End With
    End With
    
    With VB_DFTGCHECK.Range(ColLet(quantities_col) & row_num & ":" & ColLet(quantities_col + num_sites - 1) & row_num)
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Color = RGB(150, 54, 52)
            .Weight = xlThick
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 1
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Color = RGB(150, 54, 52)
            .Weight = xlThick
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 1
            .Weight = xlThin
        End With
    End With
    
    With VB_DFTGCHECK.Range(ColLet(quantities_col + num_sites) & row_num & ":" & ColLet(quantities_col + 2 * num_sites - 1) & row_num)
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Color = RGB(150, 54, 52)
            .Weight = xlThick
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 1
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Color = RGB(150, 54, 52)
            .Weight = xlThick
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 1
            .Weight = xlThin
        End With
    End With
    
    With VB_DFTGCHECK.Range(ColLet(mark_col) & row_num & ":" & ColLet(quantities_col + 2 * num_sites - 1) & row_num)
        .Interior.ColorIndex = 2
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlCenter
    End With
    With VB_DFTGCHECK.Cells(row_num, desc_col)
        .WrapText = True
        .HorizontalAlignment = xlLeft
    End With
End Sub

Private Sub AutoFitSiteColumns(ByVal col_title As String)
'Subroutine: AutoFitSiteColumns - Handles the autofitting in the Current Model Quantities and Checked Quantities columns.
'                                 The column titles are easily wider than the individual site columns. To keep the entire merged title cell
'                                 fitted correctly, we distibute the column widths appropriately to subsequent site columns.
'Arguments: - col_title - "Current Model Quantities" or "Checked Quantities"


    If Not (col_title = "Current Model Quantities" Or col_title = "Checked Quantities") Then
        Exit Sub
    End If

    Dim min_width As Double
    Dim col As Integer
    
    If col_title = "Current Model Quantities" Then
        min_width = 23.43 ' 169
        col = quantities_col
    ElseIf col_title = "Checked Quantities" Then
        min_width = 17.86 ' 130
        col = quantities_col + NumSites()
    End If
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    Dim rCell
    Dim i As Integer
    Dim last_col As Integer
    Dim total_width As Double
    Dim title_cell As Range
    Set title_cell = VB_DFTGCHECK.Cells(sites_row - 1, col).MergeArea
    i = 0
    last_col = 0
    total_width = 0
    
    For Each rCell In title_cell
    
        If Not VB_DFTGCHECK.Columns(col + i).Hidden Then
            VB_DFTGCHECK.Columns(col + i).AutoFit
            
            If VB_DFTGCHECK.Cells(sites_row, col + i).ColumnWidth < 8.86 Then
                VB_DFTGCHECK.Cells(sites_row, col + i).ColumnWidth = 8.86
            End If
            total_width = total_width + VB_DFTGCHECK.Cells(sites_row, col + i).ColumnWidth
            
            last_col = col + i
        End If
        
        i = i + 1
    Next rCell

    If total_width < min_width Then
        VB_DFTGCHECK.Cells(sites_row, last_col).ColumnWidth = VB_DFTGCHECK.Cells(sites_row, last_col).ColumnWidth + min_width - total_width
    End If
    
    If ui_change Then RenderUI True
    
    Set title_cell = Nothing
End Sub

Public Function QuantitiesColumn() As Integer
'Function: QuantitiesColumn - 'Get' method for the quantities_col property of this sheet
'Returns: Integer containing the quantities_col

    QuantitiesColumn = quantities_col
End Function
