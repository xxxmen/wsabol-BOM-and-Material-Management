Attribute VB_Name = "CADWorx_Module"
Option Explicit

Public Function GetCADFileName() As String
'Function: GetCADFileName - Handles 'Open' dialogue and return the file path for the file choosen.
'Returns: String containing the file path. If dialogue was cancelled, returns vbNullString


    On Error Resume Next
    ChDir ThisWorkbook.Path
    Err.Clear
    On Error GoTo 0
    
    Dim filename As Variant
    filename = Application.GetOpenFilename(FileFilter:="Microsoft Excel Workbooks (*.xlsx; *.xls), *.xlsx; *.xls", _
        FilterIndex:=1, _
        title:="Select CADWorx BOM...", _
        MultiSelect:=False)
    
    If filename <> False Then
        GetCADFileName = CStr(filename)
    Else
        GetCADFileName = vbNullString
    End If
End Function

Public Function ImportCADSheet(ByVal filename As String) As Integer
'Function: ImportCADSheet - Sets up the CADWorx sheet. Formats columns, converts pipe length,
'                           and sets up data validation, and finds description matches in this BOM.
'Arguments: filename - String containing the filename of the Excel workbook with the CAD BOM worksheet
'Returns: Integer error code. 0: no error, -1: File not from CADWorx


    ImportCADSheet = 0

    Dim CAD As Workbook
    Dim srcSheet As Worksheet
    Dim DstSheet As Worksheet
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    Set CAD = Application.Workbooks.Open(filename:=filename, ReadOnly:=True)
    Set srcSheet = CAD.Sheets(1)
    
    Dim quantity_col As Integer
    Dim length_col As Integer
    Dim unit_col As Integer
    Dim desc_col As Integer
    Dim desc_copy_col As Integer
    Dim match_col As Integer
    Dim new_col As Integer
    quantity_col = -1
    length_col = -1
    unit_col = -1
    desc_col = -1
    desc_copy_col = -1
    match_col = -1
    new_col = -1
    
    With srcSheet
        'Assign column headings
        Dim temp_col As Integer
        For temp_col = .UsedRange.Columns(1).column To .UsedRange.Columns(.UsedRange.Columns.count).column
            
            Select Case .Cells(1, temp_col).Value2
                
                Case "QUANTITY":
                    quantity_col = temp_col
                Case "LENGTH":
                    length_col = temp_col
                Case "LONG_DESC":
                    desc_col = temp_col
                Case "DESCRIPTION":
                    desc_col = temp_col
            End Select
            
        Next temp_col
    End With
    
    If quantity_col = -1 Or _
        length_col = -1 Or _
        desc_col = -1 Then
    
        If ui_change Then RenderUI True
    
        MsgBox "File is not from CADWorx. Please try again." & vbCrLf & vbCrLf & _
            "If you've received this message in error, please make sure that the CADWorx BOM has columns QUANTITY, LENGTH, and DESCRIPTION.", vbExclamation
            
        CAD.Close SaveChanges:=False
    
        ImportCADSheet = -1
        Exit Function
    End If
    
    SetStatusBar "Importing CAD Sheet..."

    srcSheet.Copy After:=VB_MASTER
    Set DstSheet = ThisWorkbook.Worksheets(srcSheet.Name)
    CAD.Close SaveChanges:=False
    
    DstSheet.Name = "CADWORX_IMPORT"
    
    ' order columns correctly
    If quantity_col <> 1 Then
        DstSheet.Columns(1).Insert
        If quantity_col >= 1 Then quantity_col = quantity_col + 1
        
        DstSheet.Columns(1).Value = DstSheet.Columns(quantity_col).Value2
        DstSheet.Columns(quantity_col).ClearContents
        quantity_col = 1
        If length_col >= 1 Then
            length_col = length_col + 1
        End If
        If desc_col >= 1 Then
            desc_col = desc_col + 1
        End If
    End If
    If length_col <> 2 Then
        DstSheet.Columns(2).Insert
        If length_col >= 2 Then length_col = length_col + 1
        
        DstSheet.Columns(2).Value = DstSheet.Columns(length_col).Value2
        DstSheet.Columns(length_col).ClearContents
        length_col = 2
        If desc_col >= 2 Then
            desc_col = desc_col + 1
        End If
    End If
    If desc_col <> 4 Then
        DstSheet.Columns(4).Insert
        If desc_col >= 4 Then desc_col = desc_col + 1
        
        DstSheet.Columns(4).Value = DstSheet.Columns(desc_col).Value2
        DstSheet.Columns(desc_col).ClearContents
        desc_col = 4
    End If
    DstSheet.Columns(ColLet(5) & ":" & ColLet(8)).ClearContents
    unit_col = 3
    desc_copy_col = 6
    match_col = 7
    new_col = 8
    
    '''''''''''''''''''''''''''''
    ' COLUMN DESCRIPTIONS and ACTIONS FOR THE INCOMING CADWorx BOM
    ''''''''''''''''''''''''''
    'Column 1 - Quantity
    'Column 2 - Length - Delete
    'Column 3 - ALPHA_SIZE <- Change to UNIT
    'Column 4 - LONG_DESC - Delete
    'Column 5 - SHORT_DESC - Delete
    'Column 6 - LONG_DESC - Copy
    'Column 7 - DESCRIPTION_MATCH
    'Column 8 - NEW ITEM Notice
    
    DstSheet.Cells(1, unit_col).Value = "UNIT"
    DstSheet.Cells(1, desc_copy_col).Value = "DESCRIPTION FROM CADWorx"
    DstSheet.Cells(1, match_col).Value = "MATCH CADWorx DESCRIPTION TO ITEM IN MASTER BOM"
    
    Dim importOld As Worksheet
    On Error Resume Next
    Set importOld = ThisWorkbook.Worksheets("CADWORX_IMPORT_OLD")
    Err.Clear
    On Error GoTo 0
    
    'Format and calculate length and unit for each line item
    Dim item_type As String
    Dim rowOld As Integer
    Dim row As Integer
    row = 2
    Do While Not IsEmpty(DstSheet.Cells(row, quantity_col))
        SetStatusBar "Finding Description Matches...", (row - 2), (DstSheet.UsedRange.Rows.CountLarge - 2)
        rowOld = 4
        
        DstSheet.Cells(row, desc_copy_col).Value = TrimWhiteSpace(DstSheet.Cells(row, desc_col).text)
    
        item_type = FirstPhrase(DstSheet.Cells(row, desc_copy_col).Value2)
        'assign temporary unit
        If item_type = "PIPE" Then
            DstSheet.Cells(row, unit_col).Value = "FT"
            DstSheet.Cells(row, quantity_col).Value = CDbl(DstSheet.Cells(row, quantity_col).Value2) * ConvertFeetNInches(DstSheet.Cells(row, length_col).Value2)
            DstSheet.Range("A" & row).NumberFormat = "0.0"
        Else
            DstSheet.Cells(row, quantity_col).Value = val(DstSheet.Cells(row, quantity_col).Value2)
            DstSheet.Cells(row, unit_col).Value = "EA"
        End If
        
        'set up data validation
        With DstSheet.Cells(row, match_col)
            .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
                Operator:=xlBetween, Formula1:="=item_descriptions"
                
            With .FormatConditions.Add(Type:=xlExpression, _
                Formula1:="=DescriptionCompare($" & ColLet(desc_copy_col) & "$" & row & ", $" & ColLet(match_col) & "$" & row & ")<>0")
                .Interior.ColorIndex = 36
            End With
        End With
        
        'find match or best guess
        Dim match_num As Integer
        match_num = -1
        If Not importOld Is Nothing Then
            ' try to match with previous import.
            Do While Not IsEmpty(importOld.Cells(rowOld, 2))
                If DescriptionCompare(importOld.Cells(rowOld, 3).Value2, DstSheet.Cells(row, desc_copy_col).Value2) = 0 Then
                    Exit Do
                End If
                rowOld = rowOld + 1
            Loop
            
            If Not IsEmpty(importOld.Cells(rowOld, 4).Value2) Then
                match_num = MatchDesc(importOld.Cells(rowOld, 4).Value2)
            End If
        End If
        
        If match_num = -1 Then
            match_num = MatchDesc(DstSheet.Cells(row, desc_copy_col).Value2, True)
        End If
        If match_num <> -1 Then
            DstSheet.Cells(row, match_col).Value = get_property(match_num, "Long Description")
        End If
        
        If DescriptionCompare(get_property(match_num, "Long Description"), DstSheet.Cells(row, desc_copy_col).Value2) <> 0 Then
            ' set NEW ITEM notice
            DstSheet.Cells(row, new_col).Value = "*** THIS WILL BECOME NEW ITEM IF NO MATCH IS SPECIFIED ***"
        End If
        
        row = row + 1
    Loop
    
    SetStatusBar "Combining Duplicate Rows...", 1, 1
    
    'handle same item on multiple rows - Combined quantities
    Dim i As Integer
    Dim j As Integer
    With DstSheet.UsedRange
        .Rows(1).Font.Bold = True
        .Columns(new_col).Font.Bold = True
        
        For i = 2 To .Rows.count
            For j = i + 1 To .Rows.count
                If .Cells(i, desc_copy_col).Value2 = .Cells(j, desc_copy_col).Value2 _
                    And Len(.Cells(i, desc_copy_col).Value2) > 0 Then 'same item, but on listed multiple times

                    'compile qualities; delete j-th row
                    .Cells(i, quantity_col).Value = CDbl(.Cells(i, quantity_col).Value2) + CDbl(.Cells(j, quantity_col).Value2)
                    .Rows(j).Delete
                    j = j - 1
                End If
            Next j
        Next i
    
        With .Columns(ColLet(quantity_col) & ":" & ColLet(match_col)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        
        .Columns(match_col).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    End With
    
    SetStatusBar "Finishing...", 1, 1
    
    DstSheet.Columns("E").Delete
    DstSheet.Columns("D").Delete
    DstSheet.Columns("B").Delete
    
    length_col = -1
    desc_copy_col = -1
    quantity_col = 1
    unit_col = 2
    desc_col = 3
    match_col = 4
    new_col = 5
    
    ''''''''''''''''''''''''''''
    ' FINAL COLUMN CONFIGURATION
    '''''''''''''''''''''''''''
    'Column 1 - Quantity
    'Column 2 - UNIT
    'Column 3 - LONG_DESC
    'Column 4 - DESCRIPTION_MATCH
    'Column 5 - NEW ITEM Notice
    
    With DstSheet.Columns(ColLet(quantity_col) & ":" & ColLet(new_col)).Font
        .Name = "Calibri"
        .Size = 11
    End With

    DstSheet.UsedRange.Rows.AutoFit
    
    With DstSheet.UsedRange.Columns(ColLet(quantity_col))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .ColumnWidth = 10
    End With
    With DstSheet.UsedRange.Columns(ColLet(unit_col))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .ColumnWidth = 9
    End With
    With DstSheet.UsedRange.Columns(ColLet(desc_col) & ":" & ColLet(match_col))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .ColumnWidth = 56.43
        .WrapText = True
    End With
    With DstSheet.UsedRange.Columns(ColLet(new_col))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = False
        .ColumnWidth = 8.43
    End With
    
    DstSheet.Rows(1).Insert
    DstSheet.Rows(1).Insert
    DstSheet.Cells(1, 1).Value = "CADWorx Import preformed by " & Environ("Username") & " on " & Now
    DstSheet.Cells(1, 1).HorizontalAlignment = xlLeft
    DstSheet.UsedRange.Columns(match_col).Select
    
    ResetStatusBar
    
    If ui_change Then RenderUI True
    
    Set CAD = Nothing
    Set srcSheet = Nothing
    Set DstSheet = Nothing
End Function

Public Sub FinalizeImport(ByVal site_col As Integer)
'Subroutine: FinalizeImport - Imports the CADWorx sheet into the BOM. Takes care of what was matched and what wasn't.
'                             Adds items that aren't matched. Overwrites site quantites. Deletes CADWorx import sheet
'Arguments: site_col - String containing the site name where this BOM will be imported


    Dim row As Integer
    
    Dim first As Integer
    Dim last As Integer
    first = VB_MASTER.FirstRow()
    last = VB_MASTER.LastRow()
    
    SetStatusBar "Finalizing CADWorx Import...", 0, 1
    
    Dim cadSheet As Worksheet
    Set cadSheet = ThisWorkbook.Sheets("CADWORX_IMPORT")
    
    'LOG CHANGE
    VB_CHANGE_LOG.LogChange VB_MASTER.Name & "!" & VB_MASTER.Cells(VB_MASTER.SubtitleRow(), site_col).Address, "", "BEGIN CADWorx Import", "", ""
    
    Dim ImportedMarks As Collection
    Set ImportedMarks = New Collection
    
    'first row
    row = 4
    
    Dim org_quantity As Integer
    Dim import_quantity As Integer
    Dim import_desc As String
    Dim new_mark_num As Integer

    Dim ui_change As Boolean
    ui_change = RenderUI(False)

    '''''''''''''''''''''''''''
    ' REMINDER OF COLUMN CONFIGURATION
    '''''''''''''''''''''''''''
    'Column 1 - Quantity
    'Column 2 - UNIT
    'Column 3 - LONG_DESC
    'Column 4 - DESCRIPTION_MATCH
    'Column 5 - NEW ITEM Notice
    
    Do While Not IsEmpty(cadSheet.Cells(row, 1))
        SetStatusBar "Copying Material Quantites...", (row - 4), (cadSheet.UsedRange.Rows.CountLarge - 4)
        
        If Not IsEmpty(cadSheet.Cells(row, 4)) Then
            import_desc = cadSheet.Cells(row, 4).Value2
        Else
            import_desc = cadSheet.Cells(row, 3).Value2
        End If
        
        'Match Description with mark_num. if mark num doesn't exist, loads AddItemForm. User
        'chooses a category and item is added.
        new_mark_num = MatchDesc(import_desc)
        If new_mark_num = -1 Then
            Dim first_phrase As String
            Dim category As String
            first_phrase = FirstPhrase(import_desc)
            category = VB_CATEGORY.FindCategory(first_phrase)
            
            If Len(category) > 0 Then
                new_mark_num = InsertItemBOM(category, import_desc)
            Else
                new_mark_num = AddItemForm.LoadForm(import_desc)
            End If
            
            If new_mark_num = -1 Then
                'Handle Error
                GoTo NEXTITEM
            End If
        End If
        
        ImportedMarks.Add new_mark_num
        
        'Round up values
        import_quantity = VBRoundUp(cadSheet.Cells(row, 1).Value2)
        org_quantity = VB_MASTER.Cells(get_row(new_mark_num), site_col).Value2
        
        'Check if site quantity is different and assign updated quantity
        If org_quantity <> import_quantity Then
            VB_MASTER.Cells(get_row(new_mark_num), site_col).Value = import_quantity
            Call VB_MASTER.WriteChange(VB_MASTER.Cells(get_row(new_mark_num), site_col), IIf(org_quantity = 0, "", org_quantity))
        End If
        
NEXTITEM:
        row = row + 1
    Loop
    
    SetStatusBar "Finishing...", 1, 1
    
    'DON'T delete CADWorx import sheet -- save for Drafting PunchList
    'Application.DisplayAlerts = False
    'cadSheet.Delete
    'Application.DisplayAlerts = True

    If SheetExists("CADWORX_IMPORT_OLD") Then
        ThisWorkbook.Sheets("CADWORX_IMPORT_OLD").Visible = xlSheetVisible
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets("CADWORX_IMPORT_OLD").Delete
        Application.DisplayAlerts = True
    End If
    
    cadSheet.Name = "CADWORX_IMPORT_OLD"
    cadSheet.Visible = xlSheetVeryHidden
    cadSheet.UsedRange.Validation.Delete
    
    'clear contents for mark numbers that were not included in the import
    For row = first To last
        If Not CollectionContains(ImportedMarks, get_mark_num(row)) And Not IsEmpty(VB_MASTER.Cells(row, site_col)) Then
            org_quantity = VB_MASTER.Cells(row, site_col).Value2
            VB_MASTER.Cells(row, site_col).ClearContents
            Call VB_MASTER.WriteChange(VB_MASTER.Cells(row, site_col), org_quantity)
        End If
    Next row
    
    'LOG CHANGE
    VB_CHANGE_LOG.LogChange VB_MASTER.Name & "!" & VB_MASTER.Cells(VB_MASTER.SubtitleRow(), site_col).Address, "", "END CADWorx Import", "", ""
    
    VB_MASTER.CalculateQuantityFormat
    VB_MASTER.Activate
    
    ResetStatusBar
    
    If ui_change Then RenderUI True
    
    Set ImportedMarks = Nothing
    Set cadSheet = Nothing
End Sub

Public Sub DraftingPunchList() ' Model Discrepancy Report
'Subroutine: DraftingPunchlist - Creates a Model Discrepancy Report. The item descriptions in the model should match the BOM.
'                               This is common sense, and it makes CADWorx BOM imports easier if the material descriptions match.
    
    
    If Not SheetExists("CADWORX_IMPORT_OLD") Then
        Exit Sub
    End If
    
    Dim cadSheet As Worksheet
    Set cadSheet = ThisWorkbook.Sheets("CADWORX_IMPORT_OLD")

    'transfer discrepancies
    Dim cad_row As Integer
    cad_row = 4 ' first row
    
    Dim mdr_row As Integer
    mdr_row = VB_MDR.FirstRow()
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    Do While Not IsEmpty(cadSheet.Cells(cad_row, 3))
        If DescriptionCompare(cadSheet.Cells(cad_row, 3).Value2, cadSheet.Cells(cad_row, 4).Value2) <> 0 _
            And Len(cadSheet.Cells(cad_row, 3).Value2) > 0 _
            And Len(cadSheet.Cells(cad_row, 4).Value2) > 0 Then
            
            If Not IsEmpty(VB_MDR.Cells(mdr_row, 1)) Then
                VB_MDR.Rows(mdr_row + 1).Insert
                mdr_row = mdr_row + 1
                With VB_MDR.Range("A" & mdr_row & ":C" & mdr_row).Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
            End If
            
            VB_MDR.Cells(mdr_row, 1).Value = cadSheet.Cells(cad_row, 3).Value2
            VB_MDR.Cells(mdr_row, 3).Value = cadSheet.Cells(cad_row, 4).Value2
        End If
        
        cad_row = cad_row + 1
    Loop
    
    VB_MDR.SetBasis
    VB_MDR.SetCreationDate
    VB_MDR.SetCreater
    
    'print
    VB_MDR.Visible = xlSheetVisible
    Call Print2PDF(VB_MDR.Name, VB_VAR_STORE.GetMDRDirectory(), "Model_Discrepancy_Report" & HMMFileTag())
    VB_MDR.Visible = xlSheetVeryHidden
    
    'clear
    mdr_row = VB_MDR.FirstRow()
    
    Do While Not IsEmpty(VB_MDR.Cells(mdr_row + 1, 1))
        VB_MDR.Rows(mdr_row).Delete
    Loop
    VB_MDR.Rows(mdr_row).ClearContents
    
    If ui_change Then RenderUI True
    
    Set cadSheet = Nothing
End Sub
