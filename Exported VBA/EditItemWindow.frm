VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EditItemWindow 
   Caption         =   "Material Properties"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12000
   OleObjectBlob   =   "EditItemWindow.frx":0000
End
Attribute VB_Name = "EditItemWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Form: EditItemWindow                                 '
'                                                           '
' Editing Material Properties takes place in this form.     '
' Double clicking on an item in VB_MASTER loads this form.  '
' The user has the option to scroll through the items using '
' 'Next' and 'Previous' buttons. To save changes user click '
' 'Save' button. To Close, click 'Close' button.            '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private mark_num As Integer
Private num_sites As Integer

Dim overall_change As Boolean
Dim TargetColumn As Integer

Private loaded As Boolean

Private SiteControlCollection As Collection

Public Sub LoadForm(ByVal Target As Range)
'Subroutine: LoadForm - Loads form with material properties extracted from Target.Row. Sets focus
'                       corresponding to Target.Column.
'Arguments: Target - Range object reference to one cell corresponding to an item in VB_MASTER.


    overall_change = False
    If Not loaded Then
        loaded = False
    End If
    
    Dim master_row As Integer
    
    'get mark number
    mark_num = get_mark_num(Target.row)
    master_row = Target.row
    TargetColumn = Target.column
    
    If mark_num = -1 Then
        Exit Sub
    End If
    
    If Not loaded Then
        Set SiteControlCollection = New Collection
    End If
    
    Dim site As Integer
    Dim first_site_col As Integer
    Dim site_name As String
    num_sites = GetSitesRange().count
    first_site_col = get_col_num("Current Model Quantities")
    
    'Control Model Extras Visibility
    ORG_ExtrasFrame.Visible = Not VB_MASTER.Columns(get_col_num("Total Extras")).Hidden
    ExtrasFrame.Visible = ORG_ExtrasFrame.Visible
    
    'Control Mark Number Suppress
    ORG_MarkNumLabel.Visible = Not VB_MASTER.Columns(get_col_num("Mark No.")).Hidden
    ORG_Mark_TextBox.Visible = ORG_MarkNumLabel.Visible
    MarkNumLabel.Visible = ORG_MarkNumLabel.Visible
    Mark_TextBox.Visible = ORG_MarkNumLabel.Visible
    
    'set initial focus
    MultiPageFrame.Value = 1
    Description_TextBox.SetFocus
    
    'loop through sites, add DynamicSiteBox, and set value
    If num_sites > 0 Then
        ORG_NoSites_Label.Visible = False
        NoSites_Label.Visible = False
        
        ORG_NoExtras_Label.Visible = False
        NoExtras_Label.Visible = False
        
        Dim ORG_siteBox As MSForms.TextBox
        Dim ORG_siteLabel As MSForms.Label
        Dim SiteBox As MSForms.TextBox
        Dim siteLabel As MSForms.Label
        Dim newSiteBox As DynamicSiteBox
        
        Dim ORG_extraBox As MSForms.TextBox
        Dim ORG_extraLabel As MSForms.Label
        Dim ExtraBox As MSForms.TextBox
        Dim ExtraLabel As MSForms.Label
        Dim newExtraBox As DynamicSiteBox
        
        Dim org_frame_width As Integer

        Dim previousLeft As Integer
        previousLeft = 3

        For site = first_site_col To first_site_col + num_sites - 1
            site_name = VB_MASTER.Cells(VB_MASTER.SubtitleRow(), site).Value2
            
            '''''''''''''''''''''''''''''
            ' Model Quantities
            
            If Not loaded Then
                'ORGINAL Site Label
                Set ORG_siteLabel = ORG_DFTNGFrame.Controls.Add("Forms.Label.1", "ORG_" & site_name & "_Label", True)
                ORG_siteLabel.Top = 16
                ORG_siteLabel.height = 12
                ORG_siteLabel.Width = VB_MASTER.Cells(VB_MASTER.SubtitleRow(), site).ColumnWidth * 54 / 11.14
                ORG_siteLabel.Left = previousLeft + 3
                ORG_siteLabel.TextAlign = fmTextAlignCenter
                ORG_siteLabel.Font.Size = 8
                ORG_siteLabel.Font.Bold = True
                ORG_siteLabel.Caption = site_name
                
                'Site Label
                Set siteLabel = DFTNGFrame.Controls.Add("Forms.Label.1", site_name & "_Label", True)
                siteLabel.Top = 16
                siteLabel.height = 12
                siteLabel.Width = ORG_siteLabel.Width
                siteLabel.Left = previousLeft + 3
                siteLabel.TextAlign = fmTextAlignCenter
                siteLabel.Font.Size = 8
                siteLabel.Font.Bold = True
                siteLabel.Caption = site_name
            End If
            
            'ORINGAL Site TextBox
            If Not loaded Then
                Set ORG_siteBox = ORG_DFTNGFrame.Controls.Add("Forms.TextBox.1", "ORG_" & site_name & "_Box", True)
                ORG_siteBox.Top = 34
                ORG_siteBox.height = 18
                ORG_siteBox.Width = ORG_siteLabel.Width
                ORG_siteBox.Left = previousLeft + 3
                ORG_siteBox.SpecialEffect = fmSpecialEffectEtched
                ORG_siteBox.TextAlign = fmTextAlignRight
            Else
                Set ORG_siteBox = ORG_DFTNGFrame.Controls.Item("ORG_" & site_name & "_Box")
            End If
            ORG_siteBox.Value = VB_MASTER.Cells(master_row, site).Value2
            ORG_siteBox.enabled = False
            
            'Site TextBox
            If Not loaded Then
                Set SiteBox = DFTNGFrame.Controls.Add("Forms.TextBox.1", site_name & "_Box", True)
                SiteBox.Top = 34
                SiteBox.height = 18
                SiteBox.Width = siteLabel.Width
                SiteBox.Left = previousLeft + 3
                SiteBox.SpecialEffect = fmSpecialEffectEtched
                SiteBox.TextAlign = fmTextAlignRight
                SiteBox.Value = ORG_siteBox.Value
                Set newSiteBox = New DynamicSiteBox
                Set newSiteBox.SetTextBox = SiteBox
                SiteControlCollection.Add newSiteBox
            Else
                DFTNGFrame.Controls.Item(site_name & "_Box").Value = ORG_siteBox.Value
            End If
            If TargetColumn = site Then
                DFTNGFrame.Controls.Item(site_name & "_Box").SetFocus
            End If
            
            '''''''''''''''''''''''''''''
            ' Model Extras
            
            If Not loaded Then
                'ORGINAL Extras Label
                Set ORG_extraLabel = ORG_ExtrasFrame.Controls.Add("Forms.Label.1", "ORG_" & site_name & "Extras_Label", True)
                ORG_extraLabel.Top = 16
                ORG_extraLabel.height = 12
                ORG_extraLabel.Width = VB_MASTER.Cells(VB_MASTER.SubtitleRow(), site).ColumnWidth * 54 / 11.14
                ORG_extraLabel.Left = previousLeft + 3
                ORG_extraLabel.TextAlign = fmTextAlignCenter
                ORG_extraLabel.Font.Size = 8
                ORG_extraLabel.Font.Bold = True
                ORG_extraLabel.Caption = site_name
            
                'Extras Label
                Set ExtraLabel = ExtrasFrame.Controls.Add("Forms.Label.1", site_name & "Extras_Label", True)
                ExtraLabel.Top = 16
                ExtraLabel.height = 12
                ExtraLabel.Width = ORG_extraLabel.Width
                ExtraLabel.Left = previousLeft + 3
                ExtraLabel.TextAlign = fmTextAlignCenter
                ExtraLabel.Font.Size = 8
                ExtraLabel.Font.Bold = True
                ExtraLabel.Caption = site_name
            End If
            
            'ORINGAL Extras TextBox
            If Not loaded Then
                Set ORG_extraBox = ORG_ExtrasFrame.Controls.Add("Forms.TextBox.1", "ORG_" & site_name & "Extras_Box", True)
                ORG_extraBox.Top = 34
                ORG_extraBox.height = 18
                ORG_extraBox.Width = ORG_extraLabel.Width
                ORG_extraBox.Left = previousLeft + 3
                ORG_extraBox.SpecialEffect = fmSpecialEffectEtched
                ORG_extraBox.TextAlign = fmTextAlignRight
            Else
                Set ORG_extraBox = ORG_ExtrasFrame.Controls.Item("ORG_" & site_name & "Extras_Box")
            End If
            ORG_extraBox.Value = VB_MASTER.Cells(master_row, site + num_sites).Value2
            ORG_extraBox.enabled = False
            
            'Extras TextBox
            If Not loaded Then
                Set ExtraBox = ExtrasFrame.Controls.Add("Forms.TextBox.1", site_name & "Extras_Box", True)
                ExtraBox.Top = 34
                ExtraBox.height = 18
                ExtraBox.Width = ExtraLabel.Width
                ExtraBox.Left = previousLeft + 3
                ExtraBox.SpecialEffect = fmSpecialEffectEtched
                ExtraBox.TextAlign = fmTextAlignRight
                ExtraBox.Value = ORG_extraBox.Value
                Set newExtraBox = New DynamicSiteBox
                Set newExtraBox.SetTextBox = ExtraBox
                SiteControlCollection.Add newExtraBox
            Else
                ExtrasFrame.Controls.Item(site_name & "Extras_Box").Value = ORG_extraBox.Value
            End If
            If TargetColumn = site + num_sites And ExtrasFrame.Visible Then
                ExtrasFrame.Controls.Item(site_name & "Extras_Box").SetFocus
            End If
            
            If Not loaded Then
                previousLeft = siteLabel.Left + siteLabel.Width
                
                'Resize Frame / Form
                org_frame_width = DFTNGFrame.Width
                If siteLabel.Left + siteLabel.Width > org_frame_width Then
                    DFTNGFrame.Width = siteLabel.Left + siteLabel.Width + 9
                    ORG_DFTNGFrame.Width = siteLabel.Left + siteLabel.Width + 9
                    
                    ExtrasFrame.Width = ExtraLabel.Left + ExtraLabel.Width + 9
                    ORG_ExtrasFrame.Width = ExtraLabel.Left + ExtraLabel.Width + 9
                End If
                
                If DFTNGFrame.Width > 250 Then
                    If DFTNGFrame.ScrollWidth <= 250 Then
                        DFTNGFrame.ScrollWidth = DFTNGFrame.Width
                        ORG_DFTNGFrame.ScrollWidth = ORG_DFTNGFrame.Width
                        
                        ExtrasFrame.ScrollWidth = ExtrasFrame.Width
                        ORG_ExtrasFrame.ScrollWidth = ORG_ExtrasFrame.Width
                    Else
                        DFTNGFrame.ScrollWidth = DFTNGFrame.ScrollWidth + siteLabel.Width + 9
                        ORG_DFTNGFrame.ScrollWidth = ORG_DFTNGFrame.ScrollWidth + ORG_siteLabel.Width + 9
                        
                        ExtrasFrame.ScrollWidth = ExtrasFrame.ScrollWidth + ExtraLabel.Width + 9
                        ORG_ExtrasFrame.ScrollWidth = ORG_ExtrasFrame.ScrollWidth + ORG_extraLabel.Width + 9
                    End If
                    
                    DFTNGFrame.Width = 250
                    DFTNGFrame.ScrollBars = fmScrollBarsHorizontal
                    ORG_DFTNGFrame.Width = 250
                    ORG_DFTNGFrame.ScrollBars = fmScrollBarsHorizontal
                    
                    ExtrasFrame.Width = 250
                    ExtrasFrame.ScrollBars = fmScrollBarsHorizontal
                    ORG_ExtrasFrame.Width = 250
                    ORG_ExtrasFrame.ScrollBars = fmScrollBarsHorizontal
                End If
                
                Me.Width = Me.Width + DFTNGFrame.Width - org_frame_width
                MultiPageFrame.Width = MultiPageFrame.Width + DFTNGFrame.Width - org_frame_width
                
                SaveButton.Left = SaveButton.Left + DFTNGFrame.Width - org_frame_width
                CancelButton.Left = CancelButton.Left + DFTNGFrame.Width - org_frame_width
            End If
            
            Set ORG_siteBox = Nothing
            Set ORG_siteLabel = Nothing
            Set SiteBox = Nothing
            Set siteLabel = Nothing
            Set newSiteBox = Nothing
            
            Set ORG_extraBox = Nothing
            Set ORG_extraLabel = Nothing
            Set ExtraBox = Nothing
            Set ExtraLabel = Nothing
            Set newExtraBox = Nothing
        Next site
        
    End If
    
    'SET TEXT BOX VALUES, REFINE FOCUS
    Dim col_title As String
    col_title = get_ColTitle(Target.column)
    
    ORG_Surplus_TextBox.Value = get_property(mark_num, "Client Inventory")
    Surplus_TextBox.Value = ORG_Surplus_TextBox.Value
    If col_title = "Client Inventory" Then
        Surplus_TextBox.SetFocus
    End If
    
    ORG_LongLeadCheck.Value = CBool(get_property(mark_num, "Long Lead"))
    LongLeadCheck.Value = ORG_LongLeadCheck.Value
    
    ORG_Mark_TextBox.Value = get_property(mark_num, "Mark No.")
    Mark_TextBox.Value = ORG_Mark_TextBox.Value
    Mark_TextBox.enabled = Not ThisWorkbook.MarkNumLock()
    
    ORG_SAPBox.Value = get_property(mark_num, "SAP#")
    SAPBox.Value = ORG_SAPBox.Value
    If col_title = "SAP#" Then
        SAPBox.SetFocus
    End If
    
    ORG_Unit_TextBox.Value = get_property(mark_num, "Unit")
    Unit_TextBox.Value = ORG_Unit_TextBox.Value
    
    ORG_Description_TextBox.Value = get_property(mark_num, "Long Description")
    Description_TextBox.Value = ORG_Description_TextBox.Value
    If col_title = "Long Description" Then
        Description_TextBox.SetFocus
    End If
    
    ORG_DescCheck_CheckBox.Value = (Len(get_property(mark_num, "Description Check")) > 0)
    DescCheck_CheckBox.Value = ORG_DescCheck_CheckBox.Value
    
    ORG_Delete_CheckBox.Value = (Len(get_property(mark_num, "Delete?")) > 0)
    Delete_CheckBox.Value = ORG_Delete_CheckBox.Value
    
    'HANDLE DELETE REASON
    ORG_Reason_Label.Visible = Delete_CheckBox.Value
    ORG_Reason_TextBox.Visible = Delete_CheckBox.Value
    Reason_Label.Visible = Delete_CheckBox.Value
    Reason_TextBox.Visible = Delete_CheckBox.Value
    
    If Delete_CheckBox.Value Then
        Dim c As Integer
        Dim desc_col As Integer
        Dim DESC As String
        desc_col = get_col_num("Long Description")
        DESC = Description_TextBox.Value
        ' find start position of user comment (in RED)
        For c = Len(DESC) To 1 Step -1
            With VB_MASTER.Cells(master_row, desc_col).Characters(start:=c, length:=1).Font
                If .ColorIndex <> 3 Then
                    Exit For
                End If
            End With
        Next c
        
        ' separate reason for deletion from description.
        ORG_Description_TextBox.Value = Left(DESC, c)
        Description_TextBox.Value = ORG_Description_TextBox.Value
        
        ORG_Reason_TextBox.Value = Right(DESC, Len(DESC) - c - 1)
        Reason_TextBox.Value = ORG_Reason_TextBox.Value
    End If
    
    'category
    Dim index As Integer
    Dim list_index As Integer
    Dim cat_row As Integer
    Dim category As String
    Dim cat_col As Integer
    index = 0
    list_index = index
    cat_row = VB_CATEGORY.FirstRow()
    cat_col = VB_CATEGORY.CategoryColumn()
    category = get_category(master_row)
    
    Do While Not VB_CATEGORY.EOF(cat_row)
        CategoryBox.AddItem VB_CATEGORY.Cells(cat_row, cat_col).Value2
        If category = VB_CATEGORY.Cells(cat_row, cat_col).Value2 Then
            list_index = index
        End If
        
        index = index + 1
        cat_row = cat_row + 1
    Loop
    ORG_CategoryBox.Value = category
    If master_row > 0 Then
        CategoryBox.ListIndex = list_index
    Else
        CategoryBox.ListIndex = 0
    End If
    
    ' ADD/VIEW/EDIT NOTES BUTTON
    If HasNote(mark_num) Then
        AddEditNoteButton.Caption = "View/Edit Notes"
    Else
        AddEditNoteButton.Caption = "Add  Notes"
    End If
    
    'ADJUST POSITION ON SCREEN
    If Not loaded Then
        Me.Top = Application.Top + Application.height / 2 - Me.height / 2
        Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    End If
    
    ' HANDLE SAP# on Edit Item Window
    ORG_SAPFrame.Visible = Not VB_MASTER.Columns(get_col_num("SAP#")).Hidden
    SAPFrame.Visible = ORG_SAPFrame.Visible
    
    'CORRECT FOR MARK NUM
    If mark_num = LastMark() Then
        NextButton.enabled = False
        PreviousButton.enabled = True
    ElseIf master_row = VB_MASTER.FirstRow() Then
        NextButton.enabled = True
        PreviousButton.enabled = False
    Else
        NextButton.enabled = True
        PreviousButton.enabled = True
    End If
    
    'Show
    If Not loaded Then
        loaded = True
        Me.Show
    End If
End Sub

Private Sub NextButton_Click()
'Subroutine: NextButton_Click - User form event handler. Called when the user click NextButton.
'                               Loads form with next item in VB_MASTER.
    
    Dim next_row As Integer
    next_row = get_row(mark_num) + 1
    
    Do While Len(get_property(get_mark_num(next_row), "Long Description")) = 0
        next_row = next_row + 1
        If next_row > VB_MASTER.LastRow() Then
            next_row = get_row(mark_num)
            Exit Do
        End If
    Loop
    
    LoadForm VB_MASTER.Cells(next_row, TargetColumn)
    
    If loaded Then UpdateSaveButtonColor
End Sub

Private Sub PreviousButton_Click()
'Subroutine: PreviousButton_Click - User form event handler. Called when the user click PreviousButton.
'                                   Loads form with previous item in VB_MASTER.

    Dim next_row As Integer
    next_row = get_row(mark_num) - 1
    
    Do While Len(get_property(get_mark_num(next_row), "Long Description")) = 0
        next_row = next_row - 1
        If next_row < VB_MASTER.FirstRow() Then
            next_row = get_row(mark_num)
            Exit Do
        End If
    Loop
    
    LoadForm VB_MASTER.Cells(next_row, TargetColumn)
    
    If loaded Then UpdateSaveButtonColor
End Sub

Private Sub SaveButton_Click()
'Subroutine: SaveButton_Click - User form event handler. Called when the user clicks SaveButton.
'                               Checks for changes from original properties. If changed, WriteChange

    Dim ErrControl As MSForms.TextBox
    Set ErrControl = Nothing

    Dim ui_change As Boolean
    ui_change = False
    
    'get VB_MASTER row
    Dim mst_row As Integer
    mst_row = get_row(mark_num)
    
    'check for changes
    If overall_change Then
        ui_change = RenderUI(False)
    Else
        GoTo EXIT_SUB
    End If
    
    'check Surplus box
    If Surplus_TextBox.Value <> ORG_Surplus_TextBox.Value Then
        If Len(Surplus_TextBox.Value > 0) Then
            If Not IsNumeric(Surplus_TextBox.Value) Then
                MsgBox "Client Inventory value must be numeric.", vbExclamation
                Set ErrControl = Surplus_TextBox
                GoTo EXIT_SUB
            ElseIf Fix(Surplus_TextBox.Value) <> CDbl(Surplus_TextBox.Value) Or CInt(Surplus_TextBox.Value) <= 0 Then
                MsgBox "Client Inventory value must be positive natural number.", vbExclamation
                Set ErrControl = Surplus_TextBox
                GoTo EXIT_SUB
            End If
        End If
        
        set_property mark_num, get_col_num("Client Inventory"), Surplus_TextBox.Value
        Call VB_MASTER.WriteChange(VB_MASTER.Cells(mst_row, get_col_num("Client Inventory")), ORG_Surplus_TextBox.Value)
    End If
    
    'check Long Lead box
    If LongLeadCheck.Value <> ORG_LongLeadCheck.Value Then
        set_property mark_num, get_col_num("Long Lead"), LongLeadCheck.Value
        Call VB_MASTER.WriteChange(VB_MASTER.Cells(mst_row, get_col_num("Long Lead")), ORG_LongLeadCheck.Value)
    End If
    
    'check Delete box
    If Delete_CheckBox.Value <> ORG_Delete_CheckBox.Value Then
        If Delete_CheckBox.Value And Len(Reason_TextBox.Value) > 0 Then
            ClientDeleteItem VB_MASTER.Cells(get_row(mark_num), get_col_num("Delete?")), Reason_TextBox.Value
        Else
            ClientUndeleteItem VB_MASTER.Cells(get_row(mark_num), get_col_num("Delete?"))
        End If
    End If
    
    'check Delete reason
    If Reason_TextBox.Value <> ORG_Reason_TextBox.Value Then
        Dim this_reason As String
        this_reason = UCase(TrimWhiteSpace(Reason_TextBox.Value))
        
        'check for illegal characters in reason
        If Len(this_reason) = 0 Then
            MsgBox "You must enter a reason for deleting this item.", vbExclamation
            Set ErrControl = Reason_TextBox
            GoTo EXIT_SUB
        End If
        
        'check for illegal characters in reason
        If ContainsIllegalCharacters(this_reason) Then
            MsgBox "You entered an illegal character. Do not use characters " & vbCrLf & vbCrLf & "* _ [ ] ^", vbExclamation
            Set ErrControl = Reason_TextBox
            GoTo EXIT_SUB
        End If
        
        SetDeleteReason mark_num, this_reason
    End If
    
    'if description change, then uncheck Description Approval
    Dim desc_change As Boolean
    desc_change = False
    
    'check Category Box
    If CategoryBox.Value <> ORG_CategoryBox.Value Then
        desc_change = True
        mst_row = AutoSortItem(mark_num, CategoryBox.Value)
        If VB_CHANGE_LOG.TrackChanges() Then
            Call VB_CHANGE_LOG.LogChange(VB_MASTER.Name & "!$" & ColLet(VB_MASTER.CategoryColumn()) & "$" & get_row(mark_num), mark_num, _
                "Category Change", ORG_CategoryBox.Value, CategoryBox.Value)
        End If
    End If
    
    'check Unit Box
    If Unit_TextBox.Value <> ORG_Unit_TextBox.Value Then
        Call set_property(mark_num, get_col_num("Unit"), Unit_TextBox.Value)
    End If
    
    'check Description Box
    If Description_TextBox.Value <> ORG_Description_TextBox.Value Then
        Description_TextBox.Value = TrimWhiteSpace(CStr(Description_TextBox.Value))
        
        If Len(Description_TextBox.Value) = 0 Then
            MsgBox "Description is invalid. Please type something.", vbExclamation
            Set ErrControl = Description_TextBox
            GoTo EXIT_SUB
        End If
        
        'check for illegal characters in description
        If ContainsIllegalCharacters(Description_TextBox.Value) Then
            MsgBox "You entered an illegal character. Do not use characters " & vbCrLf & vbCrLf & "* _ [ ] ^", vbExclamation
            Set ErrControl = Description_TextBox
            GoTo EXIT_SUB
        End If
        
        desc_change = True
    
        Dim reason As String
        Dim new_desc As String
        
        'handle adding delete reason
        If Delete_CheckBox.Value Then
            reason = GetDeleteReason(mark_num)
            set_property mark_num, get_col_num("Long Description"), Description_TextBox.Value
            SetDeleteReason mark_num, reason
        Else
            set_property mark_num, get_col_num("Long Description"), Description_TextBox.Value
        End If
        
        'AUTOSORT
        mst_row = AutoSortItem(mark_num)
        
        Call VB_MASTER.WriteChange(VB_MASTER.Cells(mst_row, get_col_num("Long Description")), ORG_Description_TextBox.Value)
    End If
    
    If DescCheck_CheckBox.Value And Not desc_change Then
        ApproveDescription MarkNum:=mark_num
    Else
        UnapproveDescription MarkNum:=mark_num
    End If
                     
    'check for changes to site/extras quantities
    Dim site As Integer
    Dim site_name As String
    Dim first_site_col As Integer
    first_site_col = get_col_num("Current Model Quantities")
    
    For site = first_site_col To num_sites + first_site_col - 1
        site_name = VB_MASTER.Cells(VB_MASTER.SubtitleRow(), site).Value2
        
        'if change in site quantity
        If DFTNGFrame.Controls.Item(site_name & "_Box").Value <> ORG_DFTNGFrame.Controls.Item("ORG_" & site_name & "_Box").Value Then
            If Len(DFTNGFrame.Controls.Item(site_name & "_Box").Value) > 0 Then
                If Not IsNumeric(DFTNGFrame.Controls.Item(site_name & "_Box").Value) Then
                    MsgBox "Site Quantities must be numeric.", vbExclamation
                    Set ErrControl = DFTNGFrame.Controls.Item(site_name & "_Box")
                    GoTo EXIT_SUB
                ElseIf Fix(DFTNGFrame.Controls.Item(site_name & "_Box").Value) <> CDbl(DFTNGFrame.Controls.Item(site_name & "_Box").Value) _
                    Or CInt(DFTNGFrame.Controls.Item(site_name & "_Box").Value) <= 0 Then
                    
                    MsgBox "Site Quantities must be positive natural numbers.", vbExclamation
                    Set ErrControl = DFTNGFrame.Controls.Item(site_name & "_Box")
                    GoTo EXIT_SUB
                End If
            End If
            
            set_property mark_num, site, DFTNGFrame.Controls.Item(site_name & "_Box").Value
            Call VB_MASTER.WriteChange(VB_MASTER.Cells(mst_row, site), ORG_DFTNGFrame.Controls.Item("ORG_" & site_name & "_Box").Value)
        End If
        
        'if change in extras quantity
        If ExtrasFrame.Controls.Item(site_name & "Extras_Box").Value <> ORG_ExtrasFrame.Controls.Item("ORG_" & site_name & "Extras_Box").Value Then
            If Len(ExtrasFrame.Controls.Item(site_name & "Extras_Box").Value) > 0 Then
                If Not IsNumeric(ExtrasFrame.Controls.Item(site_name & "Extras_Box").Value) Then
                    MsgBox "Extras Quantities must be numeric.", vbExclamation
                    Set ErrControl = ExtrasFrame.Controls.Item(site_name & "Extras_Box")
                    GoTo EXIT_SUB
                ElseIf Fix(ExtrasFrame.Controls.Item(site_name & "Extras_Box").Value) <> CDbl(ExtrasFrame.Controls.Item(site_name & "Extras_Box").Value) _
                    Or CInt(ExtrasFrame.Controls.Item(site_name & "Extras_Box").Value) <= 0 Then
                    
                    MsgBox "Extras Quantities must be positive natural numbers.", vbExclamation
                    Set ErrControl = ExtrasFrame.Controls.Item(site_name & "Extras_Box")
                    GoTo EXIT_SUB
                End If
            End If
            
            set_property mark_num, site + num_sites, ExtrasFrame.Controls.Item(site_name & "Extras_Box").Value
            Call VB_MASTER.WriteChange(VB_MASTER.Cells(mst_row, site + num_sites), ORG_ExtrasFrame.Controls.Item("ORG_" & site_name & "Extras_Box").Value)
        End If
    Next site
    
    'check MarkNumber box
    If Mark_TextBox.Value <> ORG_Mark_TextBox.Value And Not ThisWorkbook.MarkNumLock() Then
        If Not IsNumeric(Mark_TextBox.Value) Or Len(Mark_TextBox.Value) = 0 Then
            MsgBox "Mark Number value must be numeric.", vbExclamation
            Set ErrControl = Mark_TextBox
            GoTo EXIT_SUB
        ElseIf Fix(Mark_TextBox.Value) <> CDbl(Mark_TextBox.Value) Or CInt(Mark_TextBox.Value) <= 0 Then
            MsgBox "Mark Number value must be a positive natural number.", vbExclamation
            Set ErrControl = Mark_TextBox
            GoTo EXIT_SUB
        End If
        If get_row(Mark_TextBox.Value) <> -1 Then
            MsgBox "Mark Number already exists in BOM. Cannot change to this.", vbExclamation
            Set ErrControl = Mark_TextBox
            GoTo EXIT_SUB
        End If
        
        VB_MASTER.Cells(mst_row, get_col_num("Mark No.")).Value = CInt(Mark_TextBox.Value)
        Call VB_MASTER.WriteChange(VB_MASTER.Cells(mst_row, get_col_num("Mark No.")), ORG_Mark_TextBox.Value)
    End If
    
    'check SAPbox
    If SAPBox.Value <> ORG_SAPBox.Value And SAPBox.Visible Then
        VB_MASTER.Cells(mst_row, get_col_num("SAP#")).Value = SAPBox.Value
        Call VB_MASTER.WriteChange(VB_MASTER.Cells(mst_row, get_col_num("SAP#")), ORG_SAPBox.Value)
    End If
    
EXIT_SUB:
    LoadForm VB_MASTER.Cells(mst_row, TargetColumn)

    If Not ErrControl Is Nothing Then
        With ErrControl
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.text)
        End With
        Set ErrControl = Nothing
    End If

    VB_MASTER.CalculateQuantityFormat
    
    If loaded Then UpdateSaveButtonColor
    
    If ui_change Then RenderUI True
End Sub

Private Sub CancelButton_Click()
'Subroutine: CancelButton_Click - User form event handler. Called when the user clicks CancelButton.

    Unload Me
End Sub

Private Sub AddEditNoteButton_Click()
    AddEditNotes.LoadForm mark_num
End Sub

Private Sub RenameCategoryButton_Click()
'Subroutine: RenameCategoryButton_Click - UserForm event handler. Called when user clicks NewCategoryButton.
'                                         Runs routine to rename a category, and refreshes category list for user to choose from.


    Dim new_cat As String
    RenameCategoryForm.LoadForm CategoryBox.Value, new_cat

    If new_cat = vbNullString Then
        Exit Sub
    End If

    ORG_CategoryBox.Value = new_cat

    Do While CategoryBox.ListCount > 0
        CategoryBox.RemoveItem CategoryBox.ListCount - 1
    Loop
    
    Dim row As Integer
    Dim cat_col As Integer
    Dim index As Integer
    Dim list_index As Integer
    row = VB_CATEGORY.FirstRow()
    list_index = 0
    index = 0
    cat_col = VB_CATEGORY.CategoryColumn()
    
    Do While Not VB_CATEGORY.EOF(row)
        CategoryBox.AddItem VB_CATEGORY.Cells(row, cat_col).Value2
        If CStr(VB_CATEGORY.Cells(row, cat_col).Value2) = new_cat Then
            list_index = index
        End If
        
        index = index + 1
        row = row + 1
    Loop
    
    'set best guess
    CategoryBox.ListIndex = list_index
End Sub

Private Sub Surplus_TextBox_Change()
'Subroutine: Surplus_TextBox_Change - UserForm event handler. Called when Surplus_TextBox.Value changes.

    If loaded Then UpdateSaveButtonColor
End Sub

Private Sub LongLeadCheck_Change()
'Subroutine: LongLeadCheck_Change - UserForm event handler. Called when LongLeadCheck.Value changes.

    If loaded Then UpdateSaveButtonColor
End Sub

Private Sub Delete_CheckBox_Click()
'Subroutine: Delete_CheckBox_Click - UserForm event handler. Called when Delete_CheckBox.Value changes.
    
    'show/hide Reason box
    Reason_TextBox.Visible = Delete_CheckBox.Value
    Reason_Label.Visible = Delete_CheckBox.Value
    Reason_TextBox.Value = ORG_Reason_TextBox.Value
    
    If loaded Then UpdateSaveButtonColor
End Sub

Private Sub Reason_TextBox_Change()
'Subroutine: Reason_TextBox_Change - UserForm event handler. Called when Reason_TextBox.Value changes.

    If loaded Then UpdateSaveButtonColor
End Sub

Private Sub CategoryBox_Change()
'Subroutine: CategoryBox_Change - UserForm event handler. Called when CategoryBox.Value changes.
    
    'update unit
    Unit_TextBox.Value = VB_CATEGORY.GetCategoryUnit(CategoryBox.Value)
    
    If loaded Then UpdateSaveButtonColor
End Sub

Private Sub Mark_TextBox_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
'Subroutine: Mark_TextBox_Change - UserForm event handler. Called when Mark_TextBox.Value changes.
    
    If ThisWorkbook.MarkNumLock() Then
        Cancel = True
    Else
        Cancel = False
        If loaded Then UpdateSaveButtonColor
    End If
End Sub

Private Sub SAPBox_Change()
'Subroutine: SAPBox - UserForm event handler. Called when SAPBox.Value changes.

    If loaded Then UpdateSaveButtonColor
End Sub

Private Sub Description_TextBox_Change()
'Subroutine: Description_TextBox_Change - UserForm event handler. Called when Description_TextBox.Value changes.

    If Description_TextBox.Value <> ORG_Description_TextBox.Value Then
        DescCheck_CheckBox.Value = False
    Else
        DescCheck_CheckBox.Value = ORG_DescCheck_CheckBox.Value
    End If
    
    If loaded Then UpdateSaveButtonColor
End Sub

Private Sub DescCheck_CheckBox_Change()
'Subroutine: DescCheck_CheckBox_Change - UserForm event handler. Called when DescCheck_CheckBox.Value changes.

    If loaded Then UpdateSaveButtonColor
End Sub

Public Sub UpdateSaveButtonColor()
'Subroutine: UpdateSaveButtonColor - Updates SaveButton color. Checks for difference in any of the property
'                                    fields from the originals. If there was a change, highight SaveButton


    overall_change = False
    
    'overall_change = overall_change
    overall_change = overall_change Or Surplus_TextBox.Value <> ORG_Surplus_TextBox.Value
    overall_change = overall_change Or LongLeadCheck.Value <> ORG_LongLeadCheck.Value
    overall_change = overall_change Or Delete_CheckBox.Value <> ORG_Delete_CheckBox.Value
    overall_change = overall_change Or (Delete_CheckBox.Value And Reason_TextBox.Value <> ORG_Reason_TextBox.Value)
    overall_change = overall_change Or CategoryBox.Value <> ORG_CategoryBox.Value
    overall_change = overall_change Or Mark_TextBox.Value <> ORG_Mark_TextBox.Value
    overall_change = overall_change Or (SAPBox.Visible And SAPBox.Value <> ORG_SAPBox.Value)
    overall_change = overall_change Or Unit_TextBox.Value <> ORG_Unit_TextBox.Value
    overall_change = overall_change Or Description_TextBox.Value <> ORG_Description_TextBox.Value
    overall_change = overall_change Or DescCheck_CheckBox.Value <> ORG_DescCheck_CheckBox.Value
    
    'check site quantities
    If Not overall_change Then
        Dim site As Integer
        Dim site_name As String
        Dim first_site_col As Integer
        first_site_col = get_col_num("Current Model Quantities")
        
        For site = first_site_col To num_sites + first_site_col - 1
            site_name = VB_MASTER.Cells(VB_MASTER.SubtitleRow(), site).Value2
            
            overall_change = overall_change Or DFTNGFrame.Controls.Item(site_name & "_Box").Value <> ORG_DFTNGFrame.Controls.Item("ORG_" & site_name & "_Box").Value
            overall_change = overall_change Or ExtrasFrame.Controls.Item(site_name & "Extras_Box").Value <> ORG_ExtrasFrame.Controls.Item("ORG_" & site_name & "Extras_Box").Value
        Next site
    End If
    
    If overall_change Then
        SaveButton.BackColor = RGB(255, 255, 0)
    Else
        SaveButton.BackColor = CancelButton.BackColor
    End If
End Sub

