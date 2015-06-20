VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddItemWindow 
   Caption         =   "New Material Properties"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8880
   OleObjectBlob   =   "AddItemWindow.frx":0000
End
Attribute VB_Name = "AddItemWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Form: AddItemWindow                                 '
'                                                          '
' Adding a brand new item to the BOM goes through this     '
' window. Prevents user from inserting for new item out of '
' order or leaving a new item description blank, and funky '
' copy/paste errors.                                       '
'                                                          '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private mark_num As Integer

Public Sub LoadForm(ByRef new_mark_num As Integer)
'Subroutine: LoadForm - Loads form with item categories and sets focus.
'Arguemnts: new_mark_num - Integer reference returning new item's mark number

    Description_TextBox.SetFocus
    
    'category
    Dim index As Integer
    Dim cat_row As Integer
    Dim cat_col As Integer
    index = 0
    cat_row = VB_CATEGORY.FirstRow()
    cat_col = VB_CATEGORY.CategoryColumn()
    
    Do While Not VB_CATEGORY.EOF(cat_row)
        CategoryBox.AddItem VB_CATEGORY.Cells(cat_row, cat_col).Value2
        
        cat_row = cat_row + 1
    Loop
    CategoryBox.ListIndex = 0
    
    Dim longlead As Boolean
    ShowLongLead_GetPressed Nothing, longlead
    LongLeadCheck.Value = longlead
    LongLeadCheck.enabled = Not longlead
    
    ' set position on screen
    Me.Top = Application.Top + Application.height / 2 - Me.height / 2
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    Me.Show
    
    new_mark_num = mark_num
End Sub

Private Sub AddButton_Click()
'Subroutine: AddButton_Click - UserForm event handler. Called when user clicks AddButton.
'                              If form is complete and description doesn't already exist, add new item to BOM.

    Dim ui_change As Boolean
    
    Description_TextBox.Value = TrimWhiteSpace(Description_TextBox.Value)
    
    'check description
    If Len(Description_TextBox.Value) = 0 Then
        MsgBox "You must input a description to successfully add a material item.", vbExclamation
        With Description_TextBox
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.text)
        End With
        Exit Sub
    ElseIf MatchDesc(Description_TextBox.Value) <> -1 Then
        MsgBox "This item description already exists in this BOM. New material items must have unique descriptions.", vbExclamation
        With Description_TextBox
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.text)
        End With
        Exit Sub
    End If
    
    'check for illegal characters
    If ContainsIllegalCharacters(Description_TextBox.Value) Then
        MsgBox "You entered an illegal character. Do not use characters " & vbCrLf & vbCrLf & "* _ [ ] ^", vbExclamation
        With Description_TextBox
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.text)
        End With
        Exit Sub
    End If
    
    'check Surplus entry
    If Len(Surplus_TextBox.Value) > 0 Then
        If Not IsNumeric(Surplus_TextBox.Value) Then
            MsgBox "Client Inventory value must be numeric.", vbExclamation
            With Surplus_TextBox
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.text)
            End With
            Exit Sub
        ElseIf Fix(Surplus_TextBox.Value) <> CDbl(Surplus_TextBox.Value) Or CInt(Surplus_TextBox.Value) <= 0 Then
            MsgBox "Client Inventory value must be positive natural number.", vbExclamation
            With Surplus_TextBox
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.text)
            End With
            Exit Sub
        End If
    End If
    
    ui_change = RenderUI(False)
    
    'add item, get mark number
    mark_num = InsertItemBOM(CategoryBox.Value, Description_TextBox.Value)
    
    If get_row(mark_num) <> -1 Then
        Dim suppressed As Boolean
        SuppressMarkNums_GetPressed Nothing, suppressed
        If suppressed Then
            MsgBox "Successfully added material item to " & CategoryBox.Value & ".", vbInformation
        Else
            MsgBox "Successfully added Item #" & mark_num & ".", vbInformation
        End If
        
        If Len(Surplus_TextBox.Value) > 0 Then
            set_property mark_num, get_col_num("Client Inventory"), CInt(Surplus_TextBox.Value)
            If CInt(Surplus_TextBox.Value) <> 0 Then
                Call VB_MASTER.WriteChange(VB_MASTER.Cells(get_row(mark_num), get_col_num("Client Inventory")), 0)
            End If
        End If
        
        If LongLeadCheck.Value Then
            set_property mark_num, get_col_num("Long Lead"), LongLeadCheck.Value
            Call VB_MASTER.WriteChange(VB_MASTER.Cells(get_row(mark_num), get_col_num("Long Lead")), False)
        End If
    Else
        MsgBox "Something went wrong.", vbQuestion
    End If
    
    VB_MASTER.CalculateQuantityFormat
    
    Dim longlead As Boolean
    ShowLongLead_GetPressed Nothing, longlead
    If longlead Then Ribbon_ShowLongLead Nothing, True
    
    If ui_change Then RenderUI True
    
    Unload Me
End Sub

Private Sub AutoDescriptorButton_Click()
'Subroutine: AutoDescriptorButton_Click - UserForm event handler. Called when user clicks AutoDescriptorButton.

    AutoDescriptor.LoadForm
    
    Dim ADresult As String
    ADresult = VB_VAR_STORE.GetAutoDescription()
    
    If Len(ADresult) > 0 Then
        Description_TextBox.Value = ADresult
    End If
    VB_VAR_STORE.ClearAutoDescription
End Sub

Private Sub CancelButton_Click()
'Subroutine: CancelButton_Click - UserForm event handler. Called when user clicks CancelButton.
    
    mark_num = -1
    Unload Me
End Sub

Private Sub CategoryBox_Change()
'Subroutine: CategoryBox_Change - UserForm event handler. Called when user changes CategoryBox.

    'update unit
    Unit_TextBox.Value = VB_CATEGORY.GetCategoryUnit(CategoryBox.Value)
End Sub

Private Sub Description_TextBox_Change()
'Subroutine: Description_TextBox_Change - UserForm event handler. Called when user changes Description_TextBox.

    Dim category_guess As String
    category_guess = VB_CATEGORY.FindCategory(FirstPhrase(Description_TextBox.Value))
    
    If category_guess <> vbNullString Then
        CategoryBox.Value = category_guess
    End If
End Sub

Private Sub NewCategoryButton_Click()
'Subroutine: NewCategoryButton_Click - UserForm event handler. Called when user clicks NewCategoryButton.
'                                      Runs routine to create new category, and refreshes category list for user to choose from.

    Dim new_cat As String
    new_cat = NewCategoryHandler()
    
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
