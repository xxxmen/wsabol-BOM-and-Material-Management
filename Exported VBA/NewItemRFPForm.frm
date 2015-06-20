VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewItemRFPForm 
   Caption         =   "Unrecognized Item and/or Description"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6240
   OleObjectBlob   =   "NewItemRFPForm.frx":0000
End
Attribute VB_Name = "NewItemRFPForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Form: NewItemRFPForm                                  '
'                                                            '
' When item in VB_ORDER_TMPLT is not found in the BOM. This  '
' form gives the user the oppurtunity to add the item to the '
' BOM, Cancel the order process, or continue without adding  '
' the item.                                                  '
'                                                            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private vbClick As Integer
Private form_row As Integer
Private loaded As Boolean

Public Sub LoadForm(ByVal message As String, ByVal in_desc As String, ByVal in_row As Integer, ByRef vbResult As Integer)
'Subroutine: LoadForm - Controls the .Show routine for this form. Sets form prompt; Sets position on the screen, etc.
'Arguments: message - String containing the form prompt
'           in_desc - String containing the item description not found in the BOM
'           in_row - Integer containing the row number where this item is in VB_ORDER_TMPLT
'           vbResult - Integer reference returning an value corresponding to the operation preformed. 1 - AddButton,
'           2 - GoRouge, 3 - Cancel.


    form_row = in_row
    Me.ItemDesc1 = in_desc
    Me.ItemDesc1.Visible = False
    Me.vbMessage = message
    
    Me.Top = Application.Top + Application.height / 2 - Me.height / 2
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    
    vbClick = -1
    loaded = True
    Me.Show

    ' 1 - AddButton
    ' 2 - GoRouge
    ' 3 - Cancel
    vbResult = vbClick
End Sub

Private Sub AddButton_Click()
'Subroutine: AddButton_Click - UserForm event handler. Called when user clicks AddButton.

    
    'Set Returns
    vbClick = 1
    
    Dim Action As Integer
    Dim new_mark_num As Integer
    Dim new_item_desc As String
    new_item_desc = ItemDesc1.Caption
    
    Dim first_phrase As String
    Dim category As String
    first_phrase = FirstPhrase(new_item_desc)
    category = VB_CATEGORY.FindCategory(first_phrase)
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    '''' From this point on GoTo ResetApp, not Exit Sub
    
    If Len(category) > 0 Then
        new_mark_num = InsertItemBOM(category, new_item_desc)
    Else
        new_mark_num = AddItemForm.LoadForm(new_item_desc)
    End If
    
    'if add item was not successful, cancel
    If new_mark_num = -1 Then
        vbClick = 3
        GoTo ResetApp
    End If
    
    Me.Hide
    
    ''''''''''''''''''''''''''''''''''''
    'highlight item in BOM to show user
    VB_MASTER.Activate
    VB_MASTER.Cells(get_row(new_mark_num), get_col_num("Description Check")).Value = "a"
    
    With VB_MASTER.Range(ColLet(get_col_num("Mark No.")) & get_row(new_mark_num) & ":" & _
        ColLet(get_col_num("Long Description")) & get_row(new_mark_num))
    
        .Interior.ColorIndex = 6
        .Activate
    End With
    
    Application.ScreenUpdating = True
    
    'ask if the addition went the way it should have
    Dim result As Integer
    result = MsgBox("The new item is hilighted here in the Master BOM sheet. Do you approve this addition? If not, the action will be undone.", vbYesNo)
    
    VB_ORDER_TMPLT.Activate
    
    Application.ScreenUpdating = False
    ''''''''''''''''''''''''''''''''''''
    
    VB_MASTER.Cells(get_row(new_mark_num), get_col_num("Description Check")).ClearContents
    
    If result = vbOK Then ' it went well
        With VB_MASTER.Range(ColLet(get_col_num("Mark No.")) & get_row(new_mark_num) & ":" & _
            ColLet(get_col_num("Long Description")) & get_row(new_mark_num))
        
            .Interior.ColorIndex = xlNone
        End With
        
        'log change
        If VB_CHANGE_LOG.LogChange(VB_MASTER.Name & "!$" & ColLet(get_col_num("Long Description")) & "$" & get_row(new_mark_num), new_mark_num, "Added Material Item", "", new_item_desc) <> 0 Then
            'throw error
            ErrorHandling "NewItemForm1.AddItem", 0, "LogChange() operation failed:" & vbCrLf & _
                "LogChange(" & VB_MASTER.Name & "!" & ColLet(get_col_num("Long Description")) & get_row(new_mark_num) & ", " & new_mark_num & ", Added Material Item, , " & new_item_desc & ")" & vbCrLf & _
                "Continue [OK] or Cancel procedure?", 1
        End If

    Else ' user needs to change something, undo addition.
        DeleteMasterRow get_row(new_mark_num)
        vbClick = 3
        GoTo ResetApp
    End If
    
    VB_ORDER_TMPLT.Cells(form_row, VB_ORDER_TMPLT.Form_col_num("Item #")).Value = new_mark_num
    
ResetApp:
    If ui_change Then RenderUI True
    
    Unload Me
End Sub

Private Sub RougeButton_Click()
'Subroutine: RougeButton_Click - UserForm event handler. Called when user clicks RougeButton.

    vbClick = 2
    
    VB_ORDER_TMPLT.Cells(form_row, VB_ORDER_TMPLT.Form_col_num("Item #")).ClearContents
    
    Unload Me
End Sub

Private Sub CancelButton_Click()
'Subroutine: CancelButton_Click - UserForm event handler. Called when user clicks CancelButton.

    vbClick = 3
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Subroutine: UserForm_QueryClose - UserForm event handler. Called when form is closed.
'Arguments: Cancel - Integer returning a Cancel action.
'           CloseMode - Integer containing the close mode.

    
    If vbClick = -1 Then
        vbClick = 3
    End If
    
    Cancel = True
    If loaded = True Then
        Cancel = False
        loaded = False
    End If
End Sub

