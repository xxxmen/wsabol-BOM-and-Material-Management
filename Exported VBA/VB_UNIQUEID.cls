VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VB_UNIQUEID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const first_row = 1

Public Sub AddUniqueID(ByVal uID As String)
'Subroutine: AddUniqueID - Appends new Unique IDs to the database. If uID exists, it doesn't do anything.
'Arguments: uID - String containing the new Unique ID to add. Format: "A000000"


    Dim row As Integer
    row = first_row
    
    'Check if exists
    Do While Not EOF(row)
        If CStr(VB_UNIQUEID.Cells(row, 1).Value2) = uID Then
            Exit Sub
        End If
        
        row = row + 1
    Loop
    
    VB_UNIQUEID.Cells(row, 1).Value = uID
End Sub

Public Sub ClearAll()
'Subroutine: ClearAll - Clears all Unique IDs. Used in BlankBOM

    VB_UNIQUEID.UsedRange.ClearContents
End Sub

Public Function UniqueIDExists(ByVal testID As String) As Boolean
'Function: UniqueIDExist - Unique IDs are used for the CADWorx plugin. Unique IDs are unique, unchanging
'                          strings that correspond to each material item. This function grabs the mark
'                          number for a given Unique ID
'Arguments: testID - String containing target Unique ID
'Returns: Boolean. True - the Unique ID exists; False - the unique ID does NOT exist


    Dim row As Integer
    row = first_row
    
    Do While Not EOF(row)
        If CStr(VB_UNIQUEID.Cells(row, 1).Value2) = testID Then
            UniqueIDExists = True
            Exit Function
        End If
        
        row = row + 1
    Loop
    
    UniqueIDExists = False
End Function

Public Function FirstRow() As Integer
'Function: FirstRow - 'Get' method for first_row property in VB_UNIQUEID
'Returns: Integer containing first_row

    FirstRow = first_row
End Function

Private Function EOF(ByVal row As Integer) As Boolean
'Function: EOF - Tests if row is beyond the scope of this sheet
'Arguments: row - Integer containing the row number in this sheet
'Returns: Boolean. True - row is beyond the scope; False - row is within the meaningful scope

    EOF = IsEmpty(VB_UNIQUEID.Cells(row, 1))
End Function


