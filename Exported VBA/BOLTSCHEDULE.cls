VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BOLTSCHEDULE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const first_row = 3
Private Const NomOD_col = 1

Public Function BoltDiameter(ByVal NomSize As String, ByVal ClassIndex As Integer) As String
'Function: BoltDiameter - Retrieves the bolt diameter for a given size and pressure class
'Arguments: NomSize - String containing the nominal size
'           ClassIndex - Integer corresponding to a pressure class. ANSI 150 - 0 .. ANSI 2500 - 6
'Returns: String containing bolt diameter in fractional form, i.e. '1 1/8"'


    Dim row As Integer
    row = first_row
    
    'search for a matching NomSize
    Do While Not IsEmpty(BOLTSCHEDULE.Cells(row, NomOD_col))
        If CStr(BOLTSCHEDULE.Cells(row, NomOD_col).Value2) = NomSize Then
            Exit Do
        End If
        row = row + 1
    Loop
    
    BoltDiameter = BOLTSCHEDULE.Cells(row, 3 + 5 * ClassIndex).Value2
End Function

Public Function BoltLength(ByVal NomSize As String, ByVal face_type_code As String, ByVal ClassIndex As Integer) As String
'Function: BoltLength - Retrieves the bolt length for a given size, face type, and pressure class
'Arguments: NomSize - String containing the nominal size
'           face_type_code - String containing the 2 or 3 character code corresponding to face type i.e. "RF", "FF", ..
'           ClassIndex - Integer corresponding to a pressure class. ANSI 150 - 0 .. ANSI 2500 - 6
'Returns: String containing bolt length i.e. '7 1/4"'


    Dim row As Integer
    row = first_row
    
    'search for a matching NomSize
    Do While Not IsEmpty(BOLTSCHEDULE.Cells(row, NomOD_col))
        If CStr(BOLTSCHEDULE.Cells(row, NomOD_col).Value2) = NomSize Then
            Exit Do
        End If
        row = row + 1
    Loop
    
    Dim col As Integer
    If face_type_code = "RF" Then
        col = 4 + 5 * ClassIndex
    Else
        col = 5 + 5 * ClassIndex
    End If
    
    BoltLength = BOLTSCHEDULE.Cells(row, col).Value2
End Function

Public Function BoltQuantity(ByVal NomSize As String, ByVal ClassIndex As Integer) As String
'Function: BoltQuantity - Retrieves the bolt quantity for a given size and pressure class
'Arguments: NomSize - String containing the nominal size
'           ClassIndex - Integer corresponding to a pressure class. ANSI 150 - 0 .. ANSI 2500 - 6
'Returns: String containing bolt quantity i.e. '8'


    Dim row As Integer
    row = first_row
    
    'search for a matching NomSize
    Do While Not IsEmpty(BOLTSCHEDULE.Cells(row, NomOD_col))
        If CStr(BOLTSCHEDULE.Cells(row, NomOD_col).Value2) = NomSize Then
            Exit Do
        End If
        row = row + 1
    Loop
    
    BoltQuantity = BOLTSCHEDULE.Cells(row, 2 + 5 * ClassIndex).Value2
End Function
