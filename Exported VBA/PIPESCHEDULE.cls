VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PIPESCHEDULE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const first_row = 2

Public Function ActualOD(ByVal NomOD As String) As Double
'Function: ActualOD - Retrieves the corresponding actual OD for the given Nominal OD string argument
'Arguments: NomOD - String containing to the nominal OD: '12"'
'Returns: Double containing the actual OD: 12.75

    ActualOD = 0
    
    Dim row As Integer
    row = first_row
    Do While Not IsEmpty(PIPESCHEDULE.Cells(row, 1))
        If CStr(PIPESCHEDULE.Cells(row, 1).Value2) = NomOD Then
            ActualOD = CDbl(PIPESCHEDULE.Cells(row, 2).Value2)
        End If
        
        row = row + 1
    Loop
End Function

Public Function FirstRow() As Integer
'Function: FirstRow - 'Get' method for first_row property in VB_MASTER
'Returns: Integer containing the row of the first material item in VB_MASTER

    FirstRow = first_row
End Function
