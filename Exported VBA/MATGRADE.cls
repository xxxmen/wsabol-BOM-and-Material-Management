VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MATGRADE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const first_row = 2
Private Const grade_col = 1

Public Function GradeLookup(ByVal grade As String, ByVal col_title As String) As String
'Function: GradeLookup - Used for the AutoDescriptor. MATGRADE sheet is a lookup for Grade dependant properties for
'                        flanges and BW fittings. Given a grade (row) and col_title (col), this function returns the corresponding property.
'Arguments: grade - String containing pipe grade. i.e. GR A, X52, etc.
'           col_title - String containing column title for desired property. See list below.
'Returns: String containing lookup value


    GradeLookup = ""
    
    'FLANGE GRADE - "FLNG GRD"
    'FLANGE MATERIAL - "FLNG MAT"
    'FLANGE SPEC - "FLNG SPEC"
    'BW FITTING GRADE - "BWFIT GRD"
    'BW FITTING MATERIAL - "BWFIT MAT"
    'BW FITTING SPEC - "BWFIT SPEC"
    
    Dim col As Integer
    col = 1
    
    'search for matching col_title
    Do While Not IsEmpty(MATGRADE.Cells(1, col))
        If CStr(MATGRADE.Cells(1, col).Value2) = col_title Then
            Exit Do
        End If
        col = col + 1
    Loop
    
    'error check
    If IsEmpty(MATGRADE.Cells(1, col)) Then
        Exit Function
    End If
    
    Dim row As Integer
    row = 1
    
    'search for matching grade
    Do While Not IsEmpty(MATGRADE.Cells(row, grade_col))
        If CStr(MATGRADE.Cells(row, grade_col).Value2) = grade Then
            Exit Do
        End If
        row = row + 1
    Loop

    If CStr(MATGRADE.Cells(row, grade_col).Value2) = grade Then
        GradeLookup = MATGRADE.Cells(row, col).text ' use .Text becasue its handles "" cases, and won't return 0 like .Value(2)
    End If
End Function

Public Function FirstRow() As Integer
'Function: FirstRow - 'Get' method for first_row property in MATGRADE
'Returns: Integer containing the row of the first material item in MATGRADE


    FirstRow = first_row
End Function

Public Function SpecListColumn() As Integer
'Function: SpecListColumn - 'Get' method for SpecList property in MATGRADE
'Returns: Integer containing the column number where the line spec list begins in MATGRADE


    SpecListColumn = 1
    Do While Not IsEmpty(MATGRADE.Cells(1, SpecListColumn))
        If MATGRADE.Cells(1, SpecListColumn).Value2 = "SpecList" Then
            Exit Function
        End If
        SpecListColumn = SpecListColumn + 1
    Loop

    SpecListColumn = -1
End Function
