VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VB_SITEBOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const first_row = 4

Public Function FirstRow() As Integer
'Function: FirstRow - 'Get' method for first_row property in Site BOMs
'Returns: Integer containing the row of the first material item in Site BOMs

    FirstRow = first_row
End Function
