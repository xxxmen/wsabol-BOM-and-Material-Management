VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MAOPHelpForm 
   Caption         =   "MAOP Help"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7365
   OleObjectBlob   =   "MAOPHelpForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MAOPHelpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Form: MAOPHelpForm                                     '
'                                                             '
' From within the AutoDescriptor, the user can call this      '
' form. Based on the values in the AutoDescriptor, and values '
' choosen on this form, its calculates an MAOP. Calculation   '
' procedure is from ASME B31.8 and the steps are shown in the '
' form.
'                                                             '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private SYMS As Long
Private OD As Double
Private WT As Double
Private DF As Double
Private TF As Double
Private CA As Double

Public Sub LoadForm(ByVal inNomOD As String, ByVal inWT As String, ByVal inGrade As String)
'Subroutine: LoadForm - Handles the .Show routine for this form.
'Arguments: inNomOD - String containing the Nominal OD for the calculation
'           inWT - String containing the Wall Thickness for the calculation
'           inGrade - String containing the Grade for the calculation


    NomODBox.Value = inNomOD
    WTBox.Value = inWT
    GradeBox.Value = inGrade
    
    UpdateTemperature
    UpdateCA
    DFBox_Change
    
    'set position on screen
    Me.Top = Application.Top + Application.height / 2 - Me.height / 2
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    Me.Show
End Sub

Private Sub NomODBox_Change()
'Subroutine: NomODBox_Change - User Form event handler. Called whenever NomODBox changes.

    OD = PIPESCHEDULE.ActualOD(NomODBox.Value)
End Sub

Private Sub WTBox_Change()
'Subroutine: WTBox_Change - User Form event handler. Called whenever WTBox changes.

    WT = CDbl(Left(WTBox.Value, Len(WTBox.Value) - 1))
End Sub

Private Sub GradeBox_Change()
'Subroutine: GradeBox_Change - User Form event handler. Called whenever GradeBox changes.

    SYMS = CLng(MATGRADE.GradeLookup(GradeBox.Value, "SMYS"))
End Sub

Private Sub DFBox_Change()
'Subroutine: DFBox_Change - User Form event handler. Called whenever DFBox changes.
'                           Updates the calculation.

    DF = CDbl(DFBox.Value)
    UpdateMAOP
    UpdateCalc
End Sub

Private Sub TempBox_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
'Subroutine: TempBox_BeforeUpdate - User Form event handler. Called whenever TempBox updates.
'                                   Updates the calculation.
    
    If Not IsNumeric(TempBox.Value) Then
        MsgBox "Tempature must be a positive integer.", vbCritical
        Cancel = True
        Exit Sub
    End If
    
    UpdateTemperature
    UpdateMAOP
    UpdateCalc
End Sub

Private Sub UpdateTemperature()
'Subroutine: UpdateTemperature - Calculates temperature degrading factor. Source: ASME B31.8
    
    Dim temperature As Integer
    temperature = CInt(TempBox.Value)

    If temperature < 250 Then
        TF = 1
    ElseIf temperature >= 250 And temperature < 300 Then
        TF = (0.967 - 1) / (300 - 250) * (temperature - 250) + 1
    ElseIf temperature >= 300 And temperature < 350 Then
        TF = (0.933 - 0.967) / (350 - 300) * (temperature - 300) + 0.967
    ElseIf temperature >= 350 And temperature < 400 Then
        TF = (0.9 - 0.933) / (400 - 350) * (temperature - 350) + 0.933
    ElseIf temperature >= 400 And temperature < 450 Then
        TF = (0.867 - 0.9) / (450 - 400) * (temperature - 400) + 0.9
    ElseIf temperature >= 450 Then
        TF = (0.867 - 0.9) / (450 - 400) * (temperature - 400) + 0.9
    End If
End Sub

Private Sub CABox_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
'Subroutine: CABox_BeforeUpdate - User Form event handler. Called whenever CABox updates.
'                                 Updates the calculation.

    UpdateCA
    UpdateMAOP
    UpdateCalc
End Sub

Private Sub UpdateCA()
'Subroutine: UpdateCA - Handles valid entries in the corrosion allowance field

    If Right(CABox.Value, 1) <> """" Then
        CABox.Value = CABox.Value & """"
    End If
    
    If IsNumeric(Left(CABox.Value, Len(CABox.Value) - 1)) Then
        CA = CDbl(Left(CABox.Value, Len(CABox.Value) - 1))
        CABox.Value = IIf(CA = 0, "0", Format(CA, "0.0###")) & """"
    Else
        MsgBox "Corrosion Allowance must be a numeric value in inches (i.e. 0.0625"")", vbExclamation
        CA = 0
        CABox.Value = "0"""
    End If
End Sub

Private Sub UpdateMAOP()
'Subroutine: UpdateMAOP - Updates MAOP field

    MAOP.Caption = Format(2 * SYMS * (WT - CA) * DF * TF / OD, "#,###.0") & " PSIG"
End Sub

Private Sub UpdateCalc()
'Subroutine: UpdateCalc - Updates calculation steps field

    CalcLabel.Caption = "(2 * " & SYMS & " * " & IIf(CA > 0, "(" & Format(WT, "0.000") & " - " & Format(CA, "0.0###") & ")", Format(WT, "0.000")) & _
        " * " & Format(DF, "0.0#") & " * 1.0 * " & Format(TF, "0.0##") & ") / " & Format(OD, "0.000")
End Sub
