VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AutoDescriptor 
   Caption         =   "Automated Material Descriptor"
   ClientHeight    =   13080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8400
   OleObjectBlob   =   "AutoDescriptor.frx":0000
End
Attribute VB_Name = "AutoDescriptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Form: AutoDescriptor                                  '
'                                                            '
' Automatic Descriptor for material Item. Link to this form  '
' is in AddItemWindow. Currently works for Line Pipe,        '
' some BW/TH/SW Fittings, Flanges, and Gaskets & Bolts. User '
' chooses from options provided. The "Refresh Description"   '
' button creates the material description, and the "Save &   '
' Close" button sends the description back to AddItemWindow. '
'                                                            '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Public Sub LoadForm()
'Function: LoadForm - Controls the .Show routine for this form. Renders the initial frame, sets position on the screen, etc.


    TypeInputFrame.Visible = True
    RenderBWFittingFrame False
    RenderTHSWFittingFrame False
    RenderUserInputFrame False
    RenderPipeFrame False
    RenderFlangeFrame False, False
    RenderBWTeeEllFrame False
    RenderTHSWTeeEllCapFrame False
    RenderFinalFrame False
    
    
    'set position on screen
    Me.Top = Application.Top + Application.height / 3 - Me.height / 2
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    Me.Show
End Sub

Private Sub MAOPButton_Click()
'Subroutine: MAOPButton_Click - UserForm event handler. Called when user clicks MAOPButton.
'                               Renders the MAOP form.

    If SizeMain_Box.MatchFound And Len(WTMain_Box.Value) > 0 And GradeBox.MatchFound _
        And IIf(SizeRed_Box.Visible, SizeRed_Box.MatchFound, True) And IIf(WTRed_Box.Visible, Len(WTRed_Box.Value) > 0, True) Then
        
        MAOPHelpForm.LoadForm SizeMain_Box.Value, WTMain_Box.Value, GradeBox.Value
    Else
        MsgBox "Complete the form with Sizes, Wall Thicknesses, and Material Grade before using the MAOP help.", vbExclamation
    End If
End Sub

Private Sub RefreshButton_Click()
'Subroutine: RefreshButton_Click - UserForm event handler. Called when user clicks RefreshButton.
'                                  Generated the description for DescriptionBox; checks for completed form.


    If Not FinalFrame.Visible Then Exit Sub
    
    Dim description As String
    
    Select Case MaterialTypeBox.Value
    
        '''''''''''''''
        'LINE PIPE
        Case "LINE PIPE":
            If SizeMain_Box.MatchFound Then
                description = "PIPE - " & SizeMain_Box.Value
            Else
                MsgBox "Please select a valid nominal diameter.", vbExclamation
                GoTo INCOMPLETE
            End If
            If Len(WTMain_Box.Value) > 0 Then
                description = description & ", " & WTMain_Box.Value & " WT"
            Else
                MsgBox "Please select a valid wall thickness.", vbExclamation
                GoTo INCOMPLETE
            End If
            If PipeMakeBox.MatchFound Then
                description = description & ", " & PipeMakeBox.Value
            Else
                MsgBox "Please select a valid pipe make.", vbExclamation
                GoTo INCOMPLETE
            End If
            If PipeEndTypeBox.MatchFound Then
                description = description & ", " & PipeEndTypeCode(PipeEndTypeBox.Value)
            Else
                MsgBox "Please select a valid pipe make.", vbExclamation
                GoTo INCOMPLETE
            End If
            If SpecBox.MatchFound Then
                description = description & ", " & SpecBox.Value
            Else
                MsgBox "Please select a valid pipe specification.", vbExclamation
                GoTo INCOMPLETE
            End If
            If GradeBox.MatchFound Then
                description = description & ", " & GradeBox.Value
            Else
                MsgBox "Please select a valid pipe grade.", vbExclamation
                GoTo INCOMPLETE
            End If
            If PipeCoatingBox.MatchFound Then
                If StrComp(PipeCoatingBox.Value, "BARE", vbTextCompare) = 0 Then
                    description = description & ", BARE"
                Else
                    description = description & ", COATED WITH " & UCase(PipeCoatingBox.Value)
                End If
            Else
                MsgBox "Please select a valid pipe coating option.", vbExclamation
                GoTo INCOMPLETE
            End If
            If PipeCertifiedBox.Value Then
                description = description & ", COMPLETE WITH CERTIFIED MILL TEST REPORTS"
            End If
        
        '''''''''''''''
        'FLANGES
        Case "FLANGES":
            description = "FLANGE"
            If BlindBox.Value Then description = description & ", BLIND"
            
            If SizeMain_Box.MatchFound Then
                description = description & " - " & SizeMain_Box.Value
            Else
                MsgBox "Please select a valid nominal diameter.", vbExclamation
                GoTo INCOMPLETE
            End If
            If FlangeClassBox.MatchFound Then
                description = description & ", " & FlangeClassBox.Value
            Else
                MsgBox "Please select a valid Pressure Class.", vbExclamation
                GoTo INCOMPLETE
            End If
            If FlangeFaceType.MatchFound Then
                description = description & ", " & FaceTypeCode(FlangeFaceType.Value)
            Else
                MsgBox "Please select a valid face type.", vbExclamation
                GoTo INCOMPLETE
            End If
            If Not FlangeEndType.MatchFound Then
                MsgBox "Please select a valid end type.", vbExclamation
                GoTo INCOMPLETE
            End If
            If Not BlindBox.Value Then
                description = description & ConnectionTypeCode(FlangeEndType.Value)
            End If
            If GradeBox.MatchFound Then
                description = description & ", " & MATGRADE.GradeLookup(GradeBox.Value, "FLNG MAT")
                description = description & ", " & MATGRADE.GradeLookup(GradeBox.Value, "FLNG GRD")
            Else
                MsgBox "Please select a valid grade.", vbExclamation
                GoTo INCOMPLETE
            End If
            If Not (BlindBox.Value Or FlangeEndType.Value = "Threaded") Then
                If Len(WTMain_Box.Value) > 0 Then
                    description = description & ", TWT " & WTMain_Box.Value & " WT"
                Else
                    MsgBox "Please select a valid wall thickness.", vbExclamation
                    GoTo INCOMPLETE
                End If
                
                description = description & ", " & GradeBox.Value & " PIPE"
            End If
            description = description & ", " & MATGRADE.GradeLookup(GradeBox.Value, "FLNG SPEC")
            If BlindBox.Value And Not (CDTBox.Value = "N/A" Or Not CDTBox.MatchFound) Then
                description = description & ", CD&T FOR " & CDTBox.Value & " NPT"
            End If
        
        ''''''''''''''''
        'BW FITTINGS
        Case "BW FITTINGS"
            Select Case BWFittingType.Value
            
                ''' BW TEE '''
                Case "TEE":
                    description = "TEE"
                    If BWOption1.Value Then
                        description = description & ", " & UCase(BWOption1.Caption)
                    ElseIf BWOption2.Value Then
                        description = description & ", " & UCase(BWOption2.Caption)
                    Else
                        MsgBox "Please select TEE type.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If SizeMain_Box.MatchFound Then
                        description = description & " - " & SizeMain_Box.Value
                    Else
                        MsgBox "Please select a valid nominal diameter.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If Len(WTMain_Box.Value) > 0 Then
                        description = description & ", " & WTMain_Box.Value & " WT"
                    Else
                        MsgBox "Please select a valid wall thickness.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If BWOption2.Value Then 'reducing
                        If SizeRed_Box.MatchFound Then
                            If SizeRed_Box.ListIndex < SizeMain_Box.ListIndex Then
                                description = description & " x " & SizeRed_Box.Value
                            Else
                                MsgBox "Reduced diameter cannot be greater than or equal to main diameter.", vbExclamation
                                GoTo INCOMPLETE
                            End If
                        Else
                            MsgBox "Please select a valid REDUCED nominal diameter.", vbExclamation
                            GoTo INCOMPLETE
                        End If
                        If Len(WTRed_Box.Value) > 0 Then
                            description = description & ", " & WTRed_Box.Value & " WT"
                        Else
                            MsgBox "Please select a valid REDUCED wall thickness.", vbExclamation
                            GoTo INCOMPLETE
                        End If
                    End If
                    description = description & ", BW"
                    If GradeBox.MatchFound Then
                        description = description & ", " & MATGRADE.GradeLookup(GradeBox.Value, "BWFIT MAT")
                        description = description & ", " & MATGRADE.GradeLookup(GradeBox.Value, "BWFIT GRD")
                        description = description & MATGRADE.GradeLookup(GradeBox.Value, "BWFIT SPEC")
                    Else
                        MsgBox "Please select a valid grade.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If BWCheckOptionBox.Value Then
                        description = description & ", WITH BARRED BRANCH"
                    End If
                    
                ''' BW ELL '''
                Case "ELL":
                    description = "ELL"
                    If BWCheckOptionBox.Value Then
                        description = description & ", SEGMENTABLE"
                    End If
                    If BWEllDegree.MatchFound Then
                        description = description & ", " & BWEllDegree.Value
                    Else
                        MsgBox "Please select a valid ELL degree.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If BWEllRadius.MatchFound Then
                        description = description & " - " & BWEllRadius.Value
                    Else
                        MsgBox "Please select a valid ELL turn radius.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If SizeMain_Box.MatchFound Then
                        description = description & ", " & SizeMain_Box.Value
                    Else
                        MsgBox "Please select a valid nominal diameter.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If Len(WTMain_Box.Value) > 0 Then
                        description = description & ", " & WTMain_Box.Value & " WT"
                    Else
                        MsgBox "Please select a valid wall thickness.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    description = description & ", BW"
                    If GradeBox.MatchFound Then
                        description = description & ", " & MATGRADE.GradeLookup(GradeBox.Value, "BWFIT MAT")
                        description = description & ", " & MATGRADE.GradeLookup(GradeBox.Value, "BWFIT GRD")
                        description = description & MATGRADE.GradeLookup(GradeBox.Value, "BWFIT SPEC")
                    Else
                        MsgBox "Please select a valid grade.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    
                ''' BW REDUCER '''
                Case "REDUCER":
                    description = "REDUCER"
                    If BWOption1.Value Then
                        description = description & ", CONC."
                    ElseIf BWOption2.Value Then
                        description = description & ", ECC."
                    Else
                        MsgBox "Please select REDUCER type.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If SizeMain_Box.MatchFound Then
                        description = description & " - " & SizeMain_Box.Value
                    Else
                        MsgBox "Please select a valid nominal diameter.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If Len(WTMain_Box.Value) > 0 Then
                        description = description & ", " & WTMain_Box.Value & " WT"
                    Else
                        MsgBox "Please select a valid wall thickness.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If SizeRed_Box.MatchFound Then
                        description = description & " x " & SizeRed_Box.Value
                    Else
                        MsgBox "Please select a valid REDUCED nominal diameter.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If Len(WTRed_Box.Value) > 0 Then
                        description = description & ", " & WTRed_Box.Value & " WT"
                    Else
                        MsgBox "Please select a valid REDUCED wall thickness.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    description = description & ", BW"
                    If GradeBox.MatchFound Then
                        description = description & ", " & MATGRADE.GradeLookup(GradeBox.Value, "BWFIT MAT")
                        description = description & ", " & MATGRADE.GradeLookup(GradeBox.Value, "BWFIT GRD")
                        description = description & MATGRADE.GradeLookup(GradeBox.Value, "BWFIT SPEC")
                    Else
                        MsgBox "Please select a valid grade.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If BWOption2.Value Then
                        description = description & ", F.S.D."
                    End If
                    
                ''' BW CAP '''
                Case "CAP":
                    description = "CAP"
                    If SizeMain_Box.MatchFound Then
                        description = description & " - " & SizeMain_Box.Value
                    Else
                        MsgBox "Please select a valid nominal diameter.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If Len(WTMain_Box.Value) > 0 Then
                        description = description & ", " & WTMain_Box.Value & " WT"
                    Else
                        MsgBox "Please select a valid wall thickness.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    description = description & ", BW"
                    If GradeBox.MatchFound Then
                        description = description & ", " & MATGRADE.GradeLookup(GradeBox.Value, "BWFIT GRD")
                        description = description & MATGRADE.GradeLookup(GradeBox.Value, "BWFIT SPEC")
                    Else
                        MsgBox "Please select a valid grade.", vbExclamation
                        GoTo INCOMPLETE
                    End If
            
            End Select
        
        ''''''''''''''''''''''''
        ' TH/SW FITTINGS
        Case "TH/SW FITTINGS"
            Select Case THSWFittingType.Value
            
                ''' TH/SW TEE '''
                Case "TEE":
                    description = "TEE"
                    If THSWOption1.Value Then
                        description = description & ", " & UCase(THSWOption1.Caption)
                    ElseIf THSWOption2.Value Then
                        description = description & ", " & UCase(THSWOption2.Caption)
                    Else
                        MsgBox "Please select TEE type.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If SizeMain_Box.MatchFound Then
                        description = description & " - " & SizeMain_Box.Value
                    Else
                        MsgBox "Please select a valid nominal diameter.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If THSWOption2.Value Then 'reducing
                        If SizeRed_Box.MatchFound Then
                            If SizeRed_Box.ListIndex < SizeMain_Box.ListIndex Then
                                description = description & " x " & SizeRed_Box.Value
                            Else
                                MsgBox "Reduced diameter cannot be greater than or equal to main diameter.", vbExclamation
                                GoTo INCOMPLETE
                            End If
                        Else
                            MsgBox "Please select a valid REDUCED nominal diameter.", vbExclamation
                            GoTo INCOMPLETE
                        End If
                    End If
                    If THSWClassBox.MatchFound Then
                        description = description & ", " & THSWClassBox.Value
                    Else
                        MsgBox "Please select a valid Pressure Class.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If ThreadedOption.Value Then
                        description = description & ", FNPT"
                    ElseIf SocketWeldOption.Value Then
                        description = description & ", SW"
                    Else
                        MsgBox "Please select End Type: Threaded or Socket Welded.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If GradeBox.MatchFound Then
                        description = description & ", ASTM A105, GR B, PER ASME B16.11"
                    Else
                        MsgBox "Please select a valid grade.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    
                ''' TH/SW ELL '''
                Case "ELL":
                    description = "ELL"
                    If THSWDegreeBox.MatchFound Then
                        description = description & ", " & THSWDegreeBox.Value
                    Else
                        MsgBox "Please select a valid ELL degree.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If SizeMain_Box.MatchFound Then
                        description = description & " - " & SizeMain_Box.Value
                    Else
                        MsgBox "Please select a valid nominal diameter.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If THSWClassBox.MatchFound Then
                        description = description & ", " & THSWClassBox.Value
                    Else
                        MsgBox "Please select a valid Pressure Class.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If ThreadedOption.Value Then
                        description = description & ", FNPT"
                    ElseIf SocketWeldOption.Value Then
                        description = description & ", SW"
                    Else
                        MsgBox "Please select End Type: Threaded or Socket Welded.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If GradeBox.MatchFound Then
                        description = description & ", ASTM A105, GR B, PER ASME B16.11"
                    Else
                        MsgBox "Please select a valid grade.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    
                ''' TH/SW PLUG '''
                Case "PLUG":
                    description = "PLUG, HEXHEAD"
                    If SizeMain_Box.MatchFound Then
                        description = description & " - " & SizeMain_Box.Value
                    Else
                        MsgBox "Please select a valid nominal diameter.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If THSWClassBox.MatchFound Then
                        description = description & ", " & THSWClassBox.Value
                    Else
                        MsgBox "Please select a valid Pressure Class.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    description = description & ", MNPT"
                    If GradeBox.MatchFound Then
                        description = description & ", ASTM A105, PER ASME B16.11"
                    Else
                        MsgBox "Please select a valid grade.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    
                ''' TH/SW CAP '''
                Case "CAP":
                    description = "CAP"
                    If SizeMain_Box.MatchFound Then
                        description = description & " - " & SizeMain_Box.Value
                    Else
                        MsgBox "Please select a valid nominal diameter.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If THSWClassBox.MatchFound Then
                        description = description & ", " & THSWClassBox.Value
                    Else
                        MsgBox "Please select a valid Pressure Class.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If ThreadedOption.Value Then
                        description = description & ", FNPT"
                    ElseIf SocketWeldOption.Value Then
                        description = description & ", SW"
                    Else
                        MsgBox "Please select End Type: Threaded or Socket Welded.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                    If GradeBox.MatchFound Then
                        description = description & ", ASTM A105, GR B, PER ASME B16.11"
                    Else
                        MsgBox "Please select a valid grade.", vbExclamation
                        GoTo INCOMPLETE
                    End If
            
            End Select
            
        '''''''''''''''''''''''
        ' GASKETS & BOTLS
        Case "GASKETS & BOLTS":
        
            ' GASKET
            If GasketOption.Value Then
                description = "GASKET"
                If InsulatingBox.Value Then
                    description = description & ", INSULATING SET"
                Else
                    If GBFaceType.MatchFound Then
                        If GBFaceType.Value = "Raised Face" Then description = description & ", SPIRAL WOUND"
                    Else
                        MsgBox "Please select a valid face type.", vbExclamation
                        GoTo INCOMPLETE
                    End If
                End If
                If SizeMain_Box.MatchFound Then
                    description = description & " - " & SizeMain_Box.Value
                Else
                    MsgBox "Please select a valid nominal diameter.", vbExclamation
                    GoTo INCOMPLETE
                End If
                If GasketClass.MatchFound Then
                    description = description & ", " & GasketClass.Value
                Else
                    MsgBox "Please select a valid pressure class.", vbExclamation
                    GoTo INCOMPLETE
                End If
                description = description & ", RF"
                If InsulatingBox.Value Then
                    description = description & ", TYPE NCA, FULL FACED 1/8"" THK ""LE"" PHENOLIC INSULATING " & _
                        "GASKET, COMPLETE WITH HIGH DENSITY POLYETHELENE SLEEVES FOR " & BOLTSCHEDULE.BoltDiameter(SizeMain_Box.Value, GasketClass.ListIndex) & _
                        " DIA. STUD/BOLTS, DUAL INSULATING WASHERS AND DUAL ELECTRO-PLATED STEEL WASHERS"
                Else
                    description = description & ", FLEXITALLIC, STYLE CGI OR EQUAL, 304 SS WINDING WITH " & _
                        "GRAPHITE FILLER, 304 SS OUTER & INNER RING, 1/8"" THK, FOR MSS-SP-44 FLANGE"
                End If
            
            ' BOLTS
            Else
                description = "BOLT, STUD"
                If SizeMain_Box.MatchFound Then
                    description = description & " - " & BOLTSCHEDULE.BoltDiameter(SizeMain_Box.Value, GasketClass.ListIndex) & " DIA."
                Else
                    MsgBox "Please select a valid nominal diameter.", vbExclamation
                    GoTo INCOMPLETE
                End If
                If GBFaceType.MatchFound Then
                    description = description & " x " & BOLTSCHEDULE.BoltLength(SizeMain_Box.Value, FaceTypeCode(GBFaceType.Value), GasketClass.ListIndex) & " LG."
                Else
                    MsgBox "Please select a valid face type.", vbExclamation
                    GoTo INCOMPLETE
                End If
                description = description & ", THREAD TO THREAD, ALLOY STEEL, ASTM A193, GRB7, CLASS 2A, RTAHT, " & _
                    "W/ 2 HH NUTS EACH, ASTM A194, CLASS 2B, GR 2H, FOR " & SizeMain_Box.Value
                If GasketClass.MatchFound Then
                    description = description & ", " & GasketClass.Value
                Else
                    MsgBox "Please select a valid pressure class.", vbExclamation
                    GoTo INCOMPLETE
                End If
                description = description & ", " & FaceTypeCode(GBFaceType.Value) & " FLANGE"
                description = description & " (" & BOLTSCHEDULE.BoltQuantity(SizeMain_Box.Value, GasketClass.ListIndex) & " BOLTS PER SET)"
                If InsulatingBox.Value Then
                    description = description & " (INSULATING SET)"
                End If
                If CoatedBox.Value Then
                    description = description & ", TEFLON COATED"
                End If
            End If
    
    End Select
    
    DescriptionBox.Value = description
    Exit Sub
    
INCOMPLETE:
    DescriptionBox.Value = ""

End Sub

Private Sub SaveButton_Click()
'Subroutine: SaveButton_Click - UserForm event handler. Called when user clicks RefreshButton.
'                               Sends the description to a hidden holding place to be picked up
'                               by the AddItemWindow or equivalent.

    If FinalFrame.Visible Then
        VB_VAR_STORE.SetAutoDescription DescriptionBox.Value
        Unload Me
    End If
End Sub

Private Sub MaterialTypeBox_Change()
'Subroutine: MaterialTypeBox_Change - UserForm event handler. Called when the MaterialTypeBox selection changes.
'                                     Renders the appropriate frames.

    
    RenderBWFittingFrame (MaterialTypeBox.ListIndex = 2)
    RenderTHSWFittingFrame (MaterialTypeBox.ListIndex = 3)
    RenderPipeFrame (MaterialTypeBox.ListIndex = 0)
    RenderFlangeFrame (MaterialTypeBox.ListIndex = 1), False
    RenderTHSWTeeEllCapFrame False
    RenderBWTeeEllFrame False
    RenderGasketBoltFrame (MaterialTypeBox.ListIndex = 4)
    
    RenderUserInputFrame MaterialTypeBox.ListIndex = 0 Or _
        MaterialTypeBox.ListIndex = 1 Or _
        MaterialTypeBox.ListIndex = 4
        
    RenderReducing False
    
    If MaterialTypeBox.Value = "TH/SW FITTINGS" Then
        GradeBox.Value = "GR B"
        GradeBox.Locked = True
    Else
        GradeBox.Value = ""
        GradeBox.Locked = False
    End If
    
    If MaterialTypeBox.Value = "BW FITTINGS" Then
        SizeMain_Box.RowSource = "BWNominalOD"
        SizeRed_Box.RowSource = "BWNominalOD"
    ElseIf MaterialTypeBox.Value = "TH/SW FITTINGS" Then
        SizeMain_Box.RowSource = "THSWNominalOD"
        SizeRed_Box.RowSource = "THSWNominalOD"
    Else
        SizeMain_Box.RowSource = "NominalPipeOD"
        SizeRed_Box.RowSource = "NominalPipeOD"
    End If
    
    If BWFittingFrame.Visible Then
        UserInputFrame.Top = BWFittingFrame.Top + BWFittingFrame.height + 12
        BWTeeEllFrame.Top = UserInputFrame.Top + UserInputFrame.height + 4
    ElseIf THSWFittingFrame.Visible Then
        THSWFittingFrame.Top = BWFittingFrame.Top
        UserInputFrame.Top = THSWFittingFrame.Top + THSWFittingFrame.height + 12
        THSWTeeEllCapFrame.Top = UserInputFrame.Top + UserInputFrame.height + 4
    ElseIf PipeFrame.Visible Then
        UserInputFrame.Top = BWFittingFrame.Top
        PipeFrame.Top = UserInputFrame.Top + UserInputFrame.height + 4
    ElseIf FlangeFrame.Visible Then
        UserInputFrame.Top = BWFittingFrame.Top
        FlangeFrame.Top = UserInputFrame.Top + UserInputFrame.height + 4
    ElseIf GasketBoltFrame.Visible Then
        UserInputFrame.Top = BWFittingFrame.Top
        GasketBoltFrame.Top = UserInputFrame.Top + UserInputFrame.height + 4
    End If
End Sub

Private Sub BWFittingType_Change()
'Subroutine: BWFittingType_Change - UserForm event handler. Called when the BWFittingType selection changes.
'                                   Renders the appropriate frames and set option boxes if appropriate


    If Not BWFittingFrame.Visible Then Exit Sub
    
    BWOption1.Visible = False
    BWOption2.Visible = False
    BWOption1.Value = False
    BWOption2.Value = False
    RenderBWTeeEllFrame (BWFittingType.Value = "ELL")
    
    If BWFittingType.Value = "TEE" Then
        RenderUserInputFrame False
        BWOption1.Visible = True
        BWOption2.Visible = True
        BWOption1.Caption = "STRAIGHT"
        BWOption2.Caption = "REDUCING"
    ElseIf BWFittingType.Value = "REDUCER" Then
        RenderUserInputFrame False
        BWOption1.Visible = True
        BWOption2.Visible = True
        BWOption1.Caption = "CONCENTRIC"
        BWOption2.Caption = "ECCENTRIC"
    ElseIf BWFittingType.ListIndex >= 0 Then
        RenderReducing False
        RenderUserInputFrame True
    End If
    
End Sub

Private Sub BWOption1_Click()
'Subroutine: BWOption1_Click - UserForm event handler. Called when the user clicks BWOption1

    If Not UserInputFrame.Visible Then
        RenderUserInputFrame BWFittingType.ListIndex >= 0 And _
            (BWOption1.Visible And BWOption1.Value Or _
                BWOption2.Visible And BWOption2.Value)
    End If

    If BWFittingType.Value = "TEE" Then
        RenderBWTeeEllFrame (BWOption1.Visible And BWOption1.Value Or _
            BWOption2.Visible And BWOption2.Value)
    End If

    RenderReducing (BWOption2.Caption = "REDUCING" And BWOption2.Value) Or BWFittingType.Value = "REDUCER"
    
    RenderFinalFrame BWFittingType.ListIndex >= 0 And _
        (BWOption1.Visible And BWOption1.Value Or _
            BWOption2.Visible And BWOption2.Value)
End Sub

Private Sub BWOption2_Click()
'Subroutine: BWOption2_Click - UserForm event handler. Called when the user clicks BWOption2

    If Not UserInputFrame.Visible Then
        RenderUserInputFrame BWFittingType.ListIndex >= 0 And _
            (BWOption1.Visible And BWOption1.Value Or _
                BWOption2.Visible And BWOption2.Value)
    End If
    
    If BWFittingType.Value = "TEE" Then
        RenderBWTeeEllFrame (BWOption1.Visible And BWOption1.Value Or _
            BWOption2.Visible And BWOption2.Value)
    End If
    
    RenderReducing (BWOption2.Caption = "REDUCING" And BWOption2.Value) Or BWFittingType.Value = "REDUCER"
    
    RenderFinalFrame BWFittingType.ListIndex >= 0 And _
        (BWOption1.Visible And BWOption1.Value Or _
            BWOption2.Visible And BWOption2.Value)
End Sub

Private Sub GasketOption_Change()
'Subroutine: GasketOption_Change - Userform event handler. Called whenever the state of GasketOption changes.
'                                  Resets options within the GasketBoltFrame

    BoltOption.Value = Not GasketOption.Value
    CoatedLabel.Visible = Not GasketOption.Value
    CoatedBox.Visible = Not GasketOption.Value
End Sub

Private Sub THSWFittingType_Change()
'Subroutine: THSWFittingType_Change - UserForm event handler. Called when the THSWFittingType selection changes.
'                                     Renders the appropriate frames and set option boxes if appropriate

    If Not THSWFittingFrame.Visible Then Exit Sub
    
    THSWOption1.Visible = False
    THSWOption2.Visible = False
    THSWOption1.Value = False
    THSWOption2.Value = False
    RenderTHSWTeeEllCapFrame THSWFittingType.Value <> "TEE"
    
    If THSWFittingType.Value = "TEE" Then
        RenderUserInputFrame False
        THSWOption1.Visible = True
        THSWOption2.Visible = True
        THSWOption1.Caption = "STRAIGHT"
        THSWOption2.Caption = "REDUCING"
    ElseIf THSWFittingType.ListIndex >= 0 Then
        RenderReducing False
        RenderUserInputFrame True
    End If
    
End Sub

Private Sub THSWOption1_Click()
'Subroutine: THSWOption1_Click - UserForm event handler. Called when the user clicks THSWOption1

    If Not UserInputFrame.Visible Then
        RenderUserInputFrame THSWFittingType.ListIndex >= 0 And _
            (THSWOption1.Visible And THSWOption1.Value Or _
               THSWOption2.Visible And THSWOption2.Value)
    End If
    
    If THSWFittingType.Value = "TEE" Then
        RenderTHSWTeeEllCapFrame (THSWOption1.Visible And THSWOption1.Value Or _
            THSWOption2.Visible And THSWOption2.Value)
    End If
    
    RenderReducing (THSWOption2.Caption = "REDUCING" And THSWOption2.Value) Or THSWFittingType.Value = "REDUCER"
    
    RenderFinalFrame THSWFittingType.ListIndex >= 0 And _
        (THSWOption1.Visible And THSWOption1.Value Or _
            THSWOption2.Visible And THSWOption2.Value)
End Sub

Private Sub THSWOption2_Click()
'Subroutine: THSWOption2_Click - UserForm event handler. Called when the user clicks THSWOption2

    If Not UserInputFrame.Visible Then
        RenderUserInputFrame THSWFittingType.ListIndex >= 0 And _
            (THSWOption1.Visible And THSWOption1.Value Or _
                THSWOption2.Visible And THSWOption2.Value)
    End If
    
    If THSWFittingType.Value = "TEE" Then
        RenderTHSWTeeEllCapFrame (THSWOption1.Visible And THSWOption1.Value Or _
            THSWOption2.Visible And THSWOption2.Value)
    End If
    
    RenderReducing (THSWOption2.Caption = "REDUCING" And THSWOption2.Value) Or THSWFittingType.Value = "REDUCER"
    
    RenderFinalFrame THSWFittingType.ListIndex >= 0 And _
        (THSWOption1.Visible And THSWOption1.Value Or _
            THSWOption2.Visible And THSWOption2.Value)
End Sub

Private Sub SizeMain_Box_Change()
'Subroutine: SizeMain_Box_Change - UserForm event handler. Called when the SizeMain_Box selection changes.
'                                  Refreshes the WallThickness selections

    If Not UserInputFrame.Visible Then Exit Sub
    
    Dim row As Integer
    row = PIPESCHEDULE.FirstRow()
    
    With PIPESCHEDULE
        
        Do While Not IsEmpty(.Cells(row, 1))
            If .Cells(row, 1).Value2 = SizeMain_Box.Value Then
                Exit Do
            End If
            row = row + 1
        Loop
        
    End With
    
    Dim col As Integer
    col = 3
    
    WTMain_Box.enabled = False
    WTMain_Box.Clear
    
    Dim WTs As Collection
    Set WTs = New Collection
    
    With PIPESCHEDULE
        
        Do While col <= .UsedRange.Columns.count
            If IsNumeric(.Cells(row, col).Value2) And Len(.Cells(row, col).text) > 0 And Not CollectionContains(WTs, .Cells(row, col).text) Then
                WTMain_Box.AddItem .Cells(row, col).text & """"
                WTs.Add .Cells(row, col).text
            End If
            col = col + 1
        Loop
        
    End With
    
    Set WTs = Nothing
    WTMain_Box.enabled = True
End Sub

Private Sub WTMain_Box_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
'Subroutine: WTMain_Box_Change - UserForm event handler. Called when the WTMain_Box updates.
'                                 Refreshes the Wall Thickness number format

    If Right(WTMain_Box.Value, 1) <> """" Then
        WTMain_Box.Value = WTMain_Box.Value & """"
    End If
    
    If IsNumeric(Left(WTMain_Box.Value, Len(WTMain_Box.Value) - 1)) Then
        WTMain_Box.Value = IIf(WTMain_Box.Value = "0""", "", Format(Left(WTMain_Box.Value, Len(WTMain_Box.Value) - 1), "0.000")) & """"
    Else
        MsgBox "Wall Thickness must be a numeric value in inches (i.e. 0.250"")", vbExclamation
        WTMain_Box.Value = ""
    End If
End Sub

Private Sub SizeRed_Box_Change()
'Subroutine: SizeRed_Box_Change - UserForm event handler. Called when the SizeRed_Box selection changes.
'                                 Refreshes the WallThickness selections

    If Not UserInputFrame.Visible Then Exit Sub
    
    Dim row As Integer
    row = PIPESCHEDULE.FirstRow()
    
    With PIPESCHEDULE
        
        Do While Not IsEmpty(.Cells(row, 1))
            If .Cells(row, 1).Value2 = SizeRed_Box.Value Then
                Exit Do
            End If
            row = row + 1
        Loop
        
    End With
    
    Dim col As Integer
    col = 3
    
    WTRed_Box.enabled = False
    WTRed_Box.Clear
    
    Dim WTs As Collection
    Set WTs = New Collection
    
    Do While col <= PIPESCHEDULE.UsedRange.Columns.count
        If IsNumeric(PIPESCHEDULE.Cells(row, col).Value2) And Len(PIPESCHEDULE.Cells(row, col).Value2) > 0 And Not CollectionContains(WTs, PIPESCHEDULE.Cells(row, col).text) Then
            WTRed_Box.AddItem PIPESCHEDULE.Cells(row, col).text & """"
            WTs.Add PIPESCHEDULE.Cells(row, col).text
        End If
        col = col + 1
    Loop
    
    Set WTs = Nothing
    WTRed_Box.enabled = True
End Sub

Private Sub WTRed_Box_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
'Subroutine: WTRed_Box_Change - UserForm event handler. Called when the WTRed_Box updates.
'                                 Refreshes the Wall Thickness number format

    If Right(WTRed_Box.Value, 1) <> """" Then
        WTRed_Box.Value = WTRed_Box.Value & """"
    End If
    
    If IsNumeric(Left(WTRed_Box.Value, Len(WTRed_Box.Value) - 1)) Then
        WTRed_Box.Value = IIf(WTRed_Box.Value = "0""", "", Format(Left(WTRed_Box.Value, Len(WTRed_Box.Value) - 1), "0.000")) & """"
    Else
        MsgBox "Wall Thickness must be a numeric value in inches (i.e. 0.250"")", vbExclamation
        WTRed_Box.Value = ""
    End If
End Sub

Private Sub GradeBox_Change()
'Subroutine: GradeBox_Change - UserForm event handler. Called when the GradeBox selection changes.
'                              Refreshes the Pipe Spec selections

    If Not GradeBox.MatchFound Or Not PipeFrame.Visible Then
        Exit Sub
    End If
    
    SpecBox.Clear
    
    Dim row As Integer
    row = MATGRADE.FirstRow() + GradeBox.ListIndex
    
    Dim col As Integer
    col = MATGRADE.SpecListColumn()
    Do While Not IsEmpty(MATGRADE.Cells(row, col))
        SpecBox.AddItem MATGRADE.Cells(row, col).Value2
        col = col + 1
    Loop
End Sub

Private Sub SpecBox_DropButtonClick()
'Subroutine: SpecBox_DropButtonClick - UserForm event handler. Called when the user clicks the SpecBox drop-down button

    If SpecBox.ListCount = 0 Then
        MsgBox "Please select a valid Material Grade before setting the Pipe Spec. Thanks.", vbExclamation
        GradeBox.SetFocus
        Exit Sub
    End If
End Sub

Private Sub BlindBox_Change()
'Subroutine: BlindBox_Change - UserForm event handler. Called when the BlindBox value changes.

    CDTBox.Visible = BlindBox.Value
    CDTLabel.Visible = BlindBox.Value
    CDTBox.Value = ""
End Sub

Private Sub RenderReducing(ByVal Visible As Boolean)
'Subroutine: RenderReducing - Formats visible frames to accomodate reducing fittings
'Arguments: visible - Visibility state to set

    Dim change As Boolean
    change = RedSize_Label.Visible <> Visible

    RedSize_Label.Visible = Visible
    RedOD_Label.Visible = Visible
    RedWT_Label.Visible = Visible
    SizeRed_Box.Visible = Visible
    WTRed_Box.Visible = Visible
    SizeRed_Box.Value = ""
    WTRed_Box.Value = ""
    
    If change Then
        If Visible Then
            Grade_Label.Left = Grade_Label.Left + RedSize_Label.Width
            GradeBox.Left = GradeBox.Left + RedSize_Label.Width
        Else
            Grade_Label.Left = Grade_Label.Left - RedSize_Label.Width
            GradeBox.Left = GradeBox.Left - RedSize_Label.Width
        End If
    End If
End Sub

Private Sub RenderFinalFrame(ByVal Visible As Boolean)
'Subroutine: RenderFinalFrame - Formats FinalFrame
'Arguments: visible - Visibility state to set

    If FinalFrame.Visible <> Visible Then
        FinalFrame.Visible = Visible
        If Visible Then
            RefreshFormSize FinalFrame.height
        Else
            RefreshFormSize -1 * FinalFrame.height
        End If
    End If

    DescriptionBox.Value = ""
End Sub

Private Sub RenderGasketBoltFrame(ByVal Visible As Boolean)
'Subroutine: RenderGasketBoltFrame - Formats GasketBoltFrame
'Arguments: visible - Visibility state to set

    If GasketBoltFrame.Visible <> Visible Then
        GasketBoltFrame.Visible = Visible
        If Visible Then
            RefreshFormSize GasketBoltFrame.height
        Else
            RefreshFormSize -1 * GasketBoltFrame.height
        End If
    End If
    
    GasketOption.Value = True
    BoltOption.Value = False
    InsulatingBox.Value = False
    CoatedBox.Value = False
    
    GBFaceType.Value = ""
    GasketClass.Value = ""
End Sub

Private Sub RenderTHSWTeeEllCapFrame(ByVal Visible As Boolean)
'Subroutine: RenderTHSWTeeEllCapFrame - Formats THSWTeeEllCapFrame
'Arguments: visible - Visibility state to set

    If THSWTeeEllCapFrame.Visible <> Visible Then
        THSWTeeEllCapFrame.Visible = Visible
        If Visible Then
            RefreshFormSize THSWTeeEllCapFrame.height
        Else
            RefreshFormSize -1 * THSWTeeEllCapFrame.height
        End If
    End If
    
    ThreadedOption.Value = False
    SocketWeldOption.Value = False
    THSWClassBox.Value = ""
    THSWDegreeBox.Value = ""
    
    If THSWFittingType.Value = "ELL" Then
        THSWDegreeLabel.Visible = True
        THSWDegreeBox.Visible = True
    Else
        THSWDegreeLabel.Visible = False
        THSWDegreeBox.Visible = False
    End If
    
    If THSWFittingType.Value = "PLUG" Then
        ThreadedOption.Visible = False
        SocketWeldOption.Visible = False
    Else
        ThreadedOption.Visible = True
        SocketWeldOption.Visible = True
    End If
End Sub

Private Sub RenderBWTeeEllFrame(ByVal Visible As Boolean)
'Subroutine: RenderBWTeeEllFrame - Formats BWTeeEllFrame
'Arguments: visible - Visibility state to set

    If BWTeeEllFrame.Visible <> Visible Then
        BWTeeEllFrame.Visible = Visible
        If Visible Then
            RefreshFormSize BWTeeEllFrame.height
        Else
            RefreshFormSize -1 * BWTeeEllFrame.height
        End If
    End If
    
    BWCheckOptionBox.Value = False
    BWEllDegree.Value = ""
    BWEllRadius.Value = ""
    
    If BWFittingType.Value = "ELL" Then
        BWCheckOptLabel.Caption = "Segmentable?:"
        BWEllDegreeLabel.Visible = True
        BWEllDegree.Visible = True
        BWEllRadiusLabel.Visible = True
        BWEllRadius.Visible = True
    Else
        BWCheckOptLabel.Caption = "Barred?:"
        BWEllDegreeLabel.Visible = False
        BWEllDegree.Visible = False
        BWEllRadiusLabel.Visible = False
        BWEllRadius.Visible = False
    End If
End Sub

Private Sub RenderFlangeFrame(ByVal Visible As Boolean, ByVal blind As Boolean)
'Subroutine: RenderFlangeFrame - Formats FlangeFrame
'Arguments: visible - Visibility state to set

    If FlangeFrame.Visible <> Visible Then
        FlangeFrame.Visible = Visible
        If Visible Then
            RefreshFormSize FlangeFrame.height
            FlangeClassBox.Value = ""
            FlangeFaceType.Value = ""
            FlangeEndType.Value = ""
        Else
            RefreshFormSize -1 * FlangeFrame.height
        End If
    End If
    
    CDTBox.Value = ""
    BlindBox.Value = blind
    If CDTBox.Visible <> blind Then
        CDTBox.Visible = blind
        CDTLabel.Visible = blind
    End If
End Sub

Private Sub RenderUserInputFrame(ByVal Visible As Boolean)
'Subroutine: RenderUserInputFrame - Formats UserInputFrame
'Arguments: visible - Visibility state to set

    If UserInputFrame.Visible <> Visible Then
        UserInputFrame.Visible = Visible
        If Visible Then
            RefreshFormSize UserInputFrame.height
        Else
            RefreshFormSize -1 * UserInputFrame.height
        End If
    End If

    SizeMain_Box.Value = ""
    SizeRed_Box.Value = ""
    WTMain_Box.Value = ""
    WTRed_Box.Value = ""
    SpecBox.Value = ""
    
    If Not GradeBox.Locked Then
        GradeBox.Value = ""
    End If
    
    RenderFinalFrame Visible
End Sub

Private Sub RenderBWFittingFrame(ByVal Visible As Boolean)
'Subroutine: RenderBWFittingFrame - Formats BWFittingFrame
'Arguments: visible - Visibility state to set

    If BWFittingFrame.Visible <> Visible Then
        BWFittingFrame.Visible = Visible
        If Visible Then
            RefreshFormSize BWFittingFrame.height
        Else
            RefreshFormSize -1 * BWFittingFrame.height
        End If
    End If

    BWFittingType.Value = ""
    BWOption1.Value = False
    BWOption2.Value = False
    BWOption1.Visible = False
    BWOption2.Visible = False
End Sub

Private Sub RenderTHSWFittingFrame(ByVal Visible As Boolean)
'Subroutine: RenderTHSWFittingFrame - Formats THSWFittingFrame
'Arguments: visible - Visibility state to set

    If THSWFittingFrame.Visible <> Visible Then
        THSWFittingFrame.Visible = Visible
        If Visible Then
            RefreshFormSize THSWFittingFrame.height
        Else
            RefreshFormSize -1 * THSWFittingFrame.height
        End If
    End If
    
    THSWFittingType.Value = ""
    THSWOption1.Value = False
    THSWOption2.Value = False
    THSWOption1.Visible = False
    THSWOption2.Visible = False
End Sub

Private Sub RenderPipeFrame(ByVal Visible As Boolean)
'Subroutine: RenderPipeFrame - Formats PipeFrame
'Arguments: Visible - Visibility state to set

    If PipeFrame.Visible <> Visible Then
        PipeFrame.Visible = Visible
        If Visible Then
            RefreshFormSize PipeFrame.height
        Else
            RefreshFormSize -1 * PipeFrame.height
        End If
    End If

    PipeMakeBox.Value = ""
    PipeEndTypeBox.Value = ""
    PipeCoatingBox.Value = ""
    PipeCertifiedBox.Value = False
End Sub

Private Sub RefreshFormSize(ByVal height_change As Integer)
'Subroutine: RefreshFormSize - Refresh the height of the form and location of FinalFrame
'Arguments: height_change - Integer containing the change in heihgt

    FinalFrame.Top = FinalFrame.Top + height_change
    Me.height = Me.height + height_change
End Sub

Private Function PipeEndTypeCode(ByVal end_type As String) As String
'Function: PipeEndTypeCode - Returns the 2 character End Type code
'Arguments: end_type - String containing the end type from the drop down box
'Returns: String containing the respective two character code

    If end_type = "Plain End" Then
        PipeEndTypeCode = "PE"
    ElseIf end_type = "Bevel End" Then
        PipeEndTypeCode = "BE"
    End If
End Function

Private Function FaceTypeCode(ByVal face_type As String) As String
'Function: FaceTypeCode - Returns the Face Type code
'Arguments: face_type - String containing the Face type from the drop down box
'Returns: String containing the respective code

    If face_type = "Raised Face" Then
        FaceTypeCode = "RF"
    ElseIf face_type = "Ring Type" Then
        FaceTypeCode = "RTJ"
    ElseIf face_type = "Flat Face" Then
        FaceTypeCode = "FF"
    ElseIf face_type = "Lap Joint" Then
        FaceTypeCode = "LJ"
    End If
End Function

Private Function ConnectionTypeCode(ByVal end_type As String) As String
'Function: ConnectionTypeCode - Returns the Connection Type code
'Arguments: end_type - String containing the Connection type from the drop down box
'Returns: String containing the respective code

    If end_type = "Weld Neck" Then
        ConnectionTypeCode = "WN"
    ElseIf end_type = "Socket Weld" Then
        ConnectionTypeCode = "SW"
    ElseIf end_type = "Threaded" Then
        ConnectionTypeCode = "NPT"
    End If
End Function

