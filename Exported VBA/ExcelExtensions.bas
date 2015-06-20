Attribute VB_Name = "ExcelExtensions"
Option Explicit

Public Function HMMFileTag(Optional tag_format As String = vbNullString) As String
'Function: HMMFileTag - Return a formatted tag with either the format given in the
'                       argument, or with the format expression stored in VB_VAR_STORE.
'Arguments: tag_format - Optional format to specify
'Returns: String containing the formatted file tag


    Dim this_format As String
    
    If tag_format = vbNullString Then
        this_format = VB_VAR_STORE.HMMFileTagFormat()
    Else
        this_format = tag_format
    End If
    
    HMMFileTag = Format(Now(), this_format)
    
    Dim HMMattribute As String
    Dim index As Integer
    index = 0
    
    Do While index < Len(HMMFileTag) - 1
        If Mid(HMMFileTag, index + 1, 1) = "~" Then
            HMMattribute = FirstWord(Mid(HMMFileTag, index + 1))
        End If
        
        If HMMattribute = "ProjectName" Then
            HMMFileTag = Left(HMMFileTag, index) & VB_COVERSHEET.ProjectName() & Mid(HMMFileTag, index + Len(HMMattribute) + 2)
        ElseIf HMMattribute = "ProjectClientName" Then
            HMMFileTag = Left(HMMFileTag, index) & VB_COVERSHEET.ProjectClientName() & Mid(HMMFileTag, index + Len(HMMattribute) + 2)
        ElseIf HMMattribute = "ProjectNumber" Then
            HMMFileTag = Left(HMMFileTag, index) & VB_COVERSHEET.ProjectNumber() & Mid(HMMFileTag, index + Len(HMMattribute) + 2)
        End If
        
        HMMattribute = ""
        index = index + 1
    Loop
    
End Function

Public Function DescriptionCompare(ByVal str_1 As String, ByVal str_2 As String) As Integer
'Function: DescriptionCompare - Adaptation of StrComp as it applies to material descriptions. This function compares
'                               strings word by word, evaluating decimal numbers and fractions as actual floating
'                               point numbers, not strings (i.e. 'PIPE 1/2"'  = 'PIPE 0.500"'
'                                                                'PIPE 0.28"' = 'PIPE 0.280"'
'                                                                'PIPE 1/2"'  > 'PIPE 1/4"' )
'Arguments: str_1 - String containing the target comparison String
'           str_2 - String containing the String to compare the target to
'Returns: Integer containing the comparison code:
'           0 - str_1 = str_2
'           1 - str_1 > str_2
'          -1 - str_1 < str_2


    'check if strings are equal, for easy return
    If StrComp(str_1, str_2, vbTextCompare) = 0 Then
        DescriptionCompare = 0
        Exit Function
    End If
    
    'check length of arguments, for easy return
    If Len(str_1) = 0 Then
        DescriptionCompare = -1
        Exit Function
    ElseIf Len(str_2) = 0 Then
        DescriptionCompare = 1
        Exit Function
    End If

    Dim str1 As String
    Dim str2 As String
    str1 = str_1
    str2 = str_2
    
    Dim delim1 As String
    Dim delim2 As String
    Dim parse1 As String
    Dim parse2 As String
    Dim index As Integer
    index = 0
    DescriptionCompare = 0
    
    delim1 = TrimDelim_Front(str1)
    delim2 = TrimDelim_Front(str2)
    ''' DO NOT COMPARE PUNCTUATION
    'If Len(delim1) > 0 Or Len(delim2) > 0 Then
    '    DescriptionCompare = StrComp(delim1, delim2)
    'End If

    'parse strings, evaluate, and compare
    Do While DescriptionCompare = 0 And Len(str1) > 0 And Len(str2) > 0
        parse1 = FirstWord(str1)
        parse2 = FirstWord(str2)
        
        str1 = Right(str1, Len(str1) - Len(parse1))
        str2 = Right(str2, Len(str2) - Len(parse2))
        
        'if parses are numeric
        If IsNumeric(parse1) And IsNumeric(parse2) Then
            Dim num1 As Double
            Dim num2 As Double
            num1 = CDbl(parse1)
            num2 = CDbl(parse2)
            
            delim1 = TrimDelim_Front(str1)
            delim2 = TrimDelim_Front(str2)
            
            'handle decimal or fraction in parse1
            If delim1 = "." Or delim1 = "/" Then
                parse1 = FirstWord(str1)
                str1 = Right(str1, Len(str1) - Len(parse1))
                
                If delim1 = "." Then
                    num1 = num1 + CDbl(parse1) / 10 ^ Len(parse1)
                ElseIf delim1 = "/" Then
                    num1 = num1 / CDbl(parse1)
                End If
            ElseIf delim1 = " " Then
                Dim numer1 As String
                Dim temp_str1 As String
                Dim temp_delim1 As String
                temp_str1 = str1
                
                numer1 = FirstWord(temp_str1)
                temp_str1 = Right(temp_str1, Len(temp_str1) - Len(numer1))
                
                If IsNumeric(numer1) And Left(temp_str1, 1) = "/" Then
                    Dim denom1 As String
                    delim1 = TrimDelim_Front(temp_str1)
                    denom1 = FirstWord(temp_str1)
                    temp_str1 = Right(temp_str1, Len(temp_str1) - Len(denom1))
                    
                    If IsNumeric(denom1) Then
                        num1 = num1 + CDbl(numer1) / CDbl(denom1)
                        str1 = temp_str1
                    End If
                End If
            End If
            
            'handle decimal or fraction in parse1
            If delim2 = "." Or delim2 = "/" Then
                parse2 = FirstWord(str2)
                str2 = Right(str2, Len(str2) - Len(parse2))
                
                If delim2 = "." Then
                    num2 = num2 + CDbl(parse2) / 10 ^ Len(parse2)
                ElseIf delim2 = "/" Then
                    num2 = num2 / CDbl(parse2)
                End If
            ElseIf delim2 = " " Then
                Dim numer2 As String
                Dim temp_str2 As String
                Dim temp_delim2 As String
                temp_str2 = str2
                
                numer2 = FirstWord(temp_str2)
                temp_str2 = Right(temp_str2, Len(temp_str2) - Len(numer2))
                
                If IsNumeric(numer2) And Left(temp_str2, 1) = "/" Then
                    Dim denom2 As String
                    delim2 = TrimDelim_Front(temp_str2)
                    denom2 = FirstWord(temp_str2)
                    temp_str2 = Right(temp_str2, Len(temp_str2) - Len(denom2))
                    
                    If IsNumeric(denom2) Then
                        num2 = num2 + CDbl(numer2) / CDbl(denom2)
                        str2 = temp_str2
                    End If
                End If
            End If
            
            delim1 = TrimDelim_Front(str1)
            delim2 = TrimDelim_Front(str2)
            
            'compare numerical values
            If num1 > num2 Then
                DescriptionCompare = 1
            ElseIf num1 < num2 Then
                DescriptionCompare = -1
            Else
                DescriptionCompare = 0
            End If
            
        'GR B = GRB
        ElseIf parse1 Like "GR*" And parse2 Like "GR*" Then
            Dim temp_parse As String
            If parse1 = "GR" And (str1 Like " ? *" Or str1 Like " ?, *") Then
                TrimDelim_Front str1
                temp_parse = FirstWord(str1)
                parse1 = parse1 + temp_parse
                str1 = Right(str1, Len(str1) - Len(temp_parse))
            ElseIf parse2 = "GR" And (str2 Like " ? *" Or str2 Like " ?, *") Then
                TrimDelim_Front str2
                temp_parse = FirstWord(str2)
                parse2 = parse2 + temp_parse
                str2 = Right(str2, Len(str2) - Len(temp_parse))
            End If
            DescriptionCompare = StrComp(parse1, parse2, vbTextCompare)
            
        'FLANGE, BLIND' < 'FLANGE - 2"'
        ElseIf parse1 = "BLIND" And parse2 <> "BLIND" Then
            DescriptionCompare = -1
        ElseIf parse1 <> "BLIND" And parse2 = "BLIND" Then
            DescriptionCompare = 1
            
        ElseIf parse1 = "ECC" And parse2 = "CONC" Then
            DescriptionCompare = -1
        ElseIf parse1 = "CONC" And parse2 = "ECC" Then
            DescriptionCompare = 1
            
        ElseIf parse1 = "COATED" And parse2 = "BARE" Then
            DescriptionCompare = -1
        ElseIf parse1 = "BARE" And parse2 = "COATED" Then
            DescriptionCompare = 1
            
        Else
            DescriptionCompare = StrComp(parse1, parse2, vbTextCompare)
        End If
        
        delim1 = TrimDelim_Front(str1)
        delim2 = TrimDelim_Front(str2)
        
        ''' DO NOT COMPARE PUNCTUATION
        'If DescriptionCompare = 0 And (Len(delim1) > 0 Or Len(delim1) > 0) Then
        '    DescriptionCompare = StrComp(delim1, delim2)
        'End If
    Loop
    
End Function

Public Function RenderUI(ByVal enable As Boolean) As Boolean
'Function: RenderUI - Controls value of Application.ScreenUpdating and Application.EnableEvents
'                     Used while making changes to GUI - RenderUI(False) to hold graphics ... RenderUI(True) to repaint
'                     If returns False, no changes were made to the state of .ScreenUpdating and .EnableEvents - UI is being rendered elsewhere.
'Arguments: enable - Boolean. State to change ScreenUpdating and EnableEvents to.
'Returns: Boolean - whether a change to the state was made. True - State change to (enable); False - State is already equal to (enable).
    
    
    With Application
        If .ScreenUpdating = enable And .EnableEvents = enable Then
            RenderUI = False
        Else
            .ScreenUpdating = enable
            .EnableEvents = enable
            RenderUI = True
        End If
    End With
    
End Function

Public Sub SetStatusBar(ByVal message As String, Optional ByVal progress_part As Double = -1, Optional ByVal progress_whole As Double = -1)
'Subroutine: SetStatusBar - Sets a message in Excel's Status Bar. It will also include progress
'                           bars, if you provide the % complete in decimal form.
'                           If progress > 1, Progress bars will show at 100%
'                           If progress is < 0, progress bars will not show.
'Arguments: message - String containing the Status Bar message
'           progress_part - Double value representing the "part of the whole" to calculation the progress percentage.
'           progress_whole - Double value representing the "whole" to calculation the progress percentage.


    Dim progress As Double
    If Not (progress_part < 0 Or progress_whole < 0) Then
        If progress_whole = 0 Then
            If progress_part = 0 Then
                progress = 0
            Else
                progress = 1
            End If
        Else
            progress = progress_part / progress_whole
        End If
    Else
        progress = -1
    End If
    
    Dim bars As Integer
    bars = 0
    
    Application.StatusBar = message
    
    If 0 <= progress Then
        If progress > 1 Then progress = 1
        
        bars = Round(progress * 10)
        
        Application.StatusBar = "|" & String(bars, ChrW(9609)) & String(10 - bars, "_") & "| " & message
    End If
    
End Sub

Public Sub ResetStatusBar()
'Subroutine: SetStatusBar - Clears the message in Excel's Status Bar.

    Application.StatusBar = False
End Sub

Public Function IsValidDirectory(ByVal subdirectory As String) As Boolean
'Function: IsValidDriectory - Checks subdirectory to see if it is a valid folder loaction. In effect, if
'                             the subdirectory exists, or if the application can make it and delete it again, it's valid.
'Arguments: subdirectory - String to check.
'Returns: Boolean. True - valid directory; False - is NOT valid


    If Left(subdirectory, 1) <> "\" Then
        subdirectory = "\" & subdirectory
    End If
    
    On Error GoTo GOTOFALSE
    
    If Dir(ThisWorkbook.Path & subdirectory, vbDirectory) <> vbNullString Then
        ' path exists
        IsValidDirectory = True
        Exit Function
    End If
    
    Dim count_new As Integer
    count_new = MakeDirs(ThisWorkbook.Path & subdirectory)
    
    Dim mDir As String
    Dim i As Long
    Dim aryDirs As Variant
    mDir = ThisWorkbook.Path & subdirectory
    If Right(mDir, 1) = "\" Then
        mDir = Mid(mDir, 1, Len(mDir) - 1)
    End If
    aryDirs = Split(mDir, "\")
    
    For i = UBound(aryDirs) To UBound(aryDirs) - count_new + 1 Step -1
        If Len(aryDirs(i)) = 0 Or aryDirs(i) = ".." Then GoTo SKIPTONEXT
        
        If Dir(mDir, vbDirectory) <> vbNullString Then
            RmDir mDir
        End If
        
SKIPTONEXT:
        mDir = Left(mDir, Len(mDir) - Len(aryDirs(i)) - 1)
    Next i
    
    Err.Clear
    On Error GoTo 0
    
    IsValidDirectory = True
    Exit Function
    
GOTOFALSE:
    IsValidDirectory = False
    
End Function

Public Function MakeDirs(ByVal FullName As String) As Integer
'Function: MakeDirs - Uses MkDir to make multiple nested directories
'Arguments: FullName - String to make.
'Returns: Integer containing the number of sirectories made

    Dim aryDirs As Variant
    Dim mDir As String
    Dim i As Long
    MakeDirs = 0
    
    aryDirs = Split(FullName, "\")
    mDir = CStr(aryDirs(LBound(aryDirs)))
    
    For i = LBound(aryDirs) + 1 To UBound(aryDirs)
        mDir = mDir & "\" & aryDirs(i)
        
        If Len(aryDirs(i)) = 0 Or aryDirs(i) = ".." Then GoTo SKIPTONEXT
        
        If Dir(mDir, vbDirectory) = vbNullString Then
            MkDir mDir
            MakeDirs = MakeDirs + 1
        End If
        
SKIPTONEXT:
    Next i
End Function

Public Function GetRelativeFolderPath(ByVal full_path As String) As String
'Function: GetRelativeFolderPath - Calculates a path string relative to ThisWorkbook.Path
'Arguments: full_path - String containing the path to calculate
'Returns: String containing the relative path

    If Len(full_path) = 0 Then
        GetRelativeFolderPath = vbNullString
        Exit Function
    End If
    
    GetRelativeFolderPath = "\"

    Dim aryPath As Variant
    aryPath = Split(full_path, "\")
    
    Dim aryWorkbook As Variant
    aryWorkbook = Split(ThisWorkbook.Path, "\")
    
    Dim PathDir As String
    Dim WkbkDir As String
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    For i = LBound(aryPath) To UBound(aryPath)
        PathDir = CStr(aryPath(i))
        If i <= UBound(aryWorkbook) Then
            WkbkDir = CStr(aryWorkbook(i))
        Else
            Exit For
        End If
        
        If i = UBound(aryPath) And i < UBound(aryWorkbook) Then
            i = i + 1
            For j = i To UBound(aryWorkbook)
                WkbkDir = CStr(aryWorkbook(j))
                If Len(WkbkDir) > 0 Then
                    GetRelativeFolderPath = GetRelativeFolderPath & "..\"
                End If
            Next j
            
            Exit For
        ElseIf PathDir <> WkbkDir Then
            ' add "..\"
            For j = i To UBound(aryWorkbook)
                WkbkDir = CStr(aryWorkbook(j))
                If Len(WkbkDir) > 0 Then
                    GetRelativeFolderPath = GetRelativeFolderPath & "..\"
                End If
            Next j
            
            Exit For
        End If
    Next i
    
    For k = i To UBound(aryPath)
        PathDir = CStr(aryPath(k))
        If Len(PathDir) > 0 Then
            GetRelativeFolderPath = GetRelativeFolderPath & PathDir & "\"
        End If
    Next k
End Function

Public Function FileFolderExists(strFullPath As String) As Boolean
'Function: FileFolderExists - Check if a file folder exists
'Arguments: strFullPath - String containing the full path of the file/folder to find
'Returns: Boolean. True - File Folder exists; False - File Folder doesn't exist
    
    FileFolderExists = False

    On Error GoTo EarlyExit
    If Not Dir(strFullPath, vbDirectory) = vbNullString Then FileFolderExists = True
    
EarlyExit:
    On Error GoTo 0
End Function

Public Function BrowseForFolder(Optional ByVal startPath As String = vbNullString) As String
'Function: BrowseForFolder - Open dialogue box to browse for a folder
'Returns: String containing the path to the folder selected


    If Len(startPath) = 0 Then
        startPath = ThisWorkbook.Path
    End If
    
    On Error Resume Next
    ChDir startPath
    Err.Clear
    On Error GoTo 0
    
    Dim fldr As FileDialog
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = startPath
        If .Show <> -1 Then GoTo NextCode
        BrowseForFolder = .SelectedItems(1)
    End With
NextCode:
    Set fldr = Nothing
End Function

Public Function ContainsIllegalCharacters(ByVal description As String) As Boolean
'Function: ContainsIllegalCharacters - Checks item description for illegal characters * _ [ ] ^
'Arguments: description - String to check
'Returns: Boolean. True - Contains illegal characters; False - does not contain illegal characters


    Dim i As Integer
    
    For i = 1 To Len(description)
        If Not CheckKeyCode("Description", Asc(Mid(description, i, 1))) Then
            ContainsIllegalCharacters = True
            Exit Function
        End If
    Next i

    ContainsIllegalCharacters = False
End Function

Public Function CheckKeyCode(ByVal chk_type As String, ByVal KeyCode As Integer) As Boolean
'Function: CheckKeyCode - Tests an ASCII character code to see if it is legal. Name, Descriptions, and Units string values
'                         have different constraints on their legal characters. Names often become filenames, Descriptions are
'                         compared using the Like operator, and unit are units.. why not. See below.
'Arguments: chk_type - String dictating the legal contraints. "Name, "Description", and "Unit"
'           KeyCode - ASCII value for character
'Returns: Boolean. True - Checks out, legal. False - illegal.

    ' Name        / ? < > \ : * | "
    ' Description * _ [ ] ^
    ' Unit        (both)
    
    ' Relevant ACSII Codes
    ' / - 47
    ' ? - 63
    ' < - 60
    ' > - 62
    ' \ - 92
    ' : - 58
    ' * - 42
    ' | - 124
    ' " - 34
    ' _ - 95
    ' [ - 91
    ' ] - 93
    ' ^ - 94
    
    CheckKeyCode = True
    
    Dim name_chars
    name_chars = Array(34, 42, 47, 58, 60, 62, 63, 92, 124)
    
    Dim desc_chars
    desc_chars = Array(42, 91, 93, 94, 95)
    
    Dim i As Integer
    
    Select Case chk_type
        Case "Name":
            For i = LBound(name_chars) To UBound(name_chars)
                If name_chars(i) = KeyCode Then
                    CheckKeyCode = False
                    Exit Function
                End If
            Next i
        
        Case "Description":
            For i = LBound(desc_chars) To UBound(desc_chars)
                If desc_chars(i) = KeyCode Then
                    CheckKeyCode = False
                    Exit Function
                End If
            Next i
            
        Case "Unit":
            For i = LBound(name_chars) To UBound(name_chars)
                If name_chars(i) = KeyCode Then
                    CheckKeyCode = False
                    Exit Function
                End If
            Next i
            
            For i = LBound(desc_chars) To UBound(desc_chars)
                If desc_chars(i) = KeyCode Then
                    CheckKeyCode = False
                    Exit Function
                End If
            Next i
        
    End Select
End Function

Public Function IsLetter(ByVal strValue As String) As Boolean
'Function: IsLetter - Evaluates a string, checks if it is entirely alphabetic characters or not.
'Arguments: strValue - String containing the target String
'Returns: Boolean. True - strValue contains only alphabetic characters;
'                  False - strValue contains at least one character that is not alphabetic.


    Dim intPos As Integer
    
    For intPos = 1 To Len(strValue)
        Select Case Asc(Mid(strValue, intPos, 1))
        
            Case 65 To 90, 97 To 122 'charactor is alphebetic
                IsLetter = True

            Case Else 'character is not alphabetic
                IsLetter = False
                Exit For

        End Select
    Next
End Function

Public Function TrimWhiteSpace(ByVal strValue As String) As String
'Function: TrimWhiteSpace - Trims white space on both the front and back ends of a String.
'Arguments: strValue - String to trim
'Returns: String containing the trimmed String


    Dim intPos As Integer
    TrimWhiteSpace = strValue
    intPos = 1
    
    'Trim Front
    Do While intPos <= Len(TrimWhiteSpace)
        Select Case Asc(Mid(TrimWhiteSpace, intPos, 1))
        
            'tab
            Case 9:
                TrimWhiteSpace = Right(TrimWhiteSpace, Len(TrimWhiteSpace) - 1)
                intPos = intPos - 1
            
            'new line
            Case 10:
                TrimWhiteSpace = Right(TrimWhiteSpace, Len(TrimWhiteSpace) - 1)
                intPos = intPos - 1
                
            'vertical tab
            Case 11:
                TrimWhiteSpace = Right(TrimWhiteSpace, Len(TrimWhiteSpace) - 1)
                intPos = intPos - 1
                
            'carriage return
            Case 13:
                TrimWhiteSpace = Right(TrimWhiteSpace, Len(TrimWhiteSpace) - 1)
                intPos = intPos - 1
            
            'space
            Case 32:
                TrimWhiteSpace = Right(TrimWhiteSpace, Len(TrimWhiteSpace) - 1)
                intPos = intPos - 1
                
            Case Else: 'character is not whitespace
                Exit Do

        End Select
        
        intPos = intPos + 1
    Loop
    
    'Trim Back
    intPos = 1
    Do While intPos <= Len(TrimWhiteSpace)
        Select Case Asc(Mid(TrimWhiteSpace, Len(TrimWhiteSpace), 1))
        
            'tab
            Case 9:
                TrimWhiteSpace = Left(TrimWhiteSpace, Len(TrimWhiteSpace) - 1)
                intPos = intPos - 1
            
            'new line
            Case 10:
                TrimWhiteSpace = Left(TrimWhiteSpace, Len(TrimWhiteSpace) - 1)
                intPos = intPos - 1
                
            'vertical tab
            Case 11:
                TrimWhiteSpace = Left(TrimWhiteSpace, Len(TrimWhiteSpace) - 1)
                intPos = intPos - 1
                
            'carriage return
            Case 13:
                TrimWhiteSpace = Left(TrimWhiteSpace, Len(TrimWhiteSpace) - 1)
                intPos = intPos - 1
            
            'space
            Case 32:
                TrimWhiteSpace = Left(TrimWhiteSpace, Len(TrimWhiteSpace) - 1)
                intPos = intPos - 1
                
            Case Else: 'character is not whitespace
                Exit Do

        End Select
        
        intPos = intPos + 1
    Loop
End Function

Public Sub UpdatePageBreaksOnSheet(ByVal Sh As Worksheet)
'Subroutine: UpdatePageBreaksOnSheet - Updates the pages breaks on a sheet by refreshing the sheet's
'                                      window view through xlPageBreakPreview. The resultant view is
'                                      unchanges, but the page breaks are updated.
'Arguments: Sh - Worksheet object for the sheet to update.


    Dim org_view As Integer
    Dim org_sheet As Worksheet
    Set org_sheet = ActiveWorkbook.ActiveSheet
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    Sh.Activate
    org_view = ActiveWindow.View
    
    If org_view = xlPageBreakPreview Then
        ActiveWindow.View = xlNormalView
    Else
        ActiveWindow.View = xlPageBreakPreview
    End If
    ActiveWindow.View = org_view
    
    org_sheet.Activate
    
    If ui_change Then RenderUI True
    
End Sub

Public Function Print2PDF(ByVal sheet_name As String, ByVal sub_directory, ByVal document_title As String) As String
'Function: Print2PDF - Exports specified sheet in ActiveWorkbook as PDF(.pdf) file.
'Arguments: sheet_name - String containing the target sheet name to export
'           sub_directory - name of subdirectory to put this file in. If "", then file goes in current directory.
'           file_name - String containing the PDF document title to save as. ".pdf" extension is optional
'Returns: Boolean. True - strValue contains only alphabetic characters;
'                  False - strValue contains at least one character that is not alphabetic.


    Dim path_name As String
    path_name = ActiveWorkbook.Path
    
    If Len(sub_directory) > 0 Then
        If Left(sub_directory, 1) = "\" Then
            path_name = path_name & sub_directory
        Else
            path_name = path_name & "\" & sub_directory
        End If
        
        If Right(path_name, 1) <> "\" Then
            path_name = path_name & "\"
        End If
    Else
        path_name = path_name & "\"
    End If
    
    'creates directory if it doesn't exist
    If Not FileFolderExists(path_name) Then
        MakeDirs path_name
        If Not FileFolderExists(path_name) Then
            MsgBox "For some reason, this application cannot automatically create this folder directory:" & vbCrLf & _
                path_name & vbCrLf & _
                "Please create this folder manually, try again. Sorry for the inconvenience.", vbCritical
                
            Exit Function
        End If
    End If
    
    On Error Resume Next
    MakeDirs path_name
    Err.Clear
    On Error GoTo 0
    
    path_name = path_name & document_title
    
    If Right(path_name, 4) <> ".pdf" Then
        path_name = path_name & ".pdf"
    End If
   
    Dim print_sheet As Worksheet
    Set print_sheet = ActiveWorkbook.Sheets(sheet_name)
    
    On Error GoTo INVALID_FILENAME
    print_sheet.ExportAsFixedFormat Type:=xlTypePDF, _
                                    filename:=path_name, _
                                    Quality:=xlQualityStandard, _
                                    IncludeDocProperties:=True, _
                                    IgnorePrintAreas:=False, _
                                    OpenAfterPublish:=True
    Err.Clear
    On Error GoTo 0
    
    Set print_sheet = Nothing
    Exit Function

INVALID_FILENAME:
    
    MsgBox "Document not saved: " & path_name & vbCrLf & "Error " & Err & ": " & Error(Err), vbCritical
    Set print_sheet = Nothing
End Function

Public Function FindMissingFile(ByVal vbTitle As String) As String
'Function: FindMissingFile - Generic find file operation. Used for undo/redo operations. Allows for custom title.
'Arguments: vbTitle - String containing the title that will appear in the title bar of the browse window
'Returns: String containing the file path. If dialogue was cancelled, returns vbNullString


    On Error Resume Next
    ChDir ThisWorkbook.Path
    Err.Clear
    On Error GoTo 0
    
    Dim filename As Variant
    filename = Application.GetOpenFilename(FileFilter:="Microsoft Excel Workbooks (*.xlsx; *.xls), *.xlsx; *.xls", _
        FilterIndex:=1, _
        title:=vbTitle, _
        MultiSelect:=False)
    
    If filename <> False Then
        FindMissingFile = CStr(filename)
    Else
        FindMissingFile = vbNullString
    End If
End Function

Public Function CollectionContains(ByVal arrClctn As Collection, ByVal vbTarget) As Boolean
'Function: CollectionContains - Tests a Collection object to see if it contains vbTarget.
'Arguments: arrClctn - Collection to search through
'           vbTarget - variable to search for
'Returns: Boolean. True - vbTarget is an item in arrClctn. False - vbTarget is not an item in arrClctn
    
    
    CollectionContains = False
    
    Dim tempVal As Variant
    For Each tempVal In arrClctn
        If tempVal = vbTarget Then
            CollectionContains = True
            Exit Function
        End If
    Next tempVal
End Function

Public Function IsCellLocation(ByVal loc As String, Optional ByRef return_range As Range) As Boolean
'Function: IsCellLocation - Tests whether the specified String represents a cell location.
'Arguments: loc - the test String. Should contain sheet and cell references to be valid.
'Returns: Boolean. True - valid location string. False - invalid.

    
    On Error Resume Next
    Set return_range = Range(loc)
    Err.Clear
    On Error GoTo 0
    
    IsCellLocation = Not (return_range Is Nothing)
End Function

Public Function SheetExists(ByVal sheet_name As String, Optional ByRef wkbk As Workbook) As Boolean
'Function: SheetExists - Tests whether the specified String represents the name of a sheet in the specified workbook.
'Arguments: sheet_name - the test String.
'           workbook - reference to the Workbook object. If not specified, ThisWorkbook is used.
'Returns: Boolean. True - sheet exists. False - sheet does not exist.

    If wkbk Is Nothing Then
        Set wkbk = ThisWorkbook
    End If
    
    Dim testSheet As Worksheet
    
    For Each testSheet In wkbk.Worksheets
        If testSheet.Name = sheet_name Then
            SheetExists = True
            Exit Function
        End If
    Next testSheet

    SheetExists = False
    Set testSheet = Nothing
End Function

Public Function ColNum(ByVal colName As String) As Integer
'Function: ColNum - Gets the column number given its alphabetical name. Reciprocal to ColLet function
'                   Examples:   ColNum("M")  = 13
'                               ColNum("AB") = 28


    ColNum = 0

    colName = UCase(colName)

    Dim c As Integer
    c = Len(colName)
    
    Dim temp As String
    Dim length As Integer
     
    Do While c >= 1
        temp = Mid(colName, c, 1)
        ColNum = ColNum + 26 ^ (Len(colName) - c) * (Asc(temp) - 64)
        
        c = c - 1
    Loop
End Function

Public Function ColLet(ByVal col_num As Integer) As String
'Function: ColLet - Gets the alphabetical name for a given column number. Reciprocal to ColNum function
'                   Examples:   ColLet(2)  = "B"
'                               ColLet(39) = "AM"

    col_num = Abs(col_num)

    ColLet = ""

    Dim temp As Double
    Dim i As Integer
    i = -1
    Do
        i = i + 1
        temp = (col_num - 1) / (26 ^ i)
    Loop While temp >= 26

    Dim term2 As String
    If col_num Mod 26 = 0 Then
        term2 = "Z"
    Else
        term2 = Chr(col_num Mod 26 + 64)
    End If

    If i > 0 Then
        ColLet = ColLet((col_num - 1) \ 26) & term2
    Else
        ColLet = term2
    End If
End Function

Public Function FirstPhrase(ByVal vbTarget As String) As String
'Function: FirstPhrase - Returns the longest, leftmost section of the string that does not contain punctuation.
'Arguments: vbTarget - String to parse
'Returns: String containing the first phrase in vbTarget


    Dim index As Integer
    index = 1
    
    Dim first_word As String
    Dim next_delim As String
    next_delim = ""
    Dim delim As String
    delim = ""
    
    TrimDelim_Front vbTarget
    FirstPhrase = ""
    
    Do While Len(vbTarget) > 0 And Len(TrimWhiteSpace(delim)) = 0
        first_word = FirstWord(vbTarget)
        vbTarget = Right(vbTarget, Len(vbTarget) - Len(first_word))
        next_delim = TrimDelim_Front(vbTarget)
        
        If Len(FirstPhrase) = 0 And IsNumeric(first_word) Then
            next_delim = ""
        ElseIf Len(FirstPhrase) > 0 And IsNumeric(first_word) And Len(TrimWhiteSpace(next_delim)) > 0 Then
            delim = next_delim
        Else
            FirstPhrase = FirstPhrase & delim & first_word
            delim = next_delim
        End If
    Loop

End Function

Public Function FirstWord(ByVal vbTarget As String) As String
'Function: FirstWord - Returns the first word in a string. Delimited only by white space and punctuation.
'Arguments: vbTarget - String to parse
'Returns: String containing the first word in vbTarget


    Dim index As Integer
    index = 1
    
    FirstWord = ""
    
    Do While IsLetter(Mid(vbTarget, index, 1)) Or IsNumeric(Mid(vbTarget, index, 1)) Or Len(FirstWord) = 0
        If IsLetter(Mid(vbTarget, index, 1)) Or IsNumeric(Mid(vbTarget, index, 1)) Then
            FirstWord = FirstWord & Mid(vbTarget, index, 1)
        End If
        index = index + 1
        If index > Len(vbTarget) Then
            Exit Do
        End If
    Loop

End Function

Public Function LastWord(ByVal vbTarget As String) As String
'Function: LastWord - Returns the last word in a string. Delimited only by white space and punctuation.
'Arguments: vbTarget - String to parse
'Returns: String containing the last word in vbTarget


    Dim index As Integer
    index = Len(vbTarget)
    
    LastWord = ""

    Do While IsLetter(Mid(vbTarget, index, 1)) Or IsNumeric(Mid(vbTarget, index, 1)) Or Len(LastWord) = 0
        If IsLetter(Mid(vbTarget, index, 1)) Or IsNumeric(Mid(vbTarget, index, 1)) Then
            LastWord = Mid(vbTarget, index, 1) & LastWord
        End If
        index = index - 1
        If index <= 0 Then
            Exit Do
        End If
    Loop

End Function

Public Function TrimDelim_Front(ByRef vbTarget As String) As String
'Function: TrimDelim_Front - Trims the whitespace and punctuation off the front(left) side of a string.
'Arguments: vbTarget - String reference to be trimmed
'Returns: String containing the characters trimmed.


    TrimDelim_Front = ""
    
    Do While Len(vbTarget) > 0 And Not (IsLetter(Left(vbTarget, 1)) Or IsNumeric(Left(vbTarget, 1)))
        If Not (IsLetter(Left(vbTarget, 1)) Or IsNumeric(Left(vbTarget, 1))) Then
            TrimDelim_Front = TrimDelim_Front & Left(vbTarget, 1)
            vbTarget = Right(vbTarget, Len(vbTarget) - 1)
        Else
            Exit Function
        End If
    Loop

End Function

Public Function TrimDelim_Back(ByRef vbTarget As String) As String
'Function: TrimDelim_Back - Trims the whitespace and punctuation off the back(right) side of a string.
'Arguments: vbTarget - String reference to be trimmed
'Returns: String containing the characters trimmed.


    TrimDelim_Back = ""
    
    Do While Len(vbTarget) > 0 And Not (IsLetter(Right(vbTarget, 1)) Or IsNumeric(Right(vbTarget, 1)))
        If Not (IsLetter(Right(vbTarget, 1)) Or IsNumeric(Right(vbTarget, 1))) Then
            TrimDelim_Back = Right(vbTarget, 1) & TrimDelim_Back
            vbTarget = Left(vbTarget, Len(vbTarget) - 1)
        Else
            Exit Function
        End If
    Loop

End Function

Public Function ConvertFeetNInches(ByVal strVal As String) As Double
'Function: ConvertFeetNInches - Converts a string representation of feet and inches to a decimal floating point number of feet.
'                               Ex:     1'-0" = 1.0
'                                       2'-6" = 2.5
'                                          3" = 0.25
'                                   3'-1 1/2" = 3.125
'Arguments: strVal - String containing feet and inches.
'Returns: Double containing the numerical value in feet.


    ConvertFeetNInches = 0
    
    Dim numenator As Double
    Dim val As String
    Dim delim As String
    
    TrimDelim_Front strVal
    Do While Len(strVal) > 0
        val = FirstWord(strVal)
        strVal = Right(strVal, Len(strVal) - Len(val))
        delim = TrimDelim_Front(strVal)
        
        If Left(delim, 1) = "'" Then
            ConvertFeetNInches = ConvertFeetNInches + CDbl(val)
        ElseIf Left(delim, 1) = " " Then
            ConvertFeetNInches = ConvertFeetNInches + CDbl(val) / 12
        ElseIf Left(delim, 1) = "/" Then
            numenator = CDbl(val)
            val = FirstWord(strVal)
            strVal = Right(strVal, Len(val))
            delim = TrimDelim_Front(strVal)
            
            ConvertFeetNInches = ConvertFeetNInches + numenator / (12 * CDbl(val))
        Else
            ConvertFeetNInches = ConvertFeetNInches + CDbl(val) / 12
        End If
    Loop
    
End Function

Public Sub AutoFitCells(ByVal col_title As String)
'Subroutine: AutoFitCells - Handles the autofitting in the Current Model Quantities and Checked Quantities columns.
'                           The column titles are easily wider than the individual site columns. To keep the entire merged title cell
'                           fitted correctly, we distibute the column widths appropriately to subsequent site columns.
'Arguments: - col_title - "Current Model Quantities", "Model Extras", or "Checked Quantities"


    If Not (col_title = "Current Model Quantities" Or col_title = "Model Extras" Or col_title = "Checked Quantities") Then
        Exit Sub
    End If
    
    Dim min_width As Double
    If col_title = "Current Model Quantities" Then
        min_width = 23.43 ' 169
    ElseIf col_title = "Model Extras" Then
        min_width = 11.86 ' 88
    ElseIf col_title = "Checked Quantities" Then
        min_width = 17.86 ' 130
    End If
    
    Dim ui_change As Boolean
    ui_change = RenderUI(False)
    
    Dim col As Integer
    col = get_col_num(col_title)
    
    Dim rCell
    Dim i As Integer
    Dim total_width As Double
    Dim title_cell As Range
    Set title_cell = VB_MASTER.Cells(VB_MASTER.TitleRow(), col).MergeArea
    i = 0
    total_width = 0
    
    For Each rCell In title_cell
        VB_MASTER.Columns(col + i).AutoFit
        If VB_MASTER.Cells(VB_MASTER.SubtitleRow(), col + i).ColumnWidth < 8.86 Then
            VB_MASTER.Cells(VB_MASTER.SubtitleRow(), col + i).ColumnWidth = 8.86
        End If
        total_width = total_width + VB_MASTER.Cells(VB_MASTER.SubtitleRow(), col + i).ColumnWidth
        
        i = i + 1
    Next rCell

    If total_width < min_width Then
        VB_MASTER.Cells(VB_MASTER.SubtitleRow(), col + i - 1).ColumnWidth = VB_MASTER.Cells(VB_MASTER.SubtitleRow(), col + i - 1).ColumnWidth + min_width - total_width
    End If
    
    If ui_change Then RenderUI True
    
    Set title_cell = Nothing
End Sub

Public Sub QuickSort(arr, Lo As Long, Hi As Long)
'Subroutine: QuickSort - standard sort algorithm. Sorts ascending. The highest rows should be
'                        deleted first so subsequent row numbers aren't changed.
'Arguments: arr - Reference to array to sort
'           Lo - lower bound of arr
'           Hi - upper bound of arr


    Dim varPivot As Variant
    Dim varTmp As Variant
    Dim tmpLow As Long
    Dim tmpHi As Long
    tmpLow = Lo
    tmpHi = Hi
    varPivot = arr((Lo + Hi) \ 2)
    Do While tmpLow <= tmpHi
        Do While arr(tmpLow) < varPivot And tmpLow < Hi
            tmpLow = tmpLow + 1
        Loop
        Do While varPivot < arr(tmpHi) And tmpHi > Lo
            tmpHi = tmpHi - 1
        Loop
        If tmpLow <= tmpHi Then
            varTmp = arr(tmpLow)
            arr(tmpLow) = arr(tmpHi)
            arr(tmpHi) = varTmp
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Loop
    If Lo < tmpHi Then QuickSort arr, Lo, tmpHi
    If tmpLow < Hi Then QuickSort arr, tmpLow, Hi
End Sub

Public Function MinRow(ByVal Target As Range) As Integer
'Function: MinRow - Determines the minimum/lowest/first row in a Range
'Arguments: Range object
'Returns: The lowest row number

    If Target Is Nothing Then
        MinRow = -1
        Exit Function
    End If
    
    Dim vbCell As Range
    Dim index As Integer
    Dim i_row() As Integer
    ReDim i_row(1 To Target.Cells.count)
    index = LBound(i_row)
    
    'adds all selected rows marked for deletion to integer array
    For Each vbCell In Target
        i_row(index) = vbCell.row
        index = index + 1
    Next vbCell
    Set vbCell = Nothing
    
    ' sort
    QuickSort i_row, LBound(i_row), UBound(i_row)
    
    MinRow = i_row(LBound(i_row))
End Function

Public Function MaxRow(ByVal Target As Range) As Integer
'Function: MaxRow - Determines the maximum/highest/last row in a Range
'Arguments: Range object
'Returns: The highest row number

    If Target Is Nothing Then
        MaxRow = -1
        Exit Function
    End If
    
    Dim vbCell As Range
    Dim index As Integer
    Dim i_row() As Integer
    ReDim i_row(1 To Target.Cells.count)
    index = LBound(i_row)
    
    'adds all selected rows marked for deletion to integer array
    For Each vbCell In Target
        i_row(index) = vbCell.row
        index = index + 1
    Next vbCell
    Set vbCell = Nothing
    
    ' sort
    QuickSort i_row, LBound(i_row), UBound(i_row)
    
    MaxRow = i_row(UBound(i_row))
End Function

Public Function MinColumn(ByVal Target As Range) As Integer
'Function: MinColumn - Determines the minimum/lowest/first Column in a Range
'Arguments: Range object
'Returns: The lowest Column

    If Target Is Nothing Then
        MinColumn = -1
        Exit Function
    End If
    
    Dim vbCell As Range
    Dim index As Integer
    Dim i_col() As Integer
    ReDim i_col(1 To Target.Cells.count)
    index = LBound(i_col)
    
    ' adds all selected rows marked for deletion to integer array
    For Each vbCell In Target
        i_col(index) = vbCell.column
        index = index + 1
    Next vbCell
    Set vbCell = Nothing
    
    ' sort
    QuickSort i_col, LBound(i_col), UBound(i_col)
    
    MinColumn = i_col(LBound(i_col))
End Function

Public Function MaxColumn(ByVal Target As Range) As Integer
'Function: MaxColumn - Determines the maximum/highest/last Column in a Range
'Arguments: Range object
'Returns: The highest Column

    If Target Is Nothing Then
        MaxColumn = -1
        Exit Function
    End If
    
    Dim vbCell As Range
    Dim index As Integer
    Dim i_col() As Integer
    ReDim i_col(1 To Target.Cells.count)
    index = LBound(i_col)
    
    ' adds all selected rows marked for deletion to integer array
    For Each vbCell In Target
        i_col(index) = vbCell.column
        index = index + 1
    Next vbCell
    Set vbCell = Nothing
    
    ' sort
    QuickSort i_col, LBound(i_col), UBound(i_col)
    
    MaxColumn = i_col(UBound(i_col))
End Function

Public Function VBRangeMax(ByVal strTarget As String) As Integer
'Function: VBRangeMax - Finds the maximum numeric value in a given range.
'Arguments: strTarget - String containing the address of a Range in VB_MASTER
'Returns: Integer containing the maximum numeric value in the given range.

    Dim Target As Range
    If Not IsCellLocation(strTarget) Then
        VBRangeMax = Null
        Exit Function
    End If
    Set Target = VB_MASTER.Range(strTarget)
    
    Dim vbCell As Range
    Dim i As Integer
    i = -1
    Dim vbVal
    
    For Each vbCell In Target
        vbVal = vbCell.Value2
        If Not IsEmpty(vbCell) And IsNumeric(vbVal) Then
            If i = -1 Then
                VBRangeMax = vbVal
                i = 0
            Else
                VBRangeMax = VBMax(VBRangeMax, vbVal)
            End If
        End If
    Next vbCell
End Function

Public Function VBMax(ByVal A As Long, ByVal B As Long) As Long
'Function: VBMax - Compares two numbers and returns the largest. If they are equal, returns the first.
'Arguments: A - first number to compare
'           B - second number to compare
'Returns: Max value between A and B

    If A >= B Then VBMax = A Else VBMax = B
End Function

Public Function VBMin(ByVal A As Long, ByVal B As Long) As Long
'Function: VBMin - Compares two numbers and returns the smallest. If they are equal, returns the first.
'Arguments: A - first number to compare
'           B - second number to compare
'Returns: Min value between A and B

    If A <= B Then VBMin = A Else VBMin = B
End Function

Public Function VBRoundUp(ByVal vbArg As Double) As Long
'Function: VBRoundUp - Truncates decimal places and if value has changed, increments by one.
'Arguments: vbArg - Number to round
'Returns: Rounded value

    VBRoundUp = VBRoundDown(vbArg)
    If vbArg > VBRoundUp Then VBRoundUp = VBRoundUp + 1
End Function

Public Function VBRoundDown(ByVal vbArg As Double) As Long
'Function: VBRoundDown - Truncates decimal places and returns value.
'Arguments: vbArg - Number to round
'Returns: Rounded value

    VBRoundDown = Fix(vbArg)
End Function

Public Function IsUniqueInstance() As Boolean
'Function: IsUniqueInstance - Evaluates all open workbooks to see if one of them is already an HMM BOM
'Returns: Boolean. True - No other open workbook is an HMM BOM. False - an HMM BOM is already open

    Dim wkbk As Variant
    Dim test_sheet As Worksheet
    
    IsUniqueInstance = True
    
    For Each wkbk In Application.Workbooks
        
        If Not (wkbk Is ThisWorkbook) Then
            For Each test_sheet In wkbk.Worksheets
                If test_sheet.Name = VB_VAR_STORE.Name Then
                    IsUniqueInstance = IsUniqueInstance And False
                End If
            Next test_sheet
            
            If Not IsUniqueInstance Then Exit For
        End If

    Next wkbk
    
    Set test_sheet = Nothing
End Function
