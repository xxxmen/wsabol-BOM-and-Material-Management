VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FilenameTagForm 
   Caption         =   "Format File Tag"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6360
   OleObjectBlob   =   "FilenameTagForm.frx":0000
End
Attribute VB_Name = "FilenameTagForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' User Form: FilenameTagForm                                '
'                                                           '
' All files generated by this application include this      '
' filename tag. It can take any form, as determined by the  '
' PM's preferences. This form allows the user to modify the '
' filename tag to however he/she chooses.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Public Sub LoadForm()
'Subroutine: LoadForm - Controls the .Show routine for this form. Initializes the form with the existing tag.


    'get file tag format
    FormatBox.text = VB_VAR_STORE.HMMFileTagFormat()
    FormatBox.SetFocus
    
    Me.Top = Application.Top + Application.height / 2 - Me.height / 2
    Me.Left = Application.Left + Application.Width / 2 - Me.Width / 2
    Me.Show
End Sub

Private Sub FormatBox_Change()
'Subroutine: FormatBox_Change - User Form event handler. Called whenever FormatBoxChanges. Updates the ExampleLabel

    ExampleLabel.Caption = "DocumentTitle" & HMMFileTag(FormatBox.text) & ".pdf"
End Sub

Private Sub SaveButton_Click()
'Subroutine: SaveButton_Click - User Form event handler. Called whenever the user clicks the SaveButton.
'                               Stores the new format in VB_VAR_STORE.
   
    'set file tag format
    VB_VAR_STORE.HMMFileTagFormat FormatBox.text
    Unload Me
End Sub

Private Sub YearButton_Click()
'Subroutine: YearButton_Click - User Form event handler. Called whenever the user clicks the YearButton.
'                               Adds 'yyyy' to the format at the current cursor location

    With FormatBox
        .text = Left(.text, .SelStart) & "yyyy" & Mid(.text, .SelStart + 1)
    End With
    FormatBox.SetFocus
End Sub

Private Sub MonthButton_Click()
'Subroutine: MonthButton_Click - User Form event handler. Called whenever the user clicks the MonthButton.
'                                Adds 'mm' to the format at the current cursor location

    With FormatBox
        .text = Left(.text, .SelStart) & "mm" & Mid(.text, .SelStart + 1)
    End With
    FormatBox.SetFocus
End Sub

Private Sub DayButton_Click()
'Subroutine: DayButton_Click - User Form event handler. Called whenever the user clicks the DayButton.
'                              Adds 'dd' to the format at the current cursor location

    With FormatBox
        .text = Left(.text, .SelStart) & "dd" & Mid(.text, .SelStart + 1)
    End With
    FormatBox.SetFocus
End Sub

Private Sub MonthNameButton_Click()
'Subroutine: MonthNameButton_Click - User Form event handler. Called whenever the user clicks the MonthNameButton.
'                                    Adds 'mmmm' to the format at the current cursor location

    With FormatBox
        .text = Left(.text, .SelStart) & "mmmm" & Mid(.text, .SelStart + 1)
    End With
    FormatBox.SetFocus
End Sub

Private Sub Day3Button_Click()
'Subroutine: Day3Button_Click - User Form event handler. Called whenever the user clicks the Day3Button.
'                               Adds 'ddd' to the format at the current cursor location

    With FormatBox
        .text = Left(.text, .SelStart) & "ddd" & Mid(.text, .SelStart + 1)
    End With
    FormatBox.SetFocus
End Sub

Private Sub Month3Button_Click()
'Subroutine: Month3Button_Click - User Form event handler. Called whenever the user clicks the Month3Button.
'                                 Adds 'mmm' to the format at the current cursor location

    With FormatBox
        .text = Left(.text, .SelStart) & "mmm" & Mid(.text, .SelStart + 1)
    End With
    FormatBox.SetFocus
End Sub

Private Sub HourButton_Click()
'Subroutine: HourButton_Click - User Form event handler. Called whenever the user clicks the HourButton.
'                                 Adds 'HH' to the format at the current cursor location

    With FormatBox
        .text = Left(.text, .SelStart) & "HH" & Mid(.text, .SelStart + 1)
    End With
    FormatBox.SetFocus
End Sub

Private Sub MinuteButton_Click()
'Subroutine: MinuteButton_Click - User Form event handler. Called whenever the user clicks the MinuteButton.
'                                 Adds 'NN' to the format at the current cursor location

    With FormatBox
        .text = Left(.text, .SelStart) & "NN" & Mid(.text, .SelStart + 1)
    End With
    FormatBox.SetFocus
End Sub

Private Sub SecondButton_Click()
'Subroutine: SecondButton_Click - User Form event handler. Called whenever the user clicks the SecondButton.
'                                 Adds 'SS' to the format at the current cursor location

    With FormatBox
        .text = Left(.text, .SelStart) & "SS" & Mid(.text, .SelStart + 1)
    End With
    FormatBox.SetFocus
End Sub

Private Sub AMPMButton_Click()
'Subroutine: AMPMButton_Click - User Form event handler. Called whenever the user clicks the AMPMButton.
'                               Adds 'AMPM' to the format at the current cursor location

    With FormatBox
        .text = Left(.text, .SelStart) & "AMPM" & Mid(.text, .SelStart + 1)
    End With
    FormatBox.SetFocus
End Sub

Private Sub ProjectNameButton_Click()
'Subroutine: ProjectNameButton_Click - User Form event handler. Called whenever the user clicks the ProjectNameButton.
'                                      Adds '"~ProjectName"' to the format at the current cursor location

    With FormatBox
        .text = Left(.text, .SelStart) & Chr(34) & "~ProjectName" & Chr(34) & Mid(.text, .SelStart + 1)
    End With
End Sub

Private Sub ClientButton_Click()
'Subroutine: ClientButton_Click - User Form event handler. Called whenever the user clicks the ClientButton.
'                                 Adds '"~ProjectClientName"' to the format at the current cursor location

    With FormatBox
        .text = Left(.text, .SelStart) & Chr(34) & "~ProjectClientName" & Chr(34) & Mid(.text, .SelStart + 1)
    End With
    FormatBox.SetFocus
End Sub

Private Sub ProjectNumberButton_Click()
'Subroutine: ProjectNumberButton_Click - User Form event handler. Called whenever the user clicks the ProjectNumberButton.
'                                 Adds '"~ProjectNumber"' to the format at the current cursor location

    With FormatBox
        .text = Left(.text, .SelStart) & Chr(34) & "~ProjectNumber" & Chr(34) & Mid(.text, .SelStart + 1)
    End With
    FormatBox.SetFocus
End Sub

Private Sub CancelButton_Click()
'Subroutine: CancelButton_Click - User Form event handler. Called whenever the user clicks the CancelButton.

    Unload Me
End Sub

Private Sub UserForm_Click()
'Subroutine: UserForm_Click - User Form event handler. Called whenever the user clicks inside the UserForm.
'                             Retains focus on FormatBox.

    FormatBox.SetFocus
End Sub

Private Sub ExampleFrame_Click()
'Subroutine: ExampleFrame_Click - User Form event handler. Called whenever the user clicks inside the ExampleFrame.
'                                 Retains focus on FormatBox.

    FormatBox.SetFocus
End Sub

Private Sub ExampleLabel_Click()
'Subroutine: ExampleLabel_Click - User Form event handler. Called whenever the user clicks inside the ExampleLabel.
'                                 Retains focus on FormatBox.

    FormatBox.SetFocus
End Sub

Private Sub Instructions_Click()
'Subroutine: Instructions_Click - User Form event handler. Called whenever the user clicks inside the Instructions.
'                                 Retains focus on FormatBox.

    FormatBox.SetFocus
End Sub

Private Sub ProjectFrame_Click()
'Subroutine: ProjectFrame_Click - User Form event handler. Called whenever the user clicks inside the ProjectFrame.
'                                 Retains focus on FormatBox.

    FormatBox.SetFocus
End Sub

Private Sub TimeFrame_Click()
'Subroutine: TimeFrame_Click - User Form event handler. Called whenever the user clicks inside the TimeFrame.
'                              Retains focus on FormatBox.

    FormatBox.SetFocus
End Sub

Private Sub DateFrame_Click()
'Subroutine: DateFrame_Click - User Form event handler. Called whenever the user clicks inside the DateFrame.
'                              Retains focus on FormatBox.

    FormatBox.SetFocus
End Sub