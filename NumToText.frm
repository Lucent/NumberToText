VERSION 5.00
Begin VB.Form NumToText 
   Caption         =   "Number to Text"
   ClientHeight    =   4020
   ClientLeft      =   132
   ClientTop       =   588
   ClientWidth     =   4464
   Icon            =   "NumToText.frx":0000
   LinkTopic       =   "NumToText"
   LockControls    =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   4464
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ExitButton 
      Caption         =   "Exit"
      Height          =   372
      Left            =   3025
      TabIndex        =   6
      Top             =   3575
      Width           =   1342
   End
   Begin VB.CommandButton Clear 
      Caption         =   "Clear"
      Height          =   372
      Left            =   1573
      TabIndex        =   5
      Top             =   3575
      Width           =   1342
   End
   Begin VB.CommandButton Convert 
      Caption         =   "Convert"
      Default         =   -1  'True
      Height          =   372
      Left            =   121
      TabIndex        =   4
      Top             =   3575
      Width           =   1342
   End
   Begin VB.Frame OutputFrame 
      Caption         =   "Text"
      Height          =   2794
      Left            =   121
      TabIndex        =   1
      Top             =   660
      Width           =   4246
      Begin VB.TextBox OutputBox 
         Height          =   2464
         Left            =   121
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   220
         Width           =   3971
      End
   End
   Begin VB.Frame InputFrame 
      Caption         =   "Number"
      Height          =   616
      Left            =   121
      TabIndex        =   0
      Top             =   0
      Width           =   4246
      Begin VB.TextBox InputBox 
         Height          =   264
         Left            =   121
         TabIndex        =   2
         Top             =   220
         Width           =   4004
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuConvert 
         Caption         =   "&Convert"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "C&lear"
      End
      Begin VB.Menu mnuSpacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "&Edit"
      Index           =   1
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSpacer2 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "&Options"
      Index           =   2
      Begin VB.Menu mnuAutoCopy 
         Caption         =   "&Auto-copy output"
      End
      Begin VB.Menu mnuCurrency 
         Caption         =   "&Currency-style output"
      End
      Begin VB.Menu mnuTrim 
         Caption         =   "&Trim trailing zeros"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "&Help"
      Index           =   3
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "NumToText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Remember, Visual Basic can't handle sixty-digit integers, so numbers must be treated as strings
Option Explicit
Dim Ones, Tens, Large, Prefix As Variant
Dim Zero, WholeCurrency, FractionalCurrency, Conjunction, DecimalPoint As String
Dim Block, x As Integer

'This function grabs numbers in blocks of three digits and returns the words
Function BlockToText(Block As Integer) As String
    Dim BlockOutput As String
    If Block >= 100 Then
        BlockOutput = Ones(Block \ 100) & " hundred "
        Block = Block Mod 100
    End If
    If Block <= 20 Then
        BlockToText = RTrim(BlockOutput & Ones(Block))
    Else
        BlockToText = RTrim(BlockOutput & Tens(Block \ 10) & Chr(vbKeyInsert) & Ones(Block Mod 10))
    End If
End Function

'Visual Basic's built-in IsNumeric() function doesn't catch everything
Function IsNumber(TestNumber As String) As Boolean
    IsNumber = Not (TestNumber Like "*[!0-9" & DecimalPoint & "+-]*") And Not (TestNumber Like "?*[+-]*") And Not (TestNumber Like "[" & DecimalPoint & "+-]") And Not (TestNumber Like "*" & DecimalPoint & "*" & DecimalPoint & "*")
End Function

Private Sub Clear_Click()
    InputBox.Text = vbNullString
    Convert_Click
End Sub

Private Sub Convert_Click()
    Dim InputText, WordBlock, Whole, WholeOutput, Fraction, FractionOutput, Negative, FractionEnd, Singular, TestNumber, WholeCurrencyPlural, FractionalCurrencyPlural As String
    InputFrame.Caption = "Number"
    OutputFrame.Caption = "Text"
    
    'Don't bother converting if InputBox doesn't contain a number, don't even display an error
    If InputBox.Text = vbNullString Then
        OutputBox.Text = vbNullString
        InputBox.SetFocus
        Exit Sub
    End If
    
    'Use the above function to return errors for all non-numbers
    If Not IsNumber(InputBox.Text) Then
        InputFrame.Caption = "Number (Error: Not a number)"
        InputBox.SetFocus
        Exit Sub
    End If
    InputText = Trim(Replace(Replace(InputBox.Text, Chr(vbKeyInsert), vbNullString), "+", vbNullString))
    
    'Nice way of getting the integer part of a number (string, since can be very long) in one line whether it has a decimal point or not
    Whole = Left(InputText, InStr(InputText & DecimalPoint, DecimalPoint) - 1)
    
    'Warn the user if the fractional part of the number is trimmed
    If Len(Mid(InputText, InStr(InputText & DecimalPoint, DecimalPoint) + 1)) > UBound(Large) * 3 - 1 Then OutputFrame.Caption = "Text (Warning: Fractional portion trimmed to " & UBound(Large) * 3 - 1 & " digits)"
    Fraction = Left(Mid(InputText, InStr(InputText & DecimalPoint, DecimalPoint) + 1), UBound(Large) * 3 - 1)

    'Sgn is one function that seems to work for determining whether a number is zero or non-zero
    If Sgn(InputBox.Text) = -1 Then Negative = "negative "
    
    'Trim off the leading zeros and don't count them when reporting too many digits
    While Left(Whole, 1) = "0"
        Whole = Right(Whole, Len(Whole) - 1)
    Wend
    
    'Generate an error if the integer part contains more than UBound(Large) * 3 digits
    If Len(Whole) > UBound(Large) * 3 Then
        InputFrame.Caption = "Number (Error: Integer contains " & Len(Whole) - UBound(Large) * 3 & " too many digits)"
        'OutputBox.Text = "Error: Number is too large"
        InputBox.SetFocus
        Exit Sub
    End If
    
    'Trim trailing zeros
    If mnuTrim.Checked Then
        While Right(Fraction, 1) = "0"
            Fraction = Left(Fraction, Len(Fraction) - 1)
        Wend
    End If
    
    'For each set of three digits, starting at the left, run them through BlockToText function and add a word (billion, million)
    For x = 1 To (Len(Whole) + 2) \ 3
        WordBlock = BlockToText(Mid("00" & Whole, Len(Whole) - 3 * x + 3, 3))
        If WordBlock <> Empty Then WholeOutput = WordBlock & " " & Large(x) & ", " & WholeOutput
    Next
    'To save a couple lines and some time, BlockToText sometimes produces a little extra garbage, trim it off
    If WholeOutput <> 0 Then WholeOutput = RTrim(Replace(Left(WholeOutput, Len(WholeOutput) - 2), Chr(vbKeyInsert) & " ", " "))
    
    'If in currency mode, round to two places, get the plurals right, and output
    If mnuCurrency.Checked Then
        If Round("0" & DecimalPoint & Fraction, 2) <> 0 Then FractionOutput = RTrim(Replace(BlockToText(Mid(Format(Round(DecimalPoint & Fraction, 2), DecimalPoint & "00"), 2)) & " ", Chr(vbKeyInsert) & " ", vbNullString))
        If Whole <> 1 Then WholeCurrencyPlural = "s"
        If FractionOutput <> "one" Then FractionalCurrencyPlural = "s"
        If Sgn(Val(Whole)) = 0 And Sgn(Round(Val(DecimalPoint & Fraction), 2)) = 0 Then OutputBox.Text = Zero & " " & WholeCurrency & WholeCurrencyPlural & " " & Conjunction & " no " & FractionalCurrency & FractionalCurrencyPlural
        If Sgn(Val(Whole)) <> 0 And Sgn(Round(Val(DecimalPoint & Fraction), 2)) <> 0 Then OutputBox.Text = Negative & WholeOutput & " " & WholeCurrency & WholeCurrencyPlural & " " & Conjunction & " " & FractionOutput & " " & FractionalCurrency & FractionalCurrencyPlural
        If Sgn(Val(Whole)) = 0 And Sgn(Round(Val(DecimalPoint & Fraction), 2)) <> 0 Then OutputBox.Text = Negative & Zero & " " & WholeCurrency & WholeCurrencyPlural & " " & Conjunction & " " & FractionOutput & " " & FractionalCurrency & FractionalCurrencyPlural
        If Sgn(Val(Whole)) <> 0 And Sgn(Round(Val(DecimalPoint & Fraction), 2)) = 0 Then OutputBox.Text = Negative & WholeOutput & " " & WholeCurrency & WholeCurrencyPlural & " " & Conjunction & " no " & FractionalCurrency & FractionalCurrencyPlural
        If mnuAutoCopy.Checked Then
            Clipboard.Clear
            Clipboard.SetText OutputBox.Text
        End If
        InputBox.SetFocus
        Exit Sub
    End If
    
    'Continue parsing with the fractional part, doing the same thing as above
    For x = 1 To (Len(Fraction) + 2) \ 3
        WordBlock = BlockToText(Mid("00" & Fraction, Len(Fraction) - 3 * x + 3, 3))
        If WordBlock <> Empty Then FractionOutput = WordBlock & " " & Large(x) & ", " & FractionOutput
    Next
    If FractionOutput <> 0 Then FractionOutput = RTrim(Replace(Left(FractionOutput, Len(FractionOutput) - 2), Chr(vbKeyInsert) & " ", " "))
    
    'Add the final word to the end of the fraction signifying the whole value and get the plural right
    If Val(Fraction) = 1 Then FractionEnd = Large(Len(Fraction) \ 3 + 1) & "th" Else FractionEnd = Large(Len(Fraction) \ 3 + 1) & "ths"
    FractionEnd = Prefix(Len(Fraction) Mod 3) & FractionEnd
    'No need for a tenths and hundredths Case, just remove the dash between "ten-" and "ths" or "th"
    If x <= 2 Then FractionEnd = Replace(FractionEnd, Chr(vbKeyInsert), vbNullString)
    
    'This can't be done in one line
    If Sgn(Val(Whole)) = 0 And Sgn(Val(Fraction)) = 0 Then OutputBox.Text = Zero
    If Sgn(Val(Whole)) <> 0 And Sgn(Val(Fraction)) <> 0 Then OutputBox.Text = Negative & WholeOutput & " " & Conjunction & " " & FractionOutput & " " & FractionEnd
    If Sgn(Val(Whole)) = 0 And Sgn(Val(Fraction)) <> 0 Then OutputBox.Text = Negative & FractionOutput & " " & FractionEnd
    If Sgn(Val(Whole)) <> 0 And Sgn(Val(Fraction)) = 0 Then OutputBox.Text = Negative & WholeOutput
    If mnuAutoCopy.Checked Then
        Clipboard.Clear
        Clipboard.SetText OutputBox.Text
    End If
    InputBox.SetFocus
End Sub

Private Sub ExitButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    gHW = Me.hwnd
    Hook
    Zero = "zero"
    Ones = Array(Empty, "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen", "twenty")
    Tens = Array(Empty, Empty, "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety")
    Prefix = Array(Empty, "ten-", "hundred-")
    Large = Array(Empty, Empty, "thousand", "million", "billion", "trillion", "quintillion", "sextillion", "septillion", "octillion", "nonillion", "decillion", "undecillion", "duodecillion", "tredecillion", "quattuordecillion", "quindecillion", "sexdecillion", "septendecillion", "octodecillion", "novemdecillion", "vigintillion", "unvigintillion", "dovigintillion", "trevigintillion", "quattuorvigintillion", "quinvigintillion", "sexvigintillion", "septenvigintillion", "octovigintillion", "novemvigintillion", "trigintillion", "untrigintillion", "dotrigintillion", "tretrigintillion", "quattuortrigintillion", "quintrigintillion", "sextrigintillion", "septentrigintillion", "octotrigintillion", "novemtrigintillion")
    WholeCurrency = "dollar"
    FractionalCurrency = "cent"
    Conjunction = "and"
    DecimalPoint = Chr(Asc(DecimalSeparator()))
    'Set the window position and size to where it was on last exit
    Me.Move GetSetting("Number to Text", "Position", "Left"), GetSetting("Number to Text", "Position", "Top"), GetSetting("Number to Text", "Position", "Width"), GetSetting("Number to Text", "Position", "Height")
    'Set the options to how they were on last exit
    mnuAutoCopy.Checked = GetSetting("Number to Text", "Options", "Auto-copy output")
    mnuCurrency.Checked = GetSetting("Number to Text", "Options", "Currency-style output")
    mnuTrim.Checked = GetSetting("Number to Text", "Options", "Trim trailing zeros")
Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 13
            Resume Next
    End Select
End Sub

Private Sub Form_Resize()
    'If the window's not an icon, resize all controls in it as it resizes
    If Not WindowState = vbMinimized Then
        InputFrame.Move 121, 0, Me.ScaleWidth - 242
        InputBox.Move 121, 220, Me.ScaleWidth - 484
        OutputFrame.Move 121, 660, Me.ScaleWidth - 242, Me.ScaleHeight - 1243
        OutputBox.Move 121, 220, Me.ScaleWidth - 484, Me.ScaleHeight - 1573
        Convert.Move 121, Me.ScaleHeight - 462, Me.ScaleWidth / 3 - 154
        Clear.Move Me.ScaleWidth / 3 + 77, Me.ScaleHeight - 462, Me.ScaleWidth / 3 - 154
        ExitButton.Move 2 * Me.ScaleWidth / 3 + 33, Me.ScaleHeight - 462, Me.ScaleWidth / 3 - 154
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Save options, position, and size to the registry
    SaveSetting "Number to Text", "Options", "Auto-copy output", mnuAutoCopy.Checked
    SaveSetting "Number to Text", "Options", "Currency-style output", mnuCurrency.Checked
    SaveSetting "Number to Text", "Options", "Trim trailing zeros", mnuTrim.Checked
    SaveSetting "Number to Text", "Position", "Left", Me.Left
    SaveSetting "Number to Text", "Position", "Top", Me.Top
    SaveSetting "Number to Text", "Position", "Width", Me.Width
    SaveSetting "Number to Text", "Position", "Height", Me.Height
    Unhook
End Sub

Private Sub InputBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(DecimalPoint) And Len(InputBox.Text) - Len(Replace(InputBox.Text, DecimalPoint, vbNullString)) > 0 Then
        Beep
        KeyAscii = 0
    End If
    If KeyAscii = vbKeyInsert And (InputBox.SelStart > 0 Or Left(InputBox.Text, 1) = Chr(vbKeyInsert)) Then
        Beep
        KeyAscii = 0
    End If
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyTab And KeyAscii <> Asc(DecimalPoint) And KeyAscii <> vbKeyInsert Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Menu_Click(Index As Integer)
    'Mimics functionality of the right-click context menu
    If TypeOf Screen.ActiveControl Is TextBox Then
        If Screen.ActiveControl.Text = vbNullString Or Screen.ActiveControl.Text = Screen.ActiveControl.SelText Then
            mnuSelectAll.Enabled = False
        Else
            mnuSelectAll.Enabled = True
        End If
        
        If Clipboard.GetText = vbNullString Then
            mnuPaste.Enabled = False
        Else
            mnuPaste.Enabled = True
        End If
        
        If Screen.ActiveControl.SelText <> vbNullString Then
            mnuCopy.Enabled = True
            If Screen.ActiveControl.Locked = True Then
                mnuCut.Enabled = False
                mnuDelete.Enabled = False
            Else
                mnuCut.Enabled = True
                mnuDelete.Enabled = True
            End If
        Else
            mnuCut.Enabled = False
            mnuCopy.Enabled = False
            mnuDelete.Enabled = False
        End If
    Else
        mnuCut.Enabled = False
        mnuCopy.Enabled = False
        mnuDelete.Enabled = False
        mnuPaste.Enabled = False
        mnuDelete.Enabled = False
        mnuSelectAll.Enabled = False
    End If
End Sub

Private Sub mnuAbout_Click()
    Load frmAbout
    frmAbout.Show 1
End Sub

Private Sub mnuAutoCopy_Click()
    mnuAutoCopy.Checked = Not mnuAutoCopy.Checked
End Sub

Private Sub mnuCurrency_Click()
    mnuCurrency.Checked = Not mnuCurrency.Checked
    If OutputBox.Text <> vbNullString And InputBox.Text <> vbNullString Then Convert_Click
End Sub

Private Sub mnuClear_Click()
    Clear_Click
End Sub

Private Sub mnuConvert_Click()
    Convert_Click
End Sub

Private Sub mnuCopy_Click()
    Clipboard.Clear
    Clipboard.SetText Screen.ActiveControl.SelText
End Sub

Private Sub mnuCut_Click()
    Clipboard.Clear
    Clipboard.SetText Screen.ActiveControl.SelText
    Screen.ActiveControl.SelText = vbNullString
End Sub

Private Sub mnuDelete_Click()
    Screen.ActiveControl.SelText = vbNullString
End Sub

Private Sub mnuExit_Click()
    ExitButton_Click
End Sub

Private Sub mnuPaste_Click()
    SendKeys Clipboard.GetText
    'Screen.ActiveControl.SelText = Clipboard.GetText
End Sub

Private Sub mnuSelectAll_Click()
    Screen.ActiveControl.SelStart = 0
    Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text)
End Sub

Private Sub mnuTrim_Click()
    mnuTrim.Checked = Not mnuTrim.Checked
    If OutputBox.Text <> vbNullString And InputBox.Text <> vbNullString Then Convert_Click
End Sub
