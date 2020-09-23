VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB.Messager"
   ClientHeight    =   5400
   ClientLeft      =   1095
   ClientTop       =   1035
   ClientWidth     =   8385
   Icon            =   "Messager.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5400
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   315
      Left            =   6030
      TabIndex        =   43
      Top             =   5040
      Width           =   1110
   End
   Begin VB.CommandButton cmdShow 
      Cancel          =   -1  'True
      Caption         =   "&Show me"
      Height          =   315
      Left            =   30
      TabIndex        =   22
      Top             =   5040
      Width           =   1110
   End
   Begin VB.Frame fraModal 
      Height          =   690
      Left            =   5850
      TabIndex        =   41
      Top             =   4275
      Width           =   2475
      Begin VB.CheckBox chkModal 
         Caption         =   "System Modal"
         Height          =   195
         Left            =   210
         TabIndex        =   18
         Top             =   300
         Width           =   2070
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Default button"
      Height          =   1155
      Left            =   6345
      TabIndex        =   38
      Top             =   2160
      Width           =   1980
      Begin VB.OptionButton optDefault 
         Caption         =   "Button 3"
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   825
         Width           =   1725
      End
      Begin VB.OptionButton optDefault 
         Caption         =   "Button 2"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   540
         Width           =   1725
      End
      Begin VB.OptionButton optDefault 
         Caption         =   "Button 1"
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Value           =   -1  'True
         Width           =   1725
      End
   End
   Begin VB.Frame fraHelp 
      Height          =   945
      Left            =   5850
      TabIndex        =   37
      Top             =   3330
      Width           =   2475
      Begin VB.TextBox txtContext 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   17
         Top             =   525
         Width           =   1380
      End
      Begin VB.CheckBox chkHelpFile 
         Alignment       =   1  'Right Justify
         Caption         =   "Help file:"
         Height          =   210
         Left            =   105
         TabIndex        =   16
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label lblHelpFile 
         Caption         =   "(No help file)"
         Enabled         =   0   'False
         Height          =   210
         Left            =   1290
         TabIndex        =   40
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label lblContext 
         Caption         =   "Context:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   135
         TabIndex        =   39
         Top             =   585
         Width           =   795
      End
   End
   Begin VB.Frame fraButtons 
      Caption         =   "Buttons"
      Height          =   2145
      Left            =   6345
      TabIndex        =   36
      Top             =   0
      Width           =   1980
      Begin VB.OptionButton optButton 
         Caption         =   "Retry and Cancel"
         Height          =   210
         Index           =   5
         Left            =   105
         TabIndex        =   12
         Top             =   1770
         Width           =   1725
      End
      Begin VB.OptionButton optButton 
         Caption         =   "Yes and No"
         Height          =   210
         Index           =   4
         Left            =   105
         TabIndex        =   11
         Top             =   1470
         Width           =   1725
      End
      Begin VB.OptionButton optButton 
         Caption         =   "Yes, No, Cancel"
         Height          =   210
         Index           =   3
         Left            =   105
         TabIndex        =   10
         Top             =   1170
         Width           =   1725
      End
      Begin VB.OptionButton optButton 
         Caption         =   "Abort, Retry, Ignore"
         Height          =   210
         Index           =   2
         Left            =   105
         TabIndex        =   9
         Top             =   870
         Width           =   1725
      End
      Begin VB.OptionButton optButton 
         Caption         =   "Ok and Cancel"
         Height          =   210
         Index           =   1
         Left            =   105
         TabIndex        =   8
         Top             =   570
         Width           =   1725
      End
      Begin VB.OptionButton optButton 
         Caption         =   "Ok only"
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   285
         Value           =   -1  'True
         Width           =   1725
      End
   End
   Begin VB.Frame fraIcons 
      Caption         =   "Icon Styles"
      Height          =   3330
      Left            =   15
      TabIndex        =   35
      Top             =   0
      Width           =   1980
      Begin VB.OptionButton optIcon 
         Caption         =   "&No icon"
         Height          =   210
         Index           =   0
         Left            =   675
         TabIndex        =   2
         Top             =   390
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optIcon 
         Caption         =   "&Critical"
         Height          =   210
         Index           =   1
         Left            =   675
         TabIndex        =   3
         Top             =   975
         Width           =   1215
      End
      Begin VB.OptionButton optIcon 
         Caption         =   "&Question"
         Height          =   210
         Index           =   2
         Left            =   675
         TabIndex        =   4
         Top             =   1575
         Width           =   1215
      End
      Begin VB.OptionButton optIcon 
         Caption         =   "&Exclamation"
         Height          =   210
         Index           =   3
         Left            =   675
         TabIndex        =   5
         Top             =   2175
         Width           =   1215
      End
      Begin VB.OptionButton optIcon 
         Caption         =   "&Information"
         Height          =   210
         Index           =   4
         Left            =   675
         TabIndex        =   6
         Top             =   2775
         Width           =   1215
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   0
         Left            =   105
         Picture         =   "Messager.frx":0442
         Stretch         =   -1  'True
         Top             =   240
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   1
         Left            =   105
         Picture         =   "Messager.frx":074C
         Stretch         =   -1  'True
         Top             =   825
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   2
         Left            =   105
         Picture         =   "Messager.frx":0B8E
         Stretch         =   -1  'True
         Top             =   1425
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   3
         Left            =   105
         Picture         =   "Messager.frx":0FD0
         Stretch         =   -1  'True
         Top             =   2025
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   4
         Left            =   105
         Picture         =   "Messager.frx":1412
         Stretch         =   -1  'True
         Top             =   2625
         Width           =   480
      End
   End
   Begin VB.Frame fraText 
      Caption         =   "Text"
      Height          =   1245
      Left            =   2070
      TabIndex        =   27
      Top             =   0
      Width           =   4215
      Begin VB.TextBox txtPrompt 
         Height          =   690
         Left            =   735
         MaxLength       =   1024
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   3420
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   735
         MaxLength       =   1024
         TabIndex        =   0
         Top             =   165
         Width           =   3420
      End
      Begin VB.Label Label2 
         Caption         =   "Prompt:"
         Height          =   210
         Left            =   135
         TabIndex        =   34
         Top             =   525
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Title:"
         Height          =   210
         Left            =   105
         TabIndex        =   33
         Top             =   210
         Width           =   600
      End
   End
   Begin VB.Frame fraSample 
      Height          =   2070
      Left            =   2070
      TabIndex        =   26
      Top             =   1260
      Width           =   4215
      Begin VB.CommandButton cmdButton 
         Caption         =   "Sample"
         Height          =   345
         Index           =   2
         Left            =   2745
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1380
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Sample"
         Height          =   345
         Index           =   1
         Left            =   1530
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1380
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Sample"
         Height          =   345
         Index           =   0
         Left            =   315
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1380
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblDots 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   210
         Left            =   1035
         TabIndex        =   42
         Top             =   1005
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Height          =   210
         Left            =   1035
         TabIndex        =   30
         Top             =   795
         UseMnemonic     =   0   'False
         Width           =   2790
      End
      Begin VB.Image imgBoxIcon 
         Height          =   480
         Left            =   315
         Picture         =   "Messager.frx":1854
         Stretch         =   -1  'True
         Top             =   645
         Width           =   480
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "(Application name)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   180
         TabIndex        =   29
         Top             =   270
         Width           =   3585
      End
      Begin VB.Image ImgSample 
         Height          =   1740
         Left            =   105
         Picture         =   "Messager.frx":1B5E
         Stretch         =   -1  'True
         Top             =   210
         Width           =   3990
      End
   End
   Begin VB.Frame fraCode 
      Caption         =   "Source code"
      Height          =   1635
      Left            =   30
      TabIndex        =   25
      Top             =   3330
      Width           =   5745
      Begin VB.OptionButton optType 
         Caption         =   "Function"
         Height          =   210
         Index           =   1
         Left            =   4695
         TabIndex        =   20
         Top             =   210
         Width           =   960
      End
      Begin VB.OptionButton optType 
         Caption         =   "Sub"
         Height          =   210
         Index           =   0
         Left            =   3690
         TabIndex        =   19
         Top             =   210
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.TextBox txtSource 
         Height          =   1095
         Left            =   75
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   465
         Width           =   5595
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   315
      Left            =   7215
      TabIndex        =   24
      Top             =   5040
      Width           =   1110
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   315
      Left            =   1215
      TabIndex        =   23
      Top             =   5040
      Width           =   1110
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Form_Load()
    'Me.Top = (Screen.Height - Me.Height) / 2
    'Me.Left = (Screen.Width - Me.Width) / 2

    'options = Val(GetSetting("VBMessager", "Area", "Options", "DefaultValue"))

    bAlreadyCopied = True

    With MBox
        .Title = ""
        .Text = ""
        .Icon = I_NONE
        .Button = B_OK
        .Default = D_BUTTON1
        .Type = T_SUB
        .UseHelp = False
        .HelpFile = ""
        .Context = ""
        .SysModal = False
    End With

    SetDefaultOption
    RefreshSample

    Me.Show
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not bAlreadyCopied Then ' If the information wasn't saved....
        Select Case MsgBox("Copy MessageBox code to Clipboard?", 3 + 32, "VB.Messager")
        Case 7: 'No
        Case 6: 'yes
            cmdCopy_Click
        Case 2: 'cancel
            Cancel = True
            Exit Sub
        End Select
    End If
    'SaveSetting "VBMessager", "Area", "Options", "" & options
    End
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

' If do don't like this - you can easy delete it...
Private Sub cmdAbout_Click()
    frmAbout.Show vbModal
End Sub

Public Sub cmdCopy_Click()
    Dim sText As String
    Dim nHandle As Integer

    Me.Enabled = False

    On Error GoTo CopyErrorHandler

    Screen.MousePointer = vbHourglass

    sText = RTrim$(txtSource)

    ' Copy The actual text to the clipboard.
    If EmptyString(sText) Then
        MsgBox "Nothing to copy!", 0 + 48, "VB.Messager Copy"
        GoTo CopyFinished
    End If

    Clipboard.Clear
    Clipboard.SetText sText

    bAlreadyCopied = True
    Screen.MousePointer = vbDefault

    'MsgBox "Code placed into Clipboard.", 64, "VB.Messager"
    Beep

    GoTo CopyFinished

CopyErrorHandler:
    MsgBox Error$

CopyFinished:
    Me.Enabled = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdShow_Click()
    ShowMessageBox
End Sub

Private Sub RefreshSample()
    PaintSample
    txtSource = MakeSourceCode
End Sub

Private Sub PaintSample()
    If EmptyString(MBox.Title) Then
        lblTitle = "(Application name)"
    Else
        lblTitle = MBox.Title
    End If

    lblPrompt = MBox.Text

    imgBoxIcon = imgIcon(MBox.Icon).Picture

    Select Case MBox.Button
    Case B_OKCANCEL
        SetButton 0, 925, "OK"
        SetButton 1, 2140, "Cancel"
        OffButton 2
    Case B_ABORTRETRYIGNORE
        SetButton 0, 315, "Abort"
        SetButton 1, 1520, "Retry"
        SetButton 2, 2745, "Ignore"
    Case B_YESNOCANCEL
        SetButton 0, 315, "Yes"
        SetButton 1, 1520, "No"
        SetButton 2, 2745, "Cancel"
    Case B_YESNO
        SetButton 0, 925, "Yes"
        SetButton 1, 2140, "No"
        OffButton 2
    Case B_RETRYCANCEL
        SetButton 0, 825, "Retry"
        SetButton 1, 2140, "Cancel"
        OffButton 2
    Case Else            ' OK button only
        SetButton 0, 1520, "OK"
        OffButton 1
        OffButton 2
    End Select
End Sub

Private Sub SetButton(nIndex As Integer, nLeft As Integer, sCaption As String)
    cmdButton(nIndex).Caption = sCaption
    cmdButton(nIndex).Left = nLeft
    SetVisible cmdButton(nIndex), True
End Sub

Private Sub OffButton(nIndex As Integer)
    SetVisible cmdButton(nIndex), False
End Sub

Private Sub imgIcon_Click(Index As Integer)
    optIcon(Index) = True
    MBox.Icon = Index
    RefreshSample
End Sub

Private Sub optIcon_Click(Index As Integer)
    MBox.Icon = Index
    RefreshSample
End Sub

Private Sub optButton_Click(Index As Integer)
    MBox.Button = Index
    SetDefaultOption
    RefreshSample
End Sub

Private Sub cmdButton_Click(Index As Integer)
    optDefault(Index) = True
End Sub

Private Sub optDefault_Click(Index As Integer)
    MBox.Default = Index
    RefreshSample
End Sub

Private Sub SetDefOption(nIndex As Integer, sCaption As String, bMark As Boolean)
    If sCaption = "" Then
        optDefault(nIndex) = False
        SetCaption optDefault(nIndex), "n/a"
        SetEnabled optDefault(nIndex), False
    Else
        optDefault(nIndex) = bMark
        SetCaption optDefault(nIndex), sCaption & " button"
        SetEnabled optDefault(nIndex), True

        If bMark Then
            MBox.Default = nIndex
        End If
    End If
End Sub

Private Sub SetDefaultOption()
    Select Case MBox.Button
    Case B_OKCANCEL
        SetDefOption 0, "OK", False
        SetDefOption 1, "Cancel", True
        SetDefOption 2, "", False
    Case B_ABORTRETRYIGNORE
        SetDefOption 0, "Abort", False
        SetDefOption 1, "Retry", True
        SetDefOption 2, "Ignore", False
    Case B_YESNOCANCEL
        SetDefOption 0, "Yes", False
        SetDefOption 1, "No", False
        SetDefOption 2, "Cancel", True
    Case B_YESNO
        SetDefOption 0, "Yes", False
        SetDefOption 1, "No", True
        SetDefOption 2, "", False
    Case B_RETRYCANCEL
        SetDefOption 0, "Retry", True
        SetDefOption 1, "Cancel", False
        SetDefOption 2, "", False
    Case Else            ' OK button only
        SetDefOption 0, "OK", True
        SetDefOption 1, "", False
        SetDefOption 2, "", False
    End Select
End Sub

Private Sub optType_Click(Index As Integer)
    If Index = 0 And MBox.Type = T_FUNCTION Then
        MBox.Type = T_SUB
        txtSource = MakeSourceCode
    ElseIf Index = 1 And MBox.Type = T_SUB Then
        MBox.Type = T_FUNCTION
        txtSource = MakeSourceCode
    End If
End Sub

Private Sub txtPrompt_Change()
    lblPrompt = txtPrompt
    SetVisible lblDots, (InStr(txtPrompt, Chr$(13)) > 0)
    MBox.Text = txtPrompt
    txtSource = MakeSourceCode
End Sub

Private Sub txtTitle_Change()
    If Len(Trim(txtTitle.Text)) = 0 Then
        lblTitle = "(Application name)"
    Else
        lblTitle = txtTitle.Text
    End If
    MBox.Title = txtTitle.Text
    txtSource = MakeSourceCode
End Sub

Private Sub chkHelpFile_Click()
    If chkHelpFile = vbChecked Then
        SetEnabled lblHelpFile, True
        SetEnabled lblContext, True
        SetEnabled txtContext, True
        MBox.UseHelp = True

        PickHelpFile

        If EmptyString(MBox.HelpFile) Then
            chkHelpFile = vbUnchecked
            lblHelpFile.Caption = "(No help file)"
        End If
    End If

    If chkHelpFile = vbUnchecked Then
        SetEnabled lblHelpFile, False
        SetEnabled lblContext, False
        SetEnabled txtContext, False
        MBox.UseHelp = False
    End If
    txtSource = MakeSourceCode
End Sub

Private Sub txtContext_Change()
    MBox.Context = txtContext
    txtSource = MakeSourceCode
End Sub

Private Sub chkModal_Click()
    MBox.SysModal = (chkModal = vbChecked)
    txtSource = MakeSourceCode
End Sub

Private Sub PickHelpFile()
    Dim sHelpFile As String

    sHelpFile = OpenDialog("Help files (*.hlp)|*.hlp|All files (*.*)|*.*", _
                           "Select help file", _
                           MBox.HelpFile)

    If Len(sHelpFile) = 0 Then Exit Sub

    MBox.HelpFile = sHelpFile
    lblHelpFile.Caption = ExtractFileName(sHelpFile)
End Sub
