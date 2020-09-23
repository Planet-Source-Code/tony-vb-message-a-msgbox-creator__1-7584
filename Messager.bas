Attribute VB_Name = "modMessager"
Option Explicit

' Button constants
Public Const B_OK As Integer = 0
Public Const B_OKCANCEL As Integer = 1
Public Const B_ABORTRETRYIGNORE As Integer = 2
Public Const B_YESNOCANCEL As Integer = 3
Public Const B_YESNO As Integer = 4
Public Const B_RETRYCANCEL As Integer = 5

' Icon constants
Public Const I_NONE As Integer = 0
Public Const I_CRITICAL As Integer = 1
Public Const I_QUESTION As Integer = 2
Public Const I_EXCLAMATION As Integer = 3
Public Const I_INFORMATION As Integer = 4

' Default button constants
Public Const D_BUTTON1 As Integer = 0
Public Const D_BUTTON2 As Integer = 1
Public Const D_BUTTON3 As Integer = 2

' Type constants
Public Const T_SUB As Integer = 0
Public Const T_FUNCTION As Integer = 1

' MessageBox data
Type MBState
    Title As String
    Text As String
    Icon As Integer
    Button As Integer
    Default As Integer
    Type As Integer
    UseHelp As Boolean
    HelpFile As String
    Context As String
    SysModal As Boolean
End Type
Public MBox As MBState

' Was the code already copied to the clipboard? (Used in the mnuExit procedure)
Public bAlreadyCopied As Boolean

' API: Used to force window on top.
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long

Function MakeSourceCode() As String
    Dim sQuote As String, sPrompt As String, sSettings As String, sTitle As String, _
        sHelp As String, sText As String

    sQuote = Chr$(34)

    ' Prompt... (Convert non-text ascii to chr$ string)
    sPrompt = CodeString(RTrim$(MBox.Text))

    ' Icon, buttons and other settings
    Select Case MBox.Button
    Case B_OKCANCEL
        sSettings = "vbOKCancel"
    Case B_ABORTRETRYIGNORE
        sSettings = "vbAbortRetryIgnore"
    Case B_YESNOCANCEL
        sSettings = "vbYesNoCancel"
    Case B_YESNO
        sSettings = "vbYesNo"
    Case B_RETRYCANCEL
        sSettings = "vbRetryCancel"
    Case Else
        sSettings = ""
    End Select

    Select Case MBox.Icon
    Case I_CRITICAL
        sSettings = IIf(EmptyString(sSettings), "", sSettings & " + ") & "vbCritical"
    Case I_QUESTION
        sSettings = IIf(EmptyString(sSettings), "", sSettings & " + ") & "vbQuestion"
    Case I_EXCLAMATION
        sSettings = IIf(EmptyString(sSettings), "", sSettings & " + ") & "vbExclamation"
    Case I_INFORMATION
        sSettings = IIf(EmptyString(sSettings), "", sSettings & " + ") & "vbInformation"
    End Select

    Select Case MBox.Default
    Case D_BUTTON2
        sSettings = IIf(EmptyString(sSettings), "", sSettings & " + ") & "vbDefaultButton2"
    Case D_BUTTON3
        sSettings = IIf(EmptyString(sSettings), "", sSettings & " + ") & "vbDefaultButton3"
    End Select

    If MBox.SysModal Then
        sSettings = IIf(EmptyString(sSettings), "", sSettings & " + ") & "vbSystemModal"
    End If

    If Not EmptyString(sSettings) Then
        sSettings = ", " & sSettings
    End If

    ' Box title
    If EmptyString(MBox.Title) Then
        sTitle = ""
    Else
        sTitle = ", " & sQuote & RTrim(MBox.Title) & sQuote
    End If

    ' Help file and context
    sHelp = ""
    If MBox.UseHelp Then
        If Not EmptyString(MBox.HelpFile) Then
            sHelp = ", " & sQuote & MBox.HelpFile & sQuote & _
                    ", " & sQuote & RTrim(MBox.Context) & sQuote
        End If
    End If

    ' Put the thing together
    sText = sPrompt

    If EmptyString(sSettings) Then
        If Not EmptyString(sTitle & sHelp) Then
            sText = sText & ", "
        End If
    Else
        sText = sText & sSettings
    End If

    If EmptyString(sTitle) Then
        If Not EmptyString(sHelp) Then
            sText = sText & ", "
        End If
    Else
        sText = sText & sTitle
    End If

    If Not EmptyString(sHelp) Then
        sText = sText & sHelp
    End If

    Select Case MBox.Type
    Case T_SUB
        'MsgBox "This can be done with a message box", vbInformation, "VB.Messager"

        MakeSourceCode = "MsgBox " & sText

    Case T_FUNCTION
        'Select Case MsgBox("Copy MessageBox code to Clipboard?", vbYesNoCancel & vbQuestion, "VB.Messager")
        'Case vbYes
        'Case vbNo
        'Case vbCancel
        'End Select

        sText = "Select Case MsgBox(" & sText & ")"

        Select Case MBox.Button
        Case B_OKCANCEL
            sText = sText & vbCrLf & "Case vbOK"
            sText = sText & vbCrLf & "Case vbCancel"
        Case B_ABORTRETRYIGNORE
            sText = sText & vbCrLf & "Case vbAbort"
            sText = sText & vbCrLf & "Case vbRetry"
            sText = sText & vbCrLf & "Case vbIgnore"
        Case B_YESNOCANCEL
            sText = sText & vbCrLf & "Case vbYes"
            sText = sText & vbCrLf & "Case vbNo"
            sText = sText & vbCrLf & "Case vbCancel"
        Case B_YESNO
            sText = sText & vbCrLf & "Case vbYes"
            sText = sText & vbCrLf & "Case vbNo"
        Case B_RETRYCANCEL
            sText = sText & vbCrLf & "Case vbRetry"
            sText = sText & vbCrLf & "Case vbCancel"
        Case Else
            sText = sText & vbCrLf & "Case vbOK"
        End Select
        sText = sText & vbCrLf & "End Select"

        MakeSourceCode = sText

    Case Else
        MakeSourceCode = ""
    End Select

End Function

Sub ShowMessageBox()
    Dim sPrompt As String, sTitle As String, sHelp As String, sContext As String
    Dim nSettings As Integer

    ' Prompt...
    sPrompt = RTrim$(MBox.Text)

    ' Icon, buttons and other settings
    Select Case MBox.Button
    Case B_OKCANCEL
        nSettings = vbOKCancel
    Case B_ABORTRETRYIGNORE
        nSettings = vbAbortRetryIgnore
    Case B_YESNOCANCEL
        nSettings = vbYesNoCancel
    Case B_YESNO
        nSettings = vbYesNo
    Case B_RETRYCANCEL
        nSettings = vbRetryCancel
    Case Else
        nSettings = vbOKOnly
    End Select

    Select Case MBox.Icon
    Case I_CRITICAL
        nSettings = nSettings + vbCritical
    Case I_QUESTION
        nSettings = nSettings + vbQuestion
    Case I_EXCLAMATION
        nSettings = nSettings + vbExclamation
    Case I_INFORMATION
        nSettings = nSettings + vbInformation
    End Select

    Select Case MBox.Default
    Case D_BUTTON2
        nSettings = nSettings + vbDefaultButton2
    Case D_BUTTON3
        nSettings = nSettings + vbDefaultButton3
    End Select

    If MBox.SysModal Then
        nSettings = nSettings + vbSystemModal
    End If

    ' Box title
    If EmptyString(MBox.Title) Then
        sTitle = "(Application name)"
    Else
        sTitle = RTrim(MBox.Title)
    End If

    ' Help file and context
    sHelp = ""
    sContext = ""
    If MBox.UseHelp Then
        If Not EmptyString(MBox.HelpFile) Then
            sHelp = MBox.HelpFile
            sContext = RTrim(MBox.Context)
        End If
    End If

    ' Show this thing
    If EmptyString(sHelp) Then
        MsgBox sPrompt, nSettings, sTitle
    Else
        MsgBox sPrompt, nSettings, sTitle, sHelp, sContext
    End If

End Sub

Function EmptyString(ByRef sText As String) As Boolean
    If IsNull(sText) Then
        EmptyString = True
    Else
        EmptyString = (Len(Trim(sText)) = 0)
    End If
End Function

Function CodeString(sText As String) As String
    Dim sQuote As String, sString As String
    Dim nSize As Integer, i As Integer, nChar As Integer
    Dim bAscii As Boolean

    sQuote = Chr$(34)
    sString = ""
    nSize = Len(sText)
    bAscii = True           ' Forces quote in the beginning

    For i = 1 To nSize
        nChar = Asc(Mid$(sText, i, 1))
        If nChar < 32 Then
            If bAscii Then
                If i > 1 Then
                    sString = sString & " & "
                End If
            ElseIf i > 1 Then
                sString = sString & sQuote & " & "
            End If
            Select Case nChar
            Case 13           ' Carriage-return/linefeed (vbCrLf = Chr$(13) & Chr$(10))
                ' Carriage return (vbCr = Chr$(13))
                If i < nSize Then
                    If Asc(Mid$(sText, i + 1, 1)) = 10 Then
                        sString = sString & "vbCrLf"
                        i = i + 1
                    Else
                        sString = sString & "vbCr"
                    End If
                Else
                    sString = sString & "vbCr"
                End If
            Case 0            ' Null character (vbNullChar = Chr$(0))
                sString = sString & "vbNullChar"
            Case 8            ' Backspace (vbBack = Chr$(8))
                sString = sString & "vbBack"
            Case 9            ' Tab (vbTab = Chr$(9))
                sString = sString & "vbTab"
            Case 10           ' Linefeed (vbLf = Chr$(10))
                sString = sString & "vbLf"
            Case 11           ' Vertical tab (vbVerticalTab = Chr$(11))
                sString = sString & "vbVerticalTab"
            Case 12           ' Form feed (vbFormFeed = Chr$(12)
                sString = sString & "vbFormFeed"
            Case Else
                sString = sString & "Chr$(" & nChar & ")"
            End Select
            bAscii = True
        Else
            If bAscii Then
                If i > 1 Then
                    sString = sString & " & "
                End If
                sString = sString & sQuote
            End If
            sString = sString & Chr$(nChar)
            bAscii = False
        End If
    Next
    If Not bAscii Then sString = sString & sQuote

    CodeString = sString

End Function

' This procedure will set the visible prop of a control
' Use this procedure to reduce control "flicker".
Sub SetVisible(ctrlIn As Control, iTrueFalse As Integer, Optional sCaption)
    Dim iCompare As Integer
    iCompare = Not iTrueFalse

    If ctrlIn.Visible = iCompare Then
        ctrlIn.Visible = iTrueFalse
    End If

    If Not IsMissing(sCaption) Then
        If ctrlIn.Caption <> sCaption Then
            ctrlIn.Caption = sCaption
        End If
    End If
End Sub

' This procedure will set the enabled prop of a control
' Use this procedure to reduce control "flicker".
Sub SetEnabled(ctrlIn As Control, ByVal iTrueFalse As Integer, Optional sCaption)
    Dim iCompare As Integer
    iCompare = Not iTrueFalse

    If ctrlIn.Enabled = iCompare Then
        ctrlIn.Enabled = iTrueFalse
    End If

    If Not IsMissing(sCaption) Then
        If ctrlIn.Caption <> sCaption Then
            ctrlIn.Caption = sCaption
        End If
    End If
End Sub

' Use this procedure to reduce control "flicker".
Sub SetCaption(ctrlIn As Control, sCaption As String)
    If ctrlIn.Caption <> sCaption Then
        ctrlIn.Caption = sCaption
    End If
End Sub

Function ExtractFileName(sFileIn As String) As String
    Dim i As Integer
    For i = Len(sFileIn) To 1 Step -1
        If InStr("\", Mid$(sFileIn, i, 1)) Then Exit For
    Next
    ExtractFileName = Mid$(sFileIn, i + 1, Len(sFileIn) - i)
End Function

' Extracts the path section of a file-string
Function ExtractPath(sPathIn As String) As String
   Dim i As Integer
   For i = Len(sPathIn) To 1 Step -1
      If InStr(":\", Mid$(sPathIn, i, 1)) Then Exit For
   Next
   ExtractPath = Left$(sPathIn, i)
End Function

' To force a windows to stay on top. (usefull for splash screens)
Sub FormStayOnTop(frmHandle As Form, bOnTop As Boolean)
    Dim nFlags As Integer
    nFlags = 2 Or 1

    On Error Resume Next

    If bOnTop Then
        SetWindowPos frmHandle.hwnd, -1, 0, 0, 0, 0, nFlags
    Else
        SetWindowPos frmHandle.hwnd, -2, 0, 0, 0, 0, nFlags
    End If
End Sub

'Centre form on screen
'Sub CentreForm(frmHandle As Form)
'    frmHandle.Move (Screen.Width - frmHandle.Width) / 2, (Screen.Height - frmHandle.Height) / 2
'End Sub
