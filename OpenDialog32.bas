Attribute VB_Name = "OpenDialog32"
Option Explicit

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOLONGNAMES = &H40000     ' force no long names for 4.x modules
Private Const OFN_EXPLORER = &H80000        ' new look commdlg
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_LONGNAMES = &H200000      ' force long names for 3.x modules
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0

Private Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_NODEREFERENCELINKS Or OFN_READONLY Or OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Function OpenDialog(ByVal Filter As String, ByVal Title As String, ByVal InitPathFile As String) As String
'Private Function OpenLibrary(sLibrary As String) As String
    Dim pFileDialog As OPENFILENAME
    Dim dl As Long

    With pFileDialog
        ' Set up the data structure before you call the GetOpenFileName
        .lStructSize = Len(pFileDialog)

        ' If the OpenFile Dialog box is linked to a form use this line.
        ' It will pass the forms window handle.
        .hwndOwner = Screen.ActiveForm.hwnd
        '
        ' If the OpenFile Dialog box is not linked to any form use this line.
        ' It will pass a null pointer.
        '.hwndOwner = 0&
        .hInstance = App.hInstance

        If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
        For dl = 1 To Len(Filter)
            If Mid$(Filter, dl, 1) = "|" Then Mid$(Filter, dl, 1) = vbNullChar
        Next
        .lpstrFilter = Filter
        .nFilterIndex = 1

        .lpstrFile = Left$(ExtractFileName(InitPathFile) & Space$(1023), 1023) & vbNullChar ' Allocate string space for the returned strings.
        .nMaxFile = Len(pFileDialog.lpstrFile)
        .lpstrFileTitle = Space$(1023) & vbNullChar
        .nMaxFileTitle = Len(pFileDialog.lpstrFileTitle)

        .lpstrTitle = Title & vbNullChar                                                    ' Give the dialog a caption title.
        .flags = OFS_FILE_OPEN_FLAGS

        .lpstrDefExt = vbNullChar
        .lpstrInitialDir = ExtractPath(InitPathFile) & vbNullChar
    End With

    ' This will pass the desired data structure to the Windows API,
    ' which will in turn it uses to display the Open Dialog form.

    dl = GetOpenFileName(pFileDialog)

    If (dl) Then
       ' Note that sFileName will have an embedded null character at the end.
       ' You may wish to strip this character from the string.
       OpenDialog = StripTerminator(pFileDialog.lpstrFile)
    Else
       OpenDialog = ""
    End If
End Function

Private Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, vbNullChar)
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
