VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Add Files"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Columns         =   1
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6000
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' deja_vu
' feeedback : deja_vu555@yahoo.com
'
' Please read the comments, they really help..

' All the commondialog open/save flags from MSDN
'
'cdlOFNAllowMultiselect &H200 Specifies that the File Namelist box allows multiple selections.
'The user can select more than one file atrun time by pressing the SHIFT key and using the UP ARROW and DOWN ARROW keys to select the desired files. When this is done, the FileName property returns a string containing the names of all selected files. The names in the string are delimited by spaces.
'
'cdlOFNCreatePrompt &H2000 Specifies that the dialog box prompts the user to create a file that doesn't currently exist. This flag automatically sets the cdlOFNPathMustExist and cdlOFNFileMustExist flags.
'cdlOFNExplorer &H80000 Use the Explorer-like Open A File dialog box template. Works with Windows 95 and Windows NT 4.0.
'CdlOFNExtensionDifferent &H400 Indicates that the extension of the returned filename is different from the extension specified by the DefaultExt property. This flag isn't set if the DefaultExt property is Null, if the extensions match, or if the file has no extension. This flag value can be checked upon closing the dialog box.
'cdlOFNFileMustExist &H1000 Specifies that the user can enter only names of existing files in the File Name text box. If this flag is set and the user enters an invalid filename, a warning is displayed. This flag automatically sets the cdlOFNPathMustExist flag.
'cdlOFNHelpButton &H10 Causes the dialog box to display the Help button.
'cdlOFNHideReadOnly &H4 Hides the Read Onlycheck box.
'cdlOFNLongNames &H200000 Use long filenames.
'cdlOFNNoChangeDir &H8 Forces the dialog box to set the current directory to what it was when the dialog box was opened.
'CdlOFNNoDereferenceLinks &H100000 Do not dereference shell links (also known as shortcuts). By default, choosing a shell link causes it to be dereferenced by the shell.
'cdlOFNNoLongNames &H40000 No long file names.
'CdlOFNNoReadOnlyReturn &H8000 Specifies that the returned file won't have the Read Only attribute set and won't be in a write-protected directory.
'cdlOFNNoValidate &H100 Specifies that the common dialog box allows invalid characters in the returned filename.
'cdlOFNOverwritePrompt &H2 Causes the Save As dialog box to generate a message box if the selected file already exists. The user must confirm whether to overwrite the file.
'cdlOFNPathMustExist &H800 Specifies that the user can enter only valid paths. If this flag is set and the user enters an invalid path, a warning message is displayed.
'cdlOFNReadOnly &H1 Causes the Read Only check box to be initially checked when the dialog box is created. This flag also indicates the state of the Read Only check box when the dialog box is closed.
'cdlOFNShareAware &H4000 Specifies that sharing violation errors will be ignored.
'
'
' first of all let me tell you this is NOT at all complex..
'

Private Sub Command1_Click()
'On Error GoTo CdlEr

Dim BufferFileArray() As String
Dim i As Integer

With CommonDialog1
    .DialogTitle = "Add Multiple files..."
    .Filter = "All Files(*.*)|*.*"
    
    ' if you dont use cdlOFNExplorer then you get an old style box and
    ' the return filenames will be in 8.3 format
    .Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly
    .InitDir = CurDir
    
    ' this needs to be high since lots of file names
    ' are being returned..
    .MaxFileSize = 32767
    
    .FileName = ""
    
    ' shows the open dilaog box..
    .ShowOpen
     
    ' adds all items to an array
    BufferFileArray = Split(.FileName, Chr(0))
    ' NOTE: if you dont use cdlOFNExplorer flag the Chr(0) in the above
    ' line should be changed to Chr(32)..
End With

'Hey!! what if the guy only selected one file??
If UBound(BufferFileArray) = 0 Then GoTo SkipItAll

'formatting..
For i = LBound(BufferFileArray) + 1 To UBound(BufferFileArray)
    BufferFileArray(i) = Fixp(BufferFileArray(0)) & BufferFileArray(i)
Next i


'after formating using a for loop we add each file name to the listbox
'we wont take the first element in the array which happen to be the curdir..
        For i = LBound(BufferFileArray) + 1 To UBound(BufferFileArray)
            'if the selected file is not in the list then skips it
            If lstVal(List1, BufferFileArray(i)) = False Then
                'if the length of the first item in the BufferFileArray
                'array is not equal to the filepath without the filename
                'then we cut the first len(BufferFileArray(0)) off from
                'BufferFileArray(i) - this is used because sometimes when
                'we select windows shortcuts(.lnk) or DOS shortcuts(.pif)
                'we get a string with the root folder + the original
                'file path
                If Len(Fixp(BufferFileArray(0))) <> Len(BufferFileArray(i)) - Len(getFileNameOutofFilePath(BufferFileArray(i))) Then
                    'like i said above.. we have to check if we added the
                    'actual filepaths of any lnk or pif..
                    If lstVal(List1, Right(BufferFileArray(i), Len(BufferFileArray(i)) - Len(Fixp(BufferFileArray(0))))) = False Then
                        List1.AddItem Right(BufferFileArray(i), Len(BufferFileArray(i)) - Len(Fixp(BufferFileArray(0))))
                    End If
                Else
                'if everything is noraml then adds the filepath to the list
                    List1.AddItem BufferFileArray(i)
                End If
            End If
        Next i

Exit Sub

SkipItAll:

If lstVal(List1, CommonDialog1.FileName) = False Then
    List1.AddItem CommonDialog1.FileName
End If

Exit Sub
CdlEr:
'
'error handling here
'
Exit Sub

End Sub


'
' dont look here
' :p
'

Public Function getFileNameOutofFilePath(path As String, Optional retval As Integer = 1) As String
Dim i As Integer
If Trim(path) = "" Then Exit Function

i = Len(path)

If retval = 0 Then
    Do
        If InStr(i, path, "\") <> 0 Then
            getFileNameOutofFilePath = Left(path, i)
            Exit Function
        End If
        i = i - 1
    Loop Until i <= 0
    
ElseIf retval = 1 Then
    Do
        If InStr(i, path, "\") Then
            getFileNameOutofFilePath = Right(path, Len(path) - i)
            Exit Function
        End If
        i = i - 1
    Loop Until i <= 0
End If

End Function
'
'function checks if the string 'chk' is in the listbox 'lst'..
'
Public Function lstVal(lst As ListBox, chk As String) As Boolean
Dim i As Integer
If lst.ListCount <= 0 Then
    lstVal = False
    Exit Function
End If

For i = 0 To lst.ListCount
If LCase(chk) = LCase(lst.List(i)) Then lstVal = True: Exit Function
Next i
End Function

'fix's file paths :B
Public Function Fixp(path As String, Optional xp As Integer = 1) As String

If xp = 1 Then
    If Right(path, 1) = "\" Then
        Fixp = path
    Else
        Fixp = path & "\"
    End If
Else
    If Right(path, 1) = "\" Then
        Fixp = Left(path, Len(path) - 1)
    Else
        Fixp = path
    End If
End If
    
End Function
