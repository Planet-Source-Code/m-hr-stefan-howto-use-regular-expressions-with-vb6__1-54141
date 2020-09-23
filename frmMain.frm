VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Regex"
   ClientHeight    =   5295
   ClientLeft      =   5115
   ClientTop       =   3720
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7440
   Begin RichTextLib.RichTextBox txtValidate 
      Height          =   3885
      Left            =   0
      TabIndex        =   6
      Top             =   30
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   6853
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":0000
   End
   Begin MSComDlg.CommonDialog comDiag 
      Left            =   6750
      Top             =   4470
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtRegex 
      Height          =   345
      Left            =   2160
      TabIndex        =   5
      Top             =   3990
      Width           =   3315
   End
   Begin VB.OptionButton optRegex 
      Caption         =   "[0-9]"
      Height          =   345
      Index           =   1
      Left            =   90
      TabIndex        =   4
      ToolTipText     =   "Matches any Number"
      Top             =   4410
      Width           =   1635
   End
   Begin VB.OptionButton optRegex 
      Caption         =   "[A-Z]"
      Height          =   345
      Index           =   2
      Left            =   90
      TabIndex        =   3
      ToolTipText     =   "Match one or more uppercase letter"
      Top             =   4830
      Width           =   1635
   End
   Begin VB.OptionButton optRegex 
      Caption         =   "Option1"
      Height          =   345
      Index           =   3
      Left            =   1830
      TabIndex        =   2
      Top             =   3990
      Width           =   255
   End
   Begin VB.OptionButton optRegex 
      Caption         =   "bug.*"
      Height          =   345
      Index           =   0
      Left            =   90
      TabIndex        =   1
      ToolTipText     =   "Search for bug, bugs, bugfix"
      Top             =   3990
      Width           =   1635
   End
   Begin VB.CommandButton cmdValidate 
      Caption         =   "&Validate"
      Height          =   375
      Left            =   5730
      TabIndex        =   0
      Top             =   3990
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--> Use of Regular Expressions
'--> Set References to
'--> Microsoft VB Script Regular expressions 5.5
'--> Microsoft Scripting Runtime
'--> Additional Components
'--> Microsoft Rich Textbox
'--> Microsoft Common Dialog Control

'************************************
'bug.* = find all words starting with bug
'[0-9] = find all numbers in text
'[A-Z] = find all uppercase letters in text
'************************************

Private bolFileLoad As Boolean
Private Const ExampleFile As String = "\ExFile.txt"

Private Sub cmdValidate_Click()

Dim oRegex As RegExp
Dim m As Match
Dim strPattern As String

  If Not bolFileLoad Then
    '--> Check if there is a file loaded
    If Len(Me.txtValidate.Text) = 0 Then
      MsgBox "There is nothing to validate!", vbCritical + vbOKOnly, "Error"
      Exit Sub
    End If
  End If
  
  '--> Choose the pattern
  If Me.optRegex(0).Value Then
    strPattern = Me.optRegex(0).Caption
  ElseIf Me.optRegex(1).Value Then
    strPattern = Me.optRegex(1).Caption
  ElseIf Me.optRegex(2).Value Then
    strPattern = Me.optRegex(2).Caption
  ElseIf Me.optRegex(3).Value Then
    strPattern = Me.txtRegex.Text
  End If
  
  If Len(Trim(strPattern)) = 0 Then Exit Sub
  
  On Error GoTo errHandler
  
  '--> Set back validation color
  Me.txtValidate.SelStart = 1
  Me.txtValidate.SelLength = Len(Me.txtValidate.Text)
  Me.txtValidate.SelColor = vbBlack
  
  '--> Creating and running the regex Object
  Set oRegex = New RegExp
  With oRegex
    .Pattern = strPattern
    .Global = True
    '--> Run through all matches
    For Each m In .Execute(Me.txtValidate.Text)
      Me.txtValidate.SelStart = m.FirstIndex
      Me.txtValidate.SelLength = m.Length
      Me.txtValidate.SelColor = vbRed
    Next
  End With
  
  Set oRegex = Nothing
  
  Exit Sub
  
errHandler:
  MsgBox Err.Description & vbCrLf & Err.Number, vbCritical + vbOKOnly, "Error"
  Set oRegex = Nothing
  Err.Clear

End Sub

Private Sub Form_Load()
  bolFileLoad = False
  '--> Load the example file
  Call StreamIn(App.Path & ExampleFile)
End Sub

Private Sub mnuExit_Click()
  End
End Sub

Private Sub mnuLoad_Click()

Dim strFile As String

'--> Load TextFile
  With Me.comDiag
    .DialogTitle = "Load TextFile"
    .DefaultExt = "txt"
    .Filter = "Text (*.txt)|*.txt"
    .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist
    On Error Resume Next
      .CancelError = True
      .ShowOpen
      strFile = .FileName
      If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
  End With

  Call StreamIn(strFile)

End Sub

Private Sub StreamIn(ByVal file As String)

Dim oFso As FileSystemObject
Dim oStream As TextStream

  On Error GoTo errHandler

  '--> Load Textfile over Scripting Runtime
  Set oFso = New FileSystemObject
  If oFso.FileExists(file) Then
    Set oStream = oFso.OpenTextFile(file, ForReading, False)
    Me.txtValidate.Text = oStream.ReadAll
    oStream.Close
    Set oStream = Nothing
  End If
  
  '--> doing Cleanupwork
  Set oFso = Nothing
  
  bolFileLoad = True
  
Exit Sub
  
errHandler:
  '--> Error Handling
  MsgBox Err.Description & vbCrLf & Err.Number, vbCritical + vbOKOnly, "Error"
  Set oStream = Nothing
  Set oFso = Nothing
  Err.Clear

End Sub
