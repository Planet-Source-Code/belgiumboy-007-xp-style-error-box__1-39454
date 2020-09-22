VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ERROR"
   ClientHeight    =   360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1170
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   360
   ScaleWidth      =   1170
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   1300
      Left            =   190
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2000
      Width           =   6040
   End
   Begin Project1.chameleonButton cmdDetails 
      Height          =   315
      Left            =   4970
      TabIndex        =   1
      Top             =   1320
      Width           =   1245
      _extentx        =   2196
      _extenty        =   556
      btype           =   3
      tx              =   "Show &Details >>"
      enab            =   -1  'True
      font            =   "Form1.frx":0000
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   14215660
      bcolo           =   14215660
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "Form1.frx":002C
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin Project1.chameleonButton cmdOK 
      Default         =   -1  'True
      Height          =   315
      Left            =   3630
      TabIndex        =   0
      Top             =   1320
      Width           =   1245
      _extentx        =   2196
      _extenty        =   556
      btype           =   3
      tx              =   "OK"
      enab            =   -1  'True
      font            =   "Form1.frx":004A
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   14215660
      bcolo           =   14215660
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "Form1.frx":0076
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   190
      X2              =   6230
      Y1              =   1820
      Y2              =   1820
   End
   Begin VB.Label Label1 
      Caption         =   "An error has occurred"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.Image imgError 
      Height          =   480
      Left            =   190
      Picture         =   "Form1.frx":0094
      Top             =   200
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Number As Integer
Public Description As String
Public Source As String
Public HelpContext As String
Public HelpFile As String
Public LastDllError As String

Private Sub cmdDetails_Click()
    If cmdDetails.Caption = "Hide &Details <<" Then
        cmdDetails.Caption = "Show &Details >>"
        Me.Height = cmdDetails.Height + 1980
    Else
        cmdDetails.Caption = "Hide &Details <<"
        Me.Height = 3975
    End If
End Sub

Private Sub cmdOK_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub Form_Load()
    Text1.Text = "Error Details:" & vbCrLf & vbCrLf & _
        "Number: " & Number & vbCrLf & _
        "Description: " & Description & vbCrLf & _
        "Source: " & Source & vbCrLf & _
        "HelpContext: " & HelpContext & vbCrLf & _
        "HelpFile: " & HelpFile & vbCrLf & _
        "LastDllError: " & LastDllError
    With Me
        .Height = cmdDetails.Height + 1980
        .Width = 6520
        .Left = (Screen.Width - .Width) / 2
        .Top = (Screen.Height - .Height) / 2
    End With
End Sub
