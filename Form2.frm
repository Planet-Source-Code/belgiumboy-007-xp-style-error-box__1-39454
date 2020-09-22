VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ERROR Box"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2175
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   2175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Create ERROR"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fso As New FileSystemObject
Private strm As TextStream

Private Sub Command1_Click()
On Error GoTo a
    Set strm = fso.OpenTextFile("", ForAppending)
    Exit Sub
    
a:
    With Form1
        .Number = Err.Number
        .Description = Err.Description
        .Source = Err.Source
        .HelpContext = Err.HelpContext
        .HelpFile = Err.HelpFile
        .LastDllError = Err.LastDllError
        .Show vbModal
    End With
End Sub
