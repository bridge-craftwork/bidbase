VERSION 5.00
Begin VB.Form frmMsg 
   Caption         =   "CardShark BidBase"
   ClientHeight    =   2505
   ClientLeft      =   2835
   ClientTop       =   3825
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   167
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   355
   Visible         =   0   'False
   Begin VB.TextBox t_MsgInput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   960
      TabIndex        =   3
      Text            =   "MsgInput"
      Top             =   1725
      Visible         =   0   'False
      Width           =   4100
   End
   Begin VB.CommandButton b_MsgInput 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   165
      TabIndex        =   2
      Top             =   1725
      Width           =   650
   End
   Begin VB.CommandButton b_MsgInput 
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   1275
      Visible         =   0   'False
      Width           =   650
   End
   Begin VB.CommandButton b_MsgInput 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   1770
      TabIndex        =   0
      Top             =   1275
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label L_MsgBox 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Message goes here. 1-line height = 360. 2-line height = 650.                             3-line height = 950."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   120
      TabIndex        =   4
      Top             =   105
      Width           =   5040
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private int_MsgClick As Integer

' Stay on top: '
Private Declare Function SetWindowPos Lib "user32" _
        (ByVal hwnd As Long, _
         ByVal hWndInsertAfter As Long, _
         ByVal x As Long, Y, _
         ByVal cx As Long, _
         ByVal cy As Long, _
         ByVal wFlags As Long) _
         As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Sub MakeNormal(hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Sub MakeTopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Private Sub b_MsgInput_Click(Index As Integer)
    int_MsgClick = Index
    b_MsgInput(Index).Tag = Index
    Me.Visible = False
End Sub

Private Sub Form_Load()
    MakeTopMost hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub
