VERSION 5.00
Begin VB.Form frmBidBaseAnalysis 
   BackColor       =   &H00400000&
   Caption         =   "Bid/Hand Analysis"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ck_ShowDeactivated 
      Caption         =   "Don't show deact. entries"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2630
      TabIndex        =   3
      Top             =   6700
      Width           =   3030
   End
   Begin VB.CommandButton cmd_Continue 
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5775
      TabIndex        =   2
      Top             =   6700
      Width           =   2100
   End
   Begin VB.CommandButton cmd_Cancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   270
      TabIndex        =   1
      Top             =   6700
      Width           =   2100
   End
   Begin VB.ListBox lstAnalysis 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6060
      Left            =   270
      TabIndex        =   0
      Top             =   315
      Width           =   7665
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   2490
      TabIndex        =   4
      Top             =   6705
      Width           =   1215
   End
End
Attribute VB_Name = "frmBidBaseAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Declare an event: '
Event NotifyClients(ByVal int_Notify As Integer)
Dim lng_Top As Long
Dim lng_Left As Long

Private Declare Function SetWindowPos Lib "user32" _
        (ByVal hwnd As Long, _
         ByVal hWndInsertAfter As Long, _
         ByVal X As Long, y, _
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
    'mStayOnTop.Checked = False
End Sub


Public Sub MakeTopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    'mStayOnTop.Checked = True
End Sub


Private Sub ck_ShowDeactivated_Click()
    cmd_Continue_Click
End Sub


Private Sub cmd_Cancel_Click()
    cmd_Cancel.Visible = False
End Sub


Private Sub cmd_Continue_Click()
    Open "BidAnalysis.cfg" For Output As #1
    Print #1, Me.Left, Me.Top
    Close #1
    Beep
    
    int_Notify = 2
    RaiseEvent NotifyClients(int_Notify)
    DoEvents
    Me.Hide
End Sub


Private Sub Form_Load()
    Dim a As String
    Dim X As Long
    Dim y As Long
    
    If Dir$("BidAnalysis.cfg") <> "" Then
        Open "BidAnalysis.cfg" For Input As #1
        Input #1, X, y
        Close
        Me.Move X, y
    End If
    MakeTopMost hwnd ' Stay On Top
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' The QueryUnload event fires just before a form unloads from memory. '
    ' It fires just before the Unload event. Its main purpose is to let   '
    ' you detect why the form is being unloaded and to programmatically   '
    ' halt unloading if necessary. The QueryUnload event procedure takes  '
    ' two parameters:                                                     '
    ' 1. Cancel: When Cancel is False, it means that the unloading will   '
    '    continue. You can set it to True to stop the form from unloading.'
    ' 2. UnloadMode: This parameter can take several values corresponding '
    '    to how the QueryUnload event was triggered. You can compare      '
    '    UnloadMode's value with one of the following constants:          '
    '       vbFormControlMenu:  The user is closing the form.             '
    '       vbFormCode:  program code is closing the form.                '
    '       vbAppWindows:  The Windows session is ending.                 '
    '       vbAppTaskManager. Task Manager is closing your application.   '
    '       vbformMDIform. Form is an MDI Child, and MDI Parent is closing.
    ' A common use of the event is to prompt the user to save changes.    '
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Open "BidAnalysis.cfg" For Output As #1
    Print #1, Me.Left, Me.Top
    Close #1
End Sub


Private Sub lstAnalysis_DblClick()
    Dim i As Integer
    Dim r As String
    Dim id As String
    Dim f As String
    
    If lstAnalysis.ListIndex = 0 Then
        r = lstAnalysis
        i = InStr(r, "ID#")
        If i > 0 Then
            id = Trim$(Str(Val(Mid$(r, i + 5))))
            Open App.Path & "\ImportID.txt" For Output As #1
            Print #1, id
            Close #1
        End If
    End If
End Sub
