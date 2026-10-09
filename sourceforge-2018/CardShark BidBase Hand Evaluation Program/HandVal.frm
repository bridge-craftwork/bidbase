VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C2DB3F28-2DA2-41F7-8A58-A4485C3FDCB4}#1.2#0"; "gradframectl.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmHandVal 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hand Evaluation Systems - CardShark BidBase"
   ClientHeight    =   9795
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   14025
   Icon            =   "HandVal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   14025
   StartUpPosition =   3  'Windows Default
   Begin GradFrameCTL.GradFrame gf_CopyEntry 
      Height          =   3000
      Left            =   3360
      Top             =   3405
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   5292
      ShowOutline     =   0   'False
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Copy Entry"
      GradBlend0      =   16711680
      GradBlend2      =   16711680
      GradBlend1      =   255
      GradBlend3      =   8454143
      ForeColor       =   16777215
      Begin VB.CheckBox b_CopyCancel 
         BackColor       =   &H8000000A&
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
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   915
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   2220
         Width           =   1140
      End
      Begin VB.CheckBox b_CopyEntry 
         BackColor       =   &H8000000A&
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   930
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1635
         Width           =   1140
      End
      Begin VB.TextBox t_CopyID 
         Height          =   285
         Left            =   1275
         TabIndex        =   44
         Top             =   1050
         Width           =   1095
      End
      Begin VB.ComboBox cmb_CopyRow_System 
         BackColor       =   &H000000FF&
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
         Height          =   315
         Left            =   1305
         Sorted          =   -1  'True
         TabIndex        =   43
         Text            =   "1.All"
         Top             =   600
         Width           =   1340
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Entry ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   255
         Index           =   1
         Left            =   330
         TabIndex        =   80
         Top             =   1065
         Width           =   915
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "System:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   255
         Index           =   0
         Left            =   315
         TabIndex        =   79
         Top             =   615
         Width           =   915
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "HandVal.frx":08CA
      Height          =   7260
      Left            =   105
      OleObjectBlob   =   "HandVal.frx":08DE
      TabIndex        =   93
      Top             =   1680
      Width           =   11070
   End
   Begin GradFrameCTL.GradFrame GradFrame3 
      Height          =   720
      Left            =   105
      Top             =   9000
      Width           =   13860
      _ExtentX        =   24448
      _ExtentY        =   1270
      ShowOutline     =   0   'False
      BorderStyle     =   4
      FrameBorderStyle=   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradBlend0      =   16744576
      GradBlend2      =   8438015
      GradBlend1      =   16744576
      GradBlend3      =   33023
      GradAngle       =   60
      ForeColor       =   0
      Begin VB.Label L_Info 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   660
         Left            =   -15
         TabIndex        =   92
         Top             =   15
         Width           =   13860
      End
   End
   Begin VB.PictureBox p_All 
      BackColor       =   &H000000C0&
      Height          =   15
      Left            =   12120
      ScaleHeight     =   15
      ScaleWidth      =   960
      TabIndex        =   47
      Top             =   6390
      Width           =   960
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   3825
         Left            =   60
         TabIndex        =   48
         Top             =   60
         Visible         =   0   'False
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   6747
         _Version        =   393216
         Rows            =   17
         Cols            =   4
         ScrollBars      =   0
         BorderStyle     =   0
      End
   End
   Begin GradFrameCTL.GradFrame GradFrame2 
      Height          =   1905
      Index           =   3
      Left            =   12180
      Top             =   4395
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   3360
      ShowOutline     =   0   'False
      BorderStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Me:"
      GradBlend0      =   16777152
      GradBlend2      =   65535
      GradBlend1      =   255
      GradBlend3      =   33023
      ForeColor       =   0
      Begin VB.TextBox t_Len 
         Height          =   240
         Index           =   13
         Left            =   390
         TabIndex        =   37
         Top             =   1170
         Width           =   345
      End
      Begin VB.TextBox t_Len 
         Height          =   240
         Index           =   14
         Left            =   390
         TabIndex        =   36
         Top             =   945
         Width           =   345
      End
      Begin VB.TextBox t_Len 
         Height          =   240
         Index           =   15
         Left            =   390
         TabIndex        =   35
         Top             =   690
         Width           =   345
      End
      Begin VB.TextBox t_Len 
         Height          =   240
         Index           =   16
         Left            =   390
         TabIndex        =   34
         Top             =   435
         Width           =   345
      End
      Begin VB.CheckBox ck_NT 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   510
         TabIndex        =   38
         Top             =   1455
         Width           =   210
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "S: H: D: C:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   5
         Left            =   120
         TabIndex        =   91
         Top             =   450
         Width           =   240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "NT:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   90
         Top             =   1470
         Width           =   330
      End
   End
   Begin GradFrameCTL.GradFrame GradFrame2 
      Height          =   1905
      Index           =   2
      Left            =   13050
      Top             =   3390
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   3360
      ShowOutline     =   0   'False
      BorderStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "RHO:"
      GradBlend0      =   16777152
      GradBlend2      =   65535
      GradBlend1      =   33023
      GradBlend3      =   255
      ForeColor       =   0
      Begin VB.CheckBox ck_NT 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   495
         TabIndex        =   33
         Top             =   1575
         Width           =   210
      End
      Begin VB.TextBox t_Len 
         Height          =   240
         Index           =   12
         Left            =   375
         TabIndex        =   29
         Top             =   555
         Width           =   345
      End
      Begin VB.TextBox t_Len 
         Height          =   240
         Index           =   11
         Left            =   375
         TabIndex        =   30
         Top             =   810
         Width           =   345
      End
      Begin VB.TextBox t_Len 
         Height          =   240
         Index           =   10
         Left            =   375
         TabIndex        =   31
         Top             =   1065
         Width           =   345
      End
      Begin VB.TextBox t_Len 
         Height          =   240
         Index           =   9
         Left            =   375
         TabIndex        =   32
         Top             =   1305
         Width           =   345
      End
      Begin VB.TextBox t_Str 
         Height          =   240
         Index           =   2
         Left            =   375
         TabIndex        =   28
         Top             =   315
         Width           =   345
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "NT:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   135
         TabIndex        =   89
         Top             =   1590
         Width           =   330
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "S: H: D: C:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   3
         Left            =   120
         TabIndex        =   88
         Top             =   570
         Width           =   240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Str:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   87
         Top             =   330
         Width           =   330
      End
   End
   Begin GradFrameCTL.GradFrame GradFrame2 
      Height          =   1905
      Index           =   1
      Left            =   12180
      Top             =   2360
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   3360
      ShowOutline     =   0   'False
      BorderStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Pard:"
      GradBlend0      =   16777152
      GradBlend2      =   65535
      GradBlend1      =   33023
      GradBlend3      =   255
      ForeColor       =   0
      Begin VB.CheckBox ck_NT 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   510
         TabIndex        =   27
         Top             =   1560
         Width           =   210
      End
      Begin VB.TextBox t_Len 
         Height          =   240
         Index           =   8
         Left            =   390
         TabIndex        =   23
         Top             =   540
         Width           =   345
      End
      Begin VB.TextBox t_Len 
         Height          =   240
         Index           =   7
         Left            =   390
         TabIndex        =   24
         Top             =   795
         Width           =   345
      End
      Begin VB.TextBox t_Len 
         Height          =   240
         Index           =   6
         Left            =   390
         TabIndex        =   25
         Top             =   1050
         Width           =   345
      End
      Begin VB.TextBox t_Len 
         Height          =   240
         Index           =   5
         Left            =   390
         TabIndex        =   26
         Top             =   1290
         Width           =   345
      End
      Begin VB.TextBox t_Str 
         Height          =   240
         Index           =   1
         Left            =   390
         TabIndex        =   22
         Top             =   300
         Width           =   345
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Str:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   60
         TabIndex        =   86
         Top             =   330
         Width           =   330
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "S: H: D: C:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   13
         Left            =   120
         TabIndex        =   85
         Top             =   570
         Width           =   240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "NT:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   135
         TabIndex        =   84
         Top             =   1590
         Width           =   330
      End
   End
   Begin GradFrameCTL.GradFrame GradFrame2 
      Height          =   1905
      Index           =   0
      Left            =   11310
      Top             =   3390
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   3360
      ShowOutline     =   0   'False
      BorderStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "LHO:"
      GradBlend0      =   16777088
      GradBlend2      =   65535
      GradBlend1      =   255
      GradBlend3      =   33023
      ForeColor       =   0
      Begin VB.CheckBox ck_NT 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   495
         TabIndex        =   21
         Top             =   1575
         Width           =   210
      End
      Begin VB.TextBox t_Str 
         Height          =   240
         Index           =   0
         Left            =   375
         TabIndex        =   16
         Top             =   315
         Width           =   345
      End
      Begin VB.TextBox t_Len 
         Height          =   240
         Index           =   1
         Left            =   360
         TabIndex        =   20
         Top             =   1305
         Width           =   345
      End
      Begin VB.TextBox t_Len 
         Height          =   240
         Index           =   2
         Left            =   375
         TabIndex        =   19
         Top             =   1065
         Width           =   345
      End
      Begin VB.TextBox t_Len 
         Height          =   240
         Index           =   3
         Left            =   375
         TabIndex        =   18
         Top             =   810
         Width           =   345
      End
      Begin VB.TextBox t_Len 
         Height          =   240
         Index           =   4
         Left            =   375
         TabIndex        =   17
         Top             =   555
         Width           =   345
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "NT:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   83
         Top             =   1590
         Width           =   330
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "S: H: D: C:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   1
         Left            =   120
         TabIndex        =   82
         Top             =   570
         Width           =   240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Str:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   81
         Top             =   330
         Width           =   330
      End
   End
   Begin VB.TextBox t_FindHand 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   11685
      MaxLength       =   16
      TabIndex        =   39
      Top             =   6390
      Width           =   1785
   End
   Begin VB.ComboBox cmb_ValSystem 
      BackColor       =   &H000000FF&
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
      Height          =   315
      Left            =   6390
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "1.All"
      Top             =   105
      Width           =   1340
   End
   Begin VB.TextBox t_1Maj 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   1485
      TabIndex        =   1
      Top             =   675
      Width           =   555
   End
   Begin VB.TextBox t_1Min 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   1470
      TabIndex        =   2
      Top             =   990
      Width           =   555
   End
   Begin VB.TextBox t_1over1 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   3780
      TabIndex        =   3
      Top             =   675
      Width           =   555
   End
   Begin VB.TextBox t_2over1 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   3765
      TabIndex        =   4
      Top             =   990
      Width           =   555
   End
   Begin VB.TextBox t_2C_Maj 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   6135
      TabIndex        =   5
      Top             =   675
      Width           =   555
   End
   Begin VB.TextBox t_2C_Min 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   6150
      TabIndex        =   6
      Top             =   990
      Width           =   555
   End
   Begin VB.TextBox t_MajGame 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   8640
      TabIndex        =   7
      Top             =   675
      Width           =   555
   End
   Begin VB.TextBox t_MinGame 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   8640
      TabIndex        =   8
      Top             =   990
      Width           =   555
   End
   Begin VB.TextBox t_SmallSlam 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   10785
      TabIndex        =   9
      Top             =   675
      Width           =   555
   End
   Begin VB.TextBox t_GrandSlam 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   10785
      TabIndex        =   10
      Top             =   1005
      Width           =   555
   End
   Begin VB.TextBox t_NT 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   0
      Left            =   13140
      TabIndex        =   11
      Top             =   120
      Width           =   780
   End
   Begin VB.TextBox t_NT 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   1
      Left            =   13140
      TabIndex        =   12
      Top             =   375
      Width           =   780
   End
   Begin VB.TextBox t_NT 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   2
      Left            =   13140
      TabIndex        =   13
      Top             =   645
      Width           =   780
   End
   Begin VB.TextBox t_NT 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   3
      Left            =   13140
      TabIndex        =   14
      Top             =   915
      Width           =   780
   End
   Begin VB.TextBox t_NT 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   4
      Left            =   13140
      TabIndex        =   15
      Top             =   1185
      Width           =   780
   End
   Begin VB.CommandButton b_Compute 
      Caption         =   "&Compute Value"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11730
      TabIndex        =   42
      Top             =   7950
      Width           =   1785
   End
   Begin VB.TextBox t_Dealer 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   12900
      TabIndex        =   40
      Text            =   "S"
      Top             =   7080
      Width           =   315
   End
   Begin VB.TextBox t_Bidding 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11415
      TabIndex        =   41
      Text            =   "Ex:  1C-1H-1S-1N-P-D-R"
      Top             =   7545
      Width           =   2355
   End
   Begin VB.Timer Timer1 
      Left            =   -75
      Top             =   -105
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "BB-HEP.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   315
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "HandVal"
      Top             =   -60
      Visible         =   0   'False
      Width           =   1725
   End
   Begin GradFrameCTL.GradFrame GradFrame1 
      Height          =   9900
      Left            =   -30
      Top             =   0
      Width           =   14100
      _ExtentX        =   24871
      _ExtentY        =   17463
      ShowOutline     =   0   'False
      BorderStyle     =   6
      FrameBorderStyle=   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradBlend0      =   65535
      GradBlend2      =   16711680
      GradBlend1      =   255
      GradBlend3      =   8454143
      GradAngle       =   160
      ForeColor       =   0
      Begin VB.PictureBox Picture1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   11690
         ScaleHeight     =   375
         ScaleWidth      =   1845
         TabIndex        =   49
         Top             =   6345
         Width           =   1850
      End
      Begin VB.Label L_SuitVal 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   12975
         TabIndex        =   78
         Top             =   8325
         Width           =   570
      End
      Begin VB.Label L_NTVal 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   12975
         TabIndex        =   77
         Top             =   8610
         Width           =   570
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "The Game and Slam points below are the actual  points in the current hand plus the minimum expected points from partner."
         Height          =   615
         Index           =   11
         Left            =   8400
         TabIndex        =   76
         Top             =   60
         Width           =   3015
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         BackColor       =   &H000000FF&
         Height          =   375
         Index           =   0
         Left            =   6385
         TabIndex        =   75
         Top             =   75
         Width           =   1395
      End
      Begin VB.Label Label3 
         BackColor       =   &H000000FF&
         Height          =   360
         Index           =   1
         Left            =   11730
         TabIndex        =   74
         Top             =   7920
         Width           =   1845
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bidding:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   26
         Left            =   12195
         TabIndex        =   73
         Top             =   7350
         Width           =   780
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   25
         Left            =   12150
         TabIndex        =   72
         Top             =   7095
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "What The Bidding Has Shown"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   24
         Left            =   11715
         TabIndex        =   71
         Top             =   1770
         Width           =   1875
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         X1              =   90
         X2              =   13910
         Y1              =   1545
         Y2              =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NT Valuation:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   22
         Left            =   11670
         TabIndex        =   70
         Top             =   8655
         Width           =   1260
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Suit Valuation:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   21
         Left            =   11730
         TabIndex        =   69
         Top             =   8385
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "My Hand"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   20
         Left            =   11985
         TabIndex        =   68
         Top             =   6825
         Width           =   1305
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum Bidding Requirements:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   18
         Left            =   300
         TabIndex        =   67
         Top             =   285
         Width           =   3945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "2C then 3N:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   17
         Left            =   12015
         TabIndex        =   66
         Top             =   1185
         Width           =   1065
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "2C then 2N:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   16
         Left            =   12015
         TabIndex        =   65
         Top             =   915
         Width           =   1065
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "3N:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   12750
         TabIndex        =   64
         Top             =   675
         Width           =   330
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "2N:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   14
         Left            =   12750
         TabIndex        =   63
         Top             =   390
         Width           =   330
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1N:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   12750
         TabIndex        =   62
         Top             =   150
         Width           =   330
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Notrump Ranges: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   11745
         TabIndex        =   61
         Top             =   165
         Width           =   825
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Grand Slam:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   9750
         TabIndex        =   60
         Top             =   1020
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Small Slam:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   9735
         TabIndex        =   59
         Top             =   705
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Minor Game: Explore Slam"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   8
         Left            =   7530
         TabIndex        =   58
         Top             =   1005
         Width           =   1215
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Major Game:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   7530
         TabIndex        =   57
         Top             =   690
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "2/1 Resp. (NGF):   or Limit Raise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   6
         Left            =   2235
         TabIndex        =   56
         Top             =   1005
         Width           =   1515
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "1/1 Response:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   2250
         TabIndex        =   55
         Top             =   690
         Width           =   1395
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "2C Strong, Minor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   4560
         TabIndex        =   54
         Top             =   1005
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "2C Strong, Major:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   4575
         TabIndex        =   53
         Top             =   690
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "1C/1D Open:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   285
         TabIndex        =   52
         Top             =   1005
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "1H/1S Open:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   315
         TabIndex        =   51
         Top             =   690
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "System Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   4830
         TabIndex        =   50
         Top             =   150
         Width           =   1515
      End
   End
   Begin VB.Menu zF 
      Caption         =   "&File"
      Begin VB.Menu mNewSystem 
         Caption         =   "&New system"
      End
      Begin VB.Menu mCopySystem 
         Caption         =   "&Copy current system"
      End
      Begin VB.Menu mDeleteSystem 
         Caption         =   "&Delete current system"
      End
      Begin VB.Menu zBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mSaveMinBidReq 
         Caption         =   "&Save Min.Bid Req."
      End
      Begin VB.Menu zBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu zE 
      Caption         =   "&Edit"
      Begin VB.Menu mCopyRow 
         Caption         =   "&Copy grid row"
         Shortcut        =   ^K
      End
      Begin VB.Menu mDeleteRow 
         Caption         =   "&Delete grid row"
         Shortcut        =   ^D
      End
      Begin VB.Menu zBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mReSortGrid 
         Caption         =   "&Re-sort grid by ID"
         Shortcut        =   ^R
      End
      Begin VB.Menu zBa3a 
         Caption         =   "-"
      End
      Begin VB.Menu mWow 
         Caption         =   "Whirl"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu zH 
      Caption         =   "&Help"
      Begin VB.Menu mHelp 
         Caption         =   "&Program Help"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mHelp 
         Caption         =   "System &Info"
         Index           =   1
         Shortcut        =   ^I
      End
      Begin VB.Menu zBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmHandVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************'
' CardShark BidBase Hand Evaluation Program '
'                                           '
' Purpose: Goren started with the 4-3-2-1   '
'          method of evaluating hands, but  '
'          others have improved on that.    '
'          For BidBase we have combined the '
'          best of those systems for the    '
'          most accurate hand evaluation.   '
'          (Well, in my mind, anyway.)      '
'                                           '
' Project - References:                     '
'        Microsoft Data Binding Collection  '
'        Microsoft DAO 3.51 Object Library  '
' Project - Components:                     '
'        MS Data Bound Grid Control 5.0     '
'        MS DataGrid Control 6.0 (OLEDB)    '
'        MS FlexGrid Control 6.0            '
'        VBSight GradFrameCTL               '
'                                           '
' Don't ask me why there are so many grid   '
' controls. This program code dates back    '
' to the turn of the century (19th?).       '
' That "Whirl" feature must have been put   '
' in to break the monotony, but it might    '
' be cute in a game when someone does good. '
'*******************************************'

Option Explicit
Option Compare Binary ' In string comparisons, "A" <> "a"

'Database:
Private db_HandVal As Database
Private ws_HandVal As Workspace
Private rs_HandVal As Recordset
Private rs_HandValBidReq As Recordset
Private bln_AlreadyInitialized As Boolean
Private bln_NeedsClearing As Boolean
Private lng_GradBlend0 As Long
Private str_LastSystem As String

'Client input variables:
Private bln_Cancel As Integer   ' Emergency exit from this module
Private str_PlayersHand As String
Private str_Vul As String
Private str_OrigData As String
'Card naming constants:
Private Const Ace = 14
Private Const King = 13
Private Const Queen = 12
Private Const Jack = 11
Private Const Ten = 10
Private Const Nine = 9
Private Const Eight = 8
'Suit naming constants:
Private Const Clbs = 1
Private Const Dias = 2
Private Const Hrts = 3
Private Const Spds = 4
' Player's holdings: '
Private bln_AlreadyHaveOPsData ' Flag for other player's data retrieval
    ' Fits: '
Private bln_Fit_Established As Boolean ' Fit is known to bidder.
Private bln_Fit_Confirmed As Boolean  ' Fit has been shown by bidder.
Private int_Fit_Suit(3) As Integer
Private int_Fit_Len(3) As Integer
Private int_Opps_FitSuit(2) As Integer
    ' My hand: '
Private bln_I_Bid_NT As Boolean
Private int_My_CueBid As Integer
Private int_My_Fit(2) As Integer ' Suit(s) in which I raised partner
Private int_My_HelpSuit As Integer
Private int_My_Pos As Integer
Private int_My_RebidSuit(2) As Integer
Private int_My_SS(3) As Integer
Private int_My_Suit(4) As Integer
    ' Pard: '
Private bln_Pard_Bid_NT As Boolean
Private int_Pards_CueBid As Integer ' Pard has shown A/K in suit 1-4 (C-S)
Private int_Pards_Fit(2) As Integer
Private int_Pards_HelpSuit As Integer
Private int_Pards_Pos As Integer ' Partner's position
Private int_Pards_RebidSuit(2) As Integer
Private int_Pards_SS As Integer ' Pard's short suit number
Private int_Pards_Suit(4) As Integer ' C=1 ... S=4
Private int_Pards_SuitLen(4) As Integer  ' The expected suit len.        '
Private str_Pards_Strength As String
    ' LHO: '
Private bln_LHO_Bid_NT As Boolean
Private int_LHOs_CueBid As Integer
Private int_LHOs_Fit(2) As Integer
Private int_LHOs_HelpSuit As Integer
Private int_LHOs_Pos As Integer
Private int_LHOs_RebidSuit(2) As Integer
Private int_LHOs_SS As Integer
Private int_LHOs_Suit(2) As Integer
Private int_LHOs_SuitLen(4) As Integer
Private str_LHOs_Strength As String
    ' RHO: '
Private bln_RHO_Bid_NT As Boolean
Private int_RHOs_CueBid As Integer
Private int_RHOs_Fit(2) As Integer
Private int_RHOs_HelpSuit As Integer
Private int_RHOs_Pos As Integer
Private int_RHOs_RebidSuit(2) As Integer
Private int_RHOs_SS As Integer ' short suit
Private int_RHOs_Suit(2) As Integer
Private int_RHOs_SuitLen(4) As Integer
Private str_RHOs_Strength As String
' Hand analysis variables: '
' (In alphabetical order,  '
'  ignoring prefixes.)     '
Private int_BidNum As Integer   ' Current bid #, starting with first non-pass
Private int_BidRound As Integer ' Bidding round, starting with 1st seat
Private str_Bid(50) As String
Private str_Call(50) As String
Private int_CallNum As Integer  ' Current call #, starting with 1st seat
Private str_Disclosure As String  ' Give client all possible bid meanings
Private int_DistribPts(4) As Integer  ' Distribution points by suit/hand
Private int_HCP(4) As Integer   ' High Card Points by suit/hand
Private lng_ID As Long  'ID# of entry matching hand's specs
Private int_NumLosers(4) As Integer ' Losing Trick Count by suit/hand
Private sng_NumRunners(4) As Single ' # of running winners
Private sng_NumStops(4) As Single ' # of stoppers by suit, # of suits stopped
Private sng_NumWinners(4) As Single ' # of winners by suit(max=4)/hand
Private str_Position As String  ' Calculated position codes
Private int_PosNum As Integer   ' Position # not incl. opening passes.
Private str_PriorBids As String ' Prior bids/calls from client
Private sng_QuickTricks(4) As Single ' # of Quick Tricks by suit/hand
Private str_Shape As String     ' Shape code for this hand
Private str_Shortness As String ' Code if singleton, void, or both
Private str_Suit(4) As String   ' Card string by suit (e.g.: AJ732)
Private bln_Suit(4, 14) As Boolean ' Cards held (T/F) array for hand
Private int_SuitLen(4) As Integer  ' # of cards in each suit
Private str_SuitNames(4) As String ' Clubs, Diamonds, Hearts, Spades
' BidBase Hand Valuation variables:                   '
    ' Specs: '
Private str_CardsSpec As String
Private str_CardsSpecCombo As String
Private str_CardsSpecExtracted As String
Private str_CardsSpecOperator As String
Private str_FitSpec As String
Private str_HandShape As String
Private str_HandSpec As String
Private int_LenSpec As Integer
Private str_LenSpec As String
Private str_OPsSpec As String
Private str_PosSpec As String
Private int_SuitSpec As Integer
Private str_SuitSpec As String
Private sng_ValSpec As Single
Private str_ValSpec As String
    ' Flags: '
Private bln_CountFlag As Boolean
Private str_Flag(20) As String
Private int_Flag(20) As Integer
Private int_FlagNum As Integer
Private str_FlagSpec As String
Private int_NumFlags As Integer
    ' Hand Value: '
Private sng_HandVal As Single
Private sng_HandNT_Val As Single
Private sng_HandSuit_Val As Single
Private int_HCPn(4) As Integer
Private int_Intermediates(4) As Integer
Private int_Shpn(4) As Integer
Private int_SuitNum As Integer
Private bln_ValAdded As Boolean
Private str_ValSystem(20) As String ' Names of hand valuation systems.'
Private sng_ValPoints(20, 2) As Single ' 1=suit, 2=NT '
' Grid Columns: '
Private int_Col_VP As Integer
Private int_Col_Len As Integer
Private int_Col_Cards As Integer
Private int_Col_Suit As Integer
Private int_Col_Fit As Integer
Private int_Col_Hands As Integer
Private int_Col_Pos As Integer
Private int_Col_Flag As Integer
Private int_Col_Value As Integer
Private int_Col_NTSuit As Integer

Private Function LowWord(dwValue As Long) As Integer
   If dwValue < &H8000& Then
      LowWord = CInt(dwValue)
   Else
      LowWord = CInt(dwValue - &H10000)
   End If
End Function


Private Sub b_Compute_Click()
    Dim s As String
    Dim i As Integer
    Dim n As Single
    Dim a As String

    If b_Compute.Caption = "&Close Results" Then
        LowerWindow
        b_Compute.Caption = "&Compute Value"
        L_Info = ""
        t_FindHand.SetFocus
        Exit Sub
    End If

    str_PlayersHand = Trim$(t_FindHand)
    If Len(str_PlayersHand) < 16 Then
        MsgBox "Hand must have 13 cards" & vbCr & _
               "with 3 dashes between suits."
        t_FindHand.SetFocus
        Exit Sub
    End If
    s = Trim$(cmb_ValSystem)
    If s = "" Then
        MsgBox "Select a system first."
        cmb_ValSystem.SetFocus
        Exit Sub
    End If

    ScrollToGridTop
    AnalyzeHand
    If cmb_ValSystem = "1.All" Then
        RaiseWindow
        For i = 1 To cmb_ValSystem.ListCount - 1
            cmb_ValSystem.ListIndex = i
            DoEvents
            DBGrid1.Refresh
            Eval_A_Hand cmb_ValSystem
            Grid1.col = 0
            Grid1.row = i
            Grid1.Text = cmb_ValSystem.Text
            Grid1.col = 1
            n = Int((sng_HandVal + sng_HandSuit_Val) * 100) / 100
            Grid1.Text = "" & n
            Grid1.col = 2
            If Val(t_1Maj) = 0 Then
                n = n / 13
            Else
                n = n / Val(t_1Maj)
            End If
            n = Int(n * 100)
            Grid1.Text = n & "%"
            Grid1.col = 3
            Grid1.Text = "" & Int((sng_HandVal + sng_HandNT_Val) * 100) / 100
        Next
        cmb_ValSystem.ListIndex = 0
        b_Compute.Caption = "&Close Results"
    Else
        ClearVP
        DBGrid1.Refresh
        Eval_A_Hand cmb_ValSystem
        If Len("" & sng_HandVal) > 5 Then
            sng_HandVal = CInt(sng_HandVal * 100) / 100
        End If
        L_SuitVal = "" & (sng_HandVal + sng_HandSuit_Val)
        L_NTVal = "" & (sng_HandVal + sng_HandNT_Val)
        ScrollToGridTop
    End If
    bln_NeedsClearing = True
    If cmb_ValSystem = "1.All" Then
        L_Info.Top = 0
        L_Info = vbCr & "%/1M  is the Suit VPs as a % of the minimum VP required to open 1H/1S."
        b_Compute.SetFocus
    Else
        t_FindHand.SetFocus
    End If
End Sub


Private Sub RaiseWindow()
    Dim i As Integer

    With p_All
        .ZOrder 0
        Grid1.Visible = False
        Do
            .Top = .Top - 2
            .Height = .Height + 2
            .Refresh
        Loop While .Top > 2360
        .Top = 2360
        .Height = 4000

        Do
            .Left = .Left - 8.35
            .Width = .Width + 15.7
            .Refresh
        Loop While .Left > 11300
        .Left = 11300
        .Width = 2595

        Grid1.Visible = True
        With Grid1
            .Clear
            .row = 0
            .col = 1
            .Text = " Suit"
            .col = 2
            .Text = "%/1M"
            .col = 3
            .Text = "  NT"
        End With
        Grid1.Refresh
    End With
    Label1(24).Caption = "Evaluation Results"
End Sub


Private Sub LowerWindow()
    Dim i As Integer
    Dim j As Integer

    With p_All
        Grid1.Visible = False
        For i = 1 To 100
            .Width = .Width - 15.7
            .Left = .Left + 8.35
            .Refresh
            If i / 20 = i \ 20 Then
                GradFrame2(0).Refresh
                GradFrame2(2).Refresh
            End If
        Next
        t_Str(0).Refresh
        For i = 1 To 4
            t_Len(i).Refresh
        Next
        ck_NT(0).Refresh
        For j = 0 To 3
            GradFrame2(j).Refresh
        Next
        For i = 1 To 1000
            .Height = .Height - 4
            .Top = .Top + 4
            .Refresh
            If i / 200 = i \ 200 Then
                For j = 0 To 3
                    GradFrame2(j).Refresh
                Next
            End If
        Next
        .Top = 6390
        .Height = 0
        .ZOrder 1
    End With
    Label1(24).Caption = "What The Bidding Has Shown"
    frmHandVal.Refresh
End Sub


Private Sub b_CopyCancel_Click()
    gf_CopyEntry.Visible = False
End Sub


Private Sub b_CopyEntry_Click()
    Dim id As String
    Dim id2 As String
    Dim n As String
    Dim n2 As String
    Dim i As Integer
    Dim j As Integer

    id = DBGrid1.Columns(0).Text
    n = DBGrid1.Columns(1).Text
    id2 = Trim$(t_CopyID)
    n2 = Trim$(cmb_CopyRow_System)

    If id2 = "" Or n = "" Then
        MsgBox "Select a system and enter a new ID."
        Exit Sub
    End If

    If id = id2 Then
        MsgBox "The new ID is the same" & vbCr & _
               "as the existing one."
        Exit Sub
    ElseIf n = n2 Then
        i = MsgBox("Did you mean to enter the same" & vbCr & _
                   "system name for the new entry?.", vbYesNo)
        If i = vbNo Then Exit Sub
    End If

    j = 1
    Do While cmb_ValSystem.List(j) <> n2
        j = j + 1
    Loop
    If cmb_ValSystem.List(j) <> n2 Then
        MsgBox "The name " & n2 & " is not in the database." & _
                vbCr & "Create a new system first."
        Exit Sub
    End If

    With rs_HandVal
    .Index = "PrimaryKey"
    .Seek "=", id

    If .NoMatch = False Then
        .Edit
        !id = id2
        !Name = n2
        .Update
        If n = n2 Then ' Same name. '
            mReSortGrid_Click
            MsgBox "Entry copied."
        ElseIf cmb_ValSystem <> "1.All" Then
            i = MsgBox("Entry copied." & vbCr & _
                       "Load " & n2 & "?", vbYesNo)
            If i = vbYes Then
                cmb_ValSystem.ListIndex = j
            End If
        End If
        gf_CopyEntry.Visible = False
    Else
        MsgBox id & " is not in the database."
    End If
    End With
End Sub


Private Sub cmb_ValSystem_GotFocus()
    L_Info.Top = 50
    L_Info = vbCr & "Select a system or compare All."
End Sub


Private Sub cmb_ValSystem_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And (KeyCode = "38" Or KeyCode = "40") Then
        ClearVP
    End If
End Sub


Private Sub DBGrid1_Change()
    Dim n As Integer

    With DBGrid1
    ' 0  1  2 3 4 5 6  7 8 10 1 2 4
    '10 10 50 1 7 5 7 10 2 5 20 5 5
    Select Case .col
    Case Is = 0, 1, 7
        n = 10
    Case Is = 2
        n = 50
    Case Is = 3
        n = 1
    Case Is = 4, 6
        n = 7
    Case Is = 5, 10, 12, 14
        n = 5
    Case Is = 8
        n = 2
    Case Is = 11
        n = 20
    End Select

    If Len(.Text) > n Then
        MsgBox "Maximum field length is " & n & _
               "." & vbCr & _
               "Data entered was: " & .Text & _
               " (length=" & Len(.Text) & ")"
        .Text = Left$(.Text, n)
    End If
    End With
End Sub


Private Sub mAbout_Click()
    MsgBox "   CardShark BidBase" & vbCr & _
           "Hand Evaluation Program" & vbCr & _
           "      Version 1.0" & vbCr & vbCr & _
           "www.CardSharkGames.com"
End Sub


Private Sub mCopyRow_Click()
    gf_CopyEntry.Visible = True
End Sub


Private Sub mCopySystem_Click()
    Dim n As String
    Dim m As String
    Dim i As Integer

    n = InputBox$("Enter new name:", "Create New System")
    If n = "" Then Exit Sub
    m = Left$(n, 1)
    cmb_ValSystem.AddItem n
    cmb_ValSystem.Text = n
    cmb_CopyRow_System.AddItem n
    With rs_HandValBidReq
        .AddNew
        !EvalName = n
        !Maj_1Opening = t_1Maj
        !Min_1Opening = t_1Min
        !Maj_2C = t_2C_Maj
        !Min_2C = t_2C_Min
        !OneOver1Resp = t_1over1
        !TwoOver1Resp = t_2over1
        !MajGame = t_MajGame
        !MinGame = t_MinGame
        !SmallSlam = t_SmallSlam
        !GrandSlam = t_GrandSlam
        !Str_1N_Range = t_NT(0)
        !Bal_2N_Range = t_NT(1)
        !Bal_3N_Range = t_NT(2)
        !Bal_2C_2N_Range = t_NT(3)
        !Bal_2C_3N_Range = t_NT(4)
        .Update
    End With

    With rs_HandVal
        DBGrid1.row = 0
        Do
            .AddNew
            DBGrid1.col = 0
            !id = m & Mid$(DBGrid1.Text, 2)
            !Name = n
            DBGrid1.col = 2
            !Description = DBGrid1.Text
            DBGrid1.col = int_Col_NTSuit
            !NT_or_Suit = DBGrid1.Text
            DBGrid1.col = int_Col_Value
            !Value = DBGrid1.Text
            DBGrid1.col = int_Col_Flag
            !Flag = DBGrid1.Text
            DBGrid1.col = int_Col_Pos
            !Position = DBGrid1.Text
            DBGrid1.col = int_Col_Hands
            !Hands = DBGrid1.Text
            DBGrid1.col = int_Col_Fit
            !Fit = DBGrid1.Text
            DBGrid1.col = int_Col_Suit
            !Suit = DBGrid1.Text
            DBGrid1.col = int_Col_Cards
            !Cards = DBGrid1.Text
            DBGrid1.col = int_Col_Len
            !Len = DBGrid1.Text
            .Update
            DBGrid1.row = DBGrid1.row + 1
            DBGrid1.col = 0
        Loop While DBGrid1.Text <> ""
    End With
    i = 2
    Do While cmb_ValSystem.List(i) <> n
        i = i + 1
    Loop
    cmb_ValSystem.ListIndex = i
End Sub


Private Sub mDeleteRow_Click()
    Dim i As Integer
    Dim id As String
    Dim id2 As String

    id = DBGrid1.Columns(0).Text
    i = MsgBox("Delete entry " & id & "?", vbYesNo)
    If i = vbYes Then
        With DBGrid1
        If .row > 1 Then
            .row = .row - 1
        Else
            .row = .row + 1
        End If
        id2 = .Text
        End With
        With rs_HandVal
        .Index = "PrimaryKey"
        .Seek "=", id
        If .NoMatch = False Then
            .Delete
            With DBGrid1
            mReSortGrid_Click
            Do While .Text <> id2 And .Columns(1) <> ""
                .row = .row + 1
            Loop
            End With
        End If
        End With
    End If
End Sub

Private Sub mDeleteSystem_Click()
    Dim n As String
    Dim i As Integer

    n = cmb_ValSystem
    If n = "" Then
        MsgBox "Select a system first."
        Exit Sub
    End If

    i = MsgBox("Delete the system " & n & "?", vbYesNo)
    If i = vbNo Then Exit Sub

    With rs_HandVal
        .Index = "Name"
        .Seek "=", n
         If .NoMatch = False Then
            Do
                If !Name = n Then .Delete
                .MoveNext
                If .EOF Then Exit Do
            Loop While !Name = n
        End If
    End With

    With rs_HandValBidReq
        .Index = "PrimaryKey"
        .Seek "=", n
        If Not .NoMatch Then
            .Delete
        End If
    End With

    i = 2
    Do While cmb_ValSystem.List(i) <> n
        i = i + 1
    Loop
    If cmb_ValSystem.List(i) = n Then
        cmb_ValSystem.RemoveItem i
    End If
    cmb_ValSystem.ListIndex = 0
    MsgBox n & " has been removed."
End Sub


Private Sub mExit_Click()
    Unload Me
End Sub

Private Sub mHelp_Click(Index As Integer)
    Dim xx As Long
    Dim f As String
    Dim e As String

    If Index = 0 Then
        xx = Shell("BidBase FileViewer.exe " & App.Path & _
                   "\Notes\CardShark_BidBase_Hand_Evaluation.htm", _
                   vbNormalFocus)
    Else
        If cmb_ValSystem = "1.All" Then
            MsgBox "Select a system first.", , "System Info"
            cmb_ValSystem.SetFocus
            Exit Sub
        End If

        f = App.Path & "\Notes\" & cmb_ValSystem & "_Evaluation.htm"
        If Dir$(f) <> "" Then
            e = DBGrid1.Columns(0).Text
            xx = Shell("BidBase FileViewer.exe " & f & "#" & e, _
                   vbNormalFocus)
        Else
            xx = MsgBox("No info file found for " & cmb_ValSystem & vbCr & _
                        "Would you like to create one?", vbYesNo)
            If xx = vbYes Then
                Open f For Output As #1
                Print #1, "<HTML><HEAD>"
                Print #1, "<TITLE>CardShark BidBase - " & cmb_ValSystem & "</TITLE>"
                Print #1, "</HEAD>"
                Print #1, "<BODY TEXT=" & Chr$(34) & "#000080" & Chr$(34)
                Print #1, "BGCOLOR = " & Chr$(34) & "#AFFEFE" & Chr$(34)
                Print #1, "LINK = " & Chr$(34) & "#1C8E1F" & Chr$(34)
                Print #1, "VLINK = " & Chr$(34) & "#7C2C76" & Chr$(34)
                Print #1, "ALINK = " & Chr$(34) & "#F00000" & Chr$(34) & ">"
                Print #1, "<CENTER>"
                Print #1, "<TABLE WIDTH=90%><TR><TD>"
                Print #1, ""
                Print #1, "<H1 ALIGN=CENTER>" & cmb_ValSystem & "</H1>"
                Print #1, ""
                Print #1, ""
                Print #1, ""
                Print #1, ""
                Print #1, ""
                Print #1, ""
                Print #1, ""
                Print #1, "</TD></TR></TABLE>"
                Print #1, "<P></BODY></HTML>"
                Close #1

                xx = Shell("BidBase FileViewer.exe " & f & " " & f, vbNormalFocus)
            End If
        End If
    End If
End Sub


Private Sub mNewSystem_Click()
    Dim m As String
    Dim n As String
    Dim i As Integer
    
    DBGrid1.row = 0
    m = DBGrid1.Columns(1).Text
    n = InputBox$("Enter system name:", "Create New System")
    If n = "" Then
        cmb_ValSystem = m
        Exit Sub
    End If
    
    With rs_HandVal
        m = Left$(n, 1)
        .Index = "PrimaryKey"
        .Seek ">=", m
        If Not .NoMatch Then
            If Left$(!id, 1) = m And Val(Mid$(!id, 2)) > 0 Then
                ' e.g.: T1020, not TS1020 '
                m = InputBox$("The first letter of " & n & vbCr & _
                              "is already being used as an ID prefix." & vbCr & _
                              "Enter a 2-letter code to use instead:")
                If m = "" Then
                    MsgBox "New system not created."
                    Exit Sub
                End If
                Do
                    .Seek ">=", m
                    If Not .NoMatch Then
                        If Left$(!id, 2) = m And Val(Mid$(!id, 3)) > 0 Then
                            m = InputBox$("The letters '" & m & _
                                          "' are already" & vbCr & _
                                          "being used as an ID prefix." & vbCr & _
                                          "Enter different letters to use instead:")
                            If m = "" Then
                                MsgBox "New system not created."
                                Exit Sub
                            End If
                        Else
                            Exit Do
                        End If
                    Else
                        Exit Do
                    End If
                Loop
            End If
        End If
    End With
        
    cmb_ValSystem.AddItem n
    cmb_ValSystem.Text = n
    cmb_CopyRow_System.AddItem n

    ClearVP
    With rs_HandValBidReq
        .AddNew
        !EvalName = n
        .Update
    End With
    With rs_HandVal
        .AddNew
        !Name = n
        !id = m & "1000"
        .Update
    End With
    i = 1
    Do While cmb_ValSystem.List(i) <> n
        i = i + 1
    Loop
    cmb_ValSystem.ListIndex = 3
    cmb_ValSystem.ListIndex = i
End Sub


Private Sub mReSortGrid_Click()
    Dim n As String
    Dim w As String

    n = cmb_ValSystem.Text
    If n = "1.All" Then
        w = ""
    Else
        w = "WHERE Name='" & cmb_ValSystem.Text & "' "
    End If

    Data1.RecordSource = "SELECT * FROM HandVal " & _
                         w & _
                         "ORDER By HandVal.ID;"
    Data1.Refresh
End Sub


Private Sub mSaveMinBidReq_Click()
    With rs_HandValBidReq
        If .EOF Or .NoMatch Then Exit Sub
        .Edit
        !EvalName = cmb_ValSystem
        !Maj_1Opening = t_1Maj
        !Min_1Opening = t_1Min
        !Maj_2C = t_2C_Maj
        !Min_2C = t_2C_Min
        !OneOver1Resp = t_1over1
        !TwoOver1Resp = t_2over1
        !MajGame = t_MajGame
        !MinGame = t_MinGame
        !SmallSlam = t_SmallSlam
        !GrandSlam = t_GrandSlam
        !Str_1N_Range = t_NT(0)
        !Bal_2N_Range = t_NT(1)
        !Bal_3N_Range = t_NT(2)
        !Bal_2C_2N_Range = t_NT(3)
        !Bal_2C_3N_Range = t_NT(4)
        .Update
    End With
    MsgBox "Bidding requirements saved."
End Sub


Private Sub cmb_ValSystem_Click()
    Dim n As String
    Dim m As String
    Dim new1 As Boolean
    Dim i As Integer

    ClearVP
    str_LastSystem = cmb_ValSystem

    If cmb_ValSystem = "" Or bln_Cancel Then
        Exit Sub
    End If

    L_SuitVal = ""
    L_NTVal = ""

    If cmb_ValSystem = "1.All" Then
        ClearBoxes
        Data1.RecordSource = "SELECT * FROM HandVal " & _
                             "ORDER By HandVal.ID;"
        Data1.Refresh
        DBGrid1.BackColor = "&H" & Hex(255) & Hex(255) & Hex(255)
        Exit Sub
    End If

    Select Case cmb_ValSystem.ListIndex
    Case Is = 1, 10
        DBGrid1.BackColor = "&H" & Hex(200) & Hex(225) & Hex(245)
    Case Is = 2, 11
        DBGrid1.BackColor = "&H" & Hex(200) & Hex(245) & Hex(245)
    Case Is = 3, 12
        DBGrid1.BackColor = "&H" & Hex(225) & Hex(245) & Hex(200)
    Case Is = 4, 13
        DBGrid1.BackColor = "&H" & Hex(200) & Hex(245) & Hex(200)
    Case Is = 5, 14
        DBGrid1.BackColor = "&H" & Hex(200) & Hex(225) & Hex(245)
    Case Is = 6, 15
        DBGrid1.BackColor = "&H" & Hex(225) & Hex(200) & Hex(200)
    Case Is = 7, 16
        DBGrid1.BackColor = "&H" & Hex(245) & Hex(200) & Hex(245)
    Case Is = 8, 17
        DBGrid1.BackColor = "&H" & Hex(245) & Hex(200) & Hex(225)
    Case Is = 9, 18
        DBGrid1.BackColor = "&H" & Hex(245) & Hex(200) & Hex(245)
    Case Else
        DBGrid1.BackColor = "&H" & Hex(200) & Hex(255) & Hex(255)
    End Select
    
    With rs_HandValBidReq
        .Index = "PrimaryKey"
        .Seek "=", Trim$(cmb_ValSystem)
        If .NoMatch Then
            MsgBox cmb_ValSystem & " is not in the database."
            Exit Sub
        End If
        If new1 Then
            mSaveMinBidReq_Click
        ElseIf cmb_ValSystem <> "1.All" Then
            ' Fill bid requirement boxes: '
            t_1Maj = "" & !Maj_1Opening
            t_1Min = "" & !Min_1Opening
            t_2C_Maj = "" & !Maj_2C
            t_2C_Min = "" & !Min_2C
            t_1over1 = "" & !OneOver1Resp
            t_2over1 = "" & !TwoOver1Resp
            t_MajGame = "" & !MajGame
            t_MinGame = "" & !MinGame
            t_SmallSlam = "" & !SmallSlam
            t_GrandSlam = "" & !GrandSlam
            t_NT(0) = "" & !Str_1N_Range
            t_NT(1) = "" & !Bal_2N_Range
            t_NT(2) = "" & !Bal_3N_Range
            t_NT(3) = "" & !Bal_2C_2N_Range
            t_NT(4) = "" & !Bal_2C_3N_Range
        Else
            t_1Maj = "" & !Maj_1Opening
        End If
        ' Fill grid: '
        Data1.RecordSource = "SELECT * FROM HandVal " & _
                             "WHERE Name='" & cmb_ValSystem & "' " & _
                             "ORDER By HandVal.ID;"
        Data1.Refresh
        Grid1.Refresh
    End With

    If new1 And cmb_ValSystem.List(cmb_ValSystem.ListIndex) <> n Then
        i = 2
        Do While cmb_ValSystem.List(i) <> n
            i = i + 1
        Loop
        If cmb_ValSystem.List(i) = n Then
            'bln_Cancel = True ' Keep this sub from running again.
            cmb_ValSystem.ListIndex = i
            'bln_Cancel = False
        End If
    End If

End Sub


Private Sub ClearBoxes()
        t_1Maj = ""
        t_1Min = ""
        t_2C_Maj = ""
        t_2C_Min = ""
        t_1over1 = ""
        t_2over1 = ""
        t_MajGame = ""
        t_MinGame = ""
        t_SmallSlam = ""
        t_GrandSlam = ""
        t_NT(0) = ""
        t_NT(1) = ""
        t_NT(2) = ""
        t_NT(3) = ""
        t_NT(4) = ""
End Sub


Private Sub DBGrid1_Click()
    Timer1.Interval = 1
    ' When the grid has the focus and you click '
    ' on a cell, when the Click event runs, the '
    ' .Col value has not changed to the new col '
    ' yet, but if some other control has the    '
    ' focus when you click on a cell in the grid'
    ' the .Col value *is* the value of the col  '
    ' you clicked on.  I can see no way to make '
    ' a distinction between the two cases with- '
    ' out using flags, etc., so it's easier to  '
    ' just set a time to run when the cell is   '
    ' clicked. This gives the .Col value time to'
    ' update before the Timer calls the routine '
    ' to show the help info for the cell.       '
End Sub


Private Sub mWow_Click()
    If mWow.Checked Then
        mWow.Checked = False
        Exit Sub
    End If
    mWow.Checked = True
    ' Swirl the colors: '
    Do While mWow.Checked
        Swirl
        DoEvents
    Loop

    ' Get back to starting gradient angle: '
    With GradFrame1
    Do While .GradientAngle <> 160
        Swirl
    Loop
    End With
End Sub


Private Sub Swirl()
    Dim i As Integer

    With GradFrame1
        If .GradientAngle <= 350 Then
            .GradientAngle = .GradientAngle + 10
            For i = 0 To 3
                GradFrame2(i).GradientAngle = GradFrame2(i).GradientAngle + 10
            Next
        Else
            .GradientAngle = 0
            For i = 0 To 3
                GradFrame2(i).GradientAngle = 0
            Next
        End If
    End With
End Sub


Private Sub t_1Maj_GotFocus()
    L_Info.Top = 50
    L_Info = vbCr & "Minimum Valuation Points " & _
                    "required to open 1 of a major."
End Sub


Private Sub t_1Min_GotFocus()
    L_Info.Top = 50
    L_Info = vbCr & "Minimum Valuation Points " & _
                    "required to open 1 of a minor."
End Sub


Private Sub t_1over1_GotFocus()
    L_Info.Top = 50
    L_Info = vbCr & "Minimum Valuation Points required to " & _
                    "respond 1-over-1 (e.g.: 1C-P-1H)."
End Sub

Private Sub t_2C_Maj_GotFocus()
    L_Info.Top = 50
    L_Info = vbCr & "Minimum Valuation Points required to open " & _
                    "a strong, forcing 2C with a Major suit."
End Sub


Private Sub t_2C_Min_GotFocus()
    L_Info.Top = 50
    L_Info = vbCr & "Minimum Valuation Points required to " & _
                    "open a strong, forcing 2C with a Minor suit."
End Sub


Private Sub t_2over1_GotFocus()
    L_Info.Top = 150
    L_Info = "Minimum Valuation Points required to make a " & _
             "2-over-1 (non-jump) response to an opening bid." & vbCr & _
             "This is the total points for the current hand " & _
             "PLUS the expected minimum from partner."
End Sub


Private Sub t_GrandSlam_GotFocus()
    L_Info.Top = 50
    L_Info = vbCr & "Minimum Valuation Points required to " & _
                    "make a Grand Slam."
End Sub


Private Sub t_Len_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
    Case 97 To 122
        KeyAscii = KeyAscii - 32
    End Select
End Sub


Private Sub t_MajGame_GotFocus()
    L_Info.Top = 150
    L_Info = "Minimum Valuation Points required to make " & _
             "game in a Major suit." & vbCr & _
             "This is the total points for the current hand " & _
             "PLUS the expected minimum from partner."
End Sub


Private Sub t_MinGame_GotFocus()
    L_Info.Top = 150
    L_Info = "Minimum Valuation Points required to make " & _
             "game in a Minor suit." & vbCr & _
             "This is the total points for the current hand " & _
             "PLUS the expected minimum from partner."
End Sub


Private Sub t_NT_GotFocus(Index As Integer)
    L_Info.Top = 150
    L_Info = "Expected range of points for each " & _
             "notrump opening level." & vbCr & _
             "The 1NT range is for a strong NT opening  " & _
             "(e.g.: in HCPs, 15-17)."
End Sub


Private Sub t_SmallSlam_GotFocus()
    L_Info.Top = 50
    L_Info = vbCr & "Minimum Valuation Points required to " & _
                    "make a Small Slam."
End Sub


Private Sub Timer1_Timer()
    Timer1.Interval = 0
    ShowGridHelp
End Sub


Private Sub ShowGridHelp()
    Dim i As Long
    Dim d As String
    Dim n As String

    With DBGrid1
    '          '
    If .col = int_Col_Suit - 1 Then .col = int_Col_Suit
    If .col = int_Col_Len + 1 Then .col = int_Col_Len
    Select Case .col
    Case Is = 0 ' ID
            L_Info.Top = 150
            L_Info = "ID code is just for keeping the grid in " & _
                     "the preferred order." & vbCr & _
                     "The ID has no effect on computations, other " & _
                     "than Flag entries must be ahead of Flag tests."

            If .row > 0 And .Text = "" And .Columns(1) = "" Then
                On Local Error GoTo CantLeaveRow
                .row = .row - 1
                On Local Error GoTo 0
                d = .Text
                .row = .row + 1
                i = Val(Mid$(d, 2)) + 10
                d = Left$(d, 1) & i
                .Text = d
            End If
        Case Is = 1 ' Name
            L_Info.Top = 230
            L_Info = "System name, as selected in the System Name " & _
                     "list box, at top-left."
            If .row > 0 And .Text = "" Then
                .Text = cmb_ValSystem
            End If
        Case Is = 2 ' Description
            L_Info.Top = 230
            L_Info = "Description of this entry's purpose."
        Case Is = int_Col_NTSuit ' NT_or_Suit
            L_Info.Top = 150
            L_Info = "Enter S if this entry is for suit valuation, " & _
                     "N for Notrump." & vbCr & "Leave blank to use for both."
        Case Is = int_Col_Value ' Value
            L_Info.Top = 150
            L_Info = "Formats: .4*n = .4 x suit len.  L-S = Longest " & _
                     "minus Shortest suit length." & vbCr & _
                     "L1+L2 = Two longest suits' length.  T-S = Trump " & _
                     "suit minus Shortest suit.  .25*HCP = .25 x HCPs."
        Case Is = int_Col_Flag ' Flag
            L_Info.Top = 150
            L_Info = "Flag (any characters) to count number of times " & _
                     "entry's specs are met." & vbCr & "Flag=n or Flag>0 " & _
                     "use the results as a specification for an adjustment."
        Case Is = int_Col_Pos ' Position
            L_Info.Top = 50
            L_Info = "Seat number, e.g.:    3 = seat 3 (where Dealer=1),     " & _
                     "2-4 = seats 2 thru 4,   etc." & vbCr & _
                     "Add  L, P, R  to indicate dealer:  4L = 4th seat, " & _
                     "LHO opened,    2-3R = 2nd or 3rd seat, RHO opened." & vbCr & _
                     "Add suit bid or not bid:   2-4RS = seat 2-4, RHO " & _
                     "opened Spades   or   '-S'= opened anything BUT Spades."
        Case Is = int_Col_Hands
            L_Info.Top = 50
            L_Info = "SUITS:   L=LHO's,   LC=Cue bid,   LR=Rebiddable suit,   " & _
                     "LF=Fit for pard,   LS=Short suit,   LN=bid NT.      " & _
                     "MS=0 my void,   MS=1,   MS=2" & vbCr & _
                     "STRENGTH:  L/P/R =B  =W  =I  =O  =P  =S  =V  =X  (See Help.)     " & _
                     "VPS=## = My 'Suit' VPs,   VPN=## = My 'NT' VPs,   HCP=# = My HCPs." & vbCr & _
                     "COMPARE, examples:   MS=L,   PC=L,   M=LS  (or use '<>') (M or P first).    " & _
                     "L=LHO.  For 'L', substitute:  R=RHO,  E=Either,  P=Pard,  M=Me."
        Case Is = int_Col_Fit
            L_Info.Top = 150
            L_Info = "Y = fit established.  N = no fit with pard." & vbCr & _
                     "YC, YD, YH, YS = fit in the specified suit."
        Case Is = int_Col_Suit
            L_Info.Top = 50
            L_Info = "M=My suit.    M1=1st/longest suit.    M2=2nd suit.    " & _
                     "MH=Help suit.   MS=Short suit.    MR=Rebiddable suit.    " & _
                     "MW=Weak suit.    MC=My Cue bid suit." & vbCr & _
                     "MF=My Fit (for pard) suit.    F=any Fit with my Partner.        " & _
                     "A3=Any (or All points in) 3 suits.   A2=Any 2 suits.      " & _
                     "Specify suits as:  Spds,  Hrts,  Dias, Clbs" & vbCr & _
                     "In the above codes, substitute:  'P' for Partner,   " & _
                     "'L' for LHO,   'R' for RHO.   'E' for Either opponent." & _
                     "-E=neither opp's suit,  -P=Any suit BUT Pard's, etc."
        Case Is = int_Col_Cards
            L_Info.Top = 50
            L_Info = "Enter specific cards, with or without wildcards.    " & _
                     "Q with no wildcard = Q and no higher honor.    " & _
                     "QJ = QJ with no higher honors.    " & vbCr & _
                     "*Q = Q with or without other honors.   " & _
                     "?Q = Q with exactly 1 higher honor.   " & _
                     "A? = A with exactly 1 lower honor.   " & _
                     "??T = Ten with exactly 2 higher honors." & vbCr & _
                     "-K = no K.     *J && -Q = J with any other honors but Q.     " & _
                     "KQ && -J && -T = exactly KQ.     " & _
                     "? = exactly 1 honor.     ?? = exactly 2 honors, etc."
        Case Is = int_Col_Len
            L_Info.Top = 150
            L_Info = "Suit length. 0=void, 2=doubleton, etc.  " & _
                     "'n' refers back to the Value field." & vbCr & _
                     "Each suit is tested for Len and Card specs " & _
                     "unless otherwise specified in Suit field."
        Case Is = int_Col_VP
            L_Info.Top = 150
            L_Info = "Value Points calculated when you click Compute." & vbCr & _
                     "Points are not saved and are overwritten with " & _
                     "each Compute."
    End Select
    '        '
    End With
Exit Sub

CantLeaveRow:
    DBGrid1.Columns(1) = cmb_ValSystem
    Resume
End Sub


Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim d As String
    Dim n As String

    If Shift <> 0 Then Exit Sub

    With DBGrid1
    '          '
    If KeyCode = 37 And .col > 0 Then ' Left
        If .SelStart > 0 Then Exit Sub
        If .col = int_Col_Suit Then
            ' Skip over dividers. '
            .col = .col - 2
        ElseIf .col = int_Col_VP Then
            .col = .col - 2
        Else
            .col = .col - 1
        End If
        DBGrid1_Click
        KeyCode = 0

    ElseIf KeyCode = 39 And .col < 13 Then ' Right
        If .SelStart < Len(.Text) And .SelLength < Len(.Text) Then Exit Sub
        If .col = int_Col_Suit - 2 Then
            .col = int_Col_Suit
        ElseIf .col = int_Col_VP - 2 Then
            .col = int_Col_VP
        Else
            .col = .col + 1
        End If
        DBGrid1_Click
        KeyCode = 0

    ElseIf KeyCode = 40 Then ' Down
        If .Columns(2).Text <> "" Then
            ' If Description is not blank, go on   '
            ' down to a new row; otherwise, it is  '
            ' possible to keep generating endless  '
            ' new rows just by leaning on Down curs.

            ' If ID and Name, then carry them down '
            ' to use in next line if it is blank.  '
            d = Trim$(.Columns(0).Text)
            n = Trim$(.Columns(1).Text)

            .row = .row + 1
            If .Columns(0).Text = "" Then
                .col = 0
                i = Val(Mid$(d, 2)) + 10
                d = Left$(d, 1) & i
                .Text = d
                .Columns(1) = n
            End If
            KeyCode = 0
        End If

    ElseIf KeyCode = 46 Then ' Del
        If .SelStart = 0 And .SelLength = Len(.Text) Then
            i = MsgBox("Delete the highlighted row?", vbYesNo)
            If i <> vbYes Then
                KeyCode = 0
            End If
        End If
    End If
    '      '
    End With

End Sub


Private Sub DBGrid1_LostFocus()
    DBGrid1.SelStart = 0
End Sub


Private Sub DBGrid1_SelChange(Cancel As Integer)
    ' This is for a change in the cells selected. '
    ' When the grid has the focus and you click   '
    ' on the cell to the left of the ID in order  '
    ' to delete the row, it works fine, but if    '
    ' some other control has the focus and you    '
    ' click the same spot to highlight the row,   '
    ' the grid does NOT get the focus!  Instead,  '
    ' the other control retains the focus, and    '
    ' then when you press Delete, that control's  '
    ' text (if any) gets deleted instead of the   '
    ' grid row.   So we do two things here to get '
    ' around the problem. The first it to set the '
    ' .SelStart to 0, then we give the grid focus,'
    ' which it may already have. Then in the grid '
    ' KeyDown event, we trap the Delete key and   '
    ' check to see if .SelStart=0 and .SelLength  '
    ' is = Len(.Text). If so, then we prompt to   '
    ' confirm deletion of the line and delete it. '
    ' If the user has highlighted .Text and then  '
    ' pressed Delete, the program will confirm to '
    ' Delete the line, which is not what the user '
    ' wants. To avoid this, the user should not   '
    ' (and doesn't need to) press Delete when text'
    ' in a cell is highlighted. He can either just'
    ' press Backspace to leave the cell blank, or '
    ' start typing to replace the highlighted text.
    DBGrid1.SelStart = 0
    DBGrid1.SetFocus
    ' p.s.: Would it have killed them to give us  '
    ' some way to know if a row has been selected?'
End Sub


Private Sub Form_Load()
    ' Rimonobant = diet, heart disease, etc. pill
    bln_NeedsClearing = True
    str_LastSystem = "1.All"
    SetColWidths
    OpenHandVal
    SetUpResultsGridAndSystems
    ColorMenus
    Me.Show
    cmb_ValSystem.SetFocus
End Sub


Private Sub SetUpResultsGridAndSystems()
    cmb_ValSystem.AddItem "1.All"
    ' Set up Results grid, which is '
    ' hidden behind t_FindHand box. '
    ' To see it, scroll the list on '
    ' the Properties box in VB and  '
    ' look for L_All, then change   '
    ' the height from 0 to something'
    ' like 4000.                    '
    With rs_HandValBidReq
        .MoveFirst
        Grid1.row = 0
        Grid1.col = 1
        Grid1.Text = "   Suit"
        Grid1.col = 2
        Grid1.Text = "    NT"
        Grid1.row = 1
        Grid1.col = 0
        Do
            cmb_ValSystem.AddItem !EvalName
            cmb_CopyRow_System.AddItem !EvalName
            If Not .EOF Then
                .MoveNext
            End If
        Loop While Not .EOF
    End With
    cmb_ValSystem.ListIndex = 0
    p_All.ZOrder 1
End Sub


Private Sub OpenHandVal()
    ChDir App.Path
    Set ws_HandVal = DBEngine.Workspaces(0)
    Set db_HandVal = ws_HandVal.OpenDatabase("BB-HEP.mdb")
    Set rs_HandVal = db_HandVal.OpenRecordset("HandVal")
    Set rs_HandValBidReq = db_HandVal.OpenRecordset("HandVal_BidReq")
End Sub


Private Sub SetColWidths()
    Me.Move 650, 440
    Me.Height = 10575

    int_Col_NTSuit = 3
    int_Col_Value = 4
    int_Col_Flag = 5
    int_Col_Pos = 6
    int_Col_Hands = 7
    int_Col_Fit = 8
    int_Col_Suit = 10
    int_Col_Cards = 11
    int_Col_Len = 12
    int_Col_VP = 14

    DBGrid1.Columns(0).Width = 700
    DBGrid1.Columns(1).Width = 800
    DBGrid1.Columns(2).Width = 2650
    DBGrid1.Columns(3).Width = 300
    DBGrid1.Columns(4).Width = 600
    DBGrid1.Columns(5).Width = 600
    DBGrid1.Columns(6).Width = 400
    DBGrid1.Columns(7).Width = 900
    DBGrid1.Columns(8).Width = 400
    DBGrid1.Columns(9).Width = 50
    DBGrid1.Columns(10).Width = 500
    DBGrid1.Columns(11).Width = 1430
    DBGrid1.Columns(12).Width = 400
    DBGrid1.Columns(13).Width = 50
    DBGrid1.Columns(14).Width = 500
    DBGrid1.Height = 7220
    Grid1.ColWidth(0) = 790
    Grid1.ColWidth(1) = 550
    Grid1.ColWidth(2) = 520
    Grid1.ColWidth(3) = 550
End Sub


Private Sub ColorMenus()
    ' Drop-down menu colors: '

    Dim mi As MENUINFO
    Dim c As Long
    Dim p As Long

    c = &H80FFFF
    p = 5 ' HS_DIAGCROSS

    With mi
        .cbSize = Len(mi)

        .fMask = MIM_BACKGROUND
        .hbrBack = CreateHatchBrush(p, c)
        SetMenuInfo GetMenu(Me.hwnd), mi  'main menu bar

        .fMask = MIM_BACKGROUND Or MIM_APPLYTOSUBMENUS
        .hbrBack = CreateSolidBrush(vbCyan)
        SetMenuInfo GetSubMenu(GetMenu(Me.hwnd), 0), mi
        'File menu (item 0)

        .hbrBack = CreateSolidBrush(vbMagenta)
        SetMenuInfo GetSubMenu(GetMenu(Me.hwnd), 1), mi
        'Edit menu (item 1)

        .hbrBack = CreateSolidBrush(vbYellow)
        SetMenuInfo GetSubMenu(GetMenu(Me.hwnd), 2), mi
        'View menu (item 2)

        '.hbrBack = CreateSolidBrush(vbGreen)
        'SetMenuInfo GetSubMenu(GetMenu(Me.hwnd), 3), mi
        'Help menu (item 3)

    End With

    DrawMenuBar Me.hwnd

    ' Store initial value of the '
    ' background frame gradient: '
    lng_GradBlend0 = GradFrame1.GradBlend0
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    ClearVP
    rs_HandVal.Close
    rs_HandValBidReq.Close
    db_HandVal.Close
    ws_HandVal.Close
End Sub


Private Sub t_Bidding_GotFocus()
    L_Info.Top = 150
    L_Info = "Enter bidding as a continuous string " & _
             "of bids separated by dashes." & vbCr & _
             "Use C, D, H, S for the suits,    N for Notrump,    " & _
             "D for Double,    R for Redouble,    P for Pass."
    t_Bidding.SelStart = 0
    t_Bidding.SelLength = 32768
End Sub


Private Sub t_Bidding_KeyPress(KeyAscii As Integer)
    Dim k As String

    If KeyAscii = 3 Then Exit Sub

    k = Chr$(KeyAscii)

    If k = " " Then
        ' Change space to dash: '
        k = "-"
        KeyAscii = 45
    End If

    If InStr("cdhsnpr", k) > 0 Then
        ' Uppercase letters: '
        KeyAscii = KeyAscii - 32
        k = UCase$(k)
    End If

    With t_Bidding
    '            '
    If .SelLength = Len(.Text) Then
        .Text = ""
    End If

    If InStr("CDHSNPR", k) = 0 And _
       (k < "1" Or k > "7") And _
       KeyAscii <> 3 And _
       KeyAscii <> 8 And _
       KeyAscii <> 22 And _
       KeyAscii <> 24 And _
       k <> "-" _
    Then
        ' Reject invalid keypress: '
        KeyAscii = 0
        Exit Sub
    ElseIf InStr("CHSN", k) > 0 And _
           .SelStart > 0 _
    Then
        If Len(.Text) = 0 Or _
           Val(Mid$(.Text, .SelStart)) = 0 _
        Then
            KeyAscii = 0
            Exit Sub
        End If

    ElseIf Val(k) > 0 Or (InStr("PR", k) And _
           .SelStart > 0) > 0 _
    Then
        If Len(.Text) > 1 And _
           Mid$(" " & .Text, .SelStart + 1, 1) <> "-" _
        Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    '      '
    End With
End Sub


Private Sub t_Dealer_GotFocus()
    L_Info.Top = 50
    L_Info = vbCr & "Dealer:    W = LHO     N = Pard    " & _
             "E = RHO    S = current hand"
    t_Dealer.SelStart = 0
    t_Dealer.SelLength = 32768
End Sub


Private Sub t_Dealer_KeyPress(KeyAscii As Integer)
    Dim k As String

    k = Chr$(KeyAscii)
    If k >= "a" Then
        ' Uppercase letters: '
        KeyAscii = KeyAscii - 32
    End If
End Sub


Private Sub t_FindHand_GotFocus()
    L_Info.Top = 150
    L_Info = "Enter a hand in the format:   A2-KQ43-T9874-JT.    " & _
             "Press Space Bar for a dash." & vbCr & _
             "Letters are automatically uppercased.   " & _
             "Press Enter to Compute Value."
    t_FindHand.SelStart = 0
    t_FindHand.SelLength = 32768
End Sub


Private Sub t_FindHand_KeyPress(KeyAscii As Integer)
    ' This routine controls the input of card hands.  '
    ' It converts the space bar to dashes, uppercases '
    ' card letters (A-J), and validates the order of  '
    ' cards entered.                                  '

    Dim k As String
    Dim i As Integer
    Dim j As Integer
    Dim ss As Integer

    k = Chr$(KeyAscii)

    If KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 3 Then
        ' Backspace, Tab, Ctrl-C
        Exit Sub
    End If

    If KeyAscii = 13 Then ' Enter
        b_Compute_Click
        Exit Sub
    End If

    If Len(t_FindHand) > 15 And t_FindHand.SelLength = 0 Then
        Beep
        KeyAscii = 0
        Exit Sub
    End If

    If k = " " Then
        ' Space to dash: '
        k = "-"
        KeyAscii = 45
    End If

    If InStr("akqjt", k) > 0 Then
        ' Uppercase letters: '
        KeyAscii = KeyAscii - 32
        k = UCase$(k)
    End If

    If InStr("AKQJT", k) = 0 And _
       (k < "2" Or k > "9") And _
       KeyAscii <> 3 And _
       KeyAscii <> 8 And _
       KeyAscii <> 22 And _
       KeyAscii <> "24" And _
       k <> "-" _
    Then
        ' Exit when valid key not pressed. '
        KeyAscii = 0
        Beep
        Exit Sub
    End If

    ' Check for card in correct order '
    ' compared to prior card, if any: '
    i = InStr("AKQJT98765432", k)
    If t_FindHand.SelStart > 0 And _
       i > 0 _
    Then
        j = InStr("AKQJT98765432", Mid$(t_FindHand, t_FindHand.SelStart, 1))
        If j >= i Then 'card entered is out of rank order
            ss = t_FindHand.SelStart
            t_FindHand = Left$(t_FindHand, t_FindHand.SelStart) & _
                         k & Mid$(t_FindHand, t_FindHand.SelStart + 1)
            t_FindHand.SelStart = ss + 1
            Beep
            MsgBox "Card spec is out of order."
            KeyAscii = 8 ' backspace
            Exit Sub
        End If
    End If
    If Len(t_FindHand) = 15 Then
        Beep
    End If
End Sub


'******************************************************************
' Name:  Eval_A_Hand                                              '
' Purpose: Calculate custom hand evaluation points based on specs '
'          in the HandVal table in BidBase.                       '
' See the programmer docs for details and explanations.           '
' All Subs used by this Sub start with "Eval_" to distinguish     '
' them from Subs used by the LookUpBid sub.                       '
'******************************************************************
'
Private Sub Eval_A_Hand(sysName As String)
    Dim n As Single
    Dim m As String
    Dim i As Integer
    Dim j As Integer
    Dim row As Integer
    Dim col As Integer
    Dim startVal As Single
    Dim startNT As Single
    Dim lNewVal As Long
    Dim recNum As Long

    Eval_Initialize sysName

    Do While DBGrid1.Text <> ""
        ' See programmer docs. '

        Eval_SetUpFieldVariables
        recNum = recNum + 1
        
        If Eval_ValSpecOk Then
            If Eval_FlagSpecOk Then
                If Eval_PosSpecOk Then
                    If Eval_HandSpecOk Then
                        If Eval_FitSpecOk Then
                            If Eval_SuitSpecOk Then
                                Eval_CardsSpecs
                            End If
                        End If
                    End If
                End If
            End If
        End If

        int_SuitNum = 0

        ' Display entry's results, if any, in grid: '
        With DBGrid1
        '          '
        If sng_HandVal + sng_HandSuit_Val <> startVal Then
            n = sng_HandVal + sng_HandSuit_Val
            m = "" & (n - startVal)
            If Len(m) > 5 Then
                m = Left$(m, 5)
            End If
            Select Case Len(m)
            Case Is = 1, 2
                m = "   " & m
            Case Is = 3
                m = "  " & m
            Case Is = 4
                m = " " & m
            End Select
            .Columns(int_Col_VP).Text = m
            startVal = n

        ElseIf sng_HandNT_Val <> startNT Then
            m = "" & (sng_HandNT_Val - startNT)
            If Len(m) > 5 Then
                m = Left$(m, 5)
            End If
            Select Case Len(m)
            Case Is = 1, 2
                m = "   " & m
            Case Is = 3
                m = "  " & m
            Case Is = 4
                m = " " & m
            End Select
            .Columns(int_Col_VP).Text = m
            startNT = sng_HandNT_Val
        End If

        If .row = 29 Then
            j = Data1.Recordset.RecordCount - recNum
            If j = 0 Then Exit Do
            Stop ' At first I needed the below, then it worked without it. '
            If j > 29 Then
                .Scroll 0, 29
            ElseIf j > 0 Then
                .Scroll 0, j
            Else
                Exit Do
            End If
            .row = 0
        ElseIf .row = Data1.Recordset.RecordCount - 1 Then
            Exit Do
        Else
            .row = .row + 1
        End If
        
        If .Text = "" Then Exit Do

        End With
        '       '
        DoEvents
    Loop

    j = 1
    Do While str_ValSystem(j) <> sysName And j < 20
        j = j + 1
    Loop
    If str_ValSystem(j) <> sysName Then
        j = 0
    End If
    sng_ValPoints(j, 1) = sng_HandVal + sng_HandSuit_Val
    sng_ValPoints(j, 2) = sng_HandVal + sng_HandNT_Val
    For i = 1 To 2
        n = Fix(sng_ValPoints(j, i))
        If sng_ValPoints(j, i) - n >= 0.5 Then
            sng_ValPoints(j, i) = n + 0.5
        Else
            sng_ValPoints(j, i) = n
        End If
    Next
End Sub


Private Sub Eval_Initialize(sysName As String)
    sng_HandVal = 0
    sng_HandNT_Val = 0
    sng_HandSuit_Val = 0
    Erase int_Shpn()
    Erase int_HCPn()
    str_HandShape = ""
    Erase str_Flag()
    Erase int_Flag()
    int_FlagNum = 0
    int_NumFlags = 0

    DBGrid1.row = 0
    DBGrid1.col = 0
    Eval_CalcHandShape
    Eval_SortSuitsByHCP
    GetOtherPlayersData
End Sub


Private Sub Eval_CalcHandShape()
    ' Sort suits by length where s1(4) = longest: '
    ' (Doesn't matter here, but 1=C, 4=S.)        '
    Dim n As Integer

    If int_Shpn(1) > 0 Then ' already done '
        Exit Sub
    End If

    For n = 1 To 4
        int_Shpn(n) = n
    Next

    Do
        If int_SuitLen(int_Shpn(2)) < int_SuitLen(int_Shpn(1)) Then
            int_Shpn(0) = int_Shpn(2)
            int_Shpn(2) = int_Shpn(1)
            int_Shpn(1) = int_Shpn(0)
        ElseIf int_SuitLen(int_Shpn(3)) < int_SuitLen(int_Shpn(2)) Then
            int_Shpn(0) = int_Shpn(2)
            int_Shpn(2) = int_Shpn(3)
            int_Shpn(3) = int_Shpn(0)
        ElseIf int_SuitLen(int_Shpn(4)) < int_SuitLen(int_Shpn(3)) Then
            int_Shpn(0) = int_Shpn(4)
            int_Shpn(4) = int_Shpn(3)
            int_Shpn(3) = int_Shpn(0)
        Else
            Exit Do
        End If
    Loop
    str_HandShape = int_SuitLen(int_Shpn(4)) & _
                    int_SuitLen(int_Shpn(3)) & _
                    int_SuitLen(int_Shpn(2)) & _
                    int_SuitLen(int_Shpn(1))
End Sub


Private Sub Eval_SortSuitsByHCP()
    ' Sort suits by HCP where int_HCPn(4) = highest: '
    Dim n As Integer

    If int_HCPn(1) > 0 Then
        Exit Sub ' Alread done. '
    End If

    For n = 1 To 4
        int_HCPn(n) = n
    Next

    Do
        If int_HCP(int_HCPn(2)) < int_HCP(int_HCPn(1)) Then
            int_HCPn(0) = int_HCPn(2)
            int_HCPn(2) = int_HCPn(1)
            int_HCPn(1) = int_HCPn(0)
        ElseIf int_HCP(int_HCPn(3)) < int_HCP(int_HCPn(2)) Then
            int_HCPn(0) = int_HCPn(2)
            int_HCPn(2) = int_HCPn(3)
            int_HCPn(3) = int_HCPn(0)
        ElseIf int_HCP(int_HCPn(4)) < int_HCP(int_HCPn(3)) Then
            int_HCPn(0) = int_HCPn(4)
            int_HCPn(4) = int_HCPn(3)
            int_HCPn(3) = int_HCPn(0)
        Else
            Exit Do
        End If
    Loop
End Sub


Private Sub Eval_SetUpFieldVariables()
    With DBGrid1
        str_ValSpec = Trim$(.Columns(int_Col_Value))
        sng_ValSpec = Val(str_ValSpec)
        str_FlagSpec = Trim$(.Columns(int_Col_Flag))
        str_PosSpec = Trim$(.Columns(int_Col_Pos))
        str_HandSpec = Trim$(.Columns(int_Col_Hands))
        str_FitSpec = Trim$(.Columns(int_Col_Fit))
        str_SuitSpec = Trim$(.Columns(int_Col_Suit))
        str_CardsSpec = Trim$(.Columns(int_Col_Cards))
        str_LenSpec = Trim$(.Columns(int_Col_Len))
    End With
        
    If Left$(str_LenSpec & " ", 1) = "<" Then
        int_LenSpec = Val(Mid$(str_LenSpec, 2))
    Else
        int_LenSpec = Val(str_LenSpec)
    End If
    int_SuitNum = 0
    int_SuitSpec = 0
    str_CardsSpecCombo = ""
    bln_ValAdded = False
    DBGrid1.col = 0
End Sub


Private Function Eval_FlagSpecOk() As Boolean
    Dim f As String
    Dim fCompare As String
    Dim i As Integer
    Dim n As Integer
    Dim op As String

    ' When an entry has a Hand spec with "flag" in it,  '
    ' we first check to see if the flag's ID is already '
    ' in the str_Flag() array. If it is then set        '
    ' int_FlagNum to its array index, then Eval_AddVal  '
    ' increments the int_Flag() array counter if the    '
    ' entry matches the current hand. If the flag ID is '
    ' *not* in the array, then we add it to the array   '
    ' and set int_FlagNum to the new number. If the flag'
    ' spec has "=" in it, then we compare the number    '
    ' specified after the "=" to the number in int_Flag()

    bln_CountFlag = False

    If str_FlagSpec = "" Then
        int_FlagNum = 0
        Eval_FlagSpecOk = True
        Exit Function
    End If

    f = str_FlagSpec

    ' See if this is a flag test: '
    n = InStr(f, "=") ' e.g.: B1=0
    If n = 0 Then n = InStr(f, ">")
    If n = 0 Then n = InStr(f, "<")
    If n > 1 Then
        fCompare = Mid$(f, n + 1) ' e.g.: "0"
        op = Mid$(f, n, 1) ' e.g.: "="
        f = Left$(f, n - 1) ' e.g.: "B1"
    End If

    ' Look for flag in str_Flag() array: '
    For i = 1 To int_NumFlags
        If str_Flag(i) = f Then
            Exit For
        End If
    Next

    If str_Flag(i) = f Then
        ' Flag spec matches existing flag.'
        If fCompare <> "" Then
            If op = "=" And int_Flag(i) = Val(fCompare) Or _
               op = ">" And int_Flag(i) > Val(fCompare) Or _
               op = "<" And int_Flag(i) < Val(fCompare) _
            Then
                If str_CardsSpec = "" And _
                   str_LenSpec = "" And _
                   str_HandSpec = "" And _
                   str_PosSpec = "" And _
                   str_FitSpec = "" And _
                   str_SuitSpec = "" _
                Then ' No other specs, so add value: '
                    Eval_AddVal
                    Exit Function
                End If
            Else
                Eval_FlagSpecOk = False
                Exit Function
            End If
        End If

    Else
        ' Flag spec doesn't match existing flag.'
        If fCompare <> "" Or f = "stop" Then
            MsgBox "No previous flags were found for" & vbCr & _
                   str_FlagSpec
            Exit Function

        Else ' Set up new flag:
            str_Flag(i) = f
            int_NumFlags = int_NumFlags + 1
        End If
    End If
    If op = "" Then
        bln_CountFlag = True
    End If
    int_FlagNum = i
    Eval_FlagSpecOk = True
End Function


Private Function Eval_HandsSpecCalcOk() As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim p As String
    Dim r As String
    Dim s As String
    Dim suitNum As Integer
    Dim neg As Integer

    s = str_HandSpec
    If Left$(s, 1) = "-" Then
        neg = -1
        s = Mid$(s, 2)
    End If

    If s = "" Then
        Eval_HandsSpecCalcOk = True
        Exit Function
    End If

    If Len(s) = 1 Then
        ' This spec just means that the specified  '
        ' player has shown (bid) at least one suit.'
        Select Case s
        Case Is = "L"
            Eval_HandsSpecCalcOk = (int_LHOs_Suit(1) > 0) Or neg
        Case Is = "R"
            Eval_HandsSpecCalcOk = (int_RHOs_Suit(1) > 0) Or neg
        Case Is = "P"
            Eval_HandsSpecCalcOk = (int_Pards_Suit(1) > 0) Or neg
        Case Is = "M"
            Eval_HandsSpecCalcOk = (int_My_Suit(1) > 0) Or neg
        Case Is = "E"
            Eval_HandsSpecCalcOk = (int_LHOs_Suit(1) > 0 Or _
                                    int_RHOs_Suit(1) > 0) Or neg
        Case Else
            MsgBox "Unrecognized Hands spec: " & str_HandSpec
        End Select
        Exit Function
    End If
    
    If Len(s) = 3 Or (Len(s) = 4 And Mid$(s, 2, 2) = "<>") Then
        i = InStr("LPR", Left$(s, 1))
        k = InStr("BWIOPSVX", Right$(s, 1))
        If Len(s) = 4 And k > 0 Then
            Eval_HandsSpecCalcOk = (t_Str(i) <> Right$(s, 1)) Or t_Str(i - 1) = ""
        Else
            j = InStr("=<>", Mid$(s, 2, 1))
            If i > 0 And j > 0 And k > 0 Then
                i = i - 1
                Select Case Mid$(s, 2, 1)
                Case Is = "="
                    Eval_HandsSpecCalcOk = (t_Str(i) = Right$(s, 1))
                Case Is = "-"
                    Eval_HandsSpecCalcOk = (t_Str(i) <> Right$(s, 1)) Or t_Str(i) = ""
                Case Is = "<"
                    Eval_HandsSpecCalcOk = (InStr("BWIOPSVX", t_Str(i)) < _
                                            InStr("BWIOPSVX", Right$(s, 1))) Or t_Str(i) = ""
                Case Is = ">"
                    Eval_HandsSpecCalcOk = (InStr("BWIOPSVX", t_Str(i)) > _
                                            InStr("BWIOPSVX", Right$(s, 1)))
                End Select
            End If
            Exit Function
        End If
    End If
    
    If Left$(s & "  ", 3) = "HCP" Then
        k = Val(Mid$(s & "  ", 5)) ' e.g.: HCP=15+ '
        If Mid$(s & " ", 4, 1) = "=" Then
            If Right$(s, 1) = "+" Then
                Eval_HandsSpecCalcOk = (int_HCP(0) >= k)
            ElseIf InStr(s, "-") > 0 Then
                j = Val(Mid$(s, InStr(s, "-") + 1))
                Eval_HandsSpecCalcOk = (int_HCP(0) >= k And _
                                        int_HCP(0) <= j)
            Else
                Eval_HandsSpecCalcOk = (int_HCP(0) = k)
            End If
        ElseIf InStr(s & " ", 4, 1) = "<" Then
            Eval_HandsSpecCalcOk = (int_HCP(0) < k)
        End If
        Exit Function
    ElseIf Left$(s & "  ", 3) = "VPS" Then
        k = Val(Mid$(s & "  ", 5)) ' e.g.: VPS=15+ '
        If Mid$(s & " ", 4, 1) = "=" Then
            If Right$(s, 1) = "+" Then
                Eval_HandsSpecCalcOk = ((sng_HandVal + sng_HandSuit_Val) >= k)
            Else
                Eval_HandsSpecCalcOk = ((sng_HandVal + sng_HandSuit_Val) = k)
            End If
        ElseIf InStr(s & " ", 4, 1) = "<" Then
            Eval_HandsSpecCalcOk = ((sng_HandVal + sng_HandSuit_Val) < k)
        End If
        Exit Function
    ElseIf Left$(s & "  ", 3) = "VPN" Then
        k = Val(Mid$(s & "  ", 5)) ' e.g.: VPN=15+ '
        If Mid$(s & " ", 4, 1) = "=" Then
            If Right$(s, 1) = "+" Then
                Eval_HandsSpecCalcOk = ((sng_HandVal + sng_HandNT_Val) >= k)
            Else
                Eval_HandsSpecCalcOk = ((sng_HandVal + sng_HandNT_Val) = k)
            End If
        ElseIf InStr(s & " ", 4, 1) = "<" Then
            Eval_HandsSpecCalcOk = ((sng_HandVal + sng_HandNT_Val) < k)
        End If
        Exit Function
    End If

    p = Left$(s, 1)                 ' eg: L, LS, L=Spds, L=15, P, PC, PC<>L
    r = Trim$(Mid$(s & " ", 2))  ' eg:  ,  S,  =Spds,  =15,  ,  C,  C<>L

    Select Case p
    '            '
    Case Is = "L"
    ' LHO Specs: '
        If Left$(r, 1) = "=" Then
            ' e.g.: L=15 = LHO has 15+ VPs (entered by user) '
            Eval_HandsSpecCalcOk = (Val(str_LHOs_Strength) >= Val(Mid$(r, 2)))
            Exit Function
        ElseIf Left$(r, 1) = "<" Then
            ' e.g.: L<20 = LHO has <15 VPs '
            Eval_HandsSpecCalcOk = (Val(str_LHOs_Strength) < Val(Mid$(r, 2)))
            Exit Function
        End If

        If Len(s) = 2 Then ' LN, LS, LC, LR, LF -- alone. '
            Select Case r
            Case Is = "C"
                Eval_HandsSpecCalcOk = (int_LHOs_CueBid > 0) Or neg
            Case Is = "F"
                Eval_HandsSpecCalcOk = (int_LHOs_Fit(1) > 0) Or neg
            Case Is = "N"
                Eval_HandsSpecCalcOk = bln_LHO_Bid_NT Or neg
            Case Is = "R"
                Eval_HandsSpecCalcOk = (int_LHOs_RebidSuit(1) > 0) Or neg
            Case Is = "S"
                Eval_HandsSpecCalcOk = (int_LHOs_SS > 0) Or neg
            End Select
            Exit Function
        End If

        suitNum = InStr("   ClbsDiasHrtsSpds", Mid$(s, 3))
        If suitNum > 3 Then
            ' L=Clbs, etc. '
            suitNum = suitNum / 4
            Eval_HandsSpecCalcOk = (int_LHOs_Suit(1) = j Or _
                                    int_LHOs_Suit(2) = j)
        Else
            suitNum = InStr("   ClbsDiasHrtsSpds", Mid$(s, 4))
            If suitNum > 3 Then
                ' LS=Clbs, etc. '
                suitNum = suitNum / 4
                Select Case Left$(p, 2)
                Case Is = "LC"
                    Eval_HandsSpecCalcOk = (int_LHOs_CueBid = j)
                Case Is = "LF"
                    Eval_HandsSpecCalcOk = (int_LHOs_Fit(1) = j Or _
                                            int_LHOs_Fit(2) = j)
                Case Is = "LH"
                    Eval_HandsSpecCalcOk = (int_LHOs_HelpSuit = j)
                Case Is = "LR"
                    Eval_HandsSpecCalcOk = (int_LHOs_RebidSuit(1) = j Or _
                                            int_LHOs_RebidSuit(1) = j)
                Case Is = "LS"
                    Eval_HandsSpecCalcOk = (int_LHOs_SS = j)
                Case Else
                    MsgBox "Hand spec not recognized: " & s
                End Select
            End If
        End If
        
        Exit Function
                
    Case Is = "R"
    ' RHO Specs: '
        If Left$(r, 1) = "=" Then
            ' e.g.: R=15 = RHO has 15+ VPs (entered by user) '
            Eval_HandsSpecCalcOk = (Val(str_RHOs_Strength) >= Val(Mid$(r, 2)))
            Exit Function
        ElseIf Left$(r, 1) = "<" Then
            ' e.g.: R<20 = RHO has <15 VPs '
            Eval_HandsSpecCalcOk = (Val(str_RHOs_Strength) < Val(Mid$(r, 2)))
            Exit Function
        End If

        If Len(s) = 2 Then ' RN, RS, RC, RR, RF -- alone. '
            Select Case r
            Case Is = "C"
                Eval_HandsSpecCalcOk = (int_RHOs_CueBid > 0) Or neg
            Case Is = "F"
                Eval_HandsSpecCalcOk = (int_RHOs_Fit(1) > 0) Or neg
            Case Is = "N"
                Eval_HandsSpecCalcOk = bln_RHO_Bid_NT Or neg
            Case Is = "R"
                Eval_HandsSpecCalcOk = (int_RHOs_RebidSuit(1) > 0) Or neg
            Case Is = "S"
                Eval_HandsSpecCalcOk = (int_RHOs_SS > 0) Or neg
            End Select
            Exit Function
        End If

        suitNum = InStr("   ClbsDiasHrtsSpds", Mid$(s, 3))
        If suitNum > 3 Then
            ' R=Clbs, etc. '
            suitNum = suitNum / 4
            Eval_HandsSpecCalcOk = (int_RHOs_Suit(1) = j Or _
                                    int_RHOs_Suit(2) = j)
        Else
            suitNum = InStr("   ClbsDiasHrtsSpds", Mid$(s, 4))
            If suitNum > 3 Then
                ' RS=Clbs, etc. '
                suitNum = suitNum / 4
                Select Case Left$(p, 2)
                Case Is = "RC"
                    Eval_HandsSpecCalcOk = (int_RHOs_CueBid = j)
                Case Is = "RF"
                    Eval_HandsSpecCalcOk = (int_RHOs_Fit(1) = j Or _
                                            int_RHOs_Fit(2) = j)
                Case Is = "RH"
                    Eval_HandsSpecCalcOk = (int_RHOs_HelpSuit = j)
                Case Is = "LR"
                    Eval_HandsSpecCalcOk = (int_RHOs_RebidSuit(1) = j Or _
                                            int_RHOs_RebidSuit(1) = j)
                Case Is = "RS"
                    Eval_HandsSpecCalcOk = (int_RHOs_SS = j)
                Case Else
                    MsgBox "Hand spec not recognized: " & s
                End Select
            End If
        End If
        
        Exit Function
                
    Case Is = "P"
    ' Pard's specs: '
        If Left$(r, 1) = "=" Then
            ' e.g.: P=15 = Pard has 15+ VPs (entered by user) '
            Eval_HandsSpecCalcOk = (Val(str_Pards_Strength) >= Val(Mid$(r, 2)))
            Exit Function
        ElseIf Left$(r, 1) = "<" Then
            ' e.g.: P<20 = Pard has <15 VPs '
            Eval_HandsSpecCalcOk = (Val(str_Pards_Strength) < Val(Mid$(r, 2)))
            Exit Function
        End If
        
        If Len(r) = 1 Then
            Select Case r  ' PC, PF, PN, PR, PS -- alone '
            Case Is = "C"
                Eval_HandsSpecCalcOk = (int_Pards_CueBid > 0) Or neg
            Case Is = "F"
                Eval_HandsSpecCalcOk = (int_Pards_Fit(1) > 0) Or neg
            Case Is = "N"
                Eval_HandsSpecCalcOk = bln_Pard_Bid_NT Or neg
            Case Is = "R"
                Eval_HandsSpecCalcOk = (int_Pards_RebidSuit(1) > 0) Or neg
            Case Is = "S"
                Eval_HandsSpecCalcOk = (int_Pards_SS > 0) Or neg
            End Select
            Exit Function
        End If

        suitNum = InStr("   ClbsDiasHrtsSpds", Mid$(s, 3))
        If suitNum > 3 Then
            ' P=Clbs, etc. '
            suitNum = suitNum / 4
            Eval_HandsSpecCalcOk = (int_Pards_Suit(1) = j Or _
                                    int_Pards_Suit(2) = j)
            Exit Function
        Else
            suitNum = InStr("   ClbsDiasHrtsSpds", Mid$(s, 4))
            If suitNum > 3 Then
                ' PS=Clbs, etc. '
                suitNum = suitNum / 4
                Select Case Left$(p, 2)
                Case Is = "PC"
                    Eval_HandsSpecCalcOk = (int_Pards_CueBid = j)
                Case Is = "PF"
                    Eval_HandsSpecCalcOk = (int_Pards_Fit(1) = j Or _
                                            int_Pards_Fit(2) = j)
                Case Is = "PH"
                    Eval_HandsSpecCalcOk = (int_Pards_HelpSuit = j)
                Case Is = "PR"
                    Eval_HandsSpecCalcOk = (int_Pards_RebidSuit(1) = j Or _
                                            int_Pards_RebidSuit(1) = j)
                Case Is = "PS"
                    Eval_HandsSpecCalcOk = (int_Pards_SS = j)
                Case Else
                    MsgBox "Hand spec not recognized: " & s
                End Select
                Exit Function
            End If
        End If

        ' Pard or Me comes first comparing to LHO/RHO '
        ' so LHO and RHO don't have these sections:   '

        Select Case s
        ' not equal: '
        Case Is = "PS<>L"
            Eval_HandsSpecCalcOk = (int_Pards_SS = 0 Or _
                            (int_Pards_SS <> int_LHOs_Suit(1) And _
                             int_Pards_SS <> int_LHOs_Suit(2)))
        Case Is = "PS<>R"
            Eval_HandsSpecCalcOk = (int_Pards_SS = 0 Or _
                            (int_Pards_SS <> int_RHOs_Suit(1) And _
                             int_Pards_SS <> int_RHOs_Suit(2)))
        Case Is = "PC<>L"
            Eval_HandsSpecCalcOk = (int_Pards_CueBid = 0 Or _
                            (int_Pards_CueBid <> int_LHOs_Suit(1) And _
                             int_Pards_CueBid <> int_LHOs_Suit(2)))
        Case Is = "PC<>R"
            Eval_HandsSpecCalcOk = (int_Pards_CueBid = 0 Or _
                            (int_Pards_CueBid <> int_RHOs_Suit(1) And _
                             int_Pards_CueBid <> int_RHOs_Suit(2)))
        Case Is = "P<>LS"
            Eval_HandsSpecCalcOk = (int_Pards_Suit(1) = 0 Or _
                            (int_Pards_Suit(1) <> int_LHOs_SS And _
                             int_Pards_Suit(2) <> int_LHOs_SS))
        Case Is = "P<>RS"
            Eval_HandsSpecCalcOk = (int_Pards_Suit(1) = 0 Or _
                            (int_Pards_Suit(1) <> int_RHOs_SS And _
                             int_Pards_Suit(2) <> int_RHOs_SS))
        Case Is = "P<>LC"
            Eval_HandsSpecCalcOk = (int_Pards_Suit(1) = 0 Or _
                            (int_Pards_Suit(1) <> int_LHOs_CueBid And _
                             int_Pards_Suit(2) <> int_LHOs_CueBid))
        Case Is = "P<>RC"
            Eval_HandsSpecCalcOk = (int_Pards_Suit(1) = 0 Or _
                            (int_Pards_Suit(1) <> int_RHOs_CueBid And _
                             int_Pards_Suit(2) <> int_RHOs_CueBid))

        ' equal: '
        Case Is = "PS=L"
            Eval_HandsSpecCalcOk = _
                            (int_Pards_SS = int_LHOs_Suit(1) And _
                             int_Pards_SS = int_LHOs_Suit(2))
        Case Is = "PS=R"
            Eval_HandsSpecCalcOk = _
                            (int_Pards_SS <> int_RHOs_Suit(1) And _
                             int_Pards_SS = int_RHOs_Suit(2))
        Case Is = "PC=L"
            Eval_HandsSpecCalcOk = _
                            (int_Pards_CueBid = int_LHOs_Suit(1) And _
                             int_Pards_CueBid = int_LHOs_Suit(2))
        Case Is = "PC=R"
            Eval_HandsSpecCalcOk = _
                            (int_Pards_CueBid = int_RHOs_Suit(1) And _
                             int_Pards_CueBid = int_RHOs_Suit(2))
        Case Is = "P=LS"
            Eval_HandsSpecCalcOk = _
                            (int_Pards_Suit(1) = int_LHOs_SS And _
                             int_Pards_Suit(2) = int_LHOs_SS)
        Case Is = "P=RS"
            Eval_HandsSpecCalcOk = _
                            (int_Pards_Suit(1) = int_RHOs_SS And _
                             int_Pards_Suit(2) = int_RHOs_SS)
        Case Is = "P=LC"
            Eval_HandsSpecCalcOk = _
                            (int_Pards_Suit(1) = int_LHOs_CueBid And _
                             int_Pards_Suit(2) = int_LHOs_CueBid)
        Case Is = "P=RC"
            Eval_HandsSpecCalcOk = _
                            (int_Pards_Suit(1) = int_RHOs_CueBid And _
                             int_Pards_Suit(2) = int_RHOs_CueBid)
        End Select
        Exit Function

    Case Is = "M"
        ' "My hand" (the hand entered by user) '
        ' Remember that just "M" has already   '
        ' been handled at the top, so from here'
        ' down is "M" and something else.      '
        ' Variable 'r' is the rest of the spec '
        ' following the "M".                   '

        If Len(r) = 1 Then ' MN, MS, MC, MR, MF -- alone. '
            Select Case r
            Case Is = "C"
                Eval_HandsSpecCalcOk = (int_My_CueBid > 0) Or neg
            Case Is = "F"
                Eval_HandsSpecCalcOk = (int_My_Fit(1) > 0) Or neg
            Case Is = "N"
                Eval_HandsSpecCalcOk = bln_I_Bid_NT Or neg
            Case Is = "R"
                Eval_HandsSpecCalcOk = (int_My_RebidSuit(1) > 0) Or neg
            Case Is = "S"
                Eval_HandsSpecCalcOk = (int_SuitLen(int_My_SS(1)) < 2) Or neg
            End Select
            Exit Function
        End If
        
        If Left$(s, 3) = "MS=" Then
            Select Case Right$(s, 1)
            Case Is = "0"
                If int_My_SS(1) > 0 Then
                    Eval_HandsSpecCalcOk = (int_SuitLen(int_My_SS(1)) = 0)
                End If
            Case Is = "1"
                For i = 1 To 3
                    If int_My_SS(i) > 0 Then
                        If int_SuitLen(int_My_SS(1)) = 1 Then
                            Eval_HandsSpecCalcOk = True
                            Exit For
                        End If
                    End If
                Next
            Case Is = "2"
                For i = 1 To 4
                    If int_SuitLen(i) = 2 Then
                        Eval_HandsSpecCalcOk = True
                        Exit For
                    End If
                Next
            Case Else
                MsgBox "Unrecognized Hand spec: " & s
            End Select
            Exit Function
        End If

        suitNum = InStr("   ClbsDiasHrtsSpds", Mid$(s, 3))
        If suitNum > 3 Then
            ' M=Clbs, etc. '
            suitNum = suitNum / 4
            Eval_HandsSpecCalcOk = (int_My_Suit(1) = j Or _
                                    int_My_Suit(2) = j)
            Exit Function
        Else
            suitNum = InStr("   ClbsDiasHrtsSpds", Mid$(s, 4))
            If suitNum > 3 Then
                ' MS=Clbs, etc. '
                suitNum = suitNum / 4
                Select Case Left$(p, 2)
                Case Is = "MC"
                    Eval_HandsSpecCalcOk = (int_My_CueBid = j)
                Case Is = "MF"
                    Eval_HandsSpecCalcOk = (int_My_Fit(1) = j Or _
                                            int_My_Fit(2) = j)
                Case Is = "MH"
                    Eval_HandsSpecCalcOk = (int_My_HelpSuit = j)
                Case Is = "MR"
                    Eval_HandsSpecCalcOk = (int_My_RebidSuit(1) = j Or _
                                            int_My_RebidSuit(1) = j)
                Case Is = "MS"
                    For i = 1 To 3
                        If int_My_SS(i) = j Then
                            Eval_HandsSpecCalcOk = True
                            Exit For
                        End If
                    Next
                Case Else
                    MsgBox "Hand spec not recognized: " & s
                End Select
                Exit Function
            End If
        End If

        ' not equal: '
        Select Case s
        Case Is = "MS<>L"
            Eval_HandsSpecCalcOk = (int_My_SS(1) = 0 Or _
               (int_My_SS(1) <> int_LHOs_Suit(1) And _
                int_My_SS(1) <> int_LHOs_Suit(2) And _
                int_My_SS(2) <> int_LHOs_Suit(1) And _
                int_My_SS(2) <> int_LHOs_Suit(2) And _
                int_My_SS(3) <> int_LHOs_Suit(1) And _
                int_My_SS(3) <> int_LHOs_Suit(2)))
            
        Case Is = "MS<>R"
            Eval_HandsSpecCalcOk = (int_My_SS(1) = 0 Or _
               (int_My_SS(1) <> int_RHOs_Suit(1) And _
                int_My_SS(1) <> int_RHOs_Suit(2) And _
                int_My_SS(2) <> int_RHOs_Suit(1) And _
                int_My_SS(2) <> int_RHOs_Suit(2) And _
                int_My_SS(3) <> int_RHOs_Suit(1) And _
                int_My_SS(3) <> int_RHOs_Suit(2)))
        
        Case Is = "MC<>L"
            Eval_HandsSpecCalcOk = (int_My_CueBid = 0 Or _
                            (int_My_CueBid <> int_LHOs_Suit(1) And _
                             int_My_CueBid <> int_LHOs_Suit(2)))
        Case Is = "MC<>R"
            Eval_HandsSpecCalcOk = (int_My_CueBid = 0 Or _
                            (int_My_CueBid <> int_RHOs_Suit(1) And _
                             int_My_CueBid <> int_RHOs_Suit(2)))
        Case Is = "M<>LS"
            Eval_HandsSpecCalcOk = (int_My_Suit(1) = 0 Or _
                            (int_My_Suit(1) <> int_LHOs_SS And _
                             int_My_Suit(2) <> int_LHOs_SS))
        Case Is = "M<>RS"
            Eval_HandsSpecCalcOk = (int_My_Suit(1) = 0 Or _
                            (int_My_Suit(1) <> int_RHOs_SS And _
                             int_My_Suit(2) <> int_RHOs_SS))
        Case Is = "M<>LC"
            Eval_HandsSpecCalcOk = (int_My_Suit(1) = 0 Or _
                            (int_My_Suit(1) <> int_LHOs_CueBid And _
                             int_My_Suit(2) <> int_LHOs_CueBid))
        Case Is = "P<>RC"
            Eval_HandsSpecCalcOk = (int_My_Suit(1) = 0 Or _
                            (int_My_Suit(1) <> int_RHOs_CueBid And _
                             int_My_Suit(2) <> int_RHOs_CueBid))

        ' equal: '
        Case Is = "MS=L"
            Eval_HandsSpecCalcOk = _
               (int_My_SS(1) = int_LHOs_Suit(1) Or _
                int_My_SS(1) = int_LHOs_Suit(2) Or _
                int_My_SS(2) = int_LHOs_Suit(1) Or _
                int_My_SS(2) = int_LHOs_Suit(2) Or _
                int_My_SS(3) = int_LHOs_Suit(1) Or _
                int_My_SS(3) = int_LHOs_Suit(2))
        
        Case Is = "MS=R"
            Eval_HandsSpecCalcOk = _
               (int_My_SS(1) = int_RHOs_Suit(1) Or _
                int_My_SS(1) = int_RHOs_Suit(2) Or _
                int_My_SS(2) = int_RHOs_Suit(1) Or _
                int_My_SS(2) = int_RHOs_Suit(2) Or _
                int_My_SS(3) = int_RHOs_Suit(1) Or _
                int_My_SS(3) = int_RHOs_Suit(2))
        
        Case Is = "MC=L"
            Eval_HandsSpecCalcOk = _
                            (int_My_CueBid = int_LHOs_Suit(1) And _
                             int_My_CueBid = int_LHOs_Suit(2))
        Case Is = "MC=R"
            Eval_HandsSpecCalcOk = _
                            (int_My_CueBid = int_RHOs_Suit(1) And _
                             int_My_CueBid = int_RHOs_Suit(2))
        Case Is = "M=LS"
            Eval_HandsSpecCalcOk = _
                            (int_My_Suit(1) = int_LHOs_SS And _
                             int_My_Suit(2) = int_LHOs_SS)
        Case Is = "M=RS"
            Eval_HandsSpecCalcOk = _
                            (int_My_Suit(1) = int_RHOs_SS And _
                             int_My_Suit(2) = int_RHOs_SS)
        Case Is = "M=LC"
            Eval_HandsSpecCalcOk = _
                            (int_My_Suit(1) = int_LHOs_CueBid And _
                             int_My_Suit(2) = int_LHOs_CueBid)
        Case Is = "M=RC"
            Eval_HandsSpecCalcOk = _
                            (int_My_Suit(1) = int_RHOs_CueBid And _
                             int_My_Suit(2) = int_RHOs_CueBid)
        End Select
    End Select
End Function


Private Function Eval_SuitSpecOk() As Boolean
    Dim s As String
    Dim s1 As String
    Dim op As String
    Dim i As Integer
    Dim ok As Boolean
    Dim a As String
    Dim n As Boolean
    Dim r As Boolean
    Dim l As Boolean
    Dim l0 As Boolean
    Dim r0 As Boolean

    s = str_SuitSpec
    If s = "" Then
        Eval_SuitSpecOk = True
        Exit Function
    End If

    If Left$(s, 1) = "-" And Len(s) > 1 Then
        op = "-"
        s1 = Mid$(s, 2)
    Else
        s1 = s
    End If

    int_SuitSpec = 0

    If s1 = "R" Then ' RHO's suit. '
        int_SuitSpec = int_RHOs_Suit(1)

    ElseIf s1 = "L" Then ' LHO's suit. '
        int_SuitSpec = int_LHOs_Suit(1)

    ElseIf s = "E" Then
        If int_RHOs_Suit(1) > 0 Or int_LHOs_Suit(1) > 0 Then
            Eval_SuitSpecOk = True
        End If
        Exit Function

    ElseIf s = "ER" Then
        If int_RHOs_RebidSuit(1) > 0 Or int_LHOs_RebidSuit(1) > 0 Then
            Eval_SuitSpecOk = True
        End If
        Exit Function

    ElseIf s = "-M" Or s = "-P" Or s = "-F" Or s = "-E" Or s = "-ER" Then
        ' These are checked in Eval_CardsSpecOk. '
        Eval_SuitSpecOk = True
        Exit Function

    ElseIf s1 = "F" Then ' Fit suit. '
        int_SuitSpec = int_Fit_Suit(1)

    ElseIf s1 = "M" Then ' My suit. '
        int_SuitSpec = int_My_Suit(1)

    ElseIf s1 = "P" Then ' Pard's suit. '
        a = Right$(" " & t_Len(4 + int_Pards_Suit(1)), 1)
        If a <> "F" Then
            ' If a="F", then pard was showing a Fit '
            ' for My suit; otherwise, it's his suit.'
            int_SuitSpec = int_Pards_Suit(1)
        End If

    ElseIf Left$(s & " ", 2) = "PC" Then
        ' Pard's cue A/K control. '
        If int_Pards_CueBid = 0 Then
            Exit Function
        End If

        If s = "PC" Then
            int_SuitSpec = int_Pards_CueBid
        ElseIf s = "PCE" Then
            a = t_Len(4 + int_Pards_Suit(1))
        Else
            Exit Function
        End If

    ElseIf Left$(s & " ", 2) = "PS" Or Mid$(s & "  ", 2) = "PS" Then
        ' Set int_SuitSpec to a suit matching  '
        ' one of the following suits:          '
        ' PS = Partner's Short suit specified. '
        ' -PS = not Pard's SS                  '
        ' PSL = Pard's SS = LHO's main suit.   '
        ' PSR = likewise for RHO               '
        ' PSE = Pard's SS = either opp's suit. '
        ' -PSL = PSS <> LHO's suit.            '
        ' -PSR = likewise for RHO              '
        ' -PSE = PSS <> either opp's suit.     '
        If int_Pards_SS = 0 And Left$(s, 1) <> "-" Then
            Exit Function
        End If
        ' See if Pard's short suit LHO's suit: '
        l = (int_Pards_SS = int_LHOs_Suit(1))
        l = l Or (int_Pards_SS = int_LHOs_Suit(2))
        ' l is true if PSS = either of LHO's  '
        ' long suits, if any.                 '
        l0 = (int_LHOs_Suit(1) = 0)
        ' l0 is true if LHO has not bid a suit.

        ' See if Pard's short suit RHO's suit: '
        r = (int_Pards_SS = int_RHOs_Suit(1))
        r = r Or (int_Pards_SS = int_RHOs_Suit(2))
        ' r is true if PSS = either of RHO's  '
        ' long suits, if any.                 '
        r0 = (int_RHOs_Suit(1) = 0)
        ' r0 is true if RHO has not bid a suit.

        n = l Or r
        ' n is true if PSS = L or R           '

        ' DO THE COMPARISONS: '

        If Not (s = "PS<>L" And (l = False Or l0)) Then
            ' Spec is PSS <> LHO's long suit. '
            ' If L is true, get out:          '
            Exit Function

        ElseIf Not (s = "PS=L" And l = True) Then
            ' Spec is that PSS must = LHO. '
            Exit Function

        ElseIf Not (s = "PS<>R" And (r = False Or r0)) Then
            Exit Function

        ElseIf Not (s = "PS=R" And r = True) Then
            Exit Function

        ElseIf Not (s = "PS<>E" And (n = False Or (l0 And r0))) Then
            ' Either Pard's suit matches the opp's '
            ' or opp's don't have a suit.          '
            Exit Function

        ElseIf Not (s = "PS=E" And n = True) Then
            Exit Function

        End If

        int_SuitSpec = int_Pards_SS

    ElseIf s = "PH" Then ' Pard's Help suit. '
        int_SuitSpec = int_Pards_HelpSuit

    ElseIf s = "PN" Then ' Pard bid NT. '
        If int_BidNum Mod 4 = 3 Then
            ' We are 2 above opener = pard opened '
            If Left$(str_PriorBids, 2) = "1N" Then
                Eval_SuitSpecOk = True
                Exit Function
            End If
        End If

    ElseIf Left$(s, 1) = "A" Then
        If s = "A3" Then
            ' All HCPs are in 3 (or fewer) suits: '
            ok = (int_HCP(int_HCPn(1)) = 0)

        ElseIf s = "A2" Then
            ' All HCPs are in 2 (or fewer) suits: '
            ok = (int_HCP(int_HCPn(2)) = 0)

        ElseIf s = "A1" Then
            ' All HCPs are in 1 suit: '
            ok = (int_HCP(int_HCPn(4)) = int_HCP(0))

        Else
            MsgBox "Unrecognized Suit spec:" & vbCr & s
            Exit Function
        End If

        If ok Then
            If str_CardsSpec = "" And _
               str_LenSpec = "" _
            Then ' No other specs, so add value: '
                Eval_AddVal
            Else
                Eval_SuitSpecOk = True
            End If
        End If
        Exit Function

    ElseIf s = "ZSS" Then
        ' Short Suit has 0 HPCs: '
        Eval_SuitSpecOk = (int_HCP(int_Shpn(1)) = 0)
        Exit Function

    ElseIf s = "Clbs" Or s = "Dias" Or s = "Hrts" Or s = "Spds" Then
        int_SuitSpec = InStr("   ClbsDiasHrtsSpds", s) / 4
        Eval_SuitSpecOk = True
        Exit Function

    Else
        MsgBox "Suit spec not recognized:" & vbCr & s
    End If

    If int_SuitSpec > 0 Then
        Eval_SuitSpecOk = True

        If Left$(s, 1) = "-" Then
            int_SuitSpec = -int_SuitSpec
        End If
    End If
End Function


Private Function Eval_FitSpecOk() As Boolean
    Dim ok As Boolean
    Dim f As String
    Dim i As Integer

    If int_Pards_Suit(1) = 0 Or str_FitSpec = "" Then
        Eval_FitSpecOk = (str_FitSpec = "" Or str_FitSpec = "N")
        Exit Function
    End If

    If Left$(str_FitSpec & " ", 1) = "Y" Then
        f = Right$(str_FitSpec, 1)
        If f = "P" Then ' Must be partner's suit:'
            If int_Pards_SuitLen(int_Pards_Suit(1)) > 3 Then
                ok = True
            Else
                Exit Function
            End If
        Else
            i = InStr(" CDHS", f) - 1
            If i > 0 Then
                ' Suit specified.      '
                ' See if it was raised '
                ' ("F"= showed a Fit.) '
                If Right$(" " & t_Len(4 + i), 1) = "F" Then
                    ok = True
                End If

            ElseIf int_Fit_Suit(1) > 0 Then
                ok = True
            End If
        End If

    ElseIf str_FitSpec = "N" Then
        If int_Fit_Suit(1) = 0 Then
            ok = True
        End If

    Else
        MsgBox "Fit spec not recognized:" & vbCr & str_FitSpec
    End If

    If ok Then
        If str_CardsSpec = "" And _
           str_LenSpec = "" And _
           str_HandSpec = "" And _
           str_PosSpec = "" _
        Then ' No other specs, so add value: '
            Eval_AddVal
        Else ' Keep checking other specs: '
            Eval_FitSpecOk = True
        End If
    End If
End Function


Private Function Eval_PosSpecOk() As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim p As String
    Dim s As String

    If str_PosSpec = "" Then
        Eval_PosSpecOk = True
        Exit Function
    End If

    p = str_PosSpec
    i = Val(p)
    k = Len(p)

    If k = 1 Then
        ' Just a seat number if len(p)=1, '
        ' not a range of seat numbers nor '
        ' opener specified, etc. See docs.'
        If int_PosNum = i Then
            GoTo PosAddVal
        End If
        Exit Function
    End If

    If k > 2 Then
        j = Val(Mid$(p, 3, 1))
        If int_CallNum < i Or _
           int_CallNum > j _
        Then
            Exit Function
        End If
    End If

    ' We've established that position # is correct.'
    ' Now see if other specs, if any, match:       '

    If k > 3 Then
        ' e.g.: 2-4R means we are in chair 2-4 and '
        '       RHO opened, 2-4P means pard opened,'
        '       2-4L means LHO opened.             '
        s = Mid$(p, 4, 1)
        i = InStr(" RPL", s)
        If int_BidNum <> i Then
            ' e.g.: If p="2-4R", then the bid  '
            ' number must match "i" which is 2.'
            ' If int_BidNum is not 2, then RHO '
            ' didn't open, so exit:            '
            Exit Function
        ElseIf i < 2 Then
            MsgBox "Position spec not valid:" & vbCr & _
                   str_PosSpec & vbCr & _
                   "See the help file."
            Exit Function
        End If
    End If

    ' We've established that the specified player opened.
    If k > 4 Then
        ' e.g.: 2-4R-S means RHO opened, but not in spades.'
        s = Right$(p, 1)
        If InStr(" CDHS", s) < 2 Then
            MsgBox "Position spec suit not valid:" & vbCr & _
                   str_PosSpec & vbCr & _
                   "See the help file."
            Exit Function
        End If

        If Mid$(p, 5, 1) = "-" Then
            If s = Right$(str_Bid(1), 1) Then
                ' The minus sign means that the '
                ' specified suit must NOT have  '
                ' been bid. Since it WAS bid... '
                Exit Function
            End If
       Else ' Specified suit must be opened: '
            If s <> Right$(str_Bid(1), 1) Then
                ' Specified suit not opened. '
                Exit Function
            End If
        End If
    End If

PosAddVal:
    If str_CardsSpec = "" And _
       str_LenSpec = "" And _
       str_HandSpec = "" _
    Then ' No other specs, so add value: '
        Eval_AddVal
    Else ' Keep checking other specs: '
        Eval_PosSpecOk = True
    End If

End Function


Private Function Eval_HandSpecOk() As Boolean
    If str_HandSpec = "" Then
        Eval_HandSpecOk = True
    ElseIf Eval_HandsSpecCalcOk Then
        If str_CardsSpec = "" And _
           str_LenSpec = "" And str_SuitSpec = "" _
        Then ' No other specs, so add value: '
            If int_SuitSpec <> 0 Then
                If int_SuitNum <> int_SuitSpec Then
                    Exit Function
                End If
            Else
                int_SuitNum = int_SuitSpec
            End If
            Eval_AddVal
        Else ' Keep checking other specs: '
            Eval_HandSpecOk = True
        End If
    End If
End Function


Private Sub Eval_CardsSpecs()
    Dim j As Integer
    Dim k As Integer
    Dim n As Integer
    Dim oper As String
    Dim s As String
    Dim b As Boolean

    If str_CardsSpec = "" And str_LenSpec = "" Then
        Exit Sub
    End If

    For int_SuitNum = 1 To 4

        bln_ValAdded = False
        str_CardsSpecCombo = ""
        str_CardsSpec = DBGrid1.Columns(int_Col_Cards) 'Trim$("" & rs_HandValue!Cards)

        If int_SuitSpec > 0 Then
            ' Check only the specified suit: '
            int_SuitNum = int_SuitSpec

        ElseIf str_SuitSpec = "P" Then
            ' Check only partner's suit: '
            Do While int_SuitNum <> int_Pards_Suit(1)
                int_SuitNum = int_SuitNum + 1
                If int_SuitNum = 5 Then
                    int_SuitNum = 0
                    Exit Sub
                End If
            Loop

        ElseIf str_SuitSpec = "-P" Then ' NOT pard's suit '
            If int_SuitNum = int_Pards_Suit(1) Then
                GoTo EvalCards_GetNext
            End If

        ElseIf str_SuitSpec = "-M" Then ' Not my suit '
            If int_SuitNum = int_My_Suit(1) Then
                GoTo EvalCards_GetNext
            End If

        ElseIf str_SuitSpec = "E" Then ' Either opp's suit '
            For k = 1 To 2
                b = (int_LHOs_Suit(k) = int_SuitNum)
                b = b Or (int_RHOs_Suit(k) = int_SuitNum)
                If b Then Exit For
            Next
            If Not b Then
                GoTo EvalCards_GetNext
            End If

        ElseIf str_SuitSpec = "-E" Then ' NOT either opp's suit. '
            b = (int_LHOs_Suit(1) = int_SuitNum)
            b = b Or (int_RHOs_Suit(1) = int_SuitNum)
            If b Then
                GoTo EvalCards_GetNext
            End If

        ElseIf str_SuitSpec = "ER" Then
            For k = 1 To 2
                b = (int_LHOs_RebidSuit(k) = int_SuitNum)
                b = b Or (int_RHOs_RebidSuit(k) = int_SuitNum)
                If b Then Exit For
            Next
            If Not b Then
                GoTo EvalCards_GetNext
            End If

        ElseIf str_SuitSpec = "-F" Then
            For k = 1 To 3
                If int_Fit_Suit(k) = int_SuitNum Then
                    GoTo EvalCards_GetNext
                End If
            Next

        ElseIf Left$(str_CardsSpec & "    ", 4) = "HCP:" Then
            ' Suit HCPs specified. '
            j = Val(Mid$(str_SuitSpec & " ", 5))
            If int_HCP(int_SuitNum) < j Then
                'Eval_SuitLenSpec
                'Eval_AddVal
                GoTo EvalCards_GetNext
            End If

        End If

        ' When cards are specified: '
        Do While str_CardsSpec <> ""
            ' When more than one spec, such as '
            ' "??9, T9, 98", extract one spec  '
            ' at a time and when one passes the'
            ' test, skip the rest.             '

            Eval_ExtractCardSpecs
            Eval_SuitLenSpec

            If bln_ValAdded Then
                If str_CardsSpecCombo = "" Then
                    Exit Do
                Else
                    ' crds1 started off a something like  '
                    ' -Q & -J and "-J" was split to Combo '
                    ' and Extracted was set to "-Q", which'
                    ' has passed its test (though the value
                    ' was not really added). Next will see'
                    ' the next spec will also passes.     '
                    str_CardsSpec = str_CardsSpecCombo
                    str_CardsSpecCombo = ""
                End If
            End If
        Loop

        If str_LenSpec <> "" And _
           Trim$(DBGrid1.Columns(int_Col_Cards)) = "" _
        Then
            If Right$("  " & str_ValSpec, 2) = "*n" Then
                n = int_SuitLen(int_SuitNum)
                If DBGrid1.Columns(int_Col_Fit) = "Y" Or _
                   DBGrid1.Columns(int_Col_Suit) = "F" _
                Then ' Fit col.
                    n = n + int_Pards_SuitLen(int_Pards_Suit(1))
                End If
                Select Case Mid$(str_LenSpec & " ", 2, 1)
                Case Is = ">"
                    ' Only get value for # of cards '
                    ' greater than specified number.'
                    ' Example: n>4                  '
                    j = Val(Mid$(str_LenSpec & " ", 3))
                    n = n - j
                    If n < 0 Then n = 0

                Case Is = "+"
                    If Mid$(str_LenSpec & "  ", 3, 2) = "T>" Then
                        j = int_Pards_SuitLen(int_Pards_Suit(1))
                        n = n + j
                        j = Val(Mid$(str_LenSpec & " ", 5))
                        n = n - j
                        If n < 0 Then n = 0
                    End If
                End Select

                sng_ValSpec = Val(str_ValSpec) * n
                If sng_ValSpec > 0 Then
                    Eval_AddVal
                End If
            ElseIf Right$(" " & str_LenSpec, 1) = "+" Then
                If int_SuitLen(int_SuitNum) >= int_LenSpec Then
                    Eval_AddVal
                End If

            ElseIf int_SuitLen(int_SuitNum) = int_LenSpec Then
                Eval_AddVal
            End If
        End If

        If int_SuitSpec <> 0 Then
           ' Only check the specified suit, if any.'
           Exit For
        End If
EvalCards_GetNext:
    Next
    int_SuitNum = 0
End Sub


Private Sub Eval_ExtractCardSpecs()
    ' When more than one spec, such as '
    ' "??9, T9, 98", extract one spec  '
    ' at a time and when one passes the'
    ' test, skip the rest.             '
    Dim n As Integer

    n = InStr(str_CardsSpec, ",")
    If n > 1 Then
        str_CardsSpecExtracted = Trim$(Left$(str_CardsSpec, n - 1))
        str_CardsSpec = Trim$(Mid$(str_CardsSpec & " ", n + 1))
    Else
        str_CardsSpecExtracted = Trim$(str_CardsSpec)
        str_CardsSpec = ""
    End If

    n = InStr(str_CardsSpecExtracted, "&")
    If n > 1 Then
        str_CardsSpecCombo = _
            Trim$(Mid$(str_CardsSpecExtracted & " ", n + 1))

        str_CardsSpecExtracted = _
            Trim$(Left$(str_CardsSpecExtracted, n - 1))
    End If

    str_CardsSpecOperator = _
        Trim$(Left$(str_CardsSpecExtracted & " ", 1))

    If str_CardsSpecOperator <> "" Then
        If InStr(" AKQJT98", str_CardsSpecOperator) < 2 Then
            str_CardsSpecExtracted = Mid$(str_CardsSpecExtracted, 2)
        Else
            str_CardsSpecOperator = ""
        End If
    End If

End Sub


Private Sub Eval_SuitLenSpec()
    Dim n As Integer
    Dim j As Integer
    Dim notS As Boolean
    Dim X As String

    ' Below: Exit if the following is *NOT* true:      '
    '        length of current suit = specified length '
    '        -OR-                                      '
    '        length spec ends with "+" (e.g.: 4+) AND  '
    '        length of cur.suit < specified length     '
    '        -OR-                                      '
    '        length spec starts with < (e.g.: <9) AND  '
    '        length of cur.suit >= specified length    '
    '        -OR-                                      '
    '        specified length = "" (nothing specified) '
    '        -OR-                                      '
    '        specified length = "n", which is the suit '
    '        length.                                   '
    If Left$(str_LenSpec, 1) = "<" Then
        If int_SuitLen(int_SuitNum) >= int_LenSpec Then
            Exit Sub
        End If
    ElseIf Right$(str_LenSpec, 1) = "+" Then
        If int_SuitLen(int_SuitNum) < int_LenSpec Then
            Exit Sub
        End If
    ElseIf (int_SuitLen(int_SuitNum) <> int_LenSpec And _
            str_LenSpec <> "" And _
            str_LenSpec <> "n") _
    Then
        Exit Sub
    End If

    If Right$("  " & str_ValSpec, 2) = "*n" Then
        n = int_SuitLen(int_SuitNum)

        Select Case Mid$(str_LenSpec & " ", 2, 1)
        '                                       '
        Case Is = ">"
            ' Only get value for # of cards '
            ' greater than specified number.'
            ' Example: n>4                  '
            j = Val(Mid$(str_LenSpec & " ", 3))
            n = n - j
            If n < 0 Then n = 0

        Case Is = "<"
            ' Suit len must be less than the '
            ' specified amount (e.g.:  <4 ). '
            j = Val(Mid$(str_LenSpec & " ", 2))
            If n < j Then n = 1 ' multiply Value by 1

        Case Is = "+"
            If Mid$(str_LenSpec & "  ", 3, 2) = "P>" Then
                j = int_Pards_SuitLen(int_SuitNum)
                n = n + j
                j = Val(Mid$(str_LenSpec & " ", 5))
                n = n - j
                If n < 0 Then n = 0
            End If

        End Select
        '                                         '
        sng_ValSpec = Val(str_ValSpec) * n

    End If

    If str_CardsSpecOperator = "-" Then
        X = Left$(str_CardsSpecExtracted, 1)
        If X = "?" Then
            str_CardsSpecOperator = X
            X = "-"
        End If
    Else
        X = ""
    End If

    Select Case str_CardsSpecOperator
    Case Is = "*"
        If InStr(str_Suit(int_SuitNum), _
                 str_CardsSpecExtracted) > 0 _
        Then
            Eval_AddVal
        End If

    Case Is = "?"
        j = 1
        ' Can be ?K, ??Q, ???J, so find last ?. '
        ' (Remember that one ? has already been '
        ' stripped off to Operator.)            '
        Do While Mid$(str_CardsSpecExtracted, j, 1) = "?"
            j = j + 1
        Loop

        If j > Len(str_CardsSpecExtracted) Then
            ' spec is all ?s '
            X = Mid$(str_Suit(int_SuitNum), j, 1)
            If InStr(" AKQJT", X) > 1 Then
                Eval_AddVal
            End If

        ElseIf Mid$(str_Suit(int_SuitNum), _
                j + 1, _
                Len(str_CardsSpecExtracted) - j + 1) _
           = Mid$(str_CardsSpecExtracted, j) _
        Then
            If X <> "-" Then
                Eval_AddVal
            End If

        ElseIf X = "-" Then
            Eval_AddVal
        End If

    Case Is = "-"
        If InStr(str_Suit(int_SuitNum), _
                 str_CardsSpecExtracted) = 0 _
        Then
            Eval_AddVal
        End If

    Case Is = ""
        If Left$(str_Suit(int_SuitNum), _
                 Len(str_CardsSpecExtracted)) _
            = str_CardsSpecExtracted _
        Then
            If int_SuitLen(int_SuitNum) = int_LenSpec Or _
               str_LenSpec = "" Or InStr(str_LenSpec, "n") _
            Then
                Eval_AddVal

            ElseIf Right$(str_LenSpec, 1) = "+" Then
                If int_SuitLen(int_SuitNum) >= int_LenSpec Then
                    Eval_AddVal
                End If

            ElseIf Left$(str_LenSpec, 1) = "<" Then
                If int_SuitLen(int_SuitNum) < int_LenSpec Then
                    Eval_AddVal
                End If
            End If
        End If

    Case Else
        MsgBox "Invalid suit length spec:" & vbCr & str_LenSpec

    End Select

End Sub


Private Function Eval_ValSpecOk() As Boolean
    Dim v As Single
    
    bln_ValAdded = False
    
    If str_ValSpec = "L-S" Then
        v = int_SuitLen(int_Shpn(4)) - _
            int_SuitLen(int_Shpn(1))

    ElseIf str_ValSpec = "L1+L2" Then
        v = int_SuitLen(int_Shpn(4)) + _
            int_SuitLen(int_Shpn(3))

    ElseIf str_ValSpec = "T-S" Then
        ' Trump suit minus shortest suit '
        int_SuitNum = int_Pards_Suit(1)
        v = int_SuitLen(int_SuitNum) - _
            int_SuitLen(int_Shpn(1))
    End If

    If v <> 0 Then
        sng_ValSpec = v
        If str_PosSpec = "" And _
           str_HandSpec = "" And _
           str_FitSpec = "" And _
           str_SuitSpec = "" And _
           str_CardsSpec = "" And _
           str_LenSpec = "" _
        Then
            Eval_AddVal
        Else
            Eval_ValSpecOk = True
        End If
    Else
        Eval_ValSpecOk = True
    End If
End Function


Private Sub Eval_AddVal()
    Dim X As Boolean
    Dim h As Single

    'If ("" & rs_HandValue!Value) = "" And int_FlagNum > 0 Then
    If DBGrid1.Columns(int_Col_Value) = "" And int_FlagNum > 0 Then
        ' Counting cards or distribution '
        ' but not adding to hand value.  '
        int_Flag(int_FlagNum) = int_Flag(int_FlagNum) + 1
        'bln_ValAdded = True
        Exit Sub
    ElseIf sng_ValSpec <> 0 And _
       (int_SuitSpec = int_SuitNum Or _
        int_SuitSpec = 0) _
    Then
        If str_CardsSpecCombo = "" Then
            If InStr(str_ValSpec, "*HCP") Then
                h = int_HCP(int_SuitNum)
            Else
                h = 1
            End If
            Select Case DBGrid1.Columns(int_Col_NTSuit) ' rs_HandValue!NT_or_Suit
            Case Is = "S"
                sng_HandSuit_Val = sng_HandSuit_Val + _
                                   sng_ValSpec * h
            Case Is = "N"
                sng_HandNT_Val = sng_HandNT_Val + _
                                 sng_ValSpec * h

            Case Else
                sng_HandVal = sng_HandVal + _
                              sng_ValSpec * h
            End Select
        End If
        bln_ValAdded = True
        If bln_CountFlag Then
            int_Flag(int_FlagNum) = int_Flag(int_FlagNum) + 1
        End If
    Else
        bln_ValAdded = False
    End If

End Sub


Private Sub GetOtherPlayersData()
    Dim pb As String
    Dim b As String
    Dim m As Integer
    Dim n As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim s As String
    Dim c As String
    Dim q(4) As Integer
    Dim z(4) As Integer

    Erase int_LHOs_Suit()
    Erase int_RHOs_Suit()
    Erase int_Pards_Suit()
    Erase int_LHOs_SuitLen()
    Erase int_RHOs_SuitLen()
    Erase int_Pards_SuitLen()
    Erase int_LHOs_Fit()
    Erase int_RHOs_Fit()
    Erase int_Pards_Fit()
    Erase int_RHOs_RebidSuit()
    Erase int_LHOs_RebidSuit()
    Erase int_Pards_RebidSuit()
    Erase int_My_Suit()
    Erase int_My_Fit()
    Erase int_My_SS()
    Erase int_Fit_Suit()
    Erase int_Fit_Len()
    str_LHOs_Strength = ""
    str_RHOs_Strength = ""
    str_Pards_Strength = ""
    bln_Fit_Established = False
    bln_Fit_Confirmed = False
    int_RHOs_CueBid = 0
    int_LHOs_CueBid = 0
    int_Pards_CueBid = 0
    int_RHOs_HelpSuit = 0
    int_LHOs_HelpSuit = 0
    int_Pards_HelpSuit = 0
    int_LHOs_SS = 0
    int_RHOs_SS = 0
    int_Pards_SS = 0

    ' Get LHO's suits: '
    For i = 1 To 4
        q(i) = Val(t_Len(i))
        int_LHOs_SuitLen(i) = q(i)
        If t_Len(i) = "S" Then
            int_LHOs_SS = i
        ElseIf t_Len(i) = "C" Then
            int_LHOs_CueBid = i
        ElseIf t_Len(i) = "H" Then
            int_LHOs_HelpSuit = i
        ElseIf Right$(" " & t_Len(i), 1) = "F" Then
            ' Raised RHO's suit. '
            If int_Opps_FitSuit(1) = 0 Then
                int_Opps_FitSuit(1) = i
                int_LHOs_Fit(1) = i
            Else
                int_Opps_FitSuit(2) = i
                int_LHOs_Fit(2) = i
            End If
        ElseIf t_Len(i) = "R" Then
            If int_LHOs_RebidSuit(1) = 0 Then
                int_LHOs_RebidSuit(1) = i
            Else
                int_LHOs_RebidSuit(2) = i
            End If
        End If
    Next
    GoSub SortSuitLengths
    ' z(1) is the index number of the longest suit. '
    If q(z(1)) > 0 Then
        If q(z(1)) > 5 Then
            int_LHOs_RebidSuit(1) = z(1)
        ElseIf q(z(1)) > 3 Then
            int_LHOs_Suit(1) = z(1)
        End If
        If q(z(2)) > 3 Then
            int_LHOs_Suit(2) = z(2)
        End If
    End If
    str_LHOs_Strength = t_Str(0)
    bln_LHO_Bid_NT = (ck_NT(0).Value = 1)

    ' Get RHO's suits: '
    For i = 1 To 4
        j = 8 + i
        q(i) = Val(t_Len(j))
        int_RHOs_SuitLen(i) = q(i)
        If t_Len(j) = "S" Then
            int_RHOs_SS = i
        ElseIf t_Len(j) = "C" Then
            int_RHOs_CueBid = i
        ElseIf t_Len(j) = "H" Then
            int_RHOs_HelpSuit = i
        ElseIf Right$(" " & t_Len(j), 1) = "F" Then
            ' Raised LHO's suit. '
            If int_Opps_FitSuit(1) = 0 Then
                int_Opps_FitSuit(1) = i
                int_RHOs_Fit(1) = i
            Else
                int_Opps_FitSuit(2) = i
                int_RHOs_Fit(2) = i
            End If
        ElseIf t_Len(j) = "R" Then
            If int_RHOs_RebidSuit(1) = 0 Then
                int_RHOs_RebidSuit(1) = i
            Else
                int_RHOs_RebidSuit(2) = i
            End If
        End If
    Next
    GoSub SortSuitLengths
    If q(z(1)) > 0 Then
        If q(z(1)) > 5 Then
            int_RHOs_RebidSuit(1) = z(1)
        ElseIf q(z(1)) > 3 Then
            int_RHOs_Suit(1) = z(1)
        End If
        If q(z(2)) > 3 Then
            int_RHOs_Suit(2) = z(2)
        End If
    End If
    str_RHOs_Strength = t_Str(2)
    bln_RHO_Bid_NT = (ck_NT(2).Value = 1)

    ' Get Pard's suits: '
    For i = 1 To 4
        j = 4 + i
        q(i) = Val(t_Len(j))
        int_Pards_SuitLen(i) = q(i)
        If t_Len(j) = "S" Then
            int_Pards_SS = i
        ElseIf t_Len(j) = "C" Then
            int_Pards_CueBid = i
        ElseIf t_Len(j) = "H" Then
            int_Pards_HelpSuit = i
        ElseIf Right$(" " & t_Len(j), 1) = "F" Then
            ' Showed Fit for (raised) my suit. '
            For k = 1 To 3
                If int_Fit_Suit(k) = 0 Then
                    int_Fit_Suit(k) = i
                    int_Pards_Fit(k) = i
                    Exit For
                End If
            Next
            int_My_Suit(1) = i
        ElseIf t_Len(j) = "R" Then
            If int_Pards_RebidSuit(1) = 0 Then
                int_Pards_RebidSuit(1) = i
            Else
                int_Pards_RebidSuit(2) = 1
            End If
        End If
    Next
    GoSub SortSuitLengths
    If q(z(1)) > 0 Then
        If q(z(1)) > 5 Then
            int_Pards_RebidSuit(1) = z(1)
            int_Pards_Suit(1) = z(1)
        ElseIf q(z(1)) > 3 Then
            int_Pards_Suit(1) = z(1)
        End If
        If int_SuitLen(z(1)) + q(z(1)) > 7 Then
            int_Fit_Suit(1) = z(1)
        End If

        If q(z(2)) > 3 Then
            int_Pards_Suit(2) = z(2)
            If int_SuitLen(z(2)) + q(z(2)) > 7 Then
                int_Fit_Suit(2) = z(2)
            End If
        End If
    End If
    str_Pards_Strength = t_Str(1)
    bln_Pard_Bid_NT = (ck_NT(1).Value = 1)
    
    ' My data: '
    For i = 1 To 4
        q(i) = int_SuitLen(i)
        If t_Len(12 + i) <> "" Then
            Select Case t_Len(12 + i)
            Case "S"
                int_My_SS(1) = i
            Case "R"
                int_My_RebidSuit(1) = i
            Case "H"
                int_My_HelpSuit = i
            Case "F"
                int_My_Fit(1) = i
            Case Else
                If Val(t_Len(12 + i)) > 3 Then
                    int_My_Suit(1) = i
                End If
            End Select
        End If
    Next
    GoSub SortSuitLengths
    If int_My_SS(1) > 0 Then
        j = 2
    Else
        j = 1
    End If
    If int_SuitLen(z(4)) < 2 Then
        int_My_SS(j) = z(4)
        If int_SuitLen(z(3)) < 2 Then
            int_My_SS(j + 1) = z(3)
            If int_SuitLen(z(2)) < 2 Then
                int_My_SS(j + 2) = z(2)
            End If
        End If
    End If
    bln_I_Bid_NT = (ck_NT(3) = 1)
    bln_AlreadyHaveOPsData = True

Exit Sub


SortSuitLengths:
    For i = 1 To 4
        z(i) = i
    Next
    Do
        If q(z(2)) > q(z(1)) Then
            z(0) = z(2)
            z(2) = z(1)
            z(1) = z(0)
        ElseIf q(z(3)) > q(z(2)) Then
            z(0) = z(2)
            z(2) = z(3)
            z(3) = z(0)
        ElseIf q(z(4)) > q(z(3)) Then
            z(0) = z(4)
            z(4) = z(3)
            z(3) = z(0)
        Else
            Exit Do
        End If
    Loop

    ' Put higher-ranked suit first '
    ' when the same length suits:  '
    If z(2) > z(1) And q(z(2)) = q(z(1)) Then
        z(0) = z(2)
        z(2) = z(1)
        z(1) = z(0)
    End If
Return

End Sub


'                        Start hand analysis routines                          '


Private Sub AnalyzeHand()
    Calculate_BidNum
    ConvertHandStringToArrays
    ' TotalHCPs are calculated in the Sub  '
    '     ConvertHandStringToArrays, above.'
    ' DistribPts calc below in Calculate_Shape.'
    Calculate_Winners
    Calculate_LosingTrickCount
    Calculate_QuickTricks
    Calculate_Stoppers
    Calculate_Shape
    ' Intermed is calculated in the sub:   '
    '     ConvertHandStringToArray, above. '
    ' Vul. is specified by client.         '
    ' Suits: values used in Suits_Ok are   '
    ' calculated in other routines above.  '
End Sub


Private Sub Calculate_BidNum()
    Dim pb As String
    Dim pc As String
    Dim i As Integer
    Dim str_PriorCalls As String

    str_PriorCalls = t_Bidding

    If Left$(str_PriorCalls & "   ", 3) = "Ex:" Then
        str_PriorCalls = ""
    End If

    If str_PriorCalls = "" Then
        int_BidNum = 1
        int_CallNum = 1
        int_BidRound = 1
        int_PosNum = 1
        Exit Sub
    End If
    'Now there must be a Prior Call but not
    'necessarily a Prior Bid (e.g.: "P").

    str_PriorBids = ""
    int_BidNum = 0
    int_CallNum = 0
    int_BidRound = 0
    int_PosNum = 0

    pc = str_PriorCalls & "     "

    ' Count opening passes for number of '
    ' Prior Calls, but set Prior Bids to '
    ' ignore opening passes:             '
    If Left$(pc, 5) = "P-P-P" Then
        str_Call(1) = "P"
        str_Call(2) = "P"
        str_Call(2) = "P"
        int_CallNum = 3
        If Mid$(pc, 6, 1) = "-" Then
            str_PriorBids = Trim$(Mid$(pc, 7))
        End If

    ElseIf Left$(pc, 3) = "P-P" Then
        str_Call(1) = "P"
        str_Call(2) = "P"
        int_CallNum = 2
        If Mid$(pc, 4, 1) = "-" Then
            str_PriorBids = Trim$(Mid$(pc, 5))
        End If

    ElseIf Left$(pc, 1) = "P" Then
        str_Call(1) = "P"
        int_CallNum = 1
        If Mid$(pc, 2, 1) = "-" Then
            str_PriorBids = Trim$(Mid$(pc, 3))
        End If

    Else
        str_PriorBids = str_PriorCalls
    End If

    pb = str_PriorBids

    'Count bids after opening passes:
    i = InStr(pb, "-")
    Do While i > 0
        int_CallNum = int_CallNum + 1
        str_Call(int_CallNum) = Left$(pb, i - 1)
        int_BidNum = int_BidNum + 1
        str_Bid(int_BidNum) = str_Call(int_CallNum)
        pb = Mid$(pb, i + 1)
        i = InStr(pb, "-")
    Loop
    If pb <> "" Then
        'Always one more bid than dashes:
        int_CallNum = int_CallNum + 1
        str_Call(int_CallNum) = Trim$(pb)
        int_BidNum = int_BidNum + 1
        str_Bid(int_BidNum) = str_Call(int_CallNum)
    End If
    ' int_CallNum & BidNum point to last '
    ' Call/Bid. Want them to point to    '
    ' the current call/bid, so add 1 more.
    int_CallNum = int_CallNum + 1
    int_BidNum = int_BidNum + 1


    ' Compute bidding round #: '
    int_BidRound = int_BidNum \ 4 'e.g.: 3\4=0, 4\4=1, 5\4=1
    If int_BidNum / 4 <> int_BidNum \ 4 Then
        ' Not divisible evenly by 4, so add 1: '
        int_BidRound = int_BidRound + 1
    End If

    ' Compute seat/position #: '
    int_PosNum = int_BidNum Mod 4
    If int_PosNum = 0 Then int_PosNum = 4

    ' Compute LHO's, Pard's and RHO's positions: '
    Select Case int_PosNum
    Case 1
        int_LHOs_Pos = 2
        int_Pards_Pos = 3
        int_RHOs_Pos = 4
    Case 2
        int_LHOs_Pos = 3
        int_Pards_Pos = 4
        int_RHOs_Pos = 1
    Case 3
        int_LHOs_Pos = 4
        int_Pards_Pos = 1
        int_RHOs_Pos = 2
    Case 4
        int_LHOs_Pos = 1
        int_Pards_Pos = 2
        int_RHOs_Pos = 3
    End Select
End Sub


Private Sub ConvertHandStringToArrays()
    '*************************************************
    ' 1. Converts hand (e.g.: AK432-Q74-K74-32) into *
    '      1a. suit array "bln_Suit(suit, card)" and *
    '      1b. suit strings "str_Suit(suit)".        *
    ' 2. Verifies accuracy of hand's card values     *
    '      (e.g.:  AJ52 not AJ25, 13 cards, etc.).   *
    ' 3. Counts High Card Points for the hand and    *
    '      for each suit, where A=4, K=3, Q=2, J=1.  *
    ' 4. Counts Intermediates (8-10).                *
    '*************************************************

    Dim strCard As String
    Dim s As Integer
    Dim intCardPtr As Integer
    Dim intCardNum As Integer
    Dim intLastCard As Integer
    Dim str_SuitName(4) As String
    'Outputs:
    Erase str_Suit() ' Substring of hand string by suit '
    Erase bln_Suit() ' T/F for each card held, by suit. '
    Erase int_HCP()  ' High card points.                '
    Erase int_SuitLen()  ' Suit length.                 '
    Erase int_Intermediates() ' # Intermed. by suit/hand'


    str_SuitName(1) = "Clubs"
    str_SuitName(2) = "Diamonds"
    str_SuitName(3) = "Hearts"
    str_SuitName(4) = "Spades"

    s = Spds  'Hands from Client start with Spades.

    intLastCard = 99 'Each card value must be < the last,
                     'so start with a high number.

    For intCardPtr = 1 To 16 'Each card (and dash) in the hand.

        Do While Mid$(str_PlayersHand, intCardPtr, 1) = "-"
            ' Suit is finished when "-" reached.          '
            ' If void in Spades, hand may start with "-". '
            ' If void in H/D, two dashes are together,    '
            ' such as AK432--K853-QJT2, so Do-Loop is used.
            If intCardPtr = 16 Then
                'If hand ends with "-", it is void in Clubs
                Exit For
            Else
                'Set variables for next suit:
                s = s - 1
                intLastCard = 99
                intCardPtr = intCardPtr + 1 'Get card after dash
            End If
        Loop

        strCard = Mid$(str_PlayersHand, intCardPtr, 1)
        intCardNum = InStr(" 23456789TJQKA", strCard)
        ' Check if card is valid: '
        If intCardNum < 2 Then
            'Character in the hand is not 2-A.
            'Return error code and message:
            MsgBox strCard & " in " & _
                    str_SuitName(s) & _
                    " in " & str_PlayersHand & _
                    " is not a valid card."
            Exit Sub

        ElseIf intCardNum >= intLastCard Then
            'If card number is greater than the
            'last card,  it is out of sequence.
            'Return error code and message:
            MsgBox _
                "The " & strCard & " of " _
                & str_SuitName(s) _
                & " is not in sequence in " _
                & Chr$(34) & str_Suit(s) _
                & strCard & Chr$(34)
            Exit Sub

        Else
            'Good card, so add to HCP if an honor,
            'add to Intermediates if 8-10,
            'and add to Suit array:
            int_HCP(s) = int_HCP(s) + _
                         InStr("JQKA", strCard)

            ' Count intermediates: '
            If InStr("9T", strCard) Then
                int_Intermediates(s) = _
                    int_Intermediates(s) + 1
            ElseIf strCard = "8" And _
                   (bln_Suit(s, 9) Or _
                    bln_Suit(s, 10)) _
            Then ' Only count 8 if also 9/10.'
                int_Intermediates(s) = _
                    int_Intermediates(s) + 1
            End If

            ' Check off card in suit array: '
            bln_Suit(s, intCardNum) = True
            ' Add card to suit string:      '
            str_Suit(s) = str_Suit(s) & strCard
            'Update LastCard number: '
            intLastCard = intCardNum
        End If
    Next

    'Total up HCPs and  '
    '  Intermediates    '
    '  for whole hand:  '
    For s = 1 To 4
        int_HCP(0) = int_HCP(0) + _
                     int_HCP(s)

        int_SuitLen(s) = Len(str_Suit(s))

        int_Intermediates(0) _
            = int_Intermediates(0) _
            + int_Intermediates(s)
    Next

End Sub


Private Sub Calculate_Winners()
    '**************************************************
    ' 1. Count number of winners in this hand.        *
    '                                                 *
    ' The number of winners in a hand is based on the *
    ' first four cards in a suit, plus one winner     *
    ' counted for each card more than four in a suit. *
    '   AKQJ=4, AKQT=3.5, AKJT=3.5, AQJT=3.5, AKQ=3   *
    '   KQJT=3, AQJ=2.5, AK=2, KQJ=2, QJTx=2, AQ=1.5  *
    '   A=1, KQ=1, QJT=1, Kx=.5                       *
    '                                                 *
    ' 2. Count number of running winners.             *
    '                                                 *
    ' Running Winners are cards which can be cashed   *
    ' without giving up the lead. AK=2, AQ=1, Kx=0    *
    ' See notes below for counting long running suits.*
    '**************************************************

    Dim s As Integer

    sng_NumWinners(0) = 0
    sng_NumRunners(0) = 0

    For s = 1 To 4
        sng_NumWinners(s) = 0
        sng_NumRunners(s) = 0

        If Left$(str_Suit(s), 4) = "AKQJ" Then
            'Additional top cards are not counted as winners
            'because they are counted in length, below.
            sng_NumWinners(s) = 4

            ' May need to specify length > 5 '
            sng_NumRunners(s) = int_SuitLen(s)

        ElseIf Left$(str_Suit(s), 3) = "AKQ" Then
            sng_NumWinners(s) = 3

            If int_SuitLen(s) > 6 Then
                ' Runs if 6 out are not 4-1-1 or '
                ' worse and then only if pard is '
                ' not the one with the 4+ and if '
                ' the Jack is in the long hand.  '
                sng_NumRunners(s) = int_SuitLen(s)
            Else
                sng_NumRunners(s) = sng_NumRunners(s) + 3
            End If

        ElseIf Left$(str_Suit(s), 2) = "AK" Then

            If int_SuitLen(s) > 7 Then ' 5 out not 3-1-1, etc.
                sng_NumRunners(s) = int_SuitLen(s)
            Else
                sng_NumRunners(s) = 2
            End If

            If Left$(str_Suit(s), 4) = "AKJT" Or _
               Left$(str_Suit(s), 4) = "AKQT" _
            Then
                sng_NumWinners(s) = sng_NumWinners(s) + 3.5
                ' Actually 3.67, since LHO will have
                ' the Q only 1/3 of the time.

            ElseIf Left$(str_Suit(s), 3) = "AKJ" Then
                sng_NumWinners(s) = sng_NumWinners(s) + 2.5

            Else
                sng_NumWinners(s) = sng_NumWinners(s) + 2

            End If

        ElseIf Left$(str_Suit(s), 1) = "A" Then

            If int_SuitLen(s) > 10 Then ' 2 out not 2-0-0
                sng_NumRunners(s) = int_SuitLen(s)
            Else
                sng_NumRunners(s) = 1
            End If

            If Left$(str_Suit(s), 4) = "AQJT" Then
                sng_NumWinners(s) = sng_NumWinners(s) + 3.5

            ElseIf Left$(str_Suit(s), 3) = "AQJ" Then
                sng_NumWinners(s) = sng_NumWinners(s) + 2.5

            ElseIf Left$(str_Suit(s), 2) = "AQ" Then
                sng_NumWinners(s) = sng_NumWinners(s) + 1.5

            ElseIf Left$(str_Suit(s), 3) = "AJT" Then
                sng_NumWinners(s) = sng_NumWinners(s) + 2
                '75% chance one finesse will win.

            Else
                sng_NumWinners(s) = sng_NumWinners(s) + 1

            End If

        ElseIf Left$(str_Suit(s), 4) = "KQJT" Then
            sng_NumWinners(s) = 3

        ElseIf Left$(str_Suit(s), 3) = "KQJ" Then
            sng_NumWinners(s) = 2

        ElseIf Left$(str_Suit(s), 2) = "KQ" Then
            sng_NumWinners(s) = 1

        ElseIf Left$(str_Suit(s), 1) = "K" And _
               int_SuitLen(s) > 1 Then
            sng_NumWinners(s) = 0.5

        ElseIf Left$(str_Suit(s), 3) = "QJT" Then
            If int_SuitLen(s) > 3 Then
                sng_NumWinners(s) = 2
            Else
                sng_NumWinners(s) = 1
            End If
        End If

        'give credit for long suits:
        If int_SuitLen(s) > 4 Then
            sng_NumWinners(s) _
                = sng_NumWinners(s) _
                + int_SuitLen(s) _
                - 4
        End If

        sng_NumWinners(0) _
            = sng_NumWinners(0) _
            + sng_NumWinners(s)

        sng_NumRunners(0) _
            = sng_NumRunners(0) _
            + sng_NumRunners(s)

    Next

End Sub


Private Sub Calculate_LosingTrickCount()
    '***************************************************
    ' Calculate "Losing Trick Count" for the hand:     *
    '     Only top 3 cards in a suit are considered,   *
    '     so a suit can have no more than 3 losers.    *
    '     In top 3 cards, cards below Q are losers:    *
    '       AKx = 1, AQx = 1, AJx = 2, KQx=1, Qxx=2    *
    '     In a doubleton, cards below K are losers:    *
    '       AQ=1, Ax=1, KQ=1, Kx=1, QJ=2, Qx=2         *
    '     In a singleton, any card but A is a loser:   *
    '       K=1, Q=1, x=1                              *
    '***************************************************
    '   The above is per www.PhillipAdlerBridge.com    *
    '***************************************************
    ' Note that "LTC" does not always consider that    *
    ' more tricks may be won with a finesse, so AQT is *
    ' considered 2 losers just like A32. Yet Kx is not *
    ' considered 2 losers. This does not seem logical, *
    ' but we stick with the standard anyway.           *
    '***************************************************

    Dim s As Integer
    Erase int_NumLosers()

    For s = 1 To 4
        If Left$(str_Suit(s), 3) = "AKQ" Then
            'No losers in top 3 cards.

        ElseIf Left$(str_Suit(s), 2) = "AK" Then
            If int_SuitLen(s) > 2 Then 'AKx.
                int_NumLosers(s) = int_NumLosers(s) + 1
            End If

        ElseIf Left$(str_Suit(s), 2) = "AQ" Then
            int_NumLosers(s) = int_NumLosers(s) + 1 'AQ only or A32-AQJ.

        ElseIf Left$(str_Suit(s), 1) = "A" Then
            If int_SuitLen(s) > 2 Then 'A32 to AJT.
                int_NumLosers(s) = int_NumLosers(s) + 2

            ElseIf int_SuitLen(s) = 2 Then 'A2 to AQ.
                int_NumLosers(s) = int_NumLosers(s) + 1

            End If

        ElseIf Left$(str_Suit(s), 3) = "KQJ" Then
            int_NumLosers(s) = int_NumLosers(s) + 1 'KQJ only.

        ElseIf Left$(str_Suit(s), 2) = "KQ" Then
            If int_SuitLen(s) > 2 Then 'KQx to KQT.
                int_NumLosers(s) = int_NumLosers(s) + 2

            ElseIf int_SuitLen(s) = 2 Then 'KQ only.
                int_NumLosers(s) = int_NumLosers(s) + 1

            End If

        ElseIf Left$(str_Suit(s), 1) = "K" Then
            If int_SuitLen(s) > 2 Then 'K32 to KJT.
                int_NumLosers(s) = int_NumLosers(s) + 2

            ElseIf int_SuitLen(s) = 2 Then 'K2 to KJ.
                int_NumLosers(s) = int_NumLosers(s) + 1

            End If

        ElseIf Left$(str_Suit(s), 1) = "Q" Then
            If int_SuitLen(s) > 2 Then 'Q32 to QJT.
                int_NumLosers(s) = int_NumLosers(s) + 2
            Else                   'Q only = 1,  Q2 to QJ = 2.
                int_NumLosers(s) = int_NumLosers(s) + _
                                  int_SuitLen(s)
            End If

        Else 'none of the top 3 honors
            If int_SuitLen(s) > 3 Then '432 to JT9
                int_NumLosers(s) = int_NumLosers(s) + 3
            Else           '2 only to K only = 1,  32 to QJ = 2.
                int_NumLosers(s) = int_NumLosers(s) + _
                                  int_SuitLen(s)
            End If
        End If

        int_NumLosers(0) = int_NumLosers(0) + _
                           int_NumLosers(s)

    Next

End Sub


Private Sub Calculate_QuickTricks()
    '********************************************
    ' Count number of Quick Tricks in the hand  *
    '      ("quick" defensive winners).         *
    '                                           *
    '   AK = 2 tricks                           *
    '   AQ = 1 1/2 tricks                       *
    '   A  = 1 trick                            *
    '   KQ = 1 trick                            *
    '   Kx = 1/2 trick                          *
    '                                           *
    ' The phrase "quick tricks" is sometimes    *
    ' used to refer to tricks which can be      *
    ' cashed without giving up the lead, which  *
    ' is clearly not the case with the above.   *
    ' What they are referring to is what we     *
    ' call "running winners" to avoid confusion *
    ' with the phrase "quick tricks" as it is   *
    ' properly used to refer to hand valuation  *
    ' that derived from methods used eons ago   *
    ' by Ely Culbertson before 4-3-2-1 counting.*
    '********************************************

    Dim s As Integer
    Dim adj As Single
    ' If a suit has 6+ cards, or if partner has '
    ' bid a suit in which this hand has length, '
    ' then it is possible that the bidder has a '
    ' singleton (or even a void) in the suit,   '
    ' so an AQ is probably worth only 1 Quick   '
    ' and AK, maybe 1.5. In fact, if your side  '
    ' is known to have a 10-card fit, these     '
    ' holdings may not even be worth 1 QT, but  '
    ' we have not addressed that yet.           '

    sng_QuickTricks(0) = 0

    For s = Clbs To Spds

        If int_SuitLen(s) > 5 Then
            adj = 0.5
        Else
            adj = 0
        End If

        If Left$(str_Suit(s), 2) = "AK" Then
            sng_QuickTricks(s) = 2 - adj

        ElseIf Left$(str_Suit(s), 2) = "AQ" Then
            sng_QuickTricks(s) = 1.5 - adj

        ElseIf Left$(str_Suit(s), 1) = "A" Then
            sng_QuickTricks(s) = 1

        ElseIf Left$(str_Suit(s), 2) = "KQ" Then
            sng_QuickTricks(s) = 1 - adj

        ElseIf Left$(str_Suit(s), 1) = "K" And _
               int_SuitLen(s) > 1 _
        Then
            sng_QuickTricks(s) = 0.5

        End If

        sng_QuickTricks(0) = sng_QuickTricks(0) + _
                             sng_QuickTricks(s)

    Next
End Sub


Private Sub Calculate_Stoppers()
    '******************************************************'
    ' The Bridge World web site defines a "Stopper" as --  '
    '   "a holding that will (or is likely to, or might)   '
    '    prevent the opponents from immediately running    '
    '    a large number of tricks in a suit at notrump."   '
    '                                                      '
    ' For example, K2 in declarer's hand is always a       '
    ' stopper, but may not be in dummy's hand. If his RHO  '
    ' bid the suit, then it is probably a stopper, but if  '
    ' dummy's LHO bid the suit, then the opening lead thru '
    ' the K2 will probably make the K fall to LHO's Ace.   '
    '                                                      '
    ' In this routine, Kx is always considered a stopper,  '
    ' no matter who bid what. (In reality, LHO may lead a  '
    ' different suit to partner who leads back through the '
    ' Kx, in which case it is NOT a stopper, but for the   '
    ' purposes of calculating stoppers, that is ignored.)  '
    '                                                      '
    ' To count as 2 stoppers, a holding in declarer's hand '
    ' must withstand the opening lead from LHO plus a      '
    ' subsequent lead by RHO. Example:  AQ2 = 2 stoppers   '
    ' because the Q will take the first trick (or the A,   '
    ' if RHO plays the K). KQ2 is not 2 stoppers because a '
    ' low card can force the Q, then RHO leads thru the K. '
    '                                                      '
    ' QT9 = no stoppers. If LHO has AJ8 and RHO has Kx and '
    ' leads thru the QT, but QT9x is a stopper after LHO   '
    ' gets the lead back. QT87 is a stopper according to ML.
    '                                                      '
    ' Q98x is 1 stopper.  If LHO leads low and RHO takes K '
    ' leads back low, you can cover with the 8, then after '
    ' LHO takes with the T, your Q9 holds off the AJ.      '
    '******************************************************'

    Dim i As Integer
    Dim s As Integer
    Dim c2 As String
    Dim c3 As String
    Dim c4 As String
    Dim pard As Integer

    sng_NumStops(0) = 0 ' NOTE: '
    ' sng_NumStops() for each suit is the total number of  '
    ' stoppers in that suit, whereas sng_NumStops(0) is    '
    ' the total number of SUITS stopped, excluding pard's. '

    ' Get partner's suit: '
    If int_BidNum > 2 Then
        If int_Pards_Suit(1) = 99 Then
            GetOtherPlayersData
            pard = int_Pards_Suit(1)
        End If
       If 1 = 2 Then
        i = int_PosNum
        If i < 3 Then i = i + 4
        s = i - 2 ' pard's position #
        If s < 1 Then s = s + 3
        If str_Call(s) = "P" Then s = s + 4

        int_Pards_Pos = s

        If Val(str_Call(s)) > 0 Then
            ' Must have bid some number of something. '
            s = InStr("CDHS", Mid$(str_Call(s), 2))
            pard = s
        End If
       End If
    End If



    For s = 1 To 4
        sng_NumStops(s) = 0

        c2 = Left$(str_Suit(s) & "  ", 2)
        c3 = Left$(str_Suit(s) & "   ", 3)
        c4 = Left$(str_Suit(s) & "    ", 4)

        If bln_Suit(s, Ace) Then
            If c2 = "AK" Or c2 = "AQ" Then
                'AK or AQ
                sng_NumStops(s) = 2 'Max=2

            ElseIf (c3 = "AJT" Or c3 = "AJ9") And _
                   int_SuitLen(s) > 3 _
            Then
                'AJTx
                sng_NumStops(s) = 2

            ElseIf c4 = "AT98" And _
                   int_SuitLen(s) > 4 _
            Then
                'AT98x
                sng_NumStops(s) = 2

            Else 'A
                sng_NumStops(s) = 1
            End If

        ElseIf bln_Suit(s, King) And _
               int_SuitLen(s) > 1 _
        Then
            If c2 = "KQ" Then
                ' KQ... '
                If c3 = "KQJ" Or c3 = "KQT" Then
                    ' KQJ... or KQT...'
                    sng_NumStops(s) = 2

                ElseIf c3 = "KQ9" And _
                       int_SuitLen(s) > 3 _
                Then
                    ' KQ9x -- J-Q, x-9-T, Ax vs Kx = 2
                    sng_NumStops(s) = 2

                Else ' KQx or less
                    If int_LHOs_Suit(1) = s Then
                        sng_NumStops(s) = 1
                    Else
                        sng_NumStops(s) = 2
                    End If
                End If

            ElseIf bln_Suit(s, Jack) Then
                ' KJ... '
                If int_SuitLen(s) > 3 Then
                    If bln_Suit(s, Ten) Then
                        'KJTx -- LHO leads to RHO's and h's cards,  '
                        '             then RHO leads (now or later) '
                        '                   LHO:   h:               '
                        '        A-x, x-T-Q, 9x vs KJ = 2 or        '
                        '        Q-K, x-T-A, 9x vs Jx = 2 or        '
                        '        x-T, x-J-Q, Ax vs Kx = 2           '
                        sng_NumStops(s) = 2

                    ElseIf bln_Suit(s, Nine) Or bln_Suit(s, 8) Then
                        ' KJ9x -- Q-K, x-9-T, Ax vs Jx               '
                        sng_NumStops(s) = 2

                    Else ' KJxx+
                        sng_NumStops(s) = 1
                    End If

                Else ' KJx or KJ
                    If int_LHOs_Suit(1) = s Or _
                       int_SuitLen(s) < 3 _
                    Then
                        sng_NumStops(s) = 1
                    Else
                        sng_NumStops(s) = 2
                    End If
                End If

            ElseIf bln_Suit(s, Ten) And int_SuitLen(s) > 3 Then
                ' KT9... or KT8...                   '
                ' LHO leads low. Say RHO plays Q and '
                ' we play K, leaving AJ9 over T8x.   '
                ' If RHO gets in and leads low again,'
                ' we play 8 and LHO plays 9. AJ takes'
                ' our last 2. Only 1 stop, but for   '
                ' this to be a problem, LHO must have'
                ' most of the high cards and length  '
                ' and RHO must get in twice. The odds'
                ' of this are sufficiently low that  '
                ' we are going to count KT9x and KT8x'
                ' as 2 stoppers. I would prefer it as'
                ' 1.5 stops, but that's not allowed. '
                If bln_Suit(s, 9) Or bln_Suit(s, 8) Then
                    sng_NumStops(s) = 2

                Else 'KTxx+
                    sng_NumStops(s) = 1

                End If

            Else 'Kx+
                sng_NumStops(s) = 1
            End If

        ElseIf bln_Suit(s, Queen) And _
               int_SuitLen(s) > 2 _
        Then
            ' Q??... '
            If bln_Suit(s, Jack) Then
                ' QJx...
                sng_NumStops(s) = 1

            ElseIf bln_Suit(s, Ten) And int_SuitLen(s) > 3 Then
                ' QTxx '
                If bln_Suit(s, 9) Then
                    ' QT9x '
                    If int_LHOs_Suit(1) = s Then
                        sng_NumStops(s) = 1
                    Else
                        sng_NumStops(s) = 2
                    End If
                End If

                If bln_Suit(s, 8) And bln_Suit(s, 7) Then
                    ' QT87 '
                    If int_LHOs_Suit(1) = s Then
                        sng_NumStops(s) = 1
                    Else
                        sng_NumStops(s) = 2
                    End If
                End If

                ' In his book Overcalls, Mike Lawrence says (on '
                ' page 102, hand 3) that QT7 is a stopper, but  '
                ' if LHO has AJ8xx and RHO has Kx, QT7 is not a '
                ' stopper, so I'm going against ML this time.   '

                ' Later: In one of his columns, Bobby Wolff says'
                ' that QT75 is a stopper. I give up:            '

                If bln_Suit(s, 7) Then
                    ' QT7x '
                    sng_NumStops(s) = 1
                End If


            ElseIf bln_Suit(s, Nine) And _
                   int_SuitLen(s) > 3 _
            Then
                ' Q9xx...
                If bln_Suit(s, Ten) Then
                    'QT9x RHO takes A, leads x-9-J, Kx vs Qx
                    sng_NumStops(s) = 1

                ElseIf bln_Suit(s, Eight) Then
                    'Q98x RHO A, leads x-8-T, KJ vs Q9
                    sng_NumStops(s) = 1
                End If

            Else ' Qxx '
                If int_LHOs_Suit(1) = s Then
                    sng_NumStops(s) = 0
                Else
                    sng_NumStops(s) = 1
                End If
            End If

        ElseIf bln_Suit(s, Jack) And _
               int_SuitLen(s) > 3 _
        Then
            ' J???... '
            If bln_Suit(s, Ten) And _
               (bln_Suit(s, Nine) Or bln_Suit(s, 8)) _
            Then
                'JT9x or JT8x  A takes x, K-T, Q-8/9
                sng_NumStops(s) = 1

            ElseIf c3 = "J98" And _
                   int_SuitLen(s) > 4 _
            Then
                'J98xx  A-x, RHO leads x-8-T, K-x, Qx vs Jx
                sng_NumStops(s) = 1

            Else
                sng_NumStops(s) = 0.5
            End If

        ElseIf (c3 = "T98" Or c3 = "T97") And _
               int_SuitLen(s) > 4 _
        Then
            'T98xx -- Ax, 2-8-J, KQx vs T9x
            sng_NumStops(s) = 1
        End If


        ' Total number of suits stopped '
        ' outside of partner's 1st suit.'
        If sng_NumStops(s) > 0 And s <> pard Then
            If sng_NumStops(s) = 0.5 Then
                sng_NumStops(0) = sng_NumStops(0) + 0.5
            Else
                sng_NumStops(0) = sng_NumStops(0) + 1
            End If
        End If

    Next

End Sub


Private Sub Calculate_Shape()
    '************************************************
    ' Shape in BidBase can be one of the following: *
    '    1 = 1-suiter (6+/3-)                       *
    '    2 = 2-suiter (6+/4+, 5+/5+)                *
    '    3 = 3-suiter (4441, 5440)                  *
    '    b = balanced (4333, 4432, 5332)            *
    '    u = unbalanced (no singleton or void)      *
    '    s = singleton                              *
    '    v = void                                   *
    '    e = either sing-void                       *
    ' But a hand can be more than one of the above, *
    ' such as "1" and "u" and "e", and we don't     *
    ' know in advance which an entry will want.     *
    '                                               *
    ' So to make comparison of the hand to BidBase  *
    ' possible, the hand's features are stored in   *
    ' two different variables.                      *
    '                                               *
    ' The first 4 are in the str_Shape variable.    *
    ' The last 3 are in the str_Shortness variable. *
    ' Unbalanced is not assigned. Instead, if entry *
    ' specifies Unbalanced, we just check to make   *
    ' sure that str_Shape is not "b".               *
    '************************************************

    Dim s As Integer
    Dim n4 As Integer  ' # of 4-card suits
    Dim n5 As Integer  ' # of 5-card suits
    Dim n6 As Integer  ' # of 6+card suits
    Dim s1 As Integer  ' singleton  '
    Dim db As Integer  ' doubleton
    Dim v As Integer   ' void       '

    Erase int_DistribPts()

    str_Shape = ""     ' 1, 2, 3, b '
    str_Shortness = "" ' s, v, d     '

    ' Count suits, shortness: '
    For s = 1 To 4
        If int_SuitLen(s) > 5 Then
            n6 = n6 + 1

        ElseIf int_SuitLen(s) = 5 Then
            n5 = n5 + 1

        ElseIf int_SuitLen(s) = 4 Then
            n4 = n4 + 1

        ElseIf int_SuitLen(s) = 2 Then
            int_DistribPts(s) = int_DistribPts(s) + 1
            db = db + 1
            
        ElseIf int_SuitLen(s) = 1 Then
            s1 = s1 + 1
            int_DistribPts(s) = int_DistribPts(s) + 2

        ElseIf int_SuitLen(s) = 0 Then
            v = v + 1
            int_DistribPts(s) = int_DistribPts(s) + 3

        End If

        int_DistribPts(0) _
            = int_DistribPts(0) _
            + int_DistribPts(s)

    Next

    ' Record num of suits to 'str_Shape':  '
    If n6 = 1 Then
        If n5 = 1 Or n4 = 1 Then
            ' 6-5 or 6-4 or more '
            str_Shape = "2"

        Else ' 6-3-2-2, 7-3-2-1, etc.
            str_Shape = "1"
        End If

    ElseIf n5 = 2 Or n6 = 2 Or (n5 = 1 And n6 = 1) Then
        ' 5-5 or 6-6+ or 5-6+ '
        str_Shape = "2"

    ElseIf n5 + n4 = 3 Then
        ' 5-4-4-0 or 4-4-4-1 '
        str_Shape = "3"

    ElseIf n6 = 0 And int_DistribPts(0) < 2 Then
        ' 5-3-3-2, 4-4-3-2, 4-3-3-3 '

        str_Shape = "b"

    End If

    ' Record shortness to 'str_Shortness':  '
    If v > 0 Then
        ' We have a void.
        If s1 > 0 Then
            str_Shortness = "b" ' both
        Else
            str_Shortness = "v"
        End If

    ElseIf s1 > 0 Then
        ' We have a singleton.'
        str_Shortness = "s"

    ElseIf db > 0 Then
        str_Shortness = "d"

    End If

End Sub


Private Sub t_Len_GotFocus(Index As Integer)
    L_Info.Top = 50
    L_Info = "For suit length, enter the expected minimum " & _
             "number of cards for the suit(s) indicated by " & _
             "the bidding, or check NT for a 1NT opening." & vbCr & _
             "If the player is showing a Fit for partner's " & _
             "suit, append an 'F' to the number of cards shown, " & _
             "such as '4F'." & vbCr & _
             "Single letter codes instead of numbers:   H=Help suit.    " & _
             "R=Rebiddable suit.    S=Singleton/void.    C=Cue bid A/K."
    t_Len(Index).SelStart = 0
    t_Len(Index).SelLength = 32768
End Sub


Private Sub t_Len_KeyDown(Index As Integer, _
                          KeyCode As Integer, _
                          Shift As Integer)

    If Shift <> 0 Then Exit Sub

    Select Case KeyCode
        Case Is = 38
            Select Case Index
            Case Is = 1, 2, 3, 5, 6, 7, 9, 10, 11
                t_Len(Index + 1).SetFocus
            Case Is = 4, 8, 12
                t_Str(Index / 4 - 1).SetFocus
            End Select

        Case Is = 40
            Select Case Index
            Case Is = 2, 3, 4, 6, 7, 8, 10, 11, 12
                t_Len(Index - 1).SetFocus
            End Select
            
    End Select
End Sub


Private Sub t_Str_GotFocus(Index As Integer)
    L_Info.Top = 50
    L_Info = "Enter the closest approximation of the expected " & _
             "minimum strength of the hand where:" & vbCr & _
             "B = Bust (0-5)        W = Weak (6-10)        " & _
             "I = Intermediate(10-12)     O = Open (13-15) " & vbCr & _
             "P = Plus (15-17)     S = Strong (18-19)     " & _
             "V = Very Strong (20-24)     X = eXtreme (25+)"
    t_Str(Index).SelStart = 0
    t_Str(Index).SelLength = 32768
End Sub


Private Sub t_Str_KeyDown(Index As Integer, _
                          KeyCode As Integer, _
                          Shift As Integer)
    If Shift <> 0 Then Exit Sub

    If KeyCode = 40 Then ' down
        t_Len((Index + 1) * 4).SetFocus
    End If
End Sub


Private Sub t_Str_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim k As String
    
    k = Chr$(KeyAscii)
    If k >= "a" And k <= "z" Then
        KeyAscii = KeyAscii - 32
        k = UCase$(k)
    End If
    If InStr("BWIOPSVX", k) = 0 And k >= "0" Then
        MsgBox "Incorrect code entered (" & k & ")." & vbCr & _
               "See the help box at the bottom of the screen."
    End If
End Sub


Private Sub ClearVP()
    ' Clears the VPs from the Grid which '
    ' were calculated at the last Compute.
    Dim a As String
    Dim n As String

    DBGrid1.row = 0
    n = Trim$(DBGrid1.Columns(1))

    If Not bln_NeedsClearing Then Exit Sub

    With rs_HandVal
        If str_LastSystem = "1.All" Then
            .Index = "PrimaryKey"
            .MoveFirst
            If .NoMatch Or .EOF Then
                Exit Sub
            End If
            Do While Not .EOF
                If !VP <> "" Then
                    .Edit
                    !VP = ""
                    .Update
                End If
                .MoveNext
            Loop

        ElseIf n <> "" Then
            .Index = "Name"
            .Seek "=", n
            If .NoMatch Or .EOF Then
                Exit Sub
            End If
            Do While !Name = n
                If !VP <> "" Then
                    .Edit
                    !VP = ""
                    .Update
                End If
                .MoveNext
                If .EOF Then Exit Do
            Loop
        End If
    End With
    bln_NeedsClearing = False
End Sub


Private Sub ScrollToGridTop()
    Dim a As String
    On Local Error GoTo ScrollError
    With DBGrid1
        .col = 0
        .row = 0
        Do
            a = .Text
            .Scroll 0, -1
            .row = 0
        Loop While .Text <> a
    End With
Exit Sub

ScrollError:
    On Local Error GoTo 0

End Sub
