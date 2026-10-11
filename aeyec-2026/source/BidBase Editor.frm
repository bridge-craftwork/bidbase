VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form BBEditor 
   BackColor       =   &H00000000&
   Caption         =   "CardShark BidBase Editor"
   ClientHeight    =   13080
   ClientLeft      =   225
   ClientTop       =   2085
   ClientWidth     =   27855
   FillColor       =   &H80000012&
   Icon            =   "BidBase Editor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   13080
   ScaleWidth      =   27855
   Begin VB.Frame f_EnterHand 
      BackColor       =   &H80000005&
      Caption         =   "Find Bid For Hand:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6675
      Left            =   14640
      TabIndex        =   69
      Top             =   6195
      Visible         =   0   'False
      Width           =   3080
      Begin VB.CommandButton b_ShowRefList 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Index           =   2
         Left            =   2690
         Picture         =   "BidBase Editor.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   306
         Top             =   2985
         Width           =   290
      End
      Begin VB.TextBox t_Find_Ref 
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
         Height          =   390
         Left            =   75
         TabIndex        =   304
         Top             =   2970
         Width           =   2900
      End
      Begin VB.CommandButton b_ShowFindHandList 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   2690
         Picture         =   "BidBase Editor.frx":2F64
         Style           =   1  'Graphical
         TabIndex        =   179
         Top             =   540
         Width           =   290
      End
      Begin VB.ListBox cmb_FindRef 
         BackColor       =   &H00C0FFFF&
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
         Height          =   1740
         ItemData        =   "BidBase Editor.frx":5BBE
         Left            =   90
         List            =   "BidBase Editor.frx":5BD7
         Sorted          =   -1  'True
         TabIndex        =   305
         Top             =   3375
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.ListBox cmb_FindBid 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         ItemData        =   "BidBase Editor.frx":5C43
         Left            =   75
         List            =   "BidBase Editor.frx":5C5C
         TabIndex        =   74
         Top             =   915
         Visible         =   0   'False
         Width           =   2910
      End
      Begin VB.ComboBox cmb_FindPos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "BidBase Editor.frx":5CC8
         Left            =   1120
         List            =   "BidBase Editor.frx":5CED
         TabIndex        =   301
         Top             =   3500
         Width           =   1815
      End
      Begin VB.CommandButton b_DelFindHand 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   2690
         TabIndex        =   277
         ToolTipText     =   "Delete hand from list."
         Top             =   925
         Visible         =   0   'False
         Width           =   290
      End
      Begin VB.CommandButton b_DelFindHand 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   2690
         TabIndex        =   178
         ToolTipText     =   "Delete hand from list."
         Top             =   540
         Visible         =   0   'False
         Width           =   290
      End
      Begin VB.CommandButton b_ShowFindHandList 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   2690
         Picture         =   "BidBase Editor.frx":5D97
         Style           =   1  'Graphical
         TabIndex        =   276
         Top             =   925
         Width           =   290
      End
      Begin VB.TextBox t_FindHand 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   75
         TabIndex        =   72
         Top             =   900
         Width           =   2910
      End
      Begin VB.CheckBox ck_Unopposed 
         BackColor       =   &H00008000&
         Caption         =   "&Unopposed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   1380
         TabIndex        =   188
         Top             =   1370
         Width           =   1455
      End
      Begin VB.Frame f_Find2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2620
         Left            =   120
         TabIndex        =   283
         Top             =   3900
         Width           =   2900
         Begin VB.CommandButton b_ClearBidFound 
            Caption         =   "X"
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
            Left            =   0
            TabIndex        =   303
            ToolTipText     =   "Delete hand from list."
            Top             =   1580
            Width           =   290
         End
         Begin VB.CommandButton b_EditFind 
            Caption         =   "Add Bid"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Index           =   1
            Left            =   1700
            TabIndex        =   290
            Top             =   1180
            Width           =   1100
         End
         Begin VB.CommandButton b_EditFind 
            Caption         =   "Edit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Index           =   0
            Left            =   2250
            TabIndex        =   298
            Top             =   1180
            Width           =   550
         End
         Begin VB.CommandButton b_CancelFind 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   2020
            TabIndex        =   291
            Top             =   2260
            Width           =   800
         End
         Begin VB.CommandButton b_ClearFind 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   0
            TabIndex        =   289
            Top             =   2260
            Width           =   800
         End
         Begin VB.ComboBox cmb_Find_ScoreMeth 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "BidBase Editor.frx":89F1
            Left            =   1850
            List            =   "BidBase Editor.frx":89FE
            TabIndex        =   288
            Top             =   375
            Width           =   945
         End
         Begin VB.TextBox t_FindAnalysis4Bid 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2420
            TabIndex        =   287
            ToolTipText     =   "Bid to analyze:"
            Top             =   780
            Width           =   375
         End
         Begin VB.ComboBox cmb_FindVul 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "BidBase Editor.frx":8A12
            Left            =   1600
            List            =   "BidBase Editor.frx":8A2B
            TabIndex        =   286
            Top             =   0
            Width           =   1200
         End
         Begin VB.CommandButton b_FindBidForHand 
            Caption         =   "Find Bid"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   0
            TabIndex        =   285
            Top             =   1180
            Width           =   1100
         End
         Begin VB.CommandButton b_Import 
            Caption         =   "Import"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   1010
            TabIndex        =   284
            Top             =   2260
            Width           =   800
         End
         Begin VB.Label L_FindBidID 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Caption         =   "ID#: ---       Bid Order #: ---"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   0
            TabIndex        =   297
            Top             =   1910
            Width           =   2800
         End
         Begin VB.Label L_FindBid 
            BackStyle       =   0  'Transparent
            Caption         =   "Scoring Method:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   296
            Top             =   420
            Width           =   1890
         End
         Begin VB.Label L_FindBid 
            BackStyle       =   0  'Transparent
            Caption         =   "Show analysis for bid:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   0
            TabIndex        =   295
            Top             =   840
            Width           =   2415
         End
         Begin VB.Label L_BidFound 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   2160
            TabIndex        =   294
            Top             =   1575
            Width           =   360
         End
         Begin VB.Label L_FindBid 
            BackStyle       =   0  'Transparent
            Caption         =   "Computer's Bid:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   345
            TabIndex        =   293
            Top             =   1605
            Width           =   1695
         End
         Begin VB.Label L_FindBid 
            BackStyle       =   0  'Transparent
            Caption         =   "Vulnerability:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   292
            Top             =   15
            Width           =   1605
         End
      End
      Begin VB.CommandButton b_EnterHandHelp 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   0
         Left            =   1820
         TabIndex        =   278
         Top             =   210
         Width           =   290
      End
      Begin VB.TextBox t_FindPrior 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Left            =   75
         MultiLine       =   -1  'True
         TabIndex        =   73
         Top             =   1600
         Width           =   2910
      End
      Begin VB.CommandButton b_LoadHands 
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   2700
         TabIndex        =   181
         Top             =   210
         Width           =   290
      End
      Begin VB.CommandButton b_SaveHands 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   2400
         TabIndex        =   180
         Top             =   210
         Width           =   290
      End
      Begin VB.TextBox t_FindHand 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   420
         Index           =   0
         Left            =   75
         TabIndex        =   71
         Top             =   510
         Width           =   2910
      End
      Begin VB.CommandButton b_ClearFindHands 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   2110
         TabIndex        =   282
         Top             =   210
         Width           =   290
      End
      Begin VB.Label L_FindBid 
         BackStyle       =   0  'Transparent
         Caption         =   "Position:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   105
         TabIndex        =   302
         Top             =   3510
         Width           =   990
      End
      Begin VB.Label L_FindBid 
         BackStyle       =   0  'Transparent
         Caption         =   "Prior Bids:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   75
         Top             =   1395
         Width           =   1275
      End
      Begin VB.Label L_FindBid 
         BackStyle       =   0  'Transparent
         Caption         =   "Spd-Hrt-Dia-Clb"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   1620
      End
   End
   Begin VB.Frame f_ShowSort 
      Caption         =   "Entries To Show/Sort On:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4845
      Left            =   3150
      TabIndex        =   76
      Top             =   2865
      Visible         =   0   'False
      Width           =   4590
      Begin VB.CommandButton b_ViewGrid2 
         Caption         =   "Vie&w"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   1300
         TabIndex        =   268
         Top             =   2120
         Width           =   875
      End
      Begin VB.CommandButton b_ShowAllSorted 
         Caption         =   "All Entries"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   170
         ToolTipText     =   "Display specified groups of entries."
         Top             =   3960
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.CommandButton b_ShowSortDefaults 
         Caption         =   "Defaults"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3345
         TabIndex        =   85
         ToolTipText     =   "Display specified groups of entries."
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox t_Show 
         Height          =   285
         Index           =   5
         Left            =   885
         MaxLength       =   100
         TabIndex        =   84
         Top             =   4350
         Width           =   1365
      End
      Begin VB.TextBox t_Show 
         Height          =   285
         Index           =   4
         Left            =   885
         MaxLength       =   100
         TabIndex        =   83
         Top             =   3945
         Width           =   1365
      End
      Begin VB.TextBox t_Show 
         Height          =   285
         Index           =   2
         Left            =   885
         MaxLength       =   100
         TabIndex        =   80
         Top             =   2550
         Width           =   1365
      End
      Begin VB.TextBox t_Show 
         Height          =   285
         Index           =   3
         Left            =   885
         MaxLength       =   100
         TabIndex        =   81
         Top             =   3000
         Width           =   1365
      End
      Begin VB.ComboBox cmb_Show 
         Height          =   315
         Index           =   1
         Left            =   585
         TabIndex        =   82
         Text            =   "Prior_Bids"
         Top             =   3360
         Width           =   1665
      End
      Begin VB.ComboBox cmb_Show 
         Height          =   315
         Index           =   0
         ItemData        =   "BidBase Editor.frx":8A64
         Left            =   585
         List            =   "BidBase Editor.frx":8A66
         TabIndex        =   79
         Text            =   "Prior_Bids"
         Top             =   1515
         Width           =   1665
      End
      Begin VB.ComboBox cmb_Sort 
         Height          =   315
         Index           =   4
         ItemData        =   "BidBase Editor.frx":8A68
         Left            =   2805
         List            =   "BidBase Editor.frx":8A90
         TabIndex        =   90
         Text            =   "Pct_Used"
         Top             =   3330
         Width           =   1545
      End
      Begin VB.ComboBox cmb_Sort 
         Height          =   315
         Index           =   3
         ItemData        =   "BidBase Editor.frx":8B06
         Left            =   2805
         List            =   "BidBase Editor.frx":8B2E
         TabIndex        =   89
         Text            =   "Bid_Order_Number"
         Top             =   2775
         Width           =   1545
      End
      Begin VB.ComboBox cmb_Sort 
         Height          =   315
         Index           =   2
         ItemData        =   "BidBase Editor.frx":8BA4
         Left            =   2805
         List            =   "BidBase Editor.frx":8BCC
         TabIndex        =   88
         Text            =   "Prior_Bids"
         Top             =   2130
         Width           =   1545
      End
      Begin VB.ComboBox cmb_Sort 
         Height          =   315
         Index           =   1
         ItemData        =   "BidBase Editor.frx":8C42
         Left            =   2805
         List            =   "BidBase Editor.frx":8C6A
         TabIndex        =   87
         Text            =   "Player"
         Top             =   1530
         Width           =   1545
      End
      Begin VB.ComboBox cmb_Sort 
         Height          =   315
         Index           =   0
         ItemData        =   "BidBase Editor.frx":8CE0
         Left            =   2805
         List            =   "BidBase Editor.frx":8D08
         TabIndex        =   86
         Text            =   "Round"
         Top             =   945
         Width           =   1545
      End
      Begin VB.TextBox t_Show 
         Height          =   285
         Index           =   1
         Left            =   885
         MaxLength       =   100
         TabIndex        =   78
         Top             =   1125
         Width           =   1365
      End
      Begin VB.CommandButton b_Show 
         Caption         =   "&View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   1100
         TabIndex        =   92
         ToolTipText     =   "Display specified groups of entries."
         Top             =   375
         Width           =   915
      End
      Begin VB.TextBox t_Show 
         Height          =   285
         Index           =   0
         Left            =   885
         MaxLength       =   100
         TabIndex        =   77
         Top             =   750
         Width           =   1365
      End
      Begin VB.CommandButton b_Close_Show 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3100
         TabIndex        =   94
         Top             =   4335
         Width           =   915
      End
      Begin VB.Label L_Show 
         BackStyle       =   0  'Transparent
         Caption         =   "Grid 4:"
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
         Index           =   15
         Left            =   150
         TabIndex        =   127
         Top             =   4380
         Width           =   645
      End
      Begin VB.Label L_Show 
         BackStyle       =   0  'Transparent
         Caption         =   "Grid 3:"
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
         Left            =   150
         TabIndex        =   126
         Top             =   3990
         Width           =   645
      End
      Begin VB.Label L_Show 
         BackStyle       =   0  'Transparent
         Caption         =   "IN:"
         Height          =   195
         Index           =   7
         Left            =   285
         TabIndex        =   117
         Top             =   3405
         Width           =   285
      End
      Begin VB.Label L_Show 
         BackStyle       =   0  'Transparent
         Caption         =   "FROM:"
         Height          =   195
         Index           =   5
         Left            =   285
         TabIndex        =   116
         Top             =   2595
         Width           =   645
      End
      Begin VB.Label L_Show 
         BackStyle       =   0  'Transparent
         Caption         =   "TO:"
         Height          =   195
         Index           =   6
         Left            =   285
         TabIndex        =   115
         Top             =   3045
         Width           =   330
      End
      Begin VB.Label L_Show 
         BackStyle       =   0  'Transparent
         Caption         =   "IN:"
         Height          =   195
         Index           =   3
         Left            =   285
         TabIndex        =   114
         Top             =   1560
         Width           =   285
      End
      Begin VB.Label L_Show 
         BackStyle       =   0  'Transparent
         Caption         =   "Fifth:"
         Height          =   195
         Index           =   13
         Left            =   2805
         TabIndex        =   102
         Top             =   3150
         Width           =   645
      End
      Begin VB.Label L_Show 
         BackStyle       =   0  'Transparent
         Caption         =   "Fourth:"
         Height          =   195
         Index           =   12
         Left            =   2805
         TabIndex        =   101
         Top             =   2565
         Width           =   645
      End
      Begin VB.Label L_Show 
         BackStyle       =   0  'Transparent
         Caption         =   "Third:"
         Height          =   195
         Index           =   11
         Left            =   2805
         TabIndex        =   100
         Top             =   1935
         Width           =   645
      End
      Begin VB.Label L_Show 
         BackStyle       =   0  'Transparent
         Caption         =   "Second:"
         Height          =   195
         Index           =   10
         Left            =   2805
         TabIndex        =   99
         Top             =   1335
         Width           =   645
      End
      Begin VB.Label L_Show 
         BackStyle       =   0  'Transparent
         Caption         =   "First:"
         Height          =   195
         Index           =   9
         Left            =   2820
         TabIndex        =   98
         Top             =   720
         Width           =   645
      End
      Begin VB.Label L_Show 
         BackStyle       =   0  'Transparent
         Caption         =   "Sort On:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   8
         Left            =   2535
         TabIndex        =   97
         Top             =   405
         Width           =   945
      End
      Begin VB.Line Line2 
         X1              =   2415
         X2              =   2415
         Y1              =   510
         Y2              =   4600
      End
      Begin VB.Label L_Show 
         BackStyle       =   0  'Transparent
         Caption         =   "Bottom Grid:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   150
         TabIndex        =   96
         Top             =   2175
         Width           =   1440
      End
      Begin VB.Label L_Show 
         BackStyle       =   0  'Transparent
         Caption         =   "TO:"
         Height          =   195
         Index           =   2
         Left            =   285
         TabIndex        =   95
         Top             =   1170
         Width           =   330
      End
      Begin VB.Label L_Show 
         BackStyle       =   0  'Transparent
         Caption         =   "FROM:"
         Height          =   195
         Index           =   1
         Left            =   285
         TabIndex        =   93
         Top             =   780
         Width           =   645
      End
      Begin VB.Label L_Show 
         BackStyle       =   0  'Transparent
         Caption         =   "Top Grid:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   91
         Top             =   405
         Width           =   945
      End
   End
   Begin VB.Frame f_OtherPlayers 
      BackColor       =   &H00008000&
      Height          =   4010
      Left            =   18480
      TabIndex        =   273
      Top             =   7560
      Width           =   5000
      Begin VB.TextBox t_OtherHands 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000040&
         Height          =   285
         Left            =   3930
         TabIndex        =   274
         ToolTipText     =   "Dbl-Clk to clear grid."
         Top             =   530
         Visible         =   0   'False
         Width           =   990
      End
      Begin MSFlexGridLib.MSFlexGrid g_OtherHands 
         Height          =   3990
         Left            =   0
         TabIndex        =   275
         ToolTipText     =   "Click cell to edit."
         Top             =   0
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   7038
         _Version        =   393216
         Rows            =   19
         Cols            =   5
         BackColor       =   0
         BackColorFixed  =   32768
         ForeColorFixed  =   65535
         ScrollBars      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame f_CompareGrid 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H00FFFFFF&
      Height          =   6650
      Left            =   8745
      TabIndex        =   174
      Top             =   1335
      Visible         =   0   'False
      Width           =   4880
      Begin MSFlexGridLib.MSFlexGrid g_CompareGrid 
         Height          =   6390
         Left            =   0
         TabIndex        =   175
         Top             =   300
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   11271
         _Version        =   393216
         Rows            =   26
         Cols            =   3
         FixedCols       =   0
         ScrollBars      =   0
         MergeCells      =   1
      End
      Begin VB.Label L_CompareGridExit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   250
         Left            =   4550
         TabIndex        =   177
         Top             =   30
         Width           =   270
      End
      Begin VB.Label L_CompareGrid 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Compare Entries"
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
         Height          =   285
         Left            =   2040
         TabIndex        =   176
         Top             =   660
         Width           =   4545
      End
   End
   Begin VB.Frame f_Msg 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2565
      Left            =   7890
      TabIndex        =   182
      Top             =   5370
      Visible         =   0   'False
      Width           =   5460
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
         Left            =   1800
         TabIndex        =   265
         Top             =   1450
         Visible         =   0   'False
         Width           =   1140
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
         Left            =   1000
         TabIndex        =   264
         Top             =   1450
         Visible         =   0   'False
         Width           =   650
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
         Left            =   200
         TabIndex        =   186
         Top             =   1900
         Width           =   650
      End
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
         Left            =   1000
         TabIndex        =   185
         Text            =   "MsgInput"
         Top             =   1900
         Visible         =   0   'False
         Width           =   4100
      End
      Begin VB.Label L_MsgBox 
         BackStyle       =   0  'Transparent
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
         Left            =   195
         TabIndex        =   184
         Top             =   480
         Width           =   5040
      End
      Begin VB.Label L_Caption 
         BackColor       =   &H000000FF&
         Caption         =   "  Message Box"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   150
         TabIndex        =   183
         Top             =   -30
         Width           =   6405
      End
      Begin VB.Label L_MsgBackground 
         BackColor       =   &H00FFFFFF&
         Height          =   2130
         Left            =   60
         TabIndex        =   187
         Top             =   360
         Width           =   5340
      End
   End
   Begin VB.Frame f_Unhide 
      Caption         =   "Unhide Columns"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Left            =   9990
      TabIndex        =   103
      Top             =   3495
      Visible         =   0   'False
      Width           =   3120
      Begin VB.CommandButton b_Close_Hidden 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1695
         TabIndex        =   109
         Top             =   2985
         Width           =   915
      End
      Begin VB.ComboBox cmb_HiddenCols 
         Height          =   315
         Index           =   0
         ItemData        =   "BidBase Editor.frx":8D7E
         Left            =   450
         List            =   "BidBase Editor.frx":8D80
         Sorted          =   -1  'True
         TabIndex        =   104
         Top             =   595
         Width           =   2175
      End
      Begin VB.CommandButton b_Unhide 
         Caption         =   "Unhide"
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
         Left            =   465
         TabIndex        =   108
         Top             =   2985
         Width           =   840
      End
      Begin VB.ComboBox cmb_HiddenCols 
         Height          =   315
         Index           =   1
         Left            =   450
         Sorted          =   -1  'True
         TabIndex        =   105
         Top             =   1215
         Width           =   2175
      End
      Begin VB.ComboBox cmb_HiddenCols 
         Height          =   315
         Index           =   2
         Left            =   450
         Sorted          =   -1  'True
         TabIndex        =   106
         Top             =   1800
         Width           =   2175
      End
      Begin VB.ComboBox cmb_HiddenCols 
         Height          =   315
         Index           =   3
         Left            =   450
         Sorted          =   -1  'True
         TabIndex        =   107
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label L_Unhide 
         BackStyle       =   0  'Transparent
         Caption         =   "Bottom-Right"
         Height          =   195
         Index           =   3
         Left            =   450
         TabIndex        =   113
         Top             =   2205
         Width           =   1545
      End
      Begin VB.Label L_Unhide 
         BackStyle       =   0  'Transparent
         Caption         =   "Bottom-Left:"
         Height          =   195
         Index           =   2
         Left            =   450
         TabIndex        =   112
         Top             =   1585
         Width           =   1545
      End
      Begin VB.Label L_Unhide 
         BackStyle       =   0  'Transparent
         Caption         =   "Top-Right:"
         Height          =   195
         Index           =   1
         Left            =   450
         TabIndex        =   111
         Top             =   965
         Width           =   1545
      End
      Begin VB.Label L_Unhide 
         BackStyle       =   0  'Transparent
         Caption         =   "Top-Left:"
         Height          =   195
         Index           =   0
         Left            =   450
         TabIndex        =   110
         Top             =   345
         Width           =   1365
      End
   End
   Begin VB.Frame frame0 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   4920
      Index           =   6
      Left            =   0
      TabIndex        =   189
      Top             =   8040
      Width           =   27885
      Begin VB.ListBox list_Refs 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         ItemData        =   "BidBase Editor.frx":8D82
         Left            =   10905
         List            =   "BidBase Editor.frx":8D84
         Sorted          =   -1  'True
         TabIndex        =   48
         Top             =   3650
         Visible         =   0   'False
         Width           =   3540
      End
      Begin VB.ListBox L_Status 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         IntegralHeight  =   0   'False
         ItemData        =   "BidBase Editor.frx":8D86
         Left            =   0
         List            =   "BidBase Editor.frx":8D88
         TabIndex        =   190
         Top             =   3750
         Width           =   19425
      End
      Begin VB.Frame frame0 
         Caption         =   "Bid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1100
         Index           =   0
         Left            =   0
         TabIndex        =   198
         Top             =   0
         Width           =   14595
         Begin VB.CheckBox ck_Continuation 
            BackColor       =   &H0000C000&
            Caption         =   "Continuation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Left            =   8445
            TabIndex        =   205
            Top             =   885
            Width           =   1575
         End
         Begin VB.TextBox t_Order 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5375
            MaxLength       =   6
            TabIndex        =   7
            ToolTipText     =   "Suit grouping number."
            Top             =   450
            Width           =   915
         End
         Begin VB.TextBox t_Player 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1305
            MaxLength       =   1
            TabIndex        =   5
            ToolTipText     =   "Player number, starting with first non-passer."
            Top             =   450
            Width           =   240
         End
         Begin VB.TextBox t_Round 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1020
            MaxLength       =   1
            TabIndex        =   4
            Text            =   "1"
            ToolTipText     =   "Bidding round number, starting with first non-passing call."
            Top             =   450
            Width           =   240
         End
         Begin VB.ComboBox cmb_Bid 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "BidBase Editor.frx":8D8A
            Left            =   6240
            List            =   "BidBase Editor.frx":8E18
            TabIndex        =   8
            ToolTipText     =   "Bid to be made when hand matches this entry's specs."
            Top             =   450
            Width           =   650
         End
         Begin VB.TextBox t_Prior 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1620
            MaxLength       =   50
            TabIndex        =   6
            ToolTipText     =   "Starting with first non-passing bid. Example: 1C-P-1H-P"
            Top             =   450
            Width           =   3765
         End
         Begin VB.TextBox t_Pct 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   694
            MaxLength       =   3
            TabIndex        =   3
            Text            =   "45"
            ToolTipText     =   "% of time to use this bid instead of alternative."
            Top             =   450
            Width           =   285
         End
         Begin VB.TextBox t_Sel 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   90
            TabIndex        =   1
            ToolTipText     =   "W = We use this convention, T = They, blank = We and They use the same convention. 0 = neither."
            Top             =   450
            Width           =   240
         End
         Begin VB.CommandButton b_IncrSeq 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5970
            TabIndex        =   204
            ToolTipText     =   " Increment the Bid Order Number"
            Top             =   225
            Width           =   180
         End
         Begin VB.CheckBox ck_Saved 
            Caption         =   "Saved"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6990
            TabIndex        =   10
            ToolTipText     =   "Entry has been saved when checked."
            Top             =   420
            Width           =   1155
         End
         Begin VB.CheckBox ck_SaveNew 
            Caption         =   "New Entry"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   6975
            TabIndex        =   9
            ToolTipText     =   "Save as a new entry when checked."
            Top             =   165
            Width           =   1290
         End
         Begin VB.PictureBox b_Google 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   11160
            Picture         =   "BidBase Editor.frx":8ECF
            ScaleHeight     =   285
            ScaleWidth      =   255
            TabIndex        =   203
            Top             =   165
            Width           =   255
         End
         Begin VB.PictureBox b_Conven_Description 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   10860
            Picture         =   "BidBase Editor.frx":92ED
            ScaleHeight     =   285
            ScaleWidth      =   255
            TabIndex        =   202
            Top             =   165
            Width           =   255
         End
         Begin VB.ComboBox cmb_ConvName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            ItemData        =   "BidBase Editor.frx":98AB
            Left            =   8175
            List            =   "BidBase Editor.frx":98AD
            Sorted          =   -1  'True
            TabIndex        =   11
            Top             =   450
            Width           =   3315
         End
         Begin VB.ComboBox cmb_ConvName 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            ItemData        =   "BidBase Editor.frx":98AF
            Left            =   11625
            List            =   "BidBase Editor.frx":98B1
            Sorted          =   -1  'True
            TabIndex        =   201
            Top             =   450
            Width           =   2805
         End
         Begin VB.TextBox t_Sub 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   393
            TabIndex        =   2
            ToolTipText     =   "W = only We use this sub-convention, T = only They, [blank] = both, 0 = neither."
            Top             =   450
            Width           =   240
         End
         Begin VB.CommandButton b_SaveEntry 
            BackColor       =   &H000000FF&
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2800
            TabIndex        =   200
            Top             =   165
            Width           =   1000
         End
         Begin VB.CommandButton b_ClearInputs 
            BackColor       =   &H000000FF&
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4000
            TabIndex        =   199
            Top             =   165
            Width           =   1000
         End
         Begin VB.Label L_Bid 
            BackStyle       =   0  'Transparent
            Caption         =   "P"
            Height          =   225
            Index           =   3
            Left            =   1380
            TabIndex        =   217
            Top             =   225
            Width           =   135
         End
         Begin VB.Label L_Bid 
            BackStyle       =   0  'Transparent
            Caption         =   "Rd"
            Height          =   225
            Index           =   2
            Left            =   1050
            TabIndex        =   216
            Top             =   225
            Width           =   225
         End
         Begin VB.Label L_Bid 
            BackStyle       =   0  'Transparent
            Caption         =   "Bid"
            Height          =   225
            Index           =   6
            Left            =   6320
            TabIndex        =   215
            Top             =   225
            Width           =   315
         End
         Begin VB.Label L_Bid 
            BackStyle       =   0  'Transparent
            Caption         =   "Prior Bids"
            Height          =   225
            Index           =   4
            Left            =   1620
            TabIndex        =   214
            Top             =   210
            Width           =   810
         End
         Begin VB.Label L_Bid 
            BackStyle       =   0  'Transparent
            Caption         =   "Pct."
            Height          =   225
            Index           =   1
            Left            =   720
            TabIndex        =   213
            Top             =   225
            Width           =   330
         End
         Begin VB.Label L_Bid 
            BackStyle       =   0  'Transparent
            Caption         =   "Cnv"
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   212
            Top             =   225
            Width           =   300
         End
         Begin VB.Label L_Bid 
            BackStyle       =   0  'Transparent
            Caption         =   "Order:"
            Height          =   225
            Index           =   5
            Left            =   5475
            TabIndex        =   211
            Top             =   225
            Width           =   465
         End
         Begin VB.Label L_Date_Saved 
            BackStyle       =   0  'Transparent
            Caption         =   "date saved"
            Height          =   210
            Left            =   7170
            TabIndex        =   210
            Top             =   660
            Width           =   1005
         End
         Begin VB.Label L_Bid 
            BackStyle       =   0  'Transparent
            Caption         =   "Bid/Convention Name"
            Height          =   225
            Index           =   7
            Left            =   8355
            TabIndex        =   209
            Top             =   210
            Width           =   1980
         End
         Begin VB.Label L_Bid 
            BackStyle       =   0  'Transparent
            Caption         =   "Use Against:"
            Height          =   225
            Index           =   8
            Left            =   11625
            TabIndex        =   208
            Top             =   210
            Width           =   945
         End
         Begin VB.Label L_Bid 
            BackStyle       =   0  'Transparent
            Caption         =   "Sub"
            Height          =   225
            Index           =   9
            Left            =   390
            TabIndex        =   207
            Top             =   225
            Width           =   300
         End
         Begin VB.Label L_IDnum 
            BackStyle       =   0  'Transparent
            Caption         =   "ID#"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5490
            TabIndex        =   206
            Top             =   810
            Width           =   1620
         End
      End
      Begin VB.TextBox t_HCP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   5570
         MaxLength       =   25
         TabIndex        =   37
         ToolTipText     =   "High Card Points"
         Top             =   3280
         Width           =   2910
      End
      Begin VB.Frame frame0 
         Caption         =   "Other"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2160
         Index           =   4
         Left            =   6660
         TabIndex        =   253
         Top             =   1100
         Width           =   1800
         Begin VB.ComboBox cmb_Position 
            Height          =   315
            ItemData        =   "BidBase Editor.frx":98B3
            Left            =   135
            List            =   "BidBase Editor.frx":98D8
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1095
            Width           =   1575
         End
         Begin VB.ComboBox cmb_Vul 
            Height          =   315
            ItemData        =   "BidBase Editor.frx":9982
            Left            =   135
            List            =   "BidBase Editor.frx":999B
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   465
            Width           =   1575
         End
         Begin VB.TextBox t_BestContract 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1290
            MaxLength       =   5
            TabIndex        =   40
            Top             =   1515
            Width           =   375
         End
         Begin VB.CheckBox ck_Artificial 
            Caption         =   "Artificial"
            Height          =   270
            Left            =   210
            TabIndex        =   41
            Top             =   1850
            Width           =   930
         End
         Begin VB.Label L_Other 
            BackStyle       =   0  'Transparent
            Caption         =   "Position:"
            Height          =   225
            Index           =   1
            Left            =   135
            TabIndex        =   256
            Top             =   900
            Width           =   945
         End
         Begin VB.Label L_Other 
            BackStyle       =   0  'Transparent
            Caption         =   "Vulnerability:"
            Height          =   225
            Index           =   0
            Left            =   135
            TabIndex        =   255
            Top             =   285
            Width           =   1065
         End
         Begin VB.Label L_Other 
            BackStyle       =   0  'Transparent
            Caption         =   "Best contract:"
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   254
            Top             =   1545
            Width           =   1035
         End
      End
      Begin VB.Frame frame0 
         Caption         =   "Disclosure"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2620
         Index           =   5
         Left            =   8500
         TabIndex        =   237
         Top             =   1100
         Width           =   6090
         Begin VB.CommandButton b_CopyFindHand 
            Caption         =   "Copy find hand:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   220
            Index           =   0
            Left            =   3770
            TabIndex        =   280
            Top             =   970
            Width           =   1725
         End
         Begin VB.CommandButton b_EnterHandHelp 
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   1
            Left            =   4430
            TabIndex        =   279
            Top             =   1900
            Width           =   290
         End
         Begin VB.CommandButton b_AddRef 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   4740
            TabIndex        =   267
            Top             =   1900
            Width           =   290
         End
         Begin VB.Frame frame0 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   1305
            Index           =   7
            Left            =   3195
            TabIndex        =   263
            Top             =   135
            Width           =   2775
            Begin VB.CommandButton b_CopyFindHand 
               Caption         =   "Copy to test hand"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   220
               Index           =   1
               Left            =   570
               TabIndex        =   281
               Top             =   1090
               Width           =   1725
            End
         End
         Begin VB.CommandButton b_ClearRefList 
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   5050
            TabIndex        =   49
            Top             =   1900
            Width           =   290
         End
         Begin VB.CommandButton b_SaveRefs 
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   5355
            TabIndex        =   50
            Top             =   1900
            Width           =   290
         End
         Begin VB.CommandButton b_LoadRefs 
            Caption         =   "L"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   5655
            TabIndex        =   51
            Top             =   1900
            Width           =   290
         End
         Begin VB.CommandButton b_DelRef 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   5670
            TabIndex        =   261
            ToolTipText     =   "Delete hand from list."
            Top             =   2220
            Visible         =   0   'False
            Width           =   290
         End
         Begin VB.CommandButton b_ShowRefs 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5670
            Picture         =   "BidBase Editor.frx":9A0F
            Style           =   1  'Graphical
            TabIndex        =   262
            Top             =   2213
            Width           =   290
         End
         Begin VB.TextBox t_TestHand 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            MaxLength       =   16
            TabIndex        =   52
            ToolTipText     =   "A hand matching the specs for this entry for testing the database."
            Top             =   2185
            Width           =   2085
         End
         Begin VB.TextBox t_Disclosure 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   105
            MaxLength       =   50
            TabIndex        =   44
            ToolTipText     =   "Disclosure to be shown to human player."
            Top             =   1475
            Width           =   5565
         End
         Begin VB.ComboBox cmb_Alert 
            Height          =   315
            ItemData        =   "BidBase Editor.frx":C669
            Left            =   1125
            List            =   "BidBase Editor.frx":C67C
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   1005
            Width           =   1410
         End
         Begin VB.TextBox t_AsksFor 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1245
            MaxLength       =   5
            TabIndex        =   42
            Top             =   380
            Width           =   1290
         End
         Begin VB.CheckBox ck_AutoFill 
            Caption         =   "Auto-fill"
            Height          =   195
            Left            =   1365
            TabIndex        =   46
            Top             =   1950
            Width           =   855
         End
         Begin VB.PictureBox b_Discl 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   5730
            Picture         =   "BidBase Editor.frx":C6AE
            ScaleHeight     =   285
            ScaleWidth      =   255
            TabIndex        =   45
            Top             =   1515
            Width           =   255
         End
         Begin VB.TextBox t_Len 
            Alignment       =   2  'Center
            Height          =   240
            Index           =   4
            Left            =   5310
            MaxLength       =   5
            TabIndex        =   242
            Top             =   1125
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox t_Len 
            Alignment       =   2  'Center
            Height          =   240
            Index           =   3
            Left            =   5310
            MaxLength       =   5
            TabIndex        =   241
            Top             =   885
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox t_Len 
            Alignment       =   2  'Center
            Height          =   240
            Index           =   2
            Left            =   5310
            MaxLength       =   5
            TabIndex        =   240
            Top             =   630
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox t_Len 
            Alignment       =   2  'Center
            Height          =   240
            Index           =   1
            Left            =   5310
            MaxLength       =   5
            TabIndex        =   239
            Top             =   360
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.ComboBox cmb_StrDiscl 
            Height          =   315
            ItemData        =   "BidBase Editor.frx":CC6C
            Left            =   3870
            List            =   "BidBase Editor.frx":CC6E
            TabIndex        =   238
            Top             =   825
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.TextBox t_Ref 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2400
            MaxLength       =   25
            TabIndex        =   47
            Top             =   2185
            WhatsThisHelpID =   45
            Width           =   3555
         End
         Begin VB.Label L_Discl 
            BackStyle       =   0  'Transparent
            Caption         =   "Test Hand:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   4
            Left            =   105
            TabIndex        =   252
            Top             =   1950
            Width           =   1080
         End
         Begin VB.Label L_Discl 
            BackStyle       =   0  'Transparent
            Caption         =   "Hand Strength:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   3870
            TabIndex        =   251
            Top             =   390
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label L_Discl 
            BackStyle       =   0  'Transparent
            Caption         =   "Type of Disclosure:"
            Height          =   420
            Index           =   2
            Left            =   180
            TabIndex        =   250
            Top             =   900
            Width           =   855
         End
         Begin VB.Label L_Discl 
            BackStyle       =   0  'Transparent
            Caption         =   "Asks partner for or to bid:"
            Height          =   405
            Index           =   0
            Left            =   210
            TabIndex        =   249
            Top             =   330
            Width           =   975
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000003&
            X1              =   0
            X2              =   5370
            Y1              =   1785
            Y2              =   1785
         End
         Begin VB.Label L_Discl 
            BackStyle       =   0  'Transparent
            Caption         =   "Reference:"
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
            Index           =   5
            Left            =   2400
            TabIndex        =   248
            Top             =   1950
            Width           =   1035
         End
         Begin VB.Label L_Discl 
            BackStyle       =   0  'Transparent
            Caption         =   "Suit Length:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   4815
            TabIndex        =   247
            Top             =   140
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label L_Discl 
            BackStyle       =   0  'Transparent
            Caption         =   "S:"
            Height          =   225
            Index           =   6
            Left            =   5070
            TabIndex        =   246
            Top             =   390
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Label L_Discl 
            BackStyle       =   0  'Transparent
            Caption         =   "H:"
            Height          =   225
            Index           =   7
            Left            =   5070
            TabIndex        =   245
            Top             =   645
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Label L_Discl 
            BackStyle       =   0  'Transparent
            Caption         =   "D:"
            Height          =   225
            Index           =   8
            Left            =   5070
            TabIndex        =   244
            Top             =   885
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Label L_Discl 
            BackStyle       =   0  'Transparent
            Caption         =   "C:"
            Height          =   225
            Index           =   9
            Left            =   5070
            TabIndex        =   243
            Top             =   1125
            Visible         =   0   'False
            Width           =   180
         End
      End
      Begin VB.Frame frame0 
         Caption         =   "Hand Strength"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2620
         Index           =   1
         Left            =   0
         TabIndex        =   228
         Top             =   1100
         Width           =   2130
         Begin VB.TextBox t_BBPoints 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1080
            MaxLength       =   6
            TabIndex        =   300
            ToolTipText     =   "Examples: 6+, 6-, 11-18, >10, <17"
            Top             =   1245
            Width           =   930
         End
         Begin VB.TextBox t_Hand_HCPs 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1080
            MaxLength       =   6
            TabIndex        =   12
            ToolTipText     =   "Examples: 6+, 6-, 11-18, >10, <17"
            Top             =   300
            Width           =   930
         End
         Begin VB.TextBox t_CustPts 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   885
            MaxLength       =   3
            TabIndex        =   17
            ToolTipText     =   "Ace=2, Kings=1"
            Top             =   2270
            Width           =   330
         End
         Begin VB.TextBox t_CustPts 
            Height          =   285
            Index           =   4
            Left            =   1650
            MaxLength       =   3
            TabIndex        =   19
            ToolTipText     =   "Ace=2, Kings=1"
            Top             =   2270
            Width           =   330
         End
         Begin VB.TextBox t_CustPts 
            Height          =   285
            Index           =   3
            Left            =   1260
            MaxLength       =   3
            TabIndex        =   18
            ToolTipText     =   "Ace=2, Kings=1"
            Top             =   2270
            Width           =   330
         End
         Begin VB.TextBox t_CustPts 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   480
            MaxLength       =   3
            TabIndex        =   16
            ToolTipText     =   "Ace=2, Kings=1"
            Top             =   2270
            Width           =   330
         End
         Begin VB.TextBox t_CustPoints 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1080
            MaxLength       =   5
            TabIndex        =   14
            ToolTipText     =   "Ace=2, Kings=1"
            Top             =   1700
            Width           =   930
         End
         Begin VB.TextBox t_CustPts 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   90
            MaxLength       =   3
            TabIndex        =   15
            ToolTipText     =   "Ace=2, Kings=1"
            Top             =   2270
            Width           =   330
         End
         Begin VB.TextBox t_Distrib_Pts 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1080
            MaxLength       =   6
            TabIndex        =   13
            ToolTipText     =   "Distribution points. Examples: 2+, <3, 1-4, >1"
            Top             =   750
            Width           =   930
         End
         Begin VB.Label L_HandStr 
            BackStyle       =   0  'Transparent
            Caption         =   "BB Points:"
            Height          =   255
            Index           =   0
            Left            =   105
            TabIndex        =   299
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label L_HandStr 
            BackStyle       =   0  'Transparent
            Caption         =   "Std. HCPs:"
            Height          =   255
            Index           =   1
            Left            =   105
            TabIndex        =   236
            Top             =   380
            Width           =   975
         End
         Begin VB.Label L_HandStr 
            BackStyle       =   0  'Transparent
            Caption         =   "T"
            Height          =   255
            Index           =   8
            Left            =   1785
            TabIndex        =   235
            Top             =   2080
            Width           =   150
         End
         Begin VB.Label L_HandStr 
            BackStyle       =   0  'Transparent
            Caption         =   "J"
            Height          =   255
            Index           =   7
            Left            =   1380
            TabIndex        =   234
            Top             =   2080
            Width           =   150
         End
         Begin VB.Label L_HandStr 
            BackStyle       =   0  'Transparent
            Caption         =   "Q"
            Height          =   255
            Index           =   6
            Left            =   990
            TabIndex        =   233
            Top             =   2080
            Width           =   150
         End
         Begin VB.Label L_HandStr 
            BackStyle       =   0  'Transparent
            Caption         =   "K"
            Height          =   255
            Index           =   5
            Left            =   585
            TabIndex        =   232
            Top             =   2080
            Width           =   150
         End
         Begin VB.Label L_HandStr 
            BackStyle       =   0  'Transparent
            Caption         =   "Cust.Points:"
            Height          =   255
            Index           =   3
            Left            =   105
            TabIndex        =   231
            Top             =   1750
            Width           =   975
         End
         Begin VB.Label L_HandStr 
            BackStyle       =   0  'Transparent
            Caption         =   "A"
            Height          =   255
            Index           =   4
            Left            =   195
            TabIndex        =   230
            Top             =   2080
            Width           =   150
         End
         Begin VB.Label L_HandStr 
            BackStyle       =   0  'Transparent
            Caption         =   "Distrib.Points:"
            Height          =   255
            Index           =   9
            Left            =   105
            TabIndex        =   229
            Top             =   820
            Width           =   975
         End
      End
      Begin VB.Frame frame0 
         Caption         =   "Suits"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2620
         Index           =   3
         Left            =   4010
         TabIndex        =   219
         Top             =   1100
         Width           =   2600
         Begin VB.TextBox t_Qty 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   615
            MaxLength       =   10
            TabIndex        =   26
            ToolTipText     =   "# in Suit"
            Top             =   350
            Width           =   950
         End
         Begin VB.TextBox t_Qty 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   615
            MaxLength       =   5
            TabIndex        =   28
            ToolTipText     =   "# in Suit"
            Top             =   718
            Width           =   950
         End
         Begin VB.TextBox t_Qty 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   615
            MaxLength       =   5
            TabIndex        =   30
            ToolTipText     =   "# in Suit"
            Top             =   1086
            Width           =   950
         End
         Begin VB.TextBox t_Qty 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   615
            MaxLength       =   5
            TabIndex        =   32
            ToolTipText     =   "# in Suit"
            Top             =   1454
            Width           =   950
         End
         Begin VB.TextBox t_Qty 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   615
            MaxLength       =   5
            TabIndex        =   34
            ToolTipText     =   "# in Suit"
            Top             =   1822
            Width           =   950
         End
         Begin VB.TextBox t_HCP 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   1570
            MaxLength       =   10
            TabIndex        =   27
            ToolTipText     =   "High Card Points"
            Top             =   350
            Width           =   950
         End
         Begin VB.TextBox t_HCP 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   1570
            MaxLength       =   10
            TabIndex        =   29
            ToolTipText     =   "High Card Points"
            Top             =   717
            Width           =   950
         End
         Begin VB.TextBox t_HCP 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   1570
            MaxLength       =   10
            TabIndex        =   31
            ToolTipText     =   "High Card Points"
            Top             =   1084
            Width           =   950
         End
         Begin VB.TextBox t_HCP 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   1570
            MaxLength       =   10
            TabIndex        =   33
            ToolTipText     =   "High Card Points"
            Top             =   1451
            Width           =   950
         End
         Begin VB.TextBox t_HCP 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   1570
            MaxLength       =   10
            TabIndex        =   35
            ToolTipText     =   "High Card Points"
            Top             =   1820
            Width           =   950
         End
         Begin VB.TextBox t_Qty 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   5
            Left            =   615
            MaxLength       =   9
            TabIndex        =   36
            ToolTipText     =   "# in Suit"
            Top             =   2190
            Width           =   950
         End
         Begin VB.Label L_Suits 
            BackStyle       =   0  'Transparent
            Caption         =   """X"""
            Height          =   225
            Index           =   3
            Left            =   135
            TabIndex        =   227
            Top             =   430
            Width           =   400
         End
         Begin VB.Label L_Suits 
            BackStyle       =   0  'Transparent
            Caption         =   "Spds:"
            Height          =   225
            Index           =   4
            Left            =   135
            TabIndex        =   226
            Top             =   804
            Width           =   400
         End
         Begin VB.Label L_Suits 
            BackStyle       =   0  'Transparent
            Caption         =   "Hrts:"
            Height          =   225
            Index           =   5
            Left            =   135
            TabIndex        =   225
            Top             =   1178
            Width           =   450
         End
         Begin VB.Label L_Suits 
            BackStyle       =   0  'Transparent
            Caption         =   "Dia:"
            Height          =   225
            Index           =   6
            Left            =   135
            TabIndex        =   224
            Top             =   1552
            Width           =   390
         End
         Begin VB.Label L_Suits 
            BackStyle       =   0  'Transparent
            Caption         =   "Clbs:"
            Height          =   225
            Index           =   7
            Left            =   135
            TabIndex        =   223
            Top             =   1926
            Width           =   405
         End
         Begin VB.Label L_Suits 
            BackStyle       =   0  'Transparent
            Caption         =   "Num."
            Height          =   225
            Index           =   1
            Left            =   780
            TabIndex        =   222
            Top             =   150
            Width           =   480
         End
         Begin VB.Label L_Suits 
            BackStyle       =   0  'Transparent
            Caption         =   "Points"
            Height          =   225
            Index           =   2
            Left            =   1680
            TabIndex        =   221
            Top             =   150
            Width           =   465
         End
         Begin VB.Label L_Suits 
            BackStyle       =   0  'Transparent
            Caption         =   "Comb:"
            Height          =   225
            Index           =   8
            Left            =   120
            TabIndex        =   220
            Top             =   2300
            Width           =   510
         End
      End
      Begin VB.Frame frame0 
         Height          =   2620
         Index           =   2
         Left            =   2175
         TabIndex        =   191
         Top             =   1100
         Width           =   1725
         Begin VB.TextBox t_Winners 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1185
            MaxLength       =   5
            TabIndex        =   20
            ToolTipText     =   "Number of winners."
            Top             =   255
            Width           =   375
         End
         Begin VB.TextBox t_Losers 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1185
            MaxLength       =   5
            TabIndex        =   21
            ToolTipText     =   "Number of Losers."
            Top             =   648
            Width           =   375
         End
         Begin VB.TextBox t_QTs 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1170
            MaxLength       =   5
            TabIndex        =   22
            ToolTipText     =   "Number of Quick Tricks"
            Top             =   1041
            Width           =   375
         End
         Begin VB.TextBox t_Stoppers 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1185
            MaxLength       =   1
            TabIndex        =   23
            ToolTipText     =   "Number of suits stopped other than the suit bid."
            Top             =   1434
            Width           =   375
         End
         Begin VB.TextBox t_Intermediates 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1185
            MaxLength       =   1
            TabIndex        =   24
            Top             =   1830
            Width           =   375
         End
         Begin VB.ComboBox cmb_Shape 
            Height          =   315
            ItemData        =   "BidBase Editor.frx":CC70
            Left            =   660
            List            =   "BidBase Editor.frx":CC8F
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   2235
            Width           =   990
         End
         Begin VB.Label L_HandStr 
            BackStyle       =   0  'Transparent
            Caption         =   "# Winners:"
            Height          =   255
            Index           =   10
            Left            =   150
            TabIndex        =   197
            Top             =   300
            Width           =   975
         End
         Begin VB.Label L_HandStr 
            BackStyle       =   0  'Transparent
            Caption         =   "Losing Tricks:"
            Height          =   255
            Index           =   11
            Left            =   105
            TabIndex        =   196
            Top             =   723
            Width           =   1095
         End
         Begin VB.Label L_HandStr 
            BackStyle       =   0  'Transparent
            Caption         =   "Defen. Tricks:"
            Height          =   255
            Index           =   12
            Left            =   105
            TabIndex        =   195
            Top             =   1146
            Width           =   1095
         End
         Begin VB.Label L_HandStr 
            BackStyle       =   0  'Transparent
            Caption         =   "Suits Stopped:"
            Height          =   255
            Index           =   13
            Left            =   105
            TabIndex        =   194
            Top             =   1530
            Width           =   1095
         End
         Begin VB.Label L_HandStr 
            BackStyle       =   0  'Transparent
            Caption         =   "Intermediates:"
            Height          =   225
            Index           =   15
            Left            =   120
            TabIndex        =   193
            Top             =   1910
            Width           =   1035
         End
         Begin VB.Label L_Suits 
            BackStyle       =   0  'Transparent
            Caption         =   "Shape:"
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   192
            Top             =   2310
            Width           =   555
         End
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   240
         LargeChange     =   1000
         Left            =   -15
         Max             =   18500
         SmallChange     =   100
         TabIndex        =   218
         Top             =   4410
         Width           =   20400
      End
      Begin VB.Label lblBattTime 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "Remaining Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   225
         Left            =   16530
         TabIndex        =   260
         Top             =   2520
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label lblBattStat 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "Battery Charge"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   225
         Left            =   16560
         TabIndex        =   259
         Top             =   1680
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label lblBatt 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         Height          =   225
         Left            =   16920
         TabIndex        =   258
         Top             =   1920
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         ForeColor       =   &H00FFFF80&
         Height          =   225
         Left            =   16920
         TabIndex        =   257
         Top             =   2880
         Visible         =   0   'False
         Width           =   840
      End
   End
   Begin VB.Frame frm_Tutorial 
      BackColor       =   &H00808080&
      Caption         =   "BidBase Editor Tutorial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   3675
      Left            =   2925
      TabIndex        =   118
      Top             =   1545
      Visible         =   0   'False
      Width           =   6300
      Begin VB.PictureBox pTutorDragBar 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   100
         Left            =   100
         MousePointer    =   7  'Size N S
         ScaleHeight     =   105
         ScaleWidth      =   6045
         TabIndex        =   173
         Top             =   3540
         Width           =   6050
      End
      Begin VB.CommandButton b_TutorDetails 
         Caption         =   "&Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3060
         TabIndex        =   122
         Top             =   3150
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton b_Go_Back 
         Caption         =   "&Go Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1380
         TabIndex        =   121
         Top             =   3150
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton b_Continue 
         Caption         =   "&Continue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   210
         TabIndex        =   120
         Top             =   3150
         Width           =   1035
      End
      Begin VB.CommandButton b_QuitTutorial 
         Caption         =   "&Quit Tutorial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4700
         TabIndex        =   124
         Top             =   3150
         Width           =   1380
      End
      Begin VB.Label L_Tutorial 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2820
         Left            =   165
         TabIndex        =   119
         Top             =   195
         Width           =   5895
      End
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   22695
      Top             =   5100
   End
   Begin VB.TextBox t_InputFromPractice 
      Height          =   1095
      Left            =   13560
      LinkTopic       =   "BidBase Practice Program|t_EditData"
      TabIndex        =   172
      Text            =   "Text1"
      Top             =   6720
      Visible         =   0   'False
      Width           =   2730
   End
   Begin VB.PictureBox pSplitterBar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   250
      Left            =   30
      MousePointer    =   7  'Size N S
      ScaleHeight     =   256.977
      ScaleMode       =   0  'User
      ScaleWidth      =   13350
      TabIndex        =   169
      Top             =   0
      Width           =   13350
   End
   Begin VB.Frame f_GridColors 
      Caption         =   "Set Colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3285
      Left            =   1725
      TabIndex        =   53
      Top             =   1485
      Visible         =   0   'False
      Width           =   5325
      Begin VB.CommandButton b_CloseColors 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5040
         TabIndex        =   68
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton b_DupeColors 
         Caption         =   "Copy Top Grid Colors To All"
         Height          =   315
         Left            =   1785
         TabIndex        =   64
         Top             =   2685
         Width           =   2925
      End
      Begin VB.ComboBox cmb_ForeColors 
         Height          =   315
         Index           =   3
         ItemData        =   "BidBase Editor.frx":CCE1
         Left            =   1230
         List            =   "BidBase Editor.frx":CD18
         TabIndex        =   62
         Top             =   2145
         Width           =   1890
      End
      Begin VB.ComboBox cmb_BackColors 
         Height          =   315
         Index           =   3
         ItemData        =   "BidBase Editor.frx":CDBE
         Left            =   3240
         List            =   "BidBase Editor.frx":CDF5
         TabIndex        =   63
         Top             =   2145
         Width           =   1890
      End
      Begin VB.ComboBox cmb_ForeColors 
         Height          =   315
         Index           =   2
         ItemData        =   "BidBase Editor.frx":CE9B
         Left            =   1230
         List            =   "BidBase Editor.frx":CED2
         TabIndex        =   60
         Top             =   1665
         Width           =   1890
      End
      Begin VB.ComboBox cmb_BackColors 
         Height          =   315
         Index           =   2
         ItemData        =   "BidBase Editor.frx":CF78
         Left            =   3240
         List            =   "BidBase Editor.frx":CFAF
         TabIndex        =   61
         Top             =   1650
         Width           =   1890
      End
      Begin VB.ComboBox cmb_ForeColors 
         Height          =   315
         Index           =   1
         ItemData        =   "BidBase Editor.frx":D055
         Left            =   1230
         List            =   "BidBase Editor.frx":D08C
         TabIndex        =   266
         Top             =   1185
         Width           =   1890
      End
      Begin VB.ComboBox cmb_BackColors 
         Height          =   315
         Index           =   1
         ItemData        =   "BidBase Editor.frx":D132
         Left            =   3240
         List            =   "BidBase Editor.frx":D169
         TabIndex        =   59
         Top             =   1185
         Width           =   1890
      End
      Begin VB.ComboBox cmb_BackColors 
         Height          =   315
         Index           =   0
         ItemData        =   "BidBase Editor.frx":D20F
         Left            =   3240
         List            =   "BidBase Editor.frx":D246
         TabIndex        =   57
         Top             =   705
         Width           =   1890
      End
      Begin VB.ComboBox cmb_ForeColors 
         Height          =   315
         Index           =   0
         ItemData        =   "BidBase Editor.frx":D2EC
         Left            =   1230
         List            =   "BidBase Editor.frx":D323
         TabIndex        =   56
         Top             =   705
         Width           =   1890
      End
      Begin VB.Label L_Colors 
         BackStyle       =   0  'Transparent
         Caption         =   "Window:"
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   67
         Top             =   2205
         Width           =   1065
      End
      Begin VB.Label L_Colors 
         BackStyle       =   0  'Transparent
         Caption         =   "Input Boxes:"
         Height          =   255
         Index           =   4
         Left            =   180
         TabIndex        =   66
         Top             =   1725
         Width           =   1065
      End
      Begin VB.Label L_Colors 
         BackStyle       =   0  'Transparent
         Caption         =   "Other Grids:"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   65
         Top             =   1245
         Width           =   1065
      End
      Begin VB.Label L_Colors 
         BackStyle       =   0  'Transparent
         Caption         =   "Active Grid:"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   58
         Top             =   735
         Width           =   990
      End
      Begin VB.Label L_Colors 
         BackStyle       =   0  'Transparent
         Caption         =   "Background:"
         Height          =   225
         Index           =   1
         Left            =   3660
         TabIndex        =   55
         Top             =   450
         Width           =   1035
      End
      Begin VB.Label L_Colors 
         BackStyle       =   0  'Transparent
         Caption         =   "Text:"
         Height          =   225
         Index           =   0
         Left            =   1830
         TabIndex        =   54
         Top             =   450
         Width           =   525
      End
   End
   Begin VB.Frame f_ConvNameChange 
      BackColor       =   &H008080FF&
      Caption         =   "Convention Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   9045
      TabIndex        =   128
      Top             =   1020
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton b_ConvNameChange 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1695
         TabIndex        =   141
         Top             =   4020
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         Height          =   2745
         Left            =   60
         TabIndex        =   129
         Top             =   1005
         Width           =   4230
         Begin VB.OptionButton opt_ChConvName 
            BackColor       =   &H00FFFF80&
            Caption         =   "Change"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   105
            TabIndex        =   132
            Top             =   930
            Width           =   3105
         End
         Begin VB.OptionButton opt_ChConvName 
            BackColor       =   &H0080FFFF&
            Caption         =   "Add"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   105
            TabIndex        =   131
            Top             =   15
            Value           =   -1  'True
            Width           =   990
         End
         Begin VB.OptionButton opt_ChConvName 
            BackColor       =   &H0080FF80&
            Caption         =   "Go back to original name of:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   105
            TabIndex        =   130
            Top             =   1965
            Width           =   4305
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFF80&
            Caption         =   "in the list to:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   390
            TabIndex        =   134
            Top             =   1380
            Width           =   1905
         End
         Begin VB.Label Label1 
            BackColor       =   &H0080FFFF&
            Caption         =   "to the list."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   390
            TabIndex        =   136
            Top             =   525
            Width           =   2790
         End
         Begin VB.Label L_OldName 
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   540
            TabIndex        =   133
            Top             =   2190
            Width           =   3645
         End
         Begin VB.Label L_OldName 
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   540
            TabIndex        =   137
            Top             =   1155
            Width           =   3645
         End
         Begin VB.Label L_NewName 
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   540
            TabIndex        =   135
            Top             =   1590
            Width           =   3645
         End
         Begin VB.Label L_NewName 
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   540
            TabIndex        =   138
            Top             =   300
            Width           =   3645
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF80&
            Height          =   1005
            Index           =   1
            Left            =   30
            TabIndex        =   140
            Top             =   855
            Width           =   4320
         End
         Begin VB.Label Label2 
            BackColor       =   &H0080FFFF&
            Height          =   855
            Index           =   0
            Left            =   15
            TabIndex        =   139
            Top             =   -15
            Width           =   4320
         End
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   0
         X1              =   4300
         X2              =   4300
         Y1              =   4710
         Y2              =   120
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   30
         X2              =   4320
         Y1              =   4650
         Y2              =   4650
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   1
         X1              =   45
         X2              =   45
         Y1              =   4740
         Y2              =   135
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   2070
         X2              =   4300
         Y1              =   120
         Y2              =   135
      End
      Begin VB.Label Label3 
         BackColor       =   &H008080FF&
         Caption         =   "Choose:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1500
         TabIndex        =   142
         Top             =   375
         Width           =   1620
      End
   End
   Begin VB.PictureBox p_Testing 
      BackColor       =   &H00000000&
      FillColor       =   &H000000FF&
      Height          =   900
      Left            =   2495
      ScaleHeight     =   840
      ScaleWidth      =   1485
      TabIndex        =   123
      Top             =   2355
      Visible         =   0   'False
      Width           =   1545
      Begin VB.Label L_TestingMsg 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "      Testing Entries"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1380
         Left            =   -30
         TabIndex        =   125
         Top             =   -300
         Width           =   1530
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   22680
      Top             =   2850
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      LargeChange     =   10
      Left            =   22710
      TabIndex        =   0
      Top             =   1545
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Left            =   22710
      Top             =   3405
   End
   Begin MSComDlg.CommonDialog CmDia 
      Left            =   22680
      Top             =   3945
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   345
      Index           =   0
      Left            =   17595
      Top             =   1380
      Visible         =   0   'False
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   609
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Desktop\VB98\Bridge\BidBase.accdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Desktop\VB98\Bridge\BidBase.accdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"BidBase Editor.frx":D3C9
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   345
      Index           =   1
      Left            =   17610
      Top             =   2175
      Visible         =   0   'False
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   609
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Desktop\VB98\Bridge\BidBase.accdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Desktop\VB98\Bridge\BidBase.accdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"BidBase Editor.frx":D6DE
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   345
      Index           =   2
      Left            =   17640
      Top             =   3240
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   609
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Desktop\VB98\Bridge\BidBase.accdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Desktop\VB98\Bridge\BidBase.accdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"BidBase Editor.frx":D9F0
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   345
      Index           =   3
      Left            =   17670
      Top             =   4200
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   609
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Desktop\VB98\Bridge\BidBase.accdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Desktop\VB98\Bridge\BidBase.accdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"BidBase Editor.frx":DD05
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   345
      Index           =   4
      Left            =   17730
      Top             =   5415
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   609
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Desktop\VB98\Bridge\BidBase.accdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\Desktop\VB98\Bridge\BidBase.accdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"BidBase Editor.frx":E01A
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid Grid1 
      Bindings        =   "BidBase Editor.frx":E32F
      Height          =   915
      Index           =   1
      Left            =   13770
      TabIndex        =   269
      Top             =   1920
      Visible         =   0   'False
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   1614
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   51
      BeginProperty Column00 
         DataField       =   "Selected"
         Caption         =   "Selected"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "SubCat_Selected"
         Caption         =   "SubCat_Selected"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Pct_Used"
         Caption         =   "Pct_Used"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "RoundOfBidding"
         Caption         =   "RoundOfBidding"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Player"
         Caption         =   "Player"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "PB_Sort_Order"
         Caption         =   "PB_Sort_Order"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Prior_Bids"
         Caption         =   "Prior_Bids"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Bid_Order_Number"
         Caption         =   "Bid_Order_Number"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Bid"
         Caption         =   "Bid"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Bid_Name"
         Caption         =   "Bid_Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "HCP"
         Caption         =   "HCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Dist_Pts"
         Caption         =   "Dist_Pts"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "Cust_Pts"
         Caption         =   "Cust_Pts"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "A_Cust"
         Caption         =   "A_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "K_Cust"
         Caption         =   "K_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "Q_Cust"
         Caption         =   "Q_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "J_Cust"
         Caption         =   "J_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "T_Cust"
         Caption         =   "T_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "Winrs"
         Caption         =   "Winrs"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column19 
         DataField       =   "Losers"
         Caption         =   "Losers"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column20 
         DataField       =   "QTs"
         Caption         =   "QTs"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column21 
         DataField       =   "Stopped_Suits"
         Caption         =   "Stopped_Suits"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column22 
         DataField       =   "Shape"
         Caption         =   "Shape"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column23 
         DataField       =   "Intermed"
         Caption         =   "Intermed"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column24 
         DataField       =   "xNum"
         Caption         =   "xNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column25 
         DataField       =   "xHCP"
         Caption         =   "xHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column26 
         DataField       =   "sNum"
         Caption         =   "sNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column27 
         DataField       =   "sHCP"
         Caption         =   "sHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column28 
         DataField       =   "hNum"
         Caption         =   "hNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column29 
         DataField       =   "hHCP"
         Caption         =   "hHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column30 
         DataField       =   "dNum"
         Caption         =   "dNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column31 
         DataField       =   "dHCP"
         Caption         =   "dHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column32 
         DataField       =   "cNum"
         Caption         =   "cNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column33 
         DataField       =   "cHCP"
         Caption         =   "cHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column34 
         DataField       =   "cmNum"
         Caption         =   "cmNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column35 
         DataField       =   "cmHCP"
         Caption         =   "cmHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column36 
         DataField       =   "Vul"
         Caption         =   "Vul"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column37 
         DataField       =   "Pos"
         Caption         =   "Pos"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column38 
         DataField       =   "Artificial"
         Caption         =   "Artificial"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column39 
         DataField       =   "Alert"
         Caption         =   "Alert"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column40 
         DataField       =   "Explanation"
         Caption         =   "Explanation"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column41 
         DataField       =   "Test_Hand"
         Caption         =   "Test_Hand"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column42 
         DataField       =   "Asks"
         Caption         =   "Asks"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column43 
         DataField       =   "DateOfEntry"
         Caption         =   "DateOfEntry"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column44 
         DataField       =   "Best_Contract"
         Caption         =   "Best_Contract"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column45 
         DataField       =   "Num_Used"
         Caption         =   "Num_Used"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column46 
         DataField       =   "Source"
         Caption         =   "Source"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column47 
         DataField       =   "LHO"
         Caption         =   "LHO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column48 
         DataField       =   "Pard"
         Caption         =   "Pard"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column49 
         DataField       =   "RHO"
         Caption         =   "RHO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column50 
         DataField       =   "ID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   285.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   329.953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   315.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   225.071
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   209.764
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   104.882
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   434.835
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   345.26
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   209.764
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   209.764
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   239.811
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   150.236
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   195.024
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column26 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column27 
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column28 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column29 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column30 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column31 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column32 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column33 
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column34 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column35 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column36 
            ColumnWidth     =   315.213
         EndProperty
         BeginProperty Column37 
            ColumnWidth     =   374.74
         EndProperty
         BeginProperty Column38 
            ColumnWidth     =   255.118
         EndProperty
         BeginProperty Column39 
            ColumnWidth     =   374.74
         EndProperty
         BeginProperty Column40 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column41 
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column42 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column43 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column44 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column45 
            ColumnWidth     =   60.094
         EndProperty
         BeginProperty Column46 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column47 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column48 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column49 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column50 
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid Grid1 
      Bindings        =   "BidBase Editor.frx":E346
      Height          =   915
      Index           =   2
      Left            =   13800
      TabIndex        =   270
      Top             =   3000
      Visible         =   0   'False
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   1614
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   51
      BeginProperty Column00 
         DataField       =   "Selected"
         Caption         =   "Selected"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "SubCat_Selected"
         Caption         =   "SubCat_Selected"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Pct_Used"
         Caption         =   "Pct_Used"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "RoundOfBidding"
         Caption         =   "RoundOfBidding"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Player"
         Caption         =   "Player"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "PB_Sort_Order"
         Caption         =   "PB_Sort_Order"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Prior_Bids"
         Caption         =   "Prior_Bids"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Bid_Order_Number"
         Caption         =   "Bid_Order_Number"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Bid"
         Caption         =   "Bid"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Bid_Name"
         Caption         =   "Bid_Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "HCP"
         Caption         =   "HCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Dist_Pts"
         Caption         =   "Dist_Pts"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "Cust_Pts"
         Caption         =   "Cust_Pts"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "A_Cust"
         Caption         =   "A_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "K_Cust"
         Caption         =   "K_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "Q_Cust"
         Caption         =   "Q_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "J_Cust"
         Caption         =   "J_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "T_Cust"
         Caption         =   "T_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "Winrs"
         Caption         =   "Winrs"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column19 
         DataField       =   "Losers"
         Caption         =   "Losers"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column20 
         DataField       =   "QTs"
         Caption         =   "QTs"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column21 
         DataField       =   "Stopped_Suits"
         Caption         =   "Stopped_Suits"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column22 
         DataField       =   "Shape"
         Caption         =   "Shape"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column23 
         DataField       =   "Intermed"
         Caption         =   "Intermed"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column24 
         DataField       =   "xNum"
         Caption         =   "xNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column25 
         DataField       =   "xHCP"
         Caption         =   "xHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column26 
         DataField       =   "sNum"
         Caption         =   "sNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column27 
         DataField       =   "sHCP"
         Caption         =   "sHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column28 
         DataField       =   "hNum"
         Caption         =   "hNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column29 
         DataField       =   "hHCP"
         Caption         =   "hHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column30 
         DataField       =   "dNum"
         Caption         =   "dNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column31 
         DataField       =   "dHCP"
         Caption         =   "dHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column32 
         DataField       =   "cNum"
         Caption         =   "cNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column33 
         DataField       =   "cHCP"
         Caption         =   "cHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column34 
         DataField       =   "cmNum"
         Caption         =   "cmNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column35 
         DataField       =   "cmHCP"
         Caption         =   "cmHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column36 
         DataField       =   "Vul"
         Caption         =   "Vul"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column37 
         DataField       =   "Pos"
         Caption         =   "Pos"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column38 
         DataField       =   "Artificial"
         Caption         =   "Artificial"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column39 
         DataField       =   "Alert"
         Caption         =   "Alert"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column40 
         DataField       =   "Explanation"
         Caption         =   "Explanation"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column41 
         DataField       =   "Test_Hand"
         Caption         =   "Test_Hand"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column42 
         DataField       =   "Asks"
         Caption         =   "Asks"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column43 
         DataField       =   "DateOfEntry"
         Caption         =   "DateOfEntry"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column44 
         DataField       =   "Best_Contract"
         Caption         =   "Best_Contract"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column45 
         DataField       =   "Num_Used"
         Caption         =   "Num_Used"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column46 
         DataField       =   "Source"
         Caption         =   "Source"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column47 
         DataField       =   "LHO"
         Caption         =   "LHO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column48 
         DataField       =   "Pard"
         Caption         =   "Pard"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column49 
         DataField       =   "RHO"
         Caption         =   "RHO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column50 
         DataField       =   "ID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   285.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   329.953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   315.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   225.071
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   209.764
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   104.882
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   434.835
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   345.26
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   209.764
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   209.764
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   239.811
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   150.236
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   195.024
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column26 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column27 
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column28 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column29 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column30 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column31 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column32 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column33 
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column34 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column35 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column36 
            ColumnWidth     =   315.213
         EndProperty
         BeginProperty Column37 
            ColumnWidth     =   374.74
         EndProperty
         BeginProperty Column38 
            ColumnWidth     =   255.118
         EndProperty
         BeginProperty Column39 
            ColumnWidth     =   374.74
         EndProperty
         BeginProperty Column40 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column41 
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column42 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column43 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column44 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column45 
            ColumnWidth     =   60.094
         EndProperty
         BeginProperty Column46 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column47 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column48 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column49 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column50 
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid Grid1 
      Bindings        =   "BidBase Editor.frx":E35D
      Height          =   915
      Index           =   3
      Left            =   13800
      TabIndex        =   271
      Top             =   4080
      Visible         =   0   'False
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   1614
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   51
      BeginProperty Column00 
         DataField       =   "Selected"
         Caption         =   "Selected"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "SubCat_Selected"
         Caption         =   "SubCat_Selected"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Pct_Used"
         Caption         =   "Pct_Used"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "RoundOfBidding"
         Caption         =   "RoundOfBidding"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Player"
         Caption         =   "Player"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "PB_Sort_Order"
         Caption         =   "PB_Sort_Order"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Prior_Bids"
         Caption         =   "Prior_Bids"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Bid_Order_Number"
         Caption         =   "Bid_Order_Number"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Bid"
         Caption         =   "Bid"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Bid_Name"
         Caption         =   "Bid_Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "HCP"
         Caption         =   "HCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Dist_Pts"
         Caption         =   "Dist_Pts"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "Cust_Pts"
         Caption         =   "Cust_Pts"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "A_Cust"
         Caption         =   "A_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "K_Cust"
         Caption         =   "K_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "Q_Cust"
         Caption         =   "Q_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "J_Cust"
         Caption         =   "J_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "T_Cust"
         Caption         =   "T_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "Winrs"
         Caption         =   "Winrs"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column19 
         DataField       =   "Losers"
         Caption         =   "Losers"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column20 
         DataField       =   "QTs"
         Caption         =   "QTs"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column21 
         DataField       =   "Stopped_Suits"
         Caption         =   "Stopped_Suits"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column22 
         DataField       =   "Shape"
         Caption         =   "Shape"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column23 
         DataField       =   "Intermed"
         Caption         =   "Intermed"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column24 
         DataField       =   "xNum"
         Caption         =   "xNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column25 
         DataField       =   "xHCP"
         Caption         =   "xHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column26 
         DataField       =   "sNum"
         Caption         =   "sNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column27 
         DataField       =   "sHCP"
         Caption         =   "sHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column28 
         DataField       =   "hNum"
         Caption         =   "hNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column29 
         DataField       =   "hHCP"
         Caption         =   "hHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column30 
         DataField       =   "dNum"
         Caption         =   "dNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column31 
         DataField       =   "dHCP"
         Caption         =   "dHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column32 
         DataField       =   "cNum"
         Caption         =   "cNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column33 
         DataField       =   "cHCP"
         Caption         =   "cHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column34 
         DataField       =   "cmNum"
         Caption         =   "cmNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column35 
         DataField       =   "cmHCP"
         Caption         =   "cmHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column36 
         DataField       =   "Vul"
         Caption         =   "Vul"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column37 
         DataField       =   "Pos"
         Caption         =   "Pos"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column38 
         DataField       =   "Artificial"
         Caption         =   "Artificial"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column39 
         DataField       =   "Alert"
         Caption         =   "Alert"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column40 
         DataField       =   "Explanation"
         Caption         =   "Explanation"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column41 
         DataField       =   "Test_Hand"
         Caption         =   "Test_Hand"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column42 
         DataField       =   "Asks"
         Caption         =   "Asks"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column43 
         DataField       =   "DateOfEntry"
         Caption         =   "DateOfEntry"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column44 
         DataField       =   "Best_Contract"
         Caption         =   "Best_Contract"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column45 
         DataField       =   "Num_Used"
         Caption         =   "Num_Used"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column46 
         DataField       =   "Source"
         Caption         =   "Source"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column47 
         DataField       =   "LHO"
         Caption         =   "LHO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column48 
         DataField       =   "Pard"
         Caption         =   "Pard"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column49 
         DataField       =   "RHO"
         Caption         =   "RHO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column50 
         DataField       =   "ID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   285.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   329.953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   315.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   225.071
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   209.764
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   104.882
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   434.835
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   345.26
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   209.764
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   209.764
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   239.811
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   150.236
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   195.024
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column26 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column27 
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column28 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column29 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column30 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column31 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column32 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column33 
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column34 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column35 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column36 
            ColumnWidth     =   315.213
         EndProperty
         BeginProperty Column37 
            ColumnWidth     =   374.74
         EndProperty
         BeginProperty Column38 
            ColumnWidth     =   255.118
         EndProperty
         BeginProperty Column39 
            ColumnWidth     =   374.74
         EndProperty
         BeginProperty Column40 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column41 
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column42 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column43 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column44 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column45 
            ColumnWidth     =   60.094
         EndProperty
         BeginProperty Column46 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column47 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column48 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column49 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column50 
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid Grid1 
      Bindings        =   "BidBase Editor.frx":E374
      Height          =   915
      Index           =   4
      Left            =   13800
      TabIndex        =   272
      Top             =   5160
      Visible         =   0   'False
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   1614
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   51
      BeginProperty Column00 
         DataField       =   "Selected"
         Caption         =   "Selected"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "SubCat_Selected"
         Caption         =   "SubCat_Selected"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Pct_Used"
         Caption         =   "Pct_Used"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "RoundOfBidding"
         Caption         =   "RoundOfBidding"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Player"
         Caption         =   "Player"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "PB_Sort_Order"
         Caption         =   "PB_Sort_Order"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Prior_Bids"
         Caption         =   "Prior_Bids"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Bid_Order_Number"
         Caption         =   "Bid_Order_Number"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Bid"
         Caption         =   "Bid"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Bid_Name"
         Caption         =   "Bid_Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "HCP"
         Caption         =   "HCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Dist_Pts"
         Caption         =   "Dist_Pts"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "Cust_Pts"
         Caption         =   "Cust_Pts"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "A_Cust"
         Caption         =   "A_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "K_Cust"
         Caption         =   "K_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "Q_Cust"
         Caption         =   "Q_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "J_Cust"
         Caption         =   "J_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "T_Cust"
         Caption         =   "T_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "Winrs"
         Caption         =   "Winrs"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column19 
         DataField       =   "Losers"
         Caption         =   "Losers"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column20 
         DataField       =   "QTs"
         Caption         =   "QTs"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column21 
         DataField       =   "Stopped_Suits"
         Caption         =   "Stopped_Suits"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column22 
         DataField       =   "Shape"
         Caption         =   "Shape"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column23 
         DataField       =   "Intermed"
         Caption         =   "Intermed"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column24 
         DataField       =   "xNum"
         Caption         =   "xNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column25 
         DataField       =   "xHCP"
         Caption         =   "xHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column26 
         DataField       =   "sNum"
         Caption         =   "sNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column27 
         DataField       =   "sHCP"
         Caption         =   "sHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column28 
         DataField       =   "hNum"
         Caption         =   "hNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column29 
         DataField       =   "hHCP"
         Caption         =   "hHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column30 
         DataField       =   "dNum"
         Caption         =   "dNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column31 
         DataField       =   "dHCP"
         Caption         =   "dHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column32 
         DataField       =   "cNum"
         Caption         =   "cNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column33 
         DataField       =   "cHCP"
         Caption         =   "cHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column34 
         DataField       =   "cmNum"
         Caption         =   "cmNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column35 
         DataField       =   "cmHCP"
         Caption         =   "cmHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column36 
         DataField       =   "Vul"
         Caption         =   "Vul"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column37 
         DataField       =   "Pos"
         Caption         =   "Pos"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column38 
         DataField       =   "Artificial"
         Caption         =   "Artificial"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column39 
         DataField       =   "Alert"
         Caption         =   "Alert"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column40 
         DataField       =   "Explanation"
         Caption         =   "Explanation"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column41 
         DataField       =   "Test_Hand"
         Caption         =   "Test_Hand"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column42 
         DataField       =   "Asks"
         Caption         =   "Asks"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column43 
         DataField       =   "DateOfEntry"
         Caption         =   "DateOfEntry"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column44 
         DataField       =   "Best_Contract"
         Caption         =   "Best_Contract"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column45 
         DataField       =   "Num_Used"
         Caption         =   "Num_Used"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column46 
         DataField       =   "Source"
         Caption         =   "Source"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column47 
         DataField       =   "LHO"
         Caption         =   "LHO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column48 
         DataField       =   "Pard"
         Caption         =   "Pard"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column49 
         DataField       =   "RHO"
         Caption         =   "RHO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column50 
         DataField       =   "ID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   285.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   329.953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   315.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   225.071
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   209.764
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   104.882
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   434.835
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   345.26
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   209.764
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   209.764
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   239.811
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   150.236
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   195.024
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column26 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column27 
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column28 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column29 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column30 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column31 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column32 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column33 
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column34 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column35 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column36 
            ColumnWidth     =   315.213
         EndProperty
         BeginProperty Column37 
            ColumnWidth     =   374.74
         EndProperty
         BeginProperty Column38 
            ColumnWidth     =   255.118
         EndProperty
         BeginProperty Column39 
            ColumnWidth     =   374.74
         EndProperty
         BeginProperty Column40 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column41 
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column42 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column43 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column44 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column45 
            ColumnWidth     =   60.094
         EndProperty
         BeginProperty Column46 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column47 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column48 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column49 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column50 
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid Grid1 
      Bindings        =   "BidBase Editor.frx":E38B
      Height          =   1155
      Index           =   0
      Left            =   -150
      TabIndex        =   171
      Top             =   345
      Width           =   28875
      _ExtentX        =   50932
      _ExtentY        =   2037
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   51
      BeginProperty Column00 
         DataField       =   "Selected"
         Caption         =   "Selected"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "SubCat_Selected"
         Caption         =   "SubCat_Selected"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Pct_Used"
         Caption         =   "Pct_Used"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "RoundOfBidding"
         Caption         =   "RoundOfBidding"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Player"
         Caption         =   "Player"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "PB_Sort_Order"
         Caption         =   "PB_Sort_Order"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Prior_Bids"
         Caption         =   "Prior_Bids"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Bid_Order_Number"
         Caption         =   "Bid_Order_Number"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Bid"
         Caption         =   "Bid"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Bid_Name"
         Caption         =   "Bid_Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "HCP"
         Caption         =   "HCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Dist_Pts"
         Caption         =   "Dist_Pts"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "Cust_Pts"
         Caption         =   "Cust_Pts"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "A_Cust"
         Caption         =   "A_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "K_Cust"
         Caption         =   "K_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "Q_Cust"
         Caption         =   "Q_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "J_Cust"
         Caption         =   "J_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "T_Cust"
         Caption         =   "T_Cust"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "Winrs"
         Caption         =   "Winrs"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column19 
         DataField       =   "Losers"
         Caption         =   "Losers"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column20 
         DataField       =   "QTs"
         Caption         =   "QTs"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column21 
         DataField       =   "Stopped_Suits"
         Caption         =   "Stopped_Suits"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column22 
         DataField       =   "Shape"
         Caption         =   "Shape"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column23 
         DataField       =   "Intermed"
         Caption         =   "Intermed"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column24 
         DataField       =   "xNum"
         Caption         =   "xNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column25 
         DataField       =   "xHCP"
         Caption         =   "xHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column26 
         DataField       =   "sNum"
         Caption         =   "sNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column27 
         DataField       =   "sHCP"
         Caption         =   "sHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column28 
         DataField       =   "hNum"
         Caption         =   "hNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column29 
         DataField       =   "hHCP"
         Caption         =   "hHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column30 
         DataField       =   "dNum"
         Caption         =   "dNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column31 
         DataField       =   "dHCP"
         Caption         =   "dHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column32 
         DataField       =   "cNum"
         Caption         =   "cNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column33 
         DataField       =   "cHCP"
         Caption         =   "cHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column34 
         DataField       =   "cmNum"
         Caption         =   "cmNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column35 
         DataField       =   "cmHCP"
         Caption         =   "cmHCP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column36 
         DataField       =   "Vul"
         Caption         =   "Vul"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column37 
         DataField       =   "Pos"
         Caption         =   "Pos"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column38 
         DataField       =   "Artificial"
         Caption         =   "Artificial"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column39 
         DataField       =   "Alert"
         Caption         =   "Alert"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column40 
         DataField       =   "Explanation"
         Caption         =   "Explanation"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column41 
         DataField       =   "Test_Hand"
         Caption         =   "Test_Hand"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column42 
         DataField       =   "Asks"
         Caption         =   "Asks"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column43 
         DataField       =   "DateOfEntry"
         Caption         =   "DateOfEntry"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column44 
         DataField       =   "Best_Contract"
         Caption         =   "Best_Contract"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column45 
         DataField       =   "Num_Used"
         Caption         =   "Num_Used"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column46 
         DataField       =   "Source"
         Caption         =   "Source"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column47 
         DataField       =   "LHO"
         Caption         =   "LHO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column48 
         DataField       =   "Pard"
         Caption         =   "Pard"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column49 
         DataField       =   "RHO"
         Caption         =   "RHO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column50 
         DataField       =   "ID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   285.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   329.953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   315.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   225.071
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   209.764
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   104.882
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   434.835
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   345.26
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   209.764
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   209.764
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   239.811
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   150.236
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   195.024
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column26 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column27 
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column28 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column29 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column30 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column31 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column32 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column33 
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column34 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column35 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column36 
            ColumnWidth     =   315.213
         EndProperty
         BeginProperty Column37 
            ColumnWidth     =   374.74
         EndProperty
         BeginProperty Column38 
            ColumnWidth     =   255.118
         EndProperty
         BeginProperty Column39 
            ColumnWidth     =   374.74
         EndProperty
         BeginProperty Column40 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column41 
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column42 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column43 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column44 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column45 
            ColumnWidth     =   60.094
         EndProperty
         BeginProperty Column46 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column47 
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column48 
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column49 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column50 
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin VB.Frame f_Compare 
      BackColor       =   &H008080FF&
      Caption         =   "Comparing:  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5085
      Left            =   2000
      TabIndex        =   143
      Top             =   0
      Visible         =   0   'False
      Width           =   5940
      Begin VB.TextBox t_GridData 
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   0
         Left            =   1665
         TabIndex        =   151
         Top             =   1000
         Width           =   4095
      End
      Begin VB.CommandButton b_CompareHelp 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4395
         TabIndex        =   165
         Top             =   4440
         Width           =   1275
      End
      Begin VB.CommandButton b_ComparePutG2Data 
         Caption         =   "Compute Input Boxes for Grid 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2730
         TabIndex        =   157
         Top             =   4020
         Width           =   2955
      End
      Begin VB.CommandButton b_CompareDataOK 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2745
         TabIndex        =   144
         Top             =   4440
         Width           =   1275
      End
      Begin VB.OptionButton o_CompareOptions 
         BackColor       =   &H00FFFF80&
         Caption         =   "Change Grid 2's data."
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
         Index           =   1
         Left            =   150
         TabIndex        =   164
         Top             =   3150
         Width           =   2355
      End
      Begin VB.OptionButton o_CompareOptions 
         BackColor       =   &H00FFFF80&
         Caption         =   "Exit Compare routine."
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
         Index           =   4
         Left            =   150
         TabIndex        =   163
         Top             =   4455
         Width           =   2370
      End
      Begin VB.OptionButton o_CompareOptions 
         BackColor       =   &H00FFFF80&
         Caption         =   "Skip to next grid rows."
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
         Index           =   3
         Left            =   150
         TabIndex        =   162
         Top             =   4020
         Width           =   2355
      End
      Begin VB.OptionButton o_CompareOptions 
         BackColor       =   &H00FFFF80&
         Caption         =   "Skip to next column."
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
         Index           =   2
         Left            =   150
         TabIndex        =   161
         Top             =   3585
         Width           =   2355
      End
      Begin VB.OptionButton o_CompareOptions 
         BackColor       =   &H00FFFF80&
         Caption         =   "Change Grid 1's data."
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
         Index           =   0
         Left            =   180
         TabIndex        =   160
         Top             =   2715
         Value           =   -1  'True
         Width           =   2355
      End
      Begin VB.CommandButton b_CompareChange 
         Caption         =   "Copy Grid 2 data to Grid 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   2730
         TabIndex        =   159
         Top             =   3590
         Width           =   2955
      End
      Begin VB.CommandButton b_CompareChange 
         Caption         =   "Copy Grid 1 data to Grid 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   2730
         TabIndex        =   158
         Top             =   3160
         Width           =   2955
      End
      Begin VB.TextBox t_GridData 
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   1
         Left            =   1665
         TabIndex        =   156
         Top             =   2140
         Width           =   4095
      End
      Begin VB.CommandButton b_CompareUseInputBox 
         Caption         =   "Use Input &Box data in Grid 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2730
         TabIndex        =   154
         Top             =   2730
         Width           =   2955
      End
      Begin VB.TextBox t_GridData 
         BackColor       =   &H0080FFFF&
         Height          =   285
         Index           =   2
         Left            =   1665
         TabIndex        =   152
         Top             =   1570
         Width           =   4095
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   3
         X1              =   5880
         X2              =   5880
         Y1              =   5030
         Y2              =   150
      End
      Begin VB.Label L_Compare 
         BackColor       =   &H00FFFF80&
         Caption         =   "  Input Boxes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   105
         TabIndex        =   155
         Top             =   1005
         Width           =   5700
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF80&
         Height          =   2175
         Index           =   0
         Left            =   2670
         TabIndex        =   167
         Top             =   2670
         Width           =   3075
      End
      Begin VB.Label L_Source 
         BackColor       =   &H00FFFF80&
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
         Height          =   285
         Index           =   2
         Left            =   1665
         TabIndex        =   153
         Top             =   1890
         Width           =   4140
      End
      Begin VB.Label L_CompareSource 
         BackColor       =   &H0080FFFF&
         Caption         =   "          source:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   105
         TabIndex        =   150
         Top             =   1860
         Width           =   1560
      End
      Begin VB.Label L_DataFrom 
         BackColor       =   &H00FFFF80&
         Caption         =   "  Grid 1 Data:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   149
         Top             =   2175
         Width           =   5685
      End
      Begin VB.Label L_CompareSource 
         BackColor       =   &H0080FFFF&
         Caption         =   "          source:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   105
         TabIndex        =   148
         Top             =   1290
         Width           =   1560
      End
      Begin VB.Label L_DataFrom 
         BackColor       =   &H00FFFF80&
         Caption         =   "  Grid 2 Data:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   105
         TabIndex        =   147
         Top             =   1605
         Width           =   1560
      End
      Begin VB.Label L_Source 
         BackColor       =   &H00FFFF80&
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
         Height          =   615
         Index           =   1
         Left            =   1665
         TabIndex        =   146
         Top             =   1305
         Width           =   4140
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         Index           =   2
         X1              =   15
         X2              =   15
         Y1              =   5000
         Y2              =   135
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   0
         X2              =   5880
         Y1              =   5025
         Y2              =   5025
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF80&
         Height          =   255
         Left            =   105
         TabIndex        =   168
         Top             =   945
         Width           =   5700
      End
      Begin VB.Label L_CompareField 
         BackColor       =   &H0080FFFF&
         Caption         =   "field name"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   450
         TabIndex        =   145
         Top             =   300
         Width           =   5355
      End
      Begin VB.Label L_CompareBackground 
         BackColor       =   &H0080FFFF&
         Height          =   4695
         Left            =   105
         TabIndex        =   166
         Top             =   255
         Width           =   5700
      End
   End
   Begin VB.Menu mF 
      Caption         =   "&File"
      Begin VB.Menu mPractice 
         Caption         =   "& Dea/Enterl Hands to Bid"
      End
      Begin VB.Menu mEnter 
         Caption         =   "Enter &Your Own Hands"
         HelpContextID   =   2000
         Shortcut        =   ^Y
      End
      Begin VB.Menu mImport 
         Caption         =   "&Import"
         Visible         =   0   'False
      End
      Begin VB.Menu mBidSim 
         Caption         =   "Bidding Simulator"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mImportPracticeDeal 
         Caption         =   "&Import Practice Deal"
         Shortcut        =   ^I
      End
      Begin VB.Menu z1 
         Caption         =   "-"
      End
      Begin VB.Menu mSysSelect 
         Caption         =   "&Load System"
      End
      Begin VB.Menu mSaveSystem 
         Caption         =   "&Save System"
      End
      Begin VB.Menu mEditConv 
         Caption         =   "&Edit Conventions"
         Shortcut        =   ^{F12}
      End
      Begin VB.Menu z4 
         Caption         =   "-"
      End
      Begin VB.Menu mBackup 
         Caption         =   "&Backup Database"
      End
      Begin VB.Menu mRestore 
         Caption         =   "&Restore Database"
      End
      Begin VB.Menu z2 
         Caption         =   "-"
      End
      Begin VB.Menu mPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu z3 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "&Quit Program"
         HelpContextID   =   2020
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mEd 
      Caption         =   "&Edit"
      Begin VB.Menu mDelete 
         Caption         =   "&Zap (Delete) Entry"
         HelpContextID   =   2110
         Shortcut        =   ^Z
      End
      Begin VB.Menu mRestoreEntry 
         Caption         =   "&Restore Entry"
         Shortcut        =   ^R
      End
      Begin VB.Menu mDivider 
         Caption         =   "Add Dividing &Line"
         Shortcut        =   ^L
      End
      Begin VB.Menu mMake9 
         Caption         =   "Make 90000"
         Shortcut        =   ^G
      End
      Begin VB.Menu zE1 
         Caption         =   "-"
      End
      Begin VB.Menu mFindEntryID 
         Caption         =   "Find Entry ID#"
      End
      Begin VB.Menu mEdit 
         Caption         =   "&Edit Entry"
         HelpContextID   =   2120
         Index           =   0
         Shortcut        =   ^E
      End
      Begin VB.Menu mEdit 
         Caption         =   "Edit &Next Entry"
         Index           =   1
         Shortcut        =   ^N
      End
      Begin VB.Menu mEdit 
         Caption         =   "&Use Entry"
         HelpContextID   =   2120
         Index           =   2
         Shortcut        =   ^U
      End
      Begin VB.Menu mSave 
         Caption         =   "&Save Entry"
         HelpContextID   =   2150
         Shortcut        =   ^S
      End
      Begin VB.Menu mDupeTopEntry 
         Caption         =   "&Dupe Top Entry"
      End
      Begin VB.Menu zE1a 
         Caption         =   "-"
      End
      Begin VB.Menu mPreviewHandSpecs 
         Caption         =   "&Preview Hand Specs"
         Shortcut        =   ^P
      End
      Begin VB.Menu zE2 
         Caption         =   "-"
      End
      Begin VB.Menu mClear 
         Caption         =   "Clear &Inputs"
         HelpContextID   =   2140
         Shortcut        =   ^K
      End
      Begin VB.Menu zE3 
         Caption         =   "-"
      End
      Begin VB.Menu mSetColors 
         Caption         =   "Set Screen Colors"
      End
      Begin VB.Menu zE4 
         Caption         =   "-"
      End
      Begin VB.Menu mCompare 
         Caption         =   "Compare Entries"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mTestHands 
         Caption         =   "Test Entries"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mLookForDeadNoteFiles 
         Caption         =   "Look For Dead Notes"
      End
   End
   Begin VB.Menu mV 
      Caption         =   "&View"
      Begin VB.Menu mVF 
         Caption         =   "&From here down"
         Begin VB.Menu mViewFromHereDown 
            Caption         =   "To end of Prior Bids set"
            Index           =   0
            Shortcut        =   ^F
         End
         Begin VB.Menu mViewFromHereDown 
            Caption         =   "To &Bottom of Grid"
            Index           =   1
            Shortcut        =   ^B
         End
      End
      Begin VB.Menu mNextSection 
         Caption         =   "&Jump to next section"
         Shortcut        =   ^J
      End
      Begin VB.Menu mViewAll 
         Caption         =   "&Overview All"
         Shortcut        =   ^O
      End
      Begin VB.Menu zV0 
         Caption         =   "-"
      End
      Begin VB.Menu mHideCol 
         Caption         =   "&Hide column"
         Shortcut        =   ^H
      End
      Begin VB.Menu mUnhide 
         Caption         =   "&Unhide columns"
      End
      Begin VB.Menu mShowAllCols 
         Caption         =   "Show &All columns"
         Shortcut        =   ^A
      End
      Begin VB.Menu zV1 
         Caption         =   "-"
      End
      Begin VB.Menu mTwoGrids 
         Caption         =   "&Two Grids"
         Checked         =   -1  'True
         Shortcut        =   ^T
      End
      Begin VB.Menu mSplit 
         Caption         =   "&Split top grid"
         Index           =   0
         Shortcut        =   {F11}
      End
      Begin VB.Menu mSplit 
         Caption         =   "Split &bottom grid"
         Index           =   1
         Shortcut        =   {F12}
      End
      Begin VB.Menu mD 
         Caption         =   "&Dupes Grid"
         Begin VB.Menu mDupesGrid 
            Caption         =   "Show &Dupes grid"
            Index           =   0
            Shortcut        =   ^D
         End
         Begin VB.Menu mDupesGrid 
            Caption         =   "Show &Defined Dupes only"
            Index           =   1
         End
         Begin VB.Menu mDupesGrid 
            Caption         =   "Show same Conv.&Name only"
            Index           =   2
         End
      End
      Begin VB.Menu mFourGrids 
         Caption         =   "&Four Grids"
      End
      Begin VB.Menu mMatchWidths 
         Caption         =   "&Match column widths"
         Shortcut        =   ^M
      End
      Begin VB.Menu zV2 
         Caption         =   "-"
      End
      Begin VB.Menu mInputBoxes 
         Caption         =   "&Input Boxes"
         Checked         =   -1  'True
      End
      Begin VB.Menu zV4 
         Caption         =   "-"
      End
      Begin VB.Menu mShowSort 
         Caption         =   "Sho&w/Sort"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mH 
      Caption         =   "&Help"
      Begin VB.Menu mChangeBrowser 
         Caption         =   "&Change browser"
      End
      Begin VB.Menu mHelp 
         Caption         =   "Editor Overview"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mHelpMenus 
         Caption         =   "Editor Menus"
      End
      Begin VB.Menu mHelpInput 
         Caption         =   "Editor Input Boxes"
      End
      Begin VB.Menu mBiddingHelp 
         Caption         =   "Bidding Help Files"
      End
      Begin VB.Menu mViewSAYC 
         Caption         =   "&SAYC Help File"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mGlossary 
         Caption         =   "&Glossary"
      End
      Begin VB.Menu z88 
         Caption         =   "-"
      End
      Begin VB.Menu mFullTutorial 
         Caption         =   "&Full Tutorial"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mMenuTutorial 
         Caption         =   "&Menu Tutorial"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mInputBoxTutorial 
         Caption         =   "&Input Box Tutorial"
      End
      Begin VB.Menu mVersion 
         Caption         =   "&Version"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu mGoBack 
      Caption         =   "<--"
      Enabled         =   0   'False
   End
   Begin VB.Menu mGoFwd 
      Caption         =   "-->"
      Enabled         =   0   'False
   End
   Begin VB.Menu mNextSectionBack2 
      Caption         =   "<="
   End
   Begin VB.Menu mNextSection2 
      Caption         =   "=>"
   End
End
Attribute VB_Name = "BBEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************
'              CardShark BidBase Editor                *
'         Nelson Ford  www.aeyec.com/BidBase           *
'                                                      *
' Purpose: For editing the BidBase bidding database.   *
' Run the Help - Tutorials and read the help files     *
'                for more information.                 *
'                                                      *
' Project - References:                                               *
' The program wouldn't work until I moved these into                  *
' this exact sequence using the Priority Up/Down arrows               *
'       Visual Basic For Applications                                 *
'          c:\windows\syswow64\msvbvm60.dll                           *
'       Visual Basic runtime objectss and procedures                  *
'          c:\windows\syswow64\msvbvm60.dll\3                         *
'       Microsoft Data Environment Instance                           *
'          c:\programming\Miscrosoft Visual Studio\VB98\VB6.OLB       *
'       Microsoft Office 12.0 Access database engine object library   *
'          C:\Program Files (x86)\Common Files\Microsoft Shared ...   *
'             ... \Office12\  ?                                       *
'       Microsoft OLE DB Simple Provider 1.5 Library                  *
'          C:\windows\syswow64\simpdata.db                            *
'       Microsoft Data Binding Collection VB 6.0                      *
'          C:\windows\syswow64\msbind.dll                             *
'       Microsoft AcctiveX Data Objects 6.1 Library                   *
'          C:\Program Files (x86)\Common Files\System\ado\msado15.dll *
'       BidBaseDB:                                                    *
'          Either BidBaseDB.vbp (Language: English/US) is loaded      *
'          or BidBaseDB.DLL (Lang: Standard) is checked as Reference. *
' I just noticed the next 2 in the list on 3-27-2021:                 *
'       Microsoft Data Source Interfaces for ActiveX Data Binding     *
'          Type Library -- MSDATSRC.TLB                               *
'       Microsoft Data Formatting Object Library 6.0 SP6 MSSTDFMT.DLL *
'                                                                     *
'  - Components:                                                      *
'       Microsoft ADO Data Control 6.0 (OLEDB)                        *
'          c:\windows\syswow64\MSADFODC.OCX                           *
'       Microsoft Common Dialog Control 6.0                           *
'          c:\windows\syswow64\ComDlg32.ocx                           *
'       Microsoft Data Bound List Control 6.0                         *
'          c:\windows\syswow64\DBList32.OCX                           *
'       Microsoft DataGrid Control 6.0                                *
'          c:\windows\syswow64\MSDATGRD.OCX                           *
'       Microsoft FlexGrid Control 6.0                                *
'          c:\windows\syswow64\MSFLXGRD.OCX                           *
'       Microsoft Sysinfo Control 5.0                                 *
'          c:\windows\syswow64\SYSINFO.OCX                            *
'       Microsoft Windows Common Controls 6.0                         *
'          c:\windows\syswow64\MXCOMCTL.OCX                           *
'                                                      *
' Notes:  Written in VB6                               *
'           In VB6 under Tools, Options, Editor Format *
'           I use the following screen colors:         *
'               Normal Text - cyan on black            *
'               Selection Text - dark blue             *
'               Syntax Error - red on white            *
'               Execution Point - black on yellow      *
'               Breakpoint - white on red              *
'               Comment Text - white on dark cyan      *
'               Keyword Text - green on black          *
'               Identifier Text - yellow on black      *
'         You can set colors to your own taste, but I  *
'         STRONGLY recommend something similar to make *
'         it easier to spot various types of code, etc.*
'*******************************************************

Option Explicit ' All variables must be explicitly declared. '
Option Compare Binary ' In string comparisons, "A" <> "a".   '

Private str_CmdLineArg(2) As String ' Other pgms can pass hands to the editor.'

Private Const Version$ = "2025.07.27"
' 13,500+ lines of actual code. '

Private DebugMode As Boolean
' Change file name to Editor debug.exe '

' See if program is running in the background: '
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_CLOSE = &H10

'Database:
Private db_BidBase As Database
Private ws_BidBase As Workspace
Private rs_BidTable As Recordset
Private rs_BidTableTemp As Recordset ' rs_BidTable is set in Sub InitializeDatabase; this creates one for other uses. '
Private rs_BidNamesTbl As Recordset
Private rs_DelTable As Recordset
Private rs_DupePBs As Recordset
Private rs_BidsQuery As Recordset
Private rs_HandsData As Recordset
'The following allows changing the name of the database.
Private str_Database As String
Private str_BidsTbl As String
Private str_BidNamesTbl As String
Private str_GarbageTbl As String
Private str_HandsSpecsTbl As String
Private str_DupePBsTbl As String
Private str_DataSource As String
Private str_DBPathAndName ' = App.Path & "\BidBase.accdb"
Private qSel As String ' Query for Data Control RecordSource's SELECT '
'For calling DB module:
Private NewBid As BidBaseDB.bidder ' If error message, check Project - References for BidBaseDB.vbp or .DLL '

'General program variables:
Private bln_Abort As Boolean
Private bln_AlreadyToldNoDupes As Boolean
Private str_BBpath As String
Private int_BGInputs As Integer ' Background colors for Input Boxes. '
Private int_BGForm As Integer ' Background colors for form. '
Private bln_BidFound As Boolean
Private str_Bids(24) As String
Private str_Browser As String ' Path and name of browser. '
Private str_SuitName(4) As String
Private int_cmbConvName_Index As Integer
Private int_CurrentHandDataCell As Integer
Private int_CurHandDataRow As Integer
Private int_CurHandDataCol As Integer
Private str_DataPath As Integer
Private bln_Dragging As Boolean
Private sng_DragoffsetX As Single
Private sng_DragoffsetY As Single
Private str_DupePrior As String
Private lng_DupeID As Long
Private int_DupeGrid As Integer
Private bln_DupingEntry As Boolean
Private lng_EntryID As Long
Private str_ErrMsg As String
Private int_FGInputs As Integer ' Foreground colors for Input Boxes. '
Private int_FGForm As Integer ' Foreground colors for Form. '
Private int_FindHandNum As Integer
Private bln_FirstTime As Boolean ' Flag to show if first run for user. '
Private bln_GoBack As Boolean
Private bln_InputBoxes As Boolean
Private bln_ShowGrid2 As Boolean ' For the Grid2 View button in the Show window. '
Private int_Suit(4, 4, 13) As Integer
Private int_SuitQty(4, 4) As Integer
Private str_Job As String ' Identifies what task is being done when   '
                          ' using common routines for different tasks '
Private int_KnownFitLen As Integer
Private str_LastPBs As String
Private int_1NT_Min As Integer
Private int_2NT_Min As Integer
Private str_ImportSource As String
Private int_MsgClick As Integer
Private str_MsgShown As String
Private bln_NoHiLite As Boolean
Private bln_NoHiLitePrior As Boolean
Private int_NoPaint As Integer ' Prevents repainting of forms
Private str_NoteFile(1) As String ' For opening HTML files '
Private bln_OrigData As Boolean ' Value of ck boxes at Got_Focus
Private str_OrigData As String ' Data in text boxes at Got_Focus
Private str_PgmPath(1) As String ' Program's path. '
Private str_PriorBids As String
Private str_Reference As String
Private str_Response As String ' Response to MyMsgBox '
Private int_ScrollVal As Integer ' Keeps track of Value of Vscroll in BBEditor
Private bln_SkipIt As Boolean ' Flag for skipping routines. '
Private bln_StartUp As Boolean ' Flag for program startup. '
Private bln_TestingSaveEntry As Boolean ' Make sure test hand matches entry specs. '
Private bln_TestingSection As Boolean
Private str_TestingHand As String ' For two tests above. '
Private str_TutorDetails As String ' Help file name '
Private str_TutorSubject As String ' For selecting tutorial subject
Private bln_Waiting As Boolean ' See Sub Wait
Private str_WebEditor As String ' Path and name of web page editor. '

'Bid entry variables:
Private lng_BON As Long ' Bid Order Number '
Private lng_ID As Long ' ID number of current entry
Private lng_CurID(4) As Long
Private int_CurRow(4) As Integer
Private lng_IncrSeqAmt  As Long ' Amount for incrementing bid order number
Private int_LastBidCol As Integer ' Last column in the edit grids
Private str_PriorCalls As String
Private str_Source As String ' Source of user-entered deal.

'Grid variables:
Private int_BGGrid1 As Integer ' grid background colors #1 '
Private int_BGGrid2 As Integer ' grid background colors #2 '
Private sng_ColWidths(100) As Single
Private int_DividerLineID As Integer ' database ID# in BidNames for "======="
Private lng_FGGrid1 As Long ' grid foreground colors #1 '
Private lng_FGGrid2 As Long ' grid foreground colors #2 '
Private str_GridPick(4, 2, 50) As String ' Grid specification history. '
        ' 0-4 = Grid #, 0-2 = 1:t_Show(start) 2:t_Show(Stop) 3:cmb_Show
Private int_GridPicksNum As Integer ' Pointer to place in grid spec. history.'
Private bln_GridFwdBack As Boolean ' Flag for using Fwd/Back arrows.'
Private int_GridSplitWidth1(1) As Integer
Private int_GridSplitWidth2(1) As Integer
Private int_GridNum  As Integer
Private bln_Scrolling As Boolean
Private int_SortIndex(4) As Integer
Private str_StartGridRow
Private lng_TopID(4) As Long ' ID # of top entry in grid
Private bln_TwoGrids As Boolean
Private bln_UsingEntry As Integer
Private bln_ViewGarbage As Boolean
Private bln_Working As Boolean
Private bln_NoResize As Boolean
'Grid column pointers:
Private int_Pct_Col As Integer
Private int_Round_Col As Integer
Private int_Player_Col As Integer
Private int_PB_Sort_Ord_Col As Integer
Private int_Prior_Bids_Col As Integer
Private int_OrdNum_Col As Integer
Private int_Bid_Col As Integer
Private int_Bid_Name_Col As Integer
Private int_HCP_Col As Integer
Private int_xNum_Col As Integer
Private int_Combo_Col As Integer
Private int_Vul_Col As Integer
Private int_Expl_Col As Integer
Private int_TestHand_Col As Integer
Private int_Source_Col As Integer
Private int_ID_Col As Integer
Private lng_X As Long
Private lng_Y As Long
Private lng_Top As Long
Private lng_Left As Long
Private bln_Down As Boolean
' Hide scrollbar on help bar at bottom of Input Boxes. '
Private Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Private Const SB_VERT = 1
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Function LeftButton() As Boolean
    LeftButton = (GetAsyncKeyState(vbKeyLButton) And &H8000)
End Function


'*******************************************
' Procedure: DisplayFindHand               '
' Purpose:                                 '
    ' When the user searches for a bid,    '
    ' if one is not found, he has the menu '
    ' option of entering the search hand   '
    ' into the Input Boxes, along with its '
    ' specs, to make it easy to create a   '
    ' new entry. This routine does that.   '
'*******************************************

Private Sub DisplayFindHand()
    Dim HCPs(4) As Variant
    Dim qty(4) As Variant
    Dim valSys As String
    Dim valPts(2) As Variant ' 1=Suit, 2=NT valuation '
    Dim distPts(4) As Variant
    Dim winners(4) As Variant
    Dim losers(4) As Variant
    Dim QTs(4) As Variant
    Dim stoppers(4) As Variant
    Dim intermeds(4) As Variant
    Dim xx As Variant
    Dim s As Integer
    Dim bidRound As Integer
    Dim posNum As Integer
    Dim i As Integer
    Dim CP As Integer
    Dim p As Integer
    Dim sp As String

    Set NewBid = New BidBaseDB.bidder
    ' If "Class not registered", open BidBaseDB.dll with  '
    ' regsvr32.exe in xplorer2. For more information, see '
    ' the end of CardShark_BidBase_Programming_Notes.htm. '

    With NewBid
    '''''''''''
        .DBPathAndName = str_DBPathAndName 'str_BBpath & "\" & str_Database
        .PlayersHand = Trim$(t_FindHand(int_FindHandNum))
        .PriorCalls = Trim$(t_FindPrior)
        .BidToAnalyze = Trim$(t_FindAnalysis4Bid)
        'xx = .HandData(HCPs(), _
                       qty(), _
                       distPts(), _
                       winners(), _
                       losers(), _
                       QTs(), _
                       stoppers(), _
                       intermeds(), _
                       bidRound, _
                       posNum, _
                       valSys, _
                       valPts() _
                       )
    End With
    ''''''''

    ' Clear the Input Boxes: '
    mClear_Click

    ' Count starting passes: '
    sp = Trim$(t_FindPrior) & Space$(5)
    If Left$(sp, 5) = "P-P-P" Then
        i = 6
        p = 3
    ElseIf Left$(sp, 3) = "P-P" Then
        i = 4
        p = 2
    ElseIf Left$(sp, 1) = "P" Then
        i = 2
        p = 1
    End If

    If p > 0 Then
        If i = Len(t_FindPrior) + 1 Then
            cmb_Position.ListIndex = p + 3
        End If
        t_Prior = Trim$(Mid$(t_FindPrior & "  ", i))
    Else
        t_Prior = t_FindPrior
    End If

    t_TestHand = t_FindHand(int_FindHandNum)
    t_Round = bidRound
    t_Player = posNum

    i = InStr(HCPs(0), ",")
    If i = 0 Then
        t_Hand_HCPs = HCPs(0)
    Else
        t_Hand_HCPs = Left$(HCPs(0), i - 1)
        t_BBPoints = Mid$(HCPs(0), i + 1)
    End If

    t_CustPts(0) = "2"
    t_CustPts(1) = "1"
    For i = 1 To Len(t_TestHand)
        Select Case Mid$(t_TestHand, i, 1)
          Case Is = "A"
            CP = CP + 2
          Case Is = "K"
            CP = CP + 1
        End Select
    Next
    t_CustPoints = "" & CP

    t_Distrib_Pts = distPts(0)
    t_Winners = winners(0)
    t_Losers = losers(0)
    t_QTs = QTs(0)
    t_Stoppers = stoppers(0)
    t_Intermediates = intermeds(0)

    s = 4
    For i = 1 To 4
        t_Qty(i) = qty(s)
        t_HCP(i) = HCPs(s)
        s = s - 1
    Next

    cmb_Vul.ListIndex = cmb_FindVul.ListIndex

    CreateDistribDisclosure
End Sub


Private Function Sort_Order_Specs() As String
    ' Puts the specified database sort parameters '
    ' into a variable to use in a database call.  '
    ' The sort order is specified in the combo    '
    ' boxes in the Show/Sort box.                 '

    Dim i As Integer
    Dim s(4) As String

    s(0) = ""
    For i = 0 To 4
        ' If the user has specified Prior_Bids in the '
        ' sort specs, we change that to PB_Sort_Order '
        ' which is what we really have to sort by to  '
        ' get Prior_Bids into the correct order. For  '
        ' an explanation, see the doc file. '
        ' UPDATE 11-02-22: Never mind. See CSBB Programming Notes.htm. '
        If 1 = 2 And cmb_Sort(i).Text = "Prior_Bids" Then
            s(i) = "Bids.PB_Sort_Order"
        ElseIf cmb_Sort(i).Text = "Bid_Name" Then
            s(i) = "BidNames.Bid_Name"
        ElseIf cmb_Sort(i).Text = "Pct_Used" Then
            s(i) = "BidNames.Selected, BidNames.SubCat_Selected, Bids.Pct_Used"
        ElseIf InStr(t_Show(0), "@") = 3 And i = 0 Then
            'Stop
            s(i) = "Bids." & t_Show(0)
        ElseIf cmb_Sort(i).Text <> "" Then
            s(i) = "Bids." & cmb_Sort(i)
        End If

        If i > 0 And s(i) <> "" Then
            ' Concatenate sort fields: '
            s(0) = s(0) & ", " & s(i)
        End If
    Next

    Sort_Order_Specs = s(0)
End Function

'*************************************************
' Procedure: FindConvenNameFile                  '
' Purpose:                                       '
    ' Removes illegal filename characters from a '
    ' filename provided by calling routine.  We  '
    ' also display Y or N on the View Convention '
    ' button to indicate if a file is available. '
'*************************************************

Private Sub FindConvenNameFile(ByVal f As String, _
                       ByVal fixButton As Boolean, _
                       ByVal displayFile As Boolean, _
                       ByVal createFile As Boolean)
    Dim f1 As String
    Dim f2 As String
    Dim i As Integer
    Dim pth As String
    Dim xx As Long
    Dim FiNum As Integer

    f1 = f

    ' Remove sub-description, if any: '
    '   (e.g.: "Raise: Preemptive")   '
    i = InStr(f, ":")
    If i > 0 Then
        f = Left$(f, i - 1)
    End If

    f = Replace$(f, "2/1", "2-Over-1")
    If InStr(f, "Stayman") > 0 Or InStr(f, "Jacoby Transfer") > 0 Or f = "Texas Transfer" Then
        f = "Notrump_Bidding"
    End If
    
    ' Remove non-alpha-numerics from filename:
    f = LegalFileName(f)

    pth = str_BBpath & "\Notes\"
    f2 = f & ".htm"
    If Dir$(pth & f2) = "" Then
        f2 = Replace$(f2, " ", "_")
        If Dir$(pth & f2) = "" Then
            If Dir$(pth & f & ".txt") <> "" Then
                Name pth & f & ".txt" As pth & f2
            End If
        End If
    End If

    ' Set the "Display File" button to "Y" or "N" '
    ' according to if a convention file is found: '
    If fixButton Then
        If Dir$(pth & f2) = "" Then
            b_Conven_Description.Picture = _
                LoadPicture(str_BBpath & "\ButtonN.bmp")
        Else
            b_Conven_Description.Picture = _
                LoadPicture(str_BBpath & "\ButtonY.bmp")
        End If
    End If

    If Not displayFile Then
        Exit Sub
    End If

    ' Display file:                                                 '

    If Dir$(pth & f2) = "" Then
        If createFile Then
            ' No file, so create one: '
            FiNum = FreeFile
            Open pth & f2 For Output As #FiNum
            Print #FiNum, "<HTML><HEAD>"
            Print #FiNum, "<TITLE>CardShark BidBase - "; _
                          Trim$(cmb_ConvName(0).Text); "</TITLE>"
            Print #FiNum, "</HEAD>"
            Print #FiNum, "<BODY TEXT=""#000000"" "
            Print #FiNum, "  BGCOLOR=""#FFFFFF"" "
            Print #FiNum, "  LINK=""#0000EE"; ""
            Print #FiNum, "  VLINK=""#008000"" "
            Print #FiNum, "  ALINK=""#FF0000"">"
            Print #FiNum, "<CENTER>"
            Print #FiNum, "<H1>"; Trim$(f); "</H1>"
            Print #FiNum, "<P>"
            Print #FiNum, "<TABLE WIDTH=90%><TR><TD>"
            Print #FiNum, "<FONT SIZE=+1>"
            Print #FiNum, ""
            Print #FiNum, ""
            Print #FiNum, ""
            Print #FiNum, "<BR><BR><BR><BR><BR>"
            Print #FiNum, "<A NAME=EndOfFile>"
            Print #FiNum, "</TD></TR></TABLE>"
            Print #FiNum, "<P></BODY></HTML>"
            Close #FiNum
        Else
            str_Response = MyMsgBox("No file is present for the convention: " & vbCr & _
                                    f, 2, "ok", "")
            Exit Sub
        End If
    End If

    str_NoteFile(0) = f2
    str_NoteFile(1) = ""
    str_PgmPath(0) = str_BBpath & "\Notes\"
    str_PgmPath(1) = ""
    ChDir str_BBpath
    If str_Browser <> "" Then
        xx = Shell(str_Browser & " " & str_PgmPath(0) & str_NoteFile(0), _
                   vbNormalFocus)
    Else
        xx = Shell(str_PgmPath(0) & "BidBase File Viewer.exe " & _
                   str_PgmPath(0) & str_NoteFile(0), _
                   vbNormalFocus)
    End If

End Sub

'*******************************************************
' Procedure: LookForNoteFile                           '
' Purpose:                                             '
    ' Each entry may have its own reference/note file. '
    ' The file name is made up of the prior bids, the  '
    ' current bid, and the entry's ID#. For example:   '
    ' "1C-1D-D_ID263.txt".  If such a file exists, the '
    ' Explanation field should say: "See notes -->"    '
    ' where the arrow is pointing to the button for    '
    ' bringing up the file.  If no note file specific  '
    ' to the entry exists, the Explanation field may   '
    ' reference another file. In that case, it should  '
    ' say: "[See file: " followed by the file name.    '
    '                                                  '
    ' If the user has clicked "Use Entry" in the menu, '
    ' and if the entry being used has "See notes -->"  '
    ' in the note field, we have to change that for the'
    ' new entry to "See file: " followed by the name.  '
    '                                                  '
    ' If "setButton" is True, then set the "See Notes" '
    ' button to "Y" or "N", as appropriate.            '
    '                                                  '
    ' If "showFile" is True, then create a file if none'
    ' exists and add the reference to the note field   '
    ' and then call the file viewer to display the file.
'*******************************************************

Private Sub LookForNoteFile(discl As String, _
                    priorBid As String, _
                    thisBid As String, _
                    thisID As String, _
                    useEntry As Boolean, _
                    setButton As Boolean, _
                    showFile As Boolean)

    Dim n As Integer
    Dim n2 As Integer
    Dim f As String
    Dim fi As String
    Dim fWithoutIL As String ' File name without internal link (e.g.: #x), if any, at the end. '
    Dim p As String
    Dim i As Integer
    Dim j As Integer
    Dim t As String

    ' Look for note file reference in Disclosure field. '
    ' This will be either just "See notes --> " or      '
    ' "See file: " followed by a file name.             '
    ' If "useEntry" is True, and "See notes --> " is in '
    ' the Disclosure field, then we have to generate the'
    ' file name using the original entry's Prior Bids,  '
    ' current Bid, and ID#, which is in rs_BidTable!*   '

    p = str_BBpath & "\Notes\EntryNotes\"

    n = InStr(LCase$(discl), "see notes:")
    If n > 0 Then
        i = InStr(discl, ".htm")
        If i > 0 Then
            j = InStrRev(discl, ">", i)
            If j > 0 Then
                f = Trim$(Mid$(discl, j + 2))
                i = InStr(j + 4, f, "#")
                If i > j Then
                    fWithoutIL = f ' Change this to fWithIntLink '
                    f = Left$(f, i - 1)
                End If
            End If
        End If
    End If

    If useEntry And n > 0 And f = "" Then
        ' "[See notes -->" is in the Disclosure field, '
        ' but refers to a file with the ID# of the     '
        ' entry being used to set up a new entry, so   '
        ' the reference for the new entry must change  '
        ' from "[See notes -->"                        '
        '   to "[See file 1D-1H_ID123.txt" (for example)
        ' where "ID123" is the original entry's ID#.   '

        If int_GridNum = 1 Then
            f = Grid1(1).Columns(int_Prior_Bids_Col)
            fi = Grid1(1).Columns(int_Bid_Col)
        Else
            f = Grid1(0).Columns(int_Prior_Bids_Col)
            fi = Grid1(0).Columns(int_Bid_Col)
        End If

        If Trim$(priorBid) = "" Then
            f = fi & "_ID" & thisID
        Else
            f = f & "-" & fi & "_ID" & thisID
        End If

        If n = 1 Then
            t = "[See file: " & f & ".txt"
            If Len(discl) > 14 Then
                t_Disclosure = t & Mid$(discl, 15)
            Else
                t_Disclosure = t
            End If
        Else
            If Len(Left$(discl, n - 1) & "[See file: " & f & ".txt") <= 50 Then
                t_Disclosure = Left$(discl, n - 1) & _
                               "See file: " & f & ".txt"
            ElseIf Len(Left$(discl, n - 1) & f & ".txt") <= 50 Then
                t_Disclosure = Left$(discl, n - 1) & _
                               f & ".txt"
            End If
            ' Could be some text after the file reference,    '
            ' but there is not enough room to put it back in. '
        End If
    End If


    If n = 0 And f = "" Then
        ' See if file is referenced by file name: '
        ' (Ideally, the user should put the words '
        ' "See file:" followed by the file name,  '
        ' but there may not always be room, so    '
        ' look for the .txt and .htm extensions.  '
        ' A file will most often have an extension'
        ' of .txt, so we look for that first, but '
        ' then we also look for .htm, and if that '
        ' is not found, then just to be thorough, '
        ' we look for the phrase "See file:" and  '
        ' assume that some other file name comes  '
        ' after that.)                            '
        n2 = InStr(LCase$(discl), ".txt")
        If n2 = 0 Then
            n2 = InStr(LCase$(discl), ".htm")
            If n2 = 0 Then
                n2 = InStr(LCase$(discl), "see file:")
            End If
        End If

        If n2 > 1 Then
            If Mid$(discl, n2, 1) = "." Then
                ' .txt or .htm was found: '
                i = n2 - 1
                Do While Mid$(discl, i, 1) <> " " And _
                         Mid$(discl, i, 1) <> "[" And _
                         i > 1
                    i = i - 1
                Loop
                ' e.g.: "Blah, blah. [1C-1S_ID12324.txt" '
                '        123456789-123456789-123456789-  '
                '        .txt = 27 = n2                  '
                '        len of filename = 18            '
                '        [ = 13 = i                      '
                '        n2 -  i + ext length            '
                '        27 - 13 + 4 = 18 = filename len '
                f = Trim$(Mid$(discl, i, n2 - i + 4))
                j = InStr(discl, "#")
                If j > 0 Then
                    f = f & Mid$(discl, j)
                End If
            Else
                ' "See file" is in the Disclosure field. '
                ' Extract the name of the file:          '
                i = n2 + 9
                j = i + 5
                Do While Mid$(discl & " ", j, 1) <> " "
                    j = j + 1
                    DoEvents
                Loop
                f = Trim$(Mid$(discl, i, j - i))
            End If
        End If

        If Left$(f, 1) = "[" Then f = Trim$(Mid$(f, 2))

        ' Verify file's presence: '
        If f <> "" And Left$(f & "    ", 4) <> "http" And Left$(f & "    ", 4) <> "www." Then
            If Dir$(p & f) = "" Then
                i = InStr(discl, "#")
                If i > 0 And InStr(f, "#") = 0 Then
                    Dim fWithIL As String
                    fWithIL = f & Mid$(discl, i)
                    'f = Left$(f, i - 1)
                End If

                If Right$(p, 11) = "EntryNotes\" Then
                    p = str_BBpath & "\Notes\"
                    If Dir$(p & f) = "" And Dir$(p & "..\") = "" Then
                        str_Response = MyMsgBox("Please make a note that the Disclosure note field " & vbCrLf & _
                                                "refers to '" & f & "' but that file is " & vbCrLf & _
                                                "not in the NOTES folder, so the file reference " & vbCrLf & _
                                                "should be removed from the Disclosure note field.", 4, "ok", "")

                        If setButton Then
                            b_Discl.Picture = LoadPicture(str_BBpath & "\ButtonN.bmp")
                        End If
                        Exit Sub
                    End If
                End If
            End If

        Else
            ' A file is not referenced by name '
            ' in the Disclosure field.         '

            If Trim$(priorBid) = "" Then
                f = thisBid & "_ID" & thisID
            Else
                f = priorBid & _
                    "-" & thisBid & _
                    "_ID" & thisID
            End If

        End If

    ElseIf f = "" Then
            ' A file is not referenced by name '
            ' in the Disclosure field.         '

            If Trim$(priorBid) = "" Then
                f = thisBid & "_ID" & thisID
            Else
                f = priorBid & _
                    "-" & thisBid & _
                    "_ID" & thisID
            End If

    End If

    ' Verify file's presence: '
    If Dir$(p & f & ".txt") = "" And Dir$(p & "..\" & f) = "" And _
       Dir$(p & f & ".htm") = "" And Dir$(p & f) = "" _
    Then
        If n > 0 Then
            str_Response = MyMsgBox("Please make a note that the Disclosure " & vbCr & _
                         "note field refers to the file: " & vbCr & _
                         f & vbCr & _
                         " but that file is not in the NOTES folder, " & vbCr & _
                         "so the file referenced should be removed" & vbCr & _
                         "from the Disclosure note field.", 6, "ok", "")
        End If

        If setButton Then
            ' Again - it's being created below, so why this? '
            b_Discl.Picture = LoadPicture(str_BBpath & "\ButtonN.bmp")
        End If


        If Not showFile Then
            Exit Sub

        Else
            If setButton = False Then
                ' Disclosure field in the grid was double-clicked. '
                ' We don't want to create a new file in that case  '
                ' because the file is supposed to already be there.'
                str_Response = MyMsgBox("The file " & f & ".txt" & vbCr & _
                             "is not in the Notes folder.", 2, "ok", "")
                Exit Sub
            End If

            ' Since "showFile" is True, create  '
            ' the file so that it can be shown: '

            If ck_SaveNew.Value <> 0 Or lng_ID = 0 Then
                str_Response = MyMsgBox("The entry must be saved before" & vbCrLf & _
                             "creating a file for it.", 2, "ok", "")
                Exit Sub
            End If

            If n = 0 And InStr(t_Disclosure, "See notes") = 0 Then
                If InStr(t_Disclosure, "[") > 0 Then
                    t_Disclosure = t_Disclosure & _
                                   ". See notes -->"
                Else
                    t_Disclosure = "[See notes -->"
                End If
            End If

            Dim FiNum As Integer

            FiNum = FreeFile

            Open p & f & ".txt" For Output As #FiNum
            Print #FiNum, Tab(40); f & ".txt"
            Print #FiNum, ""
            Print #FiNum, t_TestHand
            Print #FiNum, ""
            Print #FiNum, priorBid
            Print #FiNum, thisBid
            Print #FiNum, ""
            Close #FiNum
        End If
    End If

    i = InStr(f, "#")
    fi = f
    If i > 2 Then
        f = Left$(f, i - 1)
    End If
    
    If Dir$(p & f) = "" Then
        If Dir$(p & f & ".txt") <> "" Then
            f = f & ".txt"
        ElseIf Dir$(p & f & ".htm") <> "" Then
            f = f & ".htm"
        End If
    End If

    If setButton Then
        b_Discl.Picture = LoadPicture(str_BBpath & "\ButtonY.bmp")
    End If

    If Not showFile Or fi = "" Then
        Exit Sub
    End If

    ' Call BB File Viewer to show the file: '

    Dim xx As Long

    str_NoteFile(0) = f 'fi
    str_NoteFile(1) = ""
    str_PgmPath(0) = p
    str_PgmPath(1) = ""
    If str_Browser = "" Then
        str_Browser = "D:\Desktop\VB98\Bridge\BidBase File Viewer.exe"
    End If
    If str_Browser <> "" Then
        If str_WebEditor <> "" Then
            xx = Shell(str_Browser & " " & _
                       str_PgmPath(0) & str_NoteFile(0), _
                       vbNormalFocus)
            xx = Shell(str_WebEditor & " " & _
                       str_PgmPath(0) & str_NoteFile(0), _
                       vbNormalFocus)
        Else
            xx = Shell(str_Browser & " " & str_PgmPath(0) & str_NoteFile(0) _
                        & "" _
                        & "", _
                        vbNormalFocus)
        End If
    ElseIf Dir$(str_BBpath & "\Notes\" & fi) <> "" Then
        xx = Shell(str_BBpath & "\BidBase File Viewer.exe " & _
                    str_BBpath & "\Notes\" & fi, vbNormalFocus)
    Else
        If Dir$(str_BBpath & "\BidBase File Viewer.exe") = "" Then
            MsgBox "BidBase File Viewer.exe is not in the BidBase directory."
            Exit Sub
        End If
        xx = Shell(str_BBpath & "\BidBase File Viewer.exe " & _
                    str_PgmPath(0) & str_NoteFile(0), vbNormalFocus)
    End If

End Sub


Private Sub LoadBidNames()
    ' Bid names are stored in the BidBase  '
    ' data file in their own recordset.    '
    ' The recordset is opened in Form Load '
    ' A DataCombo box is being used, which '
    ' requires .AddItem to populate.  The  '
    ' DBCombo box does not, but it caused  '
    ' other problems.                      '
    On Local Error GoTo BidNamesErr
    If cmb_ConvName(0).ListCount < 10 Then
        ' Only do this once, at start:
        AutoSizeDropDownWidth cmb_ConvName(0)
        AutoSizeDropDownWidth cmb_ConvName(1)
    Else
        cmb_ConvName(0).Clear
        cmb_ConvName(1).Clear
    End If

    Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
    Set rs_BidNamesTbl = db_BidBase.OpenRecordset(str_BidNamesTbl)
    With rs_BidNamesTbl
    '''''''''''''''''''
        .MoveFirst
        Do
            If Left$(!Bid_Name & "     ", 5) = "=====" Then
                int_DividerLineID = !IDNum

            Else
                cmb_ConvName(0).AddItem !Bid_Name
                cmb_ConvName(1).AddItem !Bid_Name
            End If
            If Not .EOF Then
                .MoveNext
            End If
            DoEvents
        Loop While Not .EOF
    End With
    ''''''''
Exit Sub
On Local Error GoTo 0

BidNamesErr:
    MsgBox "Error in Sub LoadBidNames."
End Sub


'*******************************************************
' Procedure: Welcome                                   '
' Purpose:                                             '
'     This begins the program tutorial. The first time '
'     the program is run on a machine, this tutorial   '
'     runs, but then a flag is saved in the .CFG file  '
'     so that it doesn't run automatically again. The  '
'     tutorial can be run manually via the Help menu.  '
'*******************************************************

Private Sub Welcome()

    frm_Tutorial.Visible = True
    frm_Tutorial.ZOrder 0
    frm_Tutorial.Caption = "CardShark BidBase Tutorial"

    L_Tutorial = "Welcome to CardShark BidBase, the bidding database. " & _
                 vbCr & vbCr & _
                 "The program has extensive documentation which " & _
                 "can be read by pressing F1. This tutorial will " & _
                 "give you a quick overview of how the program works." & _
                 vbCr & vbCr & _
                 "To bring up a more detailed description of a feature " & _
                 "in the documentation, click the 'Details' button " & _
                 "when the feature is being described." & _
                 vbCr & vbCr & _
                 "Click-and-drag the top of this window to move it or " & _
                 "the bottom of the window to resize it. Try it now." & _
                 vbCr & vbCr & _
                 "Thank you! If the text before this line fits in the " & _
                 "window, you should not have to adjust the size again. " & _
                 "Now press Continue, please."

    frm_Tutorial.Move 5575, 600
    frm_Tutorial.Width = 6300
    L_Tutorial.Height = 2950
    L_Tutorial.Width = 5900
    b_TutorDetails.Visible = True
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm"
    ResizeTutorial
    DoEvents
    pTutorDragBar.Refresh
    Wait
    If frm_Tutorial.Visible = False Then Exit Sub

    ' View menu tutorial: '

    Run_ShowSort_Tutorial
    If frm_Tutorial.Visible = False Then Exit Sub

    'Run_DupesGrid_Tutorial
    'Wait
    'If frm_Tutorial.Visible = False Then Exit Sub
    'If mDupesGrid(0).Checked Then mDupesGrid_Click 0

    Run_HereDown_Tutorial
    Wait
    If frm_Tutorial.Visible = False Then Exit Sub

    Run_NextSection_Tutorial
    Wait
    If frm_Tutorial.Visible = False Then Exit Sub

    Run_OverviewAll_Tutorial
    ' This is view all entries, not view all columns. '
    Wait
    If frm_Tutorial.Visible = False Then Exit Sub

    Run_HidingColumns_Tutorial
    ' Also includes Unhide and Show All Cols '
    Wait
    If frm_Tutorial.Visible = False Then Exit Sub

    Run_TwoGrids_Tutorial
    Wait
    If frm_Tutorial.Visible = False Then Exit Sub

    Run_SplitGrids_Tutorial
    Wait
    If frm_Tutorial.Visible = False Then Exit Sub

    Run_FourGrids_Tutorial
    Wait
    If mFourGrids.Checked Then mFourGrids_Click
    If frm_Tutorial.Visible = False Then Exit Sub

    Run_MatchCols_Tutorial
    Wait
    If frm_Tutorial.Visible = False Then Exit Sub

    Run_HideInputs_Tutorial
    Wait
    If frm_Tutorial.Visible = False Then Exit Sub

    ' Back/Fwd Arrows: '

    frm_Tutorial.Caption = "Go Forward In Grid"
    Run_GoFwd_Tutorial
    Wait
    If frm_Tutorial.Visible = False Then Exit Sub

    ' Edit Menu Tutorial: '

    Run_DeleteEntry_Tutorial
    Wait
    If frm_Tutorial.Visible = False Then Exit Sub

    Run_Restore_Entry
    If frm_Tutorial.Visible = False Then Exit Sub

    Run_Divider_Tutorial
    Wait
    If frm_Tutorial.Visible = False Then Exit Sub


    Run_Edit_Tutorial
    ' This runs several other tutorials for the '
    ' Input Boxes, so no Wait is needed here.   '
    If frm_Tutorial.Visible = False Then Exit Sub

    Run_UseEntry_Tutorial
    Wait
    If frm_Tutorial.Visible = False Then Exit Sub

    Run_PreviewHandSpecs_Tutorial
    Wait
    f_EnterHand.Visible = False
    If frm_Tutorial.Visible = False Then Exit Sub

    Run_ClearInputs_Tutorial
    Wait
    If frm_Tutorial.Visible = False Then Exit Sub

    f_GridColors.Left = 2222
    f_GridColors.Top = 2222
    f_GridColors.Visible = True
    mSetColors.Checked = True
    Run_SetColors_Tutorial
    If f_GridColors.Visible = True Then
        frm_Tutorial.Left = f_GridColors.Left + _
                            f_GridColors.Width
        frm_Tutorial.Top = f_GridColors.Top
    End If
    Wait
    f_GridColors.Visible = False
    If frm_Tutorial.Visible = False Then Exit Sub

    Run_CompareEntries_Tutorial
    Wait
    If frm_Tutorial.Visible = False Then Exit Sub

    Run_TestHands_Tutorial
    Wait
    If frm_Tutorial.Visible = False Then Exit Sub

    Run_LookForDeadNotes_Tutorial
    Wait
    If frm_Tutorial.Visible = False Then Exit Sub


    L_Tutorial = "This concludes the 'Full Tutorial'. " & _
                 "You may click 'Help - Menu Tutorial' at any time " & _
                 "to click on any menu item for help for that item, " & _
                 "or click on 'Help - Input Box Tutorial' for a " & _
                 "a review of the Input Boxes, or " & _
                 "press F1 to bring up the Help files." & _
                 vbCr & vbCr & _
                 "Expect to have to use the program for a little while " & _
                 "before it all becomes second nature to you, but given " & _
                 "the complexity of being able to enter the specs for " & _
                 "every possible convention and bidding system, we feel " & _
                 "that the complexity of the program is quite reasonable."
    L_Tutorial.Height = 2850
    int_GridNum = 0
    ColorGrids 0
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm"
    Wait
    frm_Tutorial.Visible = False

End Sub


Private Sub b_AddRef_Click()
    If t_Ref.Text <> "" Then
        list_Refs.AddItem t_Ref.Text
    End If
End Sub

Private Sub b_CancelFind_Click()
    b_FindBidForHand.Default = False
    t_FindHand(0) = ""
    t_FindHand(1) = ""
    t_FindPrior = ""
    cmb_FindVul = ""
    f_EnterHand.Visible = False
End Sub


Private Sub b_CancelFind_LostFocus()
    ' If tabbing to different objects on '
    ' the Sort/Show box, the following   '
    ' takes us from the Cancel button    '
    ' back to the hand entry box.        '
    If f_EnterHand.Visible Then t_FindHand(int_FindHandNum).SetFocus
End Sub


Private Sub b_ClearBidFound_Click()
    L_BidFound = ""
    L_FindBidID = ""
End Sub


Private Sub b_ClearFind_Click()
    t_FindHand(0).Text = ""
    t_FindHand(1).Text = ""
    t_FindPrior.Text = ""
    L_BidFound.Caption = ""
    t_FindAnalysis4Bid = ""
    cmb_FindPos = ""
    cmb_FindVul = ""
    cmb_Find_ScoreMeth = ""
    f_EnterHand.Visible = True
    t_FindHand(0).SetFocus
    b_EditFind(1).Caption = "Add Bid"
    b_EditFind(1).Width = 1100
    L_FindBidID = ""
    str_ImportSource = ""
    int_KnownFitLen = 0
    t_FindAnalysis4Bid.Left = 2420
    t_FindAnalysis4Bid.Width = 375
    L_BidFound.Left = 2310
    L_BidFound.Width = 360
    L_FindBid(3).Left = 495
    'ck_Unopposed.Value = 1 I'm always changing the value anyway. '
    f_Msg.Visible = False
    t_FindHand(0).SetFocus
End Sub

Private Sub b_ClearFindHands_Click()
    Dim response As String
    Dim mboxTop As Integer
    Dim mboxLeft As Integer

    With cmb_FindBid
    ''''''''''''''''
    If .ListCount = 0 Then
        response = MyMsgBox("No references to clear.", 1, "Ok", "Clear Hand List")
        .Visible = False
        Exit Sub
    End If

    .Visible = True
    
    mboxTop = f_Msg.Top
    mboxLeft = f_Msg.Left
    f_Msg.Top = f_EnterHand.Top - 800
    f_Msg.Left = f_EnterHand.Left + f_EnterHand.Width + 100
    response = MyMsgBox("Are you sure?", 1, "YN", "Clear Hand List")

    If response = "no" Then
        .Visible = False
    Else
        .Clear
    End If
    
    f_Msg.Move mboxLeft, mboxTop
    
    ''''''''
    End With
End Sub


Private Sub b_ClearInputs_Click()
    mClear_Click
End Sub


Private Sub b_ClearRefList_Click()
    Dim response As String

    With list_Refs
    ''''''''''''''
    If .ListCount = 0 Then
        response = MyMsgBox("No references to clear.", 1, "Ok", "Clear References")
        b_DelRef.Visible = False
        Exit Sub
    End If

    .Visible = True
    response = MyMsgBox("Are you sure?", 1, "YN", "Clear References")

    If response = "no" Then
        .Visible = False
    Else
        .Clear
        response = MyMsgBox("Enter the defaults?", 1, "YN", "References")
        If response <> "yes" Then
            .Visible = False
        Else
            .AddItem "Aces 2026-"
            .AddItem "Becker 2026-"
            .AddItem "IYC 2026-"
        End If
    End If
    ''''''''
    End With
    b_DelRef.Visible = False
End Sub

Private Sub b_Close_Hidden_Click()
    f_Unhide.Visible = False
End Sub


Private Sub b_Close_Show_Click()
    f_ShowSort.Visible = False
    b_Show.Default = False
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub b_Close_Show_DragDrop(source As Control, _
                                  x As Single, Y As Single)
    With source
        .Drag vbEndDrag
        .Move .Left + b_Close_Show.Left + x - sng_DragoffsetX, _
              .Top + b_Close_Show.Top + Y - sng_DragoffsetY
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub b_CloseColors_Click()
    mSetColors_Click
End Sub


Private Sub b_CompareChange_Click(Index As Integer)
    Dim gr1 As Integer
    Dim gr2 As Integer

    ' Note: The "master" grid (from whose data '
    ' the Input Boxes were filled) is in the   '
    ' box below the grid which is normally to  '
    ' be changed, so although this is a little '
    ' confusing, the lower box's Index=1 and   '
    ' the box above it is "2". Ditto for the   '
    ' L_DataFrom and L_CompareSource boxes.    '

    If InStr(L_DataFrom(2), "1") Then
        gr1 = 2 ' Grid 1 is in the top grid-data-box. '
        gr2 = 1
    Else
        gr1 = 1 ' Grid 1 is in the bottom-data-box.   '
        gr2 = 2
    End If

    If Index = 0 Then ' Copy Grid 1 to Grid 2 '
        t_GridData(gr2) = t_GridData(gr1)
        o_CompareOptions(1).Value = True
    Else ' Copy Grid 2 to Grid 1 '
        t_GridData(gr1) = t_GridData(gr2)
        o_CompareOptions(0).Value = True
    End If

End Sub

Private Sub b_CompareDataOK_Click()
    Dim changeGrid As Integer
    Dim compareTo As Integer
    Dim box As Integer
    Dim master As Integer

    ' Note: The "master" grid (from whose data '
    ' the Input Boxes were filled) is in the   '
    ' box below the grid which is normally to  '
    ' be changed, so although this is a little '
    ' confusing, the lower box's Index=1 and   '
    ' the box above it is "2". Ditto for the   '
    ' L_DataFrom and L_CompareSource boxes.    '

    If InStr(L_DataFrom(1), "1") > 0 Then
        master = 1
        compareTo = 2
    Else
        master = 2
        compareTo = 1
    End If

    If o_CompareOptions(0) Then
        changeGrid = 1

        If master = 1 Then
            box = 1
        Else
            box = 2
        End If

    ElseIf o_CompareOptions(1) Then
        changeGrid = 2

        If master = 2 Then
            box = 1
        Else
            box = 2
        End If
    End If

    If changeGrid > 0 Then
        SaveCompareChanges changeGrid, t_GridData(box)
        If InStr(b_ComparePutG2Data.Caption, "Grid " & changeGrid) > 0 Then
            ' Data from which the Input Boxes were  '
            ' filled has changed, so re-load it.    '
            ' Since the "Compute Input Boxes for"   '
            ' button is set to change the Input     '
            ' Boxes to data from what is currently  '
            ' NOT the master, we have to change the '
            ' label on the button before invoking   '
            ' the button's Click so that instead of '
            ' changing the Input Boxes to the other '
            ' grid, it will just reload the same    '
            ' grid's data.                          '
            b_ComparePutG2Data.Caption = "&Compute Input Boxes for Grid " & compareTo
            b_ComparePutG2Data_Click
        End If
    End If

    f_Compare.Visible = False
End Sub

Private Sub b_CompareHelp_Click()
    Dim xx As Long
    xx = Shell("BidBase File Viewer.exe " & _
               str_BBpath & "\Notes\CardShark_BidBase_Editor_Menus.htm#E11a")
    str_Response = MyMsgBox("Bringing up help" & vbCr & "in the File Viewer", 2, "ok", "")
End Sub

Private Sub b_ComparePutG2Data_Click()
    Dim r As Integer
    Dim gn As Integer

    If InStr(b_ComparePutG2Data.Caption, "Grid 1") > 0 Then
        ' "Grid 1" is Grid1(0) in the code '
        ' and "Grid 2" is Grid1(1).        '
        gn = 0
        int_GridNum = 1
    Else
        gn = 1
        int_GridNum = 0
        b_CompareUseInputBox.Caption = "Use Input &Box data in Grid 2"
        b_ComparePutG2Data.Caption = "&Compute Input Boxes for Grid 1"
    End If

    ColorGrids int_GridNum

    mEdit(2).Checked = False
    r = Grid1(gn).Row
    ' Put Grid1(x) into Input Boxes via "Use Entry" '
    ' then compare Grid1(y) to the Input Boxes.     '
    mEdit_Click 2
    Grid1(gn).Row = r
    f_Compare.Visible = False
End Sub

Private Sub b_CompareUseInputBox_Click()
    t_GridData(2) = t_GridData(0)
    If t_GridData(2) = "[blank]" Then t_GridData(2) = ""
    If InStr(b_CompareUseInputBox.Caption, "Grid 1") > 0 Then
        o_CompareOptions(0) = True
    Else
        o_CompareOptions(1) = True
    End If
    b_CompareDataOK_Click
End Sub

Private Sub b_Continue_Click()
    ' This button is used by the tutorial '
    ' to allow the user to control when   '
    ' to move to the next topic.          '

    If bln_Waiting Then
        bln_Waiting = False
        Exit Sub
    End If

    b_Continue.Visible = False
    L_Tutorial = "Click on another menu command or on Quit."
    frm_Tutorial.Caption = "CardShark BidBase Tutorial"
    L_Tutorial.Height = 2600
    b_QuitTutorial.Top = L_Tutorial.Height + 400
    b_Continue.Top = b_QuitTutorial.Top
    frm_Tutorial.Height = b_Continue.Top + 500
    mMenuTutorial.Checked = True
End Sub


Private Sub b_ConvNameChange_Click()
    Dim newData As String
    Dim d As String
    Dim i As Integer
    Dim fi As String
    Dim f2 As String
    Dim p As String
    Dim t As String
    Dim num As Long
    Dim sel As String
    Dim sel2 As String

   '                                                                 '
    With cmb_ConvName(int_cmbConvName_Index)
    ''''''''''''''''''''''''''''''''''''''''
        newData = .Text

        If opt_ChConvName(3).Value = True Then
            ' Go back to original name: '
            .Text = str_OrigData
            f_ConvNameChange.Visible = False
            Exit Sub

        ElseIf opt_ChConvName(2).Value = True Then
            ' Change the existing name:  '

            ' Find name in combo box: '
            i = FindMatch(cmb_ConvName(int_cmbConvName_Index), _
                          str_OrigData, True) ' True = exact match
            If i < 0 Then i = 0

            If .List(i) = str_OrigData Then
                .RemoveItem i
            End If
            .AddItem newData
            On Local Error GoTo 0

            ' Look for note file for original conv. name '
            ' and if there is one, change it to the new  '
            ' convention name:                           '
            Dim fi2 As String

            fi = str_OrigData
            i = InStr(fi, ":")
            If i > 0 Then
                fi = Left(fi, i - 1)
            End If

            fi2 = newData
            i = InStr(fi2, ":")
            If i > 0 Then
                fi2 = Left(fi2, i - 1)
            End If

            If fi <> fi2 Then
                fi = LegalFileName(fi)

                fi = fi & ".htm"

                p = str_BBpath & "\Notes\"

                If Dir$(p & fi) <> "" Then
                    ' Change original file name to new name: '
                    f2 = LegalFileName(fi2)
                    f2 = f2 & ".htm"
                    If Dir$(p & f2) = "" Then
                        Name p & fi As p & f2
                    Else
                        str_Response = MyMsgBox("There is a file for " & vbCr & _
                                                fi & vbCr & _
                                                "but it cannot be renamed to " & vbCr & _
                                                f2 & vbCr & _
                                                "because that file already exists.", 5, "ok", "")
                    End If
                End If
            End If

            LookUpBidName num, sel, sel2, str_OrigData
            If str_OrigData <> rs_BidNamesTbl!Bid_Name Then
                ' This should not be possible. '
                str_Response = MyMsgBox(str_OrigData & " not found in database.", 1, "ok", "")
            Else
                t = Trim$(.Text)
                If Len(t) > 50 Then t = Left$(t, 50)
                    rs_BidNamesTbl.Edit
                rs_BidNamesTbl!Bid_Name = t
                rs_BidNamesTbl.Update
                Data1(0).Refresh
                Data1(1).Refresh
            End If

            .Text = newData
            str_OrigData = newData
            .Refresh
            f_ConvNameChange.Visible = False

            Exit Sub

NamesUpdateError:
            str_Response = MyMsgBox(Err.Description, 1, "ok", "")
            On Local Error GoTo 0
            Exit Sub
        End If

        If .Text <> .List(.ListIndex) Then
            If .Index = 0 Then
                If Len(.Text) > 50 Then
                    str_Response = MyMsgBox("Name size limit is 50 characters." & vbCr & _
                                            "This file name was " & Len(.Text) & " characters.", 2, "ok", "")
                    .Text = Left$(.Text, 50)
                End If

                If opt_ChConvName(1).Value = True Then
                    ' Add as new line. '
                    t = .Text
                    '''''''''''''''''''
                    With rs_BidNamesTbl
                    '''''''''''''''''''
                        .AddNew
                        !Bid_Name = t

                        f_Msg.Left = f_ConvNameChange.Left - 150
                        f_Msg.Top = f_ConvNameChange.Top + f_ConvNameChange.Height + 50
                        Do
                            d = MyMsgBox("Enter description:", 1, "", t)
                            If Len(d) > 100 Then
                                str_Response = MyMsgBox("Limit for description is 100 characters." & vbCr & _
                                                        "This description is " & Len(!Text) & _
                                                        " characters.", 2, "ok", "")
                            End If
                            DoEvents
                        Loop While Len(d) > 100

                        '!Bid_Description = d
                        p = "Enter Convention Selection code." & vbCr & _
                            "   W = We use it" & vbCr & _
                            "   T = They use it" & vbCr & _
                            "   blank = both use it" & vbCr & _
                            "   0 or x = neither"
                        f_Msg.Left = f_ConvNameChange.Left - 290
                        d = MyMsgBox(p, 5, "", "")
                        !Selected = d
                        t_Sel = d
                        i = InStr(t, ":")
                        If i Then ' Subconvention name present: '
                            p = "Enter Subcategory " & Mid$(p, 17)
                            f_Msg.Left = f_ConvNameChange.Left - 645
                            d = MyMsgBox(p, 5, "", "")
                            !SubCat_Selected = d
                            t_Sub = d
                        End If
                        .Update

                        'Data2.Refresh
                        ' Update the combo boxes with new data:
                        cmb_ConvName(0).AddItem t
                        cmb_ConvName(0).Refresh
                        ' Set .ListIndex to new Name:
                        i = FindMatch(cmb_ConvName(0), t, True)
                        cmb_ConvName(0).ListIndex = i

                        ' Change the opponent's list too: '
                        cmb_ConvName(1).AddItem t
                        cmb_ConvName(1).Refresh
                    End With
                    ' Back to "With cmb_ConvName(Index)": '

                    If cmb_ConvName(0).Text = "" Then
                        b_Conven_Description.Picture = _
                            LoadPicture(str_BBpath & "\ButtonN.bmp")
                    Else
                        FindConvenNameFile cmb_ConvName(0).Text, True, False, False
                    End If
                    f_ConvNameChange.Visible = False
                End If
            Else
                str_Response = MyMsgBox("No such convention.", 1, "ok", "")
            End If
        End If
   '        '
    End With
   '                                                                      '
    f_ConvNameChange.Visible = False
End Sub


Private Sub b_CopyFindHand_Click(Index As Integer)
    Dim a As Long
    
    If t_FindHand(int_FindHandNum) = "" Then Exit Sub
    
    If Index = 1 Then
        t_TestHand = t_FindHand(int_FindHandNum)
        If t_Find_Ref <> "" Then t_Ref = t_Find_Ref
    ElseIf t_Find_Ref <> "" And t_Disclosure = "" Then
        t_Disclosure = "[Also " & t_Find_Ref & ":  " & t_FindHand(int_FindHandNum)
    Else
        t_Disclosure = t_Disclosure & " " & t_FindHand(int_FindHandNum)
    End If
End Sub


Private Sub b_DelFindHand_Click(Index As Integer)
    Dim i As Integer
    Dim iStr As String

    With cmb_FindBid
    ''''''''''''''''
    If t_FindHand(Index).Text = "" And .ListCount > 0 Then
        f_Msg.Move f_EnterHand.Left + f_EnterHand.Width + 20, _
                   f_EnterHand.Top - 450
        str_Response = MyMsgBox("Clear the list?", 1, "YN", "")
        If str_Response = "yes" Then
            .Clear
            Exit Sub
        End If
    End If

    If .ListCount = 1 Then Exit Sub

    Do While .List(i) <> t_FindHand(Index).Text And _
             i < .ListCount - 1
        i = i + 1
    Loop
    If .List(i) <> t_FindHand(int_FindHandNum).Text Then
        str_Response = MyMsgBox(t_FindHand(int_FindHandNum).Text & _
                       " is not in the list.", _
                       1, "ok", "")
        Exit Sub
    End If

    .ListIndex = i
    bln_Abort = False
    str_Response = MyMsgBox("Delete " & t_FindHand(int_FindHandNum).Text & _
                            " from the list?", _
                            1, "YN", "BidBase")
    If str_Response <> "yes" Then
        .Visible = False
        t_FindAnalysis4Bid.Visible = True
        b_DelFindHand(int_FindHandNum).Visible = False
        Exit Sub
    End If

    t_FindHand(int_FindHandNum).Text = ""
    i = .ListIndex ' RemoveItem make .ListIndex = -1 '
    .RemoveItem .ListIndex
    bln_SkipIt = True
    If i = .ListCount Then
        i = .ListCount - 1
    End If
    .ListIndex = i
    bln_SkipIt = False

    If .ListCount < 2 Then
        .Visible = False
        b_DelFindHand(int_FindHandNum).Visible = False
        Exit Sub
    End If
    ''''''''
    End With
    'b_DelFindHand.Visible = False
End Sub


Function MyMsgBox(txt As String, numLines As Integer, response As String, boxCaption As String) As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim msgWidth As Integer
    Dim x As String
    Dim n As Integer
    Dim useOk As Integer

    Do While f_Msg.Visible
        DoEvents
    Loop
    
    If bln_Abort Then
        bln_Abort = False
        Exit Function
    End If

    If boxCaption = "" Then boxCaption = "BidBase"
    L_Caption = "  " & boxCaption
    If numLines = 0 Then numLines = 1

    L_MsgBox = txt
    ' Find longest line: '
    j = InStr(txt, Chr$(13))
    If j > 0 Then
        msgWidth = j
    Else
        msgWidth = Len(txt)
    End If
    For n = 2 To numLines
        i = InStr(j + 1, txt, Chr$(13))
        If i = 0 Then i = Len(txt)
        If (i - j) > msgWidth Then msgWidth = i - j
        j = i
    Next
    If msgWidth = 0 Then
        msgWidth = Len(txt)
    End If

    If msgWidth < 30 Then
        L_MsgBox.Width = 4300
    ElseIf msgWidth < 35 Then
        L_MsgBox.Width = 4800
    ElseIf msgWidth < 45 Then
        L_MsgBox.Width = 5500
    ElseIf msgWidth < 48 Then
        L_MsgBox.Width = 6000
    Else
        L_MsgBox.Width = 6500
    End If

    t_MsgInput.Width = L_MsgBox.Width - 1000
    f_Msg.Width = L_MsgBox.Width + 100
    L_Caption.Width = f_Msg.Width

    f_Msg.Visible = True
    L_MsgBox.Height = 300 * numLines + 50
    f_Msg.Height = 1300 + numLines * 300
    L_MsgBackground.Height = f_Msg.Height - 430
    L_MsgBackground.Width = f_Msg.Width - 135
    t_MsgInput.Visible = False

    response = UCase(response)

    b_MsgInput(1).Visible = False  ' No '
    b_MsgInput(2).Visible = False  ' Cancel '
    b_MsgInput(1).Width = 650
    
    b_MsgInput(0).Top = numLines * 300 + 650
    'b_MsgInput(0).Top = L_MsgBox.Top + L_MsgBox.Height + 50
    b_MsgInput(0).Default = True

    If response = "OK" Or response = "DOCS" Then ' Wait for Ok press. '
        b_MsgInput(0).Caption = "OK"
        b_MsgInput(0).SetFocus
        If response = "DOCS" Then
            b_MsgInput(1).Caption = "Show Docs"
            b_MsgInput(1).Top = b_MsgInput(0).Top
            b_MsgInput(1).Width = 2000
            b_MsgInput(1).Visible = True
        End If

    ElseIf Left$(response & "   ", 2) = "YN" Then ' Show message, wait for Yes, No. '
        b_MsgInput(0).Caption = "Yes"
        b_MsgInput(1).Caption = "No"
        b_MsgInput(1).Top = b_MsgInput(0).Top
        b_MsgInput(1).Visible = True
        If Right$(response, 1) = "C" Then
            b_MsgInput(2).Top = b_MsgInput(0).Top
            b_MsgInput(2).Visible = True
        End If
        b_MsgInput(0).SetFocus

    Else ' Input box uses response as default input. '
        t_MsgInput.Visible = True
        b_MsgInput(0).Caption = "Ok"
        t_MsgInput.Top = b_MsgInput(0).Top
        t_MsgInput.Text = response
        b_MsgInput(0).Default = True
        t_MsgInput.SetFocus
        t_MsgInput.SelStart = Len(response)

    End If

    f_Msg.ZOrder 0
    Do While f_Msg.Visible
        x = DoEvents()
    Loop
    
    If response = "OK" Or response = "DOCS" Then
        ' No value is returned unless Docs. '
        If int_MsgClick = 1 Then
            MyMsgBox = "docs"
        End If
    ElseIf Left$(response, 1) = "Y" Then
        ' 0=Yes, 1=No, 2=Cancel '
        If int_MsgClick = 2 Then
            MyMsgBox = "cancel"
        ElseIf int_MsgClick = 1 Then
            MyMsgBox = "no"
        Else
            MyMsgBox = "yes"
        End If
        Exit Function
    Else ' Input box '
        MyMsgBox = t_MsgInput.Text
    End If
    b_MsgInput(0).Default = False
End Function


Private Sub b_DelRef_Click()
    Dim i As Integer
    Dim iStr As String
    Dim mboxTop As Integer
    Dim mboxLeft As Integer

    If list_Refs.ListCount < 1 Then Exit Sub

    mboxTop = f_Msg.Top
    mboxLeft = f_Msg.Left
    f_Msg.Top = list_Refs.Top + frame0(6).Top - 1000
    f_Msg.Left = list_Refs.Left + frame0(6).Left + list_Refs.Width + 200
    f_Msg.ZOrder 0

    With list_Refs
    ''''''''''''''
    If t_Ref = "" And .ListCount > 1 Then
        'str_Response = MyMsgBox("Clear the list?", 1, "YN", "Bid References")
        'If str_Response = "yes" Then
            b_ClearRefList_Click
            GoTo ExitDelRef
        'End If
    End If

    If .ListCount < 0 Then Exit Sub

    Do While .List(i) <> t_Ref And i < .ListCount - 1
        i = i + 1
    Loop
    If .List(i) <> t_Ref Then
        str_Response = MyMsgBox(t_Ref & " is not in the list.", 1, _
                                "ok", "Bid Reference")
        GoTo ExitDelRef
    End If

    .ListIndex = i

    str_Response = MyMsgBox("Delete " & .List(.ListIndex) & " from the list?", 1, "YN", "Delete Reference")
    If str_Response <> "yes" Then
        .Visible = False
        b_DelRef.Visible = False
        GoTo ExitDelRef
    End If

    t_Ref = ""
    i = .ListIndex ' RemoveItem make .ListIndex = -1 '
    .RemoveItem .ListIndex
    bln_SkipIt = True
    If i = .ListCount Then
        i = .ListCount - 1
    End If
    .ListIndex = i
    bln_SkipIt = False

    If .ListCount < 2 Then
        .Visible = False
        b_DelRef.Visible = False
    End If
    ''''''''
    End With

ExitDelRef:
    f_Msg.Top = mboxTop
    f_Msg.Left = mboxLeft
    b_DelRef.Visible = False
End Sub

'**************************************************
' Procedure: b_DupeEntry_Click                    '
' Purpose:                                        '
'     The purpose of this routine is to duplicate '
'     similar entries into different Prior Bid    '
'     sections. For example, specs for many bids  '
'     after 1C-1H should be similar to those for  '
'     Prior Bids of 1C-1S; likewise, 1D-1H, 1D-1S '
'     etc.  By duplicating similar entries, not   '
'     only do we save the user time, but we help  '
'     ensure consistency among similar entries.   '
'**************************************************

Private Sub b_DupeEntry_Click(Index As Integer)

    Dim pb As String
    Dim gn As Integer
    Dim j As String
    Dim xv As Double
    Dim m As String
    Dim xx As Integer

    gn = 1

    If Grid1(0).Columns(int_Prior_Bids_Col) <> t_Prior Then
        str_Response = MyMsgBox("Select an entry in the top grid first.", 1, "ok", "")
        Exit Sub
    End If

    If t_DupeEntry(Index) = t_Prior Then
        str_Response = MyMsgBox("The 'Dupe entry' prior bids (" & t_DupeEntry(Index) & ") are" & vbCr & _
                                "the same as in the Prior Bids input box.", 2, "ok", "")
        Exit Sub
    End If

    pb = t_DupeEntry(Index) ' Prior Bids for new entry '

    int_GridNum = 0

    lng_ID = Grid1(0).Columns(int_LastBidCol)
    lng_DupeID = lng_ID
    int_DupeGrid = int_GridNum

    'mClear_Click ' Clear the input boxes. '

    If Grid1(gn).Columns(int_Prior_Bids_Col) <> pb And _
       mDupesGrid(0).Checked = False _
    Then
        t_Show(gn * 2) = pb
        t_Show(gn * 2 + 1) = pb
        j = str_Job
        str_Job = "Load 1 grid"
        b_Show_Click
        str_Job = j
    End If

    Grid1(0).Col = int_OrdNum_Col
    Grid1(1).Col = int_OrdNum_Col

    Do While Grid1(gn) > Grid1(int_GridNum) Or _
             Grid1(gn).Columns(int_Pct_Col) <> _
             Grid1(int_GridNum).Columns(int_Pct_Col)

        If Grid1(gn).Row <= 0 Then 'Normally can't be <0
            xv = Grid1(gn).Columns(int_LastBidCol)
            Grid1(gn).Scroll 0, -1
            Grid1(gn).Row = 0
            If xv = Grid1(gn).Columns(int_LastBidCol) Then
                Exit Do
            End If
        Else
            Grid1(gn).Row = Grid1(gn).Row - 1
        End If
        DoEvents
    Loop

    Do While (Grid1(gn) < Grid1(int_GridNum) Or _
              Grid1(gn).Columns(int_Pct_Col) <> _
              Grid1(int_GridNum).Columns(int_Pct_Col) _
             ) And _
             Grid1(gn) <> ""
        Grid1(gn).Row = Grid1(gn).Row + 1
        DoEvents
    Loop

    If Grid1(gn) = t_Order And _
       (mDupesGrid(0).Checked = False Or _
        Grid1(1).Columns(int_Prior_Bids_Col) = t_DupeEntry(Index)) _
    Then
        m = "There is already an entry with the same Bid Order Number." & vbCr & _
            "Do you want to compare it to the current entry?"
        str_Response = MyMsgBox(m, 2, "YN", "")
        If str_Response = "yes" Then
            If t_Order < Grid1(int_GridNum).Columns(int_OrdNum_Col) Then
                Grid1(int_GridNum).Row = Grid1(int_GridNum).Row - 1
            End If
            str_Job = "Compare Dupes"
            mCompare_Click
            str_Job = ""
            GoTo ExitDupeEntry
        End If
    End If

    ' Put entry into Input Boxes via 'Use Entry':
    Grid1(0).Col = int_Prior_Bids_Col
    If Grid1(0).Columns(int_OrdNum_Col) = "ZZZZZZ" And _
       Grid1(0).Row > 0 _
    Then
        Grid1(0).Row = Grid1(0).Row - 1
    End If

    If Grid1(0).Columns(int_LastBidCol) <> lng_ID Then
        ' Locate the original entry's line:
        Do While Grid1(0).Columns(int_OrdNum_Col) > t_Order And _
                 Grid1(0) = t_Prior And _
                 Grid1(0).Columns(int_LastBidCol) <> lng_ID

            Grid1(0).Scroll 0, -1
            Grid1(0).Row = Grid1(0).Row - 1
            DoEvents
        Loop

        Do While Grid1(0).Columns(int_OrdNum_Col) < t_Order And _
                 Grid1(0).Columns(int_LastBidCol) <> lng_ID And _
                 Grid1(0) = t_Prior

            Grid1(0).Row = Grid1(0).Row + 1
            DoEvents
        Loop

        If Grid1(0).Columns(int_LastBidCol) <> lng_ID Or _
           Grid1(0) <> t_Prior _
        Then
            j = MyMsgBox("Original entry not found in grid.", 1, "ok", "")
            Exit Sub
        End If
    End If
    j = str_Job
    str_Job = "Duping"
    str_DupePrior = t_DupeEntry(Index)

    ' Put grid entry into Input Boxes: '
    mEdit_Click 2

    str_DupePrior = ""
    str_Job = j
    b_DupeEntry(Index).Visible = False
    j = MyMsgBox("Entry has been loaded into the Input Boxes." & vbCr & _
                 "Check/edit it and then press Ctrl-S to save.", _
                 2, "ok", "")

ExitDupeEntry:

End Sub


Private Sub Run_DupesGrid_Tutorial()
    frm_Tutorial.Caption = "Dupes Grid"
    L_Tutorial = "The Dupes Grid lets you compare similar entries " & _
                 "(for the same round, player and same Bid Ord.#'s) " & _
                 "to ensure that the specs for them are comparable. " & _
                 "Click View - Dupes Grid - Show Dupes Grid to display " & _
                 "the Dupes Grid." & _
                 vbCr & vbCr & _
                 "The second option in that menu is to show only those " & _
                 "entries whose 'Prior Bids' are defined in the database " & _
                 "as being very similar. (Try it now.)" & _
                 vbCr & vbCr & _
                 "The third option is to show only those entries which " & _
                 "have the same bid/convention name. (Try it now.)" & _
                 vbCr & vbCr & _
                 "The Dupes grid is extremely important. It allows duplicating " & _
                 "many entries from a single new entry, among other things. " & _
                 "Be sure to click Details for a full explanation."
    L_Tutorial.Height = 4200
    frm_Tutorial.Height = 5100
    ResizeTutorial
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#V4C"
    frm_Tutorial.Left = 7000
    bln_Waiting = False
End Sub


Private Sub b_EditFind_Click(Index As Integer)
    ' Index 1 = Start a new entry. 0 = Edit existing entry. '
    Dim bidNumber As Integer ' excluding opening passes. '
    Dim thisBidNum As Integer ' bidNumber is the number of Prior Bids.'
                              ' thisBidNum is the bid for this entry. '
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim ref As String
    Dim card As String
    Dim cardNum As Integer
    Dim cardPtr As Integer
    Dim numDashes As Integer
    Dim intermed(4) As Integer
    Dim lastCard As Integer
    Dim str_Hands(4) As String
    Dim suit As Integer
    Dim suitHCP(4) As Integer
    Dim suitLen(4) As Integer
    Dim cvtSuitNum(4) As Integer ' This Sub numbers suits 1-4 for S-C. cvtSuitNum renumbers them 4-1 for S-C. '
    Dim suitHCPltc(4) As Integer
    Dim suitLenLTC(4) As Integer
    Dim suitDistPts(4) As Integer
    Dim strSuit(4) As String
    Dim TestHand As String
    Dim p As Integer ' counter for # of players '
    Dim pc As String ' prior calls including opening passes. '
    Dim r As Integer ' round # excluding opening passes. '
    Dim firstBid As Integer ' First bid after opening passes. '
    Dim pb As String
    Dim ok As Boolean
    Dim trumpSuit(5) As String ' Potential trump suit bid by partner: 1S-P-me=S, 1S-P-1N-P,2C=trumpSuit(1)=S, trumpSuit(2)=C.'

    ' Edit the bid found if any; otherwise, '
    ' create a new entry in the Input Boxes '
    ' using the specs and bidding for the   '
    ' hand in the Find Bid window. '
    b_ClearInputs_Click

    pb = t_FindPrior
    i = InStr(pb, "@")
    If i > 2 Then
        pb = Mid$(pb, i - 2)
    End If
    str_PriorCalls = pb
    
    If pb <> "" Then
        If Right$(pb, 1) = "-" Then
            pb = Left$(pb, Len(pb) - 1)
        End If
        ValidateBidding pb, p, r, ok
        p = p + 1 ' The above calculates last bidder. '
        If p = 5 Then
            p = 1
            r = r + 1
        End If
        If Not ok Then
            Exit Sub
        End If
        t_Player = "" & p
        t_Round = "" & r
    Else
        t_Player = "1"
        t_Round = "1"
    End If
    thisBidNum = (Val(t_Round) - 1) * 4 + p

    ' Load entry into Input Boxes: '
    If Index = 0 Then ' Edit '
        If Grid1(0).Row < 0 Then
            ' Bid was found but not in the section of Prior Bids shown. '
            ' It may be in a "1M" section, so try that: '
            i = InStr(pb, "H")
            j = InStr(pb, "S")
            If i > 0 And (i < j Or j = 0) Then
                pb = Replace$(pb, "H", "M")
            ElseIf j > 0 Then
                pb = Replace$(pb, "S", "M")
            End If
            t_Show(0) = pb
            t_Show(1) = pb
            cmb_Show(0).ListIndex = int_Prior_Bids_Col
            b_Show_Click
        End If
        If Grid1(0).Row >= 0 Then
            bln_SkipIt = True
            mFindEntryID_Click ' Finds the entry in the grid and copies it to boxes. '
            If t_Ref = "Deal #0" Then
                t_Ref = t_Find_Ref
            End If
            bln_SkipIt = False
        End If
        ck_SaveNew.Value = 0
        ck_Saved.Value = 1
        Exit Sub
    End If

    ' New entry: '
    L_BidFound = ""

    t_TestHand = t_FindHand(int_FindHandNum).Text
    TestHand = t_TestHand ' TestHand is a sub, so this makes no sense '
    If Len(str_PriorCalls) > 1 Then
        If Right$(str_PriorCalls, 1) = "-" Then
            str_PriorCalls = Left$(str_PriorCalls, Len(str_PriorCalls) - 1)
        End If
    End If
    t_Prior = str_PriorCalls

    cvtSuitNum(1) = 4
    cvtSuitNum(2) = 3
    cvtSuitNum(3) = 2
    cvtSuitNum(4) = 1

    cmb_Position.ListIndex = cmb_FindPos.ListIndex
    
    If str_PriorCalls = "" Then
        t_Round = "1"
        t_Player = "1"
    Else
        ' Count opening passes for number of '
        ' Prior Calls: '
        If Left$(str_PriorCalls & "     ", 5) = "P-P-P" Then
            If str_PriorCalls = "P-P-P" Then
                ' Current bidder is in first non-passing position: '
                t_Player = "1"
                t_Round = "1"
                str_PriorCalls = ""
                cmb_Position = "4th seat only"
            Else ' Could be -- '
                 ' P-P-P-1C ?? (2nd round, passed hand) '
                 ' P-P-P-1C-1D-1H-1S ?? (2nd, unpassed) '
                i = 7 ' Position of the call after 3rd pass: P-P-P-1H... '
                'Do While Mid$(str_PriorCalls, i, 1) = "-"
                '    i = InStr(i + 2, str_PriorCalls, "-")
                '    bidNumber = bidNumber + 1  ' excludes opening passes '
                '    If i = 0 Then Exit Do
                '    DoEvents
                'Loop
                str_PriorCalls = Mid$(str_PriorCalls, i) ' Remove starting passes. '
            End If

        ElseIf Left$(str_PriorCalls & "   ", 3) = "P-P" Then
            If str_PriorCalls = "P-P" Then
                t_Player = "1"
                t_Round = "1"
                str_PriorCalls = ""
                cmb_Position = "3rd seat only"
            Else ' Could be -- '
                 ' P-P-1C-?? (1st round, unpassed hand) or  '
                 ' P-P-1C-1D-?? (2nd round, passed hand) or '
                 ' P-P-1C-1D-1H-1S-?? (2nd round, unpassed) '
                i = 5 ' Position of the call after 2nd pass: P-P-1D... '
                'bidNumber = 0
                'Do While i > 0
                '    bidNumber = bidNumber + 1
                '    i = InStr(i + 2, str_PriorCalls, "-")
                '    DoEvents
                'Loop
                str_PriorCalls = Mid$(str_PriorCalls, i) ' Remove starting passes. '
            End If

        ElseIf Left$(str_PriorCalls & "  ", 1) = "P" Then
            If str_PriorCalls = "P" Then
                t_Player = "1"
                t_Round = "1"
                str_PriorCalls = ""
                cmb_Position = "2nd seat only"
            Else
                i = 3
                str_PriorCalls = Mid$(str_PriorCalls, i)
            End If

        Else ' PriorCalls = all passes and = "" have been '
             ' eliminated, so str_PriorCalls has to start '
             ' with an actual bid of some kind, like 1C   '
             ' or 1C-P, etc. '
            If Mid$(str_PriorCalls, 3, 1) = "-" Then
                i = 1
                'bidNumber = 0
                'Do While Mid$(str_PriorCalls, i, 1) = "-"
                '    i = InStr(i + 2, str_PriorCalls, "-")
                '    bidNumber = bidNumber + 1
                '    If i = 0 Then Exit Do
                '    DoEvents
                'Loop
            ElseIf str_PriorCalls <> "" Then
                ' No "-" means only 1 bid was made. '
                i = 0
                bidNumber = 1
            End If
        End If

        If Left$(str_PriorCalls & " ", 1) = "P" Then
            Do While Left$(str_PriorCalls & " ", 1) = "P"
                str_PriorCalls = Mid$(str_PriorCalls & " ", 3)
            Loop
            str_PriorCalls = RTrim$(str_PriorCalls)

            str_Response = MyMsgBox("Leading Passes were removed from Prior Bids." & vbCr & _
                                    "Specify Position if needed.", 2, "ok", "")
        End If

        t_Prior = str_PriorCalls
        
        t_Ref = t_Find_Ref

        ' Find raises: '
        If bidNumber > 2 Then

        End If
    End If

    suit = 1  ' Hand starts with suit 1, Spades. '
    ' Count points in the hand: '
    For cardPtr = 1 To 16 ' Each card (and dash) in the hand. '
        ' See if character is a dash: '
        Do While Mid$(TestHand, cardPtr, 1) = "-"
            ' Suit is finished when "-" reached.          '
            ' If void in Spades, hand may start with "-". '
            ' If void in H/D, two dashes are together,    '
            ' such as AK432--K853-QJT2, so Do-Loop is used.
            DoEvents
            If cardPtr = 16 Then
                'If hand ends with "-", it is void in Clubs
                Exit For
            Else ' End of suit reached. '
                'Set variables for next suit:
                suit = suit + 1
                If suit = 5 Then
                    str_Response = MyMsgBox(t_FindHand(int_FindHandNum).Text & vbCr & _
                                            "has too many dashes.", _
                                            2, "ok", "")
                    Exit Sub
                End If
                lastCard = 99
                cardPtr = cardPtr + 1 'Get card after dash
            End If
            DoEvents
        Loop

        card = Mid$(TestHand, cardPtr, 1)
        cardNum = InStr(" 23456789TJQKA", card)

        ' Check if card is valid: '
        If cardNum < 2 Then ' Should not be possible with entry error checking. '
            str_Response = MyMsgBox(card & " is not a valid card.", 1, "ok", "")
            Exit Sub

        ElseIf cardNum >= lastCard And lastCard > 0 Then ' Should not be possible. '
            str_Response = MyMsgBox("The " & card & " of " & str_SuitName(suit) & " is out of sequence.", 1, "ok", "")
            Exit Sub
        End If

        'Good card, so add to HCP if an honor,
        'add to Intermediates if 8-10,
        'and add to Suit array:
        suitHCP(suit) = suitHCP(suit) + InStr("JQKA", card) ' J=1, Q=2, etc. '
        suitLen(suit) = suitLen(suit) + 1
        strSuit(cvtSuitNum(suit)) = strSuit(cvtSuitNum(suit)) & card

        ' Count intermediates: '
        ' Rarely use this in an entry. '
        'If InStr("9T", card) Then
        '    intermed(suit) = intermed(suit) + 1
        '    If card = "8" Then ' Count an 8 when with 9 or T. '
        '        intermed(suit) = intermed(suit) + 1
        '    End If
        'End If

        'Update LastCard number: '
        lastCard = cardNum
    Next

    'Total up and display Qtys, HCPs and Intermediates '
    For suit = 1 To 4
        suitLen(0) = suitLen(0) + suitLen(suit)
        If suitLen(suit) = 0 Then
            suitDistPts(suit) = 3
        ElseIf suitLen(suit) = 1 Then
            If suitHCP(suit) = 4 Or suitHCP(suit) < 1 Then
                ' Only count distribution for A or 2-10. '
                suitDistPts(suit) = 2
            End If
        ElseIf suitLen(suit) = 2 Then
            If suitHCP(suit) > 2 Then
                ' Don't count distribution for Qx, Jx. '
                If InStr("KQJ", Right$(suitHCP(suit), 1)) = 0 Then
                    ' Don't count distribution for AK, AQ, AJ, KQ, KJ. '
                    suitDistPts(suit) = 1
                End If
            Else
                suitDistPts(suit) = 1
            End If
        End If
        suitDistPts(0) = suitDistPts(0) + suitDistPts(suit)
        suitHCP(0) = suitHCP(0) + suitHCP(suit)
        intermed(0) = intermed(0) + intermed(suit)
        If suitLen(suit) > 0 Then t_Qty(suit) = suitLen(suit) & ""
        If suitHCP(suit) > 0 Then t_HCP(suit) = suitHCP(suit)
        If intermed(suit) > 0 Then
            If suitHCP(suit) = 0 Then
                t_HCP(suit) = "i" & intermed(suit)
            Else
                t_HCP(suit) = t_HCP(suit) & "i" & intermed(suit)
            End If
        End If
    Next

    If str_ImportSource <> "" Then
        t_Ref = ref & str_ImportSource
    End If

    'StoreBiddersHandSpecs

    t_Hand_HCPs = "" & suitHCP(0)

    ' Check for dummy points: '
    i = 0 ' number of suits bid '
    If bidNumber = 3 Then

    End If

    If suitDistPts(0) > 0 Then
        t_Distrib_Pts = "" & suitDistPts(0)
    End If

    ck_SaveNew.Value = 1
End Sub


Private Sub b_EnterHandHelp_Click(Index As Integer)
    Dim xx As Long

    If Index = 0 Then
        str_NoteFile(0) = "CardShark_BidBase_Editor_Menus.htm#F2"
    Else
        str_NoteFile(0) = "CardShark_BidBase_Editor_Input_Boxes.htm#AO1"
    End If
    str_PgmPath(0) = str_BBpath & "\Notes\"
    xx = Shell("BidBase File Viewer.exe " & _
               str_PgmPath(0) & str_NoteFile(0), vbNormalFocus)
End Sub


Private Sub b_Go_Back_Click()
    ' Used in 'Edit' Tutorial. '
    bln_GoBack = True
    bln_Waiting = False
End Sub


Private Sub b_Google_Click()
    Dim s As String
    Dim i As Integer
    Dim x As Long

    s = Trim(cmb_ConvName(0))
    i = InStr(s, ":")
    If i > 1 Then
        s = Left$(s, i - 1)
    End If
    Replace s, " ", "+", 1, 99

    If s = "" Then
        s = MyMsgBox("Enter text to search for:", 1, "", "Google Web Search")
        If s = "" Then Exit Sub
    End If

    x = Shell("BidBase File Viewer.exe " & _
              "http://www.google.com/search?num=50&hl=en&lr=&q=bridge+convention+" & _
              "%22" & s & "%22", vbNormalFocus)

End Sub


Private Sub b_Import_Click()
    mImportPracticeDeal_Click
End Sub


Private Sub b_MsgInput_Click(Index As Integer)
    int_MsgClick = Index
    f_Msg.Visible = False
    bln_Abort = True
End Sub


Private Sub b_QuitTutorial_Click()
    frm_Tutorial.Visible = False
    L_Tutorial.Height = 2600
    b_QuitTutorial.Top = L_Tutorial.Height + 400
    b_Continue.Top = b_QuitTutorial.Top
    frm_Tutorial.Height = b_Continue.Top + 500
    mMenuTutorial.Checked = False
    mFullTutorial.Checked = False
    bln_Waiting = False
End Sub


Private Sub b_Conven_Description_Click()
    ' This button is located over the '
    ' Convention Name input box. When '
    ' clicked, it causes the Conven   '
    ' note file, if any, to display.  '
    Dim f As String

    If Trim$(cmb_ConvName(0).Text) = "" Then
        str_Response = MyMsgBox("Select a convention first.", 1, "ok", "")
        Exit Sub
    End If

    f = Replace$(cmb_ConvName(0).Text, "2/1", "2-Over-1")
    If InStr(f, "Stayman") > 0 Or InStr(f, "Jacoby Transfer") > 0 Or f = "Texas Transfer" Then
        f = "Notrump_Bidding"
    End If

    FindConvenNameFile f, True, True, True
End Sub


Private Sub b_Discl_Click()
    ' This button is located to the right  '
    ' of the Disclosure note field. If it  '
    ' is clicked, it causes the Disclosure '
    ' note file, if any, to be displayed.  '
    ' If no file exists, one is created.   '
    Dim useEntry As Boolean

    useEntry = InStr(t_Disclosure, ".htm") > 0
    LookForNoteFile t_Disclosure, _
                    t_Prior, _
                    cmb_Bid, _
                    "" & lng_ID, _
                    useEntry, _
                    True, _
                    True

    If t_Disclosure = "" Then
        t_Disclosure = "[See notes -->"
    End If
End Sub


Private Sub b_DupeColors_Click()
    ' Duplicate the colors selected for the '
    ' top grid to the other grids & boxes.  '
    Dim i As Integer

    For i = 1 To 3
        cmb_ForeColors(i).ListIndex = cmb_ForeColors(0).ListIndex
        cmb_BackColors(i).ListIndex = cmb_BackColors(0).ListIndex
    Next
End Sub


Private Sub b_FindBidForHand_Click()
    ' User enters prior bids in the t_FindPrior '
    ' box, enters a hand in the t_FindHand box, '
    ' then clicks this button to get a bid from '
    ' the bidding module for the hand.          '

    Dim b As String
    Dim pr As String
    Dim prNTO As String ' Save full prior bid while searching for NT overcall bid. '
    Dim prNTstart As String
    Dim e As String
    Dim ec As Integer
    Dim alreadyInGrid As Boolean
    Dim topID As Long
    Dim i As Integer
    Dim r As Long
    Dim ntOvercall As Boolean

    b_FindBidForHand.Enabled = False ' Prevent multiple clicks. '
    
    L_BidFound.Caption = ""
    
    If t_FindHand(int_FindHandNum).Text = "" Then
        int_FindHandNum = 0
    End If

    If Len(t_FindHand(int_FindHandNum).Text) < 16 Then
        str_Response = MyMsgBox("Hand does not have" & vbCr & _
                                "13 cards in it.", 2, "ok", "")
        t_FindHand(int_FindHandNum).SetFocus
        b_FindBidForHand.Enabled = True
        Exit Sub
    End If

    If Right$(t_FindPrior, 1) = "-" Then
        bln_SkipIt = True
        t_FindPrior = Left$(t_FindPrior, Len(t_FindPrior) - 1)
    End If

    bln_SkipIt = True
    t_FindPrior = Replace(t_FindPrior, "--", "-")
    pr = Trim$(t_FindPrior)
    ' This stops being able to have leading passes. '
    ' I guess that the Position list is supposed to '
    ' take care of that, but it doesn't seem to. '
    'Do While Left$(pr & "  ", 2) = "P-"
    '    pr = Trim$(Mid$(pr & " ", 3))
    '    DoEvents
    'Loop
    'If pr = "P" Then pr = ""
    
    Dim j As Integer

    i = InStrRev(pr, "@")
    j = InStrRev(pr, "-")
    If i > 3 And j - i = 3 Then
        ' Example: reduce 1H-P-4H-P-4N@30-P to 4N@30. '
        pr = Mid$(pr, i - 2, 5)
        t_Show(0).Text = pr
        b_Show_Click
    End If

    If Grid1(0).Row = -1 Then
        ' Grid is empty. Put something in it. '
        mViewAll_Click
    End If

    If Grid1(0).Columns(int_Prior_Bids_Col).Text = pr Then
        alreadyInGrid = True
        int_GridNum = 0
        ColorGrids 0

    ElseIf bln_TwoGrids And Not _
           mDupesGrid(0).Checked _
    Then
        On Local Error Resume Next
        If Grid1(1).Columns(int_Prior_Bids_Col) = pr Then
            alreadyInGrid = True
            int_GridNum = 1
            ColorGrids 1
        End If
        On Local Error GoTo 0
    End If
       
    If Not alreadyInGrid And Mid$(pr, 3, 1) <> "@" Then
        If Grid1(0).Row = -1 Then
            t_Prior = t_FindPrior
            t_TestHand = t_FindHand(int_FindHandNum).Text
             str_Response = MyMsgBox("No bid found.", 1, "ok", "")
            b_FindBidForHand.Enabled = True
            Exit Sub
        End If
        t_Show(0) = pr
        t_Show(1) = pr
        cmb_Show(0).ListIndex = int_Prior_Bids_Col
        b_Show_Click
        If Grid1(0).Row = -1 Then
            ' Grid shows no bids for Prior Bids entered. '
            If InStr(pr, "H") > 0 Or InStr(pr, "S") > 0 Then
                GoTo GetBidFromDLL ' May be a 1M bid. '
            End If
            b_EditFind(1).Caption = "Add Bid"
            b_EditFind(1).Width = 1100
            L_FindBidID = ""
            lng_ID = 0
            b_FindBidForHand.Enabled = True
            Exit Sub
        End If

        Grid1(0).Row = 0
        Grid1(0).Col = int_Prior_Bids_Col
        If Grid1(0).Row = -1 Then
            t_Prior = t_FindPrior
            t_TestHand = t_FindHand(int_FindHandNum).Text
            str_Response = MyMsgBox("No bid found.", 1, "ok", "")
            b_FindBidForHand.Enabled = True
            Exit Sub
        End If
        If Grid1(0).Text <> pr Then
            L_BidFound = ""
            If t_FindPrior <> t_Prior Or _
               t_FindHand(int_FindHandNum).Text <> t_TestHand.Text _
            Then
                mClear_Click
            End If
            str_Response = MyMsgBox("No entries found with Prior Bids of" & vbCr & _
                           t_FindPrior, 2, "ok", "Find Bid")
            mPreviewHandSpecs_Click
            ck_SaveNew.Value = 1
            b_FindBidForHand.Enabled = True
            Exit Sub
        End If
    End If

    L_BidFound = ""
    If t_FindPrior <> t_Prior Or _
       t_FindHand(int_FindHandNum) <> t_TestHand _
    Then
        'mClear_Click  ' Leave it and I can clear it. '
    End If
    bln_BidFound = False

    ' Set up DB module calls: '
GetBidFromDLL:
    Set NewBid = New BidBaseDB.bidder
    ' If "Class not registered", open BidBaseDB.dll with  '
    ' regsvr32.exe in xplorer2. For more information, see '
    ' the end of CardShark_BidBase_Programming_Notes.htm. '
    Dim v As String
    Dim xxx As String
    Dim pc As String ' Prior Calls including starting passes. '
    Dim position As String
      
    L_BidFound = ""

    With NewBid
    '''''''''''
        .DBPathAndName = str_DBPathAndName
        .PlayersHand = Trim$(t_FindHand(int_FindHandNum).Text)
        If Right$(t_Prior, Len(pr)) = pr Or Mid$(pr, 3, 1) = "@" Then
            .PriorCalls = pr
        Else
            .PriorCalls = Trim$(t_FindPrior)
        End If
       '.pos =  "P"
        .Vul = RTrim$(Left$(cmb_FindVul & " ", 1))
        .ScoringMethod = Trim$(Left$(cmb_Find_ScoreMeth & " ", 1))
        .GetDisclosures = False
        .BidToAnalyze = t_FindAnalysis4Bid

        '  Get bid:  '
        b = .bid
        If InStr(b, "@") > 0 Then
            L_BidFound.Left = 2000
            L_BidFound.Width = 900
            L_FindBid(3).Left = 300
        End If
        L_BidFound.Caption = b
        e = .ErrMsg
        If b <> "" Then
            lng_ID = .id
            If lng_ID > 0 Then
                L_FindBidID.Caption = "ID# " & lng_ID & "   BON: " & .BON
            End If
        End If
    End With
      
    If Len(t_FindPrior) = 7 Then
        prNTstart = " " & Left$(t_FindPrior & "       ", 7) & " " ' e.g.: " 1H-1N-P "
    If e = "" And InStr(" 1C-1N-P 1D-1N-P 1H-1N-P 1S-1N-P ", prNTstart) > 0 Then
        ' This catches any auction starting with one  '
        ' of the NT overcall sequences above.         '
        ' Most people play that NT systems are on in  '
        ' response to a 1N overcall, meaning that     '
        ' responses are the same as to an opening 1N. '
        ' But not only does duplicating all the open- '
        ' ing 1N response entries quadruple the space '
        ' needed in the database, keeping responses   '
        ' in sync for all sections would be difficult.'
        ' The alternative is to have this sub change  '
        ' the prior bids to eliminate the bid before  '
        ' 1N before looking for a response bid.  The  '
        ' only problem with this is that the opening  '
        ' bid may change how a NT response bid should '
        ' be made. For example, after (1H)-1N-(P), it '
        ' is possible that responder could have 5 H's '
        ' to which he would normally transfer,        '
        ' but now would not want to do so. To prevent '
        ' that, the entries for 1H-1N-P (for example) '
        ' should have entries for when systems are not'
        ' on and if no bid is found in those sections,'
        ' then change prior bids to just 1N-P and get '
        ' a bid from that section. '
        prNTO = pr
        pr = Mid$(pr, 4)
        t_FindPrior = pr
        b_FindBidForHand_Click
        str_Response = MyMsgBox("If systems are on for notrump overcalls," & vbCrLf & "BidBase looks in the 1N-P section for a bid." & vbCrLf & "Use that section to add or edit a bid.", 3, "ok", "")
        bln_Abort = False
        str_Response = MyMsgBox("If systems are NOT on," & vbCrLf & "add or edit entries in the" & vbCrLf & prNTO & " section.", 3, "ok", "")
        bln_Abort = False
        t_FindPrior = prNTO
    End If
    End If

    If e <> "" Then
        'ck_SaveNew.Value = 1
        ' The above was causing the existing '
        ' entry to be saved as a new entry.  '
        GoTo ExitFindBid
    End If

    If bln_TwoGrids Then
        If t_Show(0) = pr Then
            int_GridNum = 0
            ColorGrids 0
        ElseIf t_Show(2) = pr Then
            int_GridNum = 1
            ColorGrids 1
        Else
            mPreviewHandSpecs_Click
            GoTo ExitFindBid
        End If
    Else
        int_GridNum = 0
        ColorGrids 0
        i = 1
        Do While Mid$(t_FindPrior & "  ", i, 2) = "P-"
            i = i + 2
            DoEvents
        Loop
        If t_FindPrior <> "P" And _
           t_FindPrior <> "P-P" And _
           t_FindPrior <> "P-P-P" And _
           t_Show(0) <> Trim$(Mid$(t_FindPrior & "  ", i)) _
        Then
            mPreviewHandSpecs_Click
            GoTo ExitFindBid
        End If
    End If

    ' Find entry on the grid: '
Do While Grid1(0).Row > -1
    With Grid1(int_GridNum)
        If Grid1(int_GridNum).Columns(int_LastBidCol).Caption <> "ID" Then
            'MsgBox "ID# is supposed to be the last column and isn't."
            int_LastBidCol = Grid1(int_GridNum).Columns.Count - 1
        End If
        .Col = int_LastBidCol
        .Row = 0

        Do While Val(.Text) <> lng_ID And _
                 (.Columns(int_OrdNum_Col).Text <> "ZZZZZZ" Or .Row = 0)
            If .Row = .VisibleRows - 2 Then
                ' .Row will not increment past the last row '
                ' and doesn't give an error, so keep track. '
                Exit Do
            End If
            .Row = .Row + 1
            DoEvents
        Loop
        If Val(.Text) <> lng_ID Then
            ' Wasn't on screen. '
            .Row = 0
            ' No error is generated when Scroll can't '
            ' go back any further, so it will stay in '
            ' the loop forever if we don't check it.  '
            i = 0
            Do
                topID = Val(.Text)
                .Scroll 0, -1
                .Row = 0
                DoEvents
                If .VisibleRows = 1 Then Exit Do
            Loop While Val(.Text) <> lng_ID And topID <> Val(.Text)
        End If
    End With

'    Grid1_Click int_GridNum
    Grid1(int_GridNum).Col = 7

    If Grid1(1).Visible And int_GridNum = 0 Then
        Grid1(1).Col = 0
        Exit Do ' Kept looping. '
    Else
        Exit Do
    End If
    DoEvents
Loop

ExitFindBid:
    If L_BidFound <> "" Then
        b_EditFind(1).Caption = "Add"
        b_EditFind(1).Width = 550
    Else
        b_EditFind(1).Caption = "Add Bid"
        b_EditFind(1).Width = 1100
        f_Msg.Left = f_EnterHand.Left - 800
        f_Msg.Top = f_EnterHand.Top - f_Msg.Height - 100
        str_Response = MyMsgBox("No bid found.", 1, "ok", "")
    End If
    If Grid1(0).Row > Grid1(0).VisibleRows Then
        ' Entry is showing. Exit now. '
        DisplayBid2Edit 0
        b_FindBidForHand.Enabled = True
        Exit Sub
    End If

    If t_Show(0) <> t_FindPrior Then
        t_Show(0) = pr 't_FindPrior
        t_Show(1) = pr 't_FindPrior
    End If
    b_Show_Click
    b_FindBidForHand.Default = False
    t_FindPrior.SetFocus
    b_FindBidForHand.Enabled = True
    
    If Dir$(App.Path & "\ImportID.txt") <> "" Then
        Open App.Path & "\ImportID.txt" For Input As #7
        Input #7, lng_ID
        Close #7
        DisplayBid2Edit 99
        i = MsgBox("Entry displayed. Delete ImportID.txt?", vbYesNo)
        If i = vbYes Then
            Kill App.Path & "\ImportID.txt"
        End If
    End If
End Sub


Private Sub b_IncrSeq_Click()
    ' This + button right above the bid '
    ' order number box causes the bid   '
    ' order number to be incremented.   '

    Dim inc As String
    Dim n As Long
    Dim r3 As Long ' value of right 3 characters '
    Dim p As String
    Dim x As String
    Dim i As Integer
    Dim xl As Integer
    Dim mboxTop As Integer
    Dim mboxLeft As Integer
    
    mboxTop = f_Msg.Top
    mboxLeft = f_Msg.Left
    f_Msg.Top = frame0(6).Top - 1700
    f_Msg.Left = b_IncrSeq.Left + frame0(6).Left - 800

    If Len(t_Order) < 6 Then
        MsgBox "Order must be 6 characters."
        Exit Sub
    End If
    
    r3 = Val(Right$(t_Order, 3))

    If lng_IncrSeqAmt = 0 Then lng_IncrSeqAmt = 10
    x = "" & lng_IncrSeqAmt
    inc = MyMsgBox("Enter amount to increment Sequence by:", 1, x, "")
    f_Msg.Move mboxLeft, mboxTop
    mboxLeft = f_Msg.Left
    If inc = "" Then Exit Sub

    lng_IncrSeqAmt = Val(inc)
    If t_Order <> "" Then
        i = 0
        Do While (Mid$(t_Order, i + 1, 1) < "0" Or _
                  Mid$(t_Order, i + 1, 1) > "9") And _
                 i < Len(t_Order) - 1
            i = i + 1
            DoEvents
        Loop
        If i > 0 Then
            p = Left$(t_Order, i)
            x = Mid$(t_Order, i + 1)
            xl = Len(x)
            n = Val(x)
            t_Order = p & Right$("000000" & (n + lng_IncrSeqAmt), xl)
        Else
            'n = Val(t_Order)
            n = r3
            t_Order = Left$(t_Order, 3) & Right$("000" & (n + lng_IncrSeqAmt), 3)
        End If
        'If ck_SaveNew.Value = 0 Then
        '    str_Response = MyMsgBox("Create a new entry?", 1, "YN", "")
        '    If str_Response = "yes" Then
        '        ck_SaveNew.Value = 1
        '    Else
        '        ck_SaveNew.Value = 0
        '    End If
        'End If
        ck_Saved.Value = 0
    End If
End Sub


Private Sub b_IncrSeq_MouseDown( _
            Button As Integer, _
            Shift As Integer, _
            x As Single, _
            Y As Single)
    ' This button increments the entry's '
    ' Order number in the Input Boxes.   '

    Dim n As Long
    Dim p As String
    Dim p2 As String
    Dim i As Integer
    Dim xl As Integer
Exit Sub ' Don't know why both this and Click. '
    If Button = 2 And lng_IncrSeqAmt <> 0 Then
        If t_Order <> "" Then
            i = 0
            Do While (Mid$(t_Order, i + 1, 1) < "0" Or _
                      Mid$(t_Order, i + 1, 1) > "9") And _
                     i < Len(t_Order) - 1
                i = i + 1
                DoEvents
            Loop
            If i > 0 Then
                p = Left$(t_Order, i)
                p2 = Mid$(t_Order, i + 1)
                xl = Len(p2)
                n = Val(p2)
                t_Order = p & Right$("000000" & (n + lng_IncrSeqAmt), xl)
            Else
                n = Val(t_Order)
                t_Order = Right$("000000" & (n + lng_IncrSeqAmt), 6)
            End If

            'ck_SaveNew = 1
        Else
            Grid1(int_GridNum).Text = _
                Val(Grid1(int_GridNum).Columns(int_OrdNum_Col).Text) _
                 + lng_IncrSeqAmt
            Grid1(int_GridNum).Row = Grid1(int_GridNum).Row + 1
        End If
    Else
        b_IncrSeq_Click
    End If
End Sub


Private Sub b_SaveEntry_Click()
    mSave_Click
    ' Entry may not be displayed '
    ' due to a glitch in VB6.    '
    ' Timer3 will try again. '
    'str_Job = "Saving Entry"
    'Timer3.Enabled = True
    str_Job = ""
End Sub


Private Sub b_SaveHands_Click()
    Dim FiNum As Integer
    Dim h As String
    Dim i As Integer
    Dim f As String

    f = "\EnteredHands.txt"

    If cmb_FindBid.ListCount = 0 Then
        str_Response = MyMsgBox("There are no hands to save", 1, "ok", "")
        Exit Sub
    End If

    str_Response = MyMsgBox("Replace the existing file?" & vbCr & _
                            "(Press N to append hands.)", 2, "YN", "Save Hands")

    FiNum = FreeFile

    If str_Response = "yes" Then
        Open str_BBpath & f For Output As #FiNum
    Else
        Open str_BBpath & f For Append As #FiNum
    End If

    For i = 0 To cmb_FindBid.ListCount - 1
        Print #FiNum, cmb_FindBid.List(i)
    Next
    Close #FiNum
    b_DelRef.Visible = False
    Beep
End Sub


Private Sub b_LoadHands_Click()
    Dim FiNum As Integer
    Dim h As String
    Dim i As Integer
    Dim f As String

    f = "\EnteredHands.txt"

    If Dir$(str_BBpath & f) = "" Then
        str_Response = MyMsgBox("There are no saved hands to load", 1, "ok", "")
        Exit Sub
    End If

    cmb_FindBid.Visible = True

    If cmb_FindBid.ListCount > 0 Then
         str_Response = MyMsgBox("Replace the existing hands?" & vbCr & _
                                 "(Press N to add hands to end.)", 2, "ync", "")
        If str_Response = "yes" Then
            cmb_FindBid.Clear
        ElseIf str_Response = "cancel" Then
            Exit Sub
        End If
    End If

    FiNum = FreeFile

    Open str_BBpath & f For Input As #FiNum

    'Line Input #FiNum, h
    'If cmb_FindBid.ListCount = 0 Then
    '    cmb_FindBid.AddItem h
    'End If

    Do While Not EOF(FiNum)
        Line Input #FiNum, h
        If h <> "" Then
            cmb_FindBid.AddItem h
        End If
    Loop
    Close #FiNum
End Sub


Private Sub b_LoadRefs_Click()
    Dim FiNum As Integer
    Dim h As String
    Dim i As Integer
    Dim f As String

    f = "\Refs.txt"

    If Dir$(str_BBpath & f) = "" Then
        str_Response = MyMsgBox("There are no saved " & _
                                "references to load", 2, "ok", "Load References")
        Exit Sub
    End If

    list_Refs.Visible = True

    If list_Refs.ListCount > 1 Then
         str_Response = MyMsgBox("Replace the existing references?" & vbCr & _
                                 "(Press N to add references to the end.)", 2, "ync", "")
        If str_Response = "yes" Then
            list_Refs.Clear
        ElseIf str_Response = "cancel" Then
            Exit Sub
        End If
    End If

    FiNum = FreeFile

    Open str_BBpath & f For Input As #FiNum

    Line Input #FiNum, h
    If list_Refs.ListCount = 0 Then
        list_Refs.AddItem h
    End If

    Do While Not EOF(FiNum)
        Line Input #FiNum, h
        If h <> "" Then
            list_Refs.AddItem h
        End If
    Loop
    Close #FiNum
End Sub


Private Sub b_SaveRefs_Click()
    Dim FiNum As Integer
    Dim h As String
    Dim i As Integer
    Dim f As String

    f = "\Refs.txt"

    If list_Refs.ListCount = 0 Then
         str_Response = MyMsgBox("There are no references to save", 1, "ok", "Save References")
        Exit Sub
    End If

    If Dir$(str_BBpath & f) <> "" Then
        str_Response = MyMsgBox("Replace the existing file?" & vbCr & _
                                "(Press N to append hands.)", 2, _
                                "YNC", "Save References")
    End If

    FiNum = FreeFile

    If str_Response = "yes" Then
        Open str_BBpath & f For Output As #FiNum
    Else
        Open str_BBpath & f For Append As #FiNum
    End If

    For i = 0 To list_Refs.ListCount - 1
        Print #FiNum, list_Refs.List(i)
    Next
    Close #FiNum
    Beep

End Sub


'*****************************************************
' Procedure: b_Show_Click                            '
' Purpose:                                           '
    ' Displays in the grid(s) the records specified. '
    ' In cmb_Show, the user enters the field to use  '
    ' in selecting records to display, then in t_Show'
    ' he enters the data to look for in that field.  '
    ' In the cmb_Sort fields, he enters the fields on'
    ' which to sort the grid(s). The sort order is   '
    ' stored in the variable 's1' with a call to the '
    ' function Sort_Order_Specs.                     '
'*****************************************************

Private Sub b_Show_Click()
    Dim s1 As String
    Dim g As Integer
    Dim n As Integer
    Dim pbid1 As String
    Dim pbid2 As String
    Dim w0 As String
    Dim t0 As String
    Dim t1 As String
    Dim f As String
    Dim bid As String
    Dim i As Integer

    s1 = Sort_Order_Specs

    If t_Show(2) <> "" Then
        mDupesGrid(0).Checked = False
    End If

    bln_Abort = False
    If bln_ShowGrid2 Then n = 1 Else n = 0
    For g = n To 1 ' "g" is the grid number to update. '
        If str_Job = "Load 1 grid" Then
            ' From "b_DupeEntry". '
            If int_GridNum = 0 Then
                If g = 0 Then g = 1
            Else
                If g = 1 Then Exit For
            End If
        End If

        If mNextSection.Checked Then
            g = int_GridNum ' the current grid number
        End If

            t0 = Trim$(t_Show(g * 2))
            t1 = Trim$(t_Show(g * 2 + 1))
            t0 = Replace$(t0, " ", "-")
            t1 = Replace$(t0, " ", "-")
            If t0 = "" Then t0 = " "
            If t1 = "" Then t1 = " "
            f = cmb_Show(g)
            If f = "Bid_Name" Then
                f = "BidNames.Bid_Name"
            End If
            w0 = f & " BETWEEN '" & t0 & "' AND '" & t1 & "' "

        If Left$(w0 & Space$(5), 5) <> "WHERE" And w0 <> "" Then
            w0 = "WHERE " & w0
        End If

        'On Local Error GoTo ShowError
        If InStr(s1, "@") > 0 Then
            s1 = "Bids.Prior_Bids, Bids.Bid_Order_Number, BidNames.Selected, BidNames.SubCat_Selected, Bids.Pct_Used, BidNames.Bid_Name"
        End If
        ' qSel is a global string variable. It is set in Sub InitializeDatabase.'
        ' SELECT BidNames.Selected, BidNames.SubCat_Selected, Bids.Pct_Used,    '
        '        Bids.RoundOfBidding, Bids.Player, Bids.Prior_Bids,             '
        '        Bids.Bid_Order_Number, Bids.Bid,BidNames.Bid_Name, Bids.HCP,   '
        '        Bids.Dist_Pts, Bids.Cust_Pts, Bids.A_Cust, Bids.K_Cust,        '
        '        Bids.Q_Cust, Bids.J_Cust, Bids.T_Cust, Bids.Winrs, Bids.Losers,'
        '        Bids.QTs, Bids.Stopped_Suits, Bids.Shape, Bids.Intermed,       '
        '        Bids.xNum, Bids.xHCP, Bids.sNum, Bids.sHCP, Bids.hNum,         '
        '        Bids.hHCP, Bids.dNum, Bids.dHCP, Bids.cNum, Bids.cHCP,         '
        '        Bids.cmNum, Bids.cmHCP, Bids.Vul, Bids.Pos, Bids.Artificial,   '
        '        Bids.Alert, Bids.Explanation, Bids.Test_Hand, Bids.Asks,       '
        '        Bids.DateOfEntry, Bids.Best_Contract, Bids.Num_Used,           '
        '        Bids.Source, Bids.LHO, Bids.Pard, Bids.RHO, Bids.ID            '
        '        FROM BidNames RIGHT JOIN Bids ON BidNames.IDnum = Bids.ConvID  '
        Data1(g).RecordSource = qSel & w0 & " ORDER BY " & s1
        Data1(g).Refresh ' If error, make sure the cmb_Show() boxes are valid. '
                         ' Usually should be Prior_Bids. '
        DoEvents
        Grid1(g).Refresh
        DoEvents
        If bln_Abort Then Exit Sub

        If Not bln_TwoGrids Or _
           mNextSection.Checked Or _
           (mDupesGrid(0).Checked And f_ShowSort.Visible = False) _
        Then
            Exit For
        End If
    Next

    If str_Job = "Divider" Then Exit Sub

    On Local Error GoTo 0
    If mFourGrids.Checked Then
        For g = 2 To 3
            Data1(g - 1).RecordSource = _
                qSel & _
                " WHERE Bids.Prior_Bids = '" & t_Show(g) & _
                "' ORDER BY " & s1 & ";"
            Data1(g - 1).Refresh
        Next
    End If

    If mNextSection.Checked Then
        ' Only int_GridNum grid is being changed. '
        Me.MousePointer = vbHourglass
        mTestHands_Click
        Me.MousePointer = vbDefault

    ElseIf frm_Tutorial.Visible = False And _
           str_Job <> "Load 1 grid" _
    Then
        If Grid1(0).Row = -1 Then
            Exit Sub
        End If

        If bln_Abort Then
            GoTo ExitShow
        End If

        b_Show.Default = False
        str_Job = "Testing"
        Me.MousePointer = vbHourglass
        If cmb_Show(0) = "Prior_Bids" And Not bln_ShowGrid2 Then
            Grid1(0).Row = 0
            int_GridNum = 0
            ColorGrids 0
            mTestHands_Click
            If Not bln_ShowGrid2 Then
                GoTo ExitShow
            End If
        End If
        
        If cmb_Show(1) = "Prior_Bids" And _
           Grid1(1).Visible And Grid1(1).Row > -1 _
        Then
            Grid1(1).Row = 0
            int_GridNum = 1
            ColorGrids 1
            ' If testing is needed, see if the grid is changing. '
            'mTestHands_Click
        End If
        Me.MousePointer = vbDefault
    End If

ExitShow:
    If bln_Abort Then Exit Sub
    'f_ShowSort.Visible = False
    b_Show.Default = False
    If bln_GridFwdBack = False Then
      If t_Show(0) = str_GridPick(0, 0, int_GridPicksNum) Then
        If t_Show(1) = str_GridPick(0, 1, int_GridPicksNum) Then
          If cmb_Show(0).ListIndex = Val(str_GridPick(0, 2, int_GridPicksNum)) Then
            If str_GridPick(1, 0, int_GridPicksNum) = "" And bln_TwoGrids = False Then
              GoTo ExitShow2
            ElseIf t_Show(2) = str_GridPick(1, 0, int_GridPicksNum) Then
              If t_Show(3) = str_GridPick(1, 1, int_GridPicksNum) Then
                If cmb_Show(1).ListIndex = Val(str_GridPick(1, 2, int_GridPicksNum)) Then
                  GoTo ExitShow2
                End If
              End If
            End If
          End If
        End If
      End If
    End If

    If bln_GridFwdBack Then
        GoTo ExitShow2
    End If

    Do While str_GridPick(0, 0, int_GridPicksNum) <> "" _
         And int_GridPicksNum < 50

        int_GridPicksNum = int_GridPicksNum + 1
        DoEvents
    Loop
    str_GridPick(0, 0, int_GridPicksNum) = t_Show(0)
    str_GridPick(0, 1, int_GridPicksNum) = t_Show(1)
    str_GridPick(0, 2, int_GridPicksNum) = "" & cmb_Show(0).ListIndex
    If bln_TwoGrids Then
        str_GridPick(1, 0, int_GridPicksNum) = t_Show(2)
        str_GridPick(1, 1, int_GridPicksNum) = t_Show(3)
        str_GridPick(1, 2, int_GridPicksNum) = "" & cmb_Show(1).ListIndex
    End If
    If int_GridPicksNum > 0 Then
        mGoBack.Enabled = True
    End If

ExitShow2:
    str_LastPBs = "x"
    bln_AlreadyToldNoDupes = False
    bln_ShowGrid2 = False
    If mDupesGrid(0).Checked Then
        If Grid1(0).Row < 0 Then
            Grid1(0).Row = 0
        End If
        Grid1_Click 0
    End If
    On Local Error GoTo 0
    Me.MousePointer = vbDefault
    If frm_Tutorial.Visible And frm_Tutorial.Caption = "Show/Sort Grids" Then
        bln_Waiting = False
    End If
Exit Sub

ShowError:
    str_Response = MyMsgBox("Error in parameters entered.", 1, "ok", "")
    Resume ExitShow
End Sub


'*******************************************
' Procedure: TestHand                      '
' Purpose:                                 '
'     This is called by the Save function. '
'     The TestHands menu function uses its '
'     own routines getting data from the   '
'     grid rather than from the Input Boxes.
'*******************************************

Private Sub TestHand()
    Dim a As String
    Dim b As String
    Dim e As String
    Dim s As String
    Dim i As Long
    Dim v As String
    Dim BON As String
    Dim trump As Integer

    a = Left$(t_Sel & " ", 1) & Left$(t_Sub & " ", 1) & Left$(t_Pct & " ", 1)
    If a <> "   " And Not (Val(t_Pct) > 1) Then
        'MsgBox "Saved, but deactivated entries are not tested."
        ck_Saved.Value = 1
        Exit Sub
    End If

On Local Error GoTo BidderError

    ' Set up DB module calls: '
    Set NewBid = New BidBaseDB.bidder
    ' If "Class not registered", open BidBaseDB.dll with  '
    ' regsvr32.exe in xplorer2. For more information, see '
    ' the end of CardShark_BidBase_Programming_Notes.htm. '
    With NewBid
    '''''''''''
        .DBPathAndName = str_DBPathAndName
        .PlayersHand = t_TestHand
        If str_Job = "Saving" Then
            .Activated = a
            ' Activated can be "   ", "  0", "0  ", "00 ", "  W", etc. '
        End If
        .PriorCalls = t_Prior
        i = InStr(t_Prior, "@")
        If Mid$(t_Prior, i + 1, 2) = "Mw" Then ' Minorwood '
            s = Mid$(t_Prior, i - 1, 1)
            trump = InStr("CD", s)
        ElseIf i > 2 Then
            trump = Val(Mid$(t_Prior, i + 1, 1))
        End If

        i = cmb_Vul.ListIndex
        If i > 0 Then
            v = Mid$("UFBNVX", i, 1)
        End If
        .Vul = v
        If t_Pct = "I" Then
            .ScoringMethod = "I"
        ElseIf t_Pct = "M" Then
            .ScoringMethod = "M"
        End If
        .TestID = lng_ID
        '.Combo = t_Qty(5) ' When testing an entry with "M" in PriorCalls. '
        .GetDisclosures = False
        If str_Job = "Saving" Then
            .CheckAllPcts = True
        End If

        '  Get bid:  '
        b = .bid
        e = .ErrMsg
        i = .id
        If i > 0 Then BON = .BON
    End With
    ''''''''

    If e <> "" Then
        Dim m As Integer
        Dim n As Integer
        Dim k As Integer
        k = 1
        m = Len(e) \ 2
        
        If InStr(e, Chr$(13)) > 0 Then
            k = 2
        ElseIf m > 28 Then
            n = m
            Do While Mid$(e, n, 1) <> " " And m - n < 6
                n = n - 1
            Loop
            If Mid$(e, n, 1) <> " " And m - n = 5 Then
                n = m + 1
                Do While Mid$(e, n, 1) <> " " And n - m < 6
                    n = n - 1
                Loop
            End If
            If Mid$(e, n, 1) = " " Then
                Mid$(e, n, 1) = vbCr
                k = 2
            End If
        End If
        
        If InStr(e, "also matches") > 0 Then
            k = 4
        Else
            'Stop
            k = 2
        End If
        str_Response = MyMsgBox(e, k, "ok", "BidBase Testing")
        If str_Job = "Duping" Then str_Job = "Duping Error"
        Exit Sub
    End If

    If b <> "" Then
      If i = lng_ID And _
       (Left$(BON, 1) = "9" And _
         Left$(t_Order, 1) < "9") Or _
       (Left$(BON, 1) < "9" And _
         Left$(t_Order, 1) = "9") Or _
       BON = t_Order _
      Then
    
        ' Don't give the other message if entry #i has a BON < 900000 '
        ' and the test hand entry is 900000+ because the #i entry is '
        ' an optional convention which overrides the 9xxxxx entries  '
        ' which are SAYC. '
        str_Response = MyMsgBox("Saved, and entry tested okay.", 1, "ok", "Saving Entry")
        ' Should also see if position and vul are the same or not. '
      ElseIf lng_ID > 0 Then
        ' Shouldn't give this message when #i is a deselected entry. '
        str_Response = MyMsgBox("The test hand also" & vbCr & _
                                "matches Entry ID# " & i & vbCr & _
                                "Bid Order Number " & BON & ".", 3, "ok", "")
        'FillCompareGrid
      End If
    End If
    
    ck_Saved.Value = 1

    DoEvents

BidderError:
    On Local Error GoTo 0

End Sub


Private Sub b_Show_DragDrop(source As Control, _
                            x As Single, _
                            Y As Single)
    With source
        .Drag vbEndDrag
        .Move .Left + b_Show.Left + x - sng_DragoffsetX, _
              .Top + b_Show.Top + Y - sng_DragoffsetY
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub b_ShowFindHandList_Click(Index As Integer)
    bln_SkipIt = False

    int_FindHandNum = Index
    If Index = 0 Then
        t_FindHand(0).ForeColor = vbBlack
        t_FindHand(1).ForeColor = RGB(150, 150, 150)
        cmb_FindBid.Top = 920
        cmb_FindBid.Height = 2540
    Else
        t_FindHand(1).ForeColor = vbBlack
        t_FindHand(0).ForeColor = RGB(150, 150, 150)
        cmb_FindBid.Top = 1315
        cmb_FindBid.Height = 2160
    End If
    
    If cmb_FindBid.Visible Then
        cmb_FindBid.Visible = False
        t_FindAnalysis4Bid.Visible = True
        b_DelFindHand(int_FindHandNum).Visible = False
        t_FindHand(int_FindHandNum).SetFocus
    Else
        cmb_FindBid.ZOrder 0
        cmb_FindBid.Visible = True
        'cmb_FindBid.Top = 1300
        b_ShowFindHandList(int_FindHandNum).Visible = True
        b_DelFindHand(int_FindHandNum).Visible = True
        t_FindAnalysis4Bid.Visible = False
        cmb_FindBid.SetFocus
    End If
End Sub


Private Sub b_ShowRefList_Click(Index As Integer)
    If cmb_FindRef.Visible = True Then
        cmb_FindRef.Visible = False
    Else
        cmb_FindRef.Visible = True
    End If
End Sub




Private Sub b_ShowRefs_Click()
    bln_SkipIt = False
    If list_Refs.Visible Then
        list_Refs.Visible = False
        b_DelRef.Visible = False
        t_Ref.SetFocus
    Else
        list_Refs.Visible = True
        b_DelRef.Visible = True
        list_Refs.SetFocus
    End If
End Sub

Private Sub b_ShowSortDefaults_Click()
    cmb_Sort(0).ListIndex = 2 ' Round      '
    cmb_Sort(1).ListIndex = 4 ' Prior Bids '
    cmb_Sort(2).ListIndex = 5 ' Bid_Ord_Num'
    cmb_Sort(3).ListIndex = 1 ' Pct_Used   '
    cmb_Sort(4).ListIndex = 6 ' Bid        '
End Sub


Private Sub b_TutorDetails_Click()
    Dim xx As Long

    xx = Shell("BidBase File Viewer.exe " & _
               str_BBpath & "\Notes\" & _
               str_TutorDetails, _
               vbNormalFocus)
End Sub

Private Sub b_Unhide_Click()
    ' Unhide grid columns which have been hidden by user. '
    Dim n As Integer
    Dim c As Integer ' column number
    Dim gn As Integer ' grid number
    Dim s As Integer ' grid split number

    On Local Error Resume Next
    For n = 0 To 3
        c = Val(cmb_HiddenCols(n).Text)
        If Trim$(cmb_HiddenCols(n).Text) <> "" Then
            If n < 2 Then
                gn = 0
            Else
                gn = 1
            End If

            If n = 1 Or n = 3 Then
                s = 1
            Else
                s = 0
            End If
            Grid1(gn).Splits(s).Columns(c).Visible = True
            cmb_HiddenCols(n).RemoveItem cmb_HiddenCols(n).ListIndex
        End If
    Next
    On Local Error GoTo 0
End Sub


Private Sub b_ViewGrid2_Click()
    If t_Show(2) <> "" And Not mTwoGrids.Checked Then
        mTwoGrids_Click
    Else
        bln_ShowGrid2 = True ' Don't re-display top grid. '
        b_Show_Click
        bln_ShowGrid2 = False
    End If
End Sub

Private Sub ck_Artificial_Click()
    L_Status.Clear
    If bln_OrigData <> ck_Artificial Then
        ck_Saved.Value = 0
    End If
End Sub


Private Sub ck_Artificial_GotFocus()
    bln_OrigData = ck_Artificial
End Sub


Private Sub ck_Continuation_Click()
    If ck_Continuation.Value = 1 Then
        cmb_Alert.ListIndex = 4
    Else
        cmb_Alert.ListIndex = 0
    End If
End Sub


Private Sub ck_SaveNew_Click()
    Dim i As Integer

    If ck_SaveNew.Value = 1 Then
        L_IDnum = "ID#: "
        ck_Saved.Value = 0
    ElseIf lng_ID > 0 Then
        L_IDnum = "ID#: " & lng_ID
        ck_Saved.Value = 1
    End If

    L_Date_Saved = ""

    ' If file is referenced in notes, '
    ' leave the button as "Y":        '
    i = InStr(LCase$(t_Disclosure), ".txt")
    If i = 0 Then i = InStr(LCase$(t_Disclosure), ".htm")
    If i > 0 Then Exit Sub

    ' If a "New Entry", a new ID# will be '
    ' assigned, so the old ID#-based note '
    ' file (if any) will not apply to the '
    ' entry:                              '
    If ck_SaveNew.Value = 1 Then
        'lng_ID = 0 ' This needs to be left alone in case New Entry gets unchecked. '
        b_Discl.Picture = LoadPicture(str_BBpath & "\ButtonN.bmp")
        Exit Sub
    End If

    ' "New Entry" box unchecked, so see   '
    ' if note file exists for this entry  '
    ' and if so, reset the button to "Y": '
    LookForNoteFile t_Disclosure, _
                    t_Prior, _
                    cmb_Bid, _
                    "" & lng_ID, _
                    False, _
                    True, _
                    False
End Sub


Private Sub ck_Unopposed_Click()
    bln_NoHiLitePrior = True
    t_FindPrior.SetFocus
End Sub

Private Sub cmb_Alert_Click()
    If cmb_Alert.ListIndex = 4 Then
        ' Eventually a Convention Name Continuation field '
        ' should be added to the database. The problem is '
        ' how to indicate that a prior bid by partner was '
        ' one of different option. Example:               '
        ' 1D-P-2D-P, 3C - where 2D could have been either '
        ' a weak raise or a strong one (Inverted Minor).  '
        ' If "Inverted Minor" is put in the Conven. field '
        ' of the 3C bid entry, it isn't clear that it     '
        ' applies to a past bid, not the current one.     '
        ' A better solution would probably be to create a '
        ' table which shows the values and distributions  '
        ' of other hands deduced from their bids. '
        ck_Continuation.Value = 1
    Else
        ck_Continuation.Value = 0
    End If
End Sub


Private Sub cmb_Bid_KeyPress(KeyAscii As Integer)
    ' Uppercase bid letters (e.g.: 1C from 1c) '
    Dim k As String

    k = Chr$(KeyAscii)
    
    If InStr("cdhsnpr", k) > 0 Then
        KeyAscii = Asc(UCase$(k))
    End If

    If k = "@" Then
        If cmb_Bid.Width <> 1015 Then
            t_Prior.Width = 3380
            t_Order.Left = 4990
            cmb_Bid.Left = 5835
            cmb_Bid.Width = 1050
            cmb_Bid = cmb_Bid & "@"
            cmb_Bid.SelStart = 99
            KeyAscii = 0
        End If
    End If
End Sub


Private Sub cmb_Bid_LostFocus()
    Dim b As String
    Dim n As Long
    Dim s As String
    Dim sn As Long

    If InStr(t_Prior, "M") > 0 _
       And _
       (InStr(cmb_Bid, "H") > 0 Or _
        InStr(cmb_Bid, "S") > 0 Or _
        ((InStr(t_Qty(1), "x") > 0 Or _
         InStr(t_Qty(2), "x") > 0) _
         And _
         InStr(cmb_Bid, "x") > 0)) _
    Then
        bln_Abort = False
        str_Response = MyMsgBox("With 'M' in the Prior Bids, an" & vbCrLf & _
                                "H/S bid may not work as intended." & vbCrLf & _
                                "(nor with an x bid with x in H/S)", 3, _
                                "Docs", "Combined H/S Entries")
        If str_Response = "docs" Then
            str_NoteFile(0) = "CardShark_BidBase_Combining_HS_Entries.htm"
            ShowDocs
            Exit Sub
        End If
    End If
    
    If InStr(cmb_Bid.Text, "@") > 0 Then
        ' Slam exploring bid made.  '
        ' Increase cmb_Bid.Size and '
        ' location to make room.    '
        If cmb_Bid.Width <> 1015 Then
            t_Prior.Width = 3380
            t_Order.Left = 4990
            cmb_Bid.Left = 5835
            cmb_Bid.Width = 1050
        End If
    ElseIf cmb_Bid.Width = 1050 Then
        t_Prior.Width = 3765
        t_Order.Left = 5375
        cmb_Bid.Left = 6240
        cmb_Bid.Width = 650
    End If

    If Trim$(t_Order) <> "" Or cmb_Bid.Text = "" Then Exit Sub

    b = cmb_Bid.Text

    If b = "P" Then
        t_Order.Text = "990000"
        Exit Sub

    ElseIf b = "D" Or b = "R" Then
        t_Order.Text = "980000"
        Exit Sub
    End If

    n = Val(b)
    s = Mid$(b, 2)
    sn = InStr("CDHSN", s)
    If n = 0 Or n > 7 Or sn = 0 Then Exit Sub

    t_Order.Text = "9" & sn & n & "000"
End Sub


Private Sub cmb_ConvName_KeyPress(Index As Integer, KeyAscii As Integer)
    If Len(cmb_ConvName(Index).Text) > 50 Then
        KeyAscii = 0
    End If
End Sub


Private Sub cmb_ConvName_KeyUp(Index As Integer, _
                               KeyCode As Integer, _
                               Shift As Integer)

    Dim i As Integer
    Dim t As String
    Dim j As Integer

    ' Hold Ctrl key to stop Autocomplete: '
    If Shift = 2 Or (Shift = 0 And KeyCode = 37) Then Exit Sub

    ' Do autocomplete: '
    With cmb_ConvName(Index)
    ''''''''''''''''''''''''
        ' If there isn't text, or if Backspace '
        ' or Delete was pressed, exit sub:     '
        If KeyCode = 8 Or KeyCode = 48 Or .Text = "" Then
            Exit Sub
        End If

        t = Left$(.Text, .SelStart)
        j = Len(t)
        If Mid$(.Text, j + 1) <> "" Then
            Exit Sub
        End If

        ' Find text in the list: '
        i = FindMatch(cmb_ConvName(Index), t, False, .ListIndex)
             ' Above: False <> exact match. Start at .ListIndex '
        If i < 0 Then
            i = 0
        Else
            .ListIndex = i
            .SelStart = j
            .SelLength = 32768
        End If
    End With
    ''''''''
End Sub


Private Sub cmb_FindBid_Click()
    If t_FindHand(int_FindHandNum).Text = _
       cmb_FindBid.List(cmb_FindBid.ListIndex) _
    Then
        Exit Sub
    End If
    If cmb_FindBid.ListIndex > -1 Then
        t_FindHand(int_FindHandNum).Text = _
            cmb_FindBid.List(cmb_FindBid.ListIndex)
    End If
    If bln_SkipIt Then
        bln_SkipIt = False
        Exit Sub
    End If

    'cmb_FindBid.Visible = False
    t_FindAnalysis4Bid.Visible = True
    'b_DelFindHand.Visible = False
End Sub


Private Sub cmb_FindRef_Click()
    t_Find_Ref.Text = cmb_FindRef.List(cmb_FindRef.ListIndex)
End Sub


Private Sub cmb_ForeColors_Click(Index As Integer)
    ' Change foreground (text) colors '
    ' to user-specified color.        '

    Dim fgc As Long ' foreground color code
    Dim n As Integer

    If cmb_ForeColors(Index).ListIndex - 1 = _
       cmb_ForeColors(Index).ListCount _
    Then
        fgc = vbBlack
    Else
        fgc = QBColor(cmb_ForeColors(Index).ListIndex)
    End If

    Select Case Index
    Case Is = 0
        ' Colors for grid with focus: '
        Grid1(int_GridNum).ForeColor = fgc
        lng_FGGrid1 = cmb_ForeColors(Index).ListIndex

    Case Is = 1
        ' Colors for grids without focus: '
        For n = 0 To 3
            If n <> int_GridNum Then
                Grid1(n).ForeColor = fgc
            End If
        Next
        lng_FGGrid2 = cmb_ForeColors(Index).ListIndex

    Case Is = 2
        ' Input and Combo Boxes: '
        t_AsksFor.ForeColor = fgc
        t_Disclosure.ForeColor = fgc
        t_Ref.ForeColor = fgc
        cmb_StrDiscl.ForeColor = fgc
        For n = 1 To 4: t_Len(n).ForeColor = fgc: Next
        For n = 0 To 5: t_HCP(n).ForeColor = fgc: Next
        t_Hand_HCPs.ForeColor = fgc
        t_BBPoints.ForeColor = fgc
        t_CustPoints.ForeColor = fgc
        For n = 0 To 4: t_CustPts(n).ForeColor = fgc: Next
        t_Intermediates.ForeColor = fgc
        t_Losers.ForeColor = fgc
        t_Pct.ForeColor = fgc
        t_Prior.ForeColor = fgc
        t_QTs.ForeColor = fgc
        For n = 0 To 5: t_Qty(n).ForeColor = fgc: Next
        t_Stoppers.ForeColor = fgc
        t_Sel.ForeColor = fgc
        t_Sub.ForeColor = fgc
        t_Round.ForeColor = fgc
        t_Player.ForeColor = fgc
        t_Order.ForeColor = fgc
        t_TestHand.ForeColor = fgc
        t_Distrib_Pts.ForeColor = fgc
        t_Winners.ForeColor = fgc

        'For n = 0 To 5
        '    t_DupeEntry(n).ForeColor = fgc
        'Next

        cmb_Alert.ForeColor = fgc
        cmb_Bid.ForeColor = fgc
        For n = 0 To 1
            cmb_ConvName(n).ForeColor = fgc
        Next
        cmb_Position.ForeColor = fgc
        cmb_Shape.ForeColor = fgc
        For n = 0 To 4
            cmb_Sort(n).ForeColor = fgc
        Next
        For n = 0 To 5
            t_Show(n).ForeColor = fgc
        Next
        For n = 0 To 1
            cmb_Show(n).ForeColor = fgc
        Next
        For n = 0 To 3
            cmb_ForeColors(n).ForeColor = fgc
            cmb_BackColors(n).ForeColor = fgc
        Next

        cmb_Vul.ForeColor = fgc

        t_BestContract.ForeColor = fgc
        int_FGInputs = cmb_ForeColors(Index).ListIndex

    Case Is = 3
        ' Input Area, Frames and Labels: '
        For n = 0 To 7
            frame0(n).ForeColor = fgc
        Next
        For n = 0 To 9
            L_Bid(n).ForeColor = fgc
        Next
        For n = 0 To 15
            If n = 2 Then n = 3
            If n = 14 Then n = 15
            L_HandStr(n).ForeColor = fgc
        Next
        For n = 0 To 8
            L_Suits(n).ForeColor = fgc
        Next
        For n = 0 To 2
            L_Other(n).ForeColor = fgc
        Next
        For n = 0 To 9
            L_Discl(n).ForeColor = fgc
        Next
        L_Date_Saved.ForeColor = fgc
        ck_Artificial.ForeColor = fgc
        ck_AutoFill.ForeColor = fgc
        ck_Saved.ForeColor = fgc
        ck_SaveNew.ForeColor = fgc
        L_IDnum.ForeColor = fgc

        'L_DupeEntry.ForeColor = fgc
        'f_Dupes.ForeColor = fgc

        f_OtherPlayers.ForeColor = fgc

        ' Other frames and labels: '

        f_Unhide.ForeColor = fgc
        For n = 0 To 3
            L_Unhide(n).ForeColor = fgc
        Next

        f_ShowSort.ForeColor = fgc
        For n = 0 To 15
            L_Show(n).ForeColor = fgc
        Next

        f_EnterHand.ForeColor = fgc
        For n = 0 To 6
            L_FindBid(n).ForeColor = fgc
        Next
        ck_Unopposed.ForeColor = fgc

        f_GridColors.ForeColor = fgc
        For n = 0 To 5
            L_Colors(n).ForeColor = fgc
        Next

        Me.ForeColor = fgc
        int_FGForm = cmb_ForeColors(Index).ListIndex

    End Select
End Sub


Private Sub cmb_BackColors_Click(Index As Integer)
    ' Change background colors '
    ' to user-specified color. '
    ' (Note: Label backgrounds '
    ' which are not set to a   '
    ' fixed color are set to   '
    ' transparent, so they do  '
    ' not need to be changed.) '

    Dim bgc As Long
    Dim n As Long

    If cmb_BackColors(Index).ListIndex - 1 = _
       cmb_BackColors(Index).ListCount _
    Then
        bgc = vbActiveBorder
    Else
        bgc = QBColor(cmb_BackColors(Index).ListIndex)
    End If

    g_OtherHands.BackColorFixed = vbYellow
    g_OtherHands.ForeColorFixed = vbBlack
    g_OtherHands.BackColor = vbBlack
    g_OtherHands.ForeColor = vbYellow

    Select Case Index
    Case Is = 0
        ' Grid with focus: '
        Grid1(int_GridNum).BackColor = bgc
        int_BGGrid1 = cmb_BackColors(Index).ListIndex

    Case Is = 1
        ' Grids without focus: '
        For n = 0 To 3
            If n <> int_GridNum Then
                Grid1(n).BackColor = bgc
            End If
        Next
        int_BGGrid2 = cmb_BackColors(Index).ListIndex

    Case Is = 2
        ' Input area text and combo boxes: '
        t_AsksFor.BackColor = bgc
        t_Disclosure.BackColor = bgc
        t_Ref.BackColor = bgc
        For n = 0 To 5
            t_HCP(n).BackColor = bgc
        Next
        t_Hand_HCPs.BackColor = bgc
        t_BBPoints.BackColor = bgc
        t_CustPoints.BackColor = bgc
        For n = 0 To 4
            t_CustPts(n).BackColor = bgc
        Next
        t_Intermediates.BackColor = bgc
        t_Losers.BackColor = bgc
        t_Pct.BackColor = bgc
        t_Round.BackColor = bgc
        t_Player.BackColor = bgc
        t_Order.BackColor = bgc
        t_Prior.BackColor = bgc
        t_QTs.BackColor = bgc
        For n = 0 To 5
            t_Qty(n).BackColor = bgc
        Next
        t_Stoppers.BackColor = bgc
        t_Sel.BackColor = bgc
        t_Sub.BackColor = bgc
        t_TestHand.BackColor = bgc
        t_Distrib_Pts.BackColor = bgc
        t_Winners.BackColor = bgc

        For n = 0 To 5
            t_Show(n).BackColor = bgc
        Next

        'For n = 0 To 5
        '    t_DupeEntry(n).BackColor = bgc
        'Next

        For n = 0 To 1
            cmb_Show(n).BackColor = bgc
        Next

        For n = 0 To 4
            cmb_Sort(n).BackColor = bgc
        Next

        For n = 0 To 3
            cmb_BackColors(n).BackColor = bgc
            cmb_ForeColors(n).BackColor = bgc
        Next

        cmb_Alert.BackColor = bgc
        cmb_Bid.BackColor = bgc
        For n = 0 To 1
            cmb_ConvName(n).BackColor = bgc
        Next
        cmb_Position.BackColor = bgc
        cmb_Shape.BackColor = bgc
        cmb_Vul.BackColor = bgc

        For n = 1 To 4
            t_Len(n).BackColor = bgc
        Next

        t_BestContract.BackColor = bgc
        int_BGInputs = cmb_BackColors(Index).ListIndex

    Case Is = 3
        ' Input Box area frames and check boxes: '
        For n = 0 To 7
            frame0(n).BackColor = bgc
        Next
        'f_Dupe.BackColor = bgc
        L_Date_Saved.BackColor = bgc
        ck_Artificial.BackColor = bgc
        ck_AutoFill.BackColor = bgc
        ck_Saved.BackColor = bgc
        ck_SaveNew.BackColor = bgc
        ck_Continuation.BackColor = bgc

        'f_Dupes.BackColor = bgc

        ' Frames in other areas: '
        f_Unhide.BackColor = bgc
        f_ShowSort.BackColor = bgc
        f_EnterHand.BackColor = bgc
        f_GridColors.BackColor = bgc
        f_Find2.BackColor = bgc

        Me.BackColor = bgc
        int_BGForm = cmb_BackColors(Index).ListIndex
    End Select

    ' Drop-down menu colors: '

    Dim mi As MENUINFO

    With mi
    '''''''
        .cbSize = Len(mi)

        .fMask = MIM_BACKGROUND
        .hbrBack = CreateSolidBrush(vbGreen)
        SetMenuInfo GetMenu(Me.hwnd), mi  'main menu bar

        .fMask = MIM_BACKGROUND Or MIM_APPLYTOSUBMENUS
        .hbrBack = CreateSolidBrush(vbCyan)
        SetMenuInfo GetSubMenu(GetMenu(Me.hwnd), 0), mi 'File menu (item 0)

        .hbrBack = CreateSolidBrush(vbGreen)
        SetMenuInfo GetSubMenu(GetMenu(Me.hwnd), 1), mi 'Edit menu (item 1)

        .hbrBack = CreateSolidBrush(vbYellow)
        SetMenuInfo GetSubMenu(GetMenu(Me.hwnd), 2), mi 'View menu (item 2)

        .hbrBack = CreateSolidBrush(vbGreen)
        SetMenuInfo GetSubMenu(GetMenu(Me.hwnd), 3), mi 'Help menu (item 3)
    ''''''''
    End With

    ck_Unopposed.BackColor = bgc

    DrawMenuBar Me.hwnd

End Sub


Private Sub cmb_ConvName_Click(Index As Integer)
    If cmb_ConvName(Index) <> "" Then
        LookUpConvDataAndDisplayIt Index
        If Index = 0 Then
            FindConvenNameFile cmb_ConvName(Index).Text, True, False, False
        End If
    End If
End Sub


Private Sub cmb_ConvName_GotFocus(Index As Integer)
    Dim i1 As Integer
    Dim i2 As Integer

    If bln_SkipIt Or f_Msg.Visible Then
        bln_SkipIt = False
        Exit Sub
    End If

    ' Update lists: '
    i1 = cmb_ConvName(0).ListIndex
    If i1 = -1 And cmb_ConvName(0).Text <> "" Then
        i1 = FindMatch(cmb_ConvName(0), cmb_ConvName(0).Text, True) ' True = exact match
    End If
    i2 = cmb_ConvName(1).ListIndex
    If i2 = -1 And cmb_ConvName(1).Text <> "" Then
        i2 = FindMatch(cmb_ConvName(1), cmb_ConvName(1).Text, True) ' True = exact match
    End If
    LoadBidNames
    cmb_ConvName(0).ListIndex = i1
    cmb_ConvName(1).ListIndex = i2

    str_OrigData = cmb_ConvName(Index).Text
    L_Status.Clear
    L_Status.AddItem " Name or description of bid or system.  E.g.: Stayman, Weak Jump Overcall"

    If Index = 0 Then
        cmb_ConvName(Index).SelStart = 0
        cmb_ConvName(Index).SelLength = 32768
    End If
End Sub


Private Sub cmb_ConvName_Validate(Index As Integer, Cancel As Boolean)
    Dim newData As String
    Dim i As Long
    Dim j As Integer
    Dim nam As String

    L_Status.Clear

If f_ConvNameChange.Visible Or f_Msg.Visible Then Exit Sub

'                                                                          '
 With cmb_ConvName(Index)
 ''''''''''''''''''''''''
    newData = Trim$(.Text)
    If newData = "" Then Exit Sub
    
    If str_OrigData <> newData Then
        ck_Saved.Value = 0

    Else
        If Trim$(.Text) <> "" Then
            LookUpConvDataAndDisplayIt Index
            If Index = 0 Then
                FindConvenNameFile cmb_ConvName(Index).Text, True, False, False
            End If
        End If
        Exit Sub
    End If

    nam = newData
    nam = Replace$(nam, "2/1", "2 Over 1")

    i = FindMatch(cmb_ConvName(Index), newData, True) ' True = exact match
    If i < 0 Then i = 0

    If .List(i) <> newData Then
        If LCase$(.List(i)) = LCase$(newData) Then
            ' Upper/Lower casing doesn't match; e.g.: test vs Test '
            str_Response = MyMsgBox("Capitalization of Convention Name entered" & vbCr & _
                                    "   (" & .Text & ")" & _
                                    "does not match that already in the list." & vbCr & _
                                    "   (" & .List(i) & ")" & _
                                    "Change it to match the existing entry?", _
                                    3, "YN", "BidBase")
            If str_Response = "yes" Then
                .Text = .List(i)
            Else
                With rs_BidNamesTbl
                    .Index = "Bid_Name"
                    .Seek "=", cmb_ConvName(Index).List(i)
                    If .NoMatch = False Then
                        .Edit
                        !Bid_Name = newData
                        .Update
                    End If
                End With
                ' Back to "With cmb_ConvName(Index)": '
            End If

        ElseIf str_OrigData = "" Then
            ' Was blank, new Name not found, so... '
            str_Response = MyMsgBox("Add " & newData & " to the list?", 1, "YN", "Convention Names")
            If str_Response = "yes" Then
                .AddItem newData
                .Text = newData
                With rs_BidNamesTbl
                    .AddNew
                    !Bid_Name = newData
                    !Selected = ""
                    !SubCat_Selected = ""
                    !Bid_Description = MyMsgBox("Enter description:", 1, "", newData)
                    .Update
                End With
                'Data2.Refresh
            Else
                .Text = ""
            End If

        Else
ConvNameChange:
            f_ConvNameChange.Visible = True
            f_ConvNameChange.ZOrder 0
            For j = 0 To 1
                L_NewName(j) = Trim$(cmb_ConvName(Index))
                L_OldName(j) = Trim$(str_OrigData)
            Next
            opt_ChConvName(1).Value = True
            b_ConvNameChange.SetFocus

            Do While f_ConvNameChange.Visible
                DoEvents
            Loop
        End If

    Else ' Matching name in list found, get ID# and Sel and Sub-Cat:
        Dim s As String
        Dim s2 As String
        Dim n As Long

        LookUpBidName n, s, s2, newData
        t_Sel = s
        t_Sub = s2

    End If

    If Index = 0 And .Text <> "" Then
        FindConvenNameFile cmb_ConvName(Index), True, False, False
    End If
'        '
 End With
'                                                                  '
End Sub


Private Sub g_OtherHands_Click()
    Dim n As Integer
    Dim r As Integer
    Dim c As Integer
    Dim id As Long
    
    With g_OtherHands
    '''''''''''''''''
    If .Row = 16 Then
        If .TextMatrix(int_CurHandDataRow, int_CurHandDataCol) <> t_OtherHands Then
            .TextMatrix(int_CurHandDataRow, int_CurHandDataCol) = t_OtherHands
        End If
        .Row = 0
        str_Response = MyMsgBox("Save " & .Text & "?", 1, "YN", "Hand Specs")
        If str_Response = "yes" Then
            c = .Col
            SaveHandSpecs c
        End If
        Exit Sub
    ElseIf .Row = 17 Then
        .Row = 1
        id = Val(.Text)
        If id > 0 Then
            str_Response = MyMsgBox("Delete entry ID# " & .Text & "?", 1, "YN", "Hand Specs")
            If str_Response = "yes" Then
                Stop
                Exit Sub
            End If
        End If
        
        str_Response = MyMsgBox("Clear column " & .Col & "?", 1, "YN", "Hand Specs")
        If str_Response = "yes" Then
            For r = 2 To 15
                .Row = 2
                .Text = ""
            Next
        End If
        Exit Sub
    End If
    
    f_Msg.Left = f_OtherPlayers.Left
    f_Msg.Top = frame0(6).Top - 1700
    f_Msg.Width = f_OtherPlayers.Width
    
    If t_OtherHands.Visible = False Then
        t_OtherHands.Visible = True
        t_OtherHands.ZOrder 0
    End If
    
    .TextMatrix(int_CurHandDataRow, int_CurHandDataCol) = t_OtherHands.Text
    
    int_CurHandDataRow = .Row
    int_CurHandDataCol = .Col
    
    t_OtherHands.Text = .TextMatrix(int_CurHandDataRow, int_CurHandDataCol)
    
    t_OtherHands.Top = .CellTop + 80
    
    n = 980
    
    If .Col = 1 Then
        t_OtherHands.Left = n + 40
    ElseIf .Col = 2 Then
        t_OtherHands.Left = n * 2 + 30
    ElseIf g_OtherHands.Col = 3 Then
        t_OtherHands.Left = n * 3 + 20
    ElseIf g_OtherHands.Col = 4 Then
        t_OtherHands.Left = n * 4
    End If
    ''''''''
    End With
    
    t_OtherHands.Refresh
    ' Make sure no button is set to default: '
    b_Show.Default = True ' random button '
    b_Show.Default = False ' Turn it off. '
    t_OtherHands.SetFocus
End Sub


Private Sub list_Refs_Click()
    t_Ref = list_Refs.List(list_Refs.ListIndex)
    'list_Refs.RemoveItem list_Refs.ListIndex
End Sub


Private Sub SaveHandSpecs(c As Integer)
    ' c is grid column '
    Dim sID As Long
    'Dim speks(13) As String ' 0-13 '
    Dim speks As String
    Dim i As Long
    Dim j As Long
    Dim id As Long
    
    Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
    Set rs_HandsData = db_BidBase.OpenRecordset("HandData")
    
    speks = Space$(92)  ' 6 spaces for each row 2 thru 13 and 10 for '
                        ' rows 14 and 15 (combo qty and combo HCP).  '
                        ' Each field includes 1 space for spacing so '
                        ' actual field length is 5 and 9 for last 2. '
    
    For i = 0 To 12 ' HCP thru CmbNum '
        Mid$(speks, i * 6 + 1) = g_OtherHands.TextMatrix(2 + i, c)
    Next
    
    ' CmbHCP: '
    Mid$(speks, 82) = g_OtherHands.TextMatrix(13, c)
    
    With rs_HandsData
    '''''''''''''''''
    sID = Val(g_OtherHands.TextMatrix(1, c))
    If sID = 0 Then
        ' Look for an existing entry '
        ' which matches these specs: '
        .Index = "Specs"
        .Seek "=", speks
        If .NoMatch Then
            .AddNew
            !specs = speks
            sID = -1 ' New Entry '
        Else
            ' Use found entry's ID:
            .Edit
            Stop
        End If
    Else
        .Index = "ID"
        .Seek "=", sID
        If .NoMatch = True Then
            str_Response = MyMsgBox("ID#" & sID & " is not in database.", 1, "ok", "")
            Exit Sub
        End If
        ' Because multiple bidding entries may use the same '
        ' hand data entry, we have to look to see if that   '
        ' is the case for this entry. If so, we have to save'
        ' the changed specs in a new entry. '
    
        .Edit
    End If
    
    .Update
    If sID = -1 Then
        .MoveFirst ' New entry is the first in the database.
    End If
    g_OtherHands.TextMatrix(1, c) = !id
    ''''''''
    End With
End Sub


Private Sub cmb_Show_DragDrop(Index As Integer, _
                              source As Control, _
                              x As Single, _
                              Y As Single)
    With source
        .Drag vbEndDrag
        .Move .Left + cmb_Show(Index).Left + x - sng_DragoffsetX, _
              .Top + cmb_Show(Index).Top + Y - sng_DragoffsetY
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


'*****************************************
' Procedure: cmb_Sort_Click              '
' Purpose:                               '
'     Combo boxes 0-4 are for specifying '
'     database sort order:               '
'*****************************************

Private Sub cmb_Sort_Click(Index As Integer)

    Dim j As String
    Dim sSearchText As String
    Dim KeyAscii As Integer

    If str_Job = "cmb_Sort" Then
        Exit Sub
    End If

    j = str_Job
    str_Job = "cmb_Sort"

    ' Based on user's specification of the first sort field, set '
    ' the subsequent sort combo lists to the appropriate fields: '
If 1 = 2 Then
    Select Case cmb_Sort(0)
    Case Is = "Bid_Name"
        sSearchText = "Bid_Order_Number" & Chr$(KeyAscii)
        cmb_Sort(1).ListIndex = _
            SendMessage(cmb_Sort(1).hwnd, &H14C, -1, ByVal sSearchText)

        sSearchText = "PB_Sort_Order" & Chr$(KeyAscii)
        cmb_Sort(2).ListIndex = _
            SendMessage(cmb_Sort(2).hwnd, &H14C, -1, ByVal sSearchText)

        sSearchText = "Bid" & Chr$(KeyAscii)
        cmb_Sort(3).ListIndex = _
            SendMessage(cmb_Sort(3).hwnd, &H14C, -1, ByVal sSearchText)

        cmb_Sort(4).ListIndex = -1
        cmb_Sort(4) = ""

    Case Is = "Prior_Bids"
        sSearchText = "Bid_Order_Number" & Chr$(KeyAscii)
        cmb_Sort(1).ListIndex = _
            SendMessage(cmb_Sort(1).hwnd, &H14C, -1, ByVal sSearchText)

        sSearchText = "Pct_Used" & Chr$(KeyAscii)
        cmb_Sort(2).ListIndex = _
            SendMessage(cmb_Sort(2).hwnd, &H14C, -1, ByVal sSearchText)

        sSearchText = "Bid" & Chr$(KeyAscii)
        cmb_Sort(3).ListIndex = _
            SendMessage(cmb_Sort(3).hwnd, &H14C, -1, ByVal sSearchText)

        sSearchText = "Bid_Name" & Chr$(KeyAscii)
        cmb_Sort(4).ListIndex = _
            SendMessage(cmb_Sort(4).hwnd, &H14C, -1, ByVal sSearchText)

    End Select
End If
    str_Job = j
End Sub


Private Sub cmb_Sort_DragDrop(Index As Integer, _
                              source As Control, _
                              x As Single, _
                              Y As Single)
    With source
        .Drag vbEndDrag
        .Move .Left + cmb_Sort(Index).Left + x - sng_DragoffsetX, _
              .Top + cmb_Sort(Index).Top + Y - sng_DragoffsetY
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub cmb_StrDiscl_GotFocus()
    str_OrigData = cmb_StrDiscl
    L_Status.Clear
    L_Status.AddItem " B = Bust (0-5 pts.)    W = Weak (6-10)      I  = Intermediate (10-12)    O = Opening strength (13-15)"
    L_Status.AddItem " P = Plus (15-17)       S = Strong (18-19)   V = Very strong (20-24)      X = eXtremely strong (25+)"
End Sub


Private Sub f_Compare_DragDrop(source As Control, _
                               x As Single, _
                               Y As Single)
    ' This is the colorful box which compares fields. '
    With source
        .Drag vbEndDrag
        .Move .Left + x - sng_DragoffsetX, _
              .Top + Y - sng_DragoffsetY
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub f_Compare_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                Y As Single)
    ' This is the colorful box which compares fields. '
    If Button = vbLeftButton Then
        If Not bln_Dragging Then
            bln_Dragging = True
            f_Compare.Drag vbBeginDrag
            ' When dropping a control, adjust for the    '
            ' offset of where you clicked on the control '
            ' from where the top-left corner is:         '
            sng_DragoffsetX = x
            sng_DragoffsetY = Y
            lng_X = x
            lng_Y = Y
        End If
    End If
End Sub


Private Sub f_ShowSort_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    sng_DragoffsetX = x
    sng_DragoffsetY = Y
    bln_Dragging = False
End Sub


Private Sub frm_Tutorial_DragDrop(source As Control, x As Single, Y As Single)
    With source
        .Drag vbEndDrag
        .Move frm_Tutorial.Left + x - sng_DragoffsetX, _
              frm_Tutorial.Top + Y - sng_DragoffsetY
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub frm_Tutorial_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        If Not bln_Dragging Then
            bln_Dragging = True
            frm_Tutorial.Drag vbBeginDrag
            ' When dropping a control, adjust for the    '
            ' offset of where you clicked on the control '
            ' from where the top-left corner is:         '
            sng_DragoffsetX = x
            sng_DragoffsetY = Y
        End If
    End If
End Sub


Private Sub frm_Tutorial_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    frm_Tutorial.Drag vbEndDrag
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub

Private Sub g_CompareGrid_Click()
    Dim r As Integer
    Dim c As Integer

    If g_CompareGrid.Row = 22 Then
        If g_CompareGrid.Col = 0 Then Exit Sub
        g_CompareGrid.Row = 0
        If g_CompareGrid.Col = 1 Then
            ' Edit first entry. '
            lng_ID = Val(Mid$(g_CompareGrid.Text, 2))
        Else
            ' Edit second entry '
            lng_ID = Val(Mid$(g_CompareGrid.Text, 2))
        End If
        DisplayBid2Edit 99
    End If
End Sub


Private Sub g_CompareGrid_DragDrop(source As Control, x As Single, Y As Single)
    With source
      If bln_Dragging Then
        .Drag vbEndDrag
        .Move f_CompareGrid.Left + x - sng_DragoffsetX + g_CompareGrid.Left, _
              f_CompareGrid.Top + Y - sng_DragoffsetY + g_CompareGrid.Top
      End If
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub L_BidFound_DblClick()
    Dim i As Long
    Dim p As String
    Dim b As String

    p = Trim$(t_FindPrior)
    b = Trim$(L_BidFound)

    If Right$(" " & p, 1) = "-" Then
        p = Left$(p, Len(p) - 1)
    End If

    If p = "" Then
        t_FindPrior = b & "-P-"

    ElseIf b = "" Or (Right$("  " & p, Len(b)) = b And Val(Right$("  " & p, Len(p) - 2)) > 0) Then
        ' no action '
        
    Else
        t_FindPrior = p & "-" & b & "-P-"
    End If

    If t_FindHand(1) <> "" Then
        If int_FindHandNum = 1 Then
            int_FindHandNum = 0
        Else
            int_FindHandNum = 1
        End If
        t_FindHand_GotFocus (int_FindHandNum)
    End If

    t_FindPrior.SetFocus
    t_FindPrior.SelStart = 999
End Sub

Private Sub L_CompareGrid_DragDrop(source As Control, x As Single, Y As Single)
    With source
      If bln_Dragging Then
        .Drag vbEndDrag
        .Move .Left + x - sng_DragoffsetX, _
              .Top + Y - sng_DragoffsetY
        .ZOrder 0
      End If
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub L_CompareGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then
        If Not bln_Dragging Then
            bln_Dragging = True
            f_CompareGrid.Drag vbBeginDrag
            ' When dropping a control, adjust for the    '
            ' offset of where you clicked on the control '
            ' from where the top-left corner is:         '
            sng_DragoffsetX = x
            sng_DragoffsetY = Y
        End If
    End If
End Sub


Private Sub Grid1_BeforeColUpdate(Index As Integer, _
                                  ByVal ColIndex As Integer, _
                                  OldValue As Variant, _
                                  Cancel As Integer)
    Dim t As String
    Dim num As Long
    Dim sel As String
    Dim sel2 As String
    Dim fID As Long

    Cancel = -1
Exit Sub

    ' Updating data in the grid causes a VB/Windows error. '

    If ColIndex = int_Prior_Bids_Col Then
        str_Response = MyMsgBox("Bidding entries may not be changed on the grids." & vbCr & _
                                "Load the entry into the Input Boxes to make changes.", 2, "ok", "")
        Cancel = True
    ElseIf ColIndex = int_Bid_Name_Col Then
        t = Trim$(Grid1(Index).Columns(int_Bid_Name_Col).Text)
        If t = "" Then Exit Sub
        LookUpBidName num, sel, sel2, t
        If num = 0 Then
            str_Response = MyMsgBox("No such bid/convention name." & vbCr & _
                                    "Use 'Bid/Conv. Name' box below" & vbCr & _
                                    "to add new bid/conventions names.", 3, "ok", "")
            Cancel = True
            Exit Sub
        End If

        With rs_BidTable
        ''''''''''''''''
            .Index = "PrimaryKey"
            fID = Val(Grid1(Index).Columns(int_LastBidCol))
            .Seek "=", fID
            If .NoMatch = True Then
                ' Shouldn't be possible. '
                str_Response = MyMsgBox("Entry's ID# not in database.", 1, "ok", "")
                Exit Sub
            End If
            !ConvID = num
            .Update
        End With
        ''''''''
        Data1(0).Refresh
        Data1(1).Refresh
    End If
End Sub


Private Sub Grid1_KeyDown(Index As Integer, _
                          KeyCode As Integer, _
                          Shift As Integer)
    If Index = 1 And _
       (KeyCode = 38 And _
        Shift = 2) _
    Then
        ' Ctrl-Up swaps grids' contents:
        Grid1(1).Col = int_Prior_Bids_Col
        t_Show(0) = Grid1(1)
        t_Show(1) = Grid1(1)
        Grid1(0).Col = int_Prior_Bids_Col
        t_Show(2) = Grid1(0)
        t_Show(3) = Grid1(0)
        b_Show_Click
    End If
End Sub


Private Sub Grid1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If int_MouseOverGrid <> Index Then
        int_MouseOverGrid = Index
    End If
End Sub


Private Sub L_CompareBackground_DragDrop(source As Control, _
                                         x As Single, _
                                         Y As Single)
    With source
        .Drag vbEndDrag
        .Move .Left + L_CompareBackground.Left + x - sng_DragoffsetX, _
              .Top + L_CompareBackground.Top + Y - sng_DragoffsetY
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub L_CompareBackground_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                Y As Single)
    ' This is the multi-color frame, not the grid. '
    If Button = vbLeftButton Then
        If Not bln_Dragging Then
            bln_Dragging = True
            f_Compare.Drag vbBeginDrag ' User clicked on the label, but it's the form being moved. '
            ' When dropping a control, adjust for the    '
            ' offset of where you clicked on the control '
            ' from where the top-left corner is:         '
            sng_DragoffsetX = L_CompareBackground.Left + x
            sng_DragoffsetY = L_CompareBackground.Top + Y
            lng_X = x
            lng_Y = Y
        End If
    End If
End Sub


Private Sub L_CompareField_DragDrop(source As Control, _
                                    x As Single, _
                                    Y As Single)
    With source
        .Drag vbEndDrag
        .Move .Left + L_CompareField.Left + x - sng_DragoffsetX, _
              .Top + L_CompareField.Top + Y - sng_DragoffsetY
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub L_CompareField_MouseDown(Button As Integer, _
                                     Shift As Integer, _
                                     x As Single, _
                                     Y As Single)
    If Button = vbLeftButton Then
        If Not bln_Dragging Then
            bln_Dragging = True
            f_Compare.Drag vbBeginDrag
            ' When dropping a control, adjust for the    '
            ' offset of where you clicked on the control '
            ' from where the top-left corner is:         '
            sng_DragoffsetX = L_CompareField.Left + x
            sng_DragoffsetY = L_CompareField.Top + Y
            lng_X = x
            lng_Y = Y
        End If
    End If
End Sub


Private Sub f_EnterHand_DragDrop(source As Control, _
                                 x As Single, _
                                 Y As Single)
    With source
        .Drag vbEndDrag
        .Move f_EnterHand.Left + x - sng_DragoffsetX, _
              f_EnterHand.Top + Y - sng_DragoffsetY
        Debug.Print f_EnterHand.Left; ", "; f_EnterHand.Top
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub f_EnterHand_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  Y As Single)
    If Button = vbLeftButton Then
       ' If Not bln_Dragging Then
            bln_Dragging = True
            f_EnterHand.Drag vbBeginDrag
            ' When dropping a control, adjust for the    '
            ' offset of where you clicked on the control '
            ' from where the top-left corner is:         '
            sng_DragoffsetX = x
            sng_DragoffsetY = Y
            lng_X = x
            lng_Y = Y
        End If
    'End If
End Sub


Private Sub f_ShowSort_DragDrop(source As Control, _
                                x As Single, Y As Single)
    With source
        .Drag vbEndDrag
        .Move f_ShowSort.Left + x - sng_DragoffsetX, f_ShowSort.Top + Y - sng_DragoffsetY
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub f_ShowSort_MouseDown(Button As Integer, _
                                 Shift As Integer, _
                                 x As Single, Y As Single)
    If Button = vbLeftButton Then
        If Not bln_Dragging Then
            bln_Dragging = True
            f_ShowSort.Drag vbBeginDrag
            ' When dropping a control, adjust for the    '
            ' offset of where you clicked on the control '
            ' from where the top-left corner is:         '
            sng_DragoffsetX = x
            sng_DragoffsetY = Y
        End If
    End If
End Sub


Private Sub Form_DragDrop(source As Control, _
                          x As Single, Y As Single)
    ' This routine is for when the mouse   '
    ' button is released while the pointer '
    ' is over the form. '
    With source
        .Drag vbEndDrag
        .Move .Left + x - sng_DragoffsetX, _
              .Top + Y - sng_DragoffsetY
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub Frame0_DragDrop(Index As Integer, _
                            source As Control, _
                            x As Single, Y As Single)
    With source
        .Drag vbEndDrag
        .Move frame0(Index).Left + x - sng_DragoffsetX, _
              frame0(Index).Top + Y - sng_DragoffsetY
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub Grid1_DblClick(Index As Integer)
    ' Process double-click on grids: '
    Dim f1 As String
    Dim f As String
    Dim i As Integer
    Dim cnv As String ' Convention name (before the ":" if any). '
    Dim subCnv As String ' Subconvention name after the ":" if any. '
    Dim v As String
    Dim v2 As String

    If Grid1(Index).Row = -1 Then Exit Sub

   '                                                                '
    With Grid1(Index)
   '                 '
    f1 = Trim$(.Columns(int_Bid_Name_Col).Text)

    If .Col = int_Expl_Col Then
        'If Explanation column references a note file, show it:
        LookForNoteFile .Columns(int_Expl_Col).Text, _
                        .Columns(int_Prior_Bids_Col).Text, _
                        .Columns(int_Bid_Col).Text, _
                        .Columns(int_LastBidCol).Text, _
                        False, _
                        False, _
                        True

    ElseIf .Col = int_Bid_Name_Col And _
           .Columns(int_Bid_Name_Col).Text <> "" _
    Then
        FindConvenNameFile f1, False, True, False

    ElseIf .Col = int_Source_Col Then
        ShowSourceDeal .Text

    ElseIf (.Col = 0 Or .Col = 1) And f1 <> "" Then
        i = InStr(f1, ":")
        If i > 0 Then
            If .Col = 0 Then
                cnv = Left$(f1, i - 1) ' Just SomeCnv from "SomeCnv: SomeSub". '
            Else
                cnv = f1
            End If
        Else
            cnv = f1 ' No subcat, just something like "SomeCnv". '
        End If

        ' Whether cell 0 or 1 was clicked,     '
        ' .Text will be the value of the cell. '
        If .Text = "0" Then
            v = ""
            v2 = " activated"
        Else
            v = "0"
            v2 = " deactivated"
        End If

        Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
        Set rs_BidNamesTbl = db_BidBase.OpenRecordset(str_BidNamesTbl)
        rs_BidNamesTbl.Index = "Bid_Name"
        rs_BidNamesTbl.Seek ">=", cnv
        If InStr(rs_BidNamesTbl!Bid_Name, cnv) = "" Then
            str_Response = MyMsgBox("Convention name '" & cnv & "' not found.", 1, "ok", "")
            Exit Sub
        End If

        If .Col = 1 Then
            rs_BidNamesTbl.Edit
            rs_BidNamesTbl!SubCat_Selected = v
            rs_BidNamesTbl.Update
        Else
            Do While Left$(rs_BidNamesTbl!Bid_Name, Len(cnv)) = cnv
                rs_BidNamesTbl.Edit
                rs_BidNamesTbl!Selected = v
                rs_BidNamesTbl.Update
                'If rs_BidNamesTbl.EOF Then Exit Do
                rs_BidNamesTbl.MoveNext
                If rs_BidNamesTbl.EOF Then Exit Do
            Loop
        End If

        ' Try to get the grid to show the changes: '
        'rs_BidNamesTbl.Update
        i = .Col
        Data1(0).Refresh
        Grid1(0).Refresh
        Data1(1).Refresh
        Grid1(1).Refresh
        Timer3.Enabled = True
        Do While Timer3.Enabled = True
            DoEvents
        Loop
        .Col = i
        If .Col = 0 Then
            str_Response = MyMsgBox("All entries using the convention name of" & vbCr & _
                                    "     '" & cnv & "'" & vbCr & _
                                    "have been" & v2 & ".", 3, "ok", "")
        End If

    ElseIf .Col = 2 Then
        int_GridNum = Index
        str_Job = "Change activation"
        mEdit_Click Index
        str_Response = MyMsgBox("Pct Used has several different options." & vbCr & _
                                "The entry has been loaded into the Input Boxes." & vbCr & _
                                "Enter the desired option and Save Entry.", 3, "ok", "")
        If t_Pct = "" Then
            t_Pct = "0"
        ElseIf t_Pct = "0" Then
            t_Pct = ""
        End If
        t_Pct.SetFocus
    End If
   '        '
    End With
   '                                                                '
End Sub


Private Sub Grid1_DragDrop(Index As Integer, _
                           source As Control, _
                           x As Single, Y As Single)
    ' This isn't for moving Grids. It is for moving '
    ' controls where the mouse button was released  '
    ' over a Grid. '
    With source
        .Drag vbEndDrag
        .Move Grid1(Index).Left + x - sng_DragoffsetX, Grid1(Index).Top + Y - sng_DragoffsetY
        Debug.Print Grid1(Index).Left + x - sng_DragoffsetX, Grid1(Index).Top + Y - sng_DragoffsetY
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub Grid1_DragOver(Index As Integer, source As Control, x As Single, Y As Single, State As Integer)
    'If Index = 1 Then Stop
End Sub


Private Sub Grid1_KeyUp(Index As Integer, _
                        KeyCode As Integer, _
                        Shift As Integer)
    If Index = 0 And mDupesGrid(0).Checked Then
        If (KeyCode = 38 Or KeyCode = 40) And Shift = 0 Then
            ' Up or Down arrow pressed '
            ' so do get the dupes:     '
            Timer1.Interval = 1
        End If
    End If
End Sub

Private Sub Grid1_MouseDown(Index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            x As Single, _
                            Y As Single)
    ' This routine pops up either the '
    ' Edit or View menu, depending on '
    ' the location of the mouse.      '

    If Button = 2 Then ' right button
        If Y > 225 Then
            PopupMenu mEd ' Edit menu
        Else
            PopupMenu mV ' View menu
        End If
    Else
        Grid1(Index).SetFocus
    End If
End Sub


Private Sub HScroll1_Change()
    ' On screens not wide enough to see all '
    ' the input box, this scroll bar allows '
    ' scrolling the input box section L-R.  '

    frame0(6).Left = HScroll1.Value
    If frame0(6).Left < 0 Then
        HScroll1.Left = Abs(frame0(6).Left) + 100
    Else
        HScroll1.Left = 0
    End If
End Sub


Private Sub L_CompareGridExit_Click()
    f_CompareGrid.Visible = False
End Sub


Private Sub L_CompareSource_DragDrop(Index As Integer, _
                                     source As Control, _
                                     x As Single, Y As Single)
    With source
        .Drag vbEndDrag
        .Move .Left + L_CompareSource(Index).Left + x - sng_DragoffsetX, _
              .Top + L_CompareSource(Index).Top + Y - sng_DragoffsetY
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub L_FindBid_DragDrop(Index As Integer, _
                               source As Control, _
                               x As Single, Y As Single)
    With source
        .Drag vbEndDrag
        .Move .Left + L_FindBid(Index).Left + x - sng_DragoffsetX, _
              .Top + L_FindBid(Index).Top + Y - sng_DragoffsetY
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub L_Status_Click()
    L_Status.ListIndex = -1
End Sub


Private Sub L_TestingMsg_DblClick()
    bln_Abort = True
End Sub


Private Sub L_Tutorial_DragDrop(source As Control, x As Single, Y As Single)
    With source
        .Drag vbEndDrag
        .Move frm_Tutorial.Left + L_Tutorial.Left + x - sng_DragoffsetX, _
              frm_Tutorial.Top + L_Tutorial.Top + Y - sng_DragoffsetY
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub L_Tutorial_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If bln_Dragging Then
        bln_Dragging = False
        L_Tutorial.Drag vbEndDrag
    End If
End Sub


Private Sub mBackup_Click()
    Dim x As Integer

    If frm_Tutorial.Visible Then
        If b_Continue.Visible = False Then
            frm_Tutorial.Caption = "Back Up Database"
            L_Tutorial = "This command lets you back up the bidding database. " & _
                         "Backing up should be done regularly if you " & _
                         "frequently add entries or make changes to it. "
            b_Continue.Visible = True
        End If
        Exit Sub
    End If

    If Dir$(str_BBpath & "\Backup\*.*") = "" Then
        MkDir str_BBpath & "\Backup"
    End If

    CmDia.DialogTitle = "Enter backup file name:"
    CmDia.Filter = "*.accdb"
    CmDia.InitDir = str_BBpath & "\Backup\"
    CmDia.FileName = "bids-" & _
                     Right$(Date$, 4) & "-" & _
                     Left$(Date$, 5) & _
                     ".accdb"
    CmDia.ShowSave
    If CmDia.FileTitle = "" Then Exit Sub
    If Dir$(CmDia.FileName) <> "" Then
        str_Response = MyMsgBox("Overwrite existing file with that name?", 1, "YNC", "BidBase")
        Select Case str_Response
        Case "cancel"
            Exit Sub
        Case ""
            ' No ? '
            mBackup_Click
            Exit Sub
        End Select
    End If

    FileCopy str_DBPathAndName, _
             str_BBpath & "\backup\" & CmDia.FileTitle

End Sub


Private Sub mBiddingHelp_Click()
    str_NoteFile(0) = "Index.htm"
    ShowDocs
End Sub


Private Sub mBidSim_Click()
    ' BidSim is not active at this time. '
    ' The menu item calling this should not be visible. '
    Exit Sub

    On Local Error GoTo LoadBidSim
    ' If BidSim is not already loaded,   '
    ' AppActivate will generate an error '
    ' which will Shell to the program.   '
    ' This keeps more than one BidSim    '
    ' program from opening.              '
    AppActivate "CardShark BidBase Bidding Simulator"

ExitRunBidSim:
    On Local Error GoTo 0
Exit Sub

LoadBidSim:
    Dim ret As Long
    If Dir$(str_BBpath & "\BidBase Bidding Simulator.exe") = "" Then
        str_Response = MyMsgBox("Bidding Simulator is not present.", 1, "ok", "")
    Else
        ret = Shell(str_BBpath & "\BidBase Bidding Simulator.exe", 1)
    End If
    Resume ExitRunBidSim
End Sub


Private Sub mChangeBrowser_Click()
    On Error GoTo errhandler
    With CmDia
    ''''''''''
    .CancelError = True
    ' Set flags
    .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
    ' Set filters
    '.Filter = "All Files (*.*)|*.*|RTF (*.rtf)|*.rtf|Text Files (*.txt)|*.txt"

    ' Display the Save dialog box
    .FileName = ""
    .ShowOpen
    str_Browser = .FileName

    End With
    ''''''''
    Exit Sub

errhandler:
    Select Case Err
    Case 32755 '  Dialog Cancelled
        str_Response = MyMsgBox("You cancelled the dialog box", 1, "ok", "")
    Case Else
        str_Response = MyMsgBox("Unexpected error Err " & Err & " : " & Error, 1, "ok", "")
    End Select
End Sub


'******************************************
' Procedure: mCompare_Click               '
' Purpose:                                '
    ' Compare the entries in two grids.   '
    ' Click on an entry in each grid.     '
    ' Prior Bids should be similar, such  '
    ' as 1C-1H & 1D-1H, or 1C-1H & 1C-1S. '
'******************************************

Private Sub mCompare_Click()
    Dim gn As Integer
    Dim grid2Compare2Boxes As Integer
    Dim c As Integer
    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    Dim r As Integer
    Dim t As String
    Dim response
    Dim xFieldUsed(1) As Boolean
    Dim nonSuitsOnly As Boolean
    Dim bid_Grid(1) As String
    Dim ordNum_Grid(1) As String
    Dim prBids_Grid(1) As String
    Dim rightChar_ordNum_Grid(1) As String

    '   See if this is a Tutorial request:      '
    If frm_Tutorial.Visible Then
        If b_Continue.Visible = False Then
            Run_CompareEntries_Tutorial
        End If
        Exit Sub
    End If

    If mCompare.Checked Then
        mCompare.Checked = False
        f_Compare.Visible = False
        bln_Abort = True
        Exit Sub
    End If

    '   Screen out invalid comparison groups:   '

    If Not bln_TwoGrids Or _
       Grid1(0).Columns(int_Bid_Col) = "" Or _
       Grid1(1).Columns(int_Bid_Col) = "" _
    Then
        str_Response = MyMsgBox("This feature is for comparing" & vbCr & _
                                "entries in different grids." & vbCr & _
                                "Select an entry in each grid.", 3, "ok", "")
        Exit Sub
    End If

    If t_Show(0) <> t_Show(1) Or _
       t_Show(2) <> t_Show(3) _
    Then
        f_ShowSort.Visible = True
        str_Response = MyMsgBox("Each grid must contain only" & vbCr & _
                                "one set of Prior Bids.", 2, "ok", "")
        b_Show.Default = True
        Exit Sub
    End If

    prBids_Grid(0) = Grid1(0).Columns(int_Prior_Bids_Col).Text
    prBids_Grid(1) = Grid1(1).Columns(int_Prior_Bids_Col).Text

    mCompare.Checked = True
    
    Dim b As String
    Dim id1 As Long
    Dim id2 As Long
    Stop
    b = Grid1(0).Columns(int_Bid_Col).Text
    id1 = Val(Grid1(0).Columns(int_LastBidCol).Text)
    id2 = Val(Grid1(1).Columns(int_LastBidCol).Text)
    ShowGetBidError b, id1, id2

'Exit Sub

CompareGridRows:
    Do
        ' Make sure grids are shifted all the way '
        ' over by making column 0 the first col.  '
        Grid1(0).Col = 0
        Grid1(1).Col = 0
        ordNum_Grid(0) = Grid1(0).Columns(int_OrdNum_Col)
        ordNum_Grid(1) = Grid1(1).Columns(int_OrdNum_Col)
        If ordNum_Grid(0) = "ZZZZZZ" And _
           ordNum_Grid(1) = "ZZZZZZ" _
        Then
            GoTo ExitCompare
        End If
        nonSuitsOnly = False

        ' Look for entries with Bid Order Numbers that match:
        Do While Grid1(0).Columns(int_OrdNum_Col) _
              <> Grid1(1).Columns(int_OrdNum_Col)
            ' Order Numbers do not match. If they are the '
            ' same except for the last letter, their suit '
            ' specs will usually be non-comparable, but   '
            ' the non-suit specs, such as total HCPs, will'
            ' be comparable, so we set a flag accordingly.'

            ' Working variables:
            ordNum_Grid(0) = Grid1(0).Columns(int_OrdNum_Col)
            ordNum_Grid(1) = Grid1(1).Columns(int_OrdNum_Col)
            bid_Grid(0) = Grid1(0).Columns(int_Bid_Col)
            bid_Grid(1) = Grid1(1).Columns(int_Bid_Col)

            rightChar_ordNum_Grid(0) = Right$(ordNum_Grid(0), 1)
            rightChar_ordNum_Grid(1) = Right$(ordNum_Grid(1), 1)

            If (rightChar_ordNum_Grid(0) > "9" Or _
                rightChar_ordNum_Grid(1) > "9") And _
               Left$(ordNum_Grid(0), 5) = Left$(ordNum_Grid(1), 5) _
            Then
                ' If the two numbers match except for a letter at the end '
                ' of one of them (eg: 41530D - 415300), set a flag:       '
                nonSuitsOnly = True
                int_GridNum = 0
                ColorGrids 0
                grid2Compare2Boxes = 1
                Exit Do

            ElseIf ordNum_Grid(0) < ordNum_Grid(1) Then
                ' Grid1(0) has an entry which Grid1(1) doesn't.           '
                ' eg: If 240000 vs 300000, then grid1(1) must not have    '
                '    an entry for 240000. Now we will check to see if the '
                '    missing entry should be offered to the user to add.  '

                If rightChar_ordNum_Grid(0) > "9" Then
                    ' If the missing entry ends with a letter, then we    '
                    ' just go to the next row for that grid and try again.'
                    Grid1(0).Row = Grid1(0).Row + 1
                    GoTo CompareGridRows
                End If

                ' If Grid1(0) has an entry #400000 with Prior Bids of 1H  '
                ' and a bid of 1S and Grid1(1) has PBs of 1S and doesn't  '
                ' have an entry #400000, we check to see if it could have '
                ' the same bid (1S). Since the bid is not higher than the '
                ' last PB for Grid1(1), it cannot have an entry #400000,  '
                ' so we just skip to the next row for Grid1(0).           '

                Dim lsc As String
                Dim numBid As Integer
                ' First, get the last suit changed:
                n = Len(prBids_Grid(0)) - 1
                Do While n > 1
                    If Mid$(prBids_Grid(0), n - 1, 1) = "-" Then
                        Exit Do
                    End If
                    n = n - 1
                    DoEvents
                Loop
                numBid = Val(Mid$(prBids_Grid(0), n, 1))
                If numBid > 0 And _
                   Val(bid_Grid(1)) > 0 And _
                   Len(bid_Grid(0)) > 1 _
                Then
                    lsc = Mid$(prBids_Grid(1), n + 1, 1)

                    If Right$(bid_Grid(0), 1) > lsc Or Right$(bid_Grid(0), 1) = "N" Or _
                       Val(bid_Grid(0)) > numBid _
                    Then
                        ' ok '
                    Else
                        ' eg: Either the suit in bid_Grid(1) has to be of   '
                        ' a higher rank than the last suit bid (lsc) or the '
                        ' level bid for bid_Grid(1) must be higher than the '
                        ' level bid for lsc.   So if Grid1(0)'s PB is "1C"  '
                        ' and Grid1(1)'s PB is "1D", with these entries:    '
                        ' 410000 pb:1C  bid:1D  and  (no entry 410000)      '
                        ' 411000 pb:1C  bid:1H  and  411000 pb:1D  bid:1H   '
                        ' Grid(1) doesn't have entry 410000, but if it did  '
                        ' have that entry with a bid of 1D, the bidding     '
                        ' would go 1D-1D, which is obviously illegal.       '
                        ' So when the suit bid is not higher ranked than the'
                        ' last prior bid suit and the level is not higher,  '
                        ' then we skip the entry and advance to the next    '
                        ' entry for Grid1(0), which would be 411000.        '
                        Grid1(0).Row = Grid1(0).Row + 1
                        GoTo CompareGridRows
                    End If
                End If

                ' Show the mocked-up missing entry '
                ' for grid1(1) in the Input Boxes: '
                r = Grid1(1).Row
                int_GridNum = 0
                ColorGrids 0
                grid2Compare2Boxes = 1
                If mFourGrids.Checked Then
                    mFourGrids_Click
                End If
                If mInputBoxes.Checked = False Then
                    mInputBoxes_Click
                End If
                mEdit_Click 2
                Grid1(1).Row = r

                response = MyMsgBox("Bid Order Number " & _
                            Grid1(0).Columns(int_OrdNum_Col) & _
                            " entry is in" & vbCr & _
                            "Grid #1, but not in Grid #2." & vbCr & _
                            "A matching entry for Grid #2 is in the" & vbCr & _
                            "Input Boxes. Do you want to save it?", _
                            4, "YNC", "BidBase")

                Select Case response
                Case Is = "cancel"
                    mCompare.Checked = False
                    Exit Sub

                Case Is = "yes"
                    int_GridNum = 1
                    ColorGrids 1
                    mSave_Click
                    If Grid1(1).Row > 0 Then
                        Grid1(1).Row = Grid1(1).Row - 1
                    Else
                        Grid1(1).Scroll 0, -1
                    End If
                    GoTo CompareGridRows

                Case Is = ""
                    Grid1(0).Row = Grid1(0).Row + 1
                    GoTo CompareGridRows
                End Select

                ' Increment Grid1(0) only since Grid1(1) '
                ' is already on a larger entry number:   '
                ' [Eg: if Grid1(0)=100000 and (1)=200000,'
                ' then (1) was missing the 100000 entry, '
                ' which we asked about above. Now we must'
                ' increment (0) to get it to 200000 like '
                ' (1) is pointing to.]                   '
                Grid1(0).Row = Grid1(0).Row + 1

            ElseIf ordNum_Grid(0) > ordNum_Grid(1) Then
                ' Grid1(1) has an entry Grid1(0) doesn't.  '
                ' If last character of the lower Bid Order '
                ' Number is a letter, then skip comparison.'
                If rightChar_ordNum_Grid(1) > "9" Then
                    Grid1(1).Row = Grid1(1).Row + 1
                    GoTo CompareGridRows
                End If

            End If

            r = Grid1(0).Row
            int_GridNum = 1
            ColorGrids 1
            grid2Compare2Boxes = 0
            If mFourGrids.Checked Then
                mFourGrids_Click
            End If
            If mInputBoxes.Checked = False Then
                mInputBoxes_Click
            End If
            mEdit_Click 2
            Grid1(0).Row = r

            response = MyMsgBox("Bid Order Number " & _
                        Grid1(1).Columns(int_OrdNum_Col) & _
                        " entry is in" & _
                        " Grid #2 but not in Grid #1." & vbCr & _
                        "A matching entry for Grid #1 is in the Input Boxes." & vbCr & _
                        "Do you want to save it?", _
                        3, "YNC", "BidBase")

            Select Case response
            Case Is = "cancel"
                mCompare.Checked = False
                Exit Sub

            Case Is = "yes"
                int_GridNum = 0
                ColorGrids 0
                mSave_Click
                If Grid1(0).Row > 0 Then
                    Grid1(0).Row = Grid1(0).Row - 1
                Else
                    Grid1(0).Scroll 0, -1
                End If
                GoTo CompareGridRows

            Case Is = ""
                Grid1(1).Row = Grid1(1).Row + 1
                GoTo CompareGridRows
            End Select

            ' Increment Grid1(1) only since Grid1(0) '
            ' is already on a larger entry number:   '
            ' [Eg: if Grid1(1)=100000 and (0)=200000,'
            ' then (0) was missing the 100000 entry, '
            ' which we asked about above. Now we must'
            ' increment (0) to get it to 200000 like '
            ' (1) is pointing to.]                   '
            Grid1(1).Row = Grid1(1).Row + 1
            DoEvents
        Loop


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '                 Bid Order Numbers match.                   '
    '         Now compare the columns in the two rows.           '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ' Show the mocked-up missing entry for grid1(1) '
        ' in the Input Boxes unless grid1(1) uses the "X"
        ' field and grid1(0) doesn't, then mock up the  '
        ' entry for grid1(0) for comparing the two grids.
        '                                               '
        ' The reason for this is that while BidBase does'
        ' a good job of mocking up non-X-field entries  '
        ' from an X-field entry, it's not practical to  '
        ' try to do the reverse because the X-field can '
        ' be used in different ways.                    '

        For i = 0 To 1
            xFieldUsed(i) = Grid1(i).Columns(int_xNum_Col) <> "" Or _
                            Grid1(i).Columns(int_xNum_Col + 1) <> ""
        Next
        If xFieldUsed(1) And Not xFieldUsed(0) Then
            int_GridNum = 1
            ColorGrids 1
            grid2Compare2Boxes = 0
        Else
            int_GridNum = 0
            ColorGrids 0
            grid2Compare2Boxes = 1
        End If

        mEdit(2).Checked = False
        mEdit_Click 2
        t = ""

        For c = 0 To 4
            Do While Trim$(Grid1(0).Columns(c)) <> _
                     Trim$(Grid1(1).Columns(c))
                GoSub NoMatch
                If response = "0" Then Exit Do
                If response = "x" Then GoTo NextRow
                DoEvents
            Loop
        Next

        ' "Convention Name" through "Intermediates" should '
        ' be the same. Only columns for suits specs could  '
        ' "legally" be different, so the suit specs for the'
        ' target grid are compared to the Input Boxes.     '

        For c = int_Bid_Name_Col To int_xNum_Col - 1
            Do While Grid1(0).Columns(c) _
                  <> Grid1(1).Columns(c)
                '                      '
                GoSub NoMatch
                If response = "0" Then Exit Do
                If response = "x" Then GoTo NextRow
                DoEvents
            Loop
        Next

        If Not nonSuitsOnly Then
          For i = int_xNum_Col To int_Combo_Col Step 2
            j = (i - int_xNum_Col) / 2

            Do While Grid1(grid2Compare2Boxes).Columns(i) _
                  <> t_Qty(j)
                '           '
                c = i
                t = t_Qty(j) & "."
                GoSub NoMatch
                If response = "0" Then Exit Do
                If response = "x" Then GoTo NextRow
                DoEvents
            Loop

            Do While Grid1(grid2Compare2Boxes).Columns(i + 1) _
                  <> t_HCP(j)
                '           '
                c = i + 1
                t = t_HCP(j) & "."
                GoSub NoMatch
                If response = "0" Then Exit Do
                If response = "x" Then GoTo NextRow
                DoEvents
            Loop
          Next
        End If

        For c = int_Vul_Col To int_Vul_Col + 4 ' Vul thru Alert
            Do While Grid1(0).Columns(c) _
                  <> Grid1(1).Columns(c)

                GoSub NoMatch
                If response = "0" Then Exit Do
                If response = "x" Then GoTo NextRow
                DoEvents
            Loop
        Next

        If Not nonSuitsOnly Then
          Do While Grid1(grid2Compare2Boxes).Columns(int_Expl_Col) _
                <> t_Disclosure

            c = int_Expl_Col
            t = t_Disclosure & "."
            GoSub NoMatch
            If response = "0" Then Exit Do
            If response = "x" Then GoTo NextRow
            DoEvents
          Loop

          Do While Grid1(grid2Compare2Boxes).Columns(int_TestHand_Col) _
                <> t_TestHand

            c = int_TestHand_Col
            t = t_TestHand & "."
            GoSub NoMatch
            If response = "0" Then Exit Do
            If response = "x" Then GoTo NextRow
            DoEvents
          Loop
        End If

        Do While Grid1(0).Columns(43) _
              <> Grid1(1).Columns(int_TestHand_Col + 1)
            ' Strength discl. '
            c = int_TestHand_Col
            GoSub NoMatch
            If response = "0" Then Exit Do
            If response = "x" Then GoTo NextRow
            DoEvents
        Loop

        If Not nonSuitsOnly Then
          str_Response = MyMsgBox("nonSuitsOnly is not working.", 1, "ok", "")
          Exit Sub
          'Do While Grid1(grid2Compare2Boxes).Columns(int_TestHand_Col + 2) <> t_Distrib
            ' Distribution '
            c = int_TestHand_Col + 2
            str_Response = MyMsgBox("t_Distrib not present.", 1, "ok", "")
            't = t_Distrib & "."
            GoSub NoMatch
            If response = "0" Then Exit Do
            If response = "x" Then GoTo NextRow
            DoEvents
          'Loop
        End If

        Do While Grid1(0).Columns(45) <> _
                 Grid1(1).Columns(int_TestHand_Col + 3)
            ' Asks pard.
            c = int_TestHand_Col + 3
            GoSub NoMatch
            If response = "0" Then Exit Do
            If response = "x" Then GoTo NextRow
            DoEvents
        Loop

        Do While Grid1(0).Columns(49) <> _
                 Grid1(1).Columns(int_Source_Col)
            ' Source '
            c = int_Source_Col
            GoSub NoMatch
            If response = "0" Then Exit Do
            If response = "x" Then GoTo NextRow
            DoEvents
        Loop

NextRow:
        If str_Job = "Compare Dupes" Then
            Exit Do
        End If

        'On Local Error Resume Next
        Grid1(1).Row = Grid1(1).Row + 1
        If mDupesGrid(0).Checked = False Then
            ' When Dupes for an entry in Grid1(0) are displayed '
            ' in Grid1(1), then we want to keep comparing the   '
            ' same entry in (0) to all the ones in (1). Other-  '
            ' wise, we want to advance to the next row in both. '
            Grid1(0).Row = Grid1(0).Row + 1
        Else
            ' When dupes are being compared, we have to restart '
            ' this whole routine for each line because the suits'
            ' for the Prior Bids can change for each line in the'
            ' Dupes grid, so we have to re-examine those suits, '
            ' which is done at the start of this routine before '
            ' we loop through the columns to compare them.      '
            If Grid1(1).Columns(int_Bid_Col) <> "" Then
                str_DupePrior = Grid1(1).Columns(int_Prior_Bids_Col)
                DoEvents
                mCompare_Click
                ' This causes a small stack increment because '
                ' the Exit Sub (below) for each time through  '
                ' is not executed until the whole Dupes grid  '
                ' is done (or the user aborts), but since     '
                ' there are at most 5 entries, it's okay.     '
                ' An alternative is to use Timer1 to restart  '
                ' this routine, which gives the Exit Sub time '
                ' to execute.                                 '
            Else
                GoTo ExitCompare
            End If
            Exit Sub
        End If
        DoEvents
    Loop While Grid1(0).Columns(int_Bid_Col) <> "" Or _
               Grid1(1).Columns(int_Bid_Col) <> ""

ExitCompare:
    mCompare.Checked = False
    Grid1(0).Col = 0
    Grid1(1).Col = 0
    If gn = int_GridNum And o_CompareOptions(4) = False Then
        str_Response = MyMsgBox("Comparison done.", 1, "ok", "")
    End If
Exit Sub


' Get user input for fields that don't match:                                 '

NoMatch:
'''''''
    Dim g1 As Integer
    Dim g2 As Integer
    Dim d1 As String
    Dim d2 As String
    Dim ib As String

    ' Initially, Grid1(0) is the "master" grid,  '
    ' meaning that the Input Boxes are filled in '
    ' from it to compare Grid1(1)'s data against.'
    ' The variable int_GridNum starts as = 0.    '
    ' When the response box is up, the user can  '
    ' elect to have the Input Boxes redone using '
    ' Grid1(1)'s data and to compare Grid1(0) to.'

    Grid1(0).Col = c
    Grid1(1).Col = c

    If int_GridNum = 0 Then
        g1 = 1
        g2 = 2
    Else
        g1 = 2
        g2 = 1
    End If

    d2 = Grid1(g2 - 1).Text
    If t <> "" Then
        If t = "." Then
            t = "[blank]"
        ElseIf Right$(t, 1) = "." Then
            t = Left$(t, Len(t) - 1)
        End If
        d1 = "[compare Grid " & g2 & " to Input Box data]"
        b_CompareChange(0).Enabled = False
        b_CompareChange(1).Enabled = False
    Else
        t = Grid1(int_GridNum).Text
        If t = "" Then
            t = "[blank]"
        End If
        d1 = Grid1(g1 - 1).Text
        b_CompareChange(0).Enabled = True
        b_CompareChange(1).Enabled = True
    End If

    If d1 = "" Then d1 = "[blank]"
    If d2 = "" Then d2 = "[blank]"

    L_CompareField.Caption = Grid1(0).Columns(c).Caption
    t_GridData(0) = t
    L_DataFrom(1) = "  Grid " & g1 & " Data:"
    L_DataFrom(2) = "  Grid " & g2 & " Data:"
    t_GridData(1) = d1
    t_GridData(2) = d2

    b_ComparePutG2Data.Caption = "&Compute Input Boxes for Grid " & g1
    b_CompareUseInputBox.Caption = "Use Input &Box data in Grid " & g2
    o_CompareOptions(g2 - 1) = True

    L_Source(1) = Grid1(g1 - 1).Columns(int_Source_Col)
    L_Source(2) = Grid1(g2 - 1).Columns(int_Source_Col)
    If L_Source(1) = "" And L_Source(2) = "" Then
        L_CompareSource(0) = ""
        L_CompareSource(1) = ""
    Else
        L_CompareSource(0) = "          source:"
        L_CompareSource(1) = "          source:"
    End If

    f_Compare.Caption = "Comparing #" & Grid1(0).Columns(int_OrdNum_Col) & _
                        " " & Grid1(0).Columns(int_Bid_Name_Col) & ":"
    f_Compare.Visible = True
    f_Compare.ZOrder 0
    gn = int_GridNum
    Do While f_Compare.Visible And gn = int_GridNum
        DoEvents
    Loop

        If bln_Abort Then
            bln_Abort = False
            Exit Sub
        End If

    Select Case int_GridNum
    Case Is = gn
        If o_CompareOptions(2) Then
            response = "0"
        ElseIf o_CompareOptions(3) Then
            response = "x"
        ElseIf o_CompareOptions(4) Then
            GoTo ExitCompare
        End If
    Case Is = 0
        grid2Compare2Boxes = 1
    Case Is = 1
        grid2Compare2Boxes = 0
    End Select
    t = ""

Return

End Sub


Private Sub SaveCompareChanges(ByVal gn As Integer, _
                               ByVal response As String)
    gn = gn - 1 ' gn comes in as "1" or "2". Needs to be "0" and "1". '

    Grid1(gn) = response
    ' Advance row number to save change, '
    ' then go back to original row:      '
    Grid1(gn).Row = Grid1(gn).Row + 1
    If Grid1(gn).Row = 0 Then
        Grid1(gn).Scroll 0, -1
    End If
    If Grid1(gn).Row > 0 Then
        Grid1(gn).Row = Grid1(gn).Row - 1
    End If
End Sub


'********************************************
' Procedure: mDivider_Click                 '
' Purpose:                                  '
    ' Adds a divider line between sections  '
    ' on a grid (where a "section" is a set '
    ' of entries for the same Prior Bids).  '
'********************************************
'
Private Sub mDivider_Click()

    Dim xx As Integer
    Dim g As Integer

    ' Tutorial
    If frm_Tutorial.Visible Then
        If b_Continue.Visible = False Then
            Run_Divider_Tutorial
        End If
        Exit Sub
    End If

    str_Response = MyMsgBox("Add a divider to the Prior Bids section?" & vbCr & _
                            Grid1(int_GridNum).Columns(int_Prior_Bids_Col).Text & _
                            2, "YN", "", "BidBase")
    If str_Response <> "yes" Then Exit Sub

    g = int_GridNum
    mClear_Click
    If Grid1(g).Row <> 0 Then
        mViewFromHereDown_Click g
        Grid1(g).Row = 0
    End If

    t_Round = Grid1(g).Columns(int_Round_Col).Text
    t_Player = Grid1(g).Columns(int_Player_Col).Text
    t_Prior = Grid1(g).Columns(int_Prior_Bids_Col).Text
    t_Order = "ZZZZZZ"

    ' Save it: '
    ck_SaveNew.Value = 1
    mSave_Click

    mClear_Click
End Sub


Private Sub Run_Divider_Tutorial()
    frm_Tutorial.Caption = "Add Divider"
    L_Tutorial = "The Dividing Line is an entry with ============= " & _
                 "in the Cnv_Name. It acts as a separator between " & _
                 "sets of entries." & vbCr & "(A 'set' is a group of entries " & _
                 "which all have the same Prior Bids.)" & _
                 vbCr & vbCr & _
                 "The divider is for when you are viewing more than " & _
                 "one set of entries in a grid. The divider makes it " & _
                 "easier to see where one set ends and the next begins." & _
                 "BidBase also uses it to determine when the end of a " & _
                 "section has been reached."
    L_Tutorial.Height = 2800
    ResizeTutorial
    b_Continue.Visible = True
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#E3"
End Sub


Private Sub Run_CompareEntries_Tutorial()
    frm_Tutorial.Caption = "Compare Entries"
    L_Tutorial = "The Compare Entries menu option lets you compare " & _
                 "similar entries in different sections to ensure " & _
                 "that the specs are consistent for the same types " & _
                 "of bids." & _
                 vbCr & vbCr & _
                 "For example, responder's bids after 1C-1H should " & _
                 "have essentially the same specs as after 1C-1S." & _
                 vbCr & vbCr & _
                 "This feature is extremely important for insuring " & _
                 "the accuracy and consistency of BidBase. Click " & _
                 "the Details button for a complete discussion."
    L_Tutorial.Height = 2800
    ResizeTutorial
    b_Continue.Visible = True
    b_TutorDetails.Visible = True
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#E11a"
End Sub


Private Sub mDupesGrid_Click(Index As Integer)
    ' Index=0 -- Display "Dupes" grid. This grid '
    ' contains entries with the same Bid Order # '
    ' as the selected entry in the top grid. The '
    ' entries may not be as closely related as   '
    ' those defined in the database as Dupes.    '
    ' Look at the example dupes below. Another   '
    ' set of Prior Bids which would appear in the'
    ' Dupes grid when Index=0 is 1C-1N, 1C-P, etc.
    ' While fewer entries in these sections will '
    ' match the other Dupes sections, some of    '
    ' them still should have the same specs.     '

    ' Index=1 -- Display only entries which are  '
    ' in the database as being "Dupe Entries".   '
    ' For example, if the Prior Bids for the top '
    ' grid are "1C-1H", the database list these  '
    ' as dupes: 1C-1S, 1D-1H, 1D-1S, and 1H-1S.  '

    ' Index=2 -- Display only Dupes which have   '
    ' the same Convention name. Some entries have'
    ' the same bid and Bid Ord.# but different   '
    ' specs with Pct_Used set to 0 to deactivate '
    ' the unused alternatives. Normally, these   '
    ' entries have different Convention Names.   '
    ' So this option displays only entries which '
    ' should have the same specs.                '

    If frm_Tutorial.Visible And _
       frm_Tutorial.Caption <> "Dupes Grid" _
    Then
        If b_Continue.Visible = False Then
            Run_DupesGrid_Tutorial
            Exit Sub
        End If
    End If

    mDupesGrid(Index).Checked = Not mDupesGrid(Index).Checked

    If mDupesGrid(Index).Checked And _
       Index <> 0 And _
       mDupesGrid(0).Checked = False _
    Then
        mDupesGrid(0).Checked = True
    End If

    If mDupesGrid(0).Checked Then
        Grid1_Click 0
    ElseIf bln_TwoGrids Then
        mTwoGrids_Click
    End If
End Sub


Private Sub mDupeTopEntry_Click()
    ' Press Ctrl-W to open the grid view box.     '
    ' Enter the prior bids for the top grid, like '
    ' 1C-D-1D, and for the 2nd grid, like 1C-D-1H.'
    ' Select entry in the top grid. Ex.: Bid = 1H '
    ' which copies that entry to the Input boxes. '
    ' Click Dupe Top Entry in the Edit menu.      '
    ' The entry in the Input boxes will be modded '
    ' to be a new entry in the second grid.       '
    ' Example: Prior bids in the Input box changes'
    ' to PB=1C-D-1H, Bid=1S. The suit Qty & HCP   '
    ' will be swapped between Diamonds and Hearts '
    Dim b As String
    Dim n As Integer
    Dim p1 As String
    Dim p2 As String
    Dim i As Integer

    p1 = t_Prior
    If mTwoGrids.Checked = False Then
        p2 = MyMsgBox("Enter prior bids for new entry.", 1, t_Prior, "")
        t_Show(0) = p1
        t_Show(1) = p1
        t_Show(2) = p2
        t_Show(3) = p2
        b_Show_Click
        mTwoGrids_Click
    Else
        p2 = Grid1(1).Columns(int_Prior_Bids_Col)
    End If

    bln_DupingEntry = True
    DisplayBid2Edit 2

End Sub


Private Sub mEnter_Click()
    ' Menu item to bring up hand entry box. '
    With f_EnterHand
        If .Top > Me.Height Or .Left > Me.Width Then
            .Move 15000, 6000
        End If

        If frm_Tutorial.Visible Then
            .Visible = True
            Run_EnterHands_Tutorial
        End If

        .Visible = Not f_EnterHand.Visible
        L_BidFound = ""
        .ZOrder 0

        If .Visible Then
            t_FindHand(int_FindHandNum).SetFocus
        End If
    End With
End Sub


Private Sub Run_EnterHands_Tutorial()

    frm_Tutorial.Caption = "Enter Your Own Hands"
    L_Tutorial = "When you enter a hand and other data and click 'Find " & _
                 "Bid', if a bid is found, it will be displayed, plus " & _
                 "the bid's entry will be loaded into the Input Boxes " & _
                 "at the bottom of the screen." & _
                 vbCr & vbCr & _
                 "If no entry is found for the hand and other data " & _
                 "entered, the section of bids for the prior bids " & _
                 "entered will be displayed in the grid so that you " & _
                 "look for an entry with the bid you were expecting. " & _
                 "Click on the left edge of an entry to display it " & _
                 "in the Input Boxes. You can edit that entry or " & _
                 "click the Add Entry button to have the specs for " & _
                 "your hand entered into the Input boxes for making " & _
                 "a new entry, if you wish."

    If f_EnterHand.Left < 3000 Then
        frm_Tutorial.Move f_EnterHand.Left + f_EnterHand.Width + 50, _
                          f_EnterHand.Top - 600
    End If
    L_Tutorial.Height = 5600
    ResizeTutorial
    b_TutorDetails.Visible = True
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#F2"

    Wait

    L_Tutorial = "If you feel that a bid should have been found, enter " & _
                 "the bid in the 'Show Analysis' box to see how it failed." & vbCr & vbCr & _
                 "BidBase is intended primarily for encoding " & _
                 "conventions and bidding systems, and these usually " & _
                 "take only 2 or 3 rounds to bid, so entering a long " & _
                 "set of Prior Bids which are not part of a convention " & _
                 "or system reduces the chances of finding a bid. " & _
                 "Also, you can only enter one hand here. To enter" & _
                 "deals where more than one hand is known, " & _
                 "select the 'Deal Hands' menu option. " & _
                 vbCr & vbCr & _
                 "You can try entering a hand now if you wish; " & _
                 "otherwise, close the Enter Hands box." & _
                 vbCr & vbCr & _
                 "There are numerous features which make it " & _
                 "faster and easier to enter hands and bids. " & _
                 "Click the Details button to read about them."

    If f_EnterHand.Left < 3000 Then
        frm_Tutorial.Move f_EnterHand.Left + f_EnterHand.Width + 50, _
                          f_EnterHand.Top - 600
    End If
    L_Tutorial.Height = 4300
    ResizeTutorial
    b_TutorDetails.Visible = True
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#F2"

    Wait
End Sub


Private Sub ImportDeal_Tutorial_Click()
    frm_Tutorial.Caption = "Import a Practice Deal Hand"
    L_Tutorial = "When using BidBase Practice to deal and bid, " & _
                 "if you want to see, edit or add a different " & _
                 "bid, click Edit in that program and it will " & _
                 "save the hand and prior bids to disk. Then " & _
                 "you can use this function to import it into " & _
                 "the Editor to find the bid to work on. " & _
                 vbCr & vbCr & _
                 "This makes it very quick and easy to modfiy " & _
                 "BidBase if you disagree with its bidding." & _
                 vbCr & vbCr & _
                 "The easiest way to use this is to click " & _
                 "Edit - Stay On Top in the Practice program " & _
                 "and place it to the right of the hand entry " & _
                 "window of the Editor, then after clicking " & _
                 "the Edit button in Practice, click anywhere " & _
                 "in the Editor to make it active and press " & _
                 "Ctrl-I to import the entry."

    If f_EnterHand.Left < 3000 Then
        frm_Tutorial.Move f_EnterHand.Left + f_EnterHand.Width + 50, _
                          f_EnterHand.Top - 600
    End If
    L_Tutorial.Height = 3500
    ResizeTutorial
    b_TutorDetails.Visible = True
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#F2b"

    Wait
End Sub


Private Sub mFindEntryID_Click()
    Dim i As Long
    Dim j As Integer
    Dim fID As String

    If bln_SkipIt Then
        ' This routine was called by the Edit Bid '
        ' button in the Find Bid for Hand box.    '
        i = Val(Mid$(L_FindBidID & "    ", 5))
        bln_SkipIt = False
    Else
        i = Val(MyMsgBox("Enter ID# to find:", 1, "", ""))
    End If
    If i < 1 Then Exit Sub

    Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
    Set rs_BidTable = db_BidBase.OpenRecordset(str_BidsTbl)
    rs_BidTable.Index = "PrimaryKey"
    rs_BidTable.Seek "=", i

    If rs_BidTable.NoMatch Then
        str_Response = MyMsgBox("ID#" & i & " not found.", 1, "ok", "")
        Exit Sub
    End If

    lng_ID = i
    
    ' The following looks in the database. '
    ' Searching the grid only looks in the '
    ' entries displayed. I can't figure how'
    ' to get the grid to display the whole '
    ' database. '
    DisplayBid2Edit 99

    If Not bln_SkipIt And Val(Mid$(L_FindBidID & "    ", 5)) <> lng_ID Then
        str_Response = MyMsgBox("Entry ID#" & i & " is in the Input Boxes.", 1, "ok", "")
    End If
End Sub

Private Sub mFourGrids_Click()
    ' Display four grids: '
    Dim g As Integer
    Dim c As Integer

    If frm_Tutorial.Visible And b_Continue.Visible = False Then
        Run_FourGrids_Tutorial
        Exit Sub
    End If

  With mFourGrids
    .Checked = Not .Checked
    If .Checked Then
        For g = 0 To 1
           If mSplit(g).Checked Then
                mSplit_Click g
            End If
        Next
        If mDupesGrid(0).Checked Then
            mDupesGrid_Click 0
        End If
        Grid1(2).Visible = False
        If Grid1(1).Visible = False Then
            mTwoGrids_Click
        End If
        For g = 3 To 4
            For c = 0 To int_LastBidCol
                Grid1(g).Columns(c).Width = sng_ColWidths(c)
                Grid1(g).Columns(c).Visible = Grid1(0).Columns(c).Visible
            Next
            Grid1(g).Col = 0
        Next
        ColorGrids 0
    End If

    Grid1(3).Visible = .Checked
    Grid1(4).Visible = .Checked
    VScroll1.Visible = .Checked Or bln_TwoGrids
    bln_InputBoxes = Not .Checked
    mInputBoxes.Checked = Not .Checked
    frame0(6).Visible = Not .Checked
  End With
    Form_Resize
End Sub


Private Sub mFullTutorial_Click()
    mFullTutorial.Checked = Not mFullTutorial.Checked
    If mFullTutorial.Checked Then
        Welcome
    Else
        frm_Tutorial.Visible = False
        bln_Waiting = False
    End If
End Sub


Private Sub Run_FourGrids_Tutorial()
    frm_Tutorial.Caption = "Four Grids"
    L_Tutorial = "Displaying four grids at once lets you compare " & _
                 "similar bids for all four suits." & _
                 vbCr & vbCr & _
                 "To make room, the Input Boxes are hidden automatically." & _
                 vbCr & vbCr & _
                 "To toggle off the 4-grid display, click the menu " & _
                 "option again." & _
                 vbCr & vbCr & _
                 "Click on 'View'  -  'Four Grids' now to turn off the " & _
                 "four grids display." & vbCr & vbCr & _
                 "One grid is colored different from the rest. " & _
                 "That is the last selected grid. Any menu options " & _
                 "which work on a grid, such as Edit Entry, will work " & _
                 "on that grid."
    L_Tutorial.Height = 3700
    ResizeTutorial
    b_Continue.Visible = True
    If mFourGrids.Checked = False Then
        mFourGrids_Click
    End If
    t_Show(4) = "1C-P"
    t_Show(5) = "1D-P"
    b_Show_Click
    b_TutorDetails.Visible = True
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#V4b"
End Sub


Private Sub mGoBack_Click()
    ' Causes grid(s) to display the '
    ' last displayed set of entries.'

    If frm_Tutorial.Visible Then
        If b_Continue.Visible = False Then
            frm_Tutorial.Caption = "Go Forward In Grid"
            Run_GoFwd_Tutorial
            Exit Sub
        End If
    End If

    If int_GridPicksNum > 0 Then
        int_GridPicksNum = int_GridPicksNum - 1
        t_Show(0) = str_GridPick(0, 0, int_GridPicksNum)
        t_Show(1) = str_GridPick(0, 1, int_GridPicksNum)
        cmb_Show(0).ListIndex = Val(str_GridPick(0, 2, int_GridPicksNum))
        If cmb_Show(0).ListIndex < 0 Then
            cmb_Show(0).ListIndex = int_Prior_Bids_Col
        End If

        If str_GridPick(1, 0, int_GridPicksNum) <> "" Then 'two grids
            t_Show(2) = str_GridPick(1, 0, int_GridPicksNum)
            t_Show(3) = str_GridPick(1, 1, int_GridPicksNum)
            cmb_Show(1).ListIndex = Val(str_GridPick(1, 2, int_GridPicksNum))
            If cmb_Show(1).ListIndex < 0 Then
                cmb_Show(1).ListIndex = int_Prior_Bids_Col
            End If

            If Not bln_TwoGrids Then
                ' Show 2nd grid: '
                mTwoGrids_Click
            End If
        Else ' one grid '
            If bln_TwoGrids Then
                ' Hide 2nd grid: '
                mTwoGrids_Click
            End If
        End If

        bln_GridFwdBack = True
        b_Show_Click
        bln_GridFwdBack = False

        If int_GridPicksNum > 0 Then
            mGoBack.Enabled = True
        Else
            mGoBack.Enabled = False
        End If
    End If
    mGoFwd.Enabled = True
End Sub


Private Sub mGoFwd_Click()
    ' When the user has gone back to a '
    ' previous set of bids, this menu  '
    ' option goes forward again.       '

    If frm_Tutorial.Visible Then
        If b_Continue.Visible = False Then
            frm_Tutorial.Caption = "Go Forward In Grid"
            Run_GoFwd_Tutorial
            Exit Sub
        End If
    End If

    If str_GridPick(0, 0, int_GridPicksNum + 1) = "" Then
        Exit Sub
    End If

    mGoFwd.Enabled = False

    int_GridPicksNum = int_GridPicksNum + 1
    t_Show(0) = str_GridPick(0, 0, int_GridPicksNum)
    t_Show(1) = str_GridPick(0, 1, int_GridPicksNum)
    cmb_Show(0).ListIndex = Val(str_GridPick(0, 2, int_GridPicksNum))
    If cmb_Show(0).ListIndex < 0 Then
        cmb_Show(0).ListIndex = int_Prior_Bids_Col
    End If

    If str_GridPick(1, 0, int_GridPicksNum) <> "" Then 'two grids
        t_Show(2) = str_GridPick(1, 0, int_GridPicksNum)
        t_Show(3) = str_GridPick(1, 1, int_GridPicksNum)
        cmb_Show(1).ListIndex = Val(str_GridPick(1, 2, int_GridPicksNum))
        If cmb_Show(1).ListIndex < 0 Then
            cmb_Show(1).ListIndex = int_Prior_Bids_Col
        End If
        If Not bln_TwoGrids Then mTwoGrids_Click
    Else
        If bln_TwoGrids Then mTwoGrids_Click
    End If

    bln_GridFwdBack = True
    b_Show_Click
    bln_GridFwdBack = False

    Dim i1 As Integer
    Dim i2 As Integer
    Dim i3 As Integer

    If int_GridPicksNum > 49 Then
        For i1 = 0 To 1
            For i2 = 0 To 2
                For i3 = 1 To 10
                    str_GridPick(i1, i2, i3) = str_GridPick(i1, i2, 40 + i3)
                Next

                For i3 = 11 To 50
                    str_GridPick(i1, i2, i3) = ""
                Next
            Next
        Next

        int_GridPicksNum = 10
    End If

    If str_GridPick(0, 0, int_GridPicksNum + 1) = "" Then
        mGoFwd.Enabled = False
    Else
        mGoFwd.Enabled = True
    End If
    mGoBack.Enabled = True
End Sub


Private Sub Run_GoFwd_Tutorial()
    L_Tutorial = "The left and right 'arrows' on the menu bar work like " & _
               "the back and forward arrows in a web browser. Clicking " & _
               "<-- takes you back to the previous grid contents and " & _
               "--> takes you forward to the next grid contents which " & _
               "you previously displayed." & _
               vbCr & vbCr & _
               "If you want to increment the grid contents rather than " & _
               "return to a screen you previously scrolled back from, " & _
               "then click =>  or 'View - Jump To Next Section' (Ctrl-J)." & _
               vbCr & vbCr & _
               "These are advanced features which you don't need unless " & _
               "you are working extensively with the database. See the " & _
               "help file for more information."
    L_Tutorial.Height = 3450
    ResizeTutorial
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#S1"
End Sub


Private Sub mHelpInput_Click()
    str_NoteFile(0) = "CardShark_BidBase_Editor_Input_Boxes.htm"
    ShowDocs
End Sub


Private Sub mHelpMenus_Click()
    str_NoteFile(0) = "CardShark_BidBase_Editor_Menus.htm"
    ShowDocs
End Sub


Private Sub mImportPracticeDeal_Click()
    ' This is called in the File menu.  '
    ' When a user wants to edit a bid   '
    ' in the Practice program, the bid  '
    ' and hand are put into a file.     '
    ' This sub opens that file and puts '
    ' the info into the f_EnterHand box.'
    If frm_Tutorial.Visible Then
        f_EnterHand.Visible = True
        frm_Tutorial.Visible = True
        ImportDeal_Tutorial_Click
        f_EnterHand.Visible = False
        Exit Sub
    End If

    If Dir$(App.Path & "\EditPractice.txt") = "" Then
        str_Response = MyMsgBox("The EditPractice.txt file containing " & vbCrLf & _
                                "the hand and bidding is not in the " & vbCrLf & _
                                "BidBase directory.", 3, "ok", "")
        Exit Sub
    End If

    Dim h As String
    Dim p As String
    Dim pr(4) As String ' 1st 4 calls. '
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    cmb_Bid = ""

    b_ClearFind_Click

    'On Local Error GoTo ImportErr
    Close #1
    
    Open (App.Path & "\EditPractice.txt") For Input As #1
        Input #1, h
        Input #1, p
        If Not EOF(1) Then
            Input #1, str_ImportSource
            i = InStr(str_ImportSource, "BON=")
            If i > 0 Then
                Dim BON As String
                BON = Mid$(str_ImportSource, i + 4)
                str_ImportSource = Left$(str_ImportSource, i - 2)
            End If
            i = InStr(str_ImportSource, "ID=")
            If i > 0 Then
                lng_ID = Val(Mid$(str_ImportSource, i + 3))
                str_ImportSource = Left$(str_ImportSource, i - 2)
            End If
        Else
            str_ImportSource = ""
        End If
        t_Find_Ref.Text = str_ImportSource
    Close #1
    On Local Error GoTo 0

    f_EnterHand.Visible = True
    t_FindHand(0) = h
    t_FindHand(1) = ""
    t_FindPrior = p
    If p <> "" Then
        i = 0
        j = 0
        Do While p <> "" And i < 4
            i = i + 1
            k = InStr(p, "-")
            If k = 0 Then
                pr(i) = p
                p = ""
            Else
                pr(i) = Left$(p, k - 1)
                p = Mid$(p, k + 1)
            End If
            If pr(i) = "P" Then
                ' Count # of consecutive opening passes. '
                j = j + 1
            End If
            'If Left$(p, 1) = "P" Then
            '    j = j + 1
            'End If
            'If Len(p) > 2 Then
            '    If k > 0 Then
            '        p = Mid$(p, k + 1)
            '    Else
            '        p = ""
            '    End If
            'Else
            '    Exit Do
            'End If
        Loop
        If i = 4 And p <> "" Then
            ' More than 1 round of bidding. '
            'Stop
        End If
        If i = j Then
            ' All prior calls = "P". '
            If i = 3 And i = j Then
                cmb_FindPos.ListIndex = 6
            ElseIf i = 2 Then
                cmb_FindPos.ListIndex = 5
            ElseIf i = 1 Then
                cmb_FindPos.ListIndex = 4
            End If
        End If
    End If
    L_FindBidID.Caption = "ID# " & lng_ID & "   BON: " & lng_BON
    b_FindBidForHand_Click
    If L_BidFound = "" And lng_ID > 0 Then
        ForceEntry2ShowInGrid lng_ID
        L_BidFound = cmb_Bid
    End If
    f_Msg.Visible = False
    Exit Sub

ImportErr:
    str_Response = MyMsgBox("Error importing practice hand.", 1, "ok", "")
    On Local Error GoTo 0
End Sub


Private Sub mInputBoxes_Click()
    If frm_Tutorial.Visible And _
       b_Continue.Visible = False And _
       mMenuTutorial.Checked _
    Then
        Run_HideInputs_Tutorial
        Exit Sub
    End If

    bln_InputBoxes = Not bln_InputBoxes
    mInputBoxes.Checked = bln_InputBoxes
    frame0(6).Visible = bln_InputBoxes
    Form_Resize
End Sub


Private Sub mInputBoxTutorial_Click()
    If frm_Tutorial.Visible = False Then
        mInputBoxTutorial.Checked = True
        frm_Tutorial.Visible = True
        frm_Tutorial.ZOrder 0
        Run_Edit_Tutorial
        mInputBoxTutorial.Checked = False
        b_Continue.Visible = False
    Else
        mInputBoxTutorial.Checked = False
    End If
End Sub


Private Sub mLookForDeadNoteFiles_Click()
    If frm_Tutorial.Visible Then
        If b_Continue.Visible = False Then
            Run_LookForDeadNotes_Tutorial
            Exit Sub
        End If
    End If

    Dim pth As String
    Dim f As String
    Dim n As Integer
    Dim xzx As Long
    Dim xx As Integer
    Dim n1 As Integer
    Dim pb As String
    Dim b As String
    Dim fID As Long

    pth = str_BBpath & "\Notes\"
    f = Dir$(pth & "*.txt")

    Do While Val(f) > 0
        n = InStr(f, "_ID")
        If n > 0 Then
            fID = Val(Mid$(f, n + 3))
            rs_BidTable.Index = "PrimaryKey"
            rs_BidTable.Seek "=", fID

            If rs_BidTable.NoMatch = True Then
                str_Response = MyMsgBox("The file " & f & " has an ID# " & vbCr & _
                                        "not found in the database.", 2, "ok", "")
                xzx = Shell("Notepad " & pth + f, vbNormalFocus)
                str_Response = MyMsgBox("Delete the file " & f & "?", 1, "YNC", "BidBase")
                Select Case str_Response
                Case Is = "yes"
                    Kill pth & f
                Case Is = "cancel"
                    Exit Sub
                End Select

            Else 'extract bid and prior bids:
                n1 = n
                Do
                    n = n - 1
                Loop While Mid$(f, n, 1) <> "-" And n > 3
                pb = ""
                If Mid$(f, n, 1) = "-" Then
                    b = Mid$(f, n + 1, n1 - n - 1)
                    If n > 2 Then
                        pb = Left$(f, n - 1)
                    End If
                Else
                    b = Left$(f, n1 - 1)
                End If

                If Trim$(rs_BidTable!Prior_Bids) <> pb Then
                    str_Response = MyMsgBox("The file " & f & " has a prior bid " & vbCr & _
                                            "which does not match the entry for the ID#.", 2, "ok", "")
                    xzx = Shell("Notepad " & str_BBpath & pth & f, vbNormalFocus)
                    str_Response = MyMsgBox("Delete this file?", 1, "YNC", "BidBase")
                    Select Case str_Response
                    Case Is = "yes"
                        Kill pth & f
                    Case Is = "cancel"
                        Exit Sub
                    End Select

                ElseIf Trim$(rs_BidTable!bid) <> b Then
                    t_Show(0) = pb
                    t_Show(1) = pb
                    cmb_Show(0).ListIndex = int_Prior_Bids_Col
                    b_Show_Click

                    ' Locate entry: '
                    Grid1(0).Row = 0
                    Do While Grid1(0).Columns(int_LastBidCol).Text <> fID
                        Grid1(0).Row = Grid1(0).Row + 1
                        DoEvents
                    Loop
                    ' Display in Input Boxes: '
                    mEdit_Click 0

                    str_Response = MyMsgBox("The file " & f & " has a bid which " & vbCr & _
                                            "does not match the entry for the ID#.", 2, "ok", "")
                    xzx = Shell("Notepad " & pth & f, vbNormalFocus)
                    str_Response = MyMsgBox("Delete the file " & f & "?", 1, "YNC", "BidBase")
                    Select Case str_Response
                    Case Is = "yes"
                        Kill pth & f
                    Case Is = "cancel"
                        Exit Sub
                    End Select
                End If
            End If
        Else
            str_Response = MyMsgBox("The file " & f & " does not have an ID# in its name.", 1, "ok", "")
            xzx = Shell("Notepad " & pth & f, vbNormalFocus)
            str_Response = MyMsgBox("Delete the file " & f & "?", 1, "YNC", "BidBase")
            Select Case str_Response
            Case Is = "yes"
                Kill pth & f
            Case Is = "cancel"
                Exit Sub
            End Select
        End If
        Do
            f = Dir$
            DoEvents
        Loop While InStr("CDHS", UCase$(Mid$(f, 2, 1))) = 0 And _
                   Val(f) > 0
    Loop

    str_Response = MyMsgBox("All entry note files have been checked.", 1, "ok", "")
End Sub


Sub Run_LookForDeadNotes_Tutorial()
    frm_Tutorial.Caption = "Dead Note Files"
    L_Tutorial = "This routine looks for note files which " & _
               "have names made up from an entry's Prior " & _
               "Bids, Bid, and ID#. (Example: 1H-P-1S_ID3120.TXT)" & _
               vbCr & vbCr & _
               "If an entry with the " & _
               "file's ID# does not exist, or if the Bid " & _
               "or Prior Bid do not match, you will be shown " & _
               "the file in Notepad and asked if you want " & _
               "to delete the file." & _
               vbCr & vbCr & _
               "If you want to rename it, use Notepad's Save-As " & _
               "to save the file with a new name, then use BidBase " & _
               "File Viewer to delete the original file name."
    L_Tutorial.Height = 2800
    ResizeTutorial
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#E12"
End Sub


Private Sub mMake9_Click()
    ' Convert SAYC entries to BONs of 900000+. '
    Dim i As Long
    Dim BON As String
    Dim bon1 As Long
    Dim bon9 As Long
    Dim bonLen As Integer
    Dim id As Long
    Dim r As Long
    Dim Do9s As Boolean

    Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
    Set rs_BidTable = db_BidBase.OpenRecordset(str_BidsTbl)
    rs_BidTable.Index = "PrimaryKey"
    r = 0
    Grid1(0).Row = 0
    Do9s = True

    Do
        If r > Grid1(0).Row Then
            Grid1(0).Row = r 'Grid1(0).Row + 1
        End If

        BON = Grid1(0).Columns(int_OrdNum_Col).Text
        If Left$(BON & " ", 1) = "9" And Grid1(0).Row = 0 Then
            Exit Sub
        End If
'        If bon = "401000" Then Stop
        If Do9s Then
            If BON = "ZZZZZZ" Or BON = "" Then
                Do9s = False
                Data1(0).Refresh
                Grid1(0).Refresh
                r = 0
                Grid1(0).Row = 0
                BON = Grid1(0).Columns(int_OrdNum_Col).Text
            End If
        End If

        If InStr("012345678", Right$(BON, 1)) = 0 Then
            Mid$(BON, 6, 1) = "0"
            GoSub SaveRenum9
        End If

        If Do9s Then
            If InStr("0123456789", Left$(BON, 1)) = 0 Then MsgBox "Error converting BON's."
            If Grid1(0).Columns(0).Text <> "0" And Grid1(0).Columns(1).Text <> "0" Then
                If Left$(BON, 1) = "9" Then
                    bonLen = Len(BON)
                    ' Add 8 to left and remove right-most 0.     '
                    ' Examples: 910010 = 891001, 801001 = 890101 '
                    For i = bonLen To 1 Step -1
                        If Mid$(BON, i, 1) = "0" Then
                            If i < bonLen Then
                                BON = "8" & Left$(BON, i - 1) & Mid$(BON, i + 1)
                            Else
                                BON = "8" & Left$(BON, i - 1)
                            End If
                            Exit For
                        End If
                    Next
                    If Left$(BON, 1) <> "8" Then MsgBox "Error converting BON's."

                    GoSub SaveRenum9
                Else
                    r = r + 1
                End If
            Else
                r = r + 1
            End If

        ElseIf Grid1(0).Columns(0).Text <> "0" And Grid1(0).Columns(1).Text <> "0" Then
            BON = Grid1(0).Columns(int_OrdNum_Col).Text
            If Left$(BON, 1) = "9" Then Exit Do
            If InStr("012345678", Left$(BON, 1)) = 0 Then MsgBox "Error"
            bonLen = Len(BON)
            ' Add 9 to left and remove right-most 0.     '
            ' Examples: 310010 = 931001, 401001 = 940101 '
            For i = bonLen To 1 Step -1
                If Mid$(BON, i, 1) = "0" Then
                    If i < bonLen Then
                        BON = "9" & Left$(BON, i - 1) & Mid$(BON, i + 1)
                    Else
                        BON = "9" & Left$(BON, i - 1)
                    End If
                    Exit For
                End If
            Next
            If Left$(BON, 1) <> "9" Then MsgBox "Already a 9xxxxxx BON"

            GoSub SaveRenum9
        Else
            r = r + 1
            If r > 1000 Then
                MsgBox "Break point"
                Exit Sub
            End If
        End If
        DoEvents
    Loop
Exit Sub

SaveRenum9:
            id = Val(Grid1(0).Columns(int_ID_Col))
      If Left$(BON, 1) = "Z" Then
        MsgBox "Z line"
      End If
            With rs_BidTable
                .Seek "=", id
                If .NoMatch = True Then MsgBox "Already a 9xxxxxx BON"
                .Edit
                !Bid_Order_Number = BON
                .Update
            End With
            Data1(0).Refresh
            DoEvents
        Return
End Sub

Private Sub mNextSection_Click()
    If frm_Tutorial.Visible Then
        If b_Continue.Visible = False Then
            Run_NextSection_Tutorial
            Exit Sub
        End If
    End If

    mNextSection.Checked = True
    mNextSection2_Click
    mNextSection.Checked = False
End Sub


Private Sub Run_NextSection_Tutorial()
    frm_Tutorial.Caption = "Jump To Next Section"
    L_Tutorial = "A 'section' or 'set' of bids are those which have the " & _
               "same Prior_Bids, such as '1C-P'. It is usually easier " & _
               "to work on one set of bids at a time. The 'Jump To Next Section' " & _
               "menu command lets you view the next section of bids " & _
               "without having to bring up the Show/Sort window." & _
               vbCr & vbCr & _
               "All you have to do is click on the Jump To Next Section menu " & _
               "command or press Ctrl-J to increment just the current " & _
               "grid, or click => on the menu bar to increment one OR " & _
               "two grids." & _
               vbCr & vbCr & _
               "Click <- and -> to scroll through sections which have " & _
               "previously been displayed."
    L_Tutorial.Height = 3500
    ResizeTutorial
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#V01"
End Sub


Private Sub mNextSection2_Click()
    Dim i As Integer
    Dim n As Integer
    Dim b As String
    Dim c As String
    Dim pbid1 As String
    Dim pbid2 As String
    Dim w0 As String
    Dim rs As String

    If frm_Tutorial.Visible And _
       b_Continue.Visible = False And _
       frm_Tutorial.Caption <> "Jump To Next Section - Both Grids" _
    Then
        Run_NextSection2_Tutorial
        Exit Sub
    End If

    mGoBack.Enabled = False
    mGoFwd.Enabled = False
    mNextSection.Enabled = False
    mNextSection2.Enabled = False

    For i = 0 To 1
        If mNextSection.Checked Then
            i = int_GridNum
        End If

        pbid1 = PB_To_SortOrd_Code(t_Show(i * 2))

        If Left$(Trim$(pbid1), 1) < "a" Then
            b = Right$(" " & pbid1, 1)
            n = InStr(" 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]", b)
            c = Mid$(" 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]", n + 1, 1)
            pbid1 = Left$(pbid1, Len(pbid1) - 1) + c
        Else
            pbid1 = pbid1 & "0"
        End If

        pbid2 = String$(35, "z")

        ' In mNextSectionBack2, we can simply get the current     '
        ' Prior_Bids and the do a MovePrevious to move back one   '
        ' entry to find the last entry in the preceding section   '
        ' of Prior_Bids, but the same approach doesn't work here  '
        ' because if we simply search for the first entry for the '
        ' current Prior_Bids and then do a MoveNext, it will move '
        ' to the entry in the same section of Prior_Bids. So we   '
        ' instead calculate the next PB_Sort_Order code in the    '
        ' sequence, but there may not be any entries for that code'
        ' (it may not even be a legal sequence, such as going     '
        ' from 1H-1N-P to 1H-1N-1C, where 1C is the next code     '
        ' after P, but is not legal) so we ask for all entries    '
        ' from that sequence to "zzzzz". Then the first sequence  '
        ' in the bunch returned will be the next legal sequence.  '

        rs = "SELECT * FROM Bids" & _
             " WHERE PB_Sort_Order " & _
                     "BETWEEN '" & pbid1 & _
                       "' AND '" & pbid2 & "'" & _
             " ORDER BY PB_Sort_Order;"

        Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
        Set rs_BidTableTemp = db_BidBase.OpenRecordset(rs, dbOpenForwardOnly)

        If rs_BidTableTemp.EOF Then
            str_Response = MyMsgBox("End of database.", 1, "ok", "")
            Exit Sub

        ElseIf t_Show(i * 2) = t_Show(i * 2 + 1) Then
            t_Show(i * 2 + 1) = rs_BidTableTemp!Prior_Bids
            t_Show(i * 2) = rs_BidTableTemp!Prior_Bids

        Else
            t_Show(i * 2) = rs_BidTableTemp!Prior_Bids

            pbid1 = PB_To_SortOrd_Code(t_Show(i * 2))
            b = Right$(" " & pbid1, 1)
            n = InStr(" 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]", b)
            c = Mid$(" 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]", n + 1, 1)
            pbid1 = Left$(pbid1, Len(pbid1) - 1) + c
            pbid2 = String$(35, "z")

            w0 = "PB_Sort_Order Between '" & _
                      pbid1 & "' AND '" & pbid2 & "'"

            rs = "SELECT * FROM Bids" & _
                 " WHERE " & w0 & _
                 " ORDER BY PB_Sort_Order;"

            Set rs_BidTableTemp = db_BidBase.OpenRecordset(rs, dbOpenForwardOnly)

            If rs_BidTableTemp.EOF Then
                str_Response = MyMsgBox("End of database.", 1, "ok", "") ' Shouldn't be possible
                Exit Sub
            End If

            t_Show(i * 2 + 1) = rs_BidTableTemp!Prior_Bids
        End If

        If Not bln_TwoGrids Or _
           mNextSection.Checked _
        Then
            Exit For
        End If
    Next

    b_Show_Click
    Grid1(0).Row = 0
    If Grid1(0).Columns(int_OrdNum_Col).Text = "" Then
        If Not bln_TwoGrids Or _
           Grid1(1).Columns(int_OrdNum_Col).Text = "" _
        Then
            str_Response = MyMsgBox("No more entries.", 1, "ok", "")
            Exit Sub
        End If
    End If

    mGoBack.Enabled = True
    mNextSection.Checked = False
    mNextSection.Enabled = True
    mNextSection2.Enabled = True
End Sub


Private Sub Run_NextSection2_Tutorial()
    frm_Tutorial.Caption = "Jump To Next Section - Both Grids"
    L_Tutorial = "The => menu command increments the sections in the grids. " & _
               "For example, if you have '1C-1H' displayed in a grid, " & _
               "clicking => increments the grid to '1C-1S'. If you have " & _
               "two grids displayed, the second grid will also be " & _
               "incremented." & vbCr & vbCr & _
               "If you have two grids displayed and only want to increment " & _
               "one of them, use the 'View - Jump To Next Section' menu " & _
               "command (Ctrl-J). The last grid clicked will be " & _
               "incremented. Click <- and -> to scroll through sections " & _
               "which have previously been displayed."
    b_Continue.Visible = True
    L_Tutorial.Height = 2800
    ResizeTutorial
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#I1"
End Sub


Private Sub mNextSectionBack2_Click()
    If frm_Tutorial.Visible Then
        If b_Continue.Visible = False Then
            frm_Tutorial.Caption = "Jump To Prior Section"
            Run_NextSection2_Tutorial
            Exit Sub
        End If
    End If

    Dim i As Integer
    Dim pbid As String

    Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
    Set rs_BidTable = db_BidBase.OpenRecordset(str_BidsTbl)

    '                                                        '
     With rs_BidTable
    '                '
      For i = 0 To 1
        Do
            ' Get first entry of current Prior_Bids:
            pbid = PB_To_SortOrd_Code(t_Show(i * 2))
            .Index = "PB_Sort_Order"
            .Seek "=", pbid

            ' Then get the previous entry:
            If Not .NoMatch Then
                .MovePrevious
            End If

            If .NoMatch Then
                Exit Sub
            End If
            DoEvents
        Loop While .NoMatch

        If t_Show(i * 2) = t_Show(i * 2 + 1) Then
            t_Show(i * 2 + 1) = !Prior_Bids
            t_Show(i * 2) = !Prior_Bids

        Else
            t_Show(i * 2) = !Prior_Bids
            pbid = PB_To_SortOrd_Code(t_Show(i * 2 + 1))
            .Index = "PB_Sort_Order"
            .Seek "=", pbid

            If Not .NoMatch Then
                .MovePrevious
            End If
            t_Show(i * 2 + 1) = !Prior_Bids
        End If
        If Not bln_TwoGrids Then Exit For
      Next
    '        '
     End With
    '                                                             '
    b_Show_Click
    mGoFwd.Enabled = False
    ' When a new section is shown, you obviously can't scroll fwd.
End Sub


Private Sub mPractice_Click()
    ' Run the Practice program: '

    If frm_Tutorial.Visible Then
        If b_Continue.Visible = False Then
            frm_Tutorial.Caption = "Bidding Practice"
            L_Tutorial = "Bidding Practice is a separate " _
                   & "program with its own help files." _
                   & vbCr & vbCr _
                   & "(Click on a menu command or Quit.)"
        End If
        b_TutorDetails.Visible = True
        str_TutorDetails = "CardShark_BidBase_Practice_Program.htm"
        Exit Sub
    End If

    On Local Error GoTo LoadPractice
    ' If Practice is not already loaded, '
    ' AppActivate will generate an error '
    ' which will Shell to the program.   '
    ' This keeps more than one Practice  '
    ' program from opening.              '
    AppActivate "BidBase Bidding Practice.exe"
    GoTo ExitRunPractice

LoadPractice:
    Dim ret As Long
    If Dir$(str_BBpath & "\BidBase Bidding Practice.exe") = "" Then
        str_Response = MyMsgBox("Practice program is not present.", 1, "ok", "")
    Else
        ret = Shell(str_BBpath & "\BidBase Bidding Practice.exe", vbNormalFocus)
    End If
    Resume ExitRunPractice

ExitRunPractice:
    On Local Error GoTo 0
End Sub


Private Sub mRestore_Click()
    If frm_Tutorial.Visible Then
        If b_Continue.Visible = False Then
            frm_Tutorial.Caption = "Restore Database File"
            Run_RestoreDB_Tutorial
        End If
        Exit Sub
    End If

    If Dir$(str_BBpath & "\Backup\*.*") = "" Then
        MkDir str_BBpath & "\Backup"
    End If

    Dim x As Integer

    str_Response = MyMsgBox("This will overwrite the existing file." & vbCr & _
                            "Have you backed it up yet?", 1, "YNC", "BidBase")
    Select Case str_Response
    Case Is = "cancel"
        Exit Sub
    Case Is = ""
        str_Response = MyMsgBox("You will be prompted for" & vbCr & _
                                "backing up the current file" & vbCr & _
                                "before being prompted for" & vbCr & _
                                "restoring another one.", 4, "ok", "")
        mBackup_Click
        Exit Sub
    End Select

    CmDia.DialogTitle = "Enter backup file name to restore:"
    CmDia.Filter = "*.accdb"
    CmDia.InitDir = str_BBpath & "\Backup\"
    CmDia.FileName = "bids-" & Right$(Date$, 4) & _
                     "-" & Left$(Date$, 5) & ".accdb"
    CmDia.ShowOpen
    If CmDia.FileTitle = "" Then
        Exit Sub
    End If

    'ChDir App.Path
    FileCopy str_BBpath & "\backup\" & CmDia.FileTitle, _
             str_DBPathAndName

    b_Show_Click ' Update the grids with the new database.'

End Sub


Private Sub Run_RestoreDB_Tutorial()
    L_Tutorial = "This command lets you overwrite the current database " & _
                 "with a database file you have previously backed up. " & _
                 vbCr & vbCr & _
                 "Normally, you should back up the current database " & _
                 "before doing this."
    b_Continue.Visible = True
    L_Tutorial.Height = 2300
    ResizeTutorial
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#F4"
End Sub


Private Sub mRestoreEntry_Click()
    If frm_Tutorial.Visible And _
       b_Continue.Visible = False And _
       mFullTutorial.Checked = False _
    Then
        If b_Continue.Visible = False Then
            Run_Restore_Entry
        End If
        Exit Sub
    End If

    bln_ViewGarbage = Not bln_ViewGarbage
    mRestoreEntry.Checked = bln_ViewGarbage
    Grid1(2).Visible = bln_ViewGarbage

    If bln_ViewGarbage Then
        Data1(2).Refresh
        If Grid1(1).Visible Then
            mTwoGrids_Click
        End If

        Dim cw As Column
        Dim i As Integer

        i = 0
        For Each cw In Grid1(2).Columns
            cw.Width = sng_ColWidths(i)
            i = i + 1
        Next

    ElseIf frm_Tutorial.Caption = "Restore Entries" And _
           b_Continue.Visible _
    Then
        b_Continue_Click
        Form_Resize
        Exit Sub
    End If

    Form_Resize
End Sub


Private Sub Run_Restore_Entry()
    frm_Tutorial.Caption = "Restore Entries"
    L_Tutorial = "When you delete or change an entry " & _
            "the old entry is put into the 'Garbage Grid'." & _
            vbCr & vbCr & _
            "When you click on 'Restore Entry' or press Ctrl-R, " & _
            "the Garbage Grid is displayed. You can double-click " & _
            "on an entry or click Ctrl-E to edit it and the entry " & _
            "will be copied into the Input Boxes from where you can " & _
            "save it back into the main database." & _
            vbCr & vbCr & _
            "As a result, you never have to be nervous about deleting " & _
            "or changing entries, since you can always recover the " & _
            "original." & vbCr & vbCr & _
            "Click the Details button for more information."
    L_Tutorial.Height = 3500
    int_GridNum = 0
    ColorGrids 0
    ResizeTutorial
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#E2"
    Wait

    If Grid1(2).Visible Then
        b_Continue.Visible = True
        mRestoreEntry_Click
        b_Continue.Visible = False
    End If
End Sub


Private Sub mSaveSystem_Click()
    str_Response = MyMsgBox("Not implemented yet.", 1, "ok", "Save System")
End Sub


Private Sub mSetColors_Click()
    If f_GridColors.Visible Then
        f_GridColors.Visible = False
        mSetColors.Checked = False
        If frm_Tutorial.Visible And _
           frm_Tutorial.Caption = "Set Screen Colors" _
        Then
            b_Continue_Click
        End If
    Else
        f_GridColors.Left = 2222
        f_GridColors.Top = 2222
        f_GridColors.Visible = True
        f_GridColors.ZOrder 0
        mSetColors.Checked = True
        If frm_Tutorial.Visible And _
           b_Continue.Visible = False _
        Then
            Run_SetColors_Tutorial
        End If
    End If
End Sub


Private Sub Run_SetColors_Tutorial()
    frm_Tutorial.Caption = "Set Screen Colors"
    L_Tutorial = "Self-explanatory. Just click on the different drop-down " & _
                 "lists and see the colors that result. Exit when you " & _
                 "find colors you like by clicking the 'X'." & _
                 vbCr & vbCr & _
                 "Your selected colors will be saved if you exit the " & _
                 "program normally (i.e.: don't crash) and will be " & _
                 "used the next time you run the program."
    b_Continue.Visible = True
    frm_Tutorial.Move f_GridColors.Left + f_GridColors.Width + 40, 2222
    L_Tutorial.Height = 2800
    ResizeTutorial
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#V10"
End Sub


Private Sub mMatchWidths_Click()
    If frm_Tutorial.Visible Then
        If b_Continue.Visible = False Then
            Run_MatchCols_Tutorial
            Exit Sub
        End If
    End If

    AdjustColumns
End Sub


Private Sub Run_MatchCols_Tutorial()
    frm_Tutorial.Caption = "Match Column Widths"
    L_Tutorial = "You can change the widths of columns in the grids by " & _
               "placing the cursor to the right of a column heading. " & _
               "The cursor will change shape to a vertical bar with " & _
               "arrows pointing left and right. (Try it now.)" & _
               vbCr & vbCr & _
               "Click and hold the left mouse button while you move " & _
               "the cursor to the left or right to shrink or enlarge " & _
               "the column." & _
               vbCr & vbCr & _
               "If you have two grids displayed, you can copy the " & _
               "column widths of the top grid to the bottom by selecting " & _
               "'View - Match column widths' in the menu or pressing Ctrl-M."
    L_Tutorial.Height = 3250
    b_Continue.Visible = True
    ResizeTutorial
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#V5"
End Sub


Private Sub cmb_Alert_GotFocus()
    str_OrigData = cmb_Alert.Text
    L_Status.Clear
    L_Status.AddItem " Indicate what type of disclosure this is, if any."
End Sub


Private Sub cmb_Alert_Validate(Cancel As Boolean)
    L_Status.Clear
End Sub


Private Sub cmb_Bid_GotFocus()
    str_OrigData = cmb_Bid.Text
    L_Status.Clear
    L_Status.AddItem " Bid to be made when hand matches this entry's specs.  Ace Asking examples: 4N@30 where 3=H=trumps and 0=2nd suit if needed. 4C/4D@Mw=Minorwood."
    L_Status.AddItem " Control: 5H@CS where 5H is the bid, C=control signal, S=Spades=trumps.  GSF: 5N@G1 where G=GSF signal, 1=Clubs=trumps."
End Sub


Private Sub cmb_Bid_Validate(Cancel As Boolean)
    Dim f As String
    Dim p As String
    Dim a As String
    Dim xx As Integer

    L_Status.Clear
    If str_OrigData <> cmb_Bid.Text And _
       Not int_NoPaint _
    Then
        If ck_SaveNew.Value = 0 Then
            If Trim$(t_Prior) <> "" Then
                f = Trim$(t_Prior) & "-" & str_OrigData
            Else
                f = str_OrigData
            End If
            If f = "" Then
                Exit Sub
            End If

            p = str_BBpath & "\Notes\"
            On Local Error Resume Next
            f = f & "_ID" & rs_BidTable!id
            On Local Error GoTo 0
            If Dir$(p + f & ".txt") <> "" Then
                a = "A disclosure note filename uses the original bid " & _
                    str_OrigData & "." & vbCr & _
                    "Do you wish to change the filename?" & vbCr & _
                    "(Other entries which may reference this filename" & vbCr & _
                    " will not automatically be changed.)"

                str_Response = MyMsgBox(a, 1, "YNC", "BidBase")
                Select Case str_Response
                Case Is = "cancel"
                    int_NoPaint = True
                    cmb_Bid = str_OrigData
                    int_NoPaint = False

                Case Is = "yes"
                    Name p & f & ".txt" As _
                         p & t_Prior & "-" & cmb_Bid & "_ID" & rs_BidTable!id & ".txt"
                End Select
            End If

            If 1 = 2 And ck_Saved.Value = 1 Then
                ck_Saved.Value = 0
                str_Response = MyMsgBox("Save this as a new entry?", 1, "YN", "BidBase")
                If str_Response = "yes" Then
                    ck_SaveNew.Value = 1
                End If
            End If
        End If
    End If

    If InStr(cmb_Bid.Text, "@") > 0 Then
        ' Slam exploring bid made.  '
        ' Increase cmb_Bid.Size and '
        ' location to make room.    '
        If cmb_Bid.Width <> 1015 Then
            t_Prior.Width = 3380
            t_Order.Left = 4990
            cmb_Bid.Left = 5835
            cmb_Bid.Width = 1050
        End If
    ElseIf cmb_Bid.Width = 1050 Then
        t_Prior.Width = 3765
        t_Order.Left = 5375
        cmb_Bid.Left = 6240
        cmb_Bid.Width = 650
    End If

End Sub


Private Sub cmb_StrDiscl_Validate(Cancel As Boolean)
    L_Status.Clear
    If str_OrigData <> cmb_StrDiscl.Text Then
        ck_Saved.Value = 0
    End If
End Sub


Private Sub cmb_Position_GotFocus()
    str_OrigData = cmb_Position.Text
    L_Status.Clear
    L_Status.AddItem " First or 2nd seat,  Third or 4th seat,  1st seat,  2nd seat,  3rd seat,  4th seat, "
    L_Status.AddItem " Y=Passed bidder,  N=Unpassed bidder,  P=responder is a passed hand,  U=responder is an unpassed hand."
End Sub


Private Sub cmb_Position_Validate(Cancel As Boolean)
    L_Status.Clear
    If str_OrigData <> cmb_Position.Text Then
        ck_Saved.Value = 0
    End If
End Sub


Private Sub cmb_Shape_GotFocus()
    str_OrigData = cmb_Shape.Text
    L_Status.Clear
    L_Status.AddItem " 1-suiter (6+/3-),  2-suiter (6+/4+ or 5+/5+),  3-suiter (4/4/4/1 or 5/4/4/0),"
    L_Status.AddItem " b=bal.(4/3/3/3, 4/4/3/2 or 5/3/3/2),  u=unbal.,  s=singleton,  v=void,  e=either sing/void."
End Sub


Private Sub cmb_Shape_Validate(Cancel As Boolean)
    L_Status.Clear
    If str_OrigData <> cmb_Shape.Text Then
        ck_Saved.Value = 0
    End If
End Sub


Private Sub cmb_Vul_GotFocus()
    str_OrigData = cmb_Vul.Text
    L_Status.Clear
    L_Status.AddItem " U = Unfavorable vul.  F = Favorable vul.  B = Both vul.  N = Neither vul.  V = bidder is Vul,  X = not vul, [blank]=any."
End Sub


Private Sub cmb_Vul_Validate(Cancel As Boolean)
    L_Status.Clear
    If str_OrigData <> cmb_Vul.Text Then
        ck_Saved.Value = 0
    End If
End Sub


Private Sub Form_Load()
    Dim i As Integer
    Dim e As String

    str_Job = "Startup"

    For i = 0 To 3
        Data1(i).ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\BidBase.accdb;Mode=ReadWrite;Persist Security Info=False"
    Next

    'If InStr(Dir$("C:\Programming\VB98\Bridge\BidBase Editor debug.*"), "debug") > 0 Then
        DebugMode = True
    'End If

    Dim sWinhwnd As String
    Dim tmp As String

    If DebugMode And Grid1(0).Columns.Count <> 51 Then
        MsgBox "Grid1(0).Columns.Count should be 51." & vbCr & _
                "Actual value = " & Grid1(0).Columns.Count & _
                "Program cannot run."
        Form_Unload 0
    End If

    e = "Checking for previous instance of program."
    If App.PrevInstance = True Then
        MsgBox "BidBase Editor is already running." & vbCr & _
                "If you don't see it on the task bar," & vbCr & _
                "It may still be open in Task Manager -" & vbCr & _
                "Background processes. Right click" & vbCr & _
                "and click End Process."
        End
    End If

    e = "Setting str_BBpath to " & App.Path
    str_BBpath = App.Path
    str_DBPathAndName = App.Path & "\BidBase.accdb"

    ChDir str_BBpath

    If Dir$(str_DBPathAndName) = "" Then
        MsgBox "The BidBase database (BidBase.accdb) is not in" & vbCr & _
               str_BBpath & " - Quitting."
        Form_Unload 0
    End If

    ' Subclass the datagrid control to allow scrolling it with a mouse wheel. '
    ' The scrolling code is in the class MouseWheelGridScroller.              '
    'e = "Subclassing the Grids to allow scrolling with mouse."
    'For i = 0 To 1
    '    lpPrevWndProc = SetWindowLong(Grid1(i).hwnd, GWL_WNDPROC, AddressOf WndProc)
    'Next

    Me.MousePointer = vbHourglass
    Me.ScaleMode = 1
    bln_StartUp = True

If DebugMode Then
MsgBox "Status to this point in Form_Load:" & vbCr & _
        "  Path to the programs:" & vbCr & _
        "    " & str_BBpath & vbCr & _
        "  Path to BidBase.accdb:" & vbCr & _
        "    " & str_DBPathAndName & vbCr & _
        "  Version is " & Version$ & vbCr & _
        "The configuration file loads next."
Me.Show
End If

    Me.Caption = Me.Caption & "      ver." & Version$

    LoadConfig

    InitializeDatabase

    GetCommandLine

    e = "Resizing form."
    Form_Resize
    e = "Finishing touches."
    DoEvents
    MousePointer = 0
    Timer1.Interval = 20
    'Timer2_Timer ' Time and battery charge display not used now. '
    DoEvents
    Me.Show
    Call ShowScrollBar(L_Status.hwnd, SB_VERT, 0&)
    If mTwoGrids.Checked = True Then
        mTwoGrids_Click
    End If
        
    Me.Refresh
    'mViewAll_Click
mEnter_Click
    b_LoadHands_Click
    On Local Error GoTo 0
    bln_Abort = False
    str_Job = ""
    Grid1(0).AllowRowSizing = False
    Grid1(1).AllowRowSizing = False
Exit Sub

FormLoadError:
    MsgBox "Error -- " & e & vbCr & "Closing program."
    Close #1
    Form_Unload 0
End Sub


Private Sub MultiMonitorSupport()
    Dim z As Long
    hMonitor! = GetMonitorByWindow(Me.hwnd)
    SetupMultipleMonitors
    Call EnumDisplayMonitors(0&, 0&, AddressOf MonitorEnumProc, z)
End Sub


Private Sub InitializeDatabase()
    Dim i As Integer
    Dim j As Integer

    Set ws_BidBase = DBEngine.Workspaces(0)
    Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
    Set rs_BidNamesTbl = db_BidBase.OpenRecordset(str_BidNamesTbl)
    Set rs_DelTable = db_BidBase.OpenRecordset(str_GarbageTbl)
    Set rs_DupePBs = db_BidBase.OpenRecordset(str_DupePBsTbl)
    Set rs_BidsQuery = db_BidBase.OpenRecordset(str_DataSource)
    Set rs_BidTable = db_BidBase.OpenRecordset(str_BidsTbl)
    Set rs_HandsData = db_BidBase.OpenRecordset("HandData")

    ' qSel is used when the grids are being changed to show '
    ' different prior bids and/or sorted differently. '
    ' Bids.PB_Sort_Order, was before Bids.Prior_Bids '
    qSel = "SELECT BidNames.Selected, BidNames.SubCat_Selected, " & _
                  "Bids.Pct_Used, Bids.RoundOfBidding, " & _
                  "Bids.Player, Bids.Prior_Bids, " & _
                  "Bids.Bid_Order_Number, Bids.Bid, BidNames.Bid_Name, " & _
                  "Bids.HCP, Bids.Dist_Pts, " & _
                  "Bids.Cust_Pts, Bids.A_Cust, Bids.K_Cust, " & _
                  "Bids.Q_Cust, Bids.J_Cust, Bids.T_Cust, Bids.Winrs, " & _
                  "Bids.Losers, Bids.QTs, Bids.Stopped_Suits, " & _
                  "Bids.Shape, Bids.Intermed, Bids.xNum, Bids.xHCP, " & _
                  "Bids.sNum, Bids.sHCP, Bids.hNum, Bids.hHCP, " & _
                  "Bids.dNum, Bids.dHCP, Bids.cNum, Bids.cHCP, " & _
                  "Bids.cmNum, Bids.cmHCP, Bids.Vul, Bids.Pos, " & _
                  "Bids.Artificial, Bids.Alert, Bids.Explanation, " & _
                  "Bids.Test_Hand, Bids.Asks, Bids.DateOfEntry, " & _
                  "Bids.Best_Contract, Bids.Num_Used, Bids.Source, " & _
                  "Bids.LHO, Bids.Pard, Bids.RHO, Bids.ID " & _
           "FROM BidNames RIGHT JOIN Bids ON BidNames.IDnum = Bids.ConvID "

    ' Fields not shown in grids: PB_Sort_Order, ConvID, AgainstID

    LoadBidNames

    On Local Error GoTo ColNameError
    For i = 0 To int_LastBidCol
        For j = 0 To 1
            cmb_Show(j).AddItem Grid1(0).Columns(i).DataField
        Next

        Select Case Grid1(0).Columns(i).DataField
        Case Is = "Pct_Used"
            int_Pct_Col = i
        Case Is = "RoundOfBidding"
            int_Round_Col = i
        Case Is = "Player"
            int_Player_Col = i
        Case Is = "PB_Sort_Order"
            int_PB_Sort_Ord_Col = i
        Case Is = "Prior_Bids"
            int_Prior_Bids_Col = i
        Case Is = "Bid_Order_Number"
            int_OrdNum_Col = i
        Case Is = "Bid"
            int_Bid_Col = i
        Case Is = "Bid_Name"
            int_Bid_Name_Col = i
        Case Is = "HCP"
            int_HCP_Col = i
        Case Is = "xNum"
            int_xNum_Col = i
        Case Is = "cmNum"
            int_Combo_Col = i
        Case Is = "Vul"
            int_Vul_Col = i
        Case Is = "Explanation"
            int_Expl_Col = i
        Case Is = "Test_Hand"
            int_TestHand_Col = i
        Case Is = "Source"
            int_Source_Col = i
        End Select
    Next

ColNameError:
    If i < int_LastBidCol + 1 Then
        MsgBox "Error setting column name variables for column " & i & "."
    End If
    On Local Error GoTo 0
End Sub


Private Sub GetCommandLine()
    ' BidBase Editor is an EXE. In addition to starting '
    ' it from an icon, it can be called by any program  '
    ' passing it (1) a hand and prior bids or (2) an ID '
    ' number for a bid to be edited. In the first case, '
    ' the calling program will send two parameters vs.  '
    ' only one parameter for the second case.           '
    '                                                   '
    ' Ex.: 1. BidBase Editor.exe AK432-QJT3-A2-K2, P-P  '
    '      2. BidBase Editor.exe 20134                  '

    Dim CmdLine As String
    Dim CmdLnLen As Integer
    Dim c As Integer

    Erase str_CmdLineArg()

    'Get command line arguments.
    CmdLine = Trim$(Command())
    CmdLnLen = Len(CmdLine)
    If CmdLnLen = 0 Then Exit Sub

    c = InStr(CmdLine, ",")
    If c = 0 Then
        ' One parameter '
        str_CmdLineArg(1) = CmdLine
    Else
        ' Two parameters '
        If c < 16 Or c > CmdLnLen - 2 Then
            ' Comma is out of place
            str_Response = MyMsgBox("Invalid parameters: " & vbCr _
                                    & CmdLine, 2, "ok", "Get Command Line")
        Else
            str_CmdLineArg(1) = Left$(CmdLine, c - 1)
            str_CmdLineArg(2) = Mid$(CmdLine, c + 1)
            If str_CmdLineArg(2) = "none" Then
                str_CmdLineArg(2) = ""
            End If
        End If
    End If

End Sub


Private Sub LoadConfig()
    Dim a As String
    Dim b As String
    Dim c As String
    Dim f As String
    Dim FiNum As Integer
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim m As Long
    Dim n As Long
    Dim x As String
    Dim split0 As Integer
    Dim split1 As Integer
    Dim tgls As Integer ' = top grid left split
    Dim tgrs As Integer
    Dim bgls As Integer
    Dim bgrs As Integer
    Dim fx As Long
    Dim fy As Long
    Dim ng As Integer
    Dim df As String
    Dim cw As Single
    Dim Msg As String
    Dim s As String

    ' To make it easier if you have to manually fix up the  '
    ' configuration file, descriptive captions are included '
    ' in it. In the code below, "Input #fiNum, x" reads in  '
    ' a caption, but it is then ignored. To the right of the'
    ' "x" is what the caption should be. If you are single- '
    ' stepping thru the code and the caption doesn't match, '
    ' then the file has gotten messed up somehow. If the    '
    ' BBEditor.cfg file does get messed up, just delete it  '
    ' and the program will read the BBEditor-Starting.cfg   '
    ' file instead, but on exit, will create BBEditor.cfg.  '

    str_Job = "Loading Cfg"
    str_SuitName(1) = "Club"
    str_SuitName(2) = "Diamonds"
    str_SuitName(3) = "Hearts"
    str_SuitName(4) = "Spades"

    ' Check if new user: '
'    On Local Error GoTo CfgError
    If Dir$(str_BBpath & "\BBEditor.cfg") = "" Then
        bln_FirstTime = True
        f = "\BBEditor-Starting.cfg"
    Else
        f = "\BBEditor.cfg"
    End If

If DebugMode Then
    f = "\BBEditor-Starting.cfg"
    Msg = "Loading " & str_BBpath & f
End If

LoadCfg:
    If Dir$(str_BBpath & f) <> "" Then
        FiNum = FreeFile
        Open str_BBpath & f For Input As #FiNum
        Input #FiNum, x ' "==Database and table names:" '
        If x <> "==Database and table names:" Then
            MsgBox "Error 1 loading cfg."
        End If
        Input #FiNum, str_Database
        Input #FiNum, str_BidsTbl
        Input #FiNum, str_BidNamesTbl
        Input #FiNum, str_GarbageTbl
        Input #FiNum, str_DupePBsTbl
        Input #FiNum, str_DataSource
        Msg = Msg & vbCr & "Loading " & x & vbCr & "          " & _
                str_Database & "  " & str_BidsTbl & "  " & _
                str_BidNamesTbl & "  " & str_GarbageTbl & "  " & _
                str_DupePBsTbl & "  " & str_DataSource

        Input #FiNum, x ' "==# of columns:"
        ' (Actual #-1, since start with 0.) '
        Input #FiNum, int_LastBidCol '  <--------------------- int_LastBidCol is used to control input of column data further down.
        int_ID_Col = int_LastBidCol ' In case ID column changes. Would like to move it to start to make it easier to find. '
        Msg = Msg & vbCr & "Loading " & x & " (index, starting with 0) = " & int_LastBidCol

        Input #FiNum, x ' "==Grid display entry# ranges:" '
        Msg = Msg & vbCr & "Loading " & x & vbCr & "       "
        For i = 0 To 1
            Input #FiNum, a, b, c
            Msg = Msg & "   grid #" & i & " = " & a & ", " & b & ", " & c
            t_Show(0 + 2 * i) = a
            t_Show(1 + 2 * i) = b
            j = 0
            Do While cmb_Show(i).List(j) <> c And _
                     i < cmb_Show(i).ListCount - 1
                j = j + 1
                DoEvents
            Loop
            cmb_Show(i) = c
        Next

        Input #FiNum, x ' "==Columns to sort on:" '
        Msg = Msg & vbCr & "Loading " & x & " = "
        For i = 0 To 4
            Input #FiNum, int_SortIndex(i)
            If i < 4 Then
                Msg = Msg & int_SortIndex(i) & ", "
            Else
                Msg = Msg & int_SortIndex(i)
            End If
            cmb_Sort(i).ListIndex = int_SortIndex(i)
        Next

        Line Input #FiNum, x ' "==Form top, left, width, height:" '
        Input #FiNum, j, k, m, n
        Msg = Msg & vbCr & "Loading " & x & " = " & j & ", " & k & ", " & m & ", " & n
        Me.Top = j
        Me.Left = k
        Me.Width = m
        Me.Height = n

        Line Input #FiNum, x ' "==# of grids:" '
        Input #FiNum, ng
        Msg = Msg & vbCr & "Loading " & x & " = " & ng

        Line Input #FiNum, x ' "==Active grid # and bid number:" '
        Input #FiNum, int_GridNum, str_StartGridRow
        Msg = Msg & vbCr & "Loading " & x & " = " & int_GridNum & ", " & str_StartGridRow

        Line Input #FiNum, x ' "==Grids split? (1st, then 2nd):" '
        Input #FiNum, split0, split1
        Msg = Msg & vbCr & "Loading " & x & " = " & split0 & ", " & split1
        mSplit(0).Checked = split0
        mSplit(1).Checked = split1
        Msg = Msg & vbCr & "Grid1(0).Columns.Count = " & Grid1(0).Columns.Count
        ' Up above was loaded int_LastBidCol = 47 but Grid1(0).Columns.Count = 51 ??? '
        ' Later, the ID numbers are in column 50, so use 50. '
        
If DebugMode Then
        MsgBox "Here is what the has been loaded to this point:" & vbCr & Msg
End If
        Line Input #FiNum, x ' "==Column data:" '
        Line Input #FiNum, x ' ==Visible? (top grid left, right, bottom grid left, right).  Caption, data field, width.)

        ' Set the left edge of grids: '
        ' This is not in the cfg file, I'm just setting them to 0.
        For i = 0 To 4
            Grid1(i).Left = 0
        Next
        Grid1(0).Top = 0

        i = 0
        k = vbNo

        ' If you stop and enter "Me.Show" in the debug window right now,
        ' it will show the grid with all the columns and data.
        ' The BBEditor.cfg file shows entries such as
        '     -1,0,-1,0,"Cnv","Selected",329.9528
        ' which is whether or not the column is visible in each of the grids,
        ' the field name of the column in dBase.Bids, the column name in
        ' the grids of BBEditor, and the column width. So the purpose of
        ' this Do-Loop seems to be to verify the grid column captions.
        
        Do
            DoEvents
            Msg = "Loading column data for" & vbCr
            Input #FiNum, tgls, tgrs, bgls, bgrs, c, df, cw
            tgls = True ' always? '
'Debug.Print i; c
            ' tgls = column visible in top grid, left split?    '
            ' tgrs = column visible in top grid, right split?   '
            ' bgls = column visible in bottom grid, left split? '
            ' bgls = column visible in bottom grid, left split? '

            sng_ColWidths(i) = cw

                  Msg = Msg & "Column # " & i & vbCr _
                  & "  Col visible in top grid, left split? = " & tgls & vbCr _
                  & "  Col visible in top grid, right split? = " & tgrs & vbCr _
                  & "  Col visible in bot.grid, left split? = " & bgls & vbCr _
                  & "  Col visible in bot.grid, right split? = " & bgrs & vbCr _
                  & "  Grid Column heading = " & c & vbCr _
                  & "  Database field name = " & df & vbCr _
                  & "  Column width = " & cw
                  
                  ' Column 5 is PB Sort Order which is not shown on the grid  '
                  ' because of a column width of 0. This lets us look up the  '
                  ' prior bid sort order without having to see it on the grid '
                  ' although I don't remember why we don't want to see it.    '

            If k = vbNo And DebugMode Then
                If i > 2 Then
                    k = MsgBox(Msg & vbCr & "Skip this from now on?" & vbCr & s, vbYesNo)
                    MsgBox Msg
                End If
            End If
            
            If i > Grid1(0).Columns.Count - 1 Then
                MsgBox "Error reading column " & i & vbCr & _
                       "Grid1(0).Columns.Count = " & Grid1(0).Columns.Count
                GoTo CfgError
            Else
                'For j = 0 To 4
                '    Grid1(j).Columns(i).Caption = c
                '    Grid1(j).Columns(i).Width = cw
                'Next
            End If

            sng_ColWidths(i) = cw
                     '    tgls = top grid left split, etc.   '
                     ' Note: I tried Dim-ing tgls as Boolean '
                     ' but it didn't read in correctly, so I '
                     ' had to change them back to Integers.  '

            If tgls = False Then
                cmb_HiddenCols(0).AddItem _
                    Right$("0" & i, 2) & " - " & _
                    Grid1(0).Columns(i).Caption
                Grid1(0).Splits(Grid1(0).Split).Columns(i).Visible = False
            End If

            If bgls = False Then
                cmb_HiddenCols(2).AddItem _
                    Right$("0" & i, 2) & " - " & _
                    Grid1(1).Columns(i).Caption
                Grid1(1).Splits(Grid1(1).Split).Columns(i).Visible = False
            End If

            If split0 Then
                b = Grid1(0).Columns(i).Caption
                Grid1(0).Splits(1).Columns(i).Width = sng_ColWidths(i)
                Grid1(0).Splits(1).Columns(i).Visible = tgrs
                If tgrs = False Then
                    cmb_HiddenCols(1).AddItem _
                        Right$("0" & i, 2) & " - " & _
                        b
                End If
            ElseIf tgrs = False Then
                cmb_HiddenCols(1).AddItem _
                    Right$("0" & i, 2) & " - " & _
                    Grid1(0).Columns(i).Caption
            End If

            If split1 And bgrs > 0 Then
                Grid1(1).Splits(1).Columns(i).Visible = bgrs
                Grid1(1).Splits(1).Columns(i).Width = sng_ColWidths(i)
                If bgrs = False Then
                    cmb_HiddenCols(3).AddItem _
                        Right$("0" & i, 2) & " - " & _
                        Grid1(1).Columns(i).Caption
                End If
            ElseIf bgrs = False Then
                cmb_HiddenCols(3).AddItem _
                    Right$("0" & i, 2) & " - " & _
                    Grid1(1).Columns(i).Caption
            End If
Debug.Print i; Grid1(1).Columns(i).Caption
            If Grid1(1).Columns(i).Caption = "ID" Then ' ID is Columns(46) '
                int_LastBidCol = i
                Exit Do
            End If
            i = i + 1
            DoEvents
        Loop

        'i = 50
        Input #FiNum, lng_TopID(0)
        Input #FiNum, lng_TopID(1)

        Input #FiNum, x
        Msg = Msg & vbCr & "Loading " & x
        Me.Show
        If x <> "==cmb_ForeColors, cmb_BackColors list indexes:" Then
            str_Response = MyMsgBox("Screen colors heading is out of" & vbCr & _
                                    "place in the configuration file.", 2, "ok", "")
        End If
        Input #FiNum, lng_FGGrid1, int_BGGrid1
        Input #FiNum, lng_FGGrid2, int_BGGrid2
        Input #FiNum, int_FGInputs, int_BGInputs
        Input #FiNum, int_FGForm, int_BGForm

        Input #FiNum, x
        If x <> "==Find Bid window coordinates:" Then
            str_Response = MyMsgBox("Bid window coordinates heading out of" & vbCr & _
                                    "place in the configuration file.", 2, "ok", "")
        End If
        Input #FiNum, fx, fy
        f_EnterHand.Move fx, fy

        Input #FiNum, x  ' "==Compare Grids Input window coords:"
        
        Msg = Msg & vbCr & "Loading " & x
        Input #FiNum, fx, fy
        f_Compare.Move fx, fy
        
        Input #FiNum, x
        If InStr(x, "==References:") = 0 Then
            str_Response = MyMsgBox("List of References out of" & vbCr & "place in the configuration file.", 2, "ok", "")
        Else
            Input #FiNum, n ' number of references '
            For m = 0 To n - 1
                Input #FiNum, a
                list_Refs.AddItem a
            Next
        End If
        
Msg = Msg & vbCr & "Loading configuration completed."
        Close #FiNum
        If DebugMode Then
            MsgBox Msg
        End If

        On Local Error Resume Next
        int_GridNum = 0
        cmb_ForeColors(0).ListIndex = lng_FGGrid1
        cmb_BackColors(0).ListIndex = int_BGGrid1
        cmb_ForeColors(1).ListIndex = lng_FGGrid2
        cmb_BackColors(1).ListIndex = int_BGGrid2
        cmb_ForeColors(2).ListIndex = int_FGInputs
        cmb_BackColors(2).ListIndex = int_BGInputs
        cmb_ForeColors(3).ListIndex = int_FGForm
        cmb_BackColors(3).ListIndex = int_BGForm
    End If

    HScroll1.Max = 100
    HScroll1.Min = Me.Width - 18700
    HScroll1.SmallChange = 2000
    HScroll1.LargeChange = 12000

    If Grid1(0).Columns(int_LastBidCol).Caption <> "ID" Then
        If Grid1(0).Columns(int_LastBidCol).Caption = "ID" Then
            int_LastBidCol = Grid1(0).Columns.Count - 1
        Else
            MsgBox "Last grid column should be 'ID#'." & vbCr & _
                    "Some features may not run correctly."
        End If
    End If
    bln_InputBoxes = True

    bln_TwoGrids = (ng = 1)
    'mTwoGrids_Click

    For i = 1 To 9
        cmb_StrDiscl.AddItem Mid$(" BWIOPSVX", i, 1)
    Next
'Exit Sub ' Not using the Other Hands grid at this time. '

    g_OtherHands.BackColorFixed = &H8000&
    g_OtherHands.ForeColorFixed = vbYellow ' &HC0C0&
    g_OtherHands.BackColor = vbBlack
    g_OtherHands.ForeColor = vbYellow

    g_OtherHands.Col = 0
    g_OtherHands.Top = 100
    f_OtherPlayers.Width = 5000
    f_OtherPlayers.Left = 19700
    f_OtherPlayers.Top = 9000
    f_OtherPlayers.Height = 4500
    g_OtherHands.Height = 4400
    For i = 1 To 4
        g_OtherHands.ColAlignment(i) = flexAlignCenterTop
    Next
    SetupHandDataGrid

If DebugMode Then MsgBox "No errors loading " & f

Exit Sub

CfgError:
    Close #FiNum
    If f = "\BBEditor-Starting.cfg" Then
        MsgBox "Configuration file error after: " & vbCr & Msg

        Form_Unload 0
    Else
        f = "\BBEditor-Starting.cfg"
        Resume LoadCfg
    End If
End Sub


Private Sub SetupHandDataGrid()
    Dim i As Integer
    Dim c As Integer
    Dim r As Integer
    
    With g_OtherHands
    '''''''''''''''''
    .AllowBigSelection = False
    .Col = 0
    For i = 1 To 17
        .Row = i
        .CellFontBold = True
        If i = 1 Then
            .Text = "ID#"
        ElseIf i = 2 Then
            .Text = "HCP"
        ElseIf i = 3 Then
            .Text = "Pts."
        ElseIf i = 4 Then
            .Text = "As/Ks"
        ElseIf i = 5 Then
            .Text = "Shape"
        ElseIf i = 6 Then
            .Text = "S Qty"
        ElseIf i = 7 Then
            .Text = "S HCP"
        ElseIf i = 8 Then
            .Text = "H Qty"
        ElseIf i = 9 Then
            .Text = "H HCP"
        ElseIf i = 10 Then
            .Text = "D Qty"
        ElseIf i = 11 Then
            .Text = "D HCP"
        ElseIf i = 12 Then
            .Text = "C Qty"
        ElseIf i = 13 Then
            .Text = "C HCP"
        ElseIf i = 14 Then
            .Text = "CmbNum"
        ElseIf i = 15 Then
            .Text = "CmbHCP"
        ElseIf i = 16 Then
            .Row = i
            .Col = 0
            .CellAlignment = 3
            .Text = "SAVE ALL"
            .CellBackColor = vbWhite
            .CellForeColor = 32768
            For c = 1 To 4
                .Col = c
                .CellFontBold = True
                .CellFontSize = 8.25
                .Text = "SAVE^"
                .CellBackColor = vbWhite
                .CellForeColor = 32768
            Next
        
            .Row = 17
            .Col = 0
            .CellFontBold = True
            .CellFontSize = 8.25
            .Text = "DEL. ALL"
            .CellAlignment = 3
            .CellBackColor = vbWhite
            .CellForeColor = vbRed
            For c = 1 To 4
                .Col = c
                .CellFontBold = True
                .CellFontSize = 8.25
                .Text = "DELETE^"
                .CellBackColor = vbWhite
                .CellForeColor = vbRed
            Next
        End If
    Next
    
    .Row = 0
    .Col = 1
    .Text = "Me"
    .CellFontBold = True
    .Col = 2
    .Text = "LHO"
    .CellFontBold = True
    .Col = 3
    .Text = "Pard"
    .CellFontBold = True
    .Col = 4
    .Text = "RHO"
    .CellFontBold = True
    ''''''''
    End With
    t_OtherHands.Visible = False
End Sub


Private Sub Form_Resize()
    Dim workAreaHeight As Long
    Dim ng As Integer
    Dim gn As Integer

    If Me.Visible = False Or _
       Me.ScaleWidth = 0 Or bln_SkipIt _
    Then
        Exit Sub
    End If

    On Local Error Resume Next

    If bln_InputBoxes Then
        frame0(6).Top = Me.Height - frame0(6).Height - 950 ' Entry input boxes at bottom of form. '
        workAreaHeight = Me.Height - 950 - frame0(6).Height
    Else
        workAreaHeight = Me.Height - 950
    End If

    bln_TwoGrids = Grid1(0).Visible And Grid1(1).Visible

    If bln_TwoGrids Or _
       bln_ViewGarbage Or _
       mFourGrids.Checked _
    Then
        If mFourGrids.Checked Then
            ng = 4 'number of grids
        Else
            ng = 2
        End If

        If bln_TwoGrids Or _
           mFourGrids.Checked _
        Then
            gn = 1 '2nd bid grid
        Else
            gn = 2 'garbage grid
        End If

        Grid1(0).Height = (workAreaHeight - pSplitterBar.Height) / ng - 10
        Grid1(gn).Height = Grid1(0).Height

        If bln_ViewGarbage Then
            Grid1(0).Width = Me.Width - 320
        Else
            Grid1(0).Width = Me.Width - VScroll1.Width - 225
        End If

        Grid1(gn).Width = Grid1(0).Width
        pSplitterBar.Visible = Grid1(gn).Visible
        pSplitterBar.Top = Grid1(0).Top + Grid1(0).Height
        pSplitterBar.Width = Grid1(0).Width
        pSplitterBar.Left = Grid1(0).Left
        If pSplitterBar.Visible Then
            Grid1(gn).Top = pSplitterBar.Top + pSplitterBar.Height + 15
        Else
            Grid1(gn).Top = Grid1(0).Top + Grid1(0).Height + 80
        End If

        If mFourGrids.Checked Then
            Grid1(3).Height = Grid1(0).Height
            Grid1(3).Width = Grid1(0).Width
            Grid1(3).Top = Grid1(gn).Top + Grid1(gn).Height + 20
            Grid1(4).Height = Grid1(0).Height
            Grid1(4).Width = Grid1(0).Width
            Grid1(4).Top = Grid1(3).Top + Grid1(3).Height + 20
        End If

        If bln_TwoGrids Or _
           mFourGrids.Checked _
        Then
            VScroll1.Top = Grid1(0).Top
            VScroll1.Left = Grid1(0).Left + Grid1(0).Width
            VScroll1.Height = workAreaHeight
            VScroll1.Visible = True
            VScroll1.ZOrder 0
        End If

    Else
        Grid1(0).Width = Me.Width - 225
        Grid1(0).Height = workAreaHeight
        If bln_TwoGrids Or mFourGrids.Checked Then
            pSplitterBar.Top = Grid1(0).Top + Grid1(0).Height - 225
            pSplitterBar.Width = Grid1(0).Width
        Else
            pSplitterBar.Visible = False
        End If
    End If

    HScroll1.Width = Me.Width - 225
    frame0(6).Width = 24000
    'If HScroll1.Width > L_Clock.Left + L_Clock.Width Then
    '    frame0(6).Height = 4700
    '    HScroll1.Visible = False
    'Else
    '    HScroll1.Max = 100
    '    HScroll1.Min = Me.Width - 15000
    '    frame0(6).Height = 4200
    '    HScroll1.Top = 3910
    '    HScroll1.Visible = True
    '    HScroll1.ZOrder 0
    'End If
    
    f_OtherPlayers.Top = frame0(6).Top + 15
    f_OtherPlayers.Left = L_Status.Width + 300
    f_EnterHand.Top = frame0(6).Top - (f_EnterHand.Height - L_Status.Top) - 100
    'f_EnterHand.Left =
End Sub


Private Sub Form_Unload(Cancel As Integer)

If Not App.PrevInstance Then

    Dim FiNum As Integer
    Dim i As Integer
    Dim j As Integer
    Dim r As String
    Dim r1 As String

    FiNum = FreeFile
    Open str_BBpath & "\BBEditor.cfg" For Output As #FiNum

    Print #FiNum, "==Database and table names:"
    Print #FiNum, str_Database
    Print #FiNum, str_BidsTbl
    Print #FiNum, str_BidNamesTbl
    Print #FiNum, str_GarbageTbl
    Print #FiNum, str_DupePBsTbl
    Print #FiNum, str_DataSource

    Print #FiNum, "==# of columns:"
    Print #FiNum, int_LastBidCol

    Print #FiNum, "==Grid display entry# ranges:"
    For i = 0 To 1
        Write #FiNum, t_Show(0 + i * 2); _
                      t_Show(1 + i * 2); _
                      cmb_Show(i)
    Next

    Print #FiNum, "==Columns to sort on:"
    For i = 0 To 4
        Write #FiNum, cmb_Sort(i).ListIndex
    Next

    Print #FiNum, "==Form top, left, width, height:"
    Write #FiNum, Me.Top; Me.Left; Me.Width; Me.Height

    Print #FiNum, "==# of grids:"
    If mTwoGrids.Checked And Not mDupesGrid(0).Checked Then
        Print #FiNum, 2
    Else
        Print #FiNum, 1
    End If

    Print #FiNum, "==Active grid # and bid number:"
    If Grid1(int_GridNum).Row > -1 Then
         Write #FiNum, int_GridNum, _
                  Grid1(int_GridNum).Columns(int_LastBidCol).Text
    End If

    Print #FiNum, "==Grids split? (1st, then 2nd):"

    i = mSplit(0).Checked
    j = mSplit(1).Checked
    Write #FiNum, i; j

    Dim cv(1, 1, 100) As Integer
    Dim g As Integer
    Dim c As Integer

    For g = 0 To 1
        For c = 0 To int_LastBidCol
            If Grid1(g).Splits(0).Columns(c).Visible Then
                cv(g, 0, c) = -1
            End If

            If g = 0 Then 'If mSplit(g).Checked Then
                ' VB Error in the next line for Splits(1) '
                ' if mSplit(g) is not Checked, so two     '
                ' separate If lines are needed.           '
                'If Grid1(g).Splits(1).Columns(c).Visible Then
                '    cv(g, 1, c) = -1
                'End If
            End If
        Next
    Next

    Print #FiNum, "==Column data: "
    Print #FiNum, "==Visible? (top grid left, right, bottom grid left, right).  Caption, data field, width."
    c = -1
    Do 'For c = 0 To 57 'int_LastBidCol - 1  SEE WHAT int_LastBidCol is '
        c = c + 1
        If Grid1(0).Splits(Grid1(0).Split).Columns(c).Visible = False Then
            ' Make a hidden column's width negative. '
            sng_ColWidths(c) = -sng_ColWidths(c)
        End If

        Write #FiNum, cv(0, 0, c); _
                      cv(0, 1, c); _
                      cv(1, 0, c); _
                      cv(1, 1, c); _
                      Grid1(0).Columns(c).Caption; _
                      Grid1(0).Columns(c).DataField; _
                      sng_ColWidths(c)
        DoEvents
    Loop While Grid1(0).Columns(c).Caption <> "ID"
   ' Write #fiNum, 99, 0, 0, 0, "end of fields", "", 0

    If Grid1(0).Row > -1 Then
        Grid1(0).Row = 0
        Write #FiNum, Val(Grid1(0).Columns(int_LastBidCol).Text)
    Else
        Write #FiNum, 0
    End If

    If Grid1(1).Row > -1 Then
        Grid1(1).Row = 0
        Write #FiNum, Val(Grid1(1).Columns(int_LastBidCol).Text)
    Else
        Write #FiNum, 0
    End If

    Write #FiNum, "==cmb_ForeColors, cmb_BackColors list indexes:"
    For i = 0 To 3
        Write #FiNum, cmb_ForeColors(i).ListIndex; _
                      cmb_BackColors(i).ListIndex
    Next

    Write #FiNum, "==Find Bid window coordinates:"
    Write #FiNum, f_EnterHand.Left, f_EnterHand.Top

    Write #FiNum, "==Compare Grids Input window coords:"
    Write #FiNum, f_Compare.Left, f_Compare.Top
    
    Write #FiNum, "==References:"
    Write #FiNum, list_Refs.ListCount
    For i = 0 To list_Refs.ListCount - 1
        ' Change something like Bul 2021-05-02 to Bul 2021- '
        r = list_Refs.List(i)
        j = InStr(r, "-")
        If j > 0 And j < Len(r) Then
            r = Left$(r, j)
        End If
        If r <> r1 Then
            Write #FiNum, r
            r1 = 1
        End If
    Next

    Close #FiNum
End If

 ' These cause an error:            '
 ' Run-time error '3420':           '
 ' Object invalid or no longer set. '
    'rs_BidNamesTbl.Close
    'rs_DelTable.Close
    'rs_DupePBs.Close
    'rs_BidsQuery.Close
    'rs_BidTable.Close
    'rs_HandDataTbl.Close
    'db_BidBase.Close

    ' This closes without error: '
    ws_BidBase.Close

    If Dir$("C:\Program Files (x86)\CardShark BidBase\myfhandle.txt") <> "" Then
        Kill ("C:\Program Files (x86)\CardShark BidBase\myfhandle.txt")
    End If
    
    Dim Frm As Form

    For Each Frm In Forms
        Unload Frm
        Set Frm = Nothing
    Next
End Sub


Private Sub Grid1_Click(Index As Integer)
    int_GridNum = Index
    ColorGrids Index
    If Index = 0 Then
        If mDupesGrid(0).Checked Then
            Timer1.Interval = 1
        End If
    End If

    If Grid1(Index).Col = -1 Then
        int_GridNum = Index
        mEdit_Click Index
    End If
End Sub


Private Sub AdjustColumns()
    Dim c As Column
    Dim i As Integer

    If mSplit(0).Checked Then
        i = 0
        For Each c In Grid1(0).Splits(1).Columns
            c.Width = sng_ColWidths(i)
            i = i + 1
        Next
    End If

    i = 0
    For Each c In Grid1(1).Splits(0).Columns
        c.Width = sng_ColWidths(i)
        i = i + 1
    Next

    If mSplit(1).Checked Then
        i = 0
        For Each c In Grid1(1).Splits(1).Columns
            c.Width = sng_ColWidths(i)
            i = i + 1
        Next
    End If
End Sub


Private Sub Grid1_ColResize(Index As Integer, _
                            ByVal ColIndex As Integer, _
                            Cancel As Integer)
    sng_ColWidths(ColIndex) = _
       Grid1(Index).Splits(0).Columns(ColIndex).Width
End Sub


Private Sub Grid1_GotFocus(Index As Integer)
    L_Status.Clear
    ColorGrids Index
    L_Status.AddItem " B = Bust (0-5 pts.)  W = Weak (6-10)      I  = Intermediate (10-12)   O = Opening strength (13-15)"
    L_Status.AddItem " P = Plus (15-17)     S = Strong (18-19)   V = Very strong (20-24)     X = eXtremely strong (25+)"
End Sub


Private Sub ColorGrids(ByVal Index)
    Dim i As Integer
    If mFourGrids.Checked Then
        For i = 0 To 4
          If Index = i Then
            Grid1(i).ForeColor = vbYellow 'QBColor(lng_FGGrid1)
            Grid1(i).BackColor = QBColor(int_BGGrid1)
          Else
            Grid1(i).ForeColor = vbYellow 'QBColor(lng_FGGrid2)
            Grid1(i).BackColor = vbBlack 'QBColor(int_BGGrid2)
          End If
        Next
    ElseIf bln_TwoGrids Then
        If Index = 0 Then
            Grid1(0).ForeColor = vbYellow 'QBColor(lng_FGGrid1)
            Grid1(0).BackColor = QBColor(int_BGGrid1)
            Grid1(1).ForeColor = vbYellow 'QBColor(lng_FGGrid2)
            Grid1(1).BackColor = vbBlack ' QBColor(int_BGGrid2)
        Else
            Grid1(1).ForeColor = vbYellow 'QBColor(lng_FGGrid1)
            Grid1(1).BackColor = QBColor(int_BGGrid1)
            Grid1(0).ForeColor = vbYellow 'QBColor(lng_FGGrid2)
            Grid1(0).BackColor = vbBlack 'QBColor(int_BGGrid2)
        End If
    Else
        Grid1(0).ForeColor = vbYellow 'QBColor(lng_FGGrid1)
        Grid1(0).BackColor = vbBlack 'QBColor(int_BGGrid1)
    End If
 End Sub


Private Sub mClear_Click()
    If frm_Tutorial.Visible And _
       b_Continue.Visible = False _
    Then
        int_GridNum = 0
        ColorGrids 0
        Grid1(0).Row = 0
        frm_Tutorial.Visible = False
        mEdit_Click 0
        frm_Tutorial.Visible = True
        Run_ClearInputs_Tutorial
        Exit Sub
    End If

    Dim i As Integer

    ck_SaveNew.Value = 1
    ck_Saved.Value = 0
    Grid1(0).Col = 5 'number
    t_Sel = ""
    t_Pct = ""
    t_Round = ""
    t_Player = ""
    t_Prior = ""
    t_Order = ""
    For i = 0 To 1
        cmb_ConvName(i).ListIndex = 0
    Next

    If cmb_Bid.Width = 1050 Then
        t_Prior.Width = 3765
        t_Order.Left = 5375
        cmb_Bid.Left = 6240
        cmb_Bid.Width = 650
    End If
    cmb_Bid = ""

    cmb_Vul.ListIndex = 0
    cmb_Position.ListIndex = 0
    t_Hand_HCPs = ""
    t_BBPoints = ""
    t_Distrib_Pts = ""
    t_CustPoints = ""
    For i = 0 To 4
        t_CustPts(i) = ""
    Next
    t_Winners = ""
    t_Losers = ""
    t_QTs = ""
    t_Stoppers = ""
    cmb_Shape.ListIndex = 0
    t_Intermediates = ""
    For i = 0 To 5
        t_Qty(i) = ""
        t_HCP(i) = ""
    Next
    t_Disclosure = ""
    t_Ref = ""
    b_Discl.Picture = LoadPicture(str_BBpath & "\ButtonN.bmp")
    b_Conven_Description.Picture = LoadPicture(str_BBpath & "\ButtonN.bmp")
    t_TestHand = ""
    cmb_Alert.ListIndex = 0
    cmb_StrDiscl.ListIndex = 0
    For i = 1 To 4
        t_Len(i) = ""
    Next
    ck_Artificial.Value = 0
    t_AsksFor = ""
    L_Date_Saved = ""
    ck_Continuation.Value = 0
    t_BestContract = ""
    g_OtherHands.Clear
    SetupHandDataGrid
    'If str_DupePrior = "" And mDupesGrid(0).Checked = False Then
    '    For i = 0 To 5
    '        t_DupeEntry(i) = ""
    '        b_DupeEntry(i).Visible = False
    '        'f_Dupes.Visible = False
    '    Next
    'End If
    lng_ID = 0
    For i = 0 To 1
        lng_TopID(i) = 0
        lng_CurID(i) = 0
        int_CurRow(i) = 0
    Next
    'For i = 0 To 5
    '    t_DupeEntry(i) = ""
    'Next
    Grid1(0).Enabled = True
    On Local Error Resume Next
    t_Round.SetFocus
    On Local Error GoTo 0
End Sub


Private Sub Run_ClearInputs_Tutorial()
    frm_Tutorial.Caption = "Clear Input Boxes"
    L_Tutorial = "This is a simple command. It clears the Input Boxes. " & _
               "If you want to make a new entry, press Ctrl-K " & _
               "first." & _
               vbCr & vbCr & _
               "If you are using some other command which copies " & _
               "an existing entry or hand to the Input Boxes, you do " & _
               "not have to clear the boxes first, the other " & _
               "command will take care of that." & _
               vbCr & vbCr & _
               "Before continuting, try pressing Ctrl-E to " & _
               "load an entry into the " & _
               "Input Boxes, then press Ctrl-K to clear it. You can " & _
               "repeatedly press Ctrl-E and Ctrl-K. Clearing the " & _
               "Input Boxes does not remove an entry from the " & _
               "database. You must use 'Edit - Delete' for that."
    L_Tutorial.Height = 3100
    ResizeTutorial
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#E9"
End Sub


Private Sub mDelete_Click()
    Dim x As Integer
    Dim n As Long
    Dim b As String
    Dim i As Integer
    Dim r As Integer
    Dim prBid As String

    prBid = Grid1(int_GridNum).Columns(int_Prior_Bids_Col).Text

    If frm_Tutorial.Visible Then
        If b_Continue.Visible = False Then
            Run_DeleteEntry_Tutorial
        End If
        Exit Sub
    End If

    If int_GridNum = 2 Then
        str_Response = MyMsgBox("Empty the Garbage Grid?", 1, "YN", "BidBase")
        If str_Response = "yes" Then Exit Sub
        On Local Error GoTo 0 ' EmptyGarbageError
        MousePointer = 11

        Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
        Set rs_DelTable = db_BidBase.OpenRecordset("Garbage")
        rs_DelTable.Index = "PrimaryKey"
        rs_DelTable.Seek ">=", 1
        If rs_DelTable.NoMatch Then GoTo ExitEmptyGarbage
        Do
            rs_DelTable.Delete
            If Not rs_DelTable.EOF Then
                rs_DelTable.MoveNext
            End If
            DoEvents
        Loop While Not rs_DelTable.EOF

        GoTo ExitEmptyGarbage

EmptyGarbageError:
        Resume ExitEmptyGarbage
    End If

    b = Grid1(int_GridNum).Columns(int_Bid_Col).Text
    str_Response = MyMsgBox("Delete ID#" & Grid1(int_GridNum).Columns(int_LastBidCol).Text & _
                            ", Order #" & Grid1(int_GridNum).Columns(int_OrdNum_Col).Text & _
                            ", Bid: " & b$ & "?", 1, "YN", "BidBase")
    If str_Response <> "yes" Then
        Exit Sub
    End If

    ' n is the ID# of the entry to delete: '
    n = Val(Grid1(int_GridNum).Columns(int_LastBidCol))
    If n = 0 Then
        str_Response = MyMsgBox("No entry number to delete.", 1, "ok", "Delete Entry")
        Exit Sub
    End If

    ' Though the Data Control has BidNames Query as its  '
    ' RecordSource for the grids, data lookups should be '
    ' done directly in the Bids table: '
    'Set rs_BidTable = db_BidBase.OpenTable(str_BidsTbl) ' This was done in Sub InitializeDatabase. '
    Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
    Set rs_BidTable = db_BidBase.OpenRecordset(str_BidsTbl)
    rs_BidTable.Index = "PrimaryKey"
    rs_BidTable.Seek "=", n
    If rs_BidTable.NoMatch Then
        str_Response = MyMsgBox("Entry not found.", 1, "ok", "")
        Exit Sub
    End If

    Dim f As String
    Dim p As String
    Dim xzx As Long

    If Trim$(rs_BidTable!Prior_Bids) <> "" Then
        f = Trim$(rs_BidTable!Prior_Bids) & "-" & rs_BidTable!bid
    Else
        f = rs_BidTable!bid
    End If

    p = str_BBpath & "\Notes\"
    f = f & "_ID" & rs_BidTable!id

    If Dir$(p & f & ".txt") = "" Then
        f = f & ".htm"
    Else
        f = f & ".txt"
    End If

    If Dir$(p & f) <> "" Then
        str_Response = MyMsgBox("An information file for this entry exists." & vbCr & _
                                "The file viewer will be loaded so that you" & vbCr & _
                                "can review the file. Use the File menu in" & vbCr & _
                                "the file viewer to delete or rename it.", 4, "ok", "View File")

        str_NoteFile(0) = f
        str_NoteFile(1) = ""
        str_PgmPath(0) = p
        str_PgmPath(1) = ""
        If str_Browser <> "" Then
            xzx = Shell(str_Browser & " " & str_PgmPath(0) & str_NoteFile(0), _
                        vbNormalFocus)
        Else
            xzx = Shell("BidBase File Viewer.exe " & _
                        str_PgmPath(0) & str_NoteFile(0), _
                        vbNormalFocus)
        End If
    End If

    r = Grid1(int_GridNum).Row
    Copy2DeletedTable
    rs_BidTable.Delete
    Data1(int_GridNum).Refresh
    Grid1(int_GridNum).Refresh
    b_ClearInputs_Click

    For i = 0 To 4
        If Grid1(i).Row >= 0 Then
        If Grid1(i).Visible And _
           rs_DelTable!Prior_Bids = Grid1(i).Columns(int_Prior_Bids_Col) _
        Then
            ' I blipped this out before, but it is needed when
            ' there are more entries than fit on the grid.
            RememberGridRows i, "Deleting"
            RestoreGridRows i
        End If
        End If
    Next

ExitEmptyGarbage:
    DoEvents
    Data1(2).Refresh
    MousePointer = 0
    On Local Error Resume Next
    ' This seems to be the easiest way to get '
    ' the data control and grid to update.    '
    If t_Show(0) <> prBid And t_Show(2) <> prBid Then
        t_Show(0) = prBid
        t_Show(1) = prBid
    End If
    Timer3.Enabled = True
    b_Show_Click
End Sub


Private Sub Run_DeleteEntry_Tutorial()
    frm_Tutorial.Caption = "Delete Entry"
    L_Tutorial = "Now on to the Edit menu: " & _
                 vbCr & vbCr & _
                 "To delete an entry from the database, " & _
                 "first display it on the grid, then click on it, " & _
                 "then click on the Delete menu or press Ctrl-Z to " & _
                 "Zap it. The entry will be moved from the database " & _
                 "into the 'Garbage Grid' from where it can be recovered " & _
                 "if necessary." & _
                 vbCr & vbCr & _
                 "If the Garbage Grid is visible and is the last grid " & _
                 "clicked on before selecting Delete, the entire " & _
                 "Garbage Grid will be emptied."
    b_Continue.Visible = True
    L_Tutorial.Height = 2800
    ResizeTutorial
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#E1"
End Sub


Private Sub mEdit_Click(Index As Integer)
    ' This routine takes an entry in a grid  '
    ' which was clicked on its left edge and '
    ' puts it into the Input Boxes. '
    Dim x As Integer

    If frm_Tutorial.Visible And _
       b_Continue.Visible = False _
    Then
        If Index = 0 And _
           str_TutorSubject <> "Edit Entry" _
        Then
            Run_Edit_Tutorial
            Exit Sub

        ElseIf Index = 1 And _
               frm_Tutorial.Caption <> "Edit Next Entry" _
        Then
            frm_Tutorial.Caption = "Edit Next Entry"
            L_Tutorial = "'Edit Next' makes it easy to move one entry after " & _
                         "another from a Grid into the Input Boxes " & _
                         "without having to use any other keys." & _
                         vbCr & vbCr & _
                         "Try it now - pressing Ctrl-N several times in " & _
                         "a row and watching the Input Boxes."
            b_Continue.Visible = True
            b_TutorDetails.Visible = True
            str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#E5"
            Exit Sub

        ElseIf Index = 2 And _
               frm_Tutorial.Caption <> "Use Entry" _
        Then
            Run_UseEntry_Tutorial
            Exit Sub
        End If
    End If

    If int_GridNum = 2 Then
        DisplayGarbage

    Else
        bln_Abort = False

        If mDupesGrid(0).Checked = False Then
            lng_DupeID = 0
            str_DupePrior = ""

        ElseIf int_GridNum = 0 Then
            str_Job = "Duping"
            If mDupesGrid(0).Checked And _
               mCompare.Checked And _
               Grid1(0).Columns(int_Prior_Bids_Col) <> "" _
            Then
                str_DupePrior = Grid1(1).Columns(int_Prior_Bids_Col)
            End If
        End If

        bln_UsingEntry = True

        ' Put grid entry into Input Boxes: '
        DisplayBid2Edit Index

        RememberGridRows int_GridNum, "Editing"

        If bln_UsingEntry Then
            If int_GridNum = 0 Then
                RememberGridRows 1, "Editing"
            Else
                RememberGridRows 0, "Editing"
            End If
        End If

        If 1 = 2 And ck_Saved.Value = 1 Then
            ck_Saved.Value = 0
            str_Response = MyMsgBox("Save this as update to existing entry?" & vbCr & _
                                    "(Otherwise save it as a new entry.)", _
                                    1, "YN", "BidBase")
            If str_Response = "yes" Then
                ck_SaveNew.Value = 0
            End If
        End If
    End If
End Sub


Private Sub Run_Edit_Tutorial()
    frm_Tutorial.Caption = "Edit Entry"
    Grid1(0).Row = 0
    Grid1(0).Col = int_Prior_Bids_Col
    If Grid1(0).Text <> "1H-P" Then
        t_Show(0) = "1H-P"
        t_Show(1) = "1H-P"
        cmb_Show(0).ListIndex = int_Prior_Bids_Col
        b_Show_Click
    End If
    If mDupesGrid(0).Checked Then
        mDupesGrid_Click 0
    End If

    L_Tutorial = "As you can see, the entries for responses to 1H-P " & _
                 "are displayed in the grid. You can scroll the grid " & _
                 "and see most of each entry, but it is easier to see " & _
                 "all of the fields for an entry if you move it into " & _
                 "the Input Boxes at the bottom of the screen." & _
                 vbCr & vbCr & _
                 "Do this by clicking the left edge of a line."

    L_Tutorial.Height = 2500
    frm_Tutorial.Move frame0(6).Top, cmb_ConvName(0).Left
    ResizeTutorial
    bln_Waiting = False
    Wait
    If frm_Tutorial.Visible = False Then
        Exit Sub
    End If

RET_FirstSectionA:
    L_Tutorial = "The Input Box area is divided into 5 sections. " & _
                 "The first section is for describing the bid itself. " & _
                 vbCr & vbCr & _
                 "The Bid/Convention Name field shows the name of the " & _
                 "convention, if any, of which this bid is a part. " & _
                 vbCr & vbCr & _
                 "If the name has a colon followed by another name, " & _
                 "that is the subcategory name." & _
                 vbCr & vbCr & _
                 "Bid/Convention names are chosen from the drop-down " & _
                 "list. You can change a name if you wish, or enter " & _
                 "a new name to add to the list."

    Dim frmC As Long
    Dim backC As Long
    frmC = QBColor(cmb_ForeColors(3).ListIndex)
    backC = QBColor(cmb_BackColors(3).ListIndex)
    GoSub Highlight_BidSect
    frm_Tutorial.Top = f_ShowSort.Top

    L_Tutorial.Height = 3850
    cmb_ConvName(0).SelStart = 0
    cmb_ConvName(0).SelLength = 32768
    b_Go_Back.Visible = True
    b_TutorDetails.Visible = True
    str_TutorDetails = "CardShark_BidBase_Editor_Input_Boxes.htm#B"
    ResizeTutorial
    Wait
    If bln_GoBack Then
        Run_Edit_Tutorial
        Exit Sub
    End If
    cmb_ConvName(0).SelLength = 0
    If frm_Tutorial.Visible = False Then
        GoSub Unhighlight_BidSect
        Exit Sub
    End If


RET_FirstSectionB:
    L_Tutorial = "If there is a button with a red 'Y' over the convention " & _
                 "name, that indicates " & _
                 "that there is a file which discusses the convention. " & _
                 "Click the button to view the file." & _
                 vbCr & vbCr & _
                 "If the button has an 'N', you can click it to create " & _
                 "a note file, but we don't want to do that now." & _
                 vbCr & vbCr & _
                 "The 'G' button brings up Google's home page on the Web " & _
                 "and uses it to search for information about the " & _
                 "Convention Name (up to the colon, if there is one). " & _
                 "You can try this if you are connected to the Web."

    L_Tutorial.Height = 3000
    ResizeTutorial
    b_Go_Back.Visible = True
    str_TutorDetails = "CardShark_BidBase_Editor_Input_Boxes.htm#L1"
    Wait
    If bln_GoBack Then GoTo RET_FirstSectionA
    cmb_ConvName(0).SelLength = 0
    If frm_Tutorial.Visible = False Then
        GoSub Unhighlight_BidSect
        Exit Sub
    End If


RET_FirstSectionC:
    L_Tutorial = "The Cnv field (far left) shows if the convention of " & _
                 "which this " & _
                 "entry is a part has been selected for use. This field " & _
                 "applies to the bid/convention shown in the Bid/Conv. " & _
                 "Name field. " & _
                 vbCr & vbCr & _
                 "If the bid/convention name has a subcategory (a name " & _
                 "after a colon after the bid name), the Sub field shows " & _
                 "if that subconvention has been selected for use." & _
                 vbCr & vbCr & _
                 "In either of these fields, a blank indicates that " & _
                 "the convention or subcategory is being used by both " & _
                 "sides. W=We, only.  T=They, only.  0=neither side." & _
                 vbCr & vbCr & _
                 "A change in either of these fields affects all entries " & _
                 "using the current Bid/Convention name, not just the " & _
                 "current entry. To de/select just the current entry, " & _
                 "use the Pct field."


    L_Tutorial.Height = 4150
    ResizeTutorial
    b_Go_Back.Visible = True
    str_TutorDetails = "CardShark_BidBase_Editor_Input_Boxes.htm#B"
    Wait
    If bln_GoBack Then GoTo RET_FirstSectionB
    If frm_Tutorial.Visible = False Then
        GoSub Unhighlight_BidSect
        Exit Sub
    End If

RET_FirstSectionD:
    L_Tutorial = "Next is Pct. The Cnv field indicates that the current " & _
                 "entry is part of a convention which is being used. " & _
                 "Pct lets you select alternative bids within that " & _
                 "convention. " & _
                 vbCr & vbCr & _
                 "For example, playing 'Strong 2C', " & _
                 "if you want a 2H response to be natural, you would " & _
                 "leave its Pct field blank. If you play 2H as negative, " & _
                 "you would leave that blank and put '0' in Pct for the " & _
                 "entry for a natural 2H." & _
                 vbCr & vbCr & _
                 "You can also enter alternative bids for the same " & _
                 "hand specifications when one bid is not clearly " & _
                 "superior, then enter numbers in each entry's Pct " & _
                 "field adding up to 100%, such as 60-40 or 50-50."

    L_Tutorial.Height = 3200
    ResizeTutorial
    b_Go_Back.Visible = True
    str_TutorDetails = "CardShark_BidBase_Editor_Input_Boxes.htm#B"
    Wait
    If bln_GoBack Then GoTo RET_FirstSectionC
    If frm_Tutorial.Visible = False Then
        GoSub Unhighlight_BidSect
        Exit Sub
    End If

RET_FirstSect2:
    L_Tutorial = "The next field is the Round (of bidding - starting with " & _
                 "the first non-passing hand), then the Player number, again " & _
                 "starting with the first non-passing hand. " & _
                 "Example: After P-P-P-1S, the next bid will be the second " & _
                 "bidder in the first round of bidding because the first " & _
                 "three passes are ignored. This saves space in the database " & _
                 "and if being a passed hand affects what bid is made, that " & _
                 "can be specified in the Position field." & _
                 vbCr & vbCr & _
                 "Next are the Prior_Bids, the Bid Order Number, and the Bid." & _
                 vbCr & vbCr & _
                 "This is the bid which will be made if the hand " & _
                 "under consideration meets all " & _
                 "the specs of this entry."

    L_Tutorial.Height = 3500
    ResizeTutorial
    b_Go_Back.Visible = True
    t_Round.SelStart = 0
    t_Round.SelLength = 1
    str_TutorDetails = "CardShark_BidBase_Editor_Input_Boxes.htm#D"
    Wait
    t_Round.SelLength = 0
    If bln_GoBack Then GoTo RET_FirstSectionB
    If frm_Tutorial.Visible = False Then
        GoSub Unhighlight_BidSect
        Exit Sub
    End If

RET_HandStr:
    L_Tutorial = "'Hand Strength' is the second section. " & _
                 "Click on the input box for Standard HCPs " & _
                 "and look at the long white box at the bottom of the screen. " & _
                 "It shows you the different formats which can be used to " & _
                 "specify High Card Points (HCP)." & _
                 vbCr & vbCr & _
                 "The other fields in this section are discussed in the " & _
                 "documentation. Click on the Details button to view it."

    L_Tutorial.Height = 3000
    ResizeTutorial
    GoSub Unhighlight_BidSect
    GoSub Highlight_HandStr
    str_TutorDetails = "CardShark_BidBase_Editor_Input_Boxes.htm#N"
    Wait
    GoSub Unhighlight_HandStr
    If bln_GoBack Then
        GoSub Highlight_BidSect
        GoTo RET_FirstSect2
    End If
    If frm_Tutorial.Visible = False Then Exit Sub

RET_SuitSpecs:
    L_Tutorial = "Suit specifications are entered in the next section. " & _
                 "General hand shape can be entered using the " & _
                 "drop-down list, but is not needed " & _
                 "if you define the hand shape using the individual suit " & _
                 "'Num.' (quantity) and 'Points' fields." & _
                 vbCr & vbCr & _
                 "The 'X' suit field lets you combine several entries into " & _
                 "one by putting a suit spec into the X field and putting " & _
                 "'x' in the individual suit fields. For example, to specify " & _
                 "any 6+ card suit, put 6+ in X's Num box, then put x in " & _
                 "the Num boxes of any suits you want considered." & _
                 vbCr & vbCr & _
                 "The 'Comb' box lets you enter combined suit specs. " & _
                 "See the help file for details. The Comb - Points box extends into the next section for slam bidding specs."

    GoSub Highlight_HandStr
    L_Tutorial.Height = 3800
    ResizeTutorial
    GoSub Unhighlight_HandStr
    GoSub Highlight_Suits
    str_TutorDetails = "CardShark_BidBase_Editor_Input_Boxes.htm#W"
    Wait
    GoSub Unhighlight_Suits
    If bln_GoBack Then
        GoSub Highlight_HandStr
        GoTo RET_HandStr
    End If
    If frm_Tutorial.Visible = False Then Exit Sub

RET_Other:
    L_Tutorial = "In the 'Other' section, you can specify Vulnerability " & _
                 "and position using the drop-down lists. If you wish, " & _
                 "click on the those lists to see the alternatives." & _
                 vbCr & vbCr & _
                 "'Best Contract' is a special field which you can ignore." & _
                 vbCr & vbCr & _
                 "The 'Artificial' check box indicates if a bid is artificial. " & _
                 "A bid is artificial if it is not intended to indicate an " & _
                 "offer to play in the contract bid. Examples: any cue bid, " & _
                 "Stayman 2C, Blackwood 4N, a take-out double. (The 'natural' " & _
                 "meaning of a double is for penalty.)"

    L_Tutorial.Height = 3000
    ResizeTutorial
    GoSub Highlight_Other
    str_TutorDetails = "CardShark_BidBase_Editor_Input_Boxes.htm#Y"
    Wait
    GoSub Unhighlight_Other
    If bln_GoBack Then
        GoSub Highlight_Suits
        GoTo RET_SuitSpecs
    End If
    If frm_Tutorial.Visible = False Then Exit Sub

RET_Discl:
    L_Tutorial = "In the 'Disclosure' section are fields mostly for " & _
                 "providing information about the bids to opponents." & _
                 "They are ignored when choosing a bid for the current hand." & _
                 vbCr & vbCr & _
                 "Click on the 'Asks partner' field and look down in " & _
                 "the prompt box to see what kinds of information " & _
                 "can be entered in this field." & _
                 vbCr & vbCr & _
                 "The 'Strength' and 'Distribution' fields must use standard " & _
                 "formats because their data are what is given to the " & _
                 "bridge-playing software so that it knows what the " & _
                 "bid shows. These fields are also shown to human players."

    L_Tutorial.Height = 3000
    ResizeTutorial
    GoSub Highlight_Discl
    str_TutorDetails = "CardShark_BidBase_Editor_Input_Boxes.htm#AA"
    Wait
    If bln_GoBack Then
        GoSub Unhighlight_Discl
        GoSub Highlight_Other
        GoTo RET_Other
    End If
    If frm_Tutorial.Visible = False Then
        GoSub Unhighlight_Discl
        Exit Sub
    End If

RET_Discl2:
    L_Tutorial = "The long disclosure field is used for " & _
                 "more descriptive information for humans. It is also " & _
                 "used for notes about the entry which are NOT shown " & _
                 "to the player. These notes are for assistance when you are " & _
                 "working on the database. Such notes are identified " & _
                 "by starting them with a square bracket." & _
                 vbCr & vbCr & _
                 "If you want to enter more notes about a bid than " & _
                 "will fit in the disclosure field, click the button " & _
                 "to the right of the field to bring up a note file." & _
                 "If the button has a red 'Y', there is already a " & _
                 "note field for this entry. You can also refer to " & _
                 "other entries' note files or convention files by " & _
                 "putting the file name in the disclosure field."

    L_Tutorial.Height = 3200
    ResizeTutorial
    t_Disclosure.SetFocus
    t_Disclosure.SelStart = 0
    t_Disclosure.SelLength = 32768
    str_TutorDetails = "CardShark_BidBase_Editor_Input_Boxes.htm#AH"
    Wait
    If bln_GoBack Then GoTo RET_Discl
    If frm_Tutorial.Visible = False Then
        GoSub Unhighlight_Discl
        Exit Sub
    End If

RET_Discl3:
    L_Tutorial = "The 'Test Hand' field is used by the software to " & _
                 "verify that the entry's specs are for the type of " & _
                 "hand you intended, and that the entry's specs do " & _
                 "not conflict with the specs of other entries for " & _
                 "the same Prior Bids. The Test Hand normally meets " & _
                 "the very minimum specs of the entry." & _
                 vbCr & vbCr & _
                 "The Auto-Fill check box tells the " & _
                 "program to automatically create or modify the Test " & _
                 "hand whenever you change the suit specs. Change " & _
                 "a 'Num.' field now and press Tab." & _
                 vbCr & vbCr & _
                 "The 'Reference' field tells the user where to look " & _
                 "for more information about the entry's specs and bid."

    ck_AutoFill.Value = 1
    t_HCP(4).SetFocus
    L_Tutorial.Height = 3500
    ResizeTutorial
    t_TestHand.SetFocus
    t_TestHand.SelStart = 0
    t_TestHand.SelLength = 32768
    str_TutorDetails = "CardShark_BidBase_Editor_Input_Boxes.htm#AI"
    Wait
    If bln_GoBack Then GoTo RET_Discl2
    GoSub Unhighlight_Discl
    If frm_Tutorial.Visible = False Then
        Exit Sub
    End If

RET_Reload:
    L_Tutorial = "If you changed the suit specs as suggested in the last " & _
                 "window, which in turn changed the Test Hand, the changes " & _
                 "are not saved to the database until you save the entry. " & _
                 vbCr & vbCr & _
                 "You can reload the original entry simply by pressing " & _
                 "Ctrl-E or double-clicking on the same line in the " & _
                 "grid again. " & _
                 vbCr & vbCr & _
                 "If you save a changed entry without really wanting to, " & _
                 "you can recover the original entry by clicking on " & _
                 "Edit - Restore on the menu bar." & _
                 vbCr & vbCr & _
                 "You can run a tutorial on that menu command later."

    ck_AutoFill.Value = 0
    L_Tutorial.Height = 3500
    ResizeTutorial
    GoSub Unhighlight_Discl
    Wait
    If bln_GoBack Then
        GoSub Highlight_Discl
        GoTo RET_Discl3
    End If
    If frm_Tutorial.Visible = False Then Exit Sub


    L_Tutorial = "When an " & _
                 "entry is saved, the program automatically tests it " & _
                 "using the Test Hand. If it fails the test, you can " & _
                 "will be shown why by the bidding module." & _
                 "Then you can make corrections and try again."

    b_Continue.Visible = True
    L_Tutorial.Height = 3000
    ResizeTutorial
    b_Go_Back.Top = b_Continue.Top
    b_QuitTutorial.Top = b_Continue.Top
    frm_Tutorial.Height = b_Continue.Top + 500
    Wait
    If bln_Abort Then GoTo RET_Reload
    If frm_Tutorial.Visible = False Then Exit Sub

    L_Tutorial = "Before leaving the Input Boxes, click on the different " & _
                 "fields, looking at the prompts in red at the bottom of " & _
                 "screen to see what kinds of data they take." & _
                 vbCr & vbCr & _
                 "If you are still not clear on what a particular field " & _
                 "is for, press F1 to bring up the Help file and read " & _
                 "more about the field in the 'Input Box' help file." & _
                 vbCr & vbCr & _
                 "When you have completed an entry, you can click the " & _
                 "File - Save menu to save it, or just press Ctrl-S. " & _
                 "If you have not changed anything in this entry, you " & _
                 "can press Ctrl-S now to save it again without hurting anything."

    b_Go_Back.Visible = False
    L_Tutorial.Height = 3000
    ResizeTutorial
Exit Sub


Highlight_BidSect:
'''''''''''''''''
    Dim i As Integer

    frame0(0).ForeColor = QBColor(15)
    frame0(0).BackColor = QBColor(4)
    For i = 0 To 9
        L_Bid(i).ForeColor = QBColor(15)
    Next
    ck_SaveNew.ForeColor = QBColor(15)
    ck_SaveNew.BackColor = QBColor(4)
    ck_Saved.ForeColor = QBColor(15)
    ck_Saved.BackColor = QBColor(4)
    ck_Continuation.BackColor = QBColor(4)
    ck_Continuation.ForeColor = QBColor(15)
    f_OtherPlayers.BackColor = QBColor(4)
    f_OtherPlayers.ForeColor = QBColor(15)

    L_Date_Saved.ForeColor = QBColor(15)
    L_IDnum.ForeColor = QBColor(15)
Return

Unhighlight_BidSect:
'''''''''''''''''''
    frame0(0).ForeColor = frmC
    frame0(0).BackColor = backC
    For i = 0 To 9
        L_Bid(i).ForeColor = frmC
    Next
    ck_SaveNew.ForeColor = frmC
    ck_SaveNew.BackColor = backC
    ck_Continuation.BackColor = backC
    ck_Continuation.ForeColor = frmC
    f_OtherPlayers.BackColor = backC
    f_OtherPlayers.ForeColor = frmC
    ck_Saved.ForeColor = frmC
    ck_Saved.BackColor = backC
    L_Date_Saved.ForeColor = frmC
    L_IDnum.ForeColor = frmC
    'L_Date_Saved.BackColor = backC
Return

Highlight_HandStr:
'''''''''''''''''
    For i = 1 To 2
        frame0(i).ForeColor = QBColor(15)
        frame0(i).BackColor = QBColor(4)
    Next
    For i = 0 To 15
        If i = 2 Then i = 3
        If i = 14 Then i = 15
        L_HandStr(i).ForeColor = QBColor(15)
    Next


Return

Unhighlight_HandStr:
'''''''''''''''''''
    For i = 1 To 2
        frame0(i).ForeColor = frmC
        frame0(i).BackColor = backC
    Next
    For i = 0 To 15
        L_HandStr(i).ForeColor = frmC
    Next
Return

Highlight_Suits:
'''''''''''''''
    frame0(3).ForeColor = QBColor(15)
    frame0(3).BackColor = QBColor(4)
    For i = 0 To 8
        L_Suits(i).ForeColor = QBColor(15)
    Next
Return

Unhighlight_Suits:
'''''''''''''''''
    frame0(3).ForeColor = frmC
    frame0(3).BackColor = backC
    For i = 0 To 8
        L_Suits(i).ForeColor = frmC
    Next
Return

Highlight_Other:
'''''''''''''''
    frame0(4).ForeColor = QBColor(15)
    frame0(4).BackColor = QBColor(4)
    ck_Artificial.ForeColor = QBColor(15)
    ck_Artificial.BackColor = QBColor(4)
    For i = 0 To 2
        L_Other(i).ForeColor = QBColor(15)
    Next
Return

Unhighlight_Other:
'''''''''''''''''
    frame0(4).ForeColor = frmC
    frame0(4).BackColor = backC
    ck_Artificial.ForeColor = frmC
    ck_Artificial.BackColor = backC
    For i = 0 To 2
        L_Other(i).ForeColor = frmC
    Next
Return

Highlight_Discl:
'''''''''''''''
    frame0(5).ForeColor = QBColor(15)
    frame0(5).BackColor = QBColor(4)
    ck_AutoFill.ForeColor = QBColor(15)
    ck_AutoFill.BackColor = QBColor(4)
    For i = 0 To 9
        L_Discl(i).ForeColor = QBColor(15)
    Next
Return

Unhighlight_Discl:
'''''''''''''''''
    frame0(5).ForeColor = frmC
    frame0(5).BackColor = backC
    ck_AutoFill.ForeColor = frmC
    ck_AutoFill.BackColor = backC
    For i = 0 To 9
        L_Discl(i).ForeColor = frmC
    Next
Return

End Sub


Private Sub Run_UseEntry_Tutorial()
    frm_Tutorial.Caption = "Use Entry"
    If mTwoGrids.Checked = False Then
        frm_Tutorial.Visible = False
        mTwoGrids_Click
        frm_Tutorial.Visible = True
    End If

    If mInputBoxes.Checked = False Then
        frm_Tutorial.Visible = False
        mInputBoxes_Click
        frm_Tutorial.Visible = True
    End If

    Grid1(0).Row = 0
    Grid1(0).Col = int_Prior_Bids_Col
    Grid1(1).Row = 0
    Grid1(1).Col = int_Prior_Bids_Col
    If Grid1(0) <> "1H-P" Or Grid1(1) <> "1S-P" Then
        t_Show(0) = "1H-P"
        t_Show(1) = "1H-P"
        t_Show(2) = "1S-P"
        t_Show(3) = "1S-P"
        cmb_Show(0).ListIndex = int_Prior_Bids_Col
        cmb_Show(1).ListIndex = int_Prior_Bids_Col
        b_Show_Click
    End If

    L_Tutorial = "Click on the line in the top grid for the Bid 2H. " & _
                 "(This is entry number 241000.)" & _
                 vbCr & vbCr & _
                 "Click on 'Edit - Use Entry' on the menu bar " & _
                 "(or press Ctrl-U) and you will notice that the entry " & _
                 "has been copied into the Input Boxes, but that some " & _
                 "changes have been made to it." & _
                 vbCr & vbCr & _
                 "The Prior_Bids have been changed to the lower grid's " & _
                 "Prior_Bids, the Bid has been changed from 2H to " & _
                 "2S, the suit specs have changed so that Spades now " & _
                 "shows 3+ instead of Hearts, and the Distribution and " & _
                 "Test Hand entries have also changed."

    L_Tutorial.Height = 3100
    ResizeTutorial
    Wait
    If bln_Abort Then Exit Sub

    L_Tutorial = "Do not save this entry because we already have an " & _
                 "entry like this in the database, but if this were a " & _
                 "new entry which you were adding, all you would have " & _
                 "to do at this point is press Ctrl-S to save it without " & _
                 "having to make a single change to the entry." & _
                 vbCr & vbCr & _
                 "Sometimes you may have to modify the entries a little " & _
                 "when using an entry from one section to create a new " & _
                 "entry in another section, but it's always a lot easier " & _
                 "than starting from scratch."
    L_Tutorial.Height = 2500
    ResizeTutorial
End Sub


Private Sub DisplayBid2Edit(Index)
    ' This is the 3rd sub in the series which takes an entry '
    ' which was double-clicked in a grid and displays it in  '
    ' the Input Boxes. '
    ' This is also called by Sub mFindEntryID to find lng_ID. '
    Dim gn As Integer
    Dim pb1 As String
    Dim pb2 As String
    Dim c As Integer ' for comma place in t_HCP '
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim n As Integer
    Dim x As Integer
    Dim z As Integer
    Dim a As String
    Dim b As String
    Dim p As String
    Dim s As String
    Dim qty(4) As String
    Dim hcp(4) As String
    Dim x2nonX As Boolean
    Dim cmQ As String
    Dim cmH As String
    Dim num As Long
    Dim nam As String
    Dim sel As String
    Dim sel2 As String
    Dim dash As Integer
    Dim rhoSuit As Integer
    Dim pardSuit As Integer
    Dim lhoSuit As Integer
    Dim xLvlDiff As Integer
    Dim xSuitHigher As Boolean
    Dim lastSuitBid As String
    Dim lastPriorSuitBid As String
    Dim lastLevelBid As Integer
    Dim lastPriorLevelBid As Integer
    Dim th As String
    Dim tmp As String
    Dim s1 As Integer
    Dim s2 As Integer
    Dim unswapped(2) As String
    Dim swapDistrib As Boolean
    Dim origDistrib As String
    Dim newDistrib As String
    Dim origTestSuits(4) As String
    Dim newTestSuits(4) As String
    Dim bidConverted As Boolean

    If Index = 99 Then
        ' Coming from Sub mFindBid. '
        GoTo GetEntry
    End If

    If Index < 4 Then
        If Grid1(Index).Row = -1 Then
            MsgBox "Grid #" & Index & " is empty."
            Exit Sub
        End If
    Else
        MsgBox "Grid number is invalid: " & Index
        Exit Sub
    End If

    ' Working variables for "Index": '
    Dim editNext  As Boolean
    Dim useEntry  As Boolean

    'editNext = (Index = 1)
    useEntry = True ' (Index = 2)

    If frm_Tutorial.Visible Then
        frm_Tutorial.Visible = False
        mClear_Click
        frm_Tutorial.Visible = True

    ElseIf str_Job <> "Duping" Or mCompare.Checked Then
        mClear_Click
    End If

    DoEvents

    If str_Job = "Duping" Then
        int_GridNum = 0
        ColorGrids 0
    End If

    gn = int_GridNum

    'If editNext Then
    '    Grid1(gn).Row = Grid1(gn).Row + 1
    'End If
    On Local Error GoTo 0

    If lng_ID > 0 Then
        str_Response = MyMsgBox("Entry in the Input Boxes" & vbCr & _
                                "is not in the top grid.", 2, "ok", "")
        Exit Sub
    Else
        lng_ID = Val(Grid1(gn).Columns(int_LastBidCol).Text)
    End If

    If lng_ID = 0 Then
        If mTestHands.Checked Then
            mTestHands.Checked = False
            Exit Sub

        ElseIf int_GridNum = 0 Then
            str_Response = MyMsgBox("No record selected.", 1, "ok", "")
            Exit Sub

        Else
            int_GridNum = 0
            ColorGrids 0
            DisplayBid2Edit Index
            Exit Sub
        End If
    End If

GetEntry:
    ' Get the entry from the database:                            '
    ' This should be the same as the grid, but it is safer to use '
    ' the database entry because, for example, "rs_BidTable!Bid"  '
    ' is always the saved Bid value, but the user may change the  '
    ' Bid value in the grid, on purpose or not. Using the database'
    ' is easier because to reference the grid, you have to use -- '
    ' Grid1(int_GridNum).Columns(int_Bid_Col).Text instead of just'
    ' rs_BidTable!Bid, plus you have to define a variable for     '
    ' each column in the grid, such as "int_Bid_Col", since the   '
    ' grid contents may change (so you can't just use column #'s).'
    ' Set rs_BidTable = db_BidBase.OpenRecordset("Bids")
    Set ws_BidBase = DBEngine.Workspaces(0)
    Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
    Set rs_BidTable = db_BidBase.OpenRecordset(str_BidsTbl)
''''''''''''''''
With rs_BidTable
''''''''''''''''
    .Index = "PrimaryKey"
    .Seek "=", lng_ID

    If .NoMatch Then
        str_Response = MyMsgBox("BidTable entry ID#" & lng_ID & vbCr & _
                                "is not in BidBase.", 2, "ok", "")
        Exit Sub
    End If

    ' Fill Input Boxes from database record: '
    t_Pct = !Pct_Used
    t_Round = !RoundOfBidding
    t_Player = !player
    t_Prior = Trim$(!Prior_Bids)
    t_Order = !Bid_Order_Number


    ' Bid field:                                                        '

    cmb_Bid.Text = Trim$(!bid)
    If InStr(cmb_Bid.Text, "@") > 0 Then
        ' Slam exploring bid made.  '
        ' Increase cmb_Bid.Size and '
        ' location to make room.    '
        If cmb_Bid.Width <> 1015 Then
            t_Prior.Width = 3380
            t_Order.Left = 4990
            cmb_Bid.Left = 5835
            cmb_Bid.Width = 1050
        End If
    ElseIf cmb_Bid.Width = 1050 Then
        t_Prior.Width = 3765
        t_Order.Left = 5375
        cmb_Bid.Left = 6240
        cmb_Bid.Width = 650
    End If

    ' Convention Name:                                                  '

    If !ConvID = int_DividerLineID Then
        ' This is a divider line, not an actual bid entry, '
        ' so skip the rest of this routine.                '
        Exit Sub
    End If

    num = Val("" & !ConvID)
    If num <> 0 Then
        LookUpBidName num, sel, sel2, nam
    End If
    t_Sel = "" & sel
    t_Sub = "" & sel2
    If nam <> "" Then
        ' Look up name in cmb_ConvName list: '
        i = FindMatch(cmb_ConvName(0), nam, True) ' True = exact match
        If i < 0 Then i = 0
    End If
    cmb_ConvName(0).ListIndex = i

    ' Look for file for convention name: '
    If mCompare.Checked = False Then
        FindConvenNameFile nam, True, False, False
    End If


    ' Vulnerability:                                                     '

    ' Find Vul in drop-down list using ListIndex: '
    ' (We don't save the ListIndex because:       '
    '    [1] The list could change and            '
    '    [2] An Index # is not decipherable when  '
    '         just viewing data in the grid.)     '
    a = UCase$(Trim$(!Vul))
        If a = "U" Then ' Unfav. '
            i = 1
        ElseIf a = "F" Then ' Favorable '
            i = 2
        ElseIf a = "B" Then ' Both vul. '
            i = 3
        ElseIf a = "N" Then ' Neither vul. '
            i = 4
        ElseIf a = "V" Then ' Bidder vul. '
            i = 5
        ElseIf a = "X" Then ' Bidder not vul. '
            i = 6
        ElseIf a = "" Then ' Any vul is okay. '
            i = 0
        Else
            str_Response = MyMsgBox(a & " is an invalid Position code.", 1, "ok", "Sub DisplayBid2Edit")
        End If
    cmb_Vul.ListIndex = i


    ' Position:                                                          '

    ' Find Pos in drop-down list using ListIndex: '
    a = LCase$(Trim$(!pos))
    i = InStr("ft1234ynpu", a)
    If a = "" Then i = 0
    cmb_Position.ListIndex = i

    ' Distribution Points:                                               '
    t_Distrib_Pts = !Dist_Pts

    ' Custom Points:                                                     '
    t_CustPoints = "" & !Cust_Pts
    t_CustPts(0) = "" & !A_Cust
    t_CustPts(1) = "" & !K_Cust
    t_CustPts(2) = "" & !Q_Cust
    t_CustPts(3) = "" & !J_Cust
    t_CustPts(4) = "" & !T_Cust

    ' HCP:                                                               '
    t_Hand_HCPs = !hcp
    c = InStr(t_Hand_HCPs, ",")
    If c > 0 Then
        t_BBPoints = Mid$(t_Hand_HCPs, c + 1)
        t_Hand_HCPs = Left$(t_Hand_HCPs, c - 1)
    End If

    ' Winners:                                                           '
    t_Winners = !Winrs

    ' Losers:                                                            '
    t_Losers = !losers

    ' QTs:                                                               '
    t_QTs = !QTs

    ' Stopped_Suits:                                                     '
    t_Stoppers = !Stopped_Suits

    ' Shape:                                                             '

    Select Case UCase$(Trim$(!Shape))
        Case "1"
            cmb_Shape.ListIndex = 1
        Case "2"
            cmb_Shape.ListIndex = 2
        Case "3"
            cmb_Shape.ListIndex = 3
        Case "B"
            cmb_Shape.ListIndex = 4
        Case "U"
            cmb_Shape.ListIndex = 5
        Case "S"
            cmb_Shape.ListIndex = 6
        Case "V"
            cmb_Shape.ListIndex = 7
        Case "Y", "E"
            cmb_Shape.ListIndex = 8
        Case Else
            If cmb_Shape.ListIndex > 0 Then
                cmb_Shape.ListIndex = 0
            End If
    End Select

    ' Intermediates:                                                     '
    t_Intermediates = "" & !intermed

    ' Suit length & HCPs:                                                '
    qty(0) = !xNum
    hcp(0) = !xHCP
    qty(1) = !sNum
    hcp(1) = !sHCP
    qty(2) = !hNum
    hcp(2) = !HHCP
    qty(3) = !dNum
    hcp(3) = !dHCP
    qty(4) = !cNum
    hcp(4) = !cHCP

    ' Suit Combo's, total number of cards:                               '
    cmQ = "" & !cmNum
    cmH = "" & !cmHCP
    'Note:"" & ... (above) avoids Null error.
    t_Qty(5) = cmQ
    t_HCP(5) = cmH

    ' Alert:                                                             '
    i = InStr(" ASDC", UCase$(Trim$(!Alert)))
    If i > 0 Then i = i - 1
    cmb_Alert.ListIndex = i

    ' Artificial:                                                        '
    If !Artificial = "A" Then
        ck_Artificial.Value = 1
    Else
        ck_Artificial.Value = 0
    End If

    ' Explanation:                                                       '
    t_Disclosure = !Explanation

    ' Need to call LookForNoteFile even when '
    ' comparing entries so that if the Discl.'
    ' field says "[See notes -->" it will be '
    ' converted to the actual file name for  '
    ' the new entry.                         '
    LookForNoteFile t_Disclosure, _
                    t_Prior, _
                    cmb_Bid, _
                    "" & lng_ID, _
                    useEntry, _
                    True, _
                    False
    ' Above:                                      '
    '    "useEntry" was set at the start of this  '
    '        routine and is either True or False. '
    '    "yes"  -- set the "View File" button.    '
    '    "False" -- don't show note file, if any. '


    ' Last Fields:                                                      '
    t_Ref = "" & !source
    t_AsksFor = "" & !Asks
    L_Date_Saved = "" & !DateOfEntry
    t_BestContract = "" & !Best_Contract
    
    If !MySpecs > 0 Then
        g_OtherHands.TextMatrix(1, 1) = "" & !LHO
        DisplayOtherSpecs 1, !MySpecs
    End If
    If !LHO > 0 Then
        g_OtherHands.TextMatrix(1, 2) = "" & !LHO
        DisplayOtherSpecs 2, !LHO
    End If
    If !Pard > 0 Then
        g_OtherHands.TextMatrix(1, 3) = "" & !Pard
        DisplayOtherSpecs 3, !Pard
    End If
    If !RHO > 0 Then
        g_OtherHands.TextMatrix(1, 4) = "" & !RHO
        DisplayOtherSpecs 4, !RHO
    End If
    
    L_IDnum = "ID#: " & !id
    lng_EntryID = !id


    ' "Save New" Check Box:                                             '
    str_Response = ""
    If lng_EntryID > 0 And _
       ck_SaveNew.Value = 0 And _
       f_CompareGrid.Visible = False And _
       str_Job <> "Testing" _
    Then
        str_Response = MyMsgBox("Mark this as a New Entry?" & vbCrLf & _
                                "(Otherwise the old entry will" & vbCrLf & _
                                  "be overwritten when saved.)", _
                                3, "YNC", "BidBase")
    End If

    If str_Response = "yes" Or useEntry = False Or _
       !ConvID = int_DividerLineID _
    Then
        ck_SaveNew.Value = 1
    Else
        ck_SaveNew.Value = 0
    End If
    
useEntry = False

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Adjustments For "Using Entry":                                     '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Break original Test Hand into suits: '
    th = !Test_Hand
    If th <> "" Then
    For i = 1 To 3
        j = InStr(th, "-")
        origTestSuits(i) = Left$(th, j - 1)
        th = Mid$(th, j + 1)
    Next
    origTestSuits(4) = th
    End If

    If useEntry And _
       (t_Show(2) <> "" Or _
        mDupesGrid(0).Checked) And _
       bln_TwoGrids _
    Then
        'The above line is because it is possible to have specified '
        'showing one grid, blanked out the 2nd grid specs, then     '
        'later done a Ctrl-T to open the 2nd grid without having    '
        'specified the contents of it.                              '
      
        If str_Job <> "Duping" Then
            Select Case int_GridNum
            Case Is = 0
                ' Using top grid entry in bottom grid, but get the '
                ' Prior Bid from the top row of the bottom grid:   '
                i = Grid1(1).Row
                If i = -1 Then i = 0
                Grid1(1).Row = 0
                t_Prior = Grid1(1).Columns(int_Prior_Bids_Col).Text
                If t_Prior = "" And Grid1(1).VisibleRows = 1 Then
                    ' If the user has not entered Prior Bids in  '
                    ' the Input Box and the bottom grid is empty,'
                    ' then get the Prior Bids specified in the   '
                    ' Show/Sort window. If this isn't right      '
                    ' either, the user will have to correct it.  '
                    t_Prior = t_Show(2)
                End If
                Grid1(1).Row = i
                int_GridNum = 1 ' The grid being added to.'

            Case Is = 1
                ' Using bottom grid entry in top grid. '
                i = Grid1(0).Row
                If i = -1 Then i = 0
                Grid1(0).Row = 0
                t_Prior = Grid1(0).Columns(int_Prior_Bids_Col).Text
                Grid1(0).Row = i
                int_GridNum = 0 ' The grid being added to.'
            End Select
        End If
        If mDupesGrid(0).Checked Then
            t_Prior = str_DupePrior
        End If
        ' Variables for later use: '
        pb1 = Trim$(!Prior_Bids) ' = the existing entry. '
        pb2 = t_Prior ' = the new entry. '

     '  Convert "x" specs:                                            '

        If Right$(cmb_Bid, 1) = "x" Then
            ' Make sure that each previous x suit will result '
            ' in a sufficient bid in the new entry; if not,   '
            ' then do not put an x for that suit in the new   '
            ' entry. If only one suit can make a sufficient   '
            ' bid, change the bid from "x" to that suit.      '
            i = Len(pb2)
            k = 0
            n = 0
            dash = 0
            ' Find the suit & bid level the others must beat: '
            Do While i > 0
                If n = 0 Then
                    ' This finds the last suit bid in '
                    ' Prior Bids.                     '
                    n = Val(Mid$(pb2, i - 1, 1))
                    s = Mid$(pb2, i, 1)
                    b = Mid$(pb1, i, 1)
                End If
                i = i - 1
                If Mid$(" " & pb1, i + 1, 1) = "-" Or i = 0 Then
                    dash = dash + 1
                    If Val(Mid$(pb2, i + 1)) > 0 Then
                        Select Case dash
                        Case Is = 1
                            rhoSuit = InStr("SHDC", Mid$(pb2, i + 2, 1))
                        Case Is = 2
                            pardSuit = InStr("SHDC", Mid$(pb2, i + 2, 1))
                        Case Is = 3
                            lhoSuit = InStr("SHDC", Mid$(pb2, i + 2, 1))
                            Exit Do
                        End Select
                    End If
                End If
                DoEvents
            Loop

            k = InStr("CDHS", s)

            xLvlDiff = Val(cmb_Bid) - n
            ' See if the suits with "x" are higher or '
            ' lower than the last Prior Suit bid:     '
            For i = 1 To 4
                If InStr(qty(i), "x") > 0 Or InStr(hcp(i), "x") > 0 Then
                    If Mid$("SHDC", i, 1) >= b Then
                        xSuitHigher = True
                        Exit For
                    End If
                End If
            Next

            j = 0
            If xSuitHigher Then
                ' The "x" suits must be higher  '
                ' that the last suit bid, which '
                ' is k = 1-to-4 = C-to-S.       '
                For i = 4 To 1 Step -1 ' Spades to Clubs
                    If InStr(qty(5 - i), "x") > 0 Or _
                       InStr(hcp(5 - i), "x") > 0 _
                    Then ' Here, S=4, but for qty(), S=1, thus the "5-i". '
                        If i <= k Then
                            ' Not higher, so remove "x": '
                            RemoveX qty(5 - i)
                            RemoveX hcp(5 - i)
                            x2nonX = True
                        Else
                            j = j + 1 ' Number of suits where x is okay. '
                            x = 5 - i ' In case only 1 "x" suit is valid.'
                        End If
                    End If
                Next

            Else
                For i = 1 To 4 ' Clubs to Spades '
                    If InStr(qty(5 - i), "x") > 0 Or _
                       InStr(hcp(5 - i), "x") > 0 _
                    Then
                        If i >= k And i > 1 Then
                            ' Not lower, so remove "x": '
                            s = origTestSuits(5 - i + 1)
                            origTestSuits(5 - i + 1) = origTestSuits(5 - i)
                            origTestSuits(5 - i) = s
                            RemoveX qty(5 - i)
                            RemoveX hcp(5 - i)
                            x2nonX = True
                            Exit For
                        Else
                            j = j + 1 ' Number of suits where x is okay. '
                            x = 5 - i ' In case only 1 "x" suit is valid. '
                        End If
                    ElseIf i < k Then
                        str_Response = MyMsgBox("Do you want " & Mid$("CDHS", i, 1) & _
                                                " to have an 'x'?", _
                                                1, "YN", _
                                                "Mocking up entry for " & pb2)
                        If str_Response = "yes" Then
                            If qty(0) <> "" Then
                                t_Qty(5 - i) = "x"
                            End If
                            If hcp(0) <> "" Then
                                t_HCP(5 - i) = "x"
                            End If
                            j = j + 1
                            x = 5 - i
                        End If
                    End If
                Next
            End If

            If j = 1 Then
                ' Only one "x" suit is valid, so transfer the '
                ' "x" specs to it, blank out the "x" specs,   '
                ' and change the bid from "x" to the suit.    '
                qty(x) = qty(0)
                hcp(x) = hcp(0)
                qty(0) = ""
                hcp(0) = ""
                cmb_Bid = Left$(cmb_Bid, 1) & Mid$("SHDC", x, 1)
                bidConverted = True
                If t_Find_Ref <> "" Then
                    t_Ref = t_Find_Ref
                End If
                'Stop ' The suit length disclosure
                     ' needs to be changed here.
            Else
                t_Qty(0) = qty(0)
                t_HCP(0) = hcp(0)
            End If

     '  Convert suit bid:                                                          '

        ElseIf Len(cmb_Bid) = 2 And _
           Right$(cmb_Bid, 1) <> "N" _
        Then
            ' See if the suit bid in the original entry  '
            ' is the same as an original Prior Bid suit: '
            i = 1
            s = Right$(cmb_Bid, 1)
            Do While i <= Len(pb1)
                If Val(Mid$(pb1, i, 1)) > 0 Then
                    i = i + 1
                    If Mid$(pb1, i, 1) = s Then
                        Exit Do
                    End If
                    i = i + 2 ' Skip the dash. '
                Else
                    i = i + 1
                End If
                DoEvents
            Loop

            If Mid$(pb1 & "   ", i, 1) = s Then
                ' Suit found in Prior Bids matching suit bid. '
                ' Change the new bid to the suit at the same  '
                ' point in the new Prior Bids:                '
                b = Mid$(pb2 & " ", i, 1)
                If b <> "N" And b <> " " And Not bidConverted Then
                    cmb_Bid = Left$(cmb_Bid, 1) & Mid$(pb2, i, 1)
                    bidConverted = True
                End If

            Else
                z = Len(pb2)
                Do While Val(Mid$(" " & pb2, z + 1, 1)) = 0 And z > 1
                    z = z - 1
                Loop
                If Val(Mid$(" " & pb2, z + 1, 1)) > 0 Then
                    lastLevelBid = Val(Mid$(pb2, z, 1))
                    lastSuitBid = Mid$(pb2, z + 1, 1)
                End If

                For i = 0 To 4
                    a = ""
                    ' Suit bid in original entry did not match any of  '
                    ' the suits in the Prior Bids, e.g.: 1D-1S-2C - 2H,'
                    ' so the suit for the new entry's bid must also be '
                    ' different from other suits in *its* Prior Bids.  '
                    If i = 0 Then
                        ' First check if the suit bid in the source '
                        ' entry is in the second grid's Prior Bids. '
                        ' If not, then use it for the new entry too.'
                        If InStr(pb2, s) = 0 Then
                            j = 0
                            a = s
                        End If
                        i = 1
                    End If

                    If a = "" Then
                        ' Find the next suit above the last one bid. '
                        x = InStr("CDHS", lastSuitBid)
                        If x = 4 Then
                            b = "C"
                        Else
                            b = Mid$("CDHS", x + 1, 1)
                        End If
                        j = InStr(pb2, b)
                    End If

                    If j = 0 And Not bidConverted Then
                        If a <= lastSuitBid And _
                           Val(cmb_Bid) <= lastLevelBid _
                        Then
                            ' The computed bid is insufficient: '
                            cmb_Bid = (lastLevelBid + 1) & a
                            str_Response = MyMsgBox("The bid for the new entry had" & vbCr & _
                                                    "to be made at a higher level" & vbCr & _
                                                    "than the original entry.", 3, "ok", "BidBase Editor")

                        'ElseIf a > lastSuitBid And _
                        '       k > Val(cmb_Bid) Then
                            ' Created a jump when a jump was '
                            ' not in the original entry.     '
                        '    cmb_Bid = k - 1 & a
                        Else
                            'cmb_Bid = Left$(cmb_Bid, 1) & a
                        End If
                        bidConverted = True
                        Exit For
                    End If
                Next
            End If
        End If

     '  Convert suit specs:                                                        '

        If Len(t_Prior) > 1 Then
            ' For each suit in the original Prior Bids, '
            ' we assign its suit specs (Qty & HCP) to   '
            ' the suit in the same position in the new  '
            ' entry's Prior Bids. For example...        '
            '    Original Prior Bids:  1H-1S-2C         '
            '    New entry's Pr.Bids:  1C-1H-2D         '
            ' The Heart specs in the original entry are '
            ' assigned to Clubs, the Spade specs to     '
            ' Hearts, and the Club specs to Diamonds.   '
            ' At the same time, we make corresponding   '
            ' changes to the Test Hand and Suit Length  '
            ' disclosure fields.                        '

            unswapped(1) = "SHDC" ' Remove each suit letter as it is swapped.  '
            unswapped(2) = "SHDC" ' (1) is the source entry, (2) is the target.'
            p = ""
            i = 1
            Do While i < Len(pb1)
                If Val(Mid$(pb1, i, 1)) > 0 Then
                    ' Some # of a suit is bid. '
                    a = Mid$(pb1, i + 1, 1) ' e.g.: H '
                    b = Mid$(pb2, i + 1, 1) ' e.g.: C '
                    s1 = InStr("SHDC", a)   ' e.g.: 2 '
                    s2 = InStr("SHDC", b)   ' e.g.: 4 '

                    If InStr(p, a & b) = 0 And _
                       s1 > 0 And s2 > 0 And _
                       a <> "" And b <> "" _
                    Then
                        If qty(s1) <> "" Or _
                           hcp(s1) <> "" _
                        Then ' Do this even if a=b: '
                            t_Qty(s2) = qty(s1)
                            t_HCP(s2) = hcp(s1)

                            ' Change suits in Suit Length disclosure: '
                            If swapDistrib And a <> b Then
                                tmp = t_Len(s1)
                                t_Len(s1) = t_Len(s2)
                                t_Len(s2) = tmp
                            End If

                            ' Change suits in Combo specs: '
                            If cmQ <> "" Then
                                k = InStr(cmQ, a)
                                If k > 0 Then
                                    t_Qty(5) = Left$(t_Qty(5), k - 1) & _
                                               b & _
                                               Mid$(t_Qty(5), k + 1)
                                End If
                            End If

                            If cmH <> "" Then
                                k = InStr(cmH, a)
                                If k > 0 Then
                                    t_HCP(5) = Left$(t_HCP(5), k - 1) & _
                                               b & _
                                               Mid$(t_HCP(5), k + 1)
                                End If
                            End If

                        End If

                        ' Swap suit specs in Test Hand: '
                        newTestSuits(s2) = origTestSuits(s1)

                        ' Keep track of suits in original Prior Bids: '
                        p = p & a & b & " "
                        Mid$(unswapped(1), s1, 1) = " "
                        Mid$(unswapped(2), s2, 1) = " "

                        ' Check for suits compared to other suits: '
                        For k = 1 To 4
                            Select Case qty(k)
                            Case Is = ">" & a
                                qty(k) = ">" & b
                            Case Is = "<" & a
                                qty(k) = "<" & b
                            End Select
                        Next

                    End If

                    i = i + 3 ' Jump to where next number may be. '
                Else
                    i = i + 1
                End If
                DoEvents
            Loop

            For i = 1 To 4
                a = Mid$(unswapped(1), i, 1) ' e.g.: D '
                If a > " " Then
                    j = 1
                    Do While Mid$(unswapped(2), j, 1) = " " And j < 4
                        j = j + 1
                    Loop
                    b = Mid$(unswapped(2), j, 1) ' e.g.: S '
                    If b > " " Then
                        s1 = InStr("SHDC", a)    ' e.g.: 3 '
                        s2 = InStr("SHDC", b)    ' e.g.: 1 '

                        If t_Qty(s2) = "" And t_HCP(s2) = "" Then
                            t_Qty(s2) = qty(s1)
                            t_HCP(s2) = hcp(s1)
                            If Right$(cmb_Bid, 1) = a And _
                               Len(cmb_Bid) = 2 And _
                               Not bidConverted _
                            Then
                                cmb_Bid = Left$(cmb_Bid, 1) & b
                                bidConverted = True
                            End If
                        End If

                        ' Copy suit specs to Test Hand: '
                        newTestSuits(s2) = origTestSuits(s1)

                        ' Keep track so suits aren't swapped back: '
                        Mid$(unswapped(1), s1, 1) = " "
                        Mid$(unswapped(2), s2, 1) = " "

                        ' Change suits in Suit Length disclosure: '
                        If swapDistrib And a <> b Then
                            tmp = t_Len(s1)
                            t_Len(s1) = t_Len(s2)
                            t_Len(s2) = tmp
                        End If

                        If cmH <> "" Then
                            k = InStr(cmH, a)
                            If k > 0 Then
                                t_HCP(5) = Left$(t_HCP(5), k - 1) & _
                                           b & _
                                           Mid$(t_HCP(5), k + 1)
                            End If
                        End If

                        If cmQ <> "" Then
                            k = InStr(cmQ, a)
                            If k > 0 Then
                                t_Qty(5) = Left$(t_Qty(5), k - 1) & b & Mid$(t_Qty(5), k + 1)
                            End If
                        End If

                    End If

                    ' Check for suits compared to other suits: '
                    If a <> b Then
                        For k = 1 To 4
                            Select Case qty(k)
                            Case Is = ">" & b
                                qty(k) = ">" & a
                            Case Is = "<" & b
                                qty(k) = "<" & a
                            End Select
                        Next
                    End If
                End If
            Next

            ' If "x" is used but "x" is not bid, '
            ' (for example, see Random Splinter  '
            ' where 3N is bid when "x" is <2)    '
            ' then the "x" boxes have not been   '
            ' filled in yet, so do it now:       '
            If t_Qty(0) = "" And t_HCP(0) = "" Then
                t_Qty(0) = qty(0)
                t_HCP(0) = hcp(0)
            End If

            For i = 1 To 3
                t_TestHand = t_TestHand & newTestSuits(i) & "-"
            Next
            t_TestHand = t_TestHand & newTestSuits(4)

        End If

    Else ' Not "using" - just editing. '
        For i = 0 To 4
            t_Qty(i) = qty(i)
            t_HCP(i) = hcp(i)
        Next
        t_TestHand = !Test_Hand
    End If


End With

    ' Fill in DupeEntry Prior Bids boxes: '
    If gn = 0 And Not mDupesGrid(0).Checked And Not bln_AlreadyToldNoDupes And Grid1(1).Visible Then
        'FillDupeBoxes t_Prior, t_Round, t_Player, t_Order
    End If

    DoEvents

    If bln_Abort Then Exit Sub

    If t_TestHand = "" Then
        CreateTestHand
    End If
    If Index <> 99 Then mEdit(Index).Checked = False
    If Index = 0 And ck_SaveNew.Value = 0 Then
        ck_Saved.Value = 1
    Else
        ck_Saved.Value = 0
    End If
    str_OrigData = t_Round
    int_GridNum = gn
    ColorGrids gn
    On Local Error GoTo 0
    If t_Prior = "" And (t_Order = "130000" Or t_Order = "110000") And t_Hand_HCPs <> "" Then
        str_Response = MyMsgBox("This entry records the range of" & vbCr & _
                                "HCPs required for opening " & cmb_Bid & "." & vbCr & _
                                "No changes should be made to it" & vbCr & _
                                "other than the HCP range.", 4, "ok", "")
        t_Hand_HCPs.SetFocus
    End If
End Sub


Private Sub DisplayOtherSpecs(c As Integer, sID As Long)
    Dim r As Integer
    Dim speks As String
    
    Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
    Set rs_HandsData = db_BidBase.OpenRecordset("HandData")
    
    With rs_HandsData
    '''''''''''''''''
    .Index = "PrimaryKey"
    .Seek "=", sID
    If .NoMatch Then
        str_Response = MyMsgBox("ID#" & sID & " is not in database.", 1, "ok", "")
        Exit Sub
    End If
    speks = !specs
    For r = 1 To 12
        g_OtherHands.TextMatrix(r, c) = Trim(Mid$((r - 1) * 6 + r, 6))
    Next
    g_OtherHands.TextMatrix(11, c) = Trim(Mid$(72, 10))
    g_OtherHands.TextMatrix(12, c) = Trim(Right$(speks, 10))
    ' It's easier to search for one spec instead of 14 '
    ' so they are saved combined. Left here in case:   '
    'g_OtherHands.TextMatrix(2, c) = !hcp
    'g_OtherHands.TextMatrix(3, c) = !pts
    'g_OtherHands.TextMatrix(4, c) = !Cust_Pts
    'g_OtherHands.TextMatrix(5, c) = !Shape
    'g_OtherHands.TextMatrix(6, c) = !S_Num
    'g_OtherHands.TextMatrix(7, c) = !S_HCP
    'g_OtherHands.TextMatrix(8, c) = !H_Num
    'g_OtherHands.TextMatrix(9, c) = !H_HCP
    'g_OtherHands.TextMatrix(10, c) = !D_Num
    'g_OtherHands.TextMatrix(11, c) = !D_HCP
    'g_OtherHands.TextMatrix(12, c) = !C_Num
    'g_OtherHands.TextMatrix(13, c) = !C_HCP
    'g_OtherHands.TextMatrix(14, c) = !Combo_Num
    'g_OtherHands.TextMatrix(15, c) = !Combo_HCP
    
    ''''''''
    End With
End Sub


Private Sub RemoveX(spec As String)

    Dim i As Integer

    If Len(spec) > 1 Then

        If Left$(spec, 2) = "x," Then
            spec = Trim$(Mid$(spec, 3))

        ElseIf Right$(spec, 2) = ",x" Then
            spec = Left$(spec, Len(spec) - 2)

        ElseIf Right$(spec, 3) = ", x" Then
            spec = Left$(spec, Len(spec) - 3)

        Else
            i = InStr(spec, "x")
            If i > 0 Then
                Mid$(spec, i, 1) = " "
            End If
        End If
    Else
        spec = ""
    End If

End Sub


Private Sub FillDupeBoxes(ByVal pb As String, _
                          ByVal r As String, _
                          ByVal p As String, _
                          ByVal o As String)
    ' The table DupePBs in BidBase contains records  '
    ' in which similar Prior Bids are stored in each '
    ' field, up to 6 fields. In this routine, we     '
    ' search each field, pb1-pb6, for the current    '
    ' Prior Bids. If we find a record which has the  '
    ' current Prior Bids in any of the 6 records, we '
    ' display the Prior Bids in each of the other    '
    ' fields in the Dupe boxes. Example:             '
    '    1C-1D, 1C-1H, 1C-1S, 1D-1H, 1D-1S, 1H-1S    '
    ' If the current Prior Bids were 1D-1S, then we  '
    ' would display each of the other bids as Dupes. '
'Exit Sub
    Dim k As Integer
    Dim n As String
    Dim s As String
    Dim c As String

    pb = Trim$(pb)

    If pb = str_LastPBs And _
       bln_AlreadyToldNoDupes _
    Then
        Exit Sub
    End If

    str_LastPBs = pb
    bln_AlreadyToldNoDupes = False

    If Grid1(1).Visible And Grid1(1).Row > -1 And Grid1(0).Row > -1 Then
      If mDupesGrid(0).Checked = False Or _
         (Grid1(0).Columns(int_OrdNum_Col) <> _
          Grid1(1).Columns(int_OrdNum_Col)) _
      Then
        For k = 0 To 8
            t_DupeEntry(k) = ""
            b_DupeEntry(k).Visible = False
        Next
        'f_Dupes.Visible = False
      End If
    End If

    Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
    Set rs_DupePBs = db_BidBase.OpenRecordset(str_DupePBsTbl)
    If pb <> "" And _
       str_Job <> "Duping" _
    Then
        With rs_DupePBs
            k = 0
            ' Look for the Prior Bids in each field: '
            Do
                k = k + 1
                n = "pb" & k
                .Index = n
                .Seek "=", pb
            Loop While .NoMatch And k < 9

            If Not .NoMatch Then
                k = 0
                s = "AND (" ' Used in !RecordSource, below. '

                DisplayDupePB k, "" & !pb1, s, pb
                DisplayDupePB k, "" & !pb2, s, pb
                DisplayDupePB k, "" & !pb3, s, pb
                DisplayDupePB k, "" & !pb4, s, pb
                DisplayDupePB k, "" & !pb5, s, pb
                DisplayDupePB k, "" & !pb6, s, pb
                DisplayDupePB k, "" & !pb7, s, pb
                DisplayDupePB k, "" & !pb8, s, pb
                DisplayDupePB k, "" & !pb9, s, pb
                ' Note:  the ("" &) prevents a null error. '

                s = Left$(s, Len(s) - 4) & ") "
            End If
        End With

        If s = "" Then
            str_Response = MyMsgBox("No similar Prior Bids in" & vbCr & _
                                    "the database's 'Dupes' table.", 2, "ok", "")
            bln_AlreadyToldNoDupes = True
        End If

        ' Show Dupe records in 2nd grid: '
        If mDupesGrid(0).Checked Then
            If Not bln_TwoGrids Then
                mTwoGrids_Click
            End If

            If mDupesGrid(1).Checked = False Then
                s = ""
            End If

            If mDupesGrid(2).Checked Then
                c = "AND BidNames.Bid_Name = '" & _
                    Grid1(0).Columns(int_Bid_Name_Col) & "' "
            End If

            Data1(1).RecordSource = _
                qSel & _
                " WHERE Bids.Round = '" & r & "' AND " & _
                       "Bids.Player = '" & p & "' AND " & _
                       "Bids.Bid_Order_Number = '" & o & "' AND " & _
                       "Bids.Prior_Bids <> '" & t_Show(0) & "' " & s & _
                       c & _
                " ORDER BY BidNames.Bid_Name;"  ', Bids.PB_Sort_Order;"
            'On Local Error GoTo ExitSub
            Data1(1).Refresh
        End If
    End If
Exit Sub

ExitSub:
    str_Response = MyMsgBox("Error trying to reload bottom grid.", 1, "ok", "")
    On Local Error GoTo 0
End Sub


Private Sub DisplayDupePB(ByRef k As Integer, _
                          ByVal p As String, _
                          ByRef s As String, _
                          ByVal pb As String)
    If p <> pb And p <> "" Then
        ' This routine is called by Grid1_Click as well as '
        ' when an entry is being copied to the Input Boxes.'
        t_DupeEntry(k) = p
        b_DupeEntry(k).Visible = True
        If k = 5 Then
            'f_Dupes.Visible = True
        End If
        k = k + 1
        s = s & "Prior_Bids = '" & p & "' OR "
    End If
End Sub


Private Sub DisplayGarbage()
    ' Shows entries which have been deleted or changed   '
    ' in the "Garbage Grid" so that the user can recover '
    ' recover them if desired. '

    Dim gn As Integer
    Dim gID As Long
    Dim i As Integer
    Dim f As String
    Dim p As String
    Dim j As Integer

    mClear_Click
    gn = 2
    gID = Val(Grid1(gn).Columns(int_LastBidCol).Text)
    If gID = 0 Then
        str_Response = MyMsgBox("No record selected.", 1, "ok", "Display Garbage")
    End If

    Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
    Set rs_DelTable = db_BidBase.OpenRecordset("Garbage")
    With rs_DelTable
    '              '
    .Index = "ID"
    .Seek "=", gID
    If .NoMatch Then Exit Sub
    t_Pct = !Pct_Used

    t_Round = !Round_of_Bidding
    t_Player = !player

    t_Order = !Bid_Order_Number

    If !ConvID = int_DividerLineID Then
        ck_SaveNew.Value = 1
    Else
        ck_SaveNew.Value = 0
    End If

    t_Prior = Trim$(!Prior_Bids)

    cmb_Bid.Text = Trim$(!bid)
    If InStr(cmb_Bid.Text, "@") > 0 Then
        ' Slam exploring bid made.  '
        ' Increase cmb_Bid.Size and '
        ' location to make room.    '
        If cmb_Bid.Width <> 1015 Then
            t_Prior.Width = 3380
            t_Order.Left = 4990
            cmb_Bid.Left = 5835
            cmb_Bid.Width = 1050
        End If
    ElseIf cmb_Bid.Width = 1050 Then
        t_Prior.Width = 3765
        t_Order.Left = 5375
        cmb_Bid.Left = 6240
        cmb_Bid.Width = 650
    End If

    Dim num As Long
    Dim nam As String
    Dim sel As String
    Dim sel2 As String
    nam = "" & !ConvID
    If nam <> "" Then
        LookUpBidName num, sel, sel2, nam
    End If
    'cmb_ConvName(0).ItemData(cmb_ConvName(0).ListIndex) = num
    t_Sel = sel
    t_Sub = sel2
    If nam <> "" Then
        ' Look up name in list: '
        i = FindMatch(cmb_ConvName(0), nam, True) ' True = exact match
        If i < 0 Then i = 0
    End If
    cmb_ConvName(0).ListIndex = i

    ' Look for file for convention name: '
    If !ConvID = int_DividerLineID Then
        Exit Sub

    Else
        f = Trim$(cmb_ConvName(0).Text)
        i = InStr(f, ":")
        If i > 0 Then
            f = Left$(f, i - 1)
        End If

        For i = 1 To Len(f)
            If InStr("/\:.?^*<>| ", Mid$(f, i, 1)) > 0 Then
                Mid$(f, i, 1) = "_"
            End If
        Next
        p = str_BBpath & "\notes\" & f & ".htm"
        If Dir$(p) = "" Then
            b_Conven_Description.Picture = _
                LoadPicture(str_BBpath & "\ButtonN.bmp")
        Else
            b_Conven_Description.Picture = _
                LoadPicture(str_BBpath & "\ButtonY.bmp")
        End If
    End If


    Dim a As String
    Dim c As Integer

    'Vul
    MsgBox "Error" ' Should be U F B N V X
    a = UCase$(Trim$(!Vul))
    i = InStr("Y    N    FAV  UNFAVEQ   EQ+", a) / 5
    i = i + 1
    If a = "" Then i = 0
    cmb_Vul.ListIndex = i

    'Pos
    a = LCase$(Trim$(!pos))
    i = InStr("ft1234ynpu", a)
    If a = "" Then i = 0
    cmb_Position.ListIndex = i

    'HCP
    t_Hand_HCPs = !hcp
    c = InStr(t_Hand_HCPs, ",")
    If c > 0 Then
        t_BBPoints = Mid$(t_Hand_HCPs, c + 1)
        t_Hand_HCPs = Left$(t_Hand_HCPs, c - 1)
    End If

    'Pts
    t_Distrib_Pts = !Dist_Pts

    'Custom Points
    t_CustPoints = "" & !Cust_Pts
    t_CustPts(0) = "" & !A_Cust
    t_CustPts(1) = "" & !K_Cust
    t_CustPts(2) = "" & !Q_Cust
    t_CustPts(3) = "" & !J_Cust
    t_CustPts(4) = "" & !T_Cust

    'Winrs
    t_Winners = !Winrs

    'Losers
    t_Losers = !losers

    'QTs
    t_QTs = !QTs

    'Stopped_Suits
    t_Stoppers = !Stopped_Suits

    'Shape
    Select Case UCase$(Trim$(!Shape))
        Case "1"
            cmb_Shape.ListIndex = 1
        Case "2"
            cmb_Shape.ListIndex = 2
        Case "3"
            cmb_Shape.ListIndex = 3
        Case "B"
            cmb_Shape.ListIndex = 4
        Case "U"
            cmb_Shape.ListIndex = 5
        Case "S"
            cmb_Shape.ListIndex = 6
        Case "V"
            cmb_Shape.ListIndex = 7
        Case "Y", "E"
            cmb_Shape.ListIndex = 8
        Case Else
            cmb_Shape.ListIndex = 0
    End Select

    t_Intermediates = "" & !intermed

    t_Qty(0) = !xNum
    t_HCP(0) = !xHCP
    t_Qty(1) = !sNum
    t_HCP(1) = !sHCP
    t_Qty(2) = !hNum
    t_HCP(2) = !HHCP
    t_Qty(3) = !dNum
    t_HCP(3) = !dHCP
    t_Qty(4) = !cNum
    t_HCP(4) = !cHCP
    t_Qty(5) = !cmNum
    t_HCP(5) = !cmHCP

    t_TestHand = !Test_Hand

    ' Alert:
    i = InStr(" ASD", UCase$(Trim$(!Alert)))
    If i > 0 Then i = i - 1
    cmb_Alert.ListIndex = i

    ' Artificial:
    If !Artificial = "A" Then
        ck_Artificial.Value = 1
    Else
        ck_Artificial.Value = 0
    End If

    ' Explanation:
    t_Disclosure = !Explanation

    If Trim$(t_Prior) <> "" Then
        f = Trim$(t_Prior) & "-" & cmb_Bid
    Else
        f = cmb_Bid
    End If

    p = str_BBpath & "\Notes\"
    f = f & "_ID" & !id & ".txt"
    If Dir$(p + f) = "" Then
        ' See if file is referenced in notes:
        i = InStr(LCase$(t_Disclosure), ".txt")
        If i = 0 Then i = InStr(LCase$(t_Disclosure), ".htm")
        If i > 0 Then
            j = i + 4
            Do
                i = i - 1
            Loop While i > 1 And _
                       Mid$(t_Disclosure, i, 1) <> " "

            f = Trim$(Mid$(t_Disclosure, i, j - i))
        End If
    End If

    If Dir$(p + f) <> "" Then
        b_Discl.Picture = LoadPicture(str_BBpath & "\ButtonY.bmp")
    Else
        b_Discl.Picture = LoadPicture(str_BBpath & "\ButtonN.bmp")
    End If

    t_Ref = !source

    t_AsksFor = "" & !Asks
    L_Date_Saved = "" & !Date_of_Entry
    t_BestContract = "" & !Best_Contract
    
    ' Add display hand data here: '
    
    ck_SaveNew.Value = 1
    '      '
    End With
End Sub


Private Sub mExit_Click()
    str_Response = MyMsgBox("End the program?", 1, "YN", "BidBase")
    If str_Response = "yes" Then
        Unload Me
    End If
End Sub


Private Sub mHelp_Click()
    str_NoteFile(0) = "CardShark_BidBase_Editor_Overview.htm"
    ShowDocs
End Sub
    
    
Private Sub ShowDocs()
    Dim xx As Long

    ChDir str_BBpath
    str_PgmPath(0) = str_BBpath & "\Notes\"
    xx = Shell(str_PgmPath(0) & "BidBase File Viewer.exe " & _
               str_PgmPath(0) & str_NoteFile(0), vbNormalFocus)
End Sub


Private Sub mHideCol_Click()
    Dim c As Integer
    Dim s As Integer

    If frm_Tutorial.Visible Then
        If b_Continue.Visible = False Then
            Run_HidingColumns_Tutorial
            Exit Sub
        End If
    End If

    If mShowAllCols.Checked Then
        str_Response = MyMsgBox("To hide columns, first turn off" & vbCr & _
                                "'View - Show All Columns' (Ctrl-A).", 2, "ok", "")
        Exit Sub
    End If

    c = Grid1(int_GridNum).Col
    s = Grid1(int_GridNum).Split
    On Local Error GoTo ExitHideCol
    If int_GridNum = 0 Then
        If s = 0 Then
            cmb_HiddenCols(0).AddItem _
                Right$("0" & c, 2) & _
                " - " & _
                Grid1(0).Splits(0).Columns(c).Caption
        Else
            cmb_HiddenCols(1).AddItem _
                Right$("0" & c, 2) & _
                " - " & _
                Grid1(0).Splits(1).Columns(c).Caption
        End If
    Else
        If s = 0 Then
            cmb_HiddenCols(2).AddItem _
                Right$("0" & c, 2) & _
                " - " & _
                Grid1(1).Splits(0).Columns(c).Caption
        Else
            cmb_HiddenCols(3).AddItem _
                Right$("0" & c, 2) & _
                " - " & _
                Grid1(1).Splits(1).Columns(c).Caption
        End If
    End If

    With Grid1(int_GridNum).Splits(Grid1(int_GridNum).Split)
        .Columns(c).Visible = False

        Do While .Columns(c).Visible = False _
             And c < int_LastBidCol

            c = c + 1
        Loop
    End With

    Grid1(int_GridNum).Col = c

ExitHideCol:
    On Local Error GoTo 0
End Sub


Private Sub Run_HidingColumns_Tutorial()
    frm_Tutorial.Caption = "Hiding Columns"
    L_Tutorial = "Look at the View menu and you'll see several " & _
                 "menu choices related to hiding columns. Because the " & _
                 "grid is wider than most monitors can display all at " & _
                 "once, you have to scroll side-to-side to see all the " & _
                 "fields for a record. " & _
                 vbCr & vbCr & _
                 "One way to see more of a record is to hide columns " & _
                 "which are rarely used or which you don't need to see " & _
                 "all the time. Click on a column and press Ctrl-H." & _
                 vbCr & vbCr & _
                 "To unhide columns, click on View - Unhide and a window " & _
                 "will pop up with drop-down lists of hidden columns. " & _
                 "Select columns to unhide and click the button." & _
                 vbCr & vbCr & _
                 "You can temporarily view All columns at once which have " & _
                 "been hidden, by pressing Ctrl-A or clicking View - Show All. " & _
                 "If no columns have been hidden, this command has no effect.  "

    frm_Tutorial.Move f_Unhide.Left + f_Unhide.Width + 100, _
                      f_Unhide.Top

    L_Tutorial.Height = 4700
    ResizeTutorial
    int_GridNum = 0
    ColorGrids 0
End Sub


Private Sub Run_HideInputs_Tutorial()
    frm_Tutorial.Caption = "Hide/Unhide Input Boxes"
    L_Tutorial = "The menu item 'View - Input Boxes' hides/unhides the " & _
                 "Input Boxes at the bottom of the screen." & _
                 vbCr & vbCr & _
                 "Hiding them gives you more room for viewing the grid(s)." & _
                 vbCr & vbCr & _
                 "Try clicking the menu to hide/unhide the Input Boxes."

    ResizeTutorial
End Sub


Private Sub mSave_Click()
    Dim pos As String ' For position set-up before testing. '
    Dim r As Integer
    Dim p As Integer
    Dim d As Integer
    Dim g As Integer ' grid # '
    Dim c As String
    Dim pb As String
    Dim cd As String
    Dim p0 As String
    Dim p1 As String
    Dim p2 As String
    Dim p3 As String
    Dim tempID As Long
    Dim topID As Long
    Dim i As Integer
    Dim j As String
    Dim Prior_Bids As Index
    Dim ok As Boolean

    If t_Ref <> "" Then
        j = InStr(t_Ref, " ")
        If j > 0 Then
          For i = 0 To list_Refs.ListCount - 1
            If t_Ref = list_Refs.List(i) Then
                Exit For
            End If
            If t_Ref = list_Refs.List(i) Then
                ' We know from above that the entries '
                ' don't match exactly, or if they do  '
                ' match at the start, we can delete   '
                ' the entry in the list. '
                list_Refs.RemoveItem i
                i = list_Refs.ListCount
            End If
            Do While i < list_Refs.ListCount - 1
                If list_Refs.List(i) = list_Refs.List(i + 1) Then
                    list_Refs.RemoveItem i
                Else
                    Exit Do
                End If
            Loop
          Next
          'If i = list_Refs.ListCount Then
          '  list_Refs.AddItem t_Ref.Text
          'End If
        End If
    End If

    If frm_Tutorial.Visible Then
        If b_Continue.Visible = False Then
            frm_Tutorial.Caption = "Save Entry"
            Run_Save_Tutorial
        End If
        Exit Sub
    End If

    Dim mboxTop As Integer
    Dim mboxLeft As Integer
    mboxTop = f_Msg.Top
    mboxLeft = f_Msg.Left
    f_Msg.Top = frame0(6).Top - 1700
    'topID =
    f_Msg.Left = b_SaveEntry.Left + frame0(6).Left - 200

    If Trim$(t_Round) = "" Then
        str_Response = MyMsgBox("You must enter a bidding round number.", 1, "ok", "")
        t_Round.SetFocus
        GoTo ExitSave

    ElseIf Trim$(t_Player) = "" Then
        str_Response = MyMsgBox("You must enter a player number," & vbCr & _
                                "starting with the first non-passing bidder.", 2, "ok", "")
        t_Player.SetFocus
        GoTo ExitSave

    ElseIf Trim$(t_Order) = "" Then
        str_Response = MyMsgBox("You must enter a bid order number.", 1, "ok", "")
        t_Order.SetFocus
        GoTo ExitSave
    End If

    pb = t_Prior
    If pb <> "" Then
        ValidateBidding pb, p, r, ok
        p = p + 1 ' The above calculates last bidder. "
        If p = 5 Then
            p = 1
            r = r + 1
        End If
        If Not ok Then
            GoTo ExitSave
        End If
    Else
        p = Val(t_Player)
        r = Val(t_Round)
    End If

    If r <> Val(t_Round) Then
        str_Response = MyMsgBox("Round # entered (" & t_Round & ") does not " & vbCr & _
                                "match # of rounds in Prior Bids (" & r & ").", 2, "ok", "")
        t_Round.SetFocus
        GoTo ExitSave
    End If

    If p <> Val(t_Player) Then
        str_Response = MyMsgBox("Player # entered (" & t_Player & ") does not " & vbCr & _
                                "match bidder number after Prior Bids (" & p & ").", 2, "ok", "")
        t_Player.SetFocus
        GoTo ExitSave
    End If

    ' Make sure that if the cursor is in '
    ' an Input text box, its Validate    '
    ' routine gets executed before the   '
    ' entry is saved. t_Sel is a safe box'
    ' to set focus to since it doesn't   '
    ' have complex calculations in its   '
    ' Validate routine.                  '
    str_OrigData = ""
    t_Sel.SetFocus

    ' Though the Data Controls, and thus the grids, '
    ' use "BidNames Query" as their RecordSource,   '
    ' updates can be done directly to the Bids table:

    '                                                                   '
    Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
    Set rs_BidTable = db_BidBase.OpenRecordset(str_BidsTbl)

    With rs_BidTable
    '                '
    If ck_SaveNew.Value = 0 Then
        If lng_ID = 0 Then
            str_Response = MyMsgBox("You are not editing an entry." & vbCr & _
                                    "To save as new, check 'Save As New Entry'.", 2, "ok", "")
            GoTo ExitSave
        End If
        .Index = "PrimaryKey"
        .Seek "=", lng_ID

        ' Use .Edit to Save edited entry: '
        If .NoMatch = True Then
            str_Response = MyMsgBox("ERROR:" & vbCr & _
                                    "The entry's ID# is not in the database." & vbCr & _
                                    "The entry is not being saved.", 3, "ok", "")
            GoTo ExitSave
        End If
        If str_Source <> "Importing" Then Copy2DeletedTable
        .Edit
    Else 'Use .AddNew to save new entry: '
        .AddNew
    End If

SaveEntry:
    !Pct_Used = t_Pct
    !RoundOfBidding = t_Round
    !player = t_Player
    !Prior_Bids = t_Prior
    !Bid_Order_Number = t_Order
    !bid = cmb_Bid

    Dim num As Long
    Dim sel As String
    Dim sel2 As String
    LookUpBidName num, sel, sel2, cmb_ConvName(0)
    If t_Order = "ZZZZZZ" Then
        !ConvID = int_DividerLineID
    Else
        !ConvID = num
    End If

    If cmb_Vul.ListIndex > 0 Then
        'Stop ' should be U F B N V X  Unfav, Fav, Both, Neither, this hand Vul, X=this hand not vul
        !Vul = Trim$(Left$(cmb_Vul & " ", 1))
'        !Vul = _
            Trim$(Mid$("Y    N    Fav  UnfavEq   Eq+  ", _
                       (cmb_Vul.ListIndex) * 5 + 1, 1))

    Else
        !Vul = ""
    End If

    If cmb_Position.ListIndex > 0 Then
        !pos = _
            UCase$(Trim$(Mid$("ft1234ynpu", cmb_Position.ListIndex, 1)))
    Else
        !pos = ""
    End If

    ' Rather than reworking the grid '
    ' to add a BidBase Val Points    '
    ' field, I added an input box    '
    ' and am putting the BB Points   '
    ' in with the Hand HCP field with'
    ' a comma separating them. '
    If t_BBPoints = "" Then
        !hcp = t_Hand_HCPs
    Else
        !hcp = t_Hand_HCPs & "," & t_BBPoints
    End If
    
    !Dist_Pts = t_Distrib_Pts

    ' Custom points are strings to '
    ' allow for things like <2. '
    !Cust_Pts = t_CustPoints
    !A_Cust = t_CustPts(0)
    !K_Cust = t_CustPts(1)
    !Q_Cust = t_CustPts(2)
    !J_Cust = t_CustPts(3)
    !T_Cust = t_CustPts(4)

    !Winrs = t_Winners
    !losers = t_Losers
    !QTs = t_QTs
    !Stopped_Suits = t_Stoppers
    If cmb_Shape.ListIndex = 8 Then
        !Shape = "E"
    Else
        !Shape = Trim$(Left$(cmb_Shape & " ", 1))
    End If
    !intermed = Trim$(t_Intermediates)

    !xNum = t_Qty(0)
    !xHCP = t_HCP(0)
    !sNum = t_Qty(1)
    !sHCP = t_HCP(1)
    !hNum = t_Qty(2)
    !HHCP = t_HCP(2)
    !dNum = t_Qty(3)
    !dHCP = t_HCP(3)
    !cNum = t_Qty(4)
    !cHCP = t_HCP(4)
    !cmNum = t_Qty(5)
    !cmHCP = t_HCP(5)

    ' Disclosures:
    Select Case cmb_Alert.ListIndex
        Case 1
            !Alert = "A"
        Case 2
            !Alert = "S"
        Case 3
            !Alert = "D"
        Case Else
            !Alert = ""
    End Select

    If cmb_Alert = "Continuation" Then
        !Alert = "C"
    End If

    !Explanation = Trim$(t_Disclosure) ' Database is set for 50 char. Why is it choking on 10? '
    !Test_Hand = t_TestHand
    str_TestingHand = t_TestHand

    If ck_Artificial.Value Then
        !Artificial = "A"
    Else
        !Artificial = ""
    End If

    !source = t_Ref

    !Asks = Trim$(t_AsksFor)
    !DateOfEntry = Date$
    L_Date_Saved = !DateOfEntry
    !Best_Contract = t_BestContract
    !PB_Sort_Order = PB_To_SortOrd_Code$(Trim$(t_Prior))
    
    ' Put HandData IDs here: '
    ' Search for an entry matching any specs '
    ' entered for LHO and save tha ID#, then '
    ' repeat for Pard and RHO. If no entry is'
    ' found for a player, create a new one.  '
    If t_OtherHands.Visible Then
        t_OtherHands.Visible = False
        !LHO = Val(g_OtherHands.TextMatrix(1, 2))
        !Pard = Val(g_OtherHands.TextMatrix(1, 3))
        !RHO = Val(g_OtherHands.TextMatrix(1, 4))
    End If
    
    If !ConvID = 0 Then !ConvID = 4

    .Update

    If !id = 0 Then
        ' Some kind of error. '
        str_Response = MyMsgBox("Entry does not have an ID#.", 1, "ok", "")
        Exit Sub
    End If

    L_IDnum = "ID#: " & !id

    If str_Source = "Importing" Then Exit Sub

    .Bookmark = .LastModified

    lng_ID = !id
    cd = !PB_Sort_Order
    '        '
     End With
    '                                                                         '

    On Local Error Resume Next
    
    If t_Show(1) = t_Prior Then
        ' Entry being saved is in the top grid. '
        g = 0
    ElseIf t_Show(2) = t_Prior Then
        ' Entry being saved is in the bottom grid. '
        g = 1
    End If

    topID = Val(Grid1(g).Columns(int_LastBidCol).Text)
    i = Grid1(g).FirstRow
    
    p0 = PB_To_SortOrd_Code(Trim$(t_Show(0 + g)))
    p1 = PB_To_SortOrd_Code(Trim$(t_Show(1 + g)))
    Data1(g).Refresh
    DoEvents
    Grid1(g).Refresh
    DoEvents

    'Do While Grid1(g).FirstRow < i
    '    DoEvents
    '    Grid1(g).Row = Grid1(g).Row + 1
    '    DoEvents
    'Loop
    
    GoTo TestEntry
    Exit Sub
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'If cd >= p0 And cd <= p1 Then
    If Grid1(0).VisibleRows > 1 Then
        ' Refreshing moves the pointer to row 0. '
        ' Move the pointer to the new entry:     '
        Do While t_Show(0) = Grid1(0).Columns(6) And _
                 Grid1(0).VisibleRows > 1 And Grid1(0).Columns(int_LastBidCol).Text <> ""
            ' This loop repeats when Grid1(0).VisibleRows = 1 despite this check being above. '
            If topID = Val(Grid1(0).Columns(int_LastBidCol).Text) Then
                Exit Do
            ElseIf lng_ID = Val(Grid1(0).Columns(int_LastBidCol).Text) Then
                Exit Do
            ElseIf Grid1(0).VisibleRows = 1 Then
                Exit Do
            ElseIf Grid1(0).VisibleRows = Grid1(0).Row + 1 Then
                Exit Do
            End If
            Grid1(0).Row = Grid1(0).Row + 1
            DoEvents
        Loop
    End If
    'End If
    
On Local Error Resume Next
    If mDupesGrid(0).Checked Then
        Data1(1).Refresh

    ElseIf 1 = 2 And Grid1(1).Visible Then
        p2 = PB_To_SortOrd_Code$(Trim$(t_Show(2)))
        p3 = PB_To_SortOrd_Code$(Trim$(t_Show(3)))
        If cd >= p2 And cd <= p3 Then
            Data1(1).Refresh
            DoEvents
            Grid1(1).Refresh
            DoEvents
            Do While lng_ID <> Grid1(1).Columns(int_LastBidCol).Text And _
                     Grid1(1).Columns(int_LastBidCol).Text <> ""
                If Grid1(1).VisibleRows = 1 Then
                    Exit Do
                ElseIf Grid1(1).VisibleRows = Grid1(1).Row + 1 Then
                    Exit Do
                End If
                Grid1(1).Row = Grid1(1).Row + 1
                DoEvents
            Loop
        End If
    End If

'                                                                 '
 With Grid1(int_GridNum)
'                       '
    If .Row > -1 Then
    If Not (cd >= p0 And cd <= p1) And _
       Not (cd >= p2 And cd <= p3) And _
       mDupesGrid(0).Checked = False _
    Then
        ' Was in int_GridNum, but no longer in either grid, '
        ' so refresh int_GridNum to get entry out of grid:  '
            .Row = .Row + 1
            tempID = .Columns(int_LastBidCol).Text
            Data1(int_GridNum).Refresh
            Do While tempID <> .Columns(int_LastBidCol).Text And _
                     .Columns(int_LastBidCol).Text <> ""
                If .VisibleRows = .Row + 1 Then Exit Do
                .Row = Grid1(int_GridNum).Row + 1
                DoEvents
            Loop
    End If
    End If

    Grid1(0).Enabled = True ' True
    Grid1(1).Enabled = True ' True

TestEntry:
'''''''''
    If bln_UsingEntry Then
        bln_UsingEntry = False
    End If

    On Local Error GoTo 0
    ck_SaveNew.Value = 0

    If Trim$(t_TestHand) <> "" Then
        str_Job = "Saving"
        ' When testing a bid entry, if it specifies '
        ' a position such as 3 (seat 3), it cannot  '
        ' pass the test because there are no leading'
        ' passes in the editor (unlike the Practice '
        ' program), so we have to add leading passes'
        ' to str_PriorCalls for Sub TestHand to send'
        ' to the Bidder. '
        ' Also, whether a hand is passed or not only'
        ' applies to the first round of bidding, so '
        ' if there is more than 1 round of bidding, '
        ' we have to go back to the first round to  '
        ' see if he (or his partner) passed, though '
        ' it would seem unusual for the influence of'
        ' being a passed hand to extend beyond two  '
        ' rounds. '
        pos = Left$(cmb_Position & " ", 1)
        p = Val(t_Player)
        r = Val(t_Round)
        str_PriorCalls = t_Prior

        If pos = "2" Then
            If str_PriorCalls = "" Then ' second seat - RHO passed. '
                str_PriorCalls = "P"
            Else
                'Stop
            End If
        ElseIf pos = "3" Or pos = "T" Then
            If str_PriorCalls = "" Then
                str_PriorCalls = "P-P"
            Else
                'Stop
            End If
        ElseIf pos = "4" Then
            If str_PriorCalls = "" Then
                str_PriorCalls = "P-P-P"
            Else
                'Stop
            End If
        ElseIf pos = "Y" Then
            ' This is a passed hand: '
            ' Add 5 - p # of "P-" : '
            For i = 1 To 5 - p
                str_PriorCalls = "P-" & str_PriorCalls
            Next
        ElseIf pos = "P" Then
            'Stop
        ElseIf pos = "U" Then
            'Stop
        End If

        If t_Show(0) <> t_Prior And t_Show(2) <> t_Prior Then
            t_Show(0) = t_Prior
            t_Show(1) = t_Prior
        End If
        'b_Show_Click : '
        Timer3.Enabled = True
        DoEvents

        bln_TestingSaveEntry = True
        TestHand
        bln_TestingSaveEntry = False
        str_Job = ""
    ElseIf t_Order <> "ZZZZZZ" Then
        str_Response = MyMsgBox("Entry saved.", 1, "ok", "")
    End If

    'If lng_DupeID > 0 And _
    '   str_DupePrior <> "" And _
    '   Grid1(g).Row > 0 And _
    '   b_DupeEntry(0).Visible And _
    '   str_Job <> "Duping Error" _
    'Then
    '    int_GridNum = int_DupeGrid
    '    If .Columns(int_LastBidCol).Text <> lng_DupeID And _
    '       .Row > 0 _
    '    Then
    '        .Row = .Row - 1
    '    End If
    '    j = str_Job
    '    str_Job = "Duping"
    '    mEdit_Click 0
    '    str_Job = j
    '    If str_Job = "Duping" Then str_Job = ""
    'End If

    If mDupesGrid(0).Checked And int_GridNum = 0 Then
        ' Click on new entry to cause Dupes to display: '
        Grid1_Click 0
    End If

'        '
 End With
'                                                              '

ExitSave:
    If 1 = 2 Then
        ' This is code to get the entry being saved to show in the grid. '
        ' This is also done in Timer3 which is called in several places. '
        If cd >= p2 And cd <= p3 Then
            Data1(1).Refresh
            DoEvents
            Grid1(1).Refresh
            DoEvents
            Do While lng_ID <> Grid1(1).Columns(int_LastBidCol).Text And _
                     Grid1(1).Columns(int_LastBidCol).Text <> ""
                If Grid1(1).VisibleRows = 1 Then Exit Do
                If Grid1(1).VisibleRows = Grid1(1).Row + 1 Then Exit Do
                Grid1(1).Row = Grid1(1).Row + 1
                DoEvents
            Loop
        End If
    End If
    
    f_Msg.Top = mboxTop
    f_Msg.Left = mboxLeft
End Sub


Sub ValidateBidding(pb As String, p As Integer, r As Integer, ok As Boolean)
    ' p and r are for returning player# and round#. '
    ' Calling subs must increments these for next bidder. '
    Dim LB As Integer ' LB = last bid = 1C=11, 2C=12, 1D=21, 3D=23, 1H=31, 4H=34, 2S=42 '
    Dim sLB As String ' sLB = "1C
    Dim LB2 As Integer
    Dim sLB2 As String
    Dim HB As Integer ' HB where 1st digit = suit number (CDHSN), 2nd = level bid: 1-7. '
    Dim dbl As Boolean
    Dim d As Integer
    Dim c As String
    Dim n As Boolean
    Dim lenPB As Integer
    Dim bidNum As Integer
    Dim e As String ' error message '
    Dim i As Integer

    pb = Trim$(pb) ' Example: 1D-P-1H-2S-3H-P-3S-P-4N@30-P '
    Do While Left$(pb & "  ", 2) = "P-"
        pb = Mid$(pb, 3)
    Loop
    Erase str_Bids()
    bidNum = 0
    d = 1
    LB = 0
    Do
        LB = InStr(LB + 1, pb, "-")              '     lb= 3  6  8 '
        If LB = 0 Then
            bidNum = bidNum + 1                  ' 4 '
            str_Bids(bidNum) = Mid$(pb, d)           ' 2D  '
            Exit Do
        End If
        ' Finding a dash counts the next bid. '
        bidNum = bidNum + 1                      ' bidNum= 1   2  3 '
        str_Bids(bidNum) = Mid$(pb, d, LB - d)       ' bid$()= 1D 1S  D '
        ' 1D-D-P-P-R '
        ' 123456789-  4-1=(1)1D, d=3, 5-3=2
        d = LB + 1                               '      d= 4   7  9 '
    Loop
    
    d = Len(pb)
    If d > 1 Then
        If Right$(pb, 1) = "-" Then
            pb = Left$(pb, d - 1)
        End If
    End If

    ' Reset to True after testing passes. '
    ok = False

    If str_Bids(1) = "P" And t_Prior <> "" Then
        ' Only applies to t_Prior. '
         str_Response = MyMsgBox("Prior Bids cannot start with a P (Pass)." & vbCr & _
                                 "Start with the first non-Passing bid." & vbCr & _
                                 "If position matters, use the Position box.", 3, "ok", "")
        Exit Sub
    End If

    If Val(str_Bids(1)) > 0 Then
        'LB = Val(bids(1)) * 10 + InStr("CDHSN", Mid$(bids(1), 2, 1))
    ElseIf pb <> "" Then
        If t_Prior <> "" Then
            str_Response = MyMsgBox("Prior Bids must start with" & vbCr & _
                                    "a bid such as 1C.", 2, "ok", "")
            Exit Sub
        Else
            str_Response = MyMsgBox("Use the Position box" & vbCr & _
                                    "to indicate opening passes.", 2, "ok", "")
        End If
    End If

    r = (bidNum - 1) \ 4 + 1
    p = bidNum Mod 4
    If p = 0 Then p = 4
    e = ""

    Dim pb1 As String ' prior bid 1 back '
    Dim pb2 As String ' prior bid 2 back '
    Dim pb3 As String ' prior bid 3 back (LHO) '

    For d = 1 To bidNum
        c = str_Bids(d)
        If c = "R" Then
            If d = 4 Then
                If str_Bids(d - 1) <> "D" And _
                   (str_Bids(d - 3) <> "D" Or _
                    str_Bids(d - 2) <> "P" Or _
                    str_Bids(d - 1) <> "P") _
                Then
                    e = "x"
                End If
            ElseIf d = 3 Or d = 2 Then
                If str_Bids(d - 1) <> "D" Then
                    e = "x"
                End If
            ElseIf d = 1 Then
                e = "x"
            End If

            If e = "x" Then
                e = "A redouble was made without a preceding double."
                Exit For
            End If
        End If

        If c = "D" Then
            If d = 1 Then
                e = "x"
            ElseIf d = 2 Or d = 3 Then  ' ex: P-D or 1C-P-D '
                If Val(str_Bids(d - 1)) = 0 Then e = "x"
            ElseIf d = 4 Then ' ex: P-1C-P-D or 1C-1D-D-D ' Ok: P-P-1C-D or 1C-P-P-D '
                If Val(str_Bids(d - 1)) = 0 And _
                   (Val(str_Bids(d - 3)) = 0 Or _
                    str_Bids(d - 2) <> "P" Or _
                    str_Bids(d - 1) <> "P") _
                Then
                    e = "x"
                End If
            End If

            If e = "x" Then
                e = "A Double was made illegally."
                Exit For
            End If
        End If

        LB2 = Val(str_Bids(d)) * 10 + InStr("CDHSN", Mid$(str_Bids(d), 2, 1))
        If LB2 > 0 Then ' Not P, D, R. '
            If LB2 <= HB And LB2 > 1 Then
                e = str_Bids(d) & " is an insufficient bid."
                Exit For
            End If
            If LB2 > HB Then HB = LB2
        End If
        DoEvents
    Next
    If e <> "" Then
         str_Response = MyMsgBox(e, 1, "ok", "")
    Else
        ok = True
    End If
End Sub


Private Sub AddPassesForPosition(pos As String)
    Dim i As Integer
    Dim dashNum As Integer
    Dim numDashes As Integer
    Dim bidder As Integer

    If str_PriorCalls = "" Then
        If pos = "2" Then
            str_PriorCalls = "P"
        ElseIf pos = "3" Or pos = "T" Or pos = "P" Then ' or bidder's partner passed '
            str_PriorCalls = "P-P"
        ElseIf pos = "4" Then
            str_PriorCalls = "P-P-P"
        End If
        Exit Sub
    End If

    ' There were some prior calls: '
    i = InStr(str_PriorCalls, "-")
    If i = 0 Then
        ' No dash, so only 1 prior bid. '
        If pos = "2" Then
            str_PriorCalls = "P-" & str_PriorCalls
        ElseIf pos = "3" Or pos = "T" Or pos = "P" Then ' or bidder's partner passed '
            str_PriorCalls = "P-P-" & str_PriorCalls
        ElseIf pos = "4" Then
            str_PriorCalls = "P-P-P-" & str_PriorCalls
        End If
        Exit Sub
    End If

    ' More than one prior call. Normally, being a passed   '
    ' hand only applies to the first two rounds of bidding.'
    ' But then an entry in later rounds shouldn't have a   '
    ' Position specification, so we have to watch for that.'
    numDashes = 1
    Do While Mid$(str_PriorCalls, i, 1) = "-"
        i = InStr(i + 1, str_PriorCalls, "-")
        If i = 0 Then Exit Do
        numDashes = numDashes + 1
        DoEvents
    Loop
    bidder = numDashes + 2 ' 1 more prior bid than dashes + 1 for current bidder '
    If bidder > 4 Then
        bidder = bidder Mod 4
        If bidder = 0 Then bidder = 1
    End If

    If pos = "2" Then
        ' If bid requires bidder in position 2 '
        ' but his bidding position is 1 the we '
        ' add 1 pass to the prior calls. '
        If bidder = 1 Then
            str_PriorCalls = "P-" & str_PriorCalls
        End If

    ElseIf pos = "3" Or pos = "T" Or pos = "P" Then ' or bidder's partner passed '
        If bidder = 1 Then ' add 2 passes '
            str_PriorCalls = "P-P-" & str_PriorCalls
        ElseIf bidder = 2 Then ' add 1 pass '
            str_PriorCalls = "P-" & str_PriorCalls
        End If

    ElseIf pos = "4" Then
        If bidder = 1 Then ' add 3 passes '
            str_PriorCalls = "P-P-P-" & str_PriorCalls
        ElseIf bidder = 2 Then ' add 2 passes '
            str_PriorCalls = "P-P-" & str_PriorCalls
        ElseIf bidder = 3 Then ' add 1 pas '
            str_PriorCalls = "P-" & str_PriorCalls
        End If
    End If
End Sub


Private Sub Run_Save_Tutorial()
  L_Tutorial = "When you have entered or changed an entry in the Input " & _
               "Boxes, it is not stored in the database until you click " & _
               "'Edit - Save' or press Ctrl-S." & _
               vbCr & vbCr & _
               "If the 'New Entry' box is checked, the data in the Input " & _
               "Boxes are stored in a new entry; otherwise, assuming that " & _
               "you are changing an existing entry, the entry replaces " & _
               "the original entry, which itself is copied to the Garbage " & _
               "Grid from where it may be recovered with the 'Edit - Restore' " & _
               "command." & vbCr & vbCr & _
               "When an entry is saved, the 'Saved' box is checked so " & _
               "that you can see that the entry has been saved. If you " & _
               "change a field, as soon as you move from that field, the " & _
               "'Saved' box is unchecked."
    L_Tutorial.Height = 3800
    ResizeTutorial
End Sub


Private Sub mShowAllCols_Click()
    Dim g As Integer
    Dim i As Integer
    Dim c As Integer

    mShowAllCols.Checked = Not mShowAllCols.Checked
    For g = 0 To 1
        For i = 0 To cmb_HiddenCols(g).ListCount - 1
            c = Val(cmb_HiddenCols(g).List(i))
            Grid1(g).Splits(Grid1(g).Split).Columns(c).Visible = _
              mShowAllCols.Checked
        Next
        If Not bln_TwoGrids Then Exit For
    Next

    If mShowAllCols.Checked And _
       frm_Tutorial.Visible And _
       b_Continue.Visible = False _
    Then
        frm_Tutorial.Caption = "Show All Columns"
        Run_HidingColumns_Tutorial
    End If
End Sub


Private Sub mShowSort_Click()
    If frm_Tutorial.Visible And _
       b_Continue.Visible = False And _
       mFullTutorial.Checked = False _
    Then
        Run_ShowSort_Tutorial
        Exit Sub
    End If
    f_ShowSort.Visible = True
    b_Show.Default = True

    t_Show(0).SetFocus
    If frm_Tutorial.Visible And frm_Tutorial.Caption = "Show/Sort Grids" Then
        bln_Waiting = False
    End If
    f_ShowSort.ZOrder 0
End Sub


Private Sub Run_ShowSort_Tutorial()
    Dim i As Integer

    frm_Tutorial.Caption = "Show/Sort Grids"
    frm_Tutorial.Move f_ShowSort.Left + f_ShowSort.Width + 100, f_ShowSort.Top
    L_Tutorial = "The grid behind this tutorial box is now displaying every " & _
                 "entry in BidBase. (Scroll down if you wish.) Normally, you " & _
                 "only want to look at one section at a time. A 'section' is " & _
                 "the set of all bids for a particular set of 'Prior Bids'." & _
                 vbCr & vbCr & _
                 "For example, if you want to look at the responses " & _
                 "to opener's 1H bid when your Right Hand Opponent (RHO) " & _
                 "has doubled, you only need to see the section of the " & _
                 "database with 'Prior_Bids' of '1H-D'." & _
                 vbCr & vbCr & _
                 "To display this section, perform the following steps now: " & vbCr & _
                 "  1. Click the View menu, then click Show/Sort," & vbCr & _
                 "        or press the menu's shortcut key, Ctrl-W."
    frm_Tutorial.Width = 6400
    L_Tutorial.Width = 6000
    L_Tutorial.Height = 3300
    ResizeTutorial

    If mTwoGrids.Checked Then mTwoGrids_Click
    t_Show(0) = ""
    t_Show(1) = "z"
    cmb_Show(0).ListIndex = int_Prior_Bids_Col
    b_Show_Click
    Wait
    If frm_Tutorial.Visible = False Then
        Exit Sub
    ElseIf f_ShowSort.Visible = False Then
        f_ShowSort.Visible = True
        b_Show.Default = True
    End If

    For i = 0 To 3: t_Show(i) = "": Next
    cmb_Show(0).ListIndex = 0
    cmb_Show(1).ListIndex = 0
    L_Tutorial = "The first step, pressing Ctrl-W, brought up this box. " & vbCr & _
       "Now follow the remaining steps:" & vbCr & vbCr & _
       "  2. In the Top Grid's drop-down list, select Prior_Bids." & vbCr & vbCr & _
       "  3. In the Top Grid's FROM box, enter 1H-D." & vbCr & vbCr & _
       "  4. Press Tab. 1H-D will be copied to the TO box." & vbCr & vbCr & _
       "  5. Note the sort fields on the right." & vbCr & _
       "        They are set now where you usually want them." & vbCr & _
       "        See the documentation for more information." & vbCr & vbCr & _
       "  6. Now click Show or press Alt-W."

Tutor_ShowSort:
    b_TutorDetails.Visible = True
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#V8"
    L_Tutorial.Height = 3500
    L_Tutorial.Width = 5600
    frm_Tutorial.Width = 6000

    ResizeTutorial
    Wait
    If frm_Tutorial.Visible = False Then
        f_ShowSort.Visible = False
        b_Show.Default = False
        Exit Sub
    ElseIf t_Show(0) <> "1H-D" Then
        t_Show(0) = "1H-D"
        t_Show(1) = "1H-D"
        cmb_Show(0).ListIndex = int_Prior_Bids_Col
        cmb_Show(1).ListIndex = int_Prior_Bids_Col
        b_Show_Click
    End If
    frm_Tutorial.Caption = "Dupes Grid"
    If mDupesGrid(0).Checked Then mDupesGrid_Click 0
    If mDupesGrid(1).Checked Then mDupesGrid_Click 1
    If mDupesGrid(2).Checked Then mDupesGrid_Click 2
    If mTwoGrids.Checked Then mTwoGrids_Click
End Sub


Private Sub mSplit_Click(Index As Integer)
    Dim i As Integer

    If frm_Tutorial.Visible And _
       b_Continue.Visible = False _
    Then
        Run_SplitGrids_Tutorial
    End If

    If mSplit(Index).Checked Then
        Grid1(Index).Splits.Remove 1
        cmb_HiddenCols(2 + Index).Clear
        For i = 0 To cmb_HiddenCols(0 + Index).ListCount - 1
            cmb_HiddenCols(2 + Index).AddItem cmb_HiddenCols(0 + Index).List(i)
        Next
    Else
        Grid1(Index).Splits.Add 1
    End If
    mSplit(Index).Checked = Not mSplit(Index).Checked
End Sub


Private Sub Run_SplitGrids_Tutorial()
    frm_Tutorial.Caption = "Split Grids"
    L_Tutorial = "By splitting the grid(s) vertically, you can see different " & _
               "parts of the same entry. Normally, you do this to be able " & _
               "to see both the beginning and end of the entry at the same time. " & _
               "Note that you can hide different columns in each Split." & _
               vbCr & vbCr & _
               "Press F11 to split/unsplit the top grid and F12 for the " & _
               "bottom grid. You can try this now if you wish." & _
               vbCr & vbCr & _
               "The scroll bar moves both Splits of the same grid " & _
               "at the same time and a change made in one Split " & _
               "is, of course, made to the other side automatically."
    L_Tutorial.Height = 3000
    ResizeTutorial
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#V4"
End Sub


Private Sub mSysSelect_Click()
    str_Response = MyMsgBox("Not implemented yet.", 1, "ok", "Select System")
End Sub


Private Sub mTestHands_Click()
    ' Verify the specs for each entry in '
    ' the grid by testing its Test Hand. '
    'On Local Error GoTo ExitSub

    If frm_Tutorial.Visible Then
        If b_Continue.Visible = False Then
            Run_TestHands_Tutorial
        End If
        GoTo ExitSub
    End If

    If f_EnterHand.Visible Then
        GoTo ExitSub
    End If

    Dim b As String
    Dim e As String
    Dim ec As Integer
    Dim id1 As Long
    Dim id2 As Long
    Dim n As Integer
    Dim num As Long
    Dim sel As String
    Dim sel2 As String
    Dim nam As String
    Dim pos As String

    mTestHands.Checked = True
    bln_TestingSection = True
    p_Testing.Move (Me.Width - p_Testing.Width) / 2, Grid1(0).Height / 2 - p_Testing.Height
    p_Testing.Visible = True
    p_Testing.ZOrder 0
    p_Testing.Refresh
    bln_Abort = False
    
    If t_Show(2) = t_Prior Then
        int_GridNum = 1
    Else
        int_GridNum = 0
    End If

    Set NewBid = New BidBaseDB.bidder
    ' If "Class not registered", open BidBaseDB.dll with  '
    ' regsvr32.exe in xplorer2. For more information, see '
    ' the end of CardShark_BidBase_Programming_Notes.htm. '
    
     With NewBid
    '           '
      .DBPathAndName = str_Database

      Do
        DoEvents
        If Grid1(int_GridNum).Row = -1 Then ' no rows '
            GoTo ExitSub
        End If

        If Grid1(int_GridNum).Columns(int_LastBidCol).Caption <> "ID" Then
             str_Response = MyMsgBox("The last grid column is supposed to be ID" & vbCr & _
                                     "but instead is " & Grid1(int_GridNum).Columns(int_LastBidCol).Caption & vbCr & _
                                     "Try stopping and restarting the program.", 3, "ok", "")
            GoTo ExitSub
        End If

        If Grid1(int_GridNum).Columns(int_OrdNum_Col).Text = "ZZZZZZ" Then
            Grid1(int_GridNum).Row = Grid1(int_GridNum).Row + 1
        End If

        If Grid1(int_GridNum).Columns(0).Text = "0" Or _
           Grid1(int_GridNum).Columns(1).Text = "0" Or _
           Grid1(int_GridNum).Columns(2).Text = "0" _
        Then
            ' Entry is deactivated. '
            GoTo SkipToLoop
        End If

        If Grid1(int_GridNum).Columns(int_TestHand_Col).Text <> "" Then
            If InStr(" xX0", Grid1(int_GridNum).Columns(int_Pct_Col).Text) > 1 Or _
               InStr(" xX0", Grid1(int_GridNum).Columns(int_Pct_Col - 1).Text) > 1 Or _
               InStr(" xX0", Grid1(int_GridNum).Columns(int_Pct_Col - 2).Text) > 1 _
            Then
                ' Entry is skipped if Pct Used or Selected or '
                ' Sub is not blank. '
                GoTo SkipToLoop
            End If
            ' Entry does have a test hand - test it: '
            .PlayersHand = Grid1(int_GridNum).Columns(int_TestHand_Col).Text
            pos = Grid1(int_GridNum).Columns(int_Combo_Col + 3).Text
            If str_PriorCalls = "" Or pos = "" Then ' 2018-10-22: added pos="" because str_PriorCalls was never getting reset. '
                str_PriorCalls = Grid1(int_GridNum).Columns(int_Prior_Bids_Col).Text
            End If
            If pos <> "" Then
                AddPassesForPosition pos
            End If
            .PriorCalls = str_PriorCalls
            .Vul = Grid1(int_GridNum).Columns(int_Vul_Col).Text
            .BidToAnalyze = ""
            lng_ID = Val(Grid1(int_GridNum).Columns(int_LastBidCol).Text)
            .TestID = lng_ID
            .TestEntry = True
            .DBPathAndName = str_DBPathAndName

            b = .bid
            str_ErrMsg = .ErrMsg
            ec = .ErrCode
            id2 = .id

            If ec = 88 Then ' And b <> "" Then
                Stop 'This is already being shown. ("also matches...") '
                'ShowGetBidError b, lng_ID, id2
                Exit Sub
            End If
        End If

        If Grid1(int_GridNum).Columns(int_LastBidCol).Text = "" Then
            Exit Do
        End If

        If Grid1(int_GridNum).Row = Grid1(int_GridNum).VisibleRows - 1 Then
            Exit Do
        End If
SkipToLoop:
        Grid1(int_GridNum).Row = Grid1(int_GridNum).Row + 1
        DoEvents
      Loop While Grid1(int_GridNum).Columns(int_TestHand_Col).Text <> "" And _
             Not bln_Abort
    '        '
     End With
    '                                                                        '

    'Set NewBid = Nothing

    If e = "" Then
        Data1(int_GridNum).Refresh
    ElseIf f_ShowSort.Visible Then
        bln_Abort = True
    End If

    p_Testing.Visible = False
    Me.MousePointer = 0

    Dim i As Integer

    If int_GridNum = 0 Then
        i = 0
    Else
        i = 2
    End If

    If t_Show(i) <> t_Show(i + 1) Then
        If t_Show(i) = "" Then
            If t_Show(i + 1) <> "z" Then
                b = "through " & t_Show(i + 1) & vbCr
            Else
                b = ""
            End If
        Else
            b = t_Show(i) & vbCr & "through" & vbCr & t_Show(i + 1) & vbCr
        End If
    Else
        b = t_Show(i)
    End If

    If mTestHands.Checked And e = "" And Not bln_Abort Then
        'MsgBox "Testing is completed for" & vbCr & _
               "entries with Prior Bids of" & vbCr & _
               b & vbCr & vbCr & _
               "REMINDER:" & vbCr & _
               "Entries with '0' in the Sel, Sub" & vbCr & _
               "or Pct fields were not tested for" & vbCr & _
               "conflicts with earlier entries.", , _
               "CardShark BidBase"
    End If

ExitSub:
    mTestHands.Checked = False
    bln_TestingSection = False
    p_Testing.Visible = False
    MousePointer = 0
    On Local Error GoTo 0
End Sub


Private Sub FillCompareGrid(r As Integer, t0 As String, t1 As String, t2 As String)
    ' r = row number     '
    ' t0 = "sNum"        '
    ' t1 = value entry 1 '
    ' t2 = value entry 2 '
    ' Example: 13, "sNum", Grid1(0).Columns(int_xNum_Col + 2).Text, !sNum '
    g_CompareGrid.Row = r
    g_CompareGrid.Col = 0
    If r > 0 And r < 25 Then
        g_CompareGrid.CellBackColor = RGB(255, 255, 0)
        g_CompareGrid.CellFontBold = True
        g_CompareGrid.CellForeColor = vbBlack
    End If
    g_CompareGrid.Text = t0
    g_CompareGrid.Col = 1
    If r > 0 Then
        g_CompareGrid.CellFontBold = True
        g_CompareGrid.CellForeColor = vbRed
    End If
    g_CompareGrid.Text = t1
    g_CompareGrid.Col = 2
    If r > 0 Then g_CompareGrid.CellFontBold = True
    g_CompareGrid.Text = t2
End Sub


Public Sub LookUpBidName(num As Long, _
                         sel As String, _
                         sel2 As String, _
                         nam As String)

    Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
    Set rs_BidNamesTbl = db_BidBase.OpenRecordset(str_BidNamesTbl)
    With rs_BidNamesTbl
        If nam = "" Then
            .Index = "Number"
            .Seek "=", num
        Else
            .Index = "Bid_Name"
            .Seek "=", nam
        End If

        If .NoMatch = True Then
            Exit Sub
        End If

        num = !IDNum
        sel = Trim$("" & !Selected)
        sel2 = Trim$("" & !SubCat_Selected)
        nam = !Bid_Name
    End With
End Sub


Private Sub Run_TestHands_Tutorial()
    frm_Tutorial.Caption = "Test Hands"
    L_Tutorial = "Select this menu command to have the program verify the " & _
               "entries for one section using the Test Hands for each " & _
               "entry." & _
               vbCr & vbCr & _
               "If an entry fails, it is copied into the Input Boxes and " & _
               "the testing stops. You are told the reason it failed so " & _
               "that you can fix the problem and continue testing." & _
               vbCr & vbCr & _
               "A test may fail either because the Test Hand does not match " & _
               "the specifications of the entry it belongs to, or because " & _
               "it matches the specifications of an entry before the one " & _
               "it belongs to, so the bidding software would never " & _
               "get to the Test Hand's entry because it would stop " & _
               "at the first entry whose specs match the current hand."
    L_Tutorial.Height = 3300
    ResizeTutorial
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#E11"
End Sub


Private Sub mMenuTutorial_Click()
    If frm_Tutorial.Visible = False Then
        b_Continue.Visible = False
        frm_Tutorial.Move 7000, 2000
        frm_Tutorial.Caption = "CardShark BidBase Tutorial"
        L_Tutorial = "Click on a menu command for a tutorial on that command." & _
                 vbCr & vbCr & _
                 "If you click on a menu command when a menu tutorial " & _
                 "is being displayed, that menu command will be executed, " & _
                 "even if it's the menu command for which the tutorial " & _
                 "is currently being displayed." & _
                 vbCr & vbCr & _
                 "To run a tutorial on a different menu command, you " & _
                 "must first click the Continue button to complete the " & _
                 "current tutorial."
        frm_Tutorial.Visible = True
        b_Continue.Visible = True
        b_Continue.Top = b_QuitTutorial.Top
        b_TutorDetails.Top = b_QuitTutorial.Top

        mMenuTutorial.Checked = True
    Else
        mMenuTutorial.Checked = False
    End If
End Sub


Private Sub mTwoGrids_Click()
    ' Display two grids when checked   '
    ' or hide 2nd grid when unchecked. '

    If frm_Tutorial.Visible And _
       b_Continue.Visible = False _
    Then
        frm_Tutorial.Caption = "Two Grids"
        Run_TwoGrids_Tutorial
        Exit Sub
    End If

    bln_TwoGrids = Not bln_TwoGrids
    mTwoGrids.Checked = bln_TwoGrids
    Grid1(1).Visible = bln_TwoGrids
    VScroll1.Visible = bln_TwoGrids
    bln_SkipIt = False
    Form_Resize

    If bln_TwoGrids Then
        Grid1(2).Visible = False
        pSplitterBar.Visible = True
        Dim cw As Column
        Dim i As Integer
        Dim p0 As String
        Dim p1 As String

        i = 0
        For Each cw In Grid1(1).Splits(0).Columns
            cw.Width = sng_ColWidths(i)
            i = i + 1
        Next

        If Grid1(1).VisibleRows < 1 Then
            str_Response = MyMsgBox("No entries for Prior Bids of" & vbCr & _
                                    t_Show(2), 2, "ok", "")
            Exit Sub
        End If

      'If Grid1(1).VisibleRows > 1 Then
      '  p0 = Grid1(1).Columns(int_Prior_Bids_Col)
      '  If Grid1(1).Columns(int_Bid_Col) <> "" Then
      '      If Grid1(1).Row < Grid1(1).VisibleRows - 2 Then
      '          Grid1(1).Row = 1
      '          p1 = Grid1(1).Columns(int_Prior_Bids_Col)
      '      End If
      '  End If
      ' End If
        If (p0 <> t_Show(2).Text Or p0 <> p1) And _
           mDupesGrid(0).Checked = False _
        Then
            If t_Show(1) = "" Or t_Show(0) = t_Show(2) Then
                mShowSort_Click
            Else ' Getting shown multiple times. '
                'bln_ShowGrid2 = True
                'b_Show_Click
            End If
        End If
    Else
        pSplitterBar.Visible = False
    End If
End Sub


Private Sub Run_TwoGrids_Tutorial()
    frm_Tutorial.Caption = "Two Grids"
    L_Tutorial = "Click on 'View' and 'Two Grids' to display two " & _
                 "grids on the screen. This lets you view different " & _
                 "sections of the database at the same time." & _
                 vbCr & vbCr & _
                 "Two grids are handy when comparing different parts of " & _
                 "the database, like Opening Bids vs Responses To " & _
                 "Opening Bids, or, comparing responses " & _
                 "to 1H-P and 1S-P, which should be very similar." & _
                 vbCr & vbCr & _
                 "To do the latter, we have to bring the Show/Sort box " & _
                 "back up so that we can specify what to show in the " & _
                 "bottom grid. " & _
                 vbCr & vbCr & _
                 "Press Ctrl-W to bring up the Show/Sort box."
    L_Tutorial.Height = 3500
    ResizeTutorial
    frm_Tutorial.Left = f_ShowSort.Left + f_ShowSort.Width + 100
    frm_Tutorial.Top = f_ShowSort.Top
    bln_Waiting = False
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#V6"
    Wait
    If frm_Tutorial.Visible = False Then Exit Sub

     If mTwoGrids.Checked = False Then
        b_Continue.Visible = True
        mTwoGrids_Click
    End If

    If f_ShowSort.Visible = False Then
        mShowSort_Click
    End If

    L_Tutorial = "In the FROM box for the Bottom Grid, enter 1S-P and " & _
                 "press Tab. " & _
                 vbCr & vbCr & _
                 "Tab again to get to the IN drop-down list " & _
                 "and select Prior_Bids." & _
                 vbCr & vbCr & _
                 "Click the Show button, or press Alt-W."

    b_Continue.Visible = True
    t_Show(0) = "1H-P"
    t_Show(1) = "1H-P"
    't_Show(2).SetFocus
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#V8"
    Wait
    If frm_Tutorial.Visible = False Then Exit Sub

    If t_Show(2) <> "1S-P" Then
        t_Show(2) = "1S-P"
        t_Show(3) = "1S-P"
        cmb_Show(1) = cmb_Show(0)
        b_Show_Click
    End If

    L_Tutorial = "On the right of the grids, each grid has its " & _
               "own scroll bar, and to the right of them is another " & _
               "scroll bar which scrolls both grids at once to make it " & _
               "easier to compare the entries. Try the scroll bars now." & _
               vbCr & vbCr & _
               "Notice that an entry from one section matches pretty " & _
               "closely the entry from the other section which has the " & _
               "same Bid_Order_Number. As the name implies, this number " & _
               "determines the order in which bid entries are evaluated " & _
               "by the program when making a bid, which is important " & _
               "because the program stops trying entries when the first " & _
               "entry is found which matches the current hand's specs. "
    L_Tutorial.Height = 3200
    ResizeTutorial
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#V3"
End Sub


Private Sub mUnhide_Click()
    If frm_Tutorial.Visible Then
        If b_Continue.Visible = False Then
            frm_Tutorial.Caption = "Unhide Columns"
            Run_HidingColumns_Tutorial
        End If
    End If
    f_Unhide.ZOrder 0
    f_Unhide.Visible = True
End Sub


Private Sub mPreviewHandSpecs_Click()
    If frm_Tutorial.Visible Then
        If b_Continue.Visible = False Then
            Run_PreviewHandSpecs_Tutorial
        End If
        Exit Sub
    End If

    mPreviewHandSpecs.Checked = True

    If f_EnterHand.Visible Then
        DisplayFindHand
    End If
End Sub


Private Sub Run_PreviewHandSpecs_Tutorial()
    frm_Tutorial.Caption = "Preview Hand Specs"
    f_EnterHand.Visible = True
    L_Tutorial = "Enter a hand and prior bids in this window and press Ctrl-P " & _
                 "or click 'Edit' - 'Preview Hand Specs' and BidBase will display " & _
                 "its analysis of the hand in the Input Boxes." & _
                 vbCr & vbCr & _
                 "If you click the 'Find' button and no matching entry " & _
                 "is found, BidBase will display the hand's specs in the " & _
                 "Input Boxes to make it easier to create a new entry. " & _
                 vbCr & vbCr & _
                 "The specs entered are very specific for the hand given. " & _
                 "You normally should modify them to make the specs less " & _
                 "restrictive; otherwise, the entry will not work for " & _
                 "similar, but slightly different, hands."
    L_Tutorial.Height = 3200
    If f_EnterHand.Left < 3000 Then
        frm_Tutorial.Top = f_EnterHand.Top
        frm_Tutorial.Left = f_EnterHand.Left + f_EnterHand.Width + 50
    End If
    ResizeTutorial
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#V2"
End Sub


Private Sub mVersion_Click()
    If str_MsgShown = "Version" Then
        str_MsgShown = ""
    Else
        str_MsgShown = "Version"
        str_Response = MyMsgBox("CardShark BidBase" & vbCr & _
                                "Database Editor" & vbCr & _
                                "Ver. " & Version$, 3, "ok", "")
    End If
End Sub


Private Sub mViewAll_Click()
    If frm_Tutorial.Visible Then
        If b_Continue.Visible = False Then
            Run_OverviewAll_Tutorial
            Exit Sub
        End If
    End If

    If mFourGrids.Checked Then
        mFourGrids_Click
    End If
    If mTwoGrids.Checked Then
        bln_TwoGrids = True
        mTwoGrids_Click
    End If
    t_Show(0) = ""
    t_Show(1) = "z"
    cmb_Show(0).ListIndex = int_Prior_Bids_Col
    b_Show_Click
End Sub


Private Sub Run_OverviewAll_Tutorial()
    frm_Tutorial.Caption = "Overview of All Entries"
    L_Tutorial = "This menu command causes all the entries in the database " & _
                 "to be loaded into the grid for viewing without having " & _
                 "to use the Show/Sort window." & _
                 vbCr & vbCr & _
                 "Click 'View - Overview all' or press Ctrl-O to try it now. "
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#V02"
End Sub


Private Sub mViewFromHereDown_Click(Index As Integer)
    ' Index = 0 = View from current entry to last entry for same Prior Bids.'
    ' Index = 1 = View from current entry to last entry specified for grid. '
    ' This is done by selecting records for the grid based on Bid_Order_Num '
    ' starting with the current number, then f

    If frm_Tutorial.Visible And _
       b_Continue.Visible = False _
    Then
        Run_HereDown_Tutorial
        Exit Sub
    End If

    Dim n As Integer
    Dim g As Integer
    Dim pb1 As String
    Dim pb2 As String
    Dim pbid1 As String
    Dim pbid2 As String
    Dim pbid3 As String
    Dim pbid4 As String
    Dim BON As String
    Dim w As String
    Dim pbc As Integer
    Dim bonc As Integer
    Dim r As Integer
    Dim p As Integer
    Dim s1 As String

    g = int_GridNum
    pbc = int_Prior_Bids_Col
    bonc = int_OrdNum_Col
    pb1 = Trim$(Grid1(g).Columns(pbc).Text)
    BON = Trim$(Grid1(g).Columns(bonc).Text)
    p = Val(Trim$(Grid1(g).Columns(int_Player_Col).Text))
    r = Val(Trim$(Grid1(g).Columns(int_Round_Col).Text))

    If g = 0 Then
        ' t_Show(0) and (1) are for Grid 1 '
        ' and (2) and (3) are for Grid 3.  '
        n = 0
    Else
        n = 2
    End If

    If Index = 0 Or t_Show(n) = t_Show(n + 1) Then

        w = "Bid_Order_Number BETWEEN " & _
            "'" & BON & "' AND 'ZZZZZZ' AND " & _
            "Prior_Bids = '" & pb1 & "' "

    Else
        ' Find next set of Prior_Bids, if any: '
        If t_Show(n) <> t_Show(n + 1) And _
           t_Show(n + 1) <> "" _
        Then
            Do
                Grid1(g).Row = Grid1(g).Row + 1
                DoEvents
            Loop While pb1 = Grid1(g).Columns(pbc).Text And _
                       Grid1(g).Columns(bonc).Text <> ""

            pb2 = Grid1(g).Columns(pbc).Text

            ' When Prior_Bids is the field which the  '
            ' range of entries is from, convert it to '
            ' the sort order code for correct sorting.'
            ' (See the doc file for more information.)'
            pbid1 = PB_To_SortOrd_Code(pb1)
            pbid2 = PB_To_SortOrd_Code(pb2)
            ' pbid2 is the start of the next set of   '
            ' Prior Bids after the one on which the   '
            ' user clicked, assuming more than 1 set. '

            ' pb3 is the end of the specified set of  '
            ' Prior Bids. For example: 1C-1H may be   '
            ' the start and 1C-1S may be the end,     '
            ' though this would be unusual. The user  '
            ' will rarely want more than one set of   '
            ' prior bids in the same grid, and when he'
            ' does, it will probably be to the end of '
            ' the database, in which case the second  '
            ' spec will be "z".                       '
            If Left$(t_Show(n + 1), 1) = "z" Then
                pbid3 = String$(35, "z")
            Else
                pbid3 = PB_To_SortOrd_Code(t_Show(n + 1))
            End If

            If t_Show(n) <> Grid1(g).Columns(pbc).Text Then
                w = "(Round = '" & r & "' " & _
                     "AND Player = '" & p & "' " & _
                     "AND PB_Sort_Order = '" & pbid1 & "' " & _
                     "AND Bid_Order_Number BETWEEN '" & _
                            BON & "' AND 'ZZZZZZ') " & _
                     "OR (PB_Sort_Order BETWEEN '" & _
                            pbid2 & "' AND '" & pbid3 & "')"

            Else
                w = "(Bid_Order_Number BETWEEN '" & _
                            BON & "' AND 'ZZZZZZ' " & _
                    "AND Prior_Bids = '" & pb1 & "') "
            End If
        End If
    End If

    s1 = Sort_Order_Specs

    Data1(g).RecordSource = _
             qSel & _
             " WHERE " & w & _
             " ORDER BY " & s1 & ";"
    Data1(g).Refresh

End Sub


Private Sub Run_HereDown_Tutorial()
    frm_Tutorial.Caption = "View From Here Down"
    L_Tutorial = "When you are working on a section of entries and there " & _
               "are more of them than will fit on the screen at one time, " & _
               "you can click on a line and select one of the 'View - From " & _
               "Here Down' command options to display the entries in the " & _
               "section starting with the line on which you clicked." & _
               vbCr & vbCr & _
               "Presumably when you do this, you are working down the list " & _
               "and have finished with the entries above the line on which " & _
               "you clicked, so this gets those entries out of your way." & _
               vbCr & vbCr & _
               "To try this now, click on a row about half way down and " & _
               "click 'View'  -  'From here down'  -  'To end of Prior Bids set'  " & _
               "(or press Ctrl-F).   To see the full section again, press Ctrl-W to bring up " & _
               "the Show/Sort window and click 'Show'."
    L_Tutorial.Height = 3500
    b_Continue.Visible = True
    ResizeTutorial
    str_TutorDetails = "CardShark_BidBase_Editor_Menus.htm#V0"
    b_TutorDetails.Top = b_Continue.Top
End Sub


Private Sub mViewSAYC_Click()
    str_NoteFile(0) = "SAYC.htm"
    ShowDocs
Exit Sub
    
    str_NoteFile(1) = "" ' "CardShark_BidBase_Editor_Menus.htm"
    str_PgmPath(0) = str_BBpath & "\Notes\"
    str_PgmPath(1) = "" ' str_BBpath & "\Notes\"
    Dim xzx As Long
    Dim b As String
    If str_Browser = "" Then
        str_Browser = "BidBase File Viewer.exe"
    End If
    xzx = Shell(str_Browser & " " & _
                str_PgmPath(0) & str_NoteFile(0) & ", " & _
                str_PgmPath(1) & str_NoteFile(1), _
                vbNormalFocus)
End Sub


Private Sub mGlossary_Click()
    Dim xzx As Long
    Dim b As String
    
    str_NoteFile(0) = "CardShark_BidBase_Abbreviations.htm"
    str_PgmPath(0) = str_BBpath & "\Notes\"
    
    If str_Browser = "" Then
        b = "BidBase File Viewer.exe"
        xzx = Shell(str_Browser & _
                    str_PgmPath(0) & str_NoteFile(0), _
                    vbNormalFocus)
    Else
        b = str_Browser
        xzx = Shell(str_Browser & _
                    str_PgmPath(0) & str_NoteFile(0), _
                    vbNormalFocus)
    End If
End Sub


Private Sub o_CompareOptions_Click(Index As Integer)
    Select Case Index
        Case 2, 3, 4
            b_CompareDataOK_Click
    End Select
End Sub


Private Sub pSplitterBar_DragDrop(source As Control, x As Single, Y As Single)
    With source
        .Drag vbEndDrag
        .Move pSplitterBar.Left + x - sng_DragoffsetX, pSplitterBar.Top + Y - sng_DragoffsetY
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub pSplitterBar_MouseDown(Button As Integer, _
                                   Shift As Integer, _
                                   x As Single, Y As Single)
    ' Remember starting location '
    ' for pSplitterBar_MouseMove.'
    If Button = 1 Then
        pSplitterBar.Tag = Y
    End If
End Sub


Private Sub pSplitterBar_MouseMove( _
                Button As Integer, _
                Shift As Integer, _
                x As Single, _
                Y As Single)

    Dim i As Integer
    Dim chng As Long
    Dim totalHeight As Long

    totalHeight = Grid1(0).Height
    If Grid1(1).Visible Then
        totalHeight = totalHeight + Grid1(1).Height
    ElseIf Grid1(2).Visible Then
        totalHeight = totalHeight + Grid1(2).Height
    Else
        ' Don't have two grids to adjust between. '
        Exit Sub
    End If

    If Button = 1 Then
        chng = Y - Val(pSplitterBar.Tag)

        ' Adjust height of grids, but only if the '
        ' height of the top grid is greater than  '
        ' 20% of the total height of the top two  '
        ' grids if the splitter is moving down or '
        ' less than 80% if the splitter goes up.  '
        If (Grid1(0).Height > totalHeight * 0.2 Or _
            chng > 0) And _
           (Grid1(0).Height < totalHeight * 0.8 Or _
            chng < 0) _
        Then
            Grid1(0).Height = Grid1(0).Height + chng
            For i = 1 To 2
                If Grid1(i).Visible Then
                    ' Grid1(1) is the second data grid.'
                    ' Grid1(2) is the garbage grid.    '
                    ' They will not both be showing at '
                    ' the same time.   If the splitter '
                    ' bar moves when all 4 grids are   '
                    ' showing, Grid1(0) gets smaller or'
                    ' larger and Grid1(1) gets larger  '
                    ' or smaller, respectively. The 3rd'
                    ' and 4th grids don't change/move. '
                    Grid1(i).Top = Grid1(i).Top + chng
                    Grid1(i).Height = Grid1(i).Height - chng
                End If
            Next
            pSplitterBar.Top = pSplitterBar.Top + chng
        End If
    End If
End Sub


Private Sub pTutorDragBar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lng_Top = pTutorDragBar.Top
    'lng_Left = pTutorDragBar.Left ' Not needed when only resizing length. '
    'lng_X = X
    lng_Y = Y
    bln_Dragging = True
End Sub


Private Sub pTutorDragBar_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If bln_Dragging Then
        bln_Dragging = False
    End If
End Sub


Private Sub pTutorDragBar_MouseMove( _
        Button As Integer, _
        Shift As Integer, _
        x As Single, _
        Y As Single)
        
    Static Moving As Boolean

    If bln_Dragging Then
        Moving = Not Moving
        If Moving Then
            ' Show changes as it is being resized: '
            pTutorDragBar.Top = lng_Top + Y - lng_Y
            pTutorDragBar.Left = lng_Left + x - lng_X
            L_Tutorial.Height = L_Tutorial.Height + Y - lng_Y
            ResizeTutorial
        Else
            lng_Top = pTutorDragBar.Top
            lng_Left = pTutorDragBar.Left
        End If
    End If
End Sub


Private Sub t_AsksFor_GotFocus()
    str_OrigData = t_AsksFor.Text
    L_Status.Clear
    L_Status.AddItem " H(elp)x(x=C,D,H,S),  C(ontrol)x(x=C,D,H,S), " & _
                     " B(id)x (ex: B2C, Bm=bid a minor),  P(ick a" & _
                     " suit/slam),  D(ecide contract), U(nbalanced)"
                     
    L_Status.AddItem " A(ces),  K(ings),  K(ey)C(ards),  1(1-round" & _
                     " force),  G(ame force),  PC=Pass Or" & _
                     " Correct/Continue,  SO=Sign-Off"
End Sub


Private Sub t_AsksFor_Validate(Cancel As Boolean)
    L_Status.Clear
    If str_OrigData <> t_AsksFor Then
        ck_Saved.Value = 0
    End If
End Sub


Private Sub t_BestContract_GotFocus()
    L_Status.Clear
    L_Status.AddItem " Best contract per DDA, normally after 2+ rounds of bidding."
    str_OrigData = t_BestContract
    t_BestContract.SelStart = 0
    t_BestContract.SelLength = 32768
End Sub


Private Sub t_BestContract_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub


Private Sub t_BestContract_Validate(Cancel As Boolean)
    L_Status.Clear
    If str_OrigData <> t_BestContract Then
        ck_Saved.Value = 0
    End If
End Sub


Private Sub t_CustPoints_GotFocus()
    L_Status.Clear
    L_Status.AddItem " Total points based on honor values" & _
                     " assigned below the Custom Points box."
    str_OrigData = t_CustPoints
    t_CustPoints.SelStart = 0
    t_CustPoints.SelLength = 32768
End Sub


Private Sub t_CustPoints_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
        Case 0
            Select Case KeyCode
                Case 38
                    ' Up cursor
                    t_Distrib_Pts.SetFocus
                Case 40
                    ' Down cursor
                    t_CustPts(3).SetFocus
            End Select
    End Select
End Sub


Private Sub t_CustPoints_Validate(Cancel As Boolean)
    Dim n As Integer

    L_Status.Clear
    If str_OrigData <> t_CustPoints Then
        ck_Saved.Value = 0
        n = Len(t_CustPoints)
        If n > 5 Then
            str_Response = MyMsgBox(" Custom Points field can only be" & _
                                    "5 digits.", 2, "ok", "")
            Cancel = True
        Else
            If str_OrigData <> t_CustPoints Then
                ck_Saved.Value = 0
            End If
        End If
    End If
End Sub


Private Sub t_CustPts_GotFocus(Index As Integer)
    L_Status.Clear
    L_Status.AddItem " Enter points to assign to each honor" & _
                     " card for 'Cust.Points', above."
    str_OrigData = t_CustPts(Index)
    t_CustPts(Index).SelStart = 0
    t_CustPts(Index).SelLength = 32768
End Sub


Private Sub t_CustPts_KeyDown( _
                Index As Integer, _
                KeyCode As Integer, _
                Shift As Integer)

    Select Case Shift
        Case 0
            Select Case KeyCode
                Case 38 'up
                    t_CustPoints.SetFocus
                Case 40 'down
                    t_Winners.SetFocus
                Case 37 'left
                    If Index > 0 And _
                       t_CustPts(Index).SelStart = 0 And _
                       t_CustPts(Index).SelLength = 0 _
                    Then
                        t_CustPts(Index - 1).SetFocus
                    End If
                Case 39 'right
                    If Index < 4 And _
                       t_CustPts(Index).SelStart = Len(t_CustPts(Index)) _
                    Then
                        t_CustPts(Index + 1).SetFocus
                    End If
            End Select
    End Select
End Sub


Private Sub t_CustPts_Validate(Index As Integer, Cancel As Boolean)
    Dim n As Integer

    If str_OrigData <> t_CustPts(Index) Then
        n = Len(t_CustPts(Index))
        If n > 3 Then
            str_Response = MyMsgBox("Custom Points field can only be" & _
                                    "3 digits.", 2, "ok", "")
            Cancel = True
        Else
            ck_Saved.Value = 0

            If ck_AutoFill.Value <> 0 And _
               t_CustPts(Index) = "1" _
            Then
                t_TestHand = ""
                bln_Working = True
                CreateTestHand
                Do While bln_Working
                    DoEvents
                Loop
            End If
        End If
    End If
End Sub


Private Sub t_Disclosure_GotFocus()
    str_OrigData = t_Disclosure.Text
    L_Status.Clear
    L_Status.AddItem " Text of disclosure, or entry notes. " & _
                     " To reference a file, enter '[See" & _
                     " file -->' and a file name."
    
    L_Status.AddItem " To create a note file for the entry," & _
                     " save the entry then click 'N' to the" & _
                     " right to create a note file."
    t_Disclosure.SelStart = 0
    t_Disclosure.SelLength = 32768
End Sub


Private Sub t_Disclosure_Validate(Cancel As Boolean)
    L_Status.Clear

    If str_OrigData <> t_Disclosure Then
        ck_Saved.Value = 0
        LookForNoteFile t_Disclosure, _
                        t_Prior, _
                        cmb_Bid, _
                        "" & lng_ID, _
                        False, _
                        True, _
                        False
    End If
End Sub


Private Sub NoHighlight(tb As Control)
    tb.SelStart = 32768
    tb.SelLength = 0
End Sub


Private Sub t_Distrib_Pts_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
        Case 0
            Select Case KeyCode
                Case 38 ' up arrow
                    t_Hand_HCPs.SetFocus '
                Case 40 ' down arrow
                    t_CustPoints.SetFocus
            End Select
    End Select
End Sub


Private Sub t_Distrib_Pts_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub


Private Sub t_DupeEntry_Validate(Index As Integer, Cancel As Boolean)
    If t_Prior <> "" Then
        b_DupeEntry(Index).Visible = (t_DupeEntry(Index) <> "")
    ElseIf Trim$(t_DupeEntry(Index)) <> "" Then
        'MsgBox "Load an entry into the Input Boxes first."
    End If
End Sub


Private Sub t_Find_Ref_Click()
    cmb_FindRef.Visible = False
End Sub


Private Sub t_Find_Ref_DblClick()
    If Trim$(t_Find_Ref) = "" Then Exit Sub
    If cmb_FindRef.Visible Then
        cmb_FindRef.AddItem Trim$(t_Find_Ref)
    ElseIf Trim$(t_Ref) = "" Then
        t_Ref = Trim$(t_Find_Ref)
        t_Ref = t_Find_Ref
    End If
End Sub


Private Sub t_Find_Ref_GotFocus()
    If t_Find_Ref = "" Then Exit Sub
    t_Find_Ref.SelStart = 0
    t_Find_Ref.SelLength = Len(t_Find_Ref)
End Sub


Private Sub t_FindAnalysis4Bid_DblClick()
    t_FindAnalysis4Bid = ""
End Sub


Private Sub t_FindAnalysis4Bid_GotFocus()
    t_FindAnalysis4Bid.SelStart = 0
    t_FindAnalysis4Bid.SelLength = 32768
End Sub


Private Sub t_FindAnalysis4Bid_KeyPress(KeyAscii As Integer)
    ' A bid entered here is passed to the BidBase Bidder DLL '
    ' which displays an analysis of why each entry with the  '
    ' specified bid fails to match the hand's specs.         '
    Dim k As String

    If KeyAscii = 13 Then 'Enter
        b_FindBidForHand_Click
        Exit Sub
    End If

    k = UCase$(Chr$(KeyAscii))
    KeyAscii = Asc(k)

    If k = "@" Then
        t_FindAnalysis4Bid.Left = 2000
        t_FindAnalysis4Bid.Width = 900
    ElseIf InStr(" 01234567CDHSNPDR AL", k) = 0 And KeyAscii <> 8 Then
        ' "Al" is for entering "ALL" '
        KeyAscii = 0
    End If
End Sub


Private Sub t_FindHand_DragDrop(Index As Integer, _
                                source As Control, _
                                x As Single, Y As Single)
    With source
        .Drag vbEndDrag
        .Move .Left + t_FindHand(int_FindHandNum).Left + x - sng_DragoffsetX, _
              .Top + t_FindHand(int_FindHandNum).Top + Y - sng_DragoffsetY
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub t_FindHand_GotFocus(Index As Integer)
    Dim i As Integer
    Dim n As Integer

    int_FindHandNum = Index
    If Index = 0 Then
        t_FindHand(0).ForeColor = vbBlack
        t_FindHand(1).ForeColor = RGB(150, 150, 150)
        b_DelFindHand(1).Visible = False
    Else
        t_FindHand(1).ForeColor = vbBlack
        t_FindHand(0).ForeColor = RGB(150, 150, 150)
        b_DelFindHand(0).Visible = False
    End If

    b_DelFindHand(Index).Visible = False
    
    If bln_NoHiLite Then
        bln_NoHiLite = False
    Else
        t_FindHand(Index).SelStart = 0
        t_FindHand(Index).SelLength = 16
    End If

    If bln_SkipIt Then Exit Sub

    b_FindBidForHand.Default = True
    str_OrigData = t_FindHand(Index).Text
    t_FindHand(Index).Text = Trim$(t_FindHand(Index).Text)

    If cmb_FindBid.Visible Then
        cmb_FindBid.Visible = False
        t_FindAnalysis4Bid.Visible = True
    End If
End Sub


Private Sub t_FindHand_KeyPress(Index As Integer, KeyAscii As Integer)
    ' This routine controls the input of card hands.  '
    ' It converts the space bar to dashes, uppercases '
    ' card letters (A-J), and validates the order of  '
    ' cards entered.                                  '

    Dim k As String
    Dim i As Integer
    Dim j As Integer
    Dim ss As Integer
    Dim suitPrompt As String
    Static bidMade As Boolean

    k = Chr$(KeyAscii)

    If k <> "" Then
        ' If hand changes, assume bid found will change. '
        L_BidFound = ""
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

    With t_FindHand(Index)
    ''''''''''''''''''''''
    If InStr("AKQJT", k) = 0 And _
       (k < "2" Or k > "9") And _
       KeyAscii <> 3 And _
       KeyAscii <> 8 And KeyAscii <> 48 And _
       KeyAscii <> 22 And _
       KeyAscii <> 24 And KeyAscii <> 13 And _
       k <> "-" _
    Then
        ' Exit when valid key not pressed. '
        ss = .SelStart
        Beep
        str_Response = MyMsgBox(k & " (ASCII-" & Asc(k) & _
                ") is not a valid card." & vbCr & _
                "Must be one of these:" & vbCr & _
                "AKQJT98765432.", 3, "ok", "")
        .SelStart = ss + 1
        .SelLength = 0
        bln_NoHiLite = True
        .SetFocus
        KeyAscii = 0
        Exit Sub
    End If

    If k = "-" Then
        ' Hands are displayed in S-H-C-D order to alternate Red-Black, '
        ' but need to be entered in S-H-D-C for database searches.     '
        ' I frequently forgot to change C-D to D-C when copying hands  '
        ' from the Practice program, so this prompt was added. '
        i = InStr(.Text, "-")
        j = 0
        Do While i > 0
            j = j + 1
            i = InStr(i + 1, .Text, "-")
            DoEvents
        Loop

        i = Len(.Text)
        If .SelStart < i Then
            .Text = Left$(.Text, .SelStart) & _
                         k & _
                         Mid$(.Text, .SelStart + 1)
            KeyAscii = 0
            Exit Sub
        End If

        .Text = .Text & "-" & suitPrompt
        .SelStart = i + 1
        .SelLength = 13
        KeyAscii = 0
        Exit Sub
    End If

    ' Check for card in correct order '
    ' compared to prior card, if any: '
    i = InStr("AKQJT98765432", k)
    If .SelStart > 0 And _
       i > 0 _
    Then
        j = InStr("AKQJT98765432", Mid$(.Text, .SelStart, 1))
        If j >= i Then 'card entered is out of rank order
            ss = .SelStart
            .Text = Left$(.Text, .SelStart) & _
                    k & _
                    Mid$(.Text, .SelStart + 1)
            Beep
            str_Response = MyMsgBox("Card spec is out of order.", _
                                    1, "ok", "")
            .SelStart = ss + 1
            .SelLength = 0
            bln_NoHiLite = True
            .SetFocus
            KeyAscii = 0
            Exit Sub
        End If
    End If

    If .SelStart = 15 And _
       Chr$(KeyAscii) > "0" _
    Then
        t_FindPrior.SetFocus
    End If
    ''''''''
    End With
End Sub


Private Sub t_FindHand_LostFocus(Index As Integer)
    Dim i As Long
    
    If Left$(t_FindHand(Index).Text, 3) = "oid" Then
        t_FindHand(Index).Text = Mid$(t_FindHand(Index).Text, 4)
    End If
    
    i = Len(t_FindHand(Index).Text)
    
    If i = 0 And Index = 1 Then
        t_FindHand(0).ForeColor = vbBlack
        Exit Sub
    End If

    If i < 16 Then Exit Sub
    
    t_FindHand(Index).Text = Replace(t_FindHand(Index).Text, "Void", "")
    
    t_FindHand(Index).Text = Trim$(Left$(t_FindHand(Index).Text, 16))

    t_FindHand(Index).Text = Replace(t_FindHand(Index).Text, ".", "-")

    If cmb_FindBid.ListCount >= 0 Then
        For i = 0 To cmb_FindBid.ListCount - 1
            If t_FindHand(Index).Text = cmb_FindBid.List(i) Then
                GoTo Exit_t_FindHand
            End If
        Next
    End If

    cmb_FindBid.AddItem t_FindHand(Index).Text

Exit_t_FindHand:
    If Index = 0 And Right$(t_Find_Ref, 2) = "BB" And t_FindHand(1) = "" Then
        ' Enter 2nd hand from the Bidding Box article: '
        t_FindHand(1).SetFocus
    End If
End Sub


Private Sub t_FindHand_Validate(Index As Integer, Cancel As Boolean)
    Dim t As String
    Dim i As Integer

    With t_FindHand(Index)
    ''''''''''''''''''''''
    If .Visible = False Then Exit Sub
    If InStr(.Text, " ") > 0 Or InStr(.Text, "C") > 0 Or InStr(.Text, "%") > 0 Then
        ' From Bul online: '
        .Text = Replace(.Text, "V", "-")
        .Text = Replace(.Text, "+", "-")
        .Text = Replace(.Text, "%", "-")
        .Text = Replace(.Text, Chr$(10), "")
        ' From Bridge Quiz: '
        .Text = Replace(.Text, "oid", "")
        .Text = Replace(.Text, "–", "-")
        .Text = Replace(.Text, "--", "-")
        .Text = Replace(.Text, ":", "-")
        .Text = Replace(.Text, "S-", "")
        .Text = Replace(.Text, "S", "")
        .Text = Replace(.Text, "H", "")
        .Text = Replace(.Text, "D", "")
        .Text = Replace(.Text, "C", "")
        .Text = Replace(.Text, "10", "T")
        .Text = Replace(.Text, "1 0", "T")
        .Text = Replace(.Text, "?", "-")
        .Text = Replace(.Text, " ", "")
        .SelStart = 0
        .SelLength = 16
    End If

    If .Text <> str_OrigData Then
        L_BidFound = ""
    End If

    .SetFocus
    ''''''''
    End With
End Sub


Private Sub t_FindPrior_Change()
    ' This makes it difficult to edit prior bid. '
    Exit Sub
    
    If t_FindPrior <> "" And Not bln_SkipIt Then
        ConvertPrior t_Prior
    End If
    bln_SkipIt = False
End Sub


Private Sub t_FindPrior_GotFocus()
    If Not bln_NoHiLitePrior Then ' Coming from ck_Unopposed. '
        t_FindPrior.SelStart = 0
        t_FindPrior.SelLength = 99
    Else
        bln_NoHiLitePrior = False
    End If
End Sub


Private Sub t_FindPrior_KeyPress(KeyAscii As Integer)
    Dim k As String
    Dim n As Integer
    Dim i As Integer
    Dim sStart As Integer
    Dim x As Boolean

    If KeyAscii = 3 Then Exit Sub

    If KeyAscii = 13 Then 'Enter
        b_FindBidForHand_Click
        KeyAscii = 0
        Exit Sub
    End If

    k = Chr$(KeyAscii)
    n = Val(k)
    sStart = t_FindPrior.SelStart
    ' SelStart is the character you clicked after. '
    ' e.g.: ABCD - If you clicked after C, it's 3  '
    ' even though you are inserting at position 4. '
    If sStart > 0 Then
        ' Part of a test to make sure 2 digits in a row were not pressed. '
        
        x = Val(Mid$(" " & t_FindPrior, t_FindPrior.SelStart, 1)) > 0
    End If

    If n > 7 Or (n > 0 And n < 8 And x) Then
        Beep
        str_Response = MyMsgBox("A number to bid must be 1-7.", 1, "ok", "")
        KeyAscii = 0
        Exit Sub

    ElseIf k = " " Then
        ' Change space to dash: '
        If Right$(t_FindPrior, 1) = "-" Then
            If ck_Unopposed.Value = 1 Then
                If Right$(" " & t_FindPrior, 2) <> "P" Then
                    t_FindPrior = t_FindPrior & "P-"
                End If
            End If
            KeyAscii = 0
            Exit Sub
        End If
        k = "-"
        KeyAscii = 45
    End If

    If InStr("cdhsnprq", k) > 0 Then
        ' Uppercase letters: '
        KeyAscii = KeyAscii - 32
        k = UCase$(k)
    End If

    If t_FindPrior.SelLength = Len(t_FindPrior) Then
        t_FindPrior = ""
    ElseIf n > 0 Then ' Number was entered. '
        If t_FindPrior.SelStart > 1 Then
            If Val(Mid$(t_FindPrior, t_FindPrior.SelStart - 1)) = 0 Then
                GoTo ExitPrior
            End If
        End If
    End If

    If InStr("CDHSNPR@Q0", k) = 0 And _
       (k < "1" Or k > "7") And _
       KeyAscii <> 3 And _
       KeyAscii <> 8 And _
       KeyAscii <> 22 And _
       KeyAscii <> 24 And KeyAscii <> 67 And _
       k <> "-" _
    Then ' @ = slam bidding code followed by '
         ' TS (T=trump suit, S=2nd AA suit.) '
         ' QT (Q=control bid, T=trump suit.) '
         ' @Mw (Minorwood.) '
        ' Reject invalid keypress: '
        KeyAscii = 0
        Exit Sub
    ElseIf InStr("CDHSN", k) > 0 And _
           t_FindPrior.SelStart > 0 _
    Then
        ' A suit was entered. Add the suit and a dash: '
        If 1 = 1 And _
           (Val(Mid$(t_FindPrior, t_FindPrior.SelStart)) > 0) Or k = "D" _
        Then
            If Len(t_FindPrior) > t_FindPrior.SelStart Then
                ' Suit added inside the Prior Bids. '
                t_FindPrior = Left$(t_FindPrior, t_FindPrior.SelStart) & k & _
                              Mid$(t_FindPrior, t_FindPrior.SelStart + 1)
            Else
                t_FindPrior = t_FindPrior & k & "-"
                If ck_Unopposed.Value = 1 Then
                    If Right$(" " & t_FindPrior, 2) <> "P" Then
                        t_FindPrior = t_FindPrior & "P-"
                    End If
                End If
            End If
            t_FindPrior.SelStart = Len(t_FindPrior)
        Else
            Beep
            str_Response = MyMsgBox("Enter a number first.", 1, "ok", "")
        End If
        KeyAscii = 0
        Exit Sub

    ElseIf Len(t_FindPrior) > 1 And t_FindPrior.SelStart = 0 Then
        If k = "P" Then k = "P-"
        t_FindPrior = k & t_FindPrior
        KeyAscii = 0
        Exit Sub

    ElseIf k = "D" Then
        If t_FindPrior.SelStart > 2 Then
            ' Make sure the last entry was a bid like 1C (not D, R, or P). '
            If Val(Mid$(t_FindPrior, t_FindPrior.SelStart - 2)) > 0 Then
                If Len(t_FindPrior) > t_FindPrior.SelStart Then
                    ' Suit added inside the Prior Bids. '
                    t_FindPrior = Left$(t_FindPrior, t_FindPrior.SelStart) & k & _
                                  Mid$(t_FindPrior, t_FindPrior.SelStart + 1)
                Else
                    t_FindPrior = t_FindPrior & k & "-"
                End If
                t_FindPrior.SelStart = Len(t_FindPrior)
            Else
                Beep
                str_Response = MyMsgBox("An opponent's bid (e.g.: 1C) must" & vbCr & _
                                        "be entered before a Double.", 2, "ok", "")
            End If
        End If
        KeyAscii = 0
        Exit Sub

    ElseIf k = "R" Then
        If t_FindPrior.SelStart > 2 Then
            ' Make sure the last entry was a D. '
            If Mid$(t_FindPrior, t_FindPrior.SelStart - 2, 2) = "-D" Or _
               Right$(Space$(6) & t_FindPrior, 7) = "-D-P-P-" _
            Then
                If Len(t_FindPrior) > t_FindPrior.SelStart Then
                    ' Suit added inside the Prior Bids. '
                    t_FindPrior = Left$(t_FindPrior, t_FindPrior.SelStart) & k & _
                                  Mid$(t_FindPrior, t_FindPrior.SelStart + 1)
                Else
                    t_FindPrior = t_FindPrior & k & "-"
                End If
                t_FindPrior.SelStart = Len(t_FindPrior)
            Else
                Beep
                str_Response = MyMsgBox("A Redouble is not allowed here.", 1, "ok", "")
            End If
        Else
            Beep
            str_Response = MyMsgBox("A Redouble is not allowed here.", 1, "ok", "")
        End If
        KeyAscii = 0
        Exit Sub

    ElseIf k = "P" Then
        If Len(t_FindPrior) > t_FindPrior.SelStart Then
            ' Added inside the Prior Bids. '
            t_FindPrior = Left$(t_FindPrior, t_FindPrior.SelStart) & k & _
                          Mid$(t_FindPrior, t_FindPrior.SelStart + 1)
        Else
            t_FindPrior = t_FindPrior & k & "-"
            If ck_Unopposed.Value = 1 And Right$("  " & t_FindPrior, 2) <> "P-" Then
                If Right$(" " & t_FindPrior, 2) <> "P" Then
                    bln_SkipIt = True
                    t_FindPrior = t_FindPrior & "P-"
                End If
            End If
        End If
        t_FindPrior.SelStart = Len(t_FindPrior)
        KeyAscii = 0
        Exit Sub

    ElseIf Val(k) > 0 Or (InStr("PR", k) And _
           t_FindPrior.SelStart > 0) > 0 _
    Then
        If Len(t_FindPrior) > 1 And _
           Mid$(" " & t_FindPrior, t_FindPrior.SelStart + 1, 1) <> "-" _
        Then
            KeyAscii = 0
            Exit Sub
        End If
    End If

    If Len(t_FindHand(int_FindHandNum).Text) = 15 And _
       Chr$(KeyAscii) > "0" _
    Then
        If t_FindPrior = "" Then
            t_FindPrior.SetFocus
        Else
            b_FindBidForHand.SetFocus
        End If
    End If

ExitPrior:
    bln_SkipIt = True ' To keep t_FindPrior_Change from calling ConvertPrior. '
End Sub


Private Sub ConvertPrior(p As String)
    ' This sub takes auctions pasted into '
    ' the Find Bid's Prior Bids box and   '
    ' converts the into BidBase's format. '
    Dim p2 As String
    Dim b As String
    Dim c As String
    Dim i As Integer
    Dim j As Integer

    If t_FindPrior = "" Then Exit Sub
    If t_FindPrior.SelStart Then
        If Mid$(t_FindPrior, t_FindPrior.SelStart + 1, 1) = "-" And _
           Val(Mid$(t_FindPrior, t_FindPrior.SelStart, 1)) < Len(t_FindPrior) _
        Then
            Exit Sub
        End If
    End If
    
    p = UCase$(t_FindPrior)
    p = Replace$(p, Chr$(13), "")
    p = Replace$(p, Chr$(10), "")
    p = Replace$(p, Chr$(9), "")
    p2 = UCase(p)
    p2 = Replace$(p2, "MW", "Mw") ' Minorwood '
    p2 = Replace$(p2, "PASS", "P")
    p2 = Replace$(p2, "REDBL", "R")
    p2 = Replace$(p2, "XX", "R")
    p2 = Replace$(p2, "DBL", "D")
    p2 = Replace$(p2, "-X", "-D") ' @X is Exclusion Blackwood '
    p2 = Replace$(p2, "NT", "N")
    p2 = Replace$(p2, "*", "")
    p2 = Replace$(p2, "(", "")
    p2 = Replace$(p2, ")", "")
    p2 = Replace$(p2, " ", "")

    If Mid$(p2 & " ", 3, 1) = "@" Then
        ' In t_FindPrior, we want to leave in "4N@30-P-...", etc. '
        ' but nothing before it: "1C-P-2C-P-4C@Mw-P" to "4C@Mw-P" '
        i = InStr(p2, "@")
        If i > 3 Then p2 = Mid$(p2, i - 2)
        bln_SkipIt = True
        t_FindPrior = p2
        Exit Sub
    End If
    
'    p2 = Replace$(p2, "-", "") ' Ex:  1CDP1DP '
    p = ""

    Do
        j = j + 1
        If Mid$(p2 & " ", 3, 1) = "@" Then
            ' In t_FindPrior, we want to leave in "4N@30-P-...", etc. '
            ' but nothing before it: "1C-P-2C-P-4C@Mw-P" to "4C@Mw-P" '
            i = InStr(p2, "@")
            If i > 3 Then p2 = Mid$(p2, i - 2)
            bln_SkipIt = True
            t_FindPrior = p2
            Exit Sub
        End If
        If Val(p2) = 0 Then
            ' If p2 starts with a number, there is a 2-character bid; '
            ' otherwise, it is a 1-character call: P, D, R. '
            c = Left$(p2, 1)
            If InStr("PDR", c) = 0 Then
                MyMsgBox "Incorrect bid:" & c & vbCr & _
                         "", 3, "", "Convert Prior Bids"
                bln_SkipIt = False
                Exit Sub
            End If
            p = p & Left$(p2, 1) & "-"
            p2 = Trim(Mid$(p2 & "  ", 2)) ' Remove the call. '
            If Left$(p2, 1) = "-" Then p2 = Mid$(p2, 2)
        Else
            b = Left$(p2, 2)
            p = p & b & "-"
            p2 = Trim(Mid$(p2 & "  ", Len(b) + 1)) 'Remove the bid from p2
            If Left$(p2, 1) = "-" Then p2 = Mid$(p2, 2)
        End If
        DoEvents
    Loop While p2 <> ""
    bln_SkipIt = True
    t_FindPrior = p
    p = ""
End Sub


Private Sub t_FindPrior_Validate(Cancel As Boolean)
    Dim p As String
    Dim a As String
    Dim i As Integer
    Dim j As Integer
    Dim dash1 As Integer
    Dim dash2 As Integer

    If f_Msg.Visible Then Exit Sub ' This was looping somehow. '
    
    ConvertPrior t_FindPrior
End Sub


Private Sub t_GridData_DragDrop(Index As Integer, _
                                source As Control, _
                                x As Single, Y As Single)
    With source
        .Drag vbEndDrag
        .Move .Left + t_GridData(Index).Left + x - sng_DragoffsetX, _
              .Top + t_GridData(Index).Top + Y - sng_DragoffsetY
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub t_Hand_HCPs_KeyPress(KeyAscii As Integer)
    Dim k As String
    Dim p As Integer

    k = Chr$(KeyAscii)
    p = Val(k)
    If k = "m" Or k = "M" Then
        t_Hand_HCPs = "Min"
        KeyAscii = 0
    ElseIf k = "p" Then
        KeyAscii = Asc("P")
    ElseIf k = "h" Or k = "H" Then
        t_Hand_HCPs = "HCP1N"
    ElseIf InStr("1234567890-+NP", k) = 0 And _
        k <> "<" And _
        k <> ">" And _
        KeyAscii <> 3 And _
        KeyAscii <> 8 And _
        KeyAscii <> 22 And _
        KeyAscii <> 24 _
    Then
        KeyAscii = 0
    End If
End Sub


Private Sub t_HCP_GotFocus(Index As Integer)
    str_OrigData = t_HCP(Index)
    L_Status.Clear
    L_Status.Width = 18280
    Select Case Index
      Case 0
        ' "x" boxes: '
        L_Status.AddItem " Same type of input as boxes below. " & _
                         " Arrow down to see specs."
                         
        L_Status.AddItem " (Put x in boxes below for them to" & _
                         " match specs in this box.)"

      Case 5
        ' combo boxes: '
        L_Status.AddItem " Combined suit points for 2+ suits. " & _
                         " Examples: SHD>10,  CD<3,  SC=8-9, " & _
                         " xH>7 (enter 'x' first).  Enter" & _
                         " individual suit points above."
                         
        L_Status.AddItem " Ace asking: #A (# of Aces), #K, TA" & _
                         " (Trump Ace), TK, +TQ (Trump Q), -TQ" & _
                         " (no Trump Q).  Ex.: 2 or 5 Keys +TQ:" & _
                         " 2A+QT,1A+TK+TQ,4A+TK+TQ."
      Case Else
        ' suit boxes: '
        L_Status.AddItem " HCP (3+, 3-6, etc).  a=AK/KQ/AQ, b=AJT/KJT/QJT, c=AKQ/AKJ/AQJ/KQJ.  A,K (5+A = AJ not KQ,  bA = AJT only).  FRC=First-Round Control, SRC=Second-RC.  P=Protected honors. "
        L_Status.AddItem " 0stop=no stopper, hstop=Half-Stop=Qxx/Jxxx, 1stop=A/Kx/QJx/QT9x/Q98x/JT9x/J98x, 2stop=AK/AQ/AJT/AT98x/KQJ/KQT/KQ9x.  i=Intermed. (T,9,&8; enter as: 4+i#, etc.)  'or...' "
        L_Status.AddItem " H#=minimum # of honors + intermediates (3H3=KT942 rather than just K5432, same HCP but more control).  #H#=HCP & H# as before.  #/#=# of top # (2/3=2 of top 3)"
        L_Status.Height = 900
        L_Status.Width = 18600
    End Select
    t_HCP(Index).SelStart = 0
    t_HCP(Index).SelLength = 32768
    t_HCP(Index).ZOrder 0
End Sub


Private Sub t_HCP_KeyDown(Index As Integer, _
                          KeyCode As Integer, _
                          Shift As Integer)
    ' Enable use of cursor keys to move '
    ' vertically and horizontally:      '
    Select Case Shift
    Case 0
        Select Case KeyCode
            Case 38
                ' Up arrow - move to HCP box above: '
                If Index > 0 Then t_HCP(Index - 1).SetFocus

            Case 40
                ' Down arrow - move to HCP box below: '
                If Index < 5 Then t_HCP(Index + 1).SetFocus

            Case 37
                ' Left arrow - when cursor is at start '
                ' of field, move to Qty field on left: '
                If t_HCP(Index).SelStart = 0 And _
                   t_HCP(Index).SelLength = 0 _
                Then
                    t_Qty(Index).SetFocus
                End If

            Case 39
                ' Right arrow - when cursor is at end '
                ' of field, move down to the next Qty '
                ' field on the next row.              '
                If Index < 4 And _
                   t_HCP(Index).SelStart = Len(t_HCP(Index)) _
                Then
                    t_Qty(Index + 1).SetFocus
                End If
            End Select
    End Select
End Sub


Private Sub t_HCP_Validate(Index As Integer, Cancel As Boolean)
    On Local Error Resume Next
    L_Status.Clear
    Dim x As Integer

    If str_OrigData <> t_HCP(Index) Then
        If ck_AutoFill <> 0 Then
            t_TestHand = ""
            bln_Working = True
            CreateTestHand
        End If
        ck_Saved.Value = 0
    End If
    L_Status.Height = 630
    On Local Error GoTo 0
End Sub


Private Sub t_Hand_HCPs_GotFocus()
    str_OrigData = t_Hand_HCPs()
    L_Status.Clear
    L_Status.AddItem " High Card Points for the hand. " & _
                     " Examples: 6+, 6-, 11-18, >10, <17. "
                     
    L_Status.AddItem " For NT bids and responses: HCP1N," & _
                     " HCP2N, Min..., P... -- See the" & _
                     " Input Box documentation file."
                     
    t_Hand_HCPs.SelStart = 0
    t_Hand_HCPs.SelLength = 32768
End Sub


Private Sub t_Hand_HCPs_KeyDown(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = 40 And Shift = 0 Then
        ' Down cursor -- move down to next box: '
        't_BBPoints.SetFocus
        t_Distrib_Pts.SetFocus
    End If
End Sub


Private Sub t_Hand_HCPs_Validate(Cancel As Boolean)
    Dim x As Integer

    L_Status.Clear
    If str_OrigData <> t_Hand_HCPs Then
        ck_Saved.Value = 0

        If TotalDistribPts Then
            Exit Sub
        End If

        If ck_AutoFill.Value <> 0 Then
            t_TestHand = ""
            bln_Working = True
            CreateTestHand
            Do While bln_Working
                DoEvents
            Loop
        End If

    End If
End Sub


Private Sub t_Intermediates_GotFocus()
    L_Status.Clear
    L_Status.AddItem " Intermediates: minimum # of 10's and" & _
                     " 9's in the hand, plus 8's in the company" & _
                     " of 10's and 9's."
                     
    L_Status.AddItem " The same intermediate card can be counted" & _
                     " twice - once for a suit spec and again for" & _
                     " the hand spec."
                     
    L_Status.AddItem " 0 or 1 is for poor intermediates, so more" & _
                     " would fail. If 2 or higher is specified," & _
                     " more would pass."
                     
    str_OrigData = t_Intermediates
    t_Intermediates.SelStart = 0
    t_Intermediates.SelLength = 32768
End Sub


Private Sub t_Intermediates_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
    Case Is = 0
        Select Case KeyCode
        Case Is = 38
            t_Stoppers.SetFocus
        End Select
    End Select
End Sub

Private Sub t_Intermediates_Validate(Cancel As Boolean)
    Dim n As Integer

    L_Status.Clear
    If str_OrigData <> t_Intermediates Then
        n = Len(t_Intermediates)
        If n > 1 Then
            str_Response = _
               MyMsgBox("This field can only be 1 digit.", _
                         1, "ok", "Intermediates")
            Cancel = True
        Else
            ck_Saved.Value = 0
            If ck_AutoFill.Value <> 0 Then
                t_TestHand = ""
                bln_Working = True
                CreateTestHand
                Do While bln_Working
                    DoEvents
                Loop
            End If
        End If
    End If
End Sub


Private Sub t_Len_GotFocus(Index As Integer)
    str_OrigData = t_Len(Index)
    L_Status.Clear
    L_Status.AddItem " ## = expected suit length (e.g.: open " & _
                     " 1S = 5) -- x# = x-spec #,  /# = alternative #"
                     
    L_Status.AddItem " #F = expected length of Fit for pard, " & _
                     " H = Help suit (or 2nd suit),  S = " & _
                     " Singleton/void,  C = Cue bid of A/K."
End Sub


Private Sub t_Len_KeyDown(Index As Integer, _
                          KeyCode As Integer, _
                          Shift As Integer)
    Select Case Shift
    Case Is = 0
        Select Case KeyCode
            Case Is = 38
                If Index > 1 Then t_Len(Index - 1).SetFocus

            Case Is = 40
                If Index < 4 Then t_Len(Index + 1).SetFocus
        End Select
    End Select
End Sub

Private Sub t_Losers_GotFocus()
    str_OrigData = t_Losers
    L_Status.Clear
    L_Status.AddItem " Maximum number of losers in top 3 cards" & _
                     " of each suit.  Examples: AQ=.5, Kx=1,"
                     
    L_Status.AddItem " AQJ=1, AQx=1.5, xx=2, Kxx=2, Axx=2, QJx=2," & _
                     " QTx-2, Qxx=2.5, xxx=3. Can use #-#."
    t_Losers.SelStart = 0
    t_Losers.SelLength = 32768
End Sub


Private Sub t_Losers_KeyDown( _
                KeyCode As Integer, _
                Shift As Integer)
    Select Case Shift
      Case Is = 0
        Select Case KeyCode
            Case Is = 38
                t_Winners.SetFocus
            Case Is = 40
                t_QTs.SetFocus
        End Select
    End Select
End Sub


Private Sub t_Losers_Validate(Cancel As Boolean)
    L_Status.Clear
    If str_OrigData <> t_Losers Then
        ck_Saved.Value = 0
    End If
End Sub


Private Sub t_Order_GotFocus()
    str_OrigData = t_Order
    L_Status.Clear
    L_Status.AddItem " Number controlling the sequence in" & _
                     " which bid entries are considered."
    t_Order.SelStart = 0
    t_Order.SelLength = 32768
End Sub


Private Sub t_Order_Validate(Cancel As Boolean)
    Dim x As Integer

    L_Status.Clear
    If str_OrigData <> t_Order Then
        If ck_Saved.Value = 1 Then
            ck_Saved.Value = 0
        End If
    End If
End Sub


Private Sub t_OtherHands_DblClick()
    Dim n As Long
    Dim r As Integer
    Dim c As Integer
    
        str_Response = MyMsgBox("Y = Clear all columns" & vbCr & _
                                "N = Clear this column only", _
                                 2, "YNC", "Hand Specs")
        
        With g_OtherHands
        '''''''''''''''''
        If str_Response = "yes" Then
            .Clear
            SetupHandDataGrid
        ElseIf str_Response = "no" Then
            c = .Col
            For r = 1 To 15
                .TextMatrix(r, c) = ""
            Next
        End If
        ''''''''
        End With
End Sub


Private Sub t_OtherHands_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim n As Long
    Dim r As Integer
    Dim c As Integer
    
    If Shift > 0 Then Exit Sub
    
    With g_OtherHands
    '''''''''''''''''
    c = .Col
    r = .Row
    n = Val(t_OtherHands.Text)
    If c > 0 And r = 1 And n > 0 And KeyCode = 13 Then
        DisplayOtherSpecs c, n
        Exit Sub
    End If
    ''''''''
    End With

    With g_OtherHands
    '''''''''''''''''
    If KeyCode = 38 Then
        If .Row > 1 Then .Row = .Row - 1
    ElseIf KeyCode = 40 Then
        If .Row < 15 Then .Row = .Row + 1
    ElseIf KeyCode = 37 Then
        If .Col > 1 Then .Col = .Col - 1
    ElseIf KeyCode = 39 Then
        If .Col < 4 Then .Col = .Col + 1
    End If
    g_OtherHands_Click
    ''''''''
    End With
End Sub


Private Sub t_Pct_GotFocus()
    str_OrigData = t_Pct
    L_Status.Clear
    L_Status.AddItem " Percentage of the time this entry" & _
                     " should be used when there is an alternative. " & _
                     " 0=never used by either side."
                     
    L_Status.AddItem " Leave blank for 100%.  0=deactivate. " & _
                     " W=we.  T=they.  blank=both.  M=Matchpoints" & _
                     " only.  I=IMPs only."
    t_Pct.SelStart = 0
    t_Pct.SelLength = 32768
End Sub


Private Sub t_Pct_Validate(Cancel As Boolean)
    L_Status.Clear
    If str_OrigData <> t_Pct Then
        ck_Saved.Value = 0
    End If
End Sub


Private Sub t_Player_GotFocus()
    str_OrigData = t_Player
    L_Status.Clear
    L_Status.AddItem " Player number, starting with first non-passing player."
    t_Prior.SelStart = 0
    t_Prior.SelLength = 32768
End Sub


Private Sub t_Player_Validate(Cancel As Boolean)
    L_Status.Clear
    t_Player = Trim$(t_Player)
    If str_OrigData <> t_Player Then
        ck_Saved.Value = 0
    End If
End Sub


Private Sub t_Prior_GotFocus()
    str_OrigData = t_Prior
    L_Status.Clear
    L_Status.AddItem " Prior bids starting with first non-passing bid. "
    L_Status.AddItem " Example: 1C-D-R-P-2C where D=Double, R=Redouble, P=Pass"
    t_Prior.SelStart = 0
    t_Prior.SelLength = 32768
End Sub


Private Sub t_Prior_KeyPress(KeyAscii As Integer)
    Dim k As String
    Dim n As Integer
    Dim i As Integer
    Dim sStart As Integer
    Dim x As Boolean

    Dim pb As String
    Dim p As Integer
    Dim r As Integer
    Dim result As Boolean
    Dim L As Integer
    Dim c As String

    If KeyAscii = 3 Then Exit Sub
    k = Chr$(KeyAscii)

    With t_Prior
    ''''''''''''
    ' Process slam bids with "@" in them: '
    'i = InStr(.Text, "@")
    'If i = Len(.Text) Then  ' Ex: 4D@ '.
    '    If k = "M" Then ' Minorwood (Mw) '
    '
    '    ElseIf k = "X" Or k = "x" Then ' Exclusion Blackwood '
    '        Exit Sub
    '    End If
    'End If
    
    sStart = .SelStart
    If sStart > 0 Then
        ' Part of a test to make sure 2 digits in a row were not pressed. '
        x = Val(Mid$(" " & .Text, .SelStart, 1)) > 0
    End If

    n = Val(k)
    If n > 7 Or (n > 0 And n < 8 And x) Then
        Beep
        str_Response = MyMsgBox("A number to bid must be 1-7.", 1, "ok", "")
        KeyAscii = 0
        GoTo ExitPrior

    ElseIf k = " " Then
        ' Change space to dash: '
        If Right$(.Text, 1) = "-" Then
            KeyAscii = 0
            GoTo ExitPrior
        End If
        k = "-"
        KeyAscii = 45
    End If

    If k <> "m" And k <> "w" And KeyAscii > 96 And _
       KeyAscii < 123 _
    Then
        ' Change to uppercase letter: '
        KeyAscii = KeyAscii - 32
        k = UCase$(k)
    End If

    If .SelLength = Len(.Text) Then
        .Text = ""
    ElseIf n > 0 Then ' Number was entered. '
        If .SelStart > 1 Then
            If Val(Mid$(.Text, .SelStart - 1)) = 0 Then
                GoTo ExitPrior
            End If
        End If
    End If
    If InStr("CDHSNPR-AKQ@MmMwX", k) = 0 And _
       (k < "0" Or k > "9") And _
       KeyAscii <> 3 And _
       KeyAscii <> 8 And _
       KeyAscii <> 22 And _
       KeyAscii <> 24 And _
       k <> "-" _
    Then ' @ = slam bidding code followed by '
         ' TS (T=trump suit, S=2nd AA suit.) '
         ' QT (Q=control bid, T=trump suit.) '
         ' Mm (Any Major/minor suit '
         ' Mw (Minorwood) '
         ' Xs (Exclusion Blackwood followed by suit.) '
        ' Cancel invalid character entered: '
        KeyAscii = 0
        GoTo ExitPrior
    End If

    If InStr("CDHSNMm", k) > 0 And _
       .SelStart > 0 _
    Then
        ' A suit was entered. Add the suit and a dash: '
        If 1 = 1 And _
           (Val(Mid$(.Text, .SelStart)) > 0) Or Mid$(.Text, .SelStart, 1) = "X" Or k = "D" Or UCase$(k) = "M" _
        Then ' m is Minorwood '
            If Len(.Text) > .SelStart Then
                ' Suit added inside the Prior Bids. '
                .Text = Left$(.Text, .SelStart) & k & _
                        Mid$(.Text, .SelStart + 1)
            Else
                .Text = .Text & k & "-"
            End If
            .SelStart = Len(.Text)
        Else
            Beep
            str_Response = MyMsgBox("Enter a number first.", 1, "ok", "")
        End If
        KeyAscii = 0
        GoTo ExitPrior

    ElseIf Len(.Text) > 1 And .SelStart = 0 Then
        .Text = k & .Text

    ElseIf k = "D" Then
        If .SelStart > 2 Then
            ' Make sure the last entry was a '
            ' bid like 1C (not D, R, or P).  '
            If Val(Mid$(.Text, .SelStart - 2)) > 0 Then
                If Len(.Text) > .SelStart Then
                    ' Suit added inside the Prior Bids. '
                    .Text = Left$(.Text, .SelStart) & k & _
                            Mid$(.Text, .SelStart + 1)
                Else
                    .Text = .Text & k & "-"
                End If
                .SelStart = Len(.Text)
            Else
                Beep
                str_Response = _
                   MyMsgBox("An opponent's bid (e.g.: 1C) must" & vbCr & _
                            "be entered before a Double.", 2, "ok", "")
            End If
        End If

        KeyAscii = 0
        GoTo ExitPrior

    ElseIf k = "R" Then
        If .SelStart > 2 Then
            ' Make sure the last entry was a D. '
            If Mid$(.Text, .SelStart - 2, 2) = "-D" Or _
               Right$(Space$(6) & .Text, 7) = "-D-P-P-" _
            Then
                If Len(.Text) > .SelStart Then
                    ' Suit added inside the Prior Bids. '
                    .Text = Left$(.Text, .SelStart) & k & _
                            Mid$(.Text, .SelStart + 1)
                Else
                    .Text = .Text & k & "-"
                End If
                .SelStart = Len(.Text)
            Else
                Beep
                str_Response = MyMsgBox("A Redouble is not allowed here." _
                                         , 2, "ok", "")
            End If
        Else
            Beep
             str_Response = MyMsgBox("A Redouble is not allowed here." _
                                      , 2, "ok", "")
        End If
        KeyAscii = 0
        GoTo ExitPrior

    ElseIf k = "P" Then
        If Len(.Text) > .SelStart Then
            ' Added inside the Prior Bids. '
            .Text = Left$(.Text, .SelStart) & k & _
                    Mid$(t_FindPrior, t_FindPrior.SelStart + 1)
        Else
            .Text = .Text & k & "-"
        End If
        .SelStart = Len(.Text)
        KeyAscii = 0
        GoTo ExitPrior

    ElseIf Val(k) > 0 Or (InStr("PR", k) And _
           .SelStart > 0) > 0 _
    Then
        If Len(.Text) > 1 And _
           Mid$(" " & .Text, .SelStart + 1, 1) <> "-" _
        Then
            KeyAscii = 0
            GoTo ExitPrior
        End If
    End If

ExitPrior:
    If Len(t_Prior) > 3 And k = "-" Then
        pb = t_Prior
        p = Val(t_Player)
        r = Val(t_Round)
        ValidateBidding pb, p, r, result
    End If

    ''''''''
    End With
End Sub


Private Sub t_Prior_Validate(Cancel As Boolean)
    Dim pb As String
    Dim p As Integer
    Dim r As Integer
    Dim result As Boolean
    Dim L As Integer

    If f_Msg.Visible Then Exit Sub

    t_Prior = Trim$(t_Prior)
    L = Len(t_Prior)
    If Right$(t_Prior, 1) = "-" And L > 1 Then
        t_Prior = Left$(t_Prior, L - 1)
    End If

    If str_OrigData <> t_Prior Then
        pb = t_Prior
        p = Val(t_Player)
        r = Val(t_Round)
        ValidateBidding pb, p, r, result
        If result = False Then
            t_Prior.SetFocus
            t_Prior.SelStart = Len(t_Prior)
            Exit Sub
        End If
    End If
    L_Status.Clear
End Sub


Private Sub t_QTs_GotFocus()
    str_OrigData = t_QTs
    ' L_Status is the wide List control at '
    ' the bottom of the window. Text does  '
    ' not wrap. It is broken up into lines '
    ' which are added separately.          '
    L_Status.Clear
    L_Status.AddItem " # of defensive tricks in 1st 2 rounds of a suit: " & _
                     " Kx=.5, A or KQ=1, AQ-1.5 AK=2.  If a suit has 7+"
                     
    L_Status.AddItem " cards, only count the Ace in the suit because of" & _
                     " ruffing by opponents; if 6 cards, only count AK."
    t_QTs.SelStart = 0
    t_QTs.SelLength = 32768
End Sub


Private Sub t_QTs_KeyDown( _
                KeyCode As Integer, _
                Shift As Integer)
    Select Case Shift
      Case Is = 0
        Select Case KeyCode
            Case Is = 38
                t_Losers.SetFocus
            Case Is = 40
                t_Stoppers.SetFocus
        End Select
    End Select
End Sub


Private Sub t_QTs_Validate(Cancel As Boolean)
    L_Status.Clear
    If str_OrigData <> t_QTs Then
        ck_Saved.Value = 0
    End If
End Sub


Private Sub t_Qty_GotFocus(Index As Integer)
    str_OrigData = t_Qty(Index)
    ' L_Status is the wide List control at '
    ' the bottom of the window. Text does  '
    ' not wrap. It is broken up into lines '
    ' which are added separately.          '
    L_Status.Clear
    Select Case Index
      Case Is = 0
        L_Status.AddItem " # in Suit, e.g.: 5+, <4, 2-5." & _
                         " (Put x in suits below to match" & _
                         " number in this box.)"
                         
        L_Status.AddItem " Add 'U' to end to use lower-ranked" & _
                         " of equal-length suits."

      Case Is = 5
        L_Status.AddItem " Combined suit total specs for 2-3 suits." & _
                         " Examples: CH=8, SHD>10, CD<3, SC=8-9. " & _
                         " Enter individual suit ranges above."
                         
        L_Status.AddItem " For more info about the M code, see" & _
                         " the Editor Input Boxes help file - Prior Bids."

      Case Else
        L_Status.AddItem " # in Suit, e.g.:  5+,  <4,  2-5,  >C, " & _
                         " >=D,  <H,  =S,  x,  <x,  <=x,  >x,  >=x, " & _
                         " 2-S (from 2 to # of Spades, max)"
    End Select
    t_Qty(Index).SelStart = 0
    t_Qty(Index).SelLength = 32768
    t_Qty(Index).ZOrder 0
End Sub


Private Sub t_Qty_KeyDown(Index As Integer, _
                          KeyCode As Integer, _
                          Shift As Integer)
    Select Case Shift
    Case Is = 0
        Select Case KeyCode
            Case Is = 38
                If Index > 0 Then t_Qty(Index - 1).SetFocus

            Case Is = 40
                If Index < 5 Then t_Qty(Index + 1).SetFocus

            Case Is = 37
                If Index > 0 And _
                   t_Qty(Index).SelStart = 0 And _
                   t_Qty(Index).SelLength = 0 _
                Then
                    t_HCP(Index - 1).SetFocus
                End If

            Case Is = 39
                If t_Qty(Index).SelStart = Len(t_Qty(Index)) Then
                    t_HCP(Index).SetFocus
                End If
        End Select
    End Select
End Sub


Private Sub t_Qty_LostFocus(Index As Integer)
    Dim pr As String
    Dim prH As String
    Dim prS As String
    Dim t As String
    
    If Index < 3 And _
       InStr(cmb_Bid, "x") > 0 And Val(cmb_Bid) > 0 And _
       InStr(t_Qty(Index), "x") > 0 And _
       InStr(t_Prior, "M") > 0 And f_Msg.Visible = False _
    Then
        pr = t_Prior
        prH = Replace(pr, "M", "H") & "-" & Left$(cmb_Bid, 1) & "H"
        prS = Replace(pr, "M", "S") & "-" & Left$(cmb_Bid, 1) & "S"
        bln_Abort = False
        t = "With 'M' in the Prior Bids," & vbCrLf
        t = t & "possible auctions in which" & vbCrLf
        t = t & "the 'M' letter is bid (H/S)" & vbCrLf
        t = t & "are " & prH & " and " & prS & "."
        str_Response = MyMsgBox(t, 3, "Docs", "Combined H/S Entries")
        
        If str_Response = "docs" Then
            str_NoteFile(0) = "CardShark_BidBase_Combining_HS_Entries.htm"
            ShowDocs
        End If
    End If
End Sub


Private Sub t_Qty_Validate(Index As Integer, Cancel As Boolean)
    Dim x As Integer

    L_Status.Clear
    If str_OrigData <> t_Qty(Index) Then
        ck_Saved.Value = 0
        If ck_AutoFill.Value <> 0 Then
            t_TestHand = ""
            CreateTestHand
            'CreateDistribDisclosure
        End If
    End If
End Sub


Private Sub t_Round_GotFocus()
    str_OrigData = t_Round
    L_Status.Clear
    L_Status.AddItem " Bidding round number, i.e.: number of set of four bids " & _
                     "starting with first non-passing bid, "
    L_Status.AddItem " so 'P-P-P-1C-1D-1H-1S' is still the first round of bidding."
    t_Prior.SelStart = 0
    t_Prior.SelLength = 32768
End Sub


Private Sub t_Round_Validate(Cancel As Boolean)
    L_Status.Clear
    t_Round = Trim$(t_Round)
    If str_OrigData <> t_Round Then
        ck_Saved.Value = 0
    End If
End Sub


Private Sub t_Show_DragDrop(Index As Integer, _
                            source As Control, _
                            x As Single, Y As Single)
    With source
        .Drag vbEndDrag
        .Move .Left + t_Show(Index).Left + x - sng_DragoffsetX, _
              .Top + t_Show(Index).Top + Y - sng_DragoffsetY
        .ZOrder 0
    End With
    bln_Dragging = False
    sng_DragoffsetX = 0
    sng_DragoffsetY = 0
End Sub


Private Sub t_Show_GotFocus(Index As Integer)
    If Index = 1 Or Index = 3 Then
        t_Show(Index) = t_Show(Index - 1)
    End If
    t_Show(Index).SelStart = 0
    t_Show(Index).SelLength = 32768
End Sub


Private Sub t_Show_KeyPress(Index As Integer, KeyAscii As Integer)

    If Index < 2 Then
        If cmb_Show(0) <> "Prior_Bids" Then
            Exit Sub
        End If
    Else
        If cmb_Show(1) <> "Prior_Bids" Then
            Exit Sub
        End If
    End If

    ' Prior Bids are being entered, '
    ' so let's verify input data:   '
    Dim k As String

    k = Chr$(KeyAscii)
    If k = "W" Then k = "w" ' for 4C@Mw = Minorwood '
    If k = " " Then
        k = "-"
        KeyAscii = 45
    End If

    If KeyAscii = 122 Then ' z '
        Exit Sub

    ElseIf KeyAscii > 96 And _
       KeyAscii < 123 And k <> "w" _
    Then
        ' Change to uppercase letter: '
        KeyAscii = KeyAscii - 32
        k = UCase$(k)
    End If
End Sub

Private Sub t_Show_Validate(Index As Integer, Cancel As Boolean)
    L_Status.Clear
    If str_OrigData <> t_Show(Index) Then
        ck_Saved.Value = 0
    End If
End Sub


Private Sub t_Ref_DblClick()
    ' If the source is a deal number, open the '
    ' BidBase Dealer, pass it the deal number  '
    ' and dealer #, if specified.              '
    ShowSourceDeal t_Ref
End Sub


Private Sub ShowSourceDeal(s As String)
    Dim xxx As Long
    Dim DealNum As String
    Dim dealer As String
    Dim i As Integer

    s = Trim$(s)
    If s = "" Then Exit Sub
    If Left$(s + Space$(5), 6) <> "Deal #" Then Exit Sub

    ' Look for end of deal #: '
    i = 8
    Do While Mid$(s & " ", i, 1) >= "0" And _
             Mid$(s & " ", i, 1) <= "9"
        i = i + 1
        DoEvents
    Loop
    DealNum = Trim$(Mid$(s, 7, i - 6))

    If DealNum = "" Then Exit Sub

    i = InStr(s, ",")
    If i > 0 Then
        dealer = Mid$(s, i + 1, 1)
    End If
    If dealer = "" Then dealer = "S"
'Stop ' I don't think this works in Win10. '
    On Local Error GoTo LoadDealer
    ' If Player is not already loaded,   '
    ' AppActivate will generate an error '
    ' which will Shell to the player.    '
    ' This keeps more than one player    '
    ' from opening.                      '
    AppActivate "CardShark BidBase Practice Dealer"
    'For iii! = 1 To 100000: DoEvents: Next
    'SendKeys "%d", True ' Alt-D to start new deal '
    'For iii! = 1 To 100000: DoEvents: Next

ContinueDeal:
    On Local Error GoTo 0
    'SendKeys dealNum & "{ENTER}", True
    'SendKeys "%l", True
    'SendKeys dealer & "{ENTER}", True
Exit Sub

LoadDealer:
    xxx = Shell(str_BBpath & "\BidBase Bidding Practice.exe", 1)
    AppActivate "CardShark BidBase Practice Dealer"
    Resume ContinueDeal

End Sub


Private Sub t_Ref_GotFocus()
    str_OrigData = t_Ref
    If list_Refs.Visible Then
        list_Refs.Visible = False
        b_DelRef.Visible = False
        Exit Sub
    End If
    L_Status.Clear
    L_Status.Width = 17500
    L_Status.AddItem " Source of specs/test hand:  Aces=Bobby Wolff's column,  " & _
                     "BCC=Bridge Conventions Complete,  BD=Bidding Dict.,  " & _
                     "Becker=column,  Bul=ACBL Bulletin,  BW=Bridge World,"
    L_Status.AddItem " MB=Marty Bergen,  MH=Max Hardy,  ML=Mike Lawrence.  " & _
                     "If it starts with 'Deal #', double-click it to view " & _
                     "the whole deal. (Suits are swapped for similar Prior Bids.)"
End Sub


Private Sub t_Ref_Validate(Cancel As Boolean)
    Dim i As Integer
    Dim j As Integer
    Dim r As String

    'If Left$(t_Ref & "      ", 6) = "Deal #" Then
    '    r = "Deal #"
    'Else
        r = t_Ref.Text
    'End If

    If Len(t_Ref) > 25 Then
        str_Response = MyMsgBox("Maximum input is 25 characters.", 1, "ok", "Bid Reference")
        t_Ref.SetFocus
        Cancel = True
        Exit Sub
    End If
    L_Status.Clear
    
    ' 2021-07-14: Automatically removing lines causes problems. '
    Exit Sub
    
    If str_OrigData <> t_Ref.Text And Trim$(t_Ref.Text) <> "" Then
        ck_Saved.Value = 0
        'list_Refs.AddItem r
        Do While list_Refs.List(0) = ""
            list_Refs.RemoveItem 0
        Loop
        
        t_Ref.Text = r
        
        j = InStr(t_Ref, " ")
        If j > 0 Then
          For i = 0 To list_Refs.ListCount - 1
            If t_Ref = list_Refs.List(i) Then
                Exit For
            End If
            If Left$(t_Ref, j) = Left$(list_Refs.List(i) & Space$(j), j) Then
                ' We know from above that the entries '
                ' don't match exactly, so if they do  '
                ' match at the start, we can delete   '
                ' the entry in the list. '
                list_Refs.RemoveItem i
                i = list_Refs.ListCount
            End If
          Next
          If i = list_Refs.ListCount Then
            list_Refs.AddItem t_Ref.Text
          End If
        End If
    End If
End Sub


Private Sub t_Stoppers_GotFocus()
    str_OrigData = t_Stoppers
    L_Status.Clear
    L_Status.AddItem " Number of suits stopped other than partner's suit.  " & _
                     "(If NT bid, all suits count.)"
    t_Stoppers.SelStart = 0
    t_Stoppers.SelLength = 32768
End Sub


Private Sub t_Stoppers_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
    Case Is = 0
        Select Case KeyCode
        Case Is = 38
            t_QTs.SetFocus
        Case Is = 40
            t_Intermediates.SetFocus
        End Select
    End Select
End Sub


Private Sub t_Stoppers_Validate(Cancel As Boolean)
    L_Status.Clear
    If str_OrigData <> t_Stoppers Then
        ck_Saved.Value = 0
    End If
End Sub


Private Sub t_Sel_GotFocus()
    str_OrigData = Trim$(t_Sel)
    L_Status.Clear
    L_Status.AddItem " W=We,  T=They,  [blank]=both use this entry,  0=neither uses it."
End Sub


Private Sub t_Sel_Validate(Cancel As Boolean)
    Dim c As String
    Dim c1 As String
    Dim c2 As String
    Dim c3 As String
    Dim c4 As String
    Dim i As Integer
    Dim x As Integer

    t_Sel = Trim$(t_Sel)
    If str_OrigData <> t_Sel Then
        c = Trim$(cmb_ConvName(0))
        If c = "" Then Exit Sub
        i = InStr(c, ":")
        If i > 1 Then
            c1 = Left$(c, i - 1)
        Else
            c1 = c
        End If

        ' Step thru all bid names to look for '
        ' the name in cmb_ConvName(0). If the '
        ' name is anywhere in a record, change'
        ' .Selected for the record. Example:  '
        ' Splinter is in cmb_ConvName(0) --   '
        ' Records to change:                  '
        '   1. Splinter                       '
        '   2. Splinter: Continuation         '
        '   3. Strong 2C: Splinter            '
        '   4. Strong 2C: Splinter: Gerber    '
        ' etc.                                '
        Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
        Set rs_BidNamesTbl = db_BidBase.OpenRecordset(str_BidNamesTbl)
        With rs_BidNamesTbl
            .Index = "Bid_Name"
            .MoveFirst

            If .NoMatch Then
                str_Response = MyMsgBox(c1 & vbCr & " was not found.", 2, "ok", "")
                Exit Sub
            Else
                If !Selected <> t_Sel And cmb_ConvName(0) = !Bid_Name Then
                    ' Something was making the program     '
                    ' think that the t_Sel box had changed '
                    ' when it hadn't, so...                '
                    str_Response = MyMsgBox("Are you sure you want to change the " & vbCr & _
                                            "Conv. code to " & t_Sel & " for all " & vbCr & _
                                            "entries using this convention?", 3, "YN", "BidBase")
                    If str_Response <> "yes" Then
                        Exit Sub
                    End If
                End If
                i = Len(c1) + 1
                ' c is the full conv/subconv name. '
                ' c1 is the conv name, such as "Splinter"
                ' c2 = "Splinter:", as in "Splinter: Continuation"
                ' c3 = ": Splinter", as in "Strong 2C: Splinter"
                ' c4 = ": Splinter:", as in "Strong 2C: Splinter: Gerber"
                c2 = c1 & ":"
                c3 = ": " & c1
                c4 = c3 & ":"
                Do While Not .NoMatch And Not .EOF
                    If !Bid_Name = c1 Or _
                       Left$(!Bid_Name, i) = c2 _
                    Then
                        .Edit
                        !Selected = Left$(t_Sel & " ", 1)
                        .Update

                    ElseIf Right$(!Bid_Name, i + 1) = c3 Or _
                           InStr(!Bid_Name, c4) > 0 _
                    Then
                        .Edit
                        !SubCat_Selected = Left$(t_Sel & " ", 1)
                        .Update

                    End If
                    .MoveNext
                    DoEvents
                Loop
            End If
        End With

        On Local Error Resume Next
        'Data2(0).Refresh   ' BidNames data control.'
        For i = 0 To 3  ' Bids data controls.   '
            Data1(i).Refresh
        Next
        On Local Error GoTo 0
    End If
    L_Status.Clear
    str_OrigData = t_Sel
End Sub


Private Sub t_Sub_GotFocus()
    str_OrigData = Trim$(t_Sub)
    L_Status.Clear
    L_Status.AddItem " W=We,  T=They,  [blank]=both use this sub-convention,  0=neither uses it."
    L_Status.AddItem " Changing this field will change every entry in the database that is part of this sub-convention."
End Sub


Private Sub t_Sub_Validate(Cancel As Boolean)
    Dim c As String
    Dim i As Integer

    t_Sub = Trim$(t_Sub)
    If str_OrigData <> t_Sub Then
        c = Trim$(cmb_ConvName(0))
        i = InStr(c, ":")
        If i = 0 Then
            str_Response = MyMsgBox("This field is for subconventions, but" & vbCr & _
                                    "the 'Bid/Conv.Name' does not have one.", 2, "ok", "")
            Exit Sub
        End If

        Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
        Set rs_BidNamesTbl = db_BidBase.OpenRecordset(str_BidNamesTbl)
        With rs_BidNamesTbl
            .Index = "Bid_Name"
            .Seek "=", c
            If .NoMatch Then
                str_Response = MyMsgBox(c & vbCr & " was not found.", 2, "ok", "")
                Exit Sub
            Else
                .Edit
                !SubCat_Selected = Left$(t_Sub & " ", 1)
                .Update
            End If
        End With

        On Local Error Resume Next
       ' Data2.Refresh
        For i = 0 To 4
            Data1(i).Refresh
        Next
        On Local Error GoTo 0
        str_Response = MyMsgBox("Convention Names table updated.", 1, "ok", "")
    End If
    L_Status.Clear
End Sub

Private Sub t_TestHand_GotFocus()
    str_OrigData = t_TestHand
    If t_TestHand = "" And ck_AutoFill.Value = 1 Then CreateTestHand
    L_Status.Clear
    L_Status.AddItem " A hand matching the specs for this entry for testing the database."
    L_Status.AddItem " Example: AK32-KT2--AQJT32.  Suit order: SHDC.  Click Auto-Fill & click this box."
    DoEvents
    't_TestHand.SelStart = 32768
    't_TestHand.SelLength = 0
    DoEvents
End Sub


Private Sub t_TestHand_KeyPress(KeyAscii As Integer)
    Dim k As String
    Dim i As Integer
    Dim j As Integer
    Dim ss As Integer

    k = Chr$(KeyAscii)
    If k = " " Then
        k = "-"
        KeyAscii = 45
    End If

    If InStr("akqjt", k) > 0 Then
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
        KeyAscii = 0
        Exit Sub
    End If

    i = InStr("AKQJT98765432", k)
    If t_TestHand.SelStart > 0 And i > 0 Then
        j = InStr("AKQJT98765432", Mid$(t_TestHand, t_TestHand.SelStart, 1))
        If j >= i Then 'card entered is out of rank order
            ss = t_TestHand.SelStart
            t_TestHand = Left$(t_TestHand, _
                         t_TestHand.SelStart) + k + _
                         Mid$(t_TestHand, t_TestHand.SelStart + 1)
            t_TestHand.SelStart = ss + 1
            str_Response = MyMsgBox("Card spec is out of order.", 1, "ok", "")
            KeyAscii = 8
            Exit Sub
        End If
    End If
End Sub


Private Sub t_TestHand_Validate(Cancel As Boolean)
    Dim x As Integer

    L_Status.Clear
    If str_OrigData <> t_TestHand Then
        ck_Saved.Value = 0

        LookForNoteFile t_Disclosure, _
                        t_Prior, _
                        cmb_Bid, _
                        "" & lng_ID, _
                        False, _
                        True, _
                        False
    End If
End Sub


Private Sub t_Distrib_Pts_GotFocus()
    Dim k As Integer

    str_OrigData = t_Distrib_Pts
    L_Status.Clear
    L_Status.AddItem " Distribution points.  Examples: 2+, <3, 1-4, >1 or "
    L_Status.AddItem " Total Pts (Dist.+HCP) - precede with 'T': T10+, T<9, T11-13"

    If mPreviewHandSpecs.Checked Then
        str_Response = MyMsgBox("Blank out the fields from here down?", 1, "YN", "BidBase")

        If str_Response = "yes" Then
            t_Distrib_Pts = ""
            t_CustPoints = ""
            For k = 0 To 3
                t_CustPts(k) = ""
            Next
            t_Winners = ""
            t_Losers = ""
            t_QTs = ""
            t_Stoppers = ""
            cmb_Shape.ListIndex = 0
            t_Intermediates = ""
            t_Qty(0).SetFocus
        End If

        mPreviewHandSpecs.Checked = False
    End If

    t_Distrib_Pts.SelStart = 0
    t_Distrib_Pts.SelLength = 32768
End Sub


Private Sub t_Distrib_Pts_LostFocus()
    t_Distrib_Pts_Validate 0
End Sub


Private Sub t_Distrib_Pts_Validate(Cancel As Boolean)
    Dim x As Integer

    L_Status.Clear
    If str_OrigData <> t_Distrib_Pts Then
        ck_Saved.Value = 0

        If Left$(t_Distrib_Pts & " ", 1) <> "T" Then
            x = Val(t_Distrib_Pts)
            If x > 6 Then
                ' Must be entering total points but left of the "T". '
                'MsgBox "If you are entering total points, put a 'T' at the start, such as T" & t_Distrib_Pts
                t_Distrib_Pts = "T" & t_Distrib_Pts
            End If
        End If

        If ck_AutoFill.Value <> 0 Then
            t_TestHand = ""
            bln_Working = True
            CreateTestHand
            Do While bln_Working
                DoEvents
            Loop
        End If
    End If
End Sub


Private Function TotalDistribPts()
    If Left$(t_Distrib_Pts & " ", 1) = "T" And _
        t_Hand_HCPs <> "" _
    Then
'        str_Response = MyMsgBox("If you enter 'T' for total HCPs and" & vbCr & _
'                "Distribution points, then 'Std.HCPs' box" & vbCr & _
'                "should be left blank."), 2, "ok", "")
        TotalDistribPts = True
    End If

End Function


Private Sub t_Winners_GotFocus()
    str_OrigData = t_Winners
    L_Status.Clear
    L_Status.AddItem " .5=Kx,  1=A/KQ/QJT,  1.5=AQ  2=AK/AQJ/KQJ,QJT  2.5=AQJ  3=AKQ,KQJT  4=AKQJ.  <3, 2+"
    L_Status.AddItem " +1 for each over 4 in suit.  R#=Running winners: R2=AK,AA, R3=AKQ,AKA,AAA, R4=AKQJ"
    t_Winners.SelStart = 0
    t_Winners.SelLength = 32768
End Sub


Private Sub t_Winners_KeyDown( _
                KeyCode As Integer, _
                Shift As Integer)
    Select Case Shift
      Case Is = 0
        Select Case KeyCode
            Case Is = 38 'up
                t_CustPts(0).SetFocus
            Case Is = 40
                t_Losers.SetFocus
        End Select
    End Select
End Sub


Private Sub t_Winners_Validate(Cancel As Boolean)
    L_Status.Clear
    If str_OrigData <> t_Winners Then
        ck_Saved.Value = 0
    End If
End Sub


Private Sub Timer1_Timer()
    ' This timer is run only on start-up. '
    ' It is called from Form_Load and is  '
    ' used so that the form will display  '
    ' before the routines below run.      '

    Dim i As Integer
    Dim j As Integer

    Timer1.Interval = 0 ' Disables the timer. '

    ' Restore the grids to their state the last time '
    ' the program ended. The state is saved in the   '
    ' config file in Form_Unload.                    '

    If bln_StartUp Then
        bln_StartUp = False
        On Local Error GoTo 0
        str_StartGridRow = ""
        cmb_ConvName(0).ListIndex = 0
        cmb_ConvName(1).ListIndex = 0
        For i = 0 To 0 '1
            If mSplit(i).Checked Then
                Grid1(i).Splits(0).Size = int_GridSplitWidth1(i)
                Grid1(i).Splits(1).Size = int_GridSplitWidth2(i)
            End If
        Next

        If cmb_Sort(0).ListCount = 0 Then
            ' This was screwing things up. '
            For i = 0 To int_LastBidCol
                For j = 0 To 4
            '        cmb_Sort(j).AddItem Grid1(0).Columns(i).Caption
                Next

                For j = 0 To 1
            '        cmb_Show(j).AddItem Grid1(0).Columns(i).Caption
                Next
            Next

            For i = 0 To 4
                If i = 4 Then
                    cmb_Sort(i).ListIndex = 1
                Else
                    cmb_Sort(i).ListIndex = i + 2
                End If
            Next
        End If

        If str_CmdLineArg(1) <> "" Then
            f_EnterHand.Visible = True
            t_FindHand(int_FindHandNum) = str_CmdLineArg(1)
            t_FindPrior = str_CmdLineArg(2)
            b_FindBidForHand_Click
            str_CmdLineArg(1) = ""
            str_CmdLineArg(2) = ""
        Else
            ' Load the grids. '
            b_Show_Click
            bln_ShowGrid2 = True
            b_Show_Click
            DoEvents
            If bln_Abort Then Exit Sub
        End If

        'Me.MousePointer = vbHourglass

        i = 0
        If Not bln_Abort And _
           lng_TopID(i) <> 0 And Grid1(i).Row _
        Then
            Grid1(i).Row = 0
            If Grid1(i).Columns(int_LastBidCol).Text <> "" & lng_TopID(i) Then
                VScroll1.Value = 0
                Do While Grid1(i).Columns(int_LastBidCol).Text _
                            <> "" & lng_TopID(i) And _
                        VScroll1.Value < 100
                    VScroll1.Value = VScroll1.Value + 1
                    DoEvents
                    Grid1(i).Row = 0
                    DoEvents
                Loop
            End If
        Else
            bln_Abort = False
        End If

        If Not bln_TwoGrids And Grid1(1).Visible Then
            mDupesGrid_Click 0
            mDupesGrid_Click 1
        End If

    ElseIf mDupesGrid(0).Checked Then
        With Grid1(0)
            str_Job = ""
            On Local Error Resume Next
            FillDupeBoxes .Columns(int_Prior_Bids_Col), _
                          .Columns(int_Round_Col), _
                          .Columns(int_Player_Col), _
                          .Columns(int_OrdNum_Col)
        End With
        Exit Sub
    End If

    If bln_FirstTime Then
        ' The first time a user runs this program,  '
        ' This code takes him to the tutorial, then '
        ' a flag is saved in the config file so     '
        ' that the tutorial doesn't automatically   '
        ' execute next time.                        '
        Me.MousePointer = vbDefault
        Welcome
        bln_FirstTime = False
        Exit Sub
    End If

    Me.MousePointer = vbDefault
End Sub


Private Sub Timer2_Timer()
    Dim t As String
    Dim i As Integer

    't = Time
    'i = InStr(4, t$, ":") - 1
    'L_Clock = Left$(t$, i)
    'CheckBattery
End Sub


Private Sub Timer3_Timer()
    ' Sometimes grids won't update. '
    ' I tried everything else, but  '
    ' using a timer is all that works.
    DoEvents
    'b_Show_Click
    Timer3.Enabled = False
    
    ' When an entry is saved, make sure it is visible in the grid: '
        If lng_ID > 0 And str_Job = "Saving Entry" Then
            Data1(int_GridNum).Refresh
            DoEvents
            Grid1(int_GridNum).Refresh
            DoEvents
            Grid1(int_GridNum).Row = 1
            Do While lng_ID <> Grid1(int_GridNum).Columns(int_LastBidCol).Text And _
                     Grid1(int_GridNum).Columns(int_LastBidCol).Text <> ""
                If Grid1(int_GridNum).VisibleRows = 1 Then Exit Do
                If Grid1(int_GridNum).VisibleRows = Grid1(int_GridNum).Row + 1 Then Exit Do
                Grid1(int_GridNum).Row = Grid1(1).Row + 1
                DoEvents
            Loop
        End If
End Sub


Private Sub VScroll1_Change()
    Dim chng As Long

    If bln_Scrolling Then Exit Sub
    bln_Scrolling = True
    chng = VScroll1.Value - int_ScrollVal
    Grid1(0).Scroll 0, chng
    Grid1(1).Scroll 0, chng
    If Grid1(3).Visible Then
        Grid1(3).Scroll 0, chng
        Grid1(4).Scroll 0, chng
    End If
    int_ScrollVal = VScroll1.Value
    bln_Scrolling = False
End Sub


Private Sub RememberGridRows(g As Integer, action As String)
    If Grid1(g).Row = -1 Then Exit Sub
    'If Grid1(g).Row < 0 Then
    '    Grid1(g).Row = 0
    'End If
    On Error Resume Next
    int_CurRow(g) = Grid1(g).Row
    Grid1(g).Row = 0
    lng_TopID(g) = Val(Grid1(g).Columns(int_LastBidCol))
    Grid1(g).Row = int_CurRow(g)
    lng_CurID(g) = Val(Grid1(g).Columns(int_LastBidCol))

    If action = "Deleting" And _
       g = int_GridNum And _
       int_CurRow(g) > 0 _
    Then
        ' Top line is being deleted, so scroll down a line:
        Grid1(g).Row = Grid1(g).Row + 1
        int_CurRow(g) = Grid1(g).Columns(int_LastBidCol)
        lng_TopID(g) = Grid1(g).Columns(int_LastBidCol)
    End If
    On Error GoTo 0
End Sub


Private Sub RestoreGridRows(g As Integer)
    Dim lastID As Long

    If lng_CurID(g) <> 0 Then
        Data1(g).Refresh
        Me.MousePointer = 11
        Do While lng_TopID(g) <> Grid1(g).Columns(int_LastBidCol) And _
                 lng_ID <> Grid1(g).Columns(int_LastBidCol) And _
                 Grid1(g).Columns(int_LastBidCol) <> 0 And _
                 Grid1(g).Columns(int_LastBidCol) <> lastID And lastID <> 0

            lastID = Val(Grid1(g).Columns(int_LastBidCol))
            Grid1(g).Scroll 0, 1
            Grid1(g).Row = 0
            DoEvents
        Loop

        Do While int_CurRow(g) > Grid1(g).Row And _
                 lng_ID <> Val(Grid1(g).Columns(int_LastBidCol)) And _
                 Grid1(g).Columns(int_Bid_Col) <> ""

            Grid1(g).Row = Grid1(g).Row + 1
            DoEvents
        Loop
        Me.MousePointer = 0
    End If
End Sub


Private Sub CreateTestHand()
    If t_TestHand <> "" Then Exit Sub
    ' If the Test Hand field is blank, create a test hand from the specs '
    ' entered by the user. This is very complex - the longest set of     '
    ' routines in the program - all to save the user from having to make '
    ' up a test hand, but I think it is worth it in time saved for the   '
    ' user because making up valid test hands to match the entry's specs '
    ' is very complex for the user as well. The program also uses this   '
    ' routine for other purposes as well, such as duping entries and     '
    ' comparing entries.                                                 '

    Dim t(4) As String ' = cards for each suit. Ex: AK432
    Dim p(4) As Integer ' = points for each suit. Ex: 4
    Dim qmin(4) As Integer ' = minimum number of cards for each suit
    Dim qmax(4) As Integer ' = maximum number of cards for each suit
    Dim pMin(4) As Integer ' = minimum points for each suit
    Dim pmax(4) As Integer ' = maximum points for each suit
    Dim qSL(4) As String ' = suit limiting other suits, such as <S, >=H
    Dim qRange(4) As Integer ' = range of cards for each suit
    Dim pRange(4) As Integer ' = range of points for each suit
    Dim tq(4) As String  ' = qty. specs for each suit
    Dim tp(4) As String ' = HCP specs for each suit
    Dim honCount(4) As Integer
    Dim honReq(4) As Integer
    Dim cpAdj As Integer ' Custom Points adjustment

    GoSub CheckCombo
    GoSub CheckXNum  'X suit Num field
    GoSub CheckShape 'Shape field
    GoSub CheckNums  'Num fields
    GoSub CheckDistrib 'Distrib.Points box
    GoSub FillOutHand
    GoSub CheckXHCP  'X suit HCP field
    GoSub CheckPts   'Points fields
    GoSub FillOutPoints
    t_TestHand = t(1) & "-" & t(2) & "-" & t(3) & "-" & t(4)

    If cpAdj > 0 Then
        MsgBox$ "Unable to adjust for Custom Points" & vbCr & _
                "due to conflict with suit specs."
    End If

    bln_Working = False
Exit Sub


CheckCombo:
''''''''''
    Dim nm As String
    Dim suit1n As Integer
    Dim suit2n As Integer
    Dim suit3n As Integer
    Dim ttln As Integer ' total # of cards
    Dim calcn As String
    Dim pt As String
    Dim suit1p As Integer
    Dim suit2p As Integer
    Dim suit3p As Integer
    Dim ttlp As Integer ' total points

    'Combined suits specs:
    nm = Trim$(t_Qty(5))
    If Len(nm) > 4 Then
        'Combining suits: (e.g.: SC>8, CHD=11-13, DC<4)
        suit1n = InStr("SHDC", Mid$(nm, 1, 1))
        If suit1n > 0 Then
            suit2n = InStr("SHDC", Mid$(nm, 2, 1))
            suit3n = InStr("SHDC", Mid$(nm, 3, 1))

            If suit3n > 0 Then 'CDH=10-12, DHS>9, etc.
                calcn = Mid$(nm, 4, 1)
                ttln = Val(Mid$(nm, 5))

            Else 'CD=8-9, SH>8, etc.
                calcn = Mid$(nm, 3, 1)
                ttln = Val(Mid$(nm, 4))
            End If
        End If
    End If

    Dim calcp As String

    pt = Trim$(t_HCP(5))
    If Len(pt) > 4 Then
        'Combining suits: (e.g.: SC>8, CHD=11-13, DC<4)
        suit1p = InStr("SHDC", Mid$(pt, 1, 1))
        If suit1p > 0 Then
            suit2p = InStr("SHDC", Mid$(pt, 2, 1))
            suit3p = InStr("SHDC", Mid$(pt, 3, 1))

            If suit3p > 0 Then 'CDH=10-12, DHS>9, etc.
                calcp = Mid$(pt, 4, 1)
                ttlp = Val(Mid$(pt, 5))

            Else 'CD=8-9, SH>8, etc.
                calcp = Mid$(pt, 3, 1)
                ttlp = Val(Mid$(pt, 4))
            End If
        End If
    End If
Return


CheckXNum:
'''''''''
    Dim xdoneqty As Boolean
    Dim i As Integer

    xdoneqty = False
    For i = 1 To 4
        'quantity:
        t_Qty(i) = Trim$(t_Qty(i))
        If LCase$(Left$(t_Qty(i), 1)) = "x" Then
            If xdoneqty = False Then
                tq(i) = Trim$(t_Qty(0))
                xdoneqty = True

            ElseIf Len(t_Qty(i)) > 2 Then
                tq(i) = Trim$(Mid$(t_Qty(i), 3))

            Else
                tq(i) = ""
            End If
        Else
            tq(i) = t_Qty(i)
        End If
    Next
Return


CheckShape:
''''''''''
    Dim shapeDone As Integer

    If cmb_Shape.Text <> "" Then
        ' 1-suiter: shape$(1) = 6                                           '
        '           if t_Num(1)="" then shape$(1) = 6, etc.                 '
        ' 2-suiter: shape$(1) = 5, shape$(2) = 5                            '
        ' 3-suiter: shape$(1) = 4, shape$(2) = 4, shape$(3) = 4             '
        ' bal:  shape$(1) = 3, shape$(2) = 3, shape$(3) = 3, shape$(4) = 4  '
        ' unbal: shape$(1) = 2, shape$(2) = 2, shape$(3) = 4, shape$(4) = 5 '
        ' singleton: Shape$(1) = 1                                          '
        ' void: Shape$(1) = 0                                               '
        ' e: Shape$(1) = 1                                                  '

        Dim longSuitCount As Integer
        If cmb_Shape = "2-suiter" Then
            For i = 1 To 4
                If Val(tq(i)) > 4 Then
                    longSuitCount = Val(tq(i))
                    Exit For
                End If
            Next
        End If

        For i = 1 To 4
            If tq(i) = "" Or _
               cmb_Shape = "balanced" _
            Then
                If cmb_Shape = "3-suiter" And _
                   shapeDone < 3 _
                Then
                    qmin(i) = 4
                    qmax(i) = 5
                    shapeDone = shapeDone + 1

                ElseIf cmb_Shape = "2-suiter" And _
                       shapeDone < 2 _
                Then
                    If longSuitCount = 0 Then
                        qmin(i) = 5
                    Else
                        qmin(i) = 10 - longSuitCount
                    End If
                    qmax(i) = qmin(i)
                    shapeDone = shapeDone + 1

                ElseIf cmb_Shape = "1-suiter" And _
                       shapeDone = 0 _
                Then
                    qmin(i) = 6
                    qmax(i) = 9
                    shapeDone = True

                ElseIf cmb_Shape = "balanced" Then
                    If Left$(tq(i), 2) = "<=" Or _
                       Left$(tq(i), 2) = "=<" _
                    Then
                        qmin(i) = 2
                        If Val(Mid$(tq(i), 3)) = 0 Then
                            ' Must be a spec like "<=S"
                            qmax(i) = 4
                            qSL(i) = Mid$(tq(i), 3, 1)
                        Else
                            qmax(i) = Val(Mid$(tq(i), 3))
                        End If

                    ElseIf Left$(tq(i), 2) = ">=" Or _
                           Left$(tq(i), 2) = "=>" _
                    Then
                        qmax(i) = 5
                        If Val(Mid$(tq(i), 3)) = 0 Then
                            ' Must be a spec like "<=S"
                            qmin(i) = 2
                            qSL(i) = Mid$(tq(i), 3, 1)
                        Else
                            qmin(i) = Val(Mid$(tq(i), 3))
                        End If

                    ElseIf Left$(tq(i), 1) = "<" Then
                        qmin(i) = 1
                        ' For bidding NT when <3 card support '
                        ' for partner's suit.                 '

                        If Val(Mid$(tq(i), 2)) = 0 Then
                            qmax(i) = 4
                            qSL(i) = Mid$(tq(i), 2, 1)
                        Else
                            qmax(i) = Val(Mid$(tq(i), 2)) - 1
                        End If

                    ElseIf Left$(tq(i), 1) = ">" Then
                        qmax(i) = 5
                        If Val(Mid$(tq(i), 3)) = 0 Then
                            qmin(i) = 2
                            qSL(i) = Mid$(tq(i), 2, 1)
                        Else
                            qmin(i) = Val(Mid$(tq(i), 2))
                        End If

                    ElseIf Mid$(tq(i), 2, 1) = "+" Then
                        qmin(i) = Val(tq(i))
                        qmax(i) = 5

                    ElseIf tq(i) = "" Then
                        qmax(i) = 5
                        qmin(i) = 3
                    End If

                ElseIf cmb_Shape = "unbal." Then
                    Dim ii As Integer

                    For ii = 1 To 4 'taking over the Shape routine
                        If tq(ii) = "" Then
                            qmax(ii) = 1
                            qmin(ii) = 1
                            Exit For
                        End If
                    Next

                    If qmax(ii) = 0 Then 'no empty tq() field found
                        For ii = 1 To 4
                            If Left$(tq(ii), 1) = "<" Then
                                ' A low number is allowed. '
                                qmax(ii) = 1
                                qmin(ii) = 1
                                Exit For
                            End If
                        Next
                    End If
                    Exit For 'Note: If no fields with "" or with "<",
                             '      then "unbal" is not being set.

                ElseIf (LCase$(cmb_Shape) = "singleton" Or _
                        LCase$(cmb_Shape) = "sing/void") And _
                       Not shapeDone _
                Then
                    qmax(i) = 1
                    qmin(i) = 0
                    shapeDone = True

                ElseIf LCase$(cmb_Shape) = "void" And _
                       Not shapeDone _
                Then
                    tq(i) = "0"
                    qmax(i) = 0
                    qmin(i) = 0
                    shapeDone = True
                End If
            End If

            If tq(i) = "" Then
                tq(i) = qmin(i) & "-" & qmax(i)

            ElseIf Val(tq(i)) > 5 And _
                   cmb_Shape = "1-suiter" _
            Then
                qmin(i) = Val(tq(i))
                If Right$(tq(i), 1) = "+" Then
                    qmax(i) = qmin(i) + 1

                ElseIf Len(tq(i)) > 2 And _
                       Mid$(tq(i), 2, 1) = "-" _
                Then
                    qmax(i) = Val(Mid$(tq(i), 3))
                End If
                Exit For
            End If
        Next
    End If
Return


CheckDistrib:
''''''''''''
    Dim dp As String
    Dim d As Integer
    Dim ttlPts As Integer
    Dim minDistrib As Integer
    Dim maxDistrib As Integer
    Dim hp As String
    Dim j As Integer

    If t_Distrib_Pts = "" Then
        Return
    End If

    dp = Trim$(t_Distrib_Pts)
    d = Val(dp)
    ttlPts = UCase$((Left$(dp, 1))) = "T"

    If ttlPts Then
        If Len(dp) = 1 Then
            Return
        End If

            ' Arbitrarily allocate 2 of the Total points '
            ' to distribution and the rest to HCPs.      '
            hp = Mid$(dp, 2)
            d = 2
            dp = "2+"
            If Val(hp) = 0 And Len(hp) > 1 Then
                ' e.g.: Something like "<14" '
                ' though rarely if ever:     '
                i = Val(Mid$(hp, 2))
                If i > 0 Then
                    hp = Left$(hp, 1) & (i - 2)
                End If

            ElseIf Val(Right$(hp, 1)) = 0 And _
                   Right$(hp, 1) <> "0" _
            Then
                ' e.g.: 14+ '
                hp = (Val(hp) - 2) & Right$(hp, 1)

            ElseIf InStr(hp, "-") > 1 Then
                ' e.g.: 10-14 '
                i = Val(hp) - 2
                j = Abs(Val(Right$(hp, 2))) - 2
                hp = i & "-" & j

            Else ' Just something like T14, though rarely '
                hp = "" & (Val(hp) - 2)
            End If

        If t_Hand_HCPs = "" Then
            t_Hand_HCPs = hp
        End If

    Else
        hp = Val(t_Hand_HCPs)
    End If

    If d = 0 Then ' Must have "<" or ">"
        d = Val(Mid$(dp, 2))
        If Left$(dp, 1) = "<" Then
            minDistrib = 0
            maxDistrib = d - 1
        Else
            minDistrib = d + 1
            maxDistrib = 6
        End If

    ElseIf Right$(t_Distrib_Pts, 1) = "+" Then
        minDistrib = d
        maxDistrib = 6

    Else
        minDistrib = d
        maxDistrib = d
    End If

    ' 1=doubleton                         '
    ' 2=singleton or 2 doubletons         '
    ' 3=void or singleton + doubleton     '
    ' 4=doubleton + void or 2 singletons  '
    ' 5=singleton + void                  '
    ' 6=2 voids                           '
    ' Normally just use minDistrib        '
    i = 0
    Do While minDistrib >= 3 And i < 5 ' Do voids first
        For i = 1 To 4
            If qmax(i) = 0 Then
                qmax(i) = 0
                qmin(i) = 0
                tq(i) = "0"
                minDistrib = minDistrib - 3
                Exit For
            End If
        Next
        DoEvents
    Loop

    i = 0
    Do While minDistrib >= 2 And i < 5 ' Singletons
        For i = 1 To 4
            If qmax(i) >= 1 And _
               qmin(i) <= 1 _
            Then
                qmax(i) = 1
                qmin(i) = 1
                tq(i) = "1"
                minDistrib = minDistrib - 2
                Exit For
            End If
        Next
        DoEvents
    Loop

    i = 0
    Do While minDistrib >= 1 And i < 5 ' Doubletons
        For i = 1 To 4
            If qmax(i) >= 2 And _
               qmin(i) <= 2 _
            Then
                qmax(i) = 2
                qmin(i) = 2
                tq(i) = "2"
                minDistrib = minDistrib - 1
                Exit For
            End If
        Next
        DoEvents
    Loop

    i = 0
    Do While minDistrib > 0 And i < 5
        For i = 1 To 4
            If Left$(tq(i), 1) = "<" Then ' A low number is allowed
                If minDistrib >= 3 Then
                    qmax(i) = 0
                    qmin(i) = 0
                    minDistrib = minDistrib - 3

                ElseIf minDistrib = 2 Then
                    qmax(i) = 1
                    qmin(i) = 1
                    minDistrib = minDistrib - 2

                Else
                    qmax(i) = 2
                    qmin(i) = 2
                    minDistrib = minDistrib - 1
                End If

                tq(i) = "" & qmax(i)

                If minDistrib = 0 Then
                    Exit For
                End If
            End If
        Next
        DoEvents
    Loop
Return


CheckNums:
'''''''''
    For i = 1 To 4  'S, H, D, C
        Do While tq(i) = "0"
            qmin(i) = 0
            qmax(i) = 0
            i = i + 1
            If i = 5 Then
                Exit For
            End If
            DoEvents
        Loop

        ' Skip suits fixed up in routines above
        Do While qmax(i) > 0
            i = i + 1
            If i = 5 Then
                Exit For
            End If
            DoEvents
        Loop

        Dim n As Integer
        Dim n2 As Integer
        Dim n3 As Integer
        Dim s As Integer

        n = Val(tq(i))
        n2 = Val(Mid$(tq(i) & "  ", 2))
        n3 = Val(Mid$(tq(i) & "   ", 3))

        If tq(i) = "<x" Then
            qmin(i) = 0
            qmax(i) = Val(t_Qty(0))

        ' Examples: 4-, 5+, 2-5, <4, >C, >=D, <H, =S, x, 2-S '
        '          (from 2 to # of Spades, max)              '
        ElseIf tq(i) = "" Then
            qmin(i) = 0
            qmax(i) = 5 ' Max could be 13, but keeping it real

        ElseIf Right$(tq(i), 1) = "+" Then '5+
            qmin(i) = n
            qmax(i) = n + 3 ' Max could be 13, but keeping it real

        ElseIf Right$(tq(i), 1) = "-" Then '4-
            qmin(i) = 0
            qmax(i) = n

        ElseIf Mid$(tq(i) & "  ", 2, 1) = "-" Then '2-5
            qmin(i) = n
            qmax(i) = n3

        ElseIf Left$(tq(i), 2) = "<=" Or _
               Left$(tq(i), 2) = "=<" _
        Then '<=S, <=5
            qmin(i) = 0
            If n3 = 0 Then
                qSL(i) = Mid$(tq(i), 3, 1)
                s = InStr("SHDC", qSL(i))
                If s > 0 Then
                    qmax(i) = Val(tq(s))
                End If
            Else
                qmax(i) = n3
            End If

        ElseIf Left$(tq(i), 2) = ">=" Or _
               Left$(tq(i), 2) = "=>" _
        Then
            If n3 = 0 Then
                qSL(i) = Mid$(tq(i), 3, 1)
                s = InStr("SHDC", qSL(i))
                If s > 0 Then
                    qmax(i) = Val(tq(s))
                End If
            Else
                qmin(i) = n3
            End If
            qmax(i) = qmin(i) + 3 ' Max could be 13, but keeping it real

        ElseIf Left$(tq(i), 1) = "<" Then '<4, <H
            qmin(i) = 0
            If n2 = 0 Then ' Must be a suit, like <H
                qSL(i) = Mid$(tq(i), 2, 1)
                s = InStr("SHDC", qSL(i))
                If s > 0 And s < 5 Then
                    qmax(i) = Val(tq(s)) - 1
                End If
            Else
                qmax(i) = n2 - 1
            End If

        ElseIf Left$(tq(i), 1) = ">" Then '>5, >C
            If n2 = 0 Then ' Must be suit, like >S
                qSL(i) = Mid$(tq(i), 2, 1)
                s = InStr("SHDC", qSL(i))
                If s > 0 And s < 5 Then
                    qmin(i) = Val(tq(s)) + 1
                End If
            Else
                qmin(i) = n2 + 1
            End If
            qmax(i) = qmin(i) + 3 ' Max could be 13, but keeping it real

        ElseIf n > 0 Then ' e.g.: "3"
            qmin(i) = n
            qmax(i) = n
        End If
    Next

Return


FillOutHand:
'''''''''''
    Dim sTtl As Integer
    Dim totalmin As Integer
    Dim totalmax As Integer
    Dim adj As Single
    Dim zMin As Integer
    Dim tMax As Integer
    Dim a As Integer

    ' Make sure Combo specs agree with suit specs:
    If suit1n > 0 Then
        sTtl = qmin(suit1n) + qmin(suit2n) + qmin(suit3n)
        If calcn = ">" Then
            Do While sTtl <= ttln
                If qmax(suit1n) > qmin(suit1n) Then
                    qmin(suit1n) = qmin(suit1n) + 1
                    sTtl = sTtl + 1
                End If

                If sTtl <= ttln And _
                   qmax(suit2n) > qmin(suit2n) _
                Then
                    qmin(suit2n) = qmin(suit2n) + 1
                    sTtl = sTtl + 1
                End If

                If sTtl <= ttln And _
                   qmax(suit3n) > qmin(suit3n) And _
                   suit3n > 0 _
                Then
                    qmin(suit3n) = qmin(suit3n) + 1
                    sTtl = sTtl + 1
                End If
                DoEvents
            Loop
        End If
    End If

    For i = 1 To 4
        qRange(i) = qmax(i) - qmin(i)
        totalmin = totalmin + qmin(i)
        totalmax = totalmax + qmax(i)
    Next

    adj = 13 - totalmin

    If adj <> 0 Then
        zMin = 0
        tMax = 0
        For i = 1 To 4
            If qmin(i) = 0 And qmax(i) > 0 Then
                zMin = zMin + 1
                tMax = tMax + qmax(i)
            End If
        Next
    End If

    If zMin > 0 Then
        For i = 1 To 4
            If qmin(i) = 0 And tMax > 0 And qmax(i) > 0 Then
                If zMin = 1 Then
                    a = adj
                Else
                    a = CInt(adj / (tMax / qmax(i)))
                End If

                If a > qRange(i) Then
                    a = qRange(i)
                End If

                If a <= adj Then
                    qmin(i) = a
                Else
                    qmin(i) = adj
                End If

                adj = adj - qmin(i)
                totalmin = totalmin + qmin(i)
                tMax = tMax - qmax(i)
                zMin = zMin - 1
            End If
        Next
    End If

    If adj > 0 Then
        For i = 1 To 4
            If qRange(i) > 0 Then
                If i = 4 Then
                    n = CInt(adj)
                Else
                    n = CInt(adj / qRange(i))
                End If

                If qmin(i) + n > qmax(i) Then
                    n = qmax(i) - qmin(i)
                End If

                If adj - n < 0 Then
                    n = adj
                End If

                If n > 0 Then
                    qmin(i) = qmin(i) + n
                    totalmin = totalmin + n
                    qRange(i) = qRange(i) - n
                    adj = adj - n
                End If
            End If
            If adj = 0 Then Exit For
        Next
    End If

    ' Make sure the adjustment came out right: '
    adj = 13 - totalmin
    If adj <> 0 Then
        For i = 1 To 4
            Select Case qRange(i)
              Case Is > adj
                qmin(i) = qmin(i) + adj
                totalmin = totalmin + adj
                Exit For
              Case Is > 0
                n = adj - qRange(i)
                qmin(i) = qmin(i) + n
                totalmin = totalmin + n
                adj = adj - n
            End Select
        Next
    End If

    ' Check for suits limited by other suits: '
    For i = 1 To 4
        If qSL(i) <> "" Then
            s = InStr("SHDC", qSL(i))
            If s > 0 And s < 5 Then
                If Left$(t_Qty(i), 2) = "<=" Or _
                   Left$(t_Qty(i), 2) = "=<" _
                Then
                    Do While qmin(i) > qmin(s)
                        qmin(i) = qmin(i) - 1
                        qmin(s) = qmin(s) + 1
                        DoEvents
                    Loop

                ElseIf Left$(t_Qty(i), 1) = "<" Then
                    Do While qmin(i) >= qmin(s) And qmin(i) > 0
                        qmin(i) = qmin(i) - 1
                        qmin(s) = qmin(s) + 1
                        DoEvents
                    Loop

                ElseIf Left$(t_Qty(i), 2) = ">=" Or _
                       Left$(t_Qty(i), 2) = "=>" _
                Then
                    Do While qmin(i) < qmin(s)
                        qmin(i) = qmin(i) + 1
                        qmin(s) = qmin(s) - 1
                        DoEvents
                    Loop

                ElseIf Left$(t_Qty(i), 1) = ">" Then
                    Do While qmin(i) <= qmin(s)
                        qmin(i) = qmin(i) + 1
                        qmin(s) = qmin(s) - 1
                        DoEvents
                    Loop
                End If
            End If
        End If
    Next

    For i = 1 To 4
        If qmin(i) > 0 Then
            t(i) = Right$("AKQJT98765432", qmin(i))
        End If
    Next
Return


CheckXHCP:
'''''''''
    Dim xdonePts As Boolean

    xdonePts = False

    For i = 1 To 4
        ' Points:
        If LCase$(t_HCP(i)) = "x" Then
            If xdonePts = False Then
                tp(i) = Trim$(t_HCP(0))
                xdonePts = True
            Else
                tp(i) = ""
            End If
        Else
            tp(i) = Trim$(t_HCP(i))
        End If
    Next
Return


CheckPts:
''''''''
    ' HCP for suit:
    Dim v As Integer
    Dim tpt As Integer

    For i = 1 To 4
        Do While tp(i) = ""
            i = i + 1
            If i = 5 Then Exit For
            DoEvents
        Loop

        If Mid$(tp(i), 2) <> "stop" Then
            v = Val(tp(i))
        End If

        ' 3+, 2-5, <4, >2, <=4, =<4, >=2, =>2               '
        ' A=Ace, K=King, A,K=Ace or King                    '
        ' a=AK/KQ/AQ  b=AJT/KJT/QJT                         '
        '   c=KQJ/AQJ/AKQ/AKJ  d=AKQJ/KQJT/AKQT             '
        ' 1stop=A/Kx/QJx  2stop=AK/AQ/KQJ  0stop=no stopper '
        ' I=Intermediates (number of 8's and 9's and T's)   '
        ' FRC=First-Round Control  SRC=Second-RC.           '

        On Local Error Resume Next

        If InStr(tp(i), "=") > 0 Then '<=4, =<4, >=2, =>2
            v = Val(Mid$(tp(i), 3))
            If InStr(tp(i), "<") > 0 Then
                pMin(i) = 0
                pmax(i) = v

            ElseIf InStr(tp(i), ">") > 0 Then
                pMin(i) = v
                pmax(i) = 10

            Else
                pMin(i) = v
                pmax(i) = v
            End If

        ElseIf Left(tp(i), 1) = "<" Then
            v = Val(Mid$(tp(i), 2)) - 1
            pMin(i) = 0
            pmax(i) = v

        ElseIf Left(tp(i), 1) = ">" Then
            v = Val(Mid$(tp(i), 2)) + 1
            pMin(i) = v
            pmax(i) = 10

        ElseIf InStr(tp(i), "+") Then
            v = Val(tp(i))
            pMin(i) = v
            pmax(i) = 10
        End If


        If t_HCP(i) = "1stop" Then
            Mid$(t(i), 1) = "A"
            v = 4
            pmax(i) = 10
            pMin(i) = 4

        ElseIf t_HCP(i) = "2stop" Then
            Mid$(t(i), 1) = "AQ"
            v = 6
            pmax(i) = 10
            pMin(i) = 6

        ElseIf Left$(tp(i), 1) = "a" Then
            v = 5
            Mid$(t(i), 1) = "KQ"
            pmax(i) = 10
            pMin(i) = 5

        ElseIf Left$(tp(i), 1) = "b" Then
            v = 3
            Mid$(t(i), 1) = "QJT"
            pmax(i) = 10
            pMin(i) = 5

        ElseIf Left$(tp(i), 1) = "c" Then
            v = 6
            Mid$(t(i), 1) = "KQJ"
            pmax(i) = 10
            pMin(i) = 6

        ElseIf Left$(tp(i), 1) = "d" Then
            v = 6
            Mid$(t(i), 1) = "KQJT"
            pmax(i) = 10
            pMin(i) = 6

        ElseIf Left$(tp(i), 1) = "F" Or _
               tp(i) = "A" _
        Then 'FRC
            Mid$(t(i), 1) = "A"
            v = 4

        ElseIf Left$(tp(i), 1) = "S" Or _
               tp(i) = "K" Or _
               tp(i) = "A,K" _
        Then 'SRC
            Mid$(t(i), 1) = "K"
            v = 3

        Else
            GoSub Pts2Cards
        End If

        n = 1
        Do While Val(Mid$(t(i), n)) = 0 And _
                 n < Len(t(i))
            n = n + 1
            DoEvents
        Loop

        tpt = tpt + v
    Next
    On Local Error GoTo 0
Return


Pts2Cards:
'''''''''
    On Local Error Resume Next
    Select Case v
        Case Is > 10
            Mid$(t(i), 1) = "AKQJ"
            v = 10
        'Case 12
        '    Mid$(t(i), 1) = "KQJT"
        'Case 11
        '    Mid$(t(i), 1) = "AKQT"
        Case 10
            Mid$(t(i), 1) = "AKQJ"
        Case 9
            Mid$(t(i), 1) = "AKQ"
        Case 8
            Mid$(t(i), 1) = "AKJ"
        Case 7
            Mid$(t(i), 1) = "AK"
        Case 6
            Mid$(t(i), 1) = "AQ"
        Case 5
            Mid$(t(i), 1) = "KQ"
        Case 4
            Mid$(t(i), 1) = "A"
        Case 3
            Mid$(t(i), 1) = "K"
        Case 2
            Mid$(t(i), 1) = "Q"
        Case 1
            Mid$(t(i), 1) = "J"
    End Select
    On Local Error GoTo 0
Return


FillOutPoints:
'''''''''''''
    ' Make sure Combo specs agree with suit specs:

    Dim addPts As Integer
    Dim hcpBlank As Integer
    Dim ttlCards As Integer
    Dim adjTtl As Integer
    Dim adjNum As Integer
    Dim lasti As Integer
    Dim adjust As Single
    Dim nt As String
    Dim zz As String
    Dim jp As Integer

    If suit1p > 0 Then
        sTtl = pMin(suit1p) + pMin(suit2p) + pMin(suit3p)

        If calcp = ">" Then
            Do While sTtl <= ttlp
                If (pmax(suit1p) > pMin(suit1p) Or _
                    t_HCP(suit1p) = "") And _
                   (pMin(suit1p) <= pMin(suit2p) And _
                    (pMin(suit1p) <= pMin(suit3p) Or _
                     suit3p = 0)) _
                Then
                    pMin(suit1p) = pMin(suit1p) + 1
                    sTtl = sTtl + 1
                    tpt = tpt + 1
                End If

                If sTtl <= ttlp And _
                   (pmax(suit2p) > pMin(suit2p) Or _
                    t_HCP(suit2p) = "") And _
                   (pMin(suit2p) <= pMin(suit1p) And _
                    (pMin(suit2p) <= pMin(suit3p) Or _
                     suit3p = 0)) _
                Then
                    pMin(suit2p) = pMin(suit2p) + 1
                    sTtl = sTtl + 1
                    tpt = tpt + 1
                End If

                If sTtl <= ttlp And _
                   (pmax(suit3p) > pMin(suit3p) Or _
                    t_HCP(suit3p) = "") And _
                   suit3p > 0 And _
                   pMin(suit3p) <= pMin(suit1p) And _
                   pMin(suit3p) <= pMin(suit2p) _
                Then
                    pMin(suit3p) = pMin(suit3p) + 1
                    sTtl = sTtl + 1
                    tpt = tpt + 1
                End If
                DoEvents
            Loop
        End If
    End If

    For i = 1 To 4
        v = pMin(i)
        If v > 0 Then GoSub Pts2Cards
    Next


    Dim pts As Integer
    Dim minP As Integer
    Dim maxP As Integer
    Dim strHCP As String
    Dim r1 As String
    Dim r2 As String
    Dim sel As String
    Dim bNT As String
    Dim bNTn As String

    If Left$(t_Hand_HCPs & " ", 1) = "P" And _
       t_Player = "3" _
    Then
        If int_1NT_Min = 0 Then
            GetMinNT
        End If

        If Left$(t_Prior & "  ", 2) = "1N" Then
            ' Partner opened 1N. Compute responses from '
            ' minimum NT opening HCP range. Example: If '
            ' HCPs spec for this entry is "P24-25" then '
            ' if the opening range is 15-17, then the   '
            ' HCP spec for this entry is 24 minus 15 =  '
            ' 9 and 25 minus 15 = 10, or "9-10". If the '
            ' opening range is 12-14, then the range for'
            ' this entry is 24-12=12 and 25-12=13, or   '
            ' 12-13. This is a unique situation because '
            ' of 1N (and 2N) and its responses having   '
            ' such tightly defined HCP ranges.          '
            minP = int_1NT_Min
        Else
            minP = int_2NT_Min
        End If

        strHCP = Mid$(t_Hand_HCPs, 2) ' Get rid of "P"  '
        If Left$(strHCP, 1) = "<" Then
            strHCP = "<" & (Val(Mid$(strHCP, 2)) - minP)
        Else
            r1 = "" & (Val(strHCP) - minP)
            If Len(strHCP) > 3 Then ' e.g.: 24-25 '
                strHCP = r1 & "-" & _
                        (Val(Right$(strHCP, 2)) - _
                         minP)
            Else
                strHCP = r1 & Right$(strHCP, 1)
            End If
        End If

    ElseIf (Left$(LCase$(t_Hand_HCPs), 3) = "min" Or _
            Left$(LCase$(t_Hand_HCPs), 3) = "hcp") And _
           t_Player = "1" _
    Then
        If int_1NT_Min = 0 Then
            GetMinNT
        End If

        If Left$(t_Prior & "  ", 2) = "1N" Or _
           LCase$(t_Hand_HCPs) = "hcp1n" _
        Then
            minP = int_1NT_Min
            maxP = minP + 2
        Else
            minP = int_2NT_Min
            maxP = minP + 1
        End If

        If Left$(LCase$(t_Hand_HCPs), 3) = "hcp" Then
            strHCP = minP & "-" & maxP

        ElseIf LCase$(t_Hand_HCPs) = "min" Then
            ' Bidder opened NT, responder made an '
            ' invitational bid.                   '
            strHCP = "" & minP

        ElseIf LCase$(t_Hand_HCPs) = "min+" Then
            strHCP = (minP + 1) & "-" & (minP + 2)

        ElseIf Left$(LCase$(t_Hand_HCPs) & "   ", 4) = "min+" Then
            strHCP = "" & (minP + Val(Right$(t_Hand_HCPs, 1)))

        ElseIf Left$(LCase$(t_Hand_HCPs) & "   ", 4) = "min-" Then
            strHCP = "" & (minP - Val(Right$(t_Hand_HCPs, 1)))

        End If

    Else
        strHCP = t_Hand_HCPs
    End If

    pts = Val(strHCP) ' HCPs for the hand

    If pts = 0 And _
       Len(strHCP) > 1 _
    Then
        If Left$(strHCP, 1) = "<" Then
            pts = Val(Mid$(strHCP, 2)) * 0.5
        End If
    End If

    addPts = pts - tpt

    If addPts > 0 Then
        hcpBlank = 0
        For i = 1 To 4
            If Val(t(i)) <> 0 And _
               t(i) <> "" And _
               Len(t(i)) > 1 _
            Then 'if A32 or K32, etc., val=0
                hcpBlank = hcpBlank + 1
                ttlCards = ttlCards + Len(t(i))
            End If
        Next

        If hcpBlank > 0 Then
            For i = 1 To 4
                If Val(t(i)) <> 0 And _
                   t(i) <> "" _
                Then
                    If (hcpBlank = 1 Or i = 4) And _
                       addPts _
                    Then
                        v = addPts

                    Else
                        v = Fix(addPts / ttlCards * Len(t(i))) + 0.5
                    End If

                    If v > 0 Then
                        GoSub Pts2Cards
                        addPts = addPts - v
                    End If

                    hcpBlank = hcpBlank - 1
                    If hcpBlank = 0 Then Exit For
                End If
            Next
        End If

        ' Make sure hand has required HCPs:
        If addPts > 0 Then
            adjTtl = 0
            adjNum = 0

            For i = 1 To 4
                If pmax(i) > 0 And _
                   InStr(t_HCP(i), "<") = 0 _
                Then
                    adjTtl = adjTtl + pmax(i) - pMin(i)
                    adjNum = adjNum + 1
                    lasti = i
                End If
            Next

            For i = 1 To 4
                If pmax(i) > 0 And _
                   InStr(t_HCP(i), "<") = 0 _
                Then
                    If i < lasti Then
                        adjust = (pmax(i) - pMin(i)) / adjTtl
                        adj = adjust * addPts
                        v = pMin(i) + adj
                    Else
                        adj = addPts
                        v = pMin(i) + addPts
                    End If
                    addPts = addPts - adj
                    GoSub Pts2Cards
                ElseIf t_HCP(i) = "" Then
                    If InStr(t(i), "A") Then v = 4 Else v = 0
                    If InStr(t(i), "K") Then v = v + 3
                    If InStr(t(i), "Q") Then v = v + 2
                    If InStr(t(i), "J") Then v = v + 1
                    v = v + addPts
                    addPts = v
                    ' If v>10, it will be reduced to 10.
                    GoSub Pts2Cards
                    addPts = addPts - v
                    ' addPts should now be 0 unless v '
                    ' was >10 and was reduced to 10.  '
                End If
                If addPts = 0 Then Exit For
            Next
        End If


        ' Put suits together to make test hand: '
        nt = t(1) & "-" & t(2) & "-" & t(3) & "-" & t(4)

        If Len(nt) < 16 Then
            i = 1
            Do While Right$(t_Qty(i), 1) <> "+" And _
                     i < 4
                i = i + 1
            Loop

            zz = Right$("AKQJT98765432---", 16 - Len(nt) + Len(t(i)))

            j = 1
            Do While Val(Mid$(t(i), j)) = 0 And _
                     j < Len(t(i))
                j = j + 1
            Loop

            Mid$(zz, 1) = Left$(t(i), j - 1)
            t(i) = zz
        End If
    End If

    ' The section below adjusts the test hand for Custom Points.'
    ' For example, if the test hand has only 4 Q's (8 HCPs) and '
    ' 3 Custom Points are specified where A=2 and K=1 and Q=0,  '
    ' then we change the hand to have 1 A, 1 K, and 1 J. This   '
    ' has 8 HCPs as before, but now has 3 Custom Points. As you '
    ' can see, Custom Points can greatly improve the quality of '
    ' a hand.                                                   '
    '                                                           '
    ' Custom Points are too complex to be worth auto-filling in '
    ' general, but since A=2, K=1 is fairly common, the below   '
    ' code adjusts for that. The complexity of this one simple  '
    ' Custom Points spec shows why I didn't try to allow for all'
    ' possible custom point specs. Even as long and complex as  '
    ' this is, the first test I ran on it failed. The generated '
    ' hand was KQT5432-Q-A2-432 with HCP spec of 11 and Custom  '
    ' Points of 4 and Spades of 7-8 cards with "abi1" points.   '
    ' To meet this specs, the test hand would have to have been '
    ' changed to QJT8754-A2-3-A97. That's just a bit too complex'
    ' a change to program. However, most test hands are not as  '
    ' hard to adjust.                                           '

    Dim hasK As String
    Dim hasQJ As String
    Dim hasQ As String
    Dim hasJ As String
    Dim cpSpec As Integer
    Dim CP As Integer
    Dim sLen As Integer
    Dim s1 As Integer
    Dim s2 As Integer
    Dim s3 As Integer
    Dim numi As Integer
    Dim z As Integer
    Dim ni As Integer

    cpSpec = Val(t_CustPoints)
    adj = 0
    cpAdj = 0
    If cpSpec > 0 And _
       t_CustPts(0) = "2" And _
       t_CustPts(1) = "1" _
    Then
        ' Count aces and kings:
        CP = 0
        For i = 1 To 4
            If Left$(t(i) & "  ", 2) = "AK" Then
                CP = CP + 3
            ElseIf Left$(t(i) & " ", 1) = "A" Then
                CP = CP + 2
            ElseIf Left$(t(i) & " ", 1) = "K" Then
                CP = CP + 1
                hasK = hasK & i
            ElseIf Left$(t(i) & "  ", 2) = "QJ" Then
                hasQJ = hasQJ & i
            ElseIf Left$(t(i) & " ", 1) = "Q" Then
                hasQ = hasQ & i
            ElseIf Left$(t(i) & " ", 1) = "J" Then
                hasJ = hasJ & i
            End If
        Next
        ' The above only counts Q's and J's which are the
        ' highest card in the suit, so the J in AJ, for
        ' example, is not counted.
        adj = cpSpec - CP
        If adj > 0 Then
            Do While hasQJ <> "" And adj > 0
                ' This is the easiest adjustment to make
                ' Just substitute K2 (worth 1 cp) for QJ.
                s = Val(Left$(hasQJ, 1))
                sLen = Len(t(s))
                If sLen = 2 Then 'change QJ to K2:
                    t(s) = "K2"
                Else 'QJ32, QJ5432, QJT2, etc.
                    ' Intermediates have not been put in yet,
                    ' but QJT could have been specified.    '
                    If Mid$(t(s), 3, 1) = "T" Then
                        If sLen = 3 Then 'QJT
                            t(s) = "KT2"
                        Else 'QJT2, QJT32, etc.
                            t(s) = "KT" & Mid$(t(s), 4)
                        End If
                    Else 'QJ2, QJ32, QJ432, etc.
                        'By coincidence, the card number to take '
                        'the place of the J is always = 'slen'.  '
                        t(s) = "K" & sLen & Mid$(t(s), 3)
                    End If
                End If
                adj = adj - 1
                hasQJ = Trim$(Mid$(hasQJ & " ", 2))
                DoEvents
            Loop

            If adj > 0 Then
                ' More complicated. Must find at a minimum one suit     '
                ' with a Q and another with a J, then sub a K for the   '
                ' Q and an x value for the J, or sub an A for 2 Q's.    '
                ' If adj=1, can sub an A for a K and a Q or J.          '
                ' If the user was careless with his specs, making these '
                ' changes may screw up his suit specs.                  '

                Do While adj > 1 And _
                         (Len(hasQ) > 1 Or _
                          Len(hasJ) > 3 Or _
                          (Len(hasQ) = 1 And Len(hasJ) > 1))
                    ' Substitute an A for 2 Q's or 4 J's or 1 Q & 2 J's:
                    Do While adj > 1 And _
                             Len(hasQ) > 1
                        ' 2 suits with Q's, so substitute an A:
                        s1 = Val(Left$(hasQ, 1))
                        s2 = Val(Mid$(hasQ, 2, 1))
                        Mid$(t(s1), 1, 1) = "A"

                        'Q, Q2, Q32, Q432
                        t(s2) = (Len(t(s2)) + 1) & Trim$(Mid$(t(s2) & " ", 2))
                        adj = adj - 2
                        hasQ = Trim$(Mid$(hasQ & "  ", 3))
                        DoEvents
                    Loop

                    Do While adj > 1 And _
                             hasQ <> "" And _
                             Len(hasJ) > 1
                        ' 1 suit as Q, 2+ suits have J's
                        s1 = Val(Left$(hasQ, 1))
                        s2 = Val(Left$(hasJ, 1))
                        s3 = Val(Mid$(hasJ, 2, 1))
                        Mid$(t(s1), 1, 1) = "A"

                        'J, J2, J32, J432
                        t(s2) = (Len(t(s2)) + 1) & Trim$(Mid$(t(s2) & " ", 2))
                        t(s3) = (Len(t(s2)) + 1) & Trim$(Mid$(t(s3) & " ", 2))
                        adj = adj - 2
                        hasQ = Trim$(Mid$(hasQ & " ", 2))
                        hasJ = Trim$(Mid$(hasJ & "  ", 3))
                        DoEvents
                    Loop

                    If adj > 1 And _
                       Len(hasJ) = 4 _
                    Then
                        ' Sub an A for 4 J's:                   '
                        ' (Remember that this can only be where '
                        ' J is the high honor in the suit, such '
                        ' as J32-J32-J2-J5432 -- no very likely '
                        ' to have 4 HCPs specified plus cust    '
                        ' pts=2+ but here it is anyway:)        '

                        Mid$(t(1), 1, 1) = "A"
                        For i = 2 To 4
                            t(i) = (Len(t(i)) + 1) & Trim$(Mid$(t(i) & " ", 2))
                        Next
                        hasJ = ""
                        adj = adj - 2
                    End If
                    DoEvents
                Loop

                Do While adj > 0 And _
                         ((Len(hasQ) > 1 Or Len(hasK) > 1) Or _
                          (hasQ <> "" And hasJ <> "") Or _
                           Len(hasJ) = 3)
                    ' By now, adj should not be >1, which is a K. '
                    ' Sub a K and a J for 2 Q's, or a K for a Q   '
                    ' and J, or a K for 3 J's, or an A for K-Q/J. '

                    Do While adj > 0 And _
                             Len(hasQ) > 1
                        ' 2 suits with Q's, so substitute a K and a J:
                        s1 = Val(Left$(hasQ, 1))
                        s2 = Val(Mid$(hasQ, 2, 1))
                        Mid$(t(s1), 1, 1) = "K"

                        ' Q, Q2, Q32, Q432
                        Mid$(t(s2), 1, 1) = "J"
                        adj = adj - 1
                        hasQ = Trim$(Mid$(hasQ & "  ", 3))
                        hasJ = hasJ & s2
                        DoEvents
                    Loop

                    Do While adj > 0 And _
                             hasQ <> "" And _
                             hasJ <> ""
                        ' 1 suit has Q, 1+ suits have J's
                        s1 = Val(Left$(hasQ, 1))
                        s2 = Val(Left$(hasJ, 1))
                        Mid$(t(s1), 1, 1) = "K"

                        ' J, J2, J32, J432
                        t(s2) = (Len(t(s2)) + 1) & Trim$(Mid$(t(s2) & " ", 2))
                        adj = adj - 1
                        hasQ = Trim$(Mid$(hasQ & " ", 2))
                        hasJ = Trim$(Mid$(hasJ & " ", 2))
                        DoEvents
                    Loop

                    If adj > 0 And Len(hasJ) > 2 Then
                        ' Sub a K for 3 J's:
                        s1 = Val(Left$(hasJ, 1))
                        s2 = Val(Mid$(hasJ, 2, 1))
                        s3 = Val(Mid$(hasJ, 3, 1))
                        Mid$(t(s1), 1, 1) = "K"
                        t(s2) = (Len(t(s2)) + 1) & Trim$(Mid$(t(s2) & " ", 2))
                        t(s3) = (Len(t(s3)) + 1) & Trim$(Mid$(t(s3) & " ", 2))
                        hasJ = Trim$(Mid$(hasJ & " ", 4))
                        adj = adj - 1
                    End If

                    If adj > 0 And hasK <> "" And _
                      (hasQ <> "" Or hasJ <> "") _
                    Then
                        ' Sub an A for K and Q or J: '
                        s1 = Val(Left$(hasK, 1))
                        If t_HCP(s1) <> "" And Len(hasK) > 1 Then
                            ' Look for a suit with no HCP spec: '
                            i = Val(Mid$(hasK, 2, 1))
                            If t_HCP(i) = "" Then
                                s1 = i
                                hasK = Trim$(Left$(hasK, 1) & _
                                       Mid$(hasK & " ", 3))
                            Else
                                hasK = Trim$(Mid$(hasK & " ", 2))
                            End If
                        End If
                        Mid$(t(s1), 1, 1) = "A"

                        If hasJ <> "" Then
                            s2 = Val(Left$(hasJ, 1))
                            t(s1) = Right$("T98765432", Len(t(s1)))
                            hasJ = Trim$(Mid$(hasJ & " ", 2))
                        ElseIf hasQ <> "" Then
                            s2 = Val(Left$(hasQ, 1))
                            Mid$(t(s2), 1, 1) = "J"
                            hasJ = hasJ & s2
                            hasQ = Trim$(Mid$(hasQ & " ", 2))
                        End If

                        adj = adj - 1
                    End If
                    DoEvents
                Loop
            End If
        End If
    End If
    cpAdj = adj

    ' Insert Intermediates:                       '
    ' (Both the # specified for hand and by suit) '
    numi = Val(t_Intermediates)
    Dim lt As Integer
    Do While numi > 0
        For i = 1 To 4
            lt = Len(t(i))
            If lt > 2 Or i = 4 Then
                ii = InStr(tp(i), "i")
                If ii = 0 Then
                    tp(i) = tp(i) & "i1"
                    ii = InStr(tp(i), "i")
                End If

                If ii > 0 Then
                    z = 1
                    ' Find first non-honor in hand, '
                    ' such as "KT42" = char.#3
                    Do While Val(Mid$(t(i), z, 1)) = 0 And _
                             z < Len(t(i))
                        z = z + 1
                    Loop
                    ni = Val(Mid$(tp(i), ii + 1, 1))

                    If Val(Right$(t(i), 1)) > 0 And z > 1 Then
                        If Mid$(t(i), z - 1, 1) = "T" Then
                            Mid$(t(i), z, ni) = Left$("987", ni)
                        Else
                            Mid$(t(i), z, ni) = Left$("T98", ni)
                        End If
                        numi = numi - 1
                    End If
                End If
                If numi = 0 Then Exit Do
            End If
        Next
        DoEvents
    Loop

    If hp <> "" Then
        ' t_Hand_HCPs was temporarily filled in based '
        ' on the value entered in the Distribution  '
        ' Points field, such as "T11+". Now we need '
        ' to blank it out again.                    '
        t_Hand_HCPs = ""
    End If

Return

End Sub


Private Sub CreateDistribDisclosure()
    ' Create a "Distribution Disclosure" based '
    ' on the entry's specs:                    '

    Dim t As String
    Dim q As Integer
    Dim i As Integer

    'If t_Distrib = "" Then
    '    For i = 1 To 4
    ''        If t_Qty(i) <> "x" And _
    '           t_Qty(i) <> "" _
    '        Then
    '            t = t + Mid$("SHDC", i, 1) + t_Qty(i) & ","
    '
    '        ElseIf t_Qty(i) = "x" And _
    '               q = 0 And _
    '               InStr(t, "x") = 0 _
    '        Then
    '            t = t & "x" & t_Qty(0) & ","
    '            q = Val(t_Qty(0))
    '        End If
    '    Next
   '
   '     If t <> "" Then
   '         t_Distrib = Left$(t, Len(t) - 1) ' get rid of last comma
   '     End If
   ' End If
End Sub


'Private Sub CheckBattery()

    'If SysInfo.ACStatus = 1 Then 'AC power
    '    lblBatt.Visible = False
    '    lblBattStat.Visible = False
    '    lblBattTime.Visible = False
    '    lblTime.Visible = False
    '    Timer2.Enabled = False
    '    Exit Sub
    'End If

    'lblBatt.Visible = True
    'lblBattStat.Visible = True
    'lblBattTime.Visible = True
    'lblTime.Visible = True

    'Dim PerCentLeft As String

    'If SysInfo.BatteryLifePercent <> 255 Then
    '    PerCentLeft = SysInfo.BatteryLifePercent
    '    If PerCentLeft > 49 Then
    '        lblBatt.ForeColor = vbGreen
    '    Else
    '        lblBatt.ForeColor = vbRed
    '    End If
    '    lblBatt.Caption = PerCentLeft & "%"

    'Else
    '    lblBatt.Caption = "Unknown"
    'End If

    'If SysInfo.BatteryLifeTime <> &HFFFFFFFF Then
    '    lblTime.Caption = Format((TimeSerial(0, 0, _
    '                        SysInfo.BatteryLifeTime)), "h:mm")
    'Else
    '    lblTime.Caption = "Unknown"
    'End If
'End Sub


Private Function PB_To_SortOrd_Code(ByVal pb As String) As String
    Dim v As String
    Dim v1 As String
    Dim i As Integer
    Dim n As Integer
    Dim a As String

    ' See a discussion of this code in        '
    ' CardShark_BidBase_Programming_Notes.htm '
    ' Note that this same Function is also in '
    ' the editor code with the same name, so  '
    ' any changes in one should be copied to  '
    ' the other. '
Exit Function

    If Val(pb) = 0 And pb <> "" Then
        PB_To_SortOrd_Code = LCase$(pb)
        Exit Function
    End If

On Local Error GoTo SortOrdCodeErr
    v = ""
    pb = UCase$(Trim$(pb))

    If pb = "" Then
        PB_To_SortOrd_Code = Space$(34) & "0"
        Exit Function
    End If

    If InStr(pb, "@") > 0 Then
        ' Slam bidding entries. '
        ' See the Notes file.   '
        PB_To_SortOrd_Code = pb
        Exit Function
    End If

    Do While pb <> ""
        i = InStr(pb, "-")
        If i > 0 Then
            ' Extract next bid from BidSequence:
            v1 = Left$(pb, i - 1)
            pb = Mid$(pb, i + 1)

        Else
            v1 = pb
            pb = ""
        End If

        v1 = Replace(v1, Chr$(13), "")
        v1 = Replace(v1, Chr$(10), "")
        If Len(v1) = 1 Then
            v = v & (InStr(" DRP", v1) - 1)

        ElseIf v1 <> "" Then
            'If Len(v1) > 2 Then ' Ex: 4N@30 '
            '    a = Mid$(v1, 3) '       @30 '
            '    v1 = Left$(v1, 2) '   4N    '
            'End If

            ' Find number of the bid in the list of all bids:
            n = InStr(" 1C1D1H1S1N2C2D2H2S2N3C3D3H3S3N" & _
                      "4C4D4H4S4N5C5D5H5S5N6C6D6H6S6N" & _
                      "7C7D7H7S7N1M2M3M4M5M6M7M1m2m3m4m5m6m7m", v1) / 2
            If n = 0 Then 'n > 36 Or n = 0 Then
                str_Response = MyMsgBox("Invalid prior bid.", 1, "ok", "")
                Exit Function
            End If
            ' Convert decimal value of bid # to code character:
            v = v & Mid$("456789ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]!@#$%^&*():;<>,.", n, 1)
        End If
        DoEvents
    Loop
    ' For BidBase purposes, I need the codes in alphanumeric '
    ' order, so I have to pad them on the left with spaces.  '
    ' For example, "413A" is less than "46" when comparing   '
    ' character strings, so I have to convert "46" to "  46" '
    ' so that "  46" will compare as being less than "413A": '
    PB_To_SortOrd_Code = Right$(Space$(35) + v, 35)

SortOrdCodeErr:
    If Err.Number > 0 Then
        str_Response = MyMsgBox(Err.Description & vbCr & _
                                " creating Sort Order Code.", 2, "ok", "")
        Resume Next
    End If
    On Local Error GoTo 0
End Function


Private Sub Copy2DeletedTable()
    Set rs_DelTable = db_BidBase.OpenRecordset("Garbage")
    rs_DelTable.AddNew

    With rs_BidTable
    '              '
    rs_DelTable!Pct_Used = "" & !Pct_Used
    rs_DelTable!RoundOfBidding = !RoundOfBidding
    rs_DelTable!player = "" & !player
    rs_DelTable!Bid_Order_Number = "" & !Bid_Order_Number
    rs_DelTable!Prior_Bids = "" & !Prior_Bids
    rs_DelTable!bid = "" & !bid
    rs_DelTable!ConvID = !ConvID
    rs_DelTable!hcp = "" & !hcp
    rs_DelTable!Dist_Pts = "" & !Dist_Pts
    rs_DelTable!Cust_Pts = "" & !Cust_Pts
    rs_DelTable!A_Cust = "" & !A_Cust
    rs_DelTable!K_Cust = "" & !K_Cust
    rs_DelTable!Q_Cust = "" & !Q_Cust
    rs_DelTable!J_Cust = "" & !J_Cust
    rs_DelTable!T_Cust = "" & !T_Cust
    rs_DelTable!Winrs = "" & !Winrs
    rs_DelTable!losers = "" & !losers
    rs_DelTable!QTs = "" & !QTs
    rs_DelTable!Stopped_Suits = "" & !Stopped_Suits
    rs_DelTable!Shape = "" & !Shape
    rs_DelTable!intermed = "" & !intermed
    rs_DelTable!xNum = "" & !xNum
    rs_DelTable!xHCP = "" & !xHCP
    rs_DelTable!sNum = "" & !sNum
    rs_DelTable!sHCP = "" & !sHCP
    rs_DelTable!hNum = "" & !hNum
    rs_DelTable!HHCP = "" & !HHCP
    rs_DelTable!dNum = "" & !dNum
    rs_DelTable!dHCP = "" & !dHCP
    rs_DelTable!cNum = "" & !cNum
    rs_DelTable!cHCP = "" & !cHCP
    rs_DelTable!cmNum = "" & !cmNum
    rs_DelTable!cmHCP = "" & !cmHCP
    rs_DelTable!Vul = "" & !Vul
    rs_DelTable!pos = "" & !pos
    rs_DelTable!Alert = "" & !Alert
    rs_DelTable!Artificial = "" & !Artificial
    rs_DelTable!Explanation = "" & !Explanation
    rs_DelTable!Test_Hand = "" & !Test_Hand
    rs_DelTable!source = "" & !source
    rs_DelTable!Asks = "" & !Asks
    rs_DelTable!DateOfEntry = !DateOfEntry
    rs_DelTable!Best_Contract = "" & !Best_Contract
    rs_DelTable!LHO = !LHO
    rs_DelTable!Pard = !Pard
    rs_DelTable!RHO = !RHO
    rs_DelTable!BidsID = !id
    rs_DelTable.Update
    rs_DelTable.Bookmark = rs_DelTable.LastModified
    '      '
    End With
    Data1(2).Refresh
End Sub


Private Sub Wait()
    Dim xz As Integer

    If bln_Waiting Then
        Exit Sub
    End If
    bln_Waiting = True
    bln_GoBack = False

    Do
        xz = DoEvents()
    Loop While bln_Waiting

End Sub


Sub ResizeTutorial()
    frm_Tutorial.Height = L_Tutorial.Height + 900
    b_Continue.Top = L_Tutorial.Height + 400
    b_Go_Back.Top = b_Continue.Top
    b_TutorDetails.Top = b_Continue.Top
    b_QuitTutorial.Top = b_Continue.Top
    b_QuitTutorial.Left = L_Tutorial.Left + L_Tutorial.Width - b_QuitTutorial.Width
    b_TutorDetails.Left = b_QuitTutorial.Left - 1700
    b_Continue.Visible = True
    pTutorDragBar.Top = b_Continue.Top + b_Continue.Height
    pTutorDragBar.Left = 1
    pTutorDragBar.Width = frm_Tutorial.Width
    pTutorDragBar.Height = 100
    DoEvents
End Sub


Private Function LegalFileName(ByVal f) As String
    Dim i As Integer

    For i = 1 To Len(f)
        If InStr("/\.?^*<>|", Mid$(f, i, 1)) > 0 Then
            Mid$(f, i, 1) = "_"
        End If
    Next

    LegalFileName = f

End Function


Private Sub ShowGetBidError(b As String, id1 As Long, id2 As Long)
    ' Compare 2 entries with the same bid and specs. '
    ' b = the bid. '
    Dim n As Integer
    Dim i As Long
    'Dim id1 As Long
    'Dim id2 As Long
    Dim v As String
    Dim num As Long
    Dim sel As String
    Dim sel2 As String
    Dim nam As String
    Dim pos As String
    Dim ec As Long

    If ec > 0 Then
        ' Got an error - abort the search and display '
        ' the current entry which was being tested.   '
        If ec = 88 Then
            n = InStr(str_ErrMsg, "#")
            If n > 0 Then
                id1 = Val(Mid$(str_ErrMsg, n + 1))
            End If
            n = InStr(n + 3, str_ErrMsg, "#")
            If n > 0 Then
                id2 = Val(Mid$(str_ErrMsg, n + 1))
            End If

            ' Conflicting entries. Show comparison. '
            Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
            Set rs_BidTable = db_BidBase.OpenRecordset(str_BidsTbl)
            With rs_BidTable
            ''''''''''''''''
            .Index = "PrimaryKey"
            .Seek "=", id2
            ' Get the 2nd entry. The first one is the current record in the grid. '
            If .NoMatch Then
                Exit Sub ' Shouldn't be possible. '
            End If
            ' g_CompareGrid is a MSFlexGrid control. '
            f_CompareGrid.Left = 8000
            f_CompareGrid.ZOrder 0
            g_CompareGrid.ColWidth(0) = 1295
            g_CompareGrid.ColWidth(1) = 1740
            g_CompareGrid.ColWidth(2) = 1740
            If Val("" & !ConvID) > 0 Then
                LookUpBidName !ConvID, sel, sel2, nam
            End If
            FillCompareGrid 0, "", "#" & id1, "#" & id2
            g_CompareGrid.ColAlignment(1) = 3
            g_CompareGrid.ColAlignment(2) = 3
            FillCompareGrid 1, "Convention", Grid1(0).Columns(int_Bid_Col + 1).Text, nam
            FillCompareGrid 2, "Cnv-Sub-Pct", _
                            Grid1(0).Columns(0).Text & "-" & _
                            Grid1(0).Columns(1).Text & "-" & _
                            Grid1(0).Columns(2).Text, _
                            sel & "-" & sel2 & "-" & !Pct_Used
            FillCompareGrid 3, "Bid Order #", Grid1(0).Columns(int_Bid_Col - 1).Text, !Bid_Order_Number
            FillCompareGrid 4, "Bid", Grid1(0).Columns(int_Bid_Col).Text, !bid
            FillCompareGrid 5, "HCP", Grid1(0).Columns(int_HCP_Col).Text, !hcp
            FillCompareGrid 6, "Distrib.Pts", Grid1(0).Columns(int_HCP_Col + 1).Text, !Dist_Pts
            FillCompareGrid 7, "Custom Points", Grid1(0).Columns(int_HCP_Col + 5).Text, !Cust_Pts
            FillCompareGrid 8, "Shape", Grid1(0).Columns(int_xNum_Col - 2).Text, !Shape
            FillCompareGrid 9, "Intermediates", Grid1(0).Columns(int_xNum_Col - 1).Text, !intermed
            FillCompareGrid 10, "xNum", Grid1(0).Columns(int_xNum_Col).Text, !xNum
            FillCompareGrid 11, "xHCP", Grid1(0).Columns(int_xNum_Col + 1).Text, !xHCP
            FillCompareGrid 13, "sNum", Grid1(0).Columns(int_xNum_Col + 2).Text, !sNum
            FillCompareGrid 14, "sHCP", Grid1(0).Columns(int_xNum_Col + 3).Text, !sHCP
            FillCompareGrid 15, "hNum", Grid1(0).Columns(int_xNum_Col + 4).Text, !hNum
            FillCompareGrid 16, "hHCP", Grid1(0).Columns(int_xNum_Col + 5).Text, !HHCP
            FillCompareGrid 17, "dNum", Grid1(0).Columns(int_xNum_Col + 6).Text, !dNum
            FillCompareGrid 18, "dHCP", Grid1(0).Columns(int_xNum_Col + 7).Text, !dHCP
            FillCompareGrid 19, "cNum", Grid1(0).Columns(int_xNum_Col + 8).Text, !cNum
            FillCompareGrid 20, "cHCP", Grid1(0).Columns(int_xNum_Col + 9).Text, !cHCP
            FillCompareGrid 21, "combin.Num", Grid1(0).Columns(int_xNum_Col + 10).Text, "" & !cmNum
            FillCompareGrid 22, "combin.HCP", Grid1(0).Columns(int_xNum_Col + 11).Text, "" & !cmHCP
            FillCompareGrid 23, "Position", Grid1(0).Columns(int_TestHand_Col - 4), !pos
            FillCompareGrid 24, "Test Hand", Grid1(0).Columns(int_TestHand_Col).Text, !Test_Hand
            'g_CompareGrid.CellForeColor = vbRed
            FillCompareGrid 25, "", "^ EDIT ^", "^ EDIT ^"
            g_CompareGrid.CellBackColor = vbGreen
            g_CompareGrid.Col = 1
            g_CompareGrid.CellBackColor = vbGreen
            g_CompareGrid.Col = 0
            ck_Continuation.BackColor = QBColor(int_BGGrid1)
            ck_Continuation.ForeColor = vbYellow
            g_CompareGrid.CellBackColor = RGB(1, 1, 1)
            g_CompareGrid.CellFontBold = True
            g_CompareGrid.CellForeColor = vbWhite
            g_CompareGrid.ForeColorFixed = Grid1(int_GridNum).ForeColor
            g_CompareGrid.BackColorFixed = Grid1(int_GridNum).BackColor
            f_CompareGrid.Visible = True
            ''''''''
            End With
        Else
            mEdit_Click 0
        End If
        bln_Abort = True
        p_Testing.Visible = False
        Grid1(0).Refresh
        'Exit Sub
    End If
End Sub


Private Sub GetMinNT()
    ' Gets the minimum HCPs for an opening 1N and 2N: '

    Dim i As Integer
    Dim n As String
    Dim sel As String

    For i = 1 To 2
        If i = 1 Then
            n = "130000"
        Else
            n = "110000"
        End If

        sel = "SELECT * FROM Bids " & _
              "WHERE Prior_Bids = '' " & _
              "AND Bid_Order_Number = '" & n & "' " & _
              "ORDER BY Bid_Order_Number;"
        'Stop ' Do I need to use rs
        'Set rs_BidTable = db_BidBase.OpenRecordset(sel, dbOpenForwardOnly) ' I don't want to reset rs_BidTable. Just keep it as initialized. '
        ' There should only be the one entry with #130000, '
        ' but just to be sure:                             '
        With rs_BidTable
            If Not (.NoMatch Or .EOF) Then
                Do While !Pct_Used <> "" And _
                         !Bid_Order_Number = n And _
                         Not .EOF
                    .MoveNext
                    If .EOF Then Exit Do
                    DoEvents
                Loop
                If i = 1 Then
                    int_1NT_Min = Val(!hcp)
                Else
                    int_2NT_Min = Val(!hcp)
                End If
            End If
        End With
    Next
End Sub


Private Sub LookUpConvDataAndDisplayIt(Index As Integer)
    ' When the user selects a Convention Name,     '
    ' check to see if there is a note file for it: '
    ' Also look up the Selected and Sub fields and '
    ' put them in their boxes.                     '
    Dim num As Long
    Dim nam As String
    Dim sel As String
    Dim sel2 As String

    If Index = 0 Then
        nam = Trim$(cmb_ConvName(0).Text)
        If nam <> "" Then
            LookUpBidName num, sel, sel2, nam
        End If
        t_Sel = "" & sel
        t_Sub = "" & sel2
    End If
End Sub


Sub StoreBiddersHandSpecs()

End Sub


Function LosingTrickCount(str_Hands() As String, _
                          int_SuitQty() As Integer, _
                          hcp() As Integer _
                         ) As Single
    ' Compute Losing Trick Count: '
    ' This routine is from the BB Practice program. '

    ' It is not being used in the Editor program that I can see. '
    ' I assume it was for posting a user-entered hand to the     '
    ' input boxes. '

    Dim ltc(4) As Single
    'Dim hcp(4, 4) As Integer ' HCP by player & suit. hcp(p, 0) = total hand. '
    'Dim qty(4, 4)
    Dim hon(4, 4) As Integer ' # honors per suit (Ace-10) '
    Dim c As String
    Dim fitNS As Integer
    Dim fitEW As Integer
    Dim lottNS As Integer
    Dim lottEW As Integer
    Dim ltcCap(2) As String
    Dim player As Integer
    Dim suit As Integer
    Dim suitLen(4, 4) As Integer
    Dim i As Integer
    Dim Ace As Integer
    Dim King As Integer
    Ace = 1
    King = 13

    ' Add up HCP and # honors by player: '
    For player = 1 To 4
        suit = 4
        If player > 1 Then
            suitLen(player, suit) = Val(MyMsgBox("Enter minimum # of " & str_Hands(suit) & vbCr & _
                                                 "that " & Mid$("SWNE", player, 1) & " has shown, " & _
                                                 "if any:", 2, "", "Losing Trick Count"))

            hcp(player, 0) = Val(MyMsgBox("Enter minimum # of HCP" & vbCr & _
                                          "that " & Mid$("SWNE", player, 1) & " has shown, " & _
                                          "if any:", 2, "", "Losing Trick Count"))
        Else
            For i = 1 To 16
                c = Mid$(str_Hands(player), i, 1)
                If c = "-" Then
                    hcp(player, 0) = hcp(player, 0) + hcp(player, suit)
                    suit = suit - 1
                ElseIf c > "9" Then
                    hcp(player, suit) = hcp(player, suit) + InStr("JQKA", c)
                    hon(player, suit) = hon(player, suit) + 1
                End If
                hcp(player, 0) = hcp(player, 0) + hcp(player, suit) ' clubs '
            Next
        End If
    Next

    ' Look for 8-card fit or self-sufficient suit: '
    For suit = 1 To 4 ' Clubs to Spades '
        ' Law of Total Tricks: '
        If int_SuitQty(1, suit) + int_SuitQty(3, suit) >= lottNS Then
            ' Use the higher suit if >= quantity of lower suit. '
            lottNS = int_SuitQty(1, suit) + int_SuitQty(3, suit)
        End If

        If int_SuitQty(2, suit) + int_SuitQty(4, suit) >= lottEW Then
            lottEW = int_SuitQty(2, suit) + int_SuitQty(4, suit)
        End If

        ' Look for an 8+ card fit or self-sufficient suit:
        If int_SuitQty(1, suit) + int_SuitQty(3, suit) > 7 Or _
           int_SuitQty(1, suit) + hon(1, suit) > 9 Or _
           int_SuitQty(3, suit) + hon(3, suit) > 9 _
        Then
            fitNS = suit
        End If

        If int_SuitQty(2, suit) + int_SuitQty(4, suit) > 7 Or _
           int_SuitQty(2, suit) + hcp(2, suit) > 9 Or _
           int_SuitQty(4, suit) + hcp(4, suit) > 9 _
        Then
            fitEW = suit
        End If
    Next

    For player = 1 To 4
      If ((player = 1 Or player = 3) And fitNS) Or _
         ((player = 2 Or player = 4) And fitEW) _
      Then
        ltc(player) = 0
        For suit = 1 To 4
            If int_SuitQty(player, suit) = 1 Then
                If int_Suit(player, suit, Ace) = 0 Then  ' Not A '
                    ltc(player) = ltc(player) + 1
                End If
            ElseIf int_SuitQty(player, suit) = 2 Then
                If int_Suit(player, suit, Ace) > 0 Then  ' A? '
                    If int_Suit(player, suit, King) > 0 Then ' AK = 0 '
                    ElseIf int_Suit(player, suit, 12) > 0 Then ' AQ = .5 '
                        ltc(player) = ltc(player) + 0.5
                    Else ' Ax = 1 '
                        ltc(player) = ltc(player) + 1
                    End If
                ElseIf int_Suit(player, suit, King) > 0 Then  ' K? '
                    ltc(player) = ltc(player) + 1
                Else ' xx = 2 '
                    ltc(player) = ltc(player) + 2
                End If
            ElseIf int_SuitQty(player, suit) > 2 Then
                If int_Suit(player, suit, Ace) > 0 Then  ' A?? '
                    If int_Suit(player, suit, King) > 0 Then ' AK? '
                        If int_Suit(player, suit, 12) > 0 Then ' AKQ = 0 '
                            ' no losers. '
                        ElseIf int_Suit(player, suit, 11) > 0 Then 'AKJ = .5 '
                            ltc(player) = ltc(player) + 0.5
                        Else ' AKx = 1 loser '
                            ltc(player) = ltc(player) + 1
                        End If
                    ElseIf int_Suit(player, suit, 12) > 0 Then ' AQ? = .5 '
                        If int_Suit(player, suit, 11) > 0 Then 'AQJ = 1 '
                            ltc(player) = ltc(player) + 1
                        Else ' AQx = 2 '
                            ltc(player) = ltc(player) + 1.5
                        End If
                    Else ' Axx = 1 '
                        ltc(player) = ltc(player) + 2
                    End If
                ElseIf int_Suit(player, suit, King) > 0 Then  ' K?? '
                    If int_Suit(player, suit, 12) > 0 Then ' KQ?  '
                        If int_Suit(player, suit, 11) > 0 Then ' KQJ = 1 '
                            ltc(player) = ltc(player) + 1
                        ElseIf int_Suit(player, suit, 10) > 0 Then ' KQT = 1.5 '
                            ltc(player) = ltc(player) + 1.5
                        Else ' KQx = 2 '
                            ltc(player) = ltc(player) + 2
                        End If
                    Else  ' Kxx '
                        ltc(player) = ltc(player) + 2
                    End If
                ElseIf int_Suit(player, suit, 12) > 0 Then  ' Q??
                    If int_Suit(player, suit, 11) > 0 Or int_Suit(player, suit, 10) > 0 Then  ' QJx or QTx = 2 '
                        ltc(player) = ltc(player) + 2
                    Else
                        ltc(player) = ltc(player) + 2.5
                    End If
                Else ' xxx = 3 '
                    ltc(player) = ltc(player) + 3
                End If
            End If
        Next
        MsgBox "Break point"
'        If player = 3 Then
'            ltcCap(1) = "  24 - " & ltc(3) & " - " & ltc(1) & " = " & (24 - ltc(3) - ltc(1))
'        ElseIf player = 4 Then
'            ltcCap(2) = "  24 - " & ltc(4) & " - " & ltc(2) & " = " & (24 - ltc(4) - ltc(2))
'        End If
'      Else
'        If player = 3 Then
'            ltcCap(1) = "  No 8-card fit."
'        ElseIf player = 4 Then
'            ltcCap(2) = "  No 8-card fit."
'        End If
       End If
    Next

End Function


Private Sub ForceEntry2ShowInGrid(entryID As Long)
    Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
    Set rs_BidTable = db_BidBase.OpenRecordset(str_BidsTbl)
    rs_BidTable.Index = "PrimaryKey"
    rs_BidTable.Seek "=", entryID
    lng_ID = entryID
    DisplayBid2Edit 99
End Sub


Sub tempp()
    Dim i As Long
    Dim j As Integer
    Dim k As Integer
    Dim n As Integer
    Dim x As Integer
    Dim c As Boolean
    Dim m As String
    Dim r As String
    Dim s As String
    Dim xx As Long
On Error GoTo tempp_err
    Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
    'Set rs_BidTable = db_BidBase.OpenRecordset(str_BidsTbl)

    With rs_BidTable
    ''''''''''''''''
    .Index = "PrimaryKey"
    .Seek ">=", 1
    Do
        xx = xx + 1
        If Left$(!source & "   ", 3) = "Bul" Then
            s = !source & " "
            .Edit
            
            ' target is like "Bul:2020-05 p.50" '
            If Mid$(s, 4, 1) <> ":" Then
                Mid$(s, 4, 1) = ":"
            End If
                        
            If Mid$(s, 5, 1) = " " Then
                s = Left$(s, 4) & Mid$(s, 6)
            End If
                      '123456789-123456789-123456789-123456'
            k = InStr("JanFebMarAprMayJunJulAugSepOctNovDec", Mid$(s, 5, 3))
            If k > 0 Then
                ' "Bul:Nov.05 ... " '
                '  1234567890 '
                k = (k - 1) / 3 + 1
                m = Right$("0" & k, 2) ' "11" '
                s = "Bul:" & m & "-" & Mid$(s, 5 + 4)
            End If
            
            j = Val(Mid$(s, 5, 4))
            If j < 22 Then
                ' "Bul:11-02,p76 '
                '  123456789-123 '
                ' "Bul:02-2020   '
                If j < 10 And Mid$(s, 5, 1) = " " Then
                    Stop
                End If
                m = Right$("0" & j, 2)
                k = Val(Mid$(s, 8))
                If k < 21 Then
                    ' Bul:04-19... -- 2019-04'
                    ' 123456789- '
                    s = Left$(s, 4) & "20" & Right$("0" & k, 2) & "-" & m & Mid$(s, 10)
                ElseIf k > 60 And k < 100 Then ' ex: 98 : 1998 '
                    ' Bul:11-98... -- 1998-11'
                    ' 123456789-
                    s = Left$(s, 4) & "19" & k & "-" & m & Mid$(s, 10)
                Else ' Bul:11-2005... -- 2005-11 '
                     ' 123456789-12
                    s = Left$(s, 4) & k & "-" & m & Mid$(s, 12)
                End If
            End If
            
            s = Replace$(s, ",p", " p") ' Bul:2002-11,p76 '
            s = Replace$(s, ",", " ")
            s = Replace$(s, ". ", " ")
            s = Replace$(s, "--", "-")
            s = Replace$(s, " -", "-")
            j = InStr(9, s, "p")
            If j > 0 And Mid$(s, j + 1, 1) <> "." Then
                s = Left$(s, j) & "." & Mid$(s, j + 1)
            End If
            s = Replace$(s, "p. ", "p.")
            s = Replace$(s, ": p", " p")
            s = Replace$(s, "  ", " ", 1, 99)
            s = Replace$(s, "  ", " ", 1, 99)
            s = Replace$(s, "It's Your Call", "IYC")
            If Len(s) = 10 Then Stop
            
            s = RTrim$(s)
            
            'x = x + 1
            'If x = 30 Then
            '    Stop
            '    x = 1
            'End If
            
            If s <> !source Then
                Debug.Print Left$(s & Space$(20), 27); !source
                !source = s
                .Update
            Else
                'Debug.Print "No chg: "; !source
            End If
        End If
        
        If Not .EOF Then
            .MoveNext
        Else
            Exit Do
        End If
        DoEvents
    Loop
    ''''''''
    End With
tempp_err:
    Debug.Print xx
    'Resume Next
    Stop
    On Error GoTo 0
End Sub

Sub dvds()
    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    Dim t As String
    Dim x As String
    Dim d As String
    
    Open "c:\myfiles\dvd history.txt" For Input As #1
    Open "c:\myfiles\dvd history 2.txt" For Input As #2
    
    Do
        Line Input #1, n ' dvd number
        Line Input #1, x
        Line Input #1, t ' title
        Line Input #1, x ' blank
        Line Input #1, x ' junk
        Line Input #1, x ' blank
        Line Input #1, d ' date
        n = LTrim$(n)
        n = Left$(n, 8)
        n = Right$(n, 2) & "/" & Left$(n, 5)
        
        Print #2, d; "  "; t
    Loop While Not EOF(1)
    Close #1
    Close #2
End Sub


Sub ModSources()





Exit Sub
    Set db_BidBase = ws_BidBase.OpenDatabase(str_DBPathAndName)
    Set rs_BidTable = db_BidBase.OpenRecordset(str_BidsTbl)
    
    With rs_BidTable
    ''''''''''''''''
    .Index = "PrimaryKey"
    .MoveFirst
    Do
        If Left$(!source, 4) = "Bul:" Then
            .Edit
            !source = "Bul " & Mid$(!source, 5)
            .Update
        End If
        If .EOF Then Exit Do
        .MoveNext
    Loop
    ''''''''
    End With
End Sub



