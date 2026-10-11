VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PlayScr 
   Appearance      =   0  'Flat
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CardShark BidBase Practice Dealer"
   ClientHeight    =   8640
   ClientLeft      =   90
   ClientTop       =   2070
   ClientWidth     =   9960
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
   Icon            =   "BidBase Bidding Practice.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   576
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   664
   Begin VB.ComboBox cmb_PBNBoards 
      Height          =   315
      Left            =   1080
      TabIndex        =   115
      Text            =   "Combo1"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame f_Msg 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2565
      Left            =   2280
      TabIndex        =   106
      Top             =   3840
      Visible         =   0   'False
      Width           =   5460
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
         TabIndex        =   110
         Text            =   "MsgInput"
         Top             =   1900
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
         Left            =   200
         TabIndex        =   109
         Top             =   1900
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
         Left            =   1000
         TabIndex        =   108
         Top             =   1450
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
         Left            =   1800
         TabIndex        =   107
         Top             =   1450
         Visible         =   0   'False
         Width           =   1140
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
         Left            =   0
         TabIndex        =   112
         Top             =   0
         Width           =   4995
      End
      Begin VB.Label L_MsgBox 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Message goes here."
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
         TabIndex        =   111
         Top             =   480
         Width           =   5025
      End
      Begin VB.Label L_MsgBackground 
         BackColor       =   &H80000005&
         Height          =   2130
         Left            =   60
         TabIndex        =   113
         Top             =   360
         Width           =   5340
      End
   End
   Begin VB.ComboBox cmb_FileList 
      ForeColor       =   &H00800000&
      Height          =   3690
      ItemData        =   "BidBase Bidding Practice.frx":10CA
      Left            =   1185
      List            =   "BidBase Bidding Practice.frx":10CC
      Style           =   1  'Simple Combo
      TabIndex        =   114
      Top             =   1365
      Visible         =   0   'False
      Width           =   7650
   End
   Begin VB.CommandButton b_ShowAnalysis 
      Caption         =   "Bid w/ &Analysis"
      Height          =   645
      Left            =   8820
      TabIndex        =   15
      Top             =   6450
      Width           =   1095
   End
   Begin VB.CommandButton b_RestartDeal 
      Caption         =   "&Repeat"
      Height          =   315
      Left            =   2535
      TabIndex        =   16
      Top             =   6810
      Width           =   810
   End
   Begin VB.Frame f_Inbox 
      Height          =   2175
      Left            =   3600
      TabIndex        =   81
      Top             =   2250
      Visible         =   0   'False
      Width           =   2700
      Begin VB.CommandButton b_DoneDealNum 
         Caption         =   "Done"
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
         Left            =   900
         TabIndex        =   85
         Top             =   1750
         Width           =   790
      End
      Begin VB.TextBox t_Inbox 
         Alignment       =   2  'Center
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
         Left            =   700
         TabIndex        =   84
         Text            =   "9999999"
         Top             =   1300
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   250
         Left            =   0
         TabIndex        =   82
         Top             =   0
         Width           =   2700
      End
      Begin VB.Label L_InboxLbl 
         Alignment       =   2  'Center
         Caption         =   " Enter a deal #  or ""user"" to enter hands:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   300
         TabIndex        =   83
         Top             =   375
         Width           =   2000
      End
   End
   Begin VB.Frame f_ClearDeal 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1770
      Left            =   3630
      TabIndex        =   101
      Top             =   2400
      Visible         =   0   'False
      Width           =   2505
      Begin VB.CommandButton b_ClearDeal 
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   1
         Left            =   1560
         TabIndex        =   104
         Top             =   1250
         Width           =   400
      End
      Begin VB.CommandButton b_ClearDeal 
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   0
         Left            =   1080
         TabIndex        =   103
         Top             =   1250
         Width           =   400
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clear the existing deal and select new cards?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   170
         TabIndex        =   102
         Top             =   250
         Width           =   2115
      End
   End
   Begin VB.PictureBox p_Hand 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6330
      Index           =   2
      Left            =   30
      ScaleHeight     =   6330
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   60
      Width           =   1155
   End
   Begin VB.TextBox t_EnterHand 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Height          =   285
      Index           =   4
      Left            =   6930
      TabIndex        =   100
      Top             =   2940
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.TextBox t_EnterHand 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Height          =   285
      Index           =   3
      Left            =   3990
      TabIndex        =   99
      Top             =   1860
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.TextBox t_EnterHand 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Height          =   285
      Index           =   2
      Left            =   1050
      TabIndex        =   98
      Top             =   2940
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.TextBox t_EnterHand 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Height          =   285
      Index           =   1
      Left            =   3990
      TabIndex        =   97
      Top             =   4545
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Frame f_BillBoard 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   " "
      Height          =   3450
      Left            =   2400
      TabIndex        =   90
      Top             =   1485
      Visible         =   0   'False
      Width           =   5190
      Begin VB.CommandButton b_Billboard 
         Caption         =   "OK"
         Height          =   330
         Left            =   1885
         TabIndex        =   91
         Top             =   3060
         Width           =   1455
      End
      Begin VB.Label L_Billboard 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2100
         Left            =   700
         TabIndex        =   95
         Top             =   700
         Width           =   3700
      End
      Begin VB.Label Label2 
         Height          =   280
         Index           =   1
         Left            =   3465
         TabIndex        =   94
         Top             =   3000
         Width           =   1605
      End
      Begin VB.Label Label2 
         Height          =   280
         Index           =   0
         Left            =   150
         TabIndex        =   93
         Top             =   3000
         Width           =   1605
      End
      Begin VB.Label L_BillboardBack 
         Height          =   2895
         Left            =   135
         TabIndex        =   92
         Top             =   90
         Width           =   4920
      End
   End
   Begin VB.ComboBox cmb_DDA 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   30
      Locked          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   89
      Text            =   "DDA = Maximum tricks that can be taken by each player in each suit/NT."
      Top             =   7185
      Width           =   9870
   End
   Begin VB.PictureBox p_BidBoxLbl 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4065
      Picture         =   "BidBase Bidding Practice.frx":10CE
      ScaleHeight     =   300
      ScaleWidth      =   990
      TabIndex        =   86
      Top             =   3255
      Width           =   990
   End
   Begin MSFlexGridLib.MSFlexGrid g_BidBox 
      Height          =   2070
      Left            =   2550
      TabIndex        =   1
      Top             =   1545
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   3651
      _Version        =   393216
      Rows            =   8
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   65535
      GridColorFixed  =   49152
      Enabled         =   -1  'True
      GridLines       =   2
      ScrollBars      =   0
   End
   Begin VB.TextBox t_Disclose 
      ForeColor       =   &H00004000&
      Height          =   1275
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   3630
      Width           =   4860
   End
   Begin VB.CheckBox ck_KeepDealer 
      BackColor       =   &H0000C000&
      Caption         =   "Keep"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   1200
      TabIndex        =   80
      Top             =   6840
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Frame f_NSEW 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   900
      Left            =   7725
      TabIndex        =   76
      Top             =   6450
      Width           =   825
      Begin VB.CommandButton b_EnterHands 
         Caption         =   "Enter Hands"
         Height          =   650
         Left            =   0
         TabIndex        =   77
         Top             =   0
         Width           =   810
      End
   End
   Begin VB.CommandButton b_CancelHands 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   7575
      TabIndex        =   74
      Top             =   5055
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton b_DealRest 
      Caption         =   "Deal the rest"
      Height          =   405
      Left            =   1200
      TabIndex        =   73
      Top             =   5055
      Visible         =   0   'False
      Width           =   1305
   End
   Begin MSComDlg.CommonDialog CmDia 
      Left            =   7560
      Top             =   765
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox p_Hand 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1470
      Index           =   1
      Left            =   2535
      ScaleHeight     =   1470
      ScaleWidth      =   4920
      TabIndex        =   7
      Top             =   4950
      Width           =   4920
   End
   Begin VB.PictureBox p_Hand 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6360
      Index           =   4
      Left            =   8865
      ScaleHeight     =   6360
      ScaleWidth      =   1155
      TabIndex        =   6
      Top             =   60
      Width           =   1155
   End
   Begin VB.PictureBox p_Hand 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1470
      Index           =   3
      Left            =   2520
      ScaleHeight     =   1470
      ScaleWidth      =   4875
      TabIndex        =   5
      Top             =   0
      Width           =   4875
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9225
      Top             =   495
   End
   Begin VB.ComboBox cmb_FindDeals 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   30
      Style           =   1  'Simple Combo
      TabIndex        =   88
      Top             =   8595
      Width           =   9870
   End
   Begin VB.Frame f_CardDeck 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3570
      Left            =   1020
      TabIndex        =   20
      Top             =   1410
      Visible         =   0   'False
      Width           =   7815
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   41
         Left            =   7201
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   72
         Top             =   0
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   42
         Left            =   6601
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   71
         Top             =   0
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   43
         Left            =   6001
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   70
         Top             =   0
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   44
         Left            =   5401
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   69
         Top             =   0
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   45
         Left            =   4801
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   68
         Top             =   0
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   46
         Left            =   4201
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   67
         Top             =   0
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   47
         Left            =   3601
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   66
         Top             =   0
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   48
         Left            =   3001
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   65
         Top             =   0
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   49
         Left            =   2401
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   64
         Top             =   0
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   50
         Left            =   1801
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   63
         Top             =   0
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   51
         Left            =   1201
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   62
         Top             =   0
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   52
         Left            =   601
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   61
         Top             =   0
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   40
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   60
         Top             =   0
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   28
         Left            =   7170
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   59
         Top             =   900
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   29
         Left            =   6601
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   58
         Top             =   900
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   30
         Left            =   6001
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   57
         Top             =   900
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   31
         Left            =   5401
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   56
         Top             =   900
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   32
         Left            =   4801
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   55
         Top             =   900
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   33
         Left            =   4201
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   54
         Top             =   900
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   34
         Left            =   3601
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   53
         Top             =   900
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   35
         Left            =   3001
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   52
         Top             =   900
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   36
         Left            =   2401
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   51
         Top             =   900
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   37
         Left            =   1801
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   50
         Top             =   900
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   38
         Left            =   1201
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   49
         Top             =   900
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   39
         Left            =   601
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   48
         Top             =   900
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   27
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   47
         Top             =   900
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   15
         Left            =   7185
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   46
         Top             =   1800
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   19
         Left            =   4801
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   45
         Top             =   1800
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   16
         Left            =   6601
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   44
         Top             =   1800
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   18
         Left            =   5401
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   43
         Top             =   1800
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   17
         Left            =   6001
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   42
         Top             =   1800
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   20
         Left            =   4201
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   41
         Top             =   1800
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   21
         Left            =   3601
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   40
         Top             =   1800
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   22
         Left            =   3001
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   39
         Top             =   1800
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   23
         Left            =   2401
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   38
         Top             =   1800
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   24
         Left            =   1801
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   37
         Top             =   1800
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   25
         Left            =   1201
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   36
         Top             =   1800
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   26
         Left            =   601
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   35
         Top             =   1800
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   14
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   34
         Top             =   1800
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   2
         Left            =   7201
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   33
         Top             =   2700
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   3
         Left            =   6601
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   32
         Top             =   2700
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   4
         Left            =   6001
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   31
         Top             =   2700
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   5
         Left            =   5401
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   30
         Top             =   2700
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   6
         Left            =   4801
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   29
         Top             =   2700
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   7
         Left            =   4201
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   28
         Top             =   2700
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   8
         Left            =   3601
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   27
         Top             =   2700
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   9
         Left            =   3001
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   26
         Top             =   2700
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   10
         Left            =   2401
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   25
         Top             =   2700
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   11
         Left            =   1801
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   24
         Top             =   2700
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   12
         Left            =   1201
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   23
         Top             =   2700
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   13
         Left            =   601
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   22
         Top             =   2700
         Width           =   600
      End
      Begin VB.PictureBox Card 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   850
         Index           =   1
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   600
         TabIndex        =   21
         Top             =   2700
         Width           =   600
      End
   End
   Begin VB.PictureBox p_Bottom 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   30
      ScaleHeight     =   810
      ScaleWidth      =   9885
      TabIndex        =   2
      Top             =   6405
      Width           =   9885
      Begin VB.CommandButton b_FinishedAuctions 
         Caption         =   "Finished Auctions"
         Height          =   300
         Left            =   4110
         TabIndex        =   96
         Top             =   430
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.CommandButton b_DealBackFwd 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3345
         TabIndex        =   79
         Top             =   405
         Width           =   330
      End
      Begin VB.CommandButton b_DealBackFwd 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2180
         TabIndex        =   78
         Top             =   420
         Width           =   300
      End
      Begin VB.CommandButton b_Vul 
         Caption         =   "&Vulner."
         Height          =   315
         Left            =   30
         TabIndex        =   13
         Top             =   75
         Width           =   810
      End
      Begin VB.CommandButton b_Dealer 
         Caption         =   "Dea&ler"
         Height          =   315
         Left            =   30
         TabIndex        =   11
         Top             =   405
         Width           =   810
      End
      Begin VB.CommandButton b_Deal 
         Caption         =   "&Deal"
         Height          =   315
         Left            =   2520
         TabIndex        =   9
         Top             =   75
         Width           =   810
      End
      Begin VB.CommandButton b_Edit 
         Caption         =   "&Edit Bid"
         Height          =   315
         HelpContextID   =   1
         Left            =   6420
         TabIndex        =   8
         Top             =   405
         Width           =   900
      End
      Begin VB.CommandButton b_Ok 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "&BB Bid"
         Default         =   -1  'True
         Height          =   315
         Left            =   6420
         TabIndex        =   3
         Top             =   75
         Width           =   915
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FF00&
         Height          =   210
         Left            =   5130
         TabIndex        =   19
         Top             =   450
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblBattTime 
         Caption         =   "Battery Left:"
         ForeColor       =   &H00008000&
         Height          =   210
         Left            =   3825
         TabIndex        =   18
         Top             =   450
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label L_Vul 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   880
         TabIndex        =   14
         Top             =   105
         Width           =   795
      End
      Begin VB.Label L_Deal 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3390
         TabIndex        =   12
         Top             =   90
         Width           =   3000
      End
      Begin VB.Label L_Dealer 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   880
         TabIndex        =   10
         Top             =   435
         Width           =   200
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g_Bidding 
      Height          =   2070
      Left            =   5160
      TabIndex        =   0
      Top             =   1545
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   3651
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      FixedCols       =   0
      BackColor       =   8454016
      ForeColor       =   0
      BackColorFixed  =   8454016
      AllowBigSelection=   0   'False
      Enabled         =   0   'False
      TextStyleFixed  =   2
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      ScrollBars      =   0
      AllowUserResizing=   3
      FormatString    =   "^   S  |^   W  |^   N  |^    E"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox ck_FindBids 
      BackColor       =   &H0000C000&
      Caption         =   "Find/practice conventions"
      Height          =   220
      Left            =   3825
      TabIndex        =   87
      Top             =   6855
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.TextBox t_Calculating 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   105
      Text            =   "BidBase Bidding Practice.frx":45E2
      Top             =   6390
      Visible         =   0   'False
      Width           =   9900
   End
   Begin VB.Label L_EnterHandsHelp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   2985
      TabIndex        =   75
      Top             =   5085
      Width           =   4035
   End
   Begin VB.Menu zF 
      Caption         =   "&File"
      Begin VB.Menu mAddBids 
         Caption         =   "Add Bids not in BidBase"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mSaveDeal 
         Caption         =   "&Save Deal"
         Shortcut        =   ^S
      End
      Begin VB.Menu mLoadDeal 
         Caption         =   "&Load User Deal"
         Shortcut        =   ^L
      End
      Begin VB.Menu mLoadPbn 
         Caption         =   "Load PBN Event"
         Shortcut        =   ^P
      End
      Begin VB.Menu mLoadPbnDealNum 
         Caption         =   "Load PBN Deal #"
         Enabled         =   0   'False
         Shortcut        =   ^D
      End
      Begin VB.Menu zF1 
         Caption         =   "-"
      End
      Begin VB.Menu mRunEditor 
         Caption         =   "Run &Editor"
      End
      Begin VB.Menu mExit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mO 
      Caption         =   "&Options"
      Begin VB.Menu mAutoBid 
         Caption         =   "&Auto-Bid"
         Checked         =   -1  'True
      End
      Begin VB.Menu mHide 
         Caption         =   "&Hide Hands"
         Shortcut        =   ^H
      End
      Begin VB.Menu mStayOnTop 
         Caption         =   "Stay on Top"
         Shortcut        =   ^T
      End
      Begin VB.Menu zO1 
         Caption         =   "-"
      End
      Begin VB.Menu mEnterHands 
         Caption         =   "&Enter Hands"
      End
      Begin VB.Menu mRedisplayHands 
         Caption         =   "Redisplay Hands"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mScramble 
         Caption         =   "&Mix Up Cards"
         Index           =   0
         Shortcut        =   ^M
      End
      Begin VB.Menu mScramble 
         Caption         =   "&Remix Same Hands"
         Index           =   1
         Shortcut        =   ^R
      End
      Begin VB.Menu zO9 
         Caption         =   "-"
      End
      Begin VB.Menu mAverageDDA 
         Caption         =   "11-Deal &Avg. DDA"
      End
      Begin VB.Menu m100DealAvg 
         Caption         =   "101-Deal Average"
         Enabled         =   0   'False
         Shortcut        =   ^A
      End
      Begin VB.Menu zO2 
         Caption         =   "-"
      End
      Begin VB.Menu mRotate 
         Caption         =   "Rotate Hands W-S"
         Index           =   1
         Shortcut        =   ^W
      End
      Begin VB.Menu mRotate 
         Caption         =   "Rotate Hands N-S"
         Index           =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mRotate 
         Caption         =   "Rotate Hands E-S"
         Index           =   23
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu zC 
      Caption         =   "&Conventions"
      Begin VB.Menu mShowPracCnvDocs 
         Caption         =   "1. How conv. practice works"
      End
      Begin VB.Menu mCreateConv 
         Caption         =   "2. &Create/pick convention file"
      End
      Begin VB.Menu mFindDeals 
         Caption         =   "3. &Find matching deals"
         Enabled         =   0   'False
      End
      Begin VB.Menu mPractice 
         Caption         =   "4. &Practice convention"
         Enabled         =   0   'False
      End
      Begin VB.Menu mUpdateConventionFiles 
         Caption         =   "5. &Update convention file"
         Enabled         =   0   'False
      End
      Begin VB.Menu zO3 
         Caption         =   "-"
      End
      Begin VB.Menu mShowConvDocs 
         Caption         =   "&Show index to conventions"
      End
   End
   Begin VB.Menu mH 
      Caption         =   "&Help"
      Begin VB.Menu mHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mTutorial 
         Caption         =   "Tutorial"
         Enabled         =   0   'False
         Shortcut        =   {F4}
         Visible         =   0   'False
      End
      Begin VB.Menu mVer 
         Caption         =   "Version"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "PlayScr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'     CardShark BidBase Bidding Practice Program     '
'         Nelson Ford  www.aeyec.com/BidBase         '
'                                                    '
' Project - References:                              '
'        OLE Automation                              '
'        BidBaseDB.vbp (Language: English/US)        '
'        or BidBaseDB.DLL (Language: Standard)       '
' Project - Components:                              '
'        Common Dialog                               '
'        FlexGrid                                    '
'        SysInfo                                     '
'        Microsoft Windows Common Controls 6.0 (SP6) '

Option Explicit
Private Const Version$ = "2025.07.31"
' Changed DDA variables.
    
Public NewBid As BidBaseDB.Bidder
    ' Make sure BidBaseDB.vbp or .DLL is checked in the '
    ' Project - References menu. '
    ' If "Class not registered", open BidBaseDB.dll with  '
    ' regsvr32.exe in xplorer2. For more information, see '
    ' the end of CardShark_BidBase_Programming_Notes.htm. '

' Global and module variables have a 3-letter abbreviation '
' indicating their variable type.  Local variables in Subs '
' do not have these and their names always start lowercase '
' such as allCalls versus str_AllCalls. '
Private str_DBPathAndName As String

Private bln_AceAsking As Boolean
Private str_AceAskingCall As String ' e.g.: 1C-1H, 1N-2D, 2N-4N : "4N" '
Private str_AllCalls As String
Private bln_Analysis As Boolean
Private bln_AutoBid As Boolean
Private bln_Averaging As Boolean
Private int_BidRound As Integer
Private str_BidToAnalyze As String
Private str_Call(50) As String
Private int_CallNum As Integer
Private bln_ClearHands As Boolean ' When entering hands, start by clearing out existing deal. '
Private bln_Continue As Boolean
Private lng_CurrentDeal As Long ' Keep track of random deal # when User or PBN deals are loaded. '
Private str_DDA(4) As String
Private int_DDA(6, 5, 1001) ' 11 or 1001 iterations of DDA for each player and suit '
                          ' plus average for NS and for EW. '
                          ' See BidBase Practice docs. '
Private str_DDAOutput(4) As String ' Used to display DDA results on screen. '
Private str_DeletedBid As String ' When user clicks Edit. '
Private int_AvgDDANum As Integer ' 1 - 10 for int_DDA(h, s, int_AvgDDANum) above. '
Private str_DDATotal As String ' All DDA suit,player totals for the Bidder. '
Private bln_DontAsk4DealNum As Boolean
Private int_MsgClick As Integer
Private bln_SkipIt As Boolean
Private str_TimerCallingSub As String
Private str_Trumps As String        ' e.g.: If no agreement, 1st suit bid. Above : "H" '
Private int_TrumpSuit As Integer

Private int_OptAvg(4) As Integer
        ' Optimum for averaging deals.  '
        ' Optimum for players 1-4.      '
        ' Examples: 32=3D, 34=3S, 25=2N '
Private int_OptOrig(4) As Integer ' Optimum for original deal. '
Private int_OptByStrain(4, 5, 2) As Integer
Private str_OptByStrain(4, 5, 2) As String
        ' Where 4 = player #: S=1 E=2 N=3 W=4     '
        '       5 = strain #: C=1 D=2 H=3 S=4 N=5 '
        '       2 = # iterations: 1=11, 2=101     '
        ' Example: Opt for North's NT for 11      '
        '          iterations: (3, 5, 1) = 2      '

' Hand Eval: '
'Dim sng_LTC(4) As Single
'Dim int_HCP(4, 4) As Integer
'Dim int_Qty(4, 4) As Integer

Private int_AuctionNumber As Integer    ' Number of the auction to test for. '
Private str_FindAuctions(10) As String  ' The auction specs, such as 1S-P-1N-P-2C-etc. '  Limit is 10 auctions. '
Private int_DealFileEntryNum As Integer ' The record number in the file (1 to ...) of the last deal practiced.  '
Private lng_DealNum2Practice As Long    ' The BB deal number for the deal practiced. '
Private int_Dealer4PracticeDeal As Integer  ' Dealer # where S=1, W=2, N=3, E=4  '
Private str_BBpath As String
Private bln_Get2ndBid As Boolean
Private int_LastBidder As Integer       ' Last player to bid in a specified auction. '
Private lng_NumPracticeEntries As Long  ' The number of entries in a file. 1st entry = # entries in the file & last record used.'
Private str_PracticeSpecsFile As String ' The name of the selected auction specs file (i.e.: the name of the convention). '
Private str_PracticeDealsFile As String ' The name of the saved deals file (same as above but with .Deals extension). '
Private str_PracticeSpecsPath As String ' app.path & \PracticeDeals\ '
Private str_PracticeDealsPath As String ' str_PracticeSpecsPath & \Deals\ '
Private bln_PracDealIsSetup As Boolean
Private int_NumAuctions As Integer
Private int_NumAuctionSpecs As Integer  ' The number of auction specifications shown in the list box. '
Private int_PracRecordNum As Integer    ' Practice file record number being practiced. '
Private int_QuizQuestionNum As Integer  ' For when making quiz pages. '
Private bln_ReDoFindFile As Boolean
Private int_FoundAuctionLen As Integer  ' Excludes "|" and "x". '
Private bln_SettingUpPracBid As Boolean
Private bln_SearchForMoreDeals As Boolean
Private bln_UpdateCnvFiles As Boolean
Public stayman As Integer ' Just used when making NT Response practice files. '
Private Type lastDeal  ' For record 1 of the .deals file. '
    NumEntries As Long      ' e.g.: 128 '
    LastNumUsed As Integer  '       2   '
End Type
Private Type FoundDeal ' For records 2 - ... '
    dealNum As Long
    Dealer As Integer
End Type
' In the Bidding Routines module: '
'Global bln_FindAuctions As Boolean
'Global bln_FindAuctionsDone As Boolean
'Global int_AuctionNumber As Integer
'Global bln_PracticeFoundBids As Boolean

Private str_BON As String
Private lng_ID As Long
Private str_MyBid As String
Private bln_NoPaint As Boolean
Private int_NumPasses As Integer
Private lng_PaintTimer As Long
Private int_Player(4) As Integer
Private str_Player(4) As String
Private str_Source As String  ' Magazine, book, etc.
Private str_TimerJob
Private bln_UserDeal As Boolean

' PBN variabes: '
Private str_PBNevent As String ' Event of current board. '
Private str_PBNfile As String
Private str_PBNpathFile As String
Private str_PBNcontract As String
Private str_PBNround As String
Private int_PbnDealNum As Integer
Private str_PBNeventFile As String
Private int_PBNfileLineNum As Integer ' Each line in the PBN file is given a number. '
Private int_PBNlistNum As Integer ' The line in the deal list of the deal being viewed. '
Private str_PBNscore As String
Private str_PBNbids() As String
Private int_PBNpair As Integer
Private bln_OffPBN As Boolean ' No longer using PBN's bidding (if any). '
Private str_PBNanalysis As String ' Player's analysis in PBN files. '
Public bln_PbnEvent As Boolean ' Flag when displaying PBN event. '
' Stay on top: '
Private Declare Function SetWindowPos Lib "user32" _
        (ByVal hwnd As Long, _
         ByVal hWndInsertAfter As Long, _
         ByVal X As Long, Y, _
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
    mStayOnTop.Checked = False
End Sub

Public Sub MakeTopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    mStayOnTop.Checked = True
End Sub


Private Sub AddHandToDeal(ByVal hnd As String, h As Integer)
    ' User may enter a hand as a string. '
    ' This routine puts the cards into   '
    ' the int_Cards() array which is the '
    ' array that holds the whole deal.   '
    ' It also puts each card into the    '
    ' int_Suit(0,s,c) array so that we   '
    ' can easily tell which cards have   '
    ' been used.                         '
    
    ' This Sub is not being used since the card picker was added. '
    ' But it's being left here temporarily just in case. '
    
    Dim d As Integer
    Dim s As Integer
    Dim i As Integer
    Dim crd As String
    Dim n As Integer
    
    d = (h - 1) * 13 ' Pointer for int_Cards() '
    s = Spd ' Start with suit 4 (Spades) on left. '
    For i = 1 To 16
        ' Get the card: '
        crd = Mid$(hnd, i, 1)
        If crd = "-" Then
            ' Next suit: '
            s = s - 1
        
        ElseIf crd <> "" Then
            n = InStr("A23456789TJQK", crd)
            d = d + 1 ' increment deck pointer '
            ' Put into the deal: '
            int_Cards(d) = (s - 1) * 13 + n
            ' Put into the int_Suit array: '
            int_Suit(0, s, n) = int_Cards(d)
            int_SuitQty(0, s, n) = int_SuitQty(0, s, n) + 1
        End If
    Next
End Sub


Private Function allCalls() As String
    ' Computes the Prior Bids string. '
    ' Ex.: P-1C-D-P-1H-P-2H, etc.     '
    
    Dim i As Integer
    Dim a As String
    
    a = ""
    'Stop ' Below was int_CallNum - 2 but was not getting the last prior call. '
         ' This was after doing a Find Bids for "1H-P"
    For i = 1 To int_CallNum - 1
        
        a = a & str_Call(i)
        
        If i < int_CallNum - 1 Then
            a = a & "-"
        End If
    
    Next
    
    If Left$(a, 1) = "-" Then a = Mid$(a, 2)
    
    allCalls = a
End Function


Private Sub SetUpBidDisplayGrid()
    ' Clear the grid that displays '
    ' the biding and rewrite the   '
    ' column headings.             '
    With g_Bidding
        .Clear
        .row = 0
        .RowHeight(0) = 260
        '.RowHeight(7) = 280
        .col = 0: .Text = "S"
        .col = 1: .Text = "W"
        .col = 2: .Text = "N"
        .col = 3: .Text = "E"
        .ColWidth(0) = 530
        .ColWidth(1) = 530
        .ColWidth(2) = 530
        .ColWidth(3) = 535
        .Refresh
    End With
End Sub


Private Sub SetUpBiddingGrid()
    ' Fill out the grid on which '
    ' bids are selected.         '
    
    Dim c As Integer
    Dim r As Integer
  
    With g_BidBox
        For c = 0 To 4
            .ColWidth(c) = 500
        Next
    
        .RowHeight(7) = 300
    
        For r = 1 To 7
            .row = r - 1
            For c = 0 To 4
                .col = c
                .Text = r & Mid$("CDHSN", c + 1, 1) & " "
            Next
        Next
        .row = 7
        .col = 0
        .Text = "Pass"
        .col = 1
        .Text = " Dbl"
        .col = 2
        .Text = "Rdbl"
        .col = 3
        .Text = " BID"
        .col = 4
        .Text = "BOX"
    End With
    
End Sub


Private Sub GetDealer()
    ' Prompt the user for the Dealer: '
    Dim d As String
    Dim n As Integer
    
    If ck_KeepDealer.Value = 1 Then
        Exit Sub
    End If
    
        L_InboxLbl = "Enter dealer" & vbCrLf & "(N S E W)"
        b_DoneDealNum.Default = True
        t_Inbox.Alignment = 2
        t_Inbox = L_Dealer.Caption
        f_Inbox.ZOrder 0
        f_Inbox.Visible = True
        t_Inbox.SetFocus
        Do While f_Inbox.Visible
            n = DoEvents()
        Loop
        b_DoneDealNum.Default = False
        t_Inbox.Alignment = 0
        b_DoneDealNum_Click
End Sub


Private Sub b_Billboard_Click()
    f_BillBoard.Visible = False
End Sub


Private Sub b_Billboard_LostFocus()
    f_BillBoard.Visible = False
End Sub


Private Sub b_CancelHands_Click()
    Dim h As Integer
    
    ClearVariables
    b_Deal.Enabled = True
    b_RestartDeal.Enabled = True
    b_Ok.Enabled = True
    b_ShowAnalysis.Enabled = True
    b_EnterHands.Enabled = True
    f_CardDeck.Visible = False
    
    For h = 1 To 4
        t_EnterHand(h).Visible = False
    Next
    
    str_TimerCallingSub = "b_CancelHands_Click"
    str_TimerJob = "Cancel entering hands"
    Timer1.Enabled = True
End Sub


Public Sub ClearVariables()
    Dim i As Integer
    Dim j As Long
    
    p_Hand(1).Visible = True
    For i = 1 To 4
        j = DoEvents()
        p_Hand(i).Cls
    Next
    
    Erase int_HandQty()
    Erase int_Hand()
    Erase bln_SelectedCards()
    Erase int_Suit()
    ReDim int_Suit(4, 4, 13)
    Erase int_SuitQty()
    ReDim int_SuitQty(4, 4)
    Erase bln_PickedHand()
    'ReDim bln_PickedHand(4) already dimensioned
    
    Erase str_Hands()
    ReDim str_Hands(4)
    Erase int_Cards()
    ReDim int_Cards(65)
   
    b_EnterHands.Visible = True
    'L_EnterHandsHelp.Visible = False
    f_CardDeck.Visible = False
    b_CancelHands.Visible = False
    b_DealRest.Visible = False
End Sub


Private Sub b_ClearDeal_Click(Index As Integer)
    bln_ClearHands = (Index = 0)
    f_ClearDeal.Visible = False
End Sub


Private Sub b_Deal_Click()
    Dim i As Integer
    Dim response As String
    
    If bln_UserDeal Or bln_PbnEvent Then
        response = MyMsgBox("This will end viewing " & vbCr & L_Deal & vbCr & _
                            "and return to random deals." & vbCr & _
                            "Continue?", 4, "YN", "")
        If response = "no" Then Exit Sub
        
        b_DealBackFwd(0).Enabled = True
        b_DealBackFwd(1).Enabled = True
        str_PBNeventFile = ""
        bln_PbnEvent = False
        bln_UserDeal = False
        int_PbnDealNum = 0
        mLoadPbnDealNum.Enabled = False
        L_Deal.Top = 120
        L_Deal.Height = 255
        L_Deal = "" & lng_CurrentDeal
        lng_DealNum = lng_CurrentDeal
        lng_DealNumSpecified = lng_DealNum
        FinishDealSetup
        Deal
        For i = 1 To 4
            DisplayHand i
        Next
        Open "Last Deal Number.txt" For Output As #1
            Print #1, lng_DealNum
        Close #1
        Exit Sub
    End If
    
    If lng_DealNumSpecified = 0 Then
        GetDealNum ' in DealerRoutines module '
    End If
    bln_Rebid = True ' So that it will use the deal number just entered. '
    
    bln_HandsDealt = False
    Deal ' in DealerRoutines module '
    
    For i = 1 To 4
        DisplayHand i
    Next
    
    lng_DealNumSpecified = 0
    
    If lng_DealNum = Val(L_Deal) + 1 Or lng_DealNum = Val(L_Deal) - 1 Then
        lng_DealNumSpecified = lng_DealNum
        L_Deal = lng_DealNum
        bln_DontAsk4DealNum = True
    End If
    
    int_DealerSpecified = int_Dealer
    FinishDealSetup
End Sub
        
        
Private Sub FinishDealSetup()
    mScramble(1).Enabled = False ' Turn off Remix menu item. '
    Erase str_Call()
    str_AllCalls = ""
    str_Source = ""
    int_FoundAuctionLen = 0  ' Length of the auction without "|", "x", "any". '
    If int_DealerSpecified > 4 Then
        int_DealerSpecified = 1
    End If
    int_Player(1) = int_DealerSpecified
    int_Turn = 1
    SetPlayerNumbers ' for Find Deals, this was done in ProcessFindBids '
    If Not bln_PracticeFoundBids Then
        bln_SameDeal = False
        bln_Rebid = False
        If bln_FindAuctions = False Then
            SetUpBidDisplayGrid ' It's cleared below. '
            b_Ok.Caption = "&BB Bid"
        End If
        int_CallNum = 0
        int_NumPasses = 0
        int_AuctionNumber = 1
        If bln_Abort Then
            bln_Abort = False
            Exit Sub
        End If
    Else
        int_CallNum = 0
        bln_SameDeal = True
        bln_Rebid = True
    End If
    
    If int_Player(int_Turn) > 4 Then
        int_Player(int_Turn) = 1
    End If
    
    SetUpBidDisplayGrid ' Clear the grid that shows the bids for previous deal. '
    
    'If mAutoBid.Checked Then
        If str_Hands(1) = "" Then
            MsgBox "Code error in Sub FinishDealSetup. '"
        End If
        If bln_FindAuctions = False And Not bln_SkipIt Then
            ShowDDA False
        End If
    'End If
    
    If bln_FindAuctions Then
        ' The flag above should mean that we are '
        ' looking for deals which start with one '
        ' of the specified auctions, yet further '
        ' down it sounds like it is loading a    '
        ' deal from the Deals file which has     '
        ' already been found as a match but which'
        ' no longer matches. '
    
        Dim i As Integer
        Dim j As Integer
        Dim noStartMatch As Boolean
        Dim auction As String
        Dim allCalls As String
        int_NumAuctionSpecs = cmb_FindDeals.ListCount
        Do ' Step through all the bids needed to find an auction for a convention. '
            ' Make all the bids in the specified auction: '
            ProcessBid ' Get the first/next bid in the auction. '
            noStartMatch = True
            For i = 0 To int_NumAuctionSpecs - 1
                auction = cmb_FindDeals.List(i)
                j = InStr(auction, "|") ' Don't show bids after "|". '
                If j > 0 Then
                    auction = Left$(auction, j - 2)
                    allCalls = Left$(str_AllCalls, Len(auction))
                End If
                If auction = allCalls Then
                    noStartMatch = False
                    Exit Do
                End If
            Next
        Loop While noStartMatch = False
        
        If noStartMatch Then
            ' When finding bids for an auction, '
            ' if noStartMatch then go to the    '
            ' next bidder, so I don't know what '
            ' the code starting with DisplayBid '
            ' is for, so I'm going to use Timer '
            ' to move to the next bidder. '
            If int_Dealer < 4 Then
                int_Dealer = int_Dealer + 1
                str_TimerJob = "Restart same deal"
            Else
                int_Dealer = 1
                str_TimerJob = "Go to next deal"
            End If
            int_DealerSpecified = int_Dealer
            Timer1.Enabled = True
            Exit Sub
        End If
        
        Stop
        
        DisplayBid 1
        bln_FindAuctions = False  ' ???????????? '
        If noStartMatch Then
            i = MsgBox("The bids so far for deal #" & L_Deal & " do not" & Chr$(13) & _
                   "match the start of any of the specified" & Chr$(13) & _
                   "auctions, possibly because bids in the" & Chr$(13) & _
                   "database have changed. Do you want" & Chr$(13) & _
                   "the deal to be deleted from the list?" & Chr$(13) & _
                   "(Or you can change the entries.)", vbYesNo)
            If i = vbYes Then
                DelPracticeDeal
                b_DealBackFwd_Click 1
                'b_DealBackFwd_Click 0
                Exit Sub
                     ' The Del writes the last deal record in the file over this one. '
                     ' Now we need to set up that deal as if we had advanced a deal.  '
                     ' Try just using the Repeat button. '
                     ' Make sure this sub doesn't continue somewhere. '
            Else
                Exit Sub
            End If
        End If
        
        ' Rotate hands so that South is next to bid: '
        If int_Player(int_Turn) + 1 = 2 Then
            mRotate_Click 1 ' W to S '
        ElseIf int_Player(int_Turn) + 1 = 3 Then
            mRotate_Click 2 ' N to S '
        ElseIf int_Player(int_Turn) + 1 = 4 Then
            mRotate_Click 3 ' E to S '
        End If
    End If
End Sub


Private Sub FindPracticeDeals()
    Dim n As Integer ' numDealsFound '
    Dim a As Integer
    Dim auction As String
    
    lng_DealNum = 0
 Stop
    Do While n < 100
        lng_DealNum = lng_DealNum + 1
        ' Deal '
        int_Dealer = 0
        Do
            ' 1. Increment int_Dealer if  '
            '     <4 else exit this loop. '
            If int_Dealer = 4 Then Exit Do
            int_Dealer = int_Dealer + 1
            int_Turn = 1
            int_Player(int_Turn) = 1
            SetPlayerNumbers
            int_CallNum = 1
            str_Call(1) = ""
            str_AllCalls = ""
            ' 2. Make bid for dealer.     '
            MakeBid int_Player(int_Turn)
            ' 3. If bid matches auction,  '
            '     skip to step 6.         '
            
            ' 4. No match to auction so   '
            '     increment auction # and '
            '     repeat this step.       '
            ' 5. If no more auctions,     '
            '     return to step 1.       '
            ' 6. Bid matches auction.     '
            '     If more bids in auction,'
            '     increment int_Turn and  '
            '     get next bid. Go to #3. '
            ' 7. str_AllCalls matches a   '
            '     specified auction.      '
            '     Save deal # & dealer.   '
            ' 8. If n < 100, increment n. '
            '     Increment Deal.  Start  '
            '     the main Do loop over.  '
            ' 9. 100 deals matched/saved. '
            '     Exit the main do loop.  '
        Loop
    Loop
Stop
            If int_Dealer < 4 Then
                int_Dealer = int_Dealer + 1
            Else
                int_Dealer = 1
            End If
            int_DealerSpecified = int_Dealer
            int_Turn = 1
        
        
        Do While int_Turn < 4
            ' Start with
            int_Turn = int_Turn + 1
            a = 0
            Do While a < int_NumAuctionSpecs
                a = a + 1
                MakeBid int_Player(int_Turn)
                Stop
                'If str_Call = auction Then
                '    Exit Do ' Get next bid '
                'End If
            Loop
        Loop
End Sub


Private Sub b_DealBackFwd_Click(Index As Integer)
    ' Decrement (Index=0) or increment (Index=1) '
    ' the deal number and start that deal. '
    Dim n As Long
    Dim d As Long
    Dim rd As String
    Dim i As Integer
    
    If bln_PracticeFoundBids Then
        If Index = 0 Then
            If int_DealFileEntryNum = 1 Then
                MsgBox "This was the first saved deal."
                Exit Sub
            End If
            int_DealFileEntryNum = int_DealFileEntryNum - 1
        Else
            If int_DealFileEntryNum = lng_NumPracticeEntries Then
                MsgBox "This was the last saved deal."
                bln_UpdateCnvFiles = False
                mUpdateConventionFiles.Checked = False
                bln_Abort = True
                mPractice_Click
                bln_Abort = False
                Exit Sub
            End If
            int_DealFileEntryNum = int_DealFileEntryNum + 1
        End If
        
        GetPracticeDeal
        lng_DealNum = lng_DealNum2Practice
        L_Deal = "" & lng_DealNum
        int_DealerSpecified = int_Dealer
        bln_HandsDealt = False
        bln_SameDeal = False
        
        GoTo ExitDealBackFwd
    End If
    
    bln_HandsDealt = False
    bln_FindAuctionsDone = False
    bln_OffPBN = False
    
    If Index = 0 Then
        If bln_PbnEvent Then
            int_PBNlistNum = int_PBNlistNum - 1
            If int_PBNlistNum < 0 Then
                int_PBNlistNum = cmb_PBNBoards.ListCount - 1
            End If
            ClearVariables
            GetPbnDeal
            Exit Sub
        Else
            lng_DealNum = Val(L_Deal) - 2
            ' When Index=0, the left arrow was clicked, '
            ' so I want the next deal # to be one less  '
            ' than the current #. Since lng_DealNum will'
            ' be incremented before too long, I need to '
            ' decrement it by 2 now. '
            If ck_KeepDealer.Value = 0 Then
                int_Dealer = int_Dealer - 2 ' It will be increased by 1. '
                int_DealerSpecified = int_Dealer
                SetPlayerNumbers
            End If
        End If
        
    Else ' next deal arrow clicked '
        If bln_PbnEvent Then
            int_PBNlistNum = int_PBNlistNum + 1
            ClearVariables
            GetPbnDeal
            Exit Sub
        
        ElseIf bln_SearchForMoreDeals Then
            If int_Dealer4PracticeDeal > 0 Then
                L_Dealer = Left$(str_Dealer(int_Dealer4PracticeDeal), 1)
            End If
            int_DealerSpecified = int_Dealer
            SetPlayerNumbers
        
        ElseIf ck_KeepDealer.Value = 0 Then
            int_Dealer = int_Dealer + 1
            If int_Dealer = 5 Then int_Dealer = 1
            int_DealerSpecified = int_Dealer
            SetPlayerNumbers
        End If
    End If


ExitDealBackFwd:
    bln_HandsDealt = False
    bln_FindAuctions = False
    Deal
    If Not bln_SearchForMoreDeals Then
        For n = 1 To 4: DisplayHand n: Next
    End If
    bln_HandsDealt = False
    bln_DontAsk4DealNum = True
    FinishDealSetup
    MakeBid int_Player(int_Turn)
    Open "Last Deal Number.txt" For Output As #1
        Print #1, lng_DealNum
    Close #1
End Sub


Private Sub b_Dealer_Click()
    ck_KeepDealer.Value = 0
    GetDealer
    b_RestartDeal_Click
End Sub


Private Sub b_DealRest_Click()
    ' This Sub deals the cards left  '
    ' after the user picks cards for '
    ' one or more hands. '
    Dim c As Integer
    Dim h As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim tempi As Integer
    Dim numToDeal As Integer
    Dim d(4) As Integer
    Dim s As Integer
    Dim n As Integer
    Dim X As Integer
    Dim f As String ' file name '
    Dim fiNum As Integer
    Dim cardNumInHand As Integer
    Dim cards2(52) As Integer
    Dim dealRestHands(4) As Boolean
    
    If b_DealRest.Caption = "DONE" Then
        f_CardDeck.Visible = False
        b_DealRest.Caption = "Deal the rest"
        b_DealRest.Visible = False
        b_CancelHands.Visible = False
        b_Deal.Enabled = True
        b_RestartDeal.Enabled = True
        b_Ok.Enabled = True
        b_ShowAnalysis.Enabled = True
        b_EnterHands.Enabled = True
        For h = 1 To 4
            t_EnterHand(h).Visible = False
        Next
        L_Deal = "User"

        bln_UserDeal = True
        i = MsgBox("Save deal?", vbYesNo)
        If i = vbYes Then
            If Dir$(App.Path & "\UserEnteredHands\..\*", vbDirectory) = "" Then
                MkDir App.Path & "\UserEnteredHands"
            End If
            With CmDia
                .DialogTitle = "Enter file name:"
                .Filter = "*.txt"
                .InitDir = App.Path & "\UserEnteredHands\"
                .FileName = "*.txt"
                .ShowSave
                f = CmDia.FileName
                If f = "" Then Exit Sub
                ' Note that f includes the path and file name. '
            End With
        
            If Dir$(f) <> "" Then
                X = MsgBox("Overwrite existing file with that name?", vbYesNoCancel)
            
                Select Case X
                    Case Is = vbCancel
                        Exit Sub
                    Case Is = vbNo
                        f = "temp.txt"
                        L_Deal = f
                End Select
            End If
            
            FillSuits
            
            fiNum = FreeFile
            Open f For Output As #fiNum
            For i = 1 To 52
                Print #fiNum, int_Cards(i)
                Debug.Print int_Cards(i)
            Next
            Close #fiNum
        End If
        cmb_DDA.Clear
        MsgBox "Enter Vulnerability and Dealer and click Repeat."
        ShowDDA False
        Exit Sub
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ReDim int_Cards(52)
    
    ' Fill out int_Cards(c) from int_Hand(h, c) '
    ' for the cards which they were dealt. '
    For h = 1 To 4
        If t_EnterHand(h) <> "" Then
            n = (h - 1) * 13
            For c = 1 To 13
                int_Cards(n + c) = int_Hand(h, c)
            Next
        Else
            dealRestHands(h) = True
        End If
    Next
    
    n = 0
    ' Get the cards not yet dealt: '
    For c = 52 To 1 Step -1
        If bln_SelectedCards(c) = False Then
            n = n + 1
            cards2(n) = c
            bln_SelectedCards(c) = True
            DrawBack Card(c).hwnd, 3, 0, 0
        End If
    Next
    
    k = 0
    For h = 1 To 4
        If int_HandQty(h) = 0 Then
            For c = 1 To 13
                int_Hand(h, c) = cards2(k + c)
                ' Note that cards2() is an array of just  '
                ' the previously undealt cards which have '
                ' already been sorted. '
                int_Cards((h - 1) * 13 + c) = cards2(k + c)
            Next
            k = k + 13
            int_HandQty(h) = 13
        End If
    Next
    
    numToDeal = n / 13 ' e.g.: If 26 cards were left, 26/13 = 2 hands '
    
    If numToDeal > 1 Then
        ' More than one hand, so shuffle: '
        For j = 1 To 10 + Int(5 * Rnd) ' # of times to shuffle '
            For i = 1 To j
                ' Get a random array # '
                k = Int(n * Rnd + 1)
                ' Swap with current array #: '
                tempi = cards2(i)
                cards2(i) = cards2(k)
                cards2(k) = tempi
            Next
        Next
    End If
    
    int_NumPickedHands = numToDeal
    SortHands
    
    For h = 1 To 4
        If dealRestHands(h) Then
            str_Hands(h) = ""
            s = 4
            For i = 1 To 13
                n = int_Hand(h, i)
                j = GetCardSuit(n)
                k = GetCardValue(n)
                If j < s Then
                    str_Hands(h) = str_Hands(h) & "-"
                    s = j
                End If
                str_Hands(h) = str_Hands(h) & Mid$("A23456789TJQK", k, 1)
                int_Suit(h, s, k) = n
                int_SuitQty(h, s) = int_SuitQty(h, s) + 1
            Next
            bln_SkipIt = True
            t_EnterHand(h) = str_Hands(h)
            bln_SkipIt = False
            DisplayHand h
        End If
    Next
    
    b_DealRest.Caption = "DONE"
    int_NumPickedHands = 4
End Sub


Private Sub b_DoneDealNum_Click()
    If InStr(L_InboxLbl.Caption, "deal number") > 0 Then
        lng_DealNum = Val(t_Inbox)
        lng_DealNumSpecified = lng_DealNum
        L_Deal = t_Inbox
    ElseIf InStr(L_InboxLbl.Caption, "dealer") > 0 Then
        ' It should not be possible for user to enter an invalid '
        ' character for a dealer, so that's not being checked.   '
        int_DealerSpecified = InStr("SWNE", t_Inbox)
        If int_DealerSpecified = 0 Then int_DealerSpecified = 1
        L_Dealer = t_Inbox
        ck_KeepDealer.Visible = True
    End If
    f_Inbox.Visible = False
End Sub


Private Sub b_Edit_Click()
    Dim xxx As Long
    Dim pc As String
    Dim i As Integer
    Dim fiNum As Long
    
  If bln_SkipIt <> True Then
    For i = 1 To int_CallNum - 1
        If i < int_CallNum - 1 Then
            pc = pc & str_Call(i) & "-"
        Else
            pc = pc & str_Call(i)
        End If
    Next
    pc = Replace(pc, "?", "")
    
    'ChDir str_BBpath
    fiNum = FreeFile
    Open App.Path & "\EditPractice.txt" For Output As #fiNum
    Print #fiNum, str_Hands(int_Player(int_Turn))
    Print #fiNum, pc
    If bln_PbnEvent Or bln_UserDeal Then
        Print #fiNum, "Deal #" & L_Deal.Caption & "-" & L_Dealer
    Else
        Print #fiNum, "Deal #" & L_Deal & "-" & L_Dealer & " ID=" & lng_ID & " BON=" & str_BON
    End If
    Close #fiNum
  End If
    
    str_DeletedBid = str_Call(int_CallNum)
    str_Call(int_CallNum) = ""
    i = InStrRev(str_AllCalls, "-")
    If i > 0 Then
        str_AllCalls = Left$(str_AllCalls, i - 1)
    Else
        str_AllCalls = ""
    End If
    DisplayBid int_CallNum
    int_CallNum = int_CallNum - 1
    t_Disclose = ""
    
    If mStayOnTop.Checked Then
        ' Doesn't always stay on top. '
        SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
        DoEvents
        SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
        DoEvents
    End If
        
    'MsgBox ("Deal info has been sent to the BidBase Editor." & vbCr & _
            "Press Ctrl-I in that program to import it.")
End Sub


Private Sub b_EnterHands_Click()
    Dim i As Integer
    Dim j As Integer
    Dim h As Integer
    Dim s As Integer
    Dim c As Integer
    Dim hnd As String
    Dim d As Integer
    
    b_Deal.Enabled = False
    b_RestartDeal.Enabled = False
    'b_Ok.Enabled = False
    b_ShowAnalysis.Enabled = False
    t_Disclose.Text = ""
    int_Turn = 1
    
    b_EnterHands.Enabled = False
    f_CardDeck.Visible = True
    f_CardDeck.ZOrder 0
    
    DoEvents
    For i = 1 To 52
        DrawBack Card(i).hwnd, 3, 0, 0
        bln_SelectedCards(i) = False
        DoEvents
    Next
    
    f_ClearDeal.Visible = True
    f_ClearDeal.ZOrder 0
    Do While f_ClearDeal.Visible
        DoEvents
    Loop
        
        For i = 19 To 36
            ' The card backs don't refresh after '
            ' f_ClearDeal goes away, so... '
            DrawBack Card(i).hwnd, 3, 0, 0
            DoEvents
        Next
            
    bln_SkipIt = True
    For h = 1 To 4
        t_EnterHand(h).Visible = True
        t_EnterHand(h).ZOrder 0
    Next
    
    If Not bln_ClearHands Then
        For h = 1 To 4
            t_EnterHand(h).Text = str_Hands(h)
            int_HandQty(h) = Len(str_Hands(h)) - 3
            If int_HandQty(h) = 13 Then
                int_NumPickedHands = int_NumPickedHands + 1
            End If
        Next
    
    Else
        ' Yes was clicked to clear all hands. '
        On Error Resume Next
        Erase bln_PickedHand()
        int_NumPickedHands = 0
        Erase int_Suit()
        ReDim int_Suit(4, 4, 13)
        Erase int_SuitQty()
        ReDim int_SuitQty(4, 4)
        Erase int_Cards()
        ReDim int_Cards(65)
        Erase int_HandQty()
        ReDim int_HandQty(4)
        Erase bln_SelectedCards()
        For i = 1 To 52
            bln_SelectedCards(i) = False
        Next
        
        For i = 1 To 4
            j = DoEvents()
            p_Hand(i).Cls
            p_Hand(i).Refresh
            Erase str_Hands()
            ReDim str_Hands(4)
            Erase int_Hand()
            ReDim int_Hand(4, 13)
            j = DoEvents()
            int_HandQty(i) = 0
            t_EnterHand(i).Text = ""
            If i = 1 Then
                t_EnterHand(1).BackColor = 65535 ' yellow '
                int_Turn = 1
            Else
                t_EnterHand(i).BackColor = 13434828 ' green '
            End If
        Next

        On Error GoTo 0
        bln_ClearHands = False
        
        SetUpBidDisplayGrid ' Clear the grid that shows the bids. '
        LayOutDeck ' Sub in DealerRoutines '
    End If
            
    b_DealRest.Caption = "Deal the rest"
    b_DealRest.Visible = True
    b_CancelHands.Visible = True
    
    If mStayOnTop.Checked Then
        SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
        DoEvents
        SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
        DoEvents
    End If
End Sub


Private Sub b_FinishedAuctions_Click()
    Dim i As Long
    Dim response As String
    
    b_FinishedAuctions.Visible = False
    
    If Left$(cmb_FindDeals.Text, 5) = "Enter" Then
        cmb_FindDeals.Text = str_PracticeSpecsFile
    End If
    'If cmb_FindDeals.Text <> "" Then
    '    cmb_FindDeals_KeyPress 13
    'End If
    
    Erase str_FindAuctions()
    int_AuctionNumber = 1 ' The number of the auction array currently being checked. '
    int_NumAuctionSpecs = cmb_FindDeals.ListCount
    If int_NumAuctionSpecs = 0 Then
        Exit Sub
    End If

    ' Put the starting bid specs into an array. '
    For i = 1 To int_NumAuctionSpecs
        str_FindAuctions(i) = Trim$(cmb_FindDeals.List(i - 1))
    Next
    
    ChDir str_BBpath ' Probably not necessary. '
    
    ' Rewrite the Specs file in case of changes. '
    Open str_PracticeSpecsPath & str_PracticeSpecsFile For Output As #1
        For i = 1 To int_NumAuctionSpecs
            Print #1, str_FindAuctions(i)
        Next
    Close #1
    
    mFindDeals.Enabled = True
    If FileLen(str_PracticeDealsPath & str_PracticeDealsFile) > 10 Then
        mPractice.Enabled = True
        mUpdateConventionFiles.Enabled = True
    Else
        mPractice.Enabled = False
        mUpdateConventionFiles.Enabled = False
    End If
    
    f_Msg.Top = 480
    response = MyMsgBox("Ready to find matching deals?", 1, "yn", "")
    If response = "yes" Then
        mFindDeals_Click
    End If
End Sub


Private Sub b_MsgInput_Click(Index As Integer)
    int_MsgClick = Index
    f_Msg.Visible = False
End Sub


Private Sub b_Ok_Click()
    'If f_Msg.Visible Then
    '    f_Msg.SetFocus
    '    Exit Sub
    'End If
    
    str_MyBid = ""
    If b_Ok.Caption = "Deal" Then
        
        b_Ok.Caption = "&BB Bid"
        str_TimerCallingSub = "b_OK_Click"
        Timer1.Enabled = True
  
    ElseIf b_Ok.Caption = "&BB Bid" Then
        
        Open App.Path & "DealerPosition.txt" For Output As #5
        Print #5, Me.Left
        Print #5, Me.Top
        Close #5
        
        ProcessBid
        bln_SkipIt = False
        
    ElseIf b_Ok.Caption = "&Continue" Then
        ' Was waiting for user to '
        ' finish editing an entry.'
        b_Ok.Caption = "&BB Bid"
        b_Edit.Enabled = True
        b_ShowAnalysis.Enabled = True
        b_Deal.Enabled = True
        b_Dealer.Enabled = True
        b_Vul.Enabled = True
        b_RestartDeal.Enabled = True
        zF.Enabled = True
        mO.Enabled = True
  
        If Not bln_Analysis Then
            bln_Rebid = True
            BidNextHand
        End If
            
    End If
        
End Sub


Private Sub b_RestartDeal_Click()
    bln_OffPBN = False
    If ck_KeepDealer.Value = 0 And int_DealerSpecified <> int_Dealer Then
        int_Dealer = InStr("SWNE", L_Dealer)
        int_DealerSpecified = int_Dealer
        SetPlayerNumbers ' This is also done in FinishDealSetup
    End If
    bln_SkipIt = True ' Don't need to redo DDA. '
    FinishDealSetup
    bln_SkipIt = False
    mRedisplayHands_Click
End Sub


Private Sub b_ShowAnalysis_Click()
    Dim n As Integer
    
    If bln_PbnEvent Then
        t_Disclose = str_PBNanalysis
        Exit Sub
    End If
    
    If bln_Analysis Then
        b_ShowAnalysis.Caption = "Rebid w/ &Analysis"
        b_Ok.Caption = "&BB Bid"
        bln_Analysis = False
    
    Else
        str_BidToAnalyze = UCase$(InputBox$( _
            "Enter bid to watch for or ALL:"))
    
        If InStr(LCase$(str_BidToAnalyze), "all") > 0 Then
            ' Specifying anything else '
            ' with "all" is redundant. '
            str_BidToAnalyze = "ALL"
        'ElseIf str_BidToAnalyze = "" Then
        '    Exit Sub
        End If
        
        n = Val(str_BidToAnalyze)
        If n > 0 Then
            str_BidToAnalyze = _
                str_BidToAnalyze & " " & n & "x"
            ' Example: if "2S" is entered,'
            ' must also watch for "2x"    '
            ' because "2x" might be "2S". '
        End If
    
        'b_ShowAnalysis.Caption = "Stop &Analysis"
        str_Call(int_CallNum) = ""
        DisplayBid int_Player(int_Turn)
        bln_Analysis = True
        bln_Rebid = True
        BidNextHand
    End If
    
    b_Ok.Enabled = True
    b_Ok.SetFocus
End Sub


Private Sub b_Vul_Click()
    ' Cycle: '
    If L_Vul = "" Then
        L_Vul = "Unfav."
    ElseIf L_Vul = "Unfav." Then
        L_Vul = "Fav."
    ElseIf L_Vul = "Fav." Then
        L_Vul = "Both"
    Else
        L_Vul = ""
    End If
End Sub


Private Sub Card_Click(Index As Integer)
    Dim X As Integer
    Dim suit As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim cardNumInDeck As Integer
    Dim cardNumInHand As Integer
    Dim c As Integer ' card counter '
    Dim s As Integer ' suit counter '
    Dim n As Integer
    Dim p As Integer
    
    If int_HandQty(int_Turn) = 13 Then
        MsgBox "Hand already has 13 cards."
        Exit Sub
    End If
    
    n = 0 ' number of cards added to Hand. '
    
    p_Hand(1).Visible = True
    
    If int_Turn = 0 Then
        MsgBox "Click on N, S, E or W first."
        Exit Sub
    End If
    
    cardNumInDeck = Index ' Index = 1-52 '
    ' Clb=1-13,    Dia=14-26,    Hrt=27-39,    Spd=40-52    '
    ' AC=1, KC=13, AD=14, KD=26, AH=27, KH=39, AS=40, KS=52 '
    
    If bln_SelectedCards(cardNumInDeck) = True Then
        Exit Sub
    End If
    
    ' Keep track of which cards have been selected '
    ' where bln_Selected(1-52) is a Boolean array. '
    bln_SelectedCards(cardNumInDeck) = True
    
    DrawBack Card(cardNumInDeck).hwnd, 3, 0, 0
    
    ' Go through player's hand to find where the new card goes. '
    suit = GetCardSuit(cardNumInDeck)
    cardNumInHand = GetCardValue(cardNumInDeck)
    
    ' Put cardNum in player's hand (int_Turn). '
    ' For Dia. A64 where int_Turn=3 (North):   '
    '    int_Suit(3, Dia, 1) = 14 (Ace: 13+1)  '
    '    int_Suit(3, Dia, 2) = 19 (13+6)       '
    '    int_Suit(3, Dia, 3) = 17 (13+6)       '
    ' int_Turn = # where 1=S, 2=W, 3=N, 4=E '
    int_Suit(int_Turn, suit, cardNumInHand) = cardNumInDeck
    'int_SuitQty(int_Turn, suit) = int_SuitQty(int_Turn, suit) + 1
    int_HandQty(int_Turn) = int_HandQty(int_Turn) + 1
        
If 1 = 2 Then
    ' Rather than try to figure out where to insert  '
    ' a new card, it's easier just to redo int_Hand: '
    ' int_Hand(player#, card#1-13 in hand) is set to '
    ' the number of the card in the deck which is the'
    ' value of int_Suit(...). '
    For n = 1 To 13
        int_Hand(int_Turn, n) = 0
    Next
    
    n = 0 ' Counts cards added to the Hand. '
    For s = Spd To Clb Step -1
        ' In the hand, Spades come first, then '
        ' Hearts, then Diamonds and Clubs.     '
        ' Ace comes first in a suit, then the  '
        ' cards K (13) to 2. '
        If int_Suit(int_Turn, s, 1) > 0 Then
            ' Ace '
            n = n + 1
            int_Hand(int_Turn, n) = int_Suit(int_Turn, s, 1)
        End If
        For c = King To 2 Step -1
            If int_Suit(int_Turn, s, c) > 0 Then
                n = n + 1
                int_Hand(int_Turn, n) = int_Suit(int_Turn, s, c)
            End If
        Next
    Next
End If

    ShowPickedCardsInTextBox
    
    DisplayHand int_Turn
    
    If int_HandQty(int_Turn) = 13 Then
        t_EnterHand(int_Turn).BackColor = 13434828 ' green '
        n = 0
        For i = 1 To 4
            If int_HandQty(i) < 13 Then
                n = n + 1
                t_EnterHand(i).BackColor = 65535 ' yellow '
                int_Turn = i
                Exit For
            End If
        Next
        If n > 0 Then
            b_DealRest.Visible = True
            Dim r As String
            MsgBox "Click an input box" & vbCr & _
                         "to deal to or" & vbCr & _
                         "click Cancel or" & vbCr & _
                         "click Deal The Rest."
        Else
            b_DealRest.Caption = "DONE"
        End If
    End If
    
    bln_SkipIt = True ' Keep variables from being zeroed in t_EnterHand_Validate. '
End Sub


Sub ShowPickedCardsInTextBox()
    ' from Card_Click '
    
    SetUpHands int_Turn
    bln_SkipIt = True
    t_EnterHand(int_Turn) = str_Hands(int_Turn)
    bln_SkipIt = False
End Sub


Private Sub GetPracticeDeal()
    Dim fiNum As Integer
    Dim found As FoundDeal
    Dim lastDeal As lastDeal
    
    fiNum = FreeFile
    Close #fiNum
    Open str_PracticeDealsFile For Random As fiNum Len = 6
        If int_DealFileEntryNum < 3 Then
            Get #fiNum, 1, lastDeal
                lng_NumPracticeEntries = lastDeal.NumEntries
                int_DealFileEntryNum = lastDeal.LastNumUsed
        End If
        
        If int_DealFileEntryNum < 2 Then
            int_DealFileEntryNum = 2 ' Did this before '
        End If
        
        Get #fiNum, int_DealFileEntryNum, found
    
        If found.dealNum = 0 And lng_NumPracticeEntries > int_DealFileEntryNum Then
            ' Temporary fix for a couple of entries I zeroed out. '
            DelPracticeDeal
            Get #fiNum, int_DealFileEntryNum, found
        End If
    
        int_Dealer4PracticeDeal = found.Dealer
            int_Dealer = int_Dealer4PracticeDeal
            If int_Dealer = 0 Then int_Dealer = 1
            L_Dealer = Mid$("SWNE", int_Dealer, 1)
            int_DealerSpecified = int_Dealer
        lng_DealNum2Practice = found.dealNum
    Close #fiNum
            
    ' Save the # of entries & last entry used in first record: '
    SaveLastFileEntryNumberUsed
    
    bln_FindAuctions = False
    bln_FindAuctionsDone = False
    bln_SettingUpPracBid = False
    bln_SearchForMoreDeals = False
    bln_PracticeFoundBids = True
End Sub


Private Sub PutPracticeDeal()
    Dim fiNum As Integer
    Dim found As FoundDeal

    If int_DealFileEntryNum < 1 Then int_DealFileEntryNum = 1
    fiNum = FreeFile
    Open str_PracticeDealsFile For Random As fiNum Len = 6
        found.dealNum = lng_DealNum2Practice
        found.Dealer = int_Dealer
        Put #fiNum, int_DealFileEntryNum, found
    Close fiNum
End Sub


Private Sub DelPracticeDeal()
    Dim fiNum As Integer
    Dim found As FoundDeal
    Dim lastDeal As lastDeal
    Dim i As Integer
    
    ' A record cannot actually be removed from a random access file. '
    ' Since the order of deal numbers in the practice files does not '
    ' matter, I copy the last record in the file over the record that'
    ' I want to "delete" and then reduce the lng_NumPracticeEntries  '
    ' variable by one and save it in record #1 in the file. '
    fiNum = FreeFile
    Open str_PracticeDealsFile For Random As fiNum Len = 6
        Get #fiNum, int_DealFileEntryNum, found
        i = MsgBox("Delete deal #" & found.dealNum & " from the practice file?", vbYesNo)
        If i = vbNo Then Close fiNum: Exit Sub
        ' Get the last record in the file: '
        Get #fiNum, lng_NumPracticeEntries, found
        ' Write it over the file being deleted. '
        Put #fiNum, int_DealFileEntryNum, found
        ' Decrement the variable for keeping track of the # of entries. '
        lng_NumPracticeEntries = lng_NumPracticeEntries - 1
    Close fiNum
    SaveLastFileEntryNumberUsed
End Sub


Sub SaveLastFileEntryNumberUsed()
    ' This is used for Practice files. '
    Dim lastDeal As lastDeal
    Dim fiNum As Integer
    
    fiNum = FreeFile
    Open str_PracticeDealsFile For Random As fiNum Len = 6
        lastDeal.NumEntries = lng_NumPracticeEntries
        lastDeal.LastNumUsed = int_DealFileEntryNum
        Put #fiNum, 1, lastDeal
    Close fiNum
End Sub


Private Sub cmb_FileList_Click()
    Dim response As String
    
    If cmb_FileList.Text = "" Or cmb_FileList.ListIndex = 1 Then
        str_PracticeSpecsFile = ""
        str_PracticeDealsFile = ""
    Else
        str_PracticeSpecsFile = cmb_FileList.Text
        str_PracticeDealsFile = str_PracticeSpecsFile & ".Deals"
    End If
    cmb_FileList.Clear
    cmb_FileList.Visible = False
End Sub


Private Sub cmb_FindDeals_DblClick()
    Dim i As Long
    Dim a As String
    
    a = MsgBox("Delete?", vbYesNo)
    If a = vbYes Then
        cmb_FindDeals.RemoveItem cmb_FindDeals.ListIndex
    End If
    
        If cmb_FindDeals.ListCount < 6 Then
            cmb_FindDeals.Height = 95
            Me.Height = 10770
        ElseIf cmb_FindDeals.ListCount = 6 Then
            cmb_FindDeals.Height = 110
            Me.Height = 10970
        Else
            cmb_FindDeals.Height = 125
            Me.Height = 11140
        End If
End Sub


Private Sub cmb_FindDeals_GotFocus()
    b_Ok.Default = False
    If cmb_FindDeals.ListCount = 0 Then Exit Sub
    
    Dim i As Integer
    For i = cmb_FindDeals.ListCount - 1 To 0 Step -1
        If Trim$(cmb_FindDeals.List(i)) = "" Then
            cmb_FindDeals.RemoveItem i
        End If
    Next
End Sub


Function RunningInIDE() As Boolean
'Returns whether we are running in vb(true), or compiled (false)
 
    Static counter As Variant
    If IsEmpty(counter) Then
        counter = 1
        Debug.Assert RunningInIDE() Or True
        counter = counter - 1
    ElseIf counter = 1 Then
        counter = 0
    End If
    RunningInIDE = counter
 
End Function


Private Sub cmb_PBNBoards_DblClick()
    int_PBNlistNum = cmb_PBNBoards.ListIndex
    cmb_PBNBoards.Visible = False
    GetPbnDeal
    t_Calculating.Visible = False
End Sub


Private Sub Form_Load()
    Dim n As Integer
    
    'If RunningInIDE Then
    '    str_BBpath = "C:\Programming\VB98\Bridge"
    'Else ' Running the exe. '
    '    str_BBpath = "C:\Program Files (x86)\CardShark BidBase"
    'End If
    
    'If Dir$("C:\Program Files (x86)\CardShark BidBase\*") = "" Then
    '    MsgBox "BidBase should have been installed in" & vbCr & _
    '           "C:\Program Files (x86)\CardShark BidBase" & vbCr & _
    '           "but that path isn't there. Program ending."
    '    End
    'End If

    str_BBpath = App.Path
    str_DBPathAndName = App.Path & "\BidBase.accdb"
    'str_DBPathAndName = "C:\Program Files (x86)\CardShark BidBase\BidBase.accdb"
    
    ChDir str_BBpath
    
    If Dir$("Last Deal Number.txt") <> "" Then
        Open "Last Deal Number.txt" For Input As #1
            Input #1, lng_DealNum
            lng_DealNum = lng_DealNum - 1
        Close #1
    End If
    
    If Dir$(str_DBPathAndName) = "" Then
        MsgBox "The BidBase database (BidBase.accdb) is not in" & vbCr & _
               "C:\Program Files (x86)\CardShark BidBase"
        Form_Unload 0
    End If
    
    ' Only let the program load once: '
    If App.PrevInstance Then
        MsgBox "Program is already running."
        End
    End If
    
    ' Initialize DLL that displays cards: '
    'On Local Error GoTo 0
'    On Local Error GoTo CantLoadDLL
    n = InitializeDeck(PlayScr.hwnd)
    On Local Error GoTo 0
    If n = False Then
        GoTo CantLoadDLL
    End If
    
    LoadCfg
    
    'MakeTopMost hwnd
    
    Me.Caption = Me.Caption & "     ver." & Version$
    
    SetUpBiddingGrid
    
    'If bln_FindAuctions Then
        ' Get starting bids specified: '
        'ck_FindBids.Value = 1
        'ck_FindBids_Click
        'cmb_FindDeals_LostFocus
    'End If
    
    Me.Refresh
    
    ReDim str_Hands(4)
    str_Hands(1) = "Clubs"
    str_Hands(2) = "Diamonds"
    str_Hands(3) = "Hearts"
    str_Hands(4) = "Spades"
    
    Deal
    
    For n = 1 To 4
        DisplayHand n
        int_HandQty(n) = 13
    Next
    
    If mCreateConv.Checked = False Then 'ck_FindBids.Value = 0 Then
        ShowDDA False
    End If
        
    'If b_CancelHands.Visible Then Exit Sub
    
    InitiateBidding
    
    'mAverageDDA_Click
    
    mStayOnTop_Click

Exit Sub
  
CantLoadDLL:
    MsgBox "Problem loading QCard32.DLL." & vbCr & _
           "Make sure it is in the BidBase directory." & vbCr & _
           "If it is, then try restarting this program." & vbCr & _
           "If it still won't load, restart Windows."
    
    Unload PlayScr
End Sub


Private Sub Form_LostFocus()
    Stop
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim fiNum As Integer
    Dim i As Integer
    
    Set NewBid = Nothing
    
    fiNum = FreeFile
    Open str_BBpath & "\BBPractice.cfg" For Output As #fiNum
        Print #fiNum, lng_DealNum
        Print #fiNum, PlayScr.Left, PlayScr.Top
        Write #fiNum, Not bln_AutoBid
        Write #fiNum, Not bln_HideHands
        Print #fiNum, InStr("SWNE", L_Dealer)
        Print #fiNum, ck_KeepDealer.Value
        Write #fiNum, str_PBNeventFile ' This should be blank if the last deal was not from a pbn file.
        Print #fiNum, int_PbnDealNum
    Close #fiNum
        
    ' AutoBid and HideHands are stored in '
    ' reverse because when the program    '
    ' starts, mAutoBid and mHide_Click are'
    ' run, which reverses their settings. '
    End
End Sub


Private Sub Form_Activate()
    Form_Paint
End Sub


Private Sub Form_Click()
    Form_Paint
End Sub


Private Sub Form_GotFocus()
    Form_Paint
End Sub


Private Sub Form_Paint()
    'Dim t As Long
    Dim h As Integer

    If bln_NoPaint Or bln_Abort Or Me.WindowState = 1 Then Exit Sub
    bln_NoPaint = True
    
    If L_Deal <> "" And Card(1).Visible = False Then
        For h = 1 To 4
            DisplayHand h
            DoEvents
        Next
    End If
    
    If mStayOnTop.Checked Then
        SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
        DoEvents
        SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
        DoEvents
    End If
    
    bln_NoPaint = False
End Sub


Private Sub Form_Resize()
    'If Me.WindowState = 0 Then
        bln_NoPaint = False
        Form_Paint
    'End If
End Sub


Private Sub g_BidBox_Click()
    Dim b As String
    Dim q As String
    Dim X As Long
    Dim i As Integer
    
    If g_BidBox.row = 7 Then
        If g_BidBox.col > 2 Then Exit Sub
        ' Pass, Dbl, Redbl - just use 1st letter: '
        str_MyBid = Left$(Trim$(g_BidBox.Text), 1)
    Else ' Other bids, use whatever's there: '
        str_MyBid = Trim$(g_BidBox.Text)
    End If
        
    If str_Call(int_CallNum) <> "" And _
       str_Call(int_CallNum) <> str_MyBid _
    Then
        ' Changing the computer's bid. '
        i = MsgBox("Do you want to add your bid of " & _
                    str_MyBid & vbCr & vbCr & _
                   "to BidBase to supersede BB's bid of " & _
                    str_Call(int_CallNum) & "?", vbYesNo)
        If i = vbYes Then
            str_AllCalls = Replace(str_AllCalls, "-" & str_Call(int_CallNum), "")
            AddBidToBB
            str_Call(int_CallNum) = str_MyBid
            DisplayBid 1
            str_AllCalls = str_AllCalls & "-" & str_MyBid
            Exit Sub
        End If
    End If
    
    Do While int_Turn > 1
        If str_Call(int_Turn - 1) = "" Then
            int_Turn = int_Turn - 1
        Else
            Exit Do
        End If
    Loop
    'int_CallNum = int_Turn
    ' The above line is wrong. int_CallNum is the '
    ' number of calls which could be as many as 7.'
    ' int_Turn is whose turn it is: 1-4.          '
    ' Instead, get the next unbid call number.    '
    Do While str_Call(int_CallNum) <> ""
        int_CallNum = int_CallNum + 1
    Loop
    
    ' While practicing, when user clicks a bid and it's '
    ' BB's turn, the bid for int_Player(int_Turn) must  '
    ' be blank. '
    int_Turn = int_CallNum Mod 4
    If int_Turn = 0 Then int_Turn = 4
    
    MakeBid int_Player(int_Turn)

    If _
       ((bln_FindAuctions Or bln_PracticeFoundBids Or 1 = 1) And _
        str_Call(int_CallNum) <> str_MyBid) _
    Then
        t_Disclose = ""
        q = "Did you want to change BidBase's bid of " & str_Call(int_CallNum)
        
        q = q & vbCrLf & " to your bid of " & str_MyBid
        q = q & " in the database?" & vbCr
        q = q & "(Click No to keep your bid for now.)"
        q = MyMsgBox(q, 3, "YNC", "")
        If q = "yes" Then
            b_Edit_Click
            Exit Sub
        ElseIf q = "cancel" Then
            Exit Sub
        ElseIf q = "no" Then
            str_Call(int_CallNum) = str_MyBid
        End If
    Else
        t_Disclose = "         *** Your bid is the same as BB's. ***" & vbCrLf & vbCrLf & _
                     "            Click a bid for the next player or" & vbCrLf & _
                     "            click ""BB Bid"" for BB to do it" & vbCrLf & _
                     "            or click right arrow for next deal." & vbCrLf & vbCrLf & _
                     t_Disclose
    End If
    
    DisplayBid 1
End Sub


Private Sub AddBidToBB()
    With NewBid
        .PlayersHand = str_Hands(int_Player(int_Turn))
        .PriorCalls = str_AllCalls
        .SaveDealNum = lng_DealNum
        .SaveBid = str_MyBid
        .Job = "Save"
        ' The bid is saved in the Bidder module '
        ' because it already has the evaluation '
        ' and database access routines. '
    End With
End Sub


Private Sub m100DealAvg_Click()
    m100DealAvg.Checked = True
    AverageDDA 110
    m100DealAvg.Checked = False
End Sub


Private Sub mAutoBid_Click()
    ' Default is checked so that the '
    ' computer does all the bidding. '
    ' When unchecked, the user bids  '
    ' first and then the computer    '
    ' bids and compares the bids.    '
    bln_AutoBid = Not bln_AutoBid
    mAutoBid.Checked = bln_AutoBid
    mHide.Enabled = Not bln_AutoBid
    mHide.Checked = False
    If bln_AutoBid Then
        bln_HideHands = False
    End If
End Sub


Private Sub mAverageDDA_Click()
    Dim n As Integer
    Dim i As Long
    
    mAverageDDA.Checked = Not mAverageDDA.Checked
    m100DealAvg.Enabled = mAverageDDA.Checked
    
    If mAverageDDA.Checked Then
        AverageDDA 11
        ' Deal #27115 wrong  - shows 12 with 2 C losers. '
    Else
        For n = 1 To 5
            cmb_DDA.RemoveItem cmb_DDA.ListCount - 1
        Next
        cmb_DDA.Height = 90
        Me.Height = 9330
    End If
End Sub


Private Sub mCreateConv_Click()
    Dim p As String
    Dim f As String
    Dim i As Integer
    Dim dealNum As Long
    Dim dealerNum As String
    Dim numSpecs As Integer
    Dim a As String
    Dim fiNum As Long
    Dim found As FoundDeal
    Dim lastDeal As lastDeal
    Dim specsPath As String
    Dim dealsPath As String
    Dim response As String
    
    If mCreateConv.Checked Then
        mCreateConv.Checked = False
        ClearPracticeVariables
        Me.Height = 10770
        Exit Sub
    End If
    
    If mFindDeals.Checked Then
        response = MyMsgBox("Stop searching for matching deals first" & vbCr & _
                            "by clicking the Find Matching Deals menu.", 2, "ok", "")
        Exit Sub
    End If
    
    If mPractice.Checked Then
        response = MyMsgBox("Stop practicing a convention first by" & vbCr & _
                            "clicking the Practice Convention menu.", 2, "ok", "")
        Exit Sub
    End If
                  
    OpenPracticeFile
    Do While cmb_FileList.Visible = True
        DoEvents
    Loop
            
    If str_PracticeSpecsFile = "" Then
        'MsgBox "No convention name entered."
        Exit Sub
    End If
    
    On Local Error GoTo CreateFileSpecsError
    
    ' Initialize the specs file: '
    ChDir str_BBpath
    ' Set up specs list: '
    cmb_FindDeals.Clear
    cmb_FindDeals.Text = "Enter an auction here or click an auction to edit or click Finished."
    cmb_FindDeals.ZOrder 0
    cmb_FindDeals.Visible = True
    cmb_FindDeals.Height = 90
    Me.Height = 10770
    cmb_FindDeals.SetFocus

    'If Dir$(str_PracticeSpecsPath & "*", vbDirectory) = "" Then
    '    MkDir str_PracticeSpecsPath
    'End If
    
    'str_PracticeDealsPath = str_PracticeSpecsPath & "Deals\"
    'If Dir$(str_PracticeDealsPath & "*", vbDirectory) = "" Then
    '    MkDir str_PracticeDealsPath
    'End If

    If Dir$(str_PracticeSpecsPath & str_PracticeSpecsFile) <> "" Then
        fiNum = FreeFile
        On Local Error GoTo 0
        Open str_PracticeSpecsPath & str_PracticeSpecsFile For Input As #fiNum
            Do While Not EOF(fiNum)
                Line Input #fiNum, a
                cmb_FindDeals.AddItem a
            Loop
        Close #fiNum
    Else
        cmb_FindDeals = str_PracticeSpecsFile & " - " & cmb_FindDeals
    End If
    
    mCreateConv.Checked = True
    bln_SettingUpPracBid = True
    b_FinishedAuctions.Visible = True

    f_Msg.Top = 538
    Me.Height = 10830
    bln_SkipIt = True  ' Keeps t_Calculating from hiding the Finished Auctions button. '
    'If cmb_FindDeals.ListCount = 0 Then
    '    response = MyMsgBox("лл Enter auctions at left then" & vbCr & _
    '                        "click Finished Auctions.^", 2, "ok", "")
    'Else
    '    response = MyMsgBox("<- Enter new auctions at left or" & vbCr & _
    '                        "click an auction to edit it then" & vbCr & _
    '                        "click Finished Auctions above.", 3, "ok", "")
    'End If
    Me.Height = 10770
    f_Msg.Top = 200
    bln_SkipIt = False
    cmb_FindDeals.SetFocus
Exit Sub

CreateFileSpecsError:
    MsgBox "Error creating practice file" & vbCr & f & vbCr & Err.Description
    bln_SettingUpPracBid = False
    On Local Error GoTo 0
    Exit Sub
End Sub


Public Sub mEnterHands_Click()
    ' Lets the user manually enter one or more hands. '
    ' When we prompt the user for a hand, the default '
    ' is the last set of cards for that
    Dim h As Integer
    Dim hnd As String
    Dim m As String
    Dim b As String
    
    b_EnterHands_Click
Exit Sub
    
    h = MsgBox("Clear previous hands?", vbYesNoCancel)
    If h = vbCancel Then
        Exit Sub
    ElseIf h = vbYes Then
        Erase str_Hands()
    End If
    
    For h = 1 To 4
      
        Do
            hnd = Trim$(UCase$(InputBox$( _
                   "Enter " _
                   & str_Dealer(h) _
                   & "'s hand:" & vbCr _
                   & "(Enter hands as: " & vbCr _
                   & " AKT3-T92--QJT532" & vbCr _
                   & " in suit order S-H-D-C.)" _
                   , "Enter Hands or 'Cancel':" _
                   , str_Hands(h))))
                   ' Default is last hand used.
                        
            If hnd = "CANCEL" Then Exit Sub
            
            ' Check for errors: '
            m = ""
            If hnd <> "" Then
                m = ValidateHand(h, hnd)
                If m <> "" Then 'error message returned
                    MsgBox m
                End If
            End If
        Loop While m <> ""
                
        str_Hands(h) = hnd
      
        If hnd = "" Then
            
            bln_PickedHand(h) = True
            int_NumPickedHands = _
              int_NumPickedHands + 1
    
        Else
      
            bln_PickedHand(h) = False
    
            AddHandToDeal str_Hands(h), h
            ' This enters the card value for each card   '
            ' entered by user into the int_Cards() array '
            ' which holds the whole deal. It also puts it'
            ' into int_Suit(0,s,c) so that we can easily '
            ' tell which cards have been used.           '
      
        End If
    
        'If int_NumPickedHands = 0 And h = 3 Then
            ' User has entered 3 hands. '
            ' We can fill in the fourth.'
        '    bln_PickedHand(4) = True
        '    int_NumPickedHands = 1
        '    Exit For
        'End If
    
    Next
    
    If int_NumPickedHands > 1 Then
  
        FillInTheDeal ' Puts remaining cards for the '
                      ' random hands into int_Cards()'
                      
    End If
                  
    ScrambleHands ' Mix up cards in int_Cards()  '
                  ' for the random hands.        '
                  
    SortHands ' Sort hands within int_Cards().   '
    
    FillSuits ' Assign int_Cards() to int_Suit().'
        
    'If Not bln_FindAuctions Or bln_FindDealsDone Then
        For h = 1 To 4
            DisplayHand h
        Next
    'End If
    
    ShowDDA False
    
    GetDealer
    
    SetUpBidDisplayGrid
    
    ' Get source of the deal: '
    If str_Source = "" Then
        str_Source = "User-entered hand"
    End If
    m = "Enter source of hand(s):" & vbCr _
        & "BW,mm-yy,p## = Bridge World, date, page" & vbCr _
        & "Bul,mm-yy,p## = ACBL Bulletin" & vbCr _
        & "[Book Title],p## = any book" & vbCr _
        & "[bridge col. name],mm-dd-yy = bridge column, date"
    str_Source = InputBox$(m, , str_Source)
  
    ' Get prior calls: '
    str_AllCalls = _
        UCase$(InputBox$( _
                "Enter prior bids like 1C-P-1S-D-2C:"))
    PutCallsIntoArray
        
    L_Deal = "User"
    b_RestartDeal_Click
End Sub


Private Function _
    ValidateHand( _
        ByVal h As Integer, _
        hnd As String) _
    As String
    '*************************************'
    ' Input is a hand entered by the user '
    ' with a format like xxx-xxxxxx-xxx-x.'
    ' "ValidateHand" will be returned as  '
    ' "" if hand is ok; otherwise, it will'
    ' be set to an error message.         '
    '*************************************'
    
    Dim lh As Integer
    Dim s As Integer
    Dim i As Integer
    Dim c As Integer
    Dim n As Integer
    Dim m As String
    Dim lastC As Integer
    Dim numDashes As Integer
         
    lh = Len(hnd)
    If lh = 0 Then
        ValidateHand = "No hand was entered."
        Exit Function
    
    ElseIf lh < 16 Then ' 13 cards + 3 dashes '
        ValidateHand = "Hand has fewer than 13 cards."
        Exit Function
        
    ElseIf lh > 16 Then
        ValidateHand = "Hand has more than 13 cards."
        Exit Function
    
    End If
    
    s = Spd ' Hand entered starts with Spades.'
            
    lastC = 15 ' For validity test, current card '
               ' must be lower than the last.    '
            
    For i = 1 To Len(hnd)
        
        ' Check for the dash between suits: '
        Do
            m = Mid$(hnd, i, 1)
            If m = "-" Then
                numDashes = numDashes + 1
                If numDashes > 3 Then
                    ValidateHand = "Hand has more than 3 dashes."
                    Exit Function
                End If
                
                If s > 1 Then
                    s = s - 1
                    i = i + 1
                    lastC = 15
                End If
            End If
            ' Loop, in case of adjacent    '
            ' dashes like A72--QJ5432-8632 '
        Loop While m = "-"
                
        ' See if a valid character: '
        c = Val(m)
        If c < 2 Then ' Must be a letter, not a number. '
            c = 9 + InStr("TJQKA", m)
            If c = 9 Then ' Invalid character. '
                ValidateHand = m _
                  & " is an invalid character in the suit of " _
                  & str_Suits(s)
                Exit Function
            End If
        End If
                
        ' Compare to last character:
        If c >= lastC Then
            ValidateHand = m _
              & " is out of order in the suit of " _
              & str_Suits(s)
            Exit Function
        End If
        
        ' See if card has already been used: '
        If c = 14 Then n = 1 Else n = c ' Convert Ace
        If int_Suit(0, s, n) > 0 Then
            ValidateHand = m _
              & " of " & str_Suits(s) & _
              " has already been used."
            Exit Function
        End If
                        
        ' This card is okay. Set lastC to '
        ' this card for next comparison:  '
        lastC = c
    Next
End Function


Private Sub FillInTheDeal()
    '****************************'
    ' Fill in int_Cards() for the'
    ' empty hands when picking.  '
    '****************************'
    ' Other than hands entered,  '
    ' the order of cards doesn't '
    ' matter because they will   '
    ' be scrambled next anyway.  '
    '****************************'
    
    Dim n As Integer
    Dim h As Integer
    Dim i As Integer
    Dim s As Integer
    Dim c As Integer
    Dim k As Integer
    
    n = 1
    h = 1
    Do While bln_PickedHand(h) = False
        h = h + 1
    Loop
    ' Has to be at least one random hand  '
    ' because if user enters 1st 3 hands, '
    ' the program marks the 4th random to '
    ' be filled in here.                  '
    
    For i = 1 To 52
        ' Calculate the suit # and card # '
        ' for use in assigning to array.  '
        s = GetCardSuit(i)
        c = GetCardValue(i)
        
        ' If not already used... '
        If int_Suit(0, s, c) = 0 Then
            ' Keep track of cards used: '
            int_Suit(0, s, c) = i
            int_SuitQty(0, s, c) = int_SuitQty(0, s, c) + 1
            ' Assign to the "Deck", int_Cards(): '
            k = (h - 1) * 13 + n ' pointer for array '
            int_Cards(k) = i
            ' Increment card counter for this hand: '
            n = n + 1
            If n = 13 And h < 4 Then
                ' This hand is full, get next: '
                Do
                    h = h + 1
                Loop While bln_PickedHand(h) = False
                n = 0
            End If
        End If
    Next

End Sub


Private Sub PutCallsIntoArray()
    ' Takes calls/bids from a string  '
    ' like P-1C-1N and puts each call '
    ' into the str_Call() array.      '
    '                                 '
    ' For computing a string from the '
    ' str_Call() array, see Function  '
    ' AllCalls.                       '
    
    Dim b As String
    Dim n As Integer
    Dim i As Integer
    Dim int_NumPasses As Integer
    
    Erase str_Call()
  
    If b <> "" Then
        n = 1
        i = InStr(b, "-")
    
        Do While i > 0
            str_Call(n) = Left$(b, i - 1)
            
            If str_Call(n) = "P" Then
                int_NumPasses = int_NumPasses + 1
            Else
                int_NumPasses = 0
            End If
            
            b = Mid$(b, i + 1)
            i = InStr(b, "-")
            n = n + 1
        Loop
    
        str_Call(n) = b
    
        If str_Call(n) = "P" Then
            int_NumPasses = int_NumPasses + 1
        Else
            int_NumPasses = 0
        End If
        
        int_CallNum = n
  
    Else
        int_CallNum = 0
    End If

End Sub


Private Sub mExit_Click()
    Form_Unload 0
End Sub


Private Sub mFindDeals_Click()
    Dim fiNum As Integer
    Dim f As String
    Dim lastDeal As lastDeal
    Dim hiNum As Long
    Dim bid2 As Integer
    Dim response As String

    If mFindDeals.Checked Or str_PracticeSpecsFile = "" Then
        mFindDeals.Checked = False
        response = MyMsgBox("Practice specs were" & vbCr & "not selected.", 2, "ok", "")
        Exit Sub
    End If
    
    'ClearPracticeVariables
    
    mFindDeals.Checked = True
    bln_FindAuctions = True
    

    'OpenPracticeFile
    L_Billboard = str_PracticeSpecsFile & " will have 100 deals added to it. " & _
                  "More deals can be added in the future by clicking this menu function. " & _
                  "To stop adding deals at any time, click the menu again." & _
                  Chr$(13) & Chr$(13) & _
                  "        Click Ok below to continue."
    f_BillBoard.Visible = True
    b_Billboard.SetFocus
    Do While f_BillBoard.Visible
        DoEvents
    Loop
           
    bln_SettingUpPracBid = False
    bln_SearchForMoreDeals = True
    bln_PracticeFoundBids = False

    If Dir$(str_PracticeDealsPath & str_PracticeDealsFile) = "" Then
        fiNum = FreeFile
        int_DealFileEntryNum = 1
        Open str_PracticeDealsPath & str_PracticeDealsFile For Random As fiNum Len = 6
            lastDeal.NumEntries = 0   ' Record # for last file record used. '
            lastDeal.LastNumUsed = 0  ' Record # for the last file record. '
            Put #fiNum, 1, lastDeal
        Close #fiNum
        lng_DealNum = 1
    ElseIf FileLen(str_PracticeDealsPath & str_PracticeDealsFile) < 10 Then
        int_DealFileEntryNum = 1
        lng_DealNum = 1
    Else
        ' Get the high deal number (hiNum) in the file. '
        'lng_DealNum = hiNum + 1
        Stop ' See what's going on next. '
    End If

    LoadPracticeDealsAuctions
    
    ' Starting with deal #1 BB will   '
    ' continue checking each deal for '
    ' one matching the specified bids:'
    L_Deal = "" & lng_DealNum
    lng_DealNumSpecified = lng_DealNum
    
    int_Dealer = 1
    L_Dealer = "S"
    int_DealerSpecified = 1
    
    ' Deal the first time: '
    bln_HandsDealt = False ' Single-step and find out where deal num is going to 2 from 1. '
    
FindPracticeDeals
Stop
    Deal
    'mRedisplayHands_Click ' Already done. '
    str_TimerCallingSub = "mFindDeals"
    str_TimerJob = "Restart same deal"
    Timer1.Enabled = True
End Sub


Private Sub LoadPracticeDealsAuctions()
    Dim fiNum As Integer
    
    If Dir$(str_PracticeDealsPath & str_PracticeDealsFile) = "" Then
        MsgBox "The auctions file was not created for " & str_PracticeSpecsFile
        Exit Sub
    End If
    
    fiNum = FreeFile
    ' The convention name file contains all '
    ' of the specified auctions like this:  '
    '   1C-P-1H-P-1N-P-2D-P     '
    '   1C-P-1S-P-1N-P-2D-P     '
    '   1D-P-1H-P-1N-P-2C-P     '
    '   1D-P-1S-P-1N-P-2C-P     '
    
    ' Load the specified file's auctions '
    ' into the auction list box: '
    If cmb_FindDeals.ListCount = 0 Then
      Open str_PracticeSpecsPath & str_PracticeSpecsFile For Input As #fiNum
        int_NumAuctionSpecs = 0
        cmb_FindDeals.Clear
        cmb_FindDeals.Text = str_PracticeSpecsFile
        
        Do
            int_NumAuctionSpecs = int_NumAuctionSpecs + 1
            Line Input #fiNum, str_FindAuctions(int_NumAuctionSpecs)
            cmb_FindDeals.AddItem str_FindAuctions(int_NumAuctionSpecs)
            If EOF(fiNum) Or str_FindAuctions(int_NumAuctionSpecs) = "" Then
                Exit Do
            End If
            DoEvents
        Loop
      Close #fiNum
    End If
    
    If int_NumAuctionSpecs < 6 Then
        cmb_FindDeals.Height = 95
        Me.Height = 10770
    ElseIf int_NumAuctionSpecs = 6 Then
        cmb_FindDeals.Height = 110
        Me.Height = 10970
    Else
        cmb_FindDeals.Height = 125
        Me.Height = 11140
    End If
End Sub


Private Sub mHelp_Click()
    Dim f1 As String
    Dim xx As Long
    
    If Dir$(str_BBpath & "\Notes\*") = "" Then
        MsgBox "The Notes directory is not present." & vbCr & _
                "Download the BidBase Notes.Zip file" & vbCr & _
                "or go to www.aeyec.com/BidBase."
        Exit Sub
    End If
    
    f1 = str_BBpath & "\Notes\CardShark_BidBase_Practice_Program.htm"
    xx = Shell(str_BBpath & "\BidBase File Viewer.exe " & f1, vbNormalFocus)
End Sub


Private Sub mHide_Click()
    Dim h As Integer
    
'    If bln_AutoBid Then
'        bln_HideHands = False
'        mHide.Checked = False
'    Else
        bln_HideHands = Not bln_HideHands
        mHide.Checked = bln_HideHands
        If bln_NoPaint Then
            bln_NoPaint = False
            Exit Sub
        End If
'    End If
    
    If L_Deal <> "" Then
        For h = 1 To 4
            DisplayHand h
        Next
    End If
End Sub

Private Sub mLoadDeal_Click()
    Dim X As Integer
    Dim f As String
    Dim i As Integer
    
    X = MsgBox("The loaded deal will replace the current deal. Continue?", vbYesNo)
    If X = vbNo Then Exit Sub
    
    lng_CurrentDeal = Val(lng_DealNum)
    
    If Dir$(str_BBpath & "\UserEnteredHands\..\*", vbDirectory) = "" Then
        MkDir str_BBpath & "\UserEnteredHands"
    End If
    With CmDia
        .DialogTitle = "Enter file name:"
        .Filter = "*.txt"
        .InitDir = str_BBpath & "\UserEnteredHands\"
        .FileName = ""
        .DialogTitle = "Load User-Entered Deal"
        .ShowSave
        
        If .FileName = "" Then Exit Sub
        f = .FileName
    End With
    
        i = InStrRev(f, "\")
        f = Mid$(f, i + 1)
        i = InStr(f, ".txt")
        If i > 1 Then f = Left$(f, i - 1)
        L_Deal = "User - " & f
    LoadDeal f & ".txt"
    
    SetUpHands 0
    MsgBox "Enter dealer and vulnerability" & Chr$(13) & "then click Repeat, if necessary."
    bln_UserDeal = True
    b_DealBackFwd(0).Enabled = False
    b_DealBackFwd(1).Enabled = False
    t_Calculating.Visible = True
    t_Calculating.ZOrder 0
    cmb_DDA.Clear
    ShowDDA 0
    t_Calculating.Visible = False
End Sub

Private Sub LoadDeal(f As String)
    Dim c As Integer
    Dim h As Integer
    Dim s As Integer
    Dim cardNumInDeck As Integer
    Dim cardNumInHand As Integer
    Dim suit As Integer
    Dim i As Integer
    Dim d As String
    
        ClearPracticeVariables
        If f = ".txt" Then Exit Sub
            
        f = str_BBpath & "\UserEnteredHands\" & f
        If Dir$(f) = "" Then Exit Sub
        
'        On Local Error GoTo LoadDealError
        
        Open f For Input As #1
        ' Each line/record in the file is a number from 1 to 52   '
        ' The first 13 are for South, then West, North, and East. '
        For h = 1 To 4
            s = 4
            For c = 1 To 13
                Input #1, cardNumInDeck
                suit = Fix((cardNumInDeck - 1) / 13) + 1
                If suit <> s Then
                    str_Hands(h) = str_Hands(h) & "-"
                    s = suit
                End If
                str_Hands(h) = str_Hands(h) & Mid$("A23456789TJQK", GetCardValue(cardNumInDeck), 1)
                int_Hand(h, c) = cardNumInDeck
                cardNumInHand = cardNumInDeck - (suit - 1) * 13
                int_Suit(h, suit, cardNumInHand) = cardNumInDeck
                int_SuitQty(h, suit) = int_SuitQty(h, suit) + 1
                int_HandQty(h) = int_HandQty(h) + 1
                int_Cards((h - 1) * 13 + c) = cardNumInDeck
            Next
            DisplayHand h
        Next
        Close #1
        On Local Error GoTo 0
        Exit Sub

LoadDealError:
    MsgBox "Error loading " & f
    On Local Error GoTo 0
End Sub


Private Sub mLoadPBNDealNum_Click()
    Dim d As Integer
    Dim response As String
    Dim i As Integer
    
    With cmb_PBNBoards
    ''''''''''''''''''
    .Visible = True
    .ZOrder 0
    int_PBNfileLineNum = .ListIndex + 1
    ''''''''
    End With
End Sub


Private Sub mLoadPbn_Click()
    Dim i As Integer
    Dim n As Integer ' First line of a board. '
    Dim X As Integer
    Dim a As String
    Dim fiNum As Integer
    Dim brd As Integer
    Dim previousBlankLine As Integer
    Dim r As String
    
    If Dir$(str_BBpath & "\PBNs\\..\*", vbDirectory) = "" Then
        MkDir str_BBpath & "\PBNs"
    End If
    With CmDia
        .DialogTitle = "Enter file name:"
        .Filter = "*.pbn"
        .InitDir = str_BBpath & "\PBNs\"
        .FileName = ""
        .DialogTitle = "Load Event .pbn"
        .ShowSave
        
        If .FileName = "" Then Exit Sub
        str_PBNpathFile = .FileName
    End With
    
    lng_CurrentDeal = Val(lng_DealNum)
    bln_PbnEvent = False
    
    f_Msg.Top = 360
    a = MyMsgBox("The deal to be loaded will" & vbCr & _
                 "replace the current deal." & vbCr & _
                 "Continue?", 3, "YN", "Loading PBN")
    If a = "no" Then Exit Sub
       
    ClearPracticeVariables
    int_PbnDealNum = 1
    
    Close #fiNum
    fiNum = FreeFile
    Open str_PBNpathFile For Input As #fiNum
    int_PBNfileLineNum = 0
    
    ' Get file name out of the full path\filename: '
    i = InStrRev(str_PBNpathFile, "\")
    If i = 0 Then MsgBox "Error PBN-1"
    
    str_PBNfile = Mid$(str_PBNpathFile, i + 1)
    str_PBNfile = Replace(str_PBNfile, ".pbn", "")
        
    cmb_PBNBoards.Clear ' Clears the list box which holds all the boards. '
    
    ' Load all the deal numbers into the list box: '
    Do
        Line Input #1, a
        If Len(a) > 1000 Then
            r = MyMsgBox("This PBN file cannot be loaded." & vbCr & _
                         "It does not have a Carriage Return" & vbCr & _
                         "and Line Feed between each board.", _
                         3, "ok", "Loading PBN")
            Close #fiNum
            L_Deal.Top = 120
            L_Deal.Height = 255
            mLoadPbnDealNum.Enabled = False
            Exit Sub
        End If
        
        int_PBNfileLineNum = int_PBNfileLineNum + 1
          ' This tells BB where each board starts in the file. '
          ' When going to the next deal or a selected deal, it '
          ' loops through Line Inputs up to int_PBNfileLineNum.'
        
        If a = "" Then
            previousBlankLine = int_PBNfileLineNum
            ' Marks the start of a new board. '
        
        ElseIf InStr(a, "[Event") > 0 Then
            str_PBNevent = Mid$(a, 9)
            str_PBNevent = Left$(str_PBNevent, Len(str_PBNevent) - 2) ' "]" off '
            n = int_PBNfileLineNum
            'L_Deal = str_PBNevent
            'L_Deal.Width = 3000
        
        ElseIf InStr(a, "[Round") > 0 Then
            str_PBNround = Mid$(a, 9)
            str_PBNround = Left$(str_PBNround, 2) & "-"
            
        ElseIf InStr(a, "[Board") > 0 Then
            If previousBlankLine = 0 And int_PBNfileLineNum > 9 Then
                r = MyMsgBox("This file cannot be loaded." & vbCr & _
                             "It does not have a blank" & vbCr & _
                             "line between each board.", _
                             3, "ok", "Loading PBN")
                Close #fiNum
                L_Deal.Top = 120
                L_Deal.Height = 255
                mLoadPbnDealNum.Enabled = False
                Exit Sub
            End If
            
            brd = Val(Mid$(a, 9, 2))
            cmb_PBNBoards.AddItem "Line " & previousBlankLine & ": Round-Board= " & _
                                 str_PBNround & brd & "  " & str_PBNevent
        End If
        DoEvents
    Loop While Not EOF(fiNum)
    Close #fiNum
        
    int_PBNfileLineNum = 0
    int_PBNlistNum = 0
    
    GetPbnDeal
    
    mLoadPbnDealNum.Enabled = True
End Sub


Private Sub GetPbnDeal()
    ' When a PBN file is first specified, it is opened   '
    ' and the line number in the file for every deal is  '
    ' loaded into list box cmb_PBNBoards along with the  '
    ' round number (if any), board number, and event.    '
    ' Then the first board is displayed. After that, the '
    ' user scan scroll back and forth through the boards '
    ' using the Deal arrow keys or he may select a board '
    ' from the list box by clicking File - Load PBN File.'
    ' So this Sub is called either by Sub mLoadPBN or    '
    ' cmb_PBNBoards. '
    
    Dim a As String
    Dim f As String
    Dim response As String
    Dim b As Integer
    Dim c As Integer
    Dim h As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim rd As String ' round ID with letters ]
    Dim d As Integer ' deal num '
    Dim X As Integer
    Dim cardNumInDeck As Integer
    Dim cardNumInHand As Integer
    Dim fh As String
    Dim firstH As Integer
    Dim suit As String
    Dim s As Integer
    Dim hand As String
    Dim Card As String
    Dim fiNum As Integer
    Dim eventName As String
    Dim bids As String
        
    ' int_PBNlistNum is the ListIndex of the PBN Boards list. '
    ' int_PBNfileLineNum is the line # in the file for the    '
    '    selected board. This is from the selected list line. '
        
    If int_PBNlistNum < 0 Or int_PBNlistNum = cmb_PBNBoards.ListCount Then
        int_PBNlistNum = 0
    End If
    
    a = cmb_PBNBoards.List(int_PBNlistNum)
    int_PBNfileLineNum = Val(Mid$(a, 6))
    i = InStr(a, "Board")
    If i > 0 Then
        i = i + 7
        j = InStr(i + 1, a, " ")
        k = InStr(i + 1, a, "-")
        If k > 0 And k < j Then
            eventName = Mid$(a, i, k - 1)
            i = i + 1
        End If
        int_PbnDealNum = "" & Val(Mid$(a, i, j - i))
    End If
    
    Close #fiNum
    fiNum = FreeFile
    Open str_PBNpathFile For Input As #fiNum
    
    ' Clear the old hand display: '
    For i = 1 To 4
        p_Hand(i).Cls
    Next
    
    Erase str_PBNbids
    ReDim str_PBNbids(2, 50) ' 2=pair#/Team, 50=call# '
    
    t_Disclose = "" & vbCrLf
    
    ' Get to the chosen line number/board: '
    For i = 1 To int_PBNfileLineNum
        Line Input #fiNum, a
    Next
    
    ' Load the board: '
        int_Dealer = 0
        Do While Not EOF(fiNum) And _
             (a <> "" Or int_Dealer = 0)
                          
        If InStr(a, "[Event") > 0 Then
            a = Mid$(a, 9)
            eventName = a
            
        ElseIf InStr(a, "[Round") > 0 Then
            i = InStr(a, Chr$(34))
            rd = Mid$(a, i + 1)
            i = InStr(rd, Chr$(34))
            rd = Left$(rd, i - 1) & "-"
            If str_PBNround = "" Then
                str_PBNround = rd
            End If
        
        ElseIf InStr(a, "[Board") Then
            int_PbnDealNum = "" & Val(Mid$(a, 9))
        
        ElseIf InStr("[West [Nort [East [Sout", Left$(a & "    ", 5)) > 0 Then
            ' Show players' names, if given. '
            If Len(a) > 12 Then
                If Len(t_Disclose) < 6 Then
                    t_Disclose = "      Players: " & vbCrLf
                End If
                t_Disclose = t_Disclose & "                   " & a & vbCrLf
            End If
        
        ElseIf InStr(a, "[Dealer") > 0 Then
            L_Dealer = Mid$(a, 10, 1)
            int_Dealer = InStr("SWNE", L_Dealer)
            int_DealerSpecified = int_Dealer
        
        ElseIf InStr(a, "[Score") > 0 Then
            a = Replace(a, "[Score " & Chr$(34), "")
            a = Replace(a, Chr$(34) & "]", "")
            str_PBNscore = a
        
        ElseIf InStr(a, "[Vul") > 0 Then
            a = Mid$(a, 14, 2)
            If a = "No" Then
                L_Vul = ""
            ElseIf a = "Al" Then
                L_Vul = "Both"
            ElseIf a = "NS" Then
                L_Vul = "Unfav."
            ElseIf a = "EW" Then
                L_Vul = "Fav."
            Else
                L_Vul = a
            End If
        
        ElseIf InStr(a, "[Auction") > 0 Then
            b = 0
            Line Input #fiNum, bids
            Do While InStr(bids, "[") = 0
                ' e.g.: "Pass 1NT X XX" '
                bids = Replace(bids, "Pass", " P")
                bids = Replace(bids, "NT", "N")
                bids = Replace(bids, "XX", " R")
                bids = Replace(bids, "X", " D")
                Do: bids = Replace(bids, "   P", "  P"): Loop While InStr(bids, "   P")
                ' Result: " P 1N  D  R" '
                '          123456789-1  '
                str_PBNbids(int_PBNpair, b + 1) = Trim$(Left$(bids, 2))
                str_PBNbids(int_PBNpair, b + 2) = Trim$(Mid$(bids, 4, 2))
                str_PBNbids(int_PBNpair, b + 3) = Trim$(Mid$(bids, 7, 2))
                str_PBNbids(int_PBNpair, b + 4) = Trim$(Mid$(bids, 10, 2))
                b = b + 4
                Line Input #fiNum, bids
            Loop
            ' Not using plays at this time. '
            
        ElseIf InStr(a, "[Deal ") > 0 And a <> "" Then
            fh = Mid$(a, 8, 1)
            ' First hand shown on Deal may or may not be the actual dealer. '
            a = Mid$(a, 10, 67) ' The hands, in a row. See the PBN file. '
            h = InStr("SWNE", fh)
            firstH = h
            ' Array.Clear(myArray,0,myArray.Length) '
            ' Erase Hand
            ' ReDim hand(4, 13, 9)
            Erase int_Hand
            Erase int_HandQty
            Erase int_Suit
            Erase int_SuitQty
            Erase str_Hands
            ReDim int_Hand(4, 13)
            ReDim int_HandQty(4)
            ReDim int_Suit(4, 4, 13)
            ReDim int_SuitQty(4, 4)
            ReDim str_Hands(4)
            Do
                hand = Left$(a, 16) ' e.g.: KQT.KT86.K93.A64 '
                a = Mid$(a, 18)
                str_Hands(h) = Replace(hand, ".", "-")
                s = Spd ' Hand starts with spade. '
                k = 0
                For c = 1 To 16
                    Do While Mid$(hand, c, 1) = "."
                        If s = Clb Then
                            Exit Do
                        End If
                        s = s - 1 ' Next suit: S=4 H=3 D=2 C=1
                        c = c + 1
                    Loop
                    Card = Mid$(hand, c, 1)
                    cardNumInHand = InStr("A23456789TJQK", Card)
                    cardNumInDeck = (s - 1) * 13 + cardNumInHand
                    If cardNumInDeck = 0 Then
                        GoTo ErrorLoadingPbn
                    End If
                    k = k + 1
                    If k = 14 Then Exit For
                    int_Hand(h, k) = cardNumInDeck
                    int_HandQty(h) = int_HandQty(h) + 1
                    int_Suit(h, s, cardNumInHand) = cardNumInDeck
                    int_SuitQty(h, s) = int_SuitQty(h, s) + 1
                Next
                DisplayHand h
                h = h + 1
                If h = 5 Then h = 1
                DoEvents
            Loop While h <> firstH
            bln_SkipIt = True ' Exit the loop below when a="" '
            
        ElseIf InStr(a, "[Declarer") > 0 Then
            
        ElseIf a = "{" Then
            str_PBNanalysis = ""
            Line Input #fiNum, a
            Do
                str_PBNanalysis = str_PBNanalysis & " " & a
                Line Input #fiNum, a
                a = Replace(a, "\\", "")
            Loop While a <> "}"
            str_PBNanalysis = Replace(str_PBNanalysis, ". ", "." & vbCrLf)
            str_PBNanalysis = Replace(str_PBNanalysis, "? ", "?" & vbCrLf)
            str_PBNanalysis = LTrim$(str_PBNanalysis)
        
        ElseIf InStr(a, "[Result") > 0 Then
            str_PBNscore = a
        End If
        
        Line Input #fiNum, a
        
        'If a = "" Then
        '    If bln_SkipIt Then
        '        bln_SkipIt = False
        '        Exit Do
        '    End If
        '    Line Input #fiNum, a
        '    If InStr(a, "[Event") > 0 Then Exit Do
        'End If
        DoEvents
    Loop
    Close #fiNum
        
    For h = 0 To 3
        str_DDAOutput(h) = ""
    Next
   
    ' Fill Cards() array: '
    i = 0
    For h = 1 To 4
        For s = 1 To 4
            For c = 1 To 13
                If int_Suit(h, s, c) > 0 Then
                    i = i + 1
                    int_Cards(i) = int_Suit(h, s, c)
                End If
            Next
        Next
    Next
            
    SetUpHands 0 ' Creates str_Hands() in DealerRoutines. '
    
    ' Show board number and file name: '
    str_PBNfile = Replace(str_PBNfile, ".pbn", "")
    L_Deal = str_PBNround & int_PbnDealNum & " - " & str_PBNfile
    i = TextPixelWidth(L_Deal.Caption)
    If i > 280 Then
        ' Pixel Width is too wide for one line. '
        L_Deal.Top = 10
        L_Deal.Height = 400
    Else
        L_Deal.Top = 120
        L_Deal.Height = 255
    End If
    
    ShowDDA False
    
    bln_PbnEvent = True
    
ExitGetPbn:
    Close #1
    ChDir str_BBpath ' Something was changing the default directory. '
    On Local Error GoTo 0
    cmb_PBNBoards.Visible = False
Exit Sub

ErrorLoadingPbn:
    MsgBox "Error in hand " & vbCr & hand & vbCr & "for Board #" & int_PbnDealNum
    Resume ExitGetPbn
End Sub


Sub OpenPracticeFile()
    Dim f As String
    Dim i As Integer
    Dim fiNum As Integer
    Dim lastDeal As lastDeal
    
    ChDir str_BBpath
    str_PracticeSpecsPath = str_BBpath & "\PracticeDeals\"
    str_PracticeDealsPath = str_PracticeSpecsPath & "Deals\"
    
    ' A folder off the Bridge folder. '
    ' Deal files folders are in       '
    ' ...\Bridge\PracticeDeals\Deals  '

    If Dir$(str_PracticeSpecsPath & "\*", vbDirectory) = "" Then
        MkDir str_PracticeSpecsPath
    End If
    
    If Dir$(str_PracticeDealsPath & "\*", vbDirectory) = "" Then
        MkDir str_PracticeDealsPath
    End If
    ChDir App.Path
    
    f = Dir$(str_PracticeSpecsPath & "\.")
    cmb_FileList.Clear
    cmb_FileList.Visible = True
    cmb_FileList.AddItem "Enter a new convention above."
    cmb_FileList.AddItem "     Click this line to cancel."
    cmb_FileList.AddItem "Click a line below to select."
    cmb_FileList.AddItem "--------------------------------------------"
    Do While f <> ""
        cmb_FileList.AddItem f
        f = Dir$
    Loop
    cmb_FileList.Left = 167
    cmb_FileList.Width = 325
    cmb_FileList.Height = 235
    cmb_FileList.Top = 96
End Sub


Private Sub mRedisplayHands_Click()
    Dim h As Integer
    
    For h = 1 To 4
        DisplayHand h
    Next
End Sub


Private Sub mShowPracCnvDocs_Click()
    Dim xx As Long
    Dim p As String
    
    'MsgBox "The documentation for creating and" & vbCr & _
           "using the convention quizzes should" & vbCr & _
           "be read before attempting to use" & vbCr & _
           "the Convention menu commands."
           
    p = str_BBpath
    
    If Dir$(p & "\*") = "" Then
        If p <> App.Path Then
            str_BBpath = App.Path
            mShowPracCnvDocs_Click
        Else
            MsgBox p & " is not a valid path." ' Shouldn't be possible.
        End If
        Exit Sub
    End If
           
    If Dir$(p & "\Notes\CardShark_BidBase_Practice_Program.htm") = "" Then
        MsgBox "CardShark_BidBase_Practice_Program.htm is not in the Notes directory." & vbCr & _
               "You can find it online at www.aeyec.com/BidBase/Notes/"
        Exit Sub
    End If
           
    xx = Shell(str_BBpath & "\BidBase File Viewer.exe " & _
               str_BBpath & "\Notes\CardShark_BidBase_Practice_Program.htm#C1", vbNormalFocus)
End Sub


Private Sub mUpdateConventionFiles_Click()
    bln_UpdateCnvFiles = Not bln_UpdateCnvFiles
    mUpdateConventionFiles.Checked = bln_UpdateCnvFiles
    
    ' When a practice file is created, it makes  '
    ' a file with the convention name and a file '
    ' extension of ".txt" which has all the deal '
    ' numbers found for the convention.  At the  '
    ' time, a "Practice_[conv.name]_.txt file is '
    ' created with HTML code to be plugged into a'
    ' web page for practicing the convention. '
    ' If future BidBase entries are made/changed '
    ' which affect the convention, then instead  '
    ' of having to do the Create from scratch,   '
    ' which can take a long time, do the Update  '
    ' and it gets the previously found deal #'s  '
    ' and looks up the new final bid if that is  '
    ' all that changed. If the deal no longer    '
    ' matches the convention's auction(s), then  ;
    ' it skips that deal. '
    ' All that is needed by the Update sub is to '
    ' set the variable shown above, then Practice'
    ' is called and does the rest. '
    Dim bid2 As Integer
    
    bid2 = MsgBox("Get a 2nd bid for practicing?", vbYesNoCancel, "Find Practice Deals")
    If bid2 = vbYes Then
        bln_Get2ndBid = True
    ElseIf bid2 = vbCancel Then
        ClearPracticeVariables
        Exit Sub
    End If
            
    mPractice_Click
End Sub


Private Sub mPractice_Click()
    Dim i As Integer
    Dim n As Long
    Dim fiNum As Long
    Dim lastDeal As lastDeal ' Record format for the 1st enty in a .Deals file. '

    If mPractice.Checked Then
        mPractice.Checked = False
        ClearPracticeVariables
        Exit Sub
    End If
    
    ClearPracticeVariables
    
    mPractice.Checked = True
    
    If bln_Abort Then
        Exit Sub
    End If

    OpenPracticeFile ' Select the file. No data retrieved. '
    
    If str_PracticeSpecsFile = "" Then
        Exit Sub
    End If
    
    ' The first record in a convention's deals file consists  '
    ' of the number of entries in the file (counting this 1st '
    ' record) followed by the entry number of the last deal   '
    ' practiced.  For example, "150, 27" says that there are  '
    ' 150 records in the file and #27 was the last one used.  '
    
    ' If it's the first time used or if starting over, the    '
    ' the 2nd record is the first deal and would look like:   '
    ' "3561, 1" where "3561" is the deal number and "1" is    '
    ' the position number of the dealer (1=S, 2=W, 3=N, 4=E). '
    
    ' The GetPracticeDeal and PutPracticeDeal Subs get and    '
    ' save deal records as needed. Instead of using the Put   '
    ' sub below, we get the 1st record here since it is not   '
    ' a deal # record. '
    
    If Dir$(str_PracticeDealsFile) = "" Then
        ' No deals found yet. '
        mCreateConv_Click
        Exit Sub
    End If
    
    mPractice.Checked = True
    
    LoadPracticeDealsAuctions
    
    ' Get the deal to be practiced: '
    fiNum = FreeFile
    Open str_PracticeDealsFile For Random As fiNum Len = 6
        Get #fiNum, 1, lastDeal
        lng_NumPracticeEntries = lastDeal.NumEntries  ' Number of entries in the file. '
        int_DealFileEntryNum = lastDeal.LastNumUsed   ' Record # for the last file record used. '
        n = int_DealFileEntryNum
        
        If int_DealFileEntryNum > 2 Then
            int_DealFileEntryNum = InputBox("What file entry number do you want to start with?" & vbCrLf & "(Press Enter for the last entry used as shown.)", str_PracticeDealsFile, int_DealFileEntryNum)
            lastDeal.LastNumUsed = int_DealFileEntryNum
            Put #fiNum, 1, lastDeal
        End If
        
        If int_DealFileEntryNum < 2 Then
            int_DealFileEntryNum = 2 ' Even though it is the first file# entry, it is the second entry in the file. '
            If lng_NumPracticeEntries < 2 Then
                lng_NumPracticeEntries = 2
            End If
        End If
        
        If int_DealFileEntryNum >= lng_NumPracticeEntries Then
            ' When the last record has been reached,'
            ' circle back around to the first deal: '
            int_DealFileEntryNum = 2
        End If
        
        ' Get the first deal (or last '
        ' one used) to practice on:   '
        GetPracticeDeal
    Close #fiNum
        
    If n <> int_DealFileEntryNum Then
        SaveLastFileEntryNumberUsed
    End If
        
    bln_SettingUpPracBid = False
    ' Other bln's are fixed in GetPracticeDeal

    'lng_DealNum2Practice  =
    L_Deal = "" & lng_DealNum2Practice
    lng_DealNum = lng_DealNum2Practice
    lng_DealNumSpecified = lng_DealNum
    L_Dealer = Mid$("SWNE", int_Dealer, 1)
    int_DealerSpecified = int_Dealer
    
    bln_HandsDealt = False
    Deal ' This should show the newly dealt hands. '
    For i = 1 To 4
        DisplayHand i
    Next
    ' We want the program to make the bids '
    ' for the number of bids specified. '
    Stop
    
    t_Disclose = ""
        
    str_TimerCallingSub = "mPractice"
    str_TimerJob = "Restart same deal" ' Don't want the deal to be re-dealt. '
    Timer1.Enabled = True
    ' All the above 2 lines do is run FinishDealSetup '
    ' but this allows time for this sub to end. '
End Sub


Private Sub ClearPracticeVariables()
        int_AuctionNumber = 0
        int_DealFileEntryNum = 0
        lng_DealNum2Practice = 0
        'Erase str_FindAuctions()
        'int_NumAuctionSpecs = 0
        lng_NumPracticeEntries = 0
        'str_PracticeSpecsFile = ""
        'str_PracticeDealsFile = ""
        int_PracRecordNum = 0
        bln_ReDoFindFile = False
        bln_SettingUpPracBid = False
        bln_SettingUpPracBid = False
        bln_SearchForMoreDeals = False
        bln_PracticeFoundBids = False
        bln_FindAuctions = False
        bln_FindAuctionsDone = False
        'mPractice.Checked = False
        'mFindDeals.Checked = False
        'mCreateConv.Checked = False
        'mUpdateConventionFiles.Checked = False
        
        str_AllCalls = ""
        int_CallNum = 0
        bln_Rebid = False
        bln_SameDeal = False
        int_Turn = 0
        
        Erase str_Call()  ' dimensioned to 50 in this module '
        Erase str_Player() '  ditto, to 4 '
        Erase int_Player() '  ditto, to 4 '
        
        Erase str_Dealer() ' dimensioned to 4 in DealerRoutines '
        
        Erase int_Cards()
        ReDim int_Cards(65)
        
        Erase int_Hand()
        ReDim int_Hand(4, 13)
        
        Erase int_HandQty()
        ReDim int_HandQty(4)
        
        Erase str_Hands()
        ReDim str_Hands(4)
        
        Erase int_Suit()
        ReDim int_Suit(4, 4, 13)
        
        Erase int_SuitQty()
        ReDim int_SuitQty(4, 4)
        
        SetUpBidDisplayGrid
End Sub
                

Private Sub mRotate_Click(Index As Integer)
    Dim p As Integer ' player
    Dim n As Integer ' next player '
    Dim s As Integer ' suit
    Dim c As Integer ' card
    Dim i As Integer
    Dim X As Integer
    Dim ndx As Integer
    Dim cards(13) As Integer ' Temporary '
    
    
    ' This rotates the hands 1 position counter-clockwise. '
    ' Index = the number of rotations: 2 moves W to E, etc.'
    ' At first I wrote a routine to swap, say, N to E, etc.'
    ' in one runthrough. Doing 3 run-throughs to go from N '
    ' to E is less efficient, but the time difference is   '
    ' nothing and the code is much simpler this way. '
    
    ' This Sub is called by Rotate menu options and by the  '
    ' Sub ProcessFindBids used when practicing conventions  '
    ' where we want South to always have the next bid after '
    ' the specified auctions. '
    
    For i = 1 To Index
        For c = 1 To 13
            int_Cards(52 + c) = int_Cards(c)
        Next
        For p = 0 To 3
            For c = 1 To 13
                int_Cards(c) = int_Cards(p * 13 + c)
            Next
        Next
    
        For p = 0 To 4 ' 0<1, 1<2, 2<3, 3<4, 4<0 '
            If p < 4 Then
                n = p + 1
            Else
                n = 0
            End If
            
            str_Hands(p) = str_Hands(n)
            
            For s = 1 To 4
                For c = 1 To 13
                    int_Suit(p, s, c) = int_Suit(n, s, c)
                Next
                int_SuitQty(p, s) = int_SuitQty(n, s)
            Next
        Next
    Next
    
    If 1 = 2 And InStr(str_FindAuctions(int_AuctionNumber), "|") = 0 Then
        int_Dealer = int_Dealer - ndx
        If int_Dealer = 0 Then int_Dealer = 4
        int_Turn = int_Turn + 1
        If int_Turn = 5 Then int_Turn = 1
        int_CallNum = int_CallNum + 1
    End If
    
    int_Dealer = int_Dealer - Index
    If int_Dealer < 1 Then
        int_Dealer = int_Dealer + 4
    End If
    int_Player(1) = int_Dealer
    int_DealerSpecified = int_Dealer
    SetPlayerNumbers
    DisplayBid 1
    For i = 1 To 4
        DisplayHand i
    Next
End Sub


Sub CopyArrayElement(pTo As Integer, _
                     pFrom As Integer)
    Dim s As Integer
    Dim c As Integer
' This Sub is not being used. '

    For s = 1 To 4
        For c = 1 To 13
            int_Suit(pTo, s, c) = int_Suit(pFrom, s, c)
        Next
    Next
End Sub


Private Sub mSaveDeal_Click()
    ' Appends the current deal to the '
    ' SavedDeals.txt file in str_BBpath.'
    
    Dim f As String
    Dim fiNum As Integer
    Dim i As Integer
    Dim xx As Long
    
    f = str_BBpath & "\Notes\SavedDeals.txt"
    
    fiNum = FreeFile
    Open f For Append As #fiNum
        Print #fiNum, ""
        Print #fiNum, ""
        Print #fiNum, "Deal #"; L_Deal
        Print #fiNum, "N:  "; str_Hands(3)
        Print #fiNum, "S:  "; str_Hands(1)
        Print #fiNum, "W:  "; str_Hands(2)
        Print #fiNum, "E:  "; str_Hands(4)
        Print #fiNum, "Dealer: "; Left$(L_Dealer, 1)
        Print #fiNum, "Bidding:"
        For i = 1 To int_CallNum
            Print #fiNum, str_Call(i);
            If i < int_CallNum And i Mod 4 > 0 Then
                Print #fiNum, "-";
            Else
                Print #fiNum, ""
            End If
        Next
        i = 0
        Do While i < cmb_DDA.ListCount
            Print #fiNum, cmb_DDA.List(i)
            i = i + 1
        Loop
        Print #fiNum, ""
    Close #fiNum
    ' Display file in Notepad: '
    xx = Shell("c:\windows\notepad.exe " & f, vbNormalFocus)
End Sub


Private Sub mScramble_Click(Index As Integer)
     ' Shuffles cards among user specified '
     ' hands to see what the bidding would '
     ' be against different lay of cards.  '
    Dim h As Integer
    Dim rh As String
    Dim numDeals As Integer
    Dim i As Integer
    Dim d As String
    
    int_NumMixedHands = 0
    
    If Index = 1 Then
        
        ' Rescramble previously selected hands: '
        For h = 1 To 4
            If bln_MixHands(h) Then
                int_NumMixedHands = _
                int_NumMixedHands + 1
            End If
        Next
    
        If int_NumMixedHands = 0 Then
            MsgBox "No record of hands selected." & vbCr & _
                   "Please enter again."
            mScramble_Click 0
            Exit Sub
        End If
    
    Else
    
        rh = Trim$(UCase$(InputBox$( _
            "Enter the hands (N,S,E,W) to scramble.")))
    
        If rh = "" Then Exit Sub
    
        Erase bln_MixHands()
        For h = 1 To 4
            If InStr(rh, Mid$("SWNE", h, 1)) Then
                bln_MixHands(h) = True
                int_NumMixedHands = _
                int_NumMixedHands + 1
            End If
        Next
    End If
  
    If int_NumMixedHands < 2 Then
        MsgBox "You must specify at least 2 hands to Scramble."
        mScramble_Click 0
        Exit Sub
    
    ElseIf int_NumMixedHands > 3 Then
        MsgBox "You cannot Scramble more than 3 hands." _
            & vbCr & "To Scramble all 4, just click on Deal."
        mScramble_Click 0
        Exit Sub
    End If

    ScrambleHands
    SortHands ' Sort hands within int_Cards().
    FillSuits ' Assign int_Cards() to int_Suit().'
    
    For h = 1 To 4
        DisplayHand h
    Next
    ShowDDA False
    b_RestartDeal_Click
    mScramble(1).Enabled = True

End Sub


Private Sub mShowConvDocs_Click()
    Dim xx As Long
    Dim p As String
    
    MsgBox "A directory of convention docs" & vbCr & _
           "arranged in convention card order" & vbCr & _
           "will be displayed next. Click on" & vbCr & _
           "a convention to see its doc."
           
    p = str_BBpath
    
    If Dir$(p & "\*") = "" Then
        If p <> App.Path Then
            str_BBpath = App.Path
            mShowConvDocs_Click
        Else
            MsgBox p & " is not a valid path." ' Shouldn't be possible.
        End If
        Exit Sub
    End If
           
    If Dir$(p & "\Notes\Index.htm") = "" Then
        MsgBox "Index.htm is not in the Notes directory." & vbCr & _
               "You can find it online at www.aeyec.com/BidBase/Notes"
        Exit Sub
    End If
           
    xx = Shell(str_BBpath & "\BidBase File Viewer.exe " & _
               str_BBpath & "\Notes\Index.htm", vbNormalFocus)
End Sub


Private Sub mStayOnTop_Click()
    If mStayOnTop.Checked Then
        MakeNormal hwnd
    Else
        MakeTopMost hwnd
    End If
End Sub


Private Sub mTutorial_Click()
    mTutorial.Checked = True
'    Welcome
    mTutorial.Checked = False
End Sub


Private Sub mRunEditor_Click()
    Dim ret As Long
    
    On Local Error GoTo LoadEditor
    ' If Practice is not already loaded, '
    ' AppActivate will generate an error '
    ' which will Shell to the program.   '
    ' This keeps more than one Practice  '
    ' program from opening.              '
    AppActivate "BidBase Editor.exe"
    'GoTo ExitRunEditor

LoadEditor:
    If Dir$(str_BBpath & "\BidBase Bidding Practice.exe") = "" Then
        MsgBox "BidBase Editor is not present."
    Else
        ret = Shell(str_BBpath & "\BidBase Editor.exe", vbNormalFocus)
    End If
    Resume ExitRunEditor

ExitRunEditor:
    On Local Error GoTo 0

End Sub

Private Sub p_Hand_MouseDown(Index As Integer, _
                             Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    ' Index is hand#. Cards in a hand are just '
    ' drawn onto this p_Hand picture control.  '
    ' X and Y are the coordinates of where the '
    ' mouse clicked in p_Hand() and are used   '
    ' to calculate which "card" was clicked.   '
    ' The card is returned to the deck and all '
    ' the hand variables are adjusted to match.'
    Dim i As Long
    Dim lastCard As Long
    Dim cardNumInHand As Long
    Dim cardNumInDeck As Long
    Dim CardNumInSuit As Long
    Dim numPix As Long
    Dim suit As Integer
    Dim hnd As String
    Dim crd As Integer
    Dim n As Integer

    If t_EnterHand(Index).Visible = False Then Exit Sub
    
    If Index <> int_Turn Then
      For i = 1 To 4
        If i = Index Then
            t_EnterHand(i).BackColor = 65535 ' yellow '
        Else
            t_EnterHand(i).BackColor = 13434828 ' green '
        End If
      Next
    End If
    
    int_Turn = Index
    hnd = Replace(str_Hands(int_Turn), "-", "")
    lastCard = Len(hnd) - 1

    If int_Hand(int_Turn, 1) = 0 Then
        MsgBox "Stop p_Hand.MouseDown 1"
    Else
        For i = 1 To 13
            n = int_Hand(int_Turn, i)
            suit = GetCardSuit(n)
            crd = GetCardValue(n)
            int_Suit(int_Turn, suit, crd) = n
        Next
    End If
    ' The cards are just drawn on 1 picture control '
    ' per hand. Calculate which card was clicked on: '
    If Index = 1 Or Index = 3 Then
        ' S and N cards go left to right 300 pix each. '
        numPix = lastCard * 300
        If X > numPix Or lastCard = 1 Then
            cardNumInHand = lastCard
        Else
            cardNumInHand = Fix(X / 300) + 1
            If X / 300 = X \ 300 Then
                ' It is 300 from the start of one card to '
                ' the start of the next, but it appears   '
                ' possible to click the very right edge of'
                ' say, the first card and get a multiple  '
                ' of 300 and thus get a cardNumInHand = 2.'
                ' In that case, we need to deduct 1. '
                cardNumInHand = cardNumInHand - 1
            End If
        End If
    Else ' E and W card go top to bottom 400 pix each. '
        numPix = lastCard * 400
        If Y > numPix Then
            cardNumInHand = lastCard
        Else
            cardNumInHand = Fix(Y / 400) + 1
        End If
    End If
    If cardNumInHand < 1 Then cardNumInHand = 1
        
    cardNumInDeck = int_Hand(Index, cardNumInHand)
    ' Mark the card as back in the deck: '
    bln_SelectedCards(cardNumInDeck) = False
            
    ' Remove card from int_Hand() '
    For i = cardNumInHand To lastCard
        int_Hand(int_Turn, i) = int_Hand(int_Turn, i + 1)
    Next
    int_Hand(int_Turn, i) = 0
        
    On Local Error GoTo DrawCardErr
    DrawCard Card(cardNumInDeck).hwnd, cardNumInDeck, 0, 0
    ' Remove cardNum from player's hand (Index = player #): '
    ' Example: Say the 4th card in the hand display is      '
    ' clicked and is the Ace of D.                          '
    ' int_Hand(Index, 4) would have a cardNumInDeck of 14.  '
    ' suit = Fix((14 - 1) / 13) + 1 = 2 = Diamonds          '
    ' If 2 of H: Fix((28 - 1) / 13) + 1 = 3 = Hearts        '
    ' If K of S: Fix((52 - 1) / 13) + 1 = 4 = Spades        '
    suit = Fix((cardNumInDeck - 1) / 13) + 1
    CardNumInSuit = cardNumInDeck - (suit - 1) * 13
    int_Suit(Index, suit, CardNumInSuit) = 0
    'int_SuitQty(Index, suit) = int_SuitQty(Index, suit) - 1
    int_HandQty(Index) = int_HandQty(Index) - 1
    DisplayHand Index
    ShowPickedCardsInTextBox
Exit Sub

DrawCardErr:
    MsgBox Err.Description & " in p_Hand_MouseDown."
    Resume Next
End Sub


Private Sub ClearHandVariables(h As Integer)
    ' Can't use Erase when only clearing 1 hand. '
    Dim s As Integer ' suit '
    Dim c As Integer ' card '
    
    If int_HandQty(h) = 0 Then
        Exit Sub
    End If
    
    For c = 1 To 13
        If int_Hand(h, c) > 0 Then
            If f_CardDeck.Visible And bln_SelectedCards(int_Hand(h, c)) Then
                bln_SelectedCards(int_Hand(h, c)) = False
                DrawCard Card(int_Hand(h, c)).hwnd, int_Hand(h, c), 0, 0
            End If
            int_Hand(h, c) = 0
        End If
    Next
    
    For s = 1 To 4
        For c = 1 To 13
            int_Suit(h, s, c) = 0
        Next
        On Local Error GoTo DimSuitQty
        int_SuitQty(h, s) = 0
        On Local Error GoTo 0
    Next
    
    p_Hand(h).Cls
    int_HandQty(h) = 0
Exit Sub
    
DimSuitQty:
    ReDim int_SuitQty(4, 4)
    Resume
End Sub


Private Sub t_EnterHand_Change(h As Integer)
    Dim c As Integer
    Dim cDk As Integer
    Dim crd As String
    Dim d As Integer ' dashes '
    Dim s As Integer
    Dim i As Integer
    Dim j  As Integer
    Dim n As Integer
    
    int_Turn = h
    
    str_Hands(h) = Trim(str_Hands(h))
    t_EnterHand(h) = Trim(t_EnterHand(h))
    
    If bln_SkipIt Then
        t_EnterHand(h) = str_Hands(h)
        Exit Sub
    End If
    
    ClearHandVariables h
    ' Change int_Hand to match t_EnterHand. '
    ' Draw the card back on cards entered.  '
    ' When done, display the cards entered. '
    s = 4
    'On Local Error GoTo EnterHandError
    For c = 1 To Len(t_EnterHand(h))
        crd = Mid$(t_EnterHand(h), c, 1)
        If crd <> "-" Then
            n = 14 - InStr("KQJT98765432A", crd)
            cDk = (s - 1) * 13 + n
            If int_Hand(h, c - (4 - s)) <> cDk Then
                If bln_SelectedCards(cDk) = True Then
                    MsgBox Left$(str_Suits(s), 1) & ":" & crd & " has already been dealt."
                    Exit Sub
                End If
                int_Hand(h, c - (4 - s)) = cDk
                int_Suit(h, s, n) = cDk
                int_SuitQty(h, s) = int_SuitQty(h, s) + 1
                DrawBack Card(cDk).hwnd, 3, 0, 0
                bln_SelectedCards(cDk) = True
            End If
        ' Calculate next suit: '
        ElseIf s = 1 Then
            MsgBox "Too many dashes."
            Exit Sub
        Else
            s = s - 1
        End If
    Next
    'If s > 1 Then
    '    MsgBox "Even if fewer than 4 suits," & vbCr & _
    '           "the hand must have 3 dashes." & vbCr & _
    '           "Example:  xxxx-xxxx-xxxxx-"
    '    Exit Sub
    'End If
    
    DisplayHand h
    str_Hands(h) = t_EnterHand(h)
    int_HandQty(h) = Len(str_Hands(h)) - 3
    
    int_NumPickedHands = 0
    For h = 1 To 4
        If int_HandQty(h) = 13 Then
            int_NumPickedHands = int_NumPickedHands + 1
        ElseIf int_HandQty(h) > 0 And int_HandQty(h) < 13 Then
            ' Only allow Deal The Rest if no partial hands. '
            b_DealRest.Visible = True
        Else
            Exit For
        End If
    Next
Exit Sub

EnterHandError:
    t_EnterHand(h) = ""
    MsgBox "Error entering hand."
End Sub


Private Sub t_EnterHand_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim ss As Integer ' SelStart '
    Dim sL As Integer ' SelLength '
    Dim sTxt As String ' Selected text '
    Dim i As Integer
    Dim d As Integer
    Dim hnd As String
    
    If (KeyCode = 37 Or KeyCode = 39) And Shift = 0 Then
        ' Left/Right Cursor '
        Exit Sub
    ElseIf Shift = 2 Or InStr(t_EnterHand(Index).SelText, "-") = 0 Then
        Exit Sub
    End If

    ss = t_EnterHand(Index).SelStart
    sL = t_EnterHand(Index).SelLength
    sTxt = t_EnterHand(Index).SelText
    
    For i = 1 To Len(sTxt)
        If Mid$(sTxt, i, 1) = "-" Then
            d = d + 1
        End If
    Next
    
    hnd = t_EnterHand(Index)
    hnd = Left$(hnd, ss) & String(d, "-") & Mid$(hnd, ss + sL + 1)
    t_EnterHand(Index) = hnd
    KeyCode = 0
End Sub


Private Sub t_EnterHand_KeyPress(Index As Integer, KeyAscii As Integer)
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

    With t_EnterHand(Index)
    '''''''''''''''''''''''
    i = InStr("AKQJT98765432", k)
    If .SelStart > 0 And i > 0 Then
        j = InStr("AKQJT98765432", Mid$(.Text, .SelStart, 1))
        If j >= i Then 'card entered is out of rank order
            ss = .SelStart
            .Text = Left$(.Text, .SelStart) + _
                    k + _
                    Mid$(.Text, 1)
            .SelStart = ss + 1
            MsgBox "'" & k & "' is out of order."
            KeyAscii = 0
            Exit Sub
        End If
    End If
    ''''''''
    End With
    bln_SkipIt = False
End Sub


Private Sub t_EnterHand_GotFocus(Index As Integer)
    Dim i As Integer
    
    int_Turn = Index
    For i = 1 To 4
        If i = Index Then
            t_EnterHand(i).BackColor = 65535 ' yellow '
        Else
            t_EnterHand(i).BackColor = 13434828 ' green '
        End If
    Next
    
    With t_EnterHand(Index)
    '''''''''''''''''''''''
    If .Text = "" Then
        Exit Sub
    End If
    
    .SelStart = 0
    .SelLength = Len(.Text)
    ''''''''
    End With
End Sub


Private Sub t_EnterHand_LostFocus(Index As Integer)
    t_EnterHand(Index).BackColor = 13434828 ' green '
End Sub

Private Sub t_EnterHand_Validate(Index As Integer, Cancel As Boolean)
    Dim d As Integer
    Dim i As Integer
    Dim c As Integer
    Dim n As Integer
    
    If bln_SkipIt Then
        bln_SkipIt = False
        Exit Sub
    End If
    
    Erase int_Suit()
    ReDim int_Suit(4, 4, 13)
Exit Sub
    If t_EnterHand(Index) = str_Hands(Index) Or t_EnterHand(Index) = "" Then
        Exit Sub
    End If
    bln_SkipIt = True ' Don't do t_EnterHand_Change '
    
    n = Len(t_EnterHand(Index))
    If n = 0 Then
        t_EnterHand(Index) = ""
        n = 3
    Else
        c = 0
        For i = n To 1 Step -1
            If Mid$(t_EnterHand(Index), i, 1) = "-" Then
                'int_Suit(Index, 4 - d, c) =
                c = 0
                d = d + 1
            Else
                c = c + 1
            End If
        Next
        If d < 3 Then ' Shouldn't be possible; checked in .Change sub. '
            MsgBox "Even if fewer than 4 suits," & vbCr & _
                   "the hand must have 3 dashes." & vbCr & _
                   "Example:  xxxx-xxxx-xxxxx-"
            Cancel = True
            bln_SkipIt = False
            Exit Sub
        End If
    End If
    
    int_HandQty(Index) = n - 3
    str_Hands(Index) = t_EnterHand(Index)
    bln_SkipIt = False
End Sub


Private Sub t_Inbox_GotFocus()
    t_Inbox.SelStart = 0
    t_Inbox.SelLength = 20 ' Anything longer than the actual length. '
End Sub


Private Sub t_Inbox_KeyPress(KeyAscii As Integer)
    Dim c As String
    Dim d As Integer
    
    c = UCase$(Chr$(KeyAscii))
    KeyAscii = Asc(c)
    
    d = InStr("SWNE", c)
    If d > 0 And InStr(L_InboxLbl.Caption, "dealer") > 0 Then
        b_DoneDealNum_Click
    End If
End Sub


Private Sub t_Disclose_DblClick()
    ' When a note file is referenced in the '
    ' disclosure box, double-clicking it    '
    ' will bring it up in the file viewer.  '
    Dim t As String
    Dim f As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim xx As Long
    
    t = t_Disclose
    k = InStr(t, vbCrLf)
    
    f = Trim$(Left$(t, k - 1))
    ' If file name is like "Take-Out Double: Shape Based" '
    ' then use only to the left of the colon.             '
    i = InStr(f, ":")
    If i > 0 Then
        f = Left$(f, i - 1)
        For i = 1 To Len(f$)
            If InStr("/\.?^*<>| ", Mid$(f$, i, 1)) > 0 Then
                Mid$(f$, i, 1) = "_"
            End If
        Next
    End If
    
    If Dir$(str_BBpath & "\Notes\" & f & ".htm") <> "" Then
        xx = Shell(str_BBpath & "\BidBase FileViewer.exe " _
             & str_BBpath & "\Notes\" & f & ".htm", 1)
    Else
        MsgBox "No file for " & Chr$(34) & f & Chr$(34)
    End If
End Sub


Private Sub InitiateBidding()
    ' This routine gets everything ready   '
    ' to start the bidding for the current '
    ' deal. It ends by calling BidNextHand '
    ' which gets everything ready to call  '
    ' MakeBid which calls the bidding DLL. '
    ' After showing the user the bid, the  '
    ' program waits for him to click OK,   '
    ' then ProcessBid is called.
    ' If not passed out, then ProcessBid   '
    ' ends by BidNextHand again. '
    
    Dim d As String
    Dim n As Integer

If ck_KeepDealer.Value = 0 Then
    d = Trim$(Left$(L_Dealer & " ", 1))
  
    If d = "" Then
        If int_Player(1) > 0 Then
            d = Mid$("SWNE", int_Player(1), 1)
        ElseIf int_Turn > 0 Then
            d = Mid$("SWNE", int_Turn, 1)
        End If
        
        GetDealer
        
        ' Don't know what's going on here. '
        ' Why set d and then go to Get Dealer '
        ' which sets int_DealerSpecified but  '
        ' has no effect on d so the bidding   '
        ' never takes place. So I added the   '
        ' next line. I'll try to figure it out'
        ' later. '
        d = Mid$("SWNE", int_DealerSpecified, 1)
        
        If d = "" Then
            bln_Abort = True
            Exit Sub
        End If
        
    ElseIf int_DealerSpecified > 0 Then
        n = int_DealerSpecified
        d = Mid$("SWNE", n, 1)
    Else
        If Not bln_SameDeal Then
            ' Get this dealer: '
            n = InStr("SWNE", d)
            ' Get next dealer: '
            d = Mid$("WNES", n, 1)
        End If
    End If
    
    int_Player(1) = InStr("SWNE", d)
    If int_Player(1) = 0 Then int_Player(1) = 1

    int_DealerSpecified = int_Player(1)
    L_Dealer.Caption = d
    ck_KeepDealer.Visible = True
End If

    If Not bln_PbnEvent Then
        SetPlayerNumbers
    End If
    
    Erase str_Call()
    str_AllCalls = ""
    bln_SameDeal = False
    bln_FindAuctionsDone = False
    
    ' If not a user-entered hand or if '
    ' user-entered but no prior calls, '
    ' initialize the bidding variables:'
    If Left$(L_Deal & "   ", 4) <> "User" Or int_CallNum = 0 Then
        int_Turn = 0
        int_CallNum = 1
        int_BidRound = 1
        int_NumPasses = 0
    End If
    
    b_Edit.Enabled = True
    b_RestartDeal.Enabled = True
    If Not bln_PbnEvent And _
       bln_PickedHand(1) = False And _
       Not bln_HandsDealt _
    Then
        Deal
        If Not bln_FindAuctions Or bln_FindAuctionsDone Then
            ' This is done in Deal. '
            MsgBox "Stop InitiateBidding"
            For n = 1 To 4
                DisplayHand (n)
            Next
        End If
        bln_HandsDealt = True
    End If
    BidNextHand
End Sub




Sub BidNextHand()
    ' Prepares for and calls MakeBid.     '
    ' Though named BidNextHand, it may be '
    ' the first hand to bid in an auction.'
    '                                     '
    ' This routine is called first by the '
    ' InitiateBidding routine and then by '
    ' the ProcessBid routine to continue  '
    ' the bidding until it is passed out, '
    ' or in the case of specified starting'
    ' bids, until all 4 hands have been   '
    ' tested to see if they match the     '
    ' starting bids. '
    
    Dim n As Integer
    
    int_Turn = int_CallNum Mod 4
    If int_Turn = 0 Then int_Turn = 4
    
    If bln_Rebid And int_CallNum > 0 Then
        str_Call(int_CallNum - 1) = ""
        str_AllCalls = allCalls ' AllCalls is a function which removes the last Call for rebidding. '
    ElseIf str_Call(int_CallNum) <> "" Then
        int_Turn = int_Turn + 1
        If int_Turn > 4 Then
            int_Turn = 1
            int_BidRound = int_BidRound + 1
        End If
    End If
        
    bln_Rebid = False
    bln_Abort = False
    If (int_Player(int_Turn) = 1 _
        Or Not bln_HideHands) _
        And Not bln_AutoBid And _
        (mFindDeals.Checked = False Or _
         bln_FindAuctionsDone) _
        And str_MyBid <> "" _
    Then
        ' 07-20-2019: Single stepping. User has already clicked on a bid in the Bid Box '
        ' and str_MyBid = the selected bid. We are now getting Bidder to look up a bid  '
        ' which will be compared to str_MyBid to see if they match. '
        If mFindDeals.Checked = False Then 'ck_FindBids.Value = 0 Then
            bln_DontAsk4DealNum = False
            lng_DealNumSpecified = 0
        End If
        ' Let user bid first. If '
        ' other hands are hidden,'
        ' user just bids South;  '
        ' if hands aren't hidden,'
        ' user bids all hands.   '
        ' After the user picks a '
        ' bid, the program then  '
        ' bids and compares bids.'
        'b_Ok.Enabled = False
        'b_Edit.Enabled = False
        g_BidBox.Visible = True
        bln_Continue = False
        Exit Sub
    End If
            
    If int_Player(int_Turn) = 0 Then
        SetPlayerNumbers
    End If
    
    ' Let BidBase bid: '
    MakeBid int_Player(int_Turn)
End Sub


Sub ProcessBid()
    ' After the DLL makes the bid, it is  '
    ' shown to the user who can click any '
    ' of the buttons, menu options, etc.  '
    ' If he clicks "BB Bid", then this    '
    ' routine is called to accept/process '
    ' the bid, then this ends by calling  '
    ' BidNextHand unless it's passed out. '
    ' This sub is also called when doing  '
    ' practice bidding where BB has to    '
    ' make all the specified bids before  '
    ' the user having to bid. '
    
    Dim n As Integer
    Dim a As String
    Dim X As String
    Dim b As String
    Dim prompt As String
    Dim allCalls As String
    Dim response As String
        
        If bln_Abort Then
            GoTo ExitBidHands
        End If
        
        If int_Turn = 0 Then int_Turn = 1
        int_CallNum = int_CallNum + 1
            ' Bid # of the next bid when added to prior bids. '
            ' Bid # for the next bid in the current round of  '
            ' bidding is int_Turn which goes from 1 to 4.     '
                
        If Right$("      " & str_AllCalls, 6) = "-P-P-P" And Not bln_PbnEvent And Not bln_SettingUpPracBid Then
            f_Msg.Top = 200
            response = MyMsgBox("Three consecutive passes." & vbCr & _
                                "Bidding's over.", 2, "ok", "")
            Exit Sub
        End If
        
        If Not bln_Rebid Then
            If b = "P" Then
                int_NumPasses = int_NumPasses + 1 ' Counting consecutive passes. '
            Else
                int_NumPasses = 0 ' If prior call was not Pass, start count over. '
            End If
            
        ElseIf bln_Rebid Then
            str_AllCalls = ""
            str_Call(int_CallNum) = ""
            For n = 1 To int_CallNum - 1
                str_AllCalls = str_AllCalls _
                & str_Call(n) & "-"
            Next
            If Left$(str_AllCalls & " ", 1) = "-" Then
                str_AllCalls = Mid$(str_AllCalls, 2)
            End If
            If Len(str_AllCalls) > 0 Then
                str_AllCalls = _
                    Left$(str_AllCalls, _
                    Len(str_AllCalls) - 1)
            End If
        End If

    If (int_NumPasses < 3 Or int_CallNum < 5) Then
        ' Bidding continues if not 3 consecutive passes '
        ' or if it is still the first round of bidding. '
        ' Examples:  1S-P-P-P  or  P-P-P '
        allCalls = str_AllCalls ' Keep track of str_AllCalls before next bid. '
        BidNextHand ' Next *hand*, not next *deal*; e.g.: If W just bid, go to N. '
                
        bln_Rebid = False
        If Not (bln_FindAuctions Or _
                bln_SettingUpPracBid) Or _
           (Len(str_AllCalls) <> int_FoundAuctionLen And _
            bln_FindAuctionsDone) _
        Then
            ' bln_FindAuctions is just for when auctions are being assigned to a convention. '
            ' This is also done in ProcessFindBids when '
            ' the auction has been found in a deal. '
            If str_Call(int_Turn - 1) = "" Then
                ' This happens when auction has a "|" in it. '
                'MsgBox "Stop InitiateBidding" ' Disclosure is being cleared after bidding. 6-12-2020 '
                't_Disclose = ""
            End If
            'DisplayBid int_Player(int_Turn) ' This was already done in MakeBid. '
        End If
    End If
    
    If bln_PbnEvent Then
        If int_CallNum = 1 Then
            ReDim str_PBNbids(40) As String
        End If
        
        If str_PBNbids(int_CallNum) <> "" And Not bln_OffPBN Then
            If str_PBNbids(int_PBNpair, int_CallNum) <> str_Call(int_CallNum) Then
                ' Bids differ. Offer choice: '
                If str_PBNanalysis <> "" Then
                    t_Disclose = str_PBNanalysis ' Show the PBN analysis. '
                    prompt = "Refer to the PBN analysis above."
                Else
                    prompt = "Refer to the Optimum scores below."
                End If
                X = str_PBNscore ' e.g.: "NS 630"
                X = Replace$(X, "NS", "NS scoring")
                X = Replace$(X, "EW", "EW scoring")
                f_Msg.Top = 360
                a = MyMsgBox("BB bid " & str_Call(int_CallNum) & _
                       " and the PBN player bid " & _
                       str_PBNbids(int_PBNpair, int_CallNum) & "." & vbCr & _
                       prompt & vbCr & _
                       "Do you want to change to PBN's bid?" & vbCr & _
                       "(Click no to keep the BidBase bid.)", 4, "YN", "")
                If a = "yes" Then
                    str_AllCalls = allCalls & "-" & str_PBNbids(int_PBNpair, int_CallNum)
                    str_Call(int_CallNum) = str_PBNbids(int_PBNpair, int_CallNum)
                    DisplayBid int_Player(int_Turn)
                    If str_Call(int_CallNum - 1) <> "P" Then
                        int_NumPasses = 0
                    End If
                Else
                    bln_OffPBN = True
                End If
            End If
        End If
    End If
    
ExitBidHands:
    If int_NumPasses = 3 And int_CallNum < 4 Then
        'b_Edit.Enabled = False
        b_Ok.Caption = "Deal"
    End If
    bln_Abort = False
End Sub


Private Sub MakeBid(ByVal h As Integer)
    'CheckBattery ' Prefer sticking this here over using a Timer. '
    
    ' This routine is called by BidNextHand.'
    ' It calls the BidBaseDB.DLL which uses '
    ' the input hand, etc., to find a bid in'
    ' the database.                         '

    Dim b As String
    Dim c As Integer
    Dim d As String
    Dim s As Integer
    Dim v As String
    Dim i As Long
    Dim j As Long
    Dim p As Long
    Dim pn As Long
    Dim BON As String
    Dim ID As Long
    Dim pr As String
    
    t_Disclose = ""
    If Not Val(L_Deal) = lng_DealNum And Not bln_PbnEvent Then
        L_Deal = "" & lng_DealNum
    End If
    
    Do While str_Call(int_Turn - 1) = "" And int_Turn > 1
        int_Turn = int_Turn - 1
        'MsgBox "Stop InitiateBidding"
    Loop
    
    'Submit specs to DLL:
    With NewBid
        .DBPathAndName = str_DBPathAndName
        .PlayersHand = str_Hands(h)
        
        pr = str_AllCalls
        
        ' Split off ace/king-asking
        If int_CallNum > 2 Then
            i = InStr(pr, "@")
        End If
        If i > 2 Then
            pr = Mid$(pr, i - 2)
            ' 4D@Qx = a control bid, 4N@xy = Ace Asking (x=suit#,
            '         and y=2nd suit # as for 6-key ask)        '
            ' Ex: "1H-P-1S-2D-2S-P-4N@S0-P" - just submit 4N-P  '
            ' and .TrumpSuit = 4 (spades). '
            ' We want to maintain str_AllCalls to display the   '
            ' complete bidding, but slam bidding entries start  '
            ' Prior Bids with the Ace Asking or Control showing '
            ' bid. '
            d = Mid$(pr, 4, 2) ' Either Q# or s# where s=strain. '
            'j = Val(Mid$(pr, 2, 1))
            If Left$(d, 1) = "Q" Then
                MsgBox "Q lookup needs checking."
                .TrumpSuit = Val(Mid$(d, 2) & 9) ' 9 is just code for "Q". '
                ' Trump suit must be known in order to '
                ' skip over it when control bidding.   '
            ElseIf Left$(d, 1) = "N" Then
                ' Opp's suit when partner '
                ' has overcalled it 1N.   '
                MsgBox "N lookup needs checking."
                .TrumpSuit = Val(d)
            'ElseIf j > 0 Then
            '    .TrumpSuit = j
            Else
                .TrumpSuit = Val(d)
                'Debug.Print .TrumpSuit, d
            End If
            d = ""
            ' Continuing the example above: .PriorCalls = "4N-P": '
            .PriorCalls = pr ' Mid$(str_AllCalls, i - 2)
        Else
            .PriorCalls = str_AllCalls
        End If
        
        If bln_AceAsking Then
            .TrumpSuit = int_TrumpSuit
            i = InStr(str_AllCalls, str_AceAskingCall)
            .PriorCalls = Mid$(str_AllCalls, i)
            ' Should not be possible for str_AceAskingCall not to be in PriorCalls. '
        End If
        
        ' h = current bidder. p = h's partner '
'Stop
        If h > 2 Then
            p = h - 2
        Else
            p = h + 2
        End If
        
        If int_OptAvg(h) = 0 Then
            ' See if pard is 0 also: '
            If int_OptAvg(p) > 6 Then
                ' Shouldn't be possible. '
                int_OptAvg(p) = int_OptAvg(h)
            End If
        End If
        
        If int_OptAvg(p) = 0 Then
            int_OptAvg(p) = int_OptAvg(h)
        End If
        
        .OptimumAvg = int_OptAvg(h)
        .Opt4PardAvg = int_OptAvg(p)
        .OptimumOrig = int_OptOrig(h)
        .Opt4PardOrig = int_OptOrig(p)
        
        i = h - 1
        If i = 0 Then i = 4
        .Opt4OppAvg = int_OptAvg(i)
        .Opt4OppOrig = int_OptOrig(i)
        
        ' I can't find how to pass an array to a DLL in VB6 '
        ' so I am going to store the data from the array in '
        ' a file and load the data into an array in the DLL.'
        ' The array is int_OptByStrain(4, 5, 2)   '
        ' Where 4 = player #: S=1 E=2 N=3 W=4     '
        '       5 = strain #: C=1 D=2 H=3 S=4 N=5 '
        '       2 = # iterations: 1=11, 2=101,    '
        '       or 0 for the original DDA results '
        ' Example: Opt for North's NT for 11      '
        '          iterations: (3, 5, 1) = 2      '
    If 1 = 2 Then
        Dim fi As Integer
        fi = FreeFile
        Close #fi
        Open "DDAresults.txt" For Output As #fi
        For i = 0 To 2 ' 0=no avg., 1=11 avg, 2=101 avg '
            For p = 1 To 4 ' 1=S, 2=W, 3=N, 4=E '
                
                If p = 1 Then
                    pn = 1 ' S '
                ElseIf p = 2 Then
                    pn = 3 ' N '
                ElseIf p = 3 Then
                    pn = 2 ' W '
                Else
                    pn = 4 ' E '
                End If
                For s = 1 To 5
                    Write #fi, int_OptByStrain(pn, s, i);
                Next
                Print #fi, ""
                Print #fi, str_DDAOutput(pn)
            Next
        Next
        Close #fi
        'Stop
    End If
        Dim vn As Integer ' Vul.# '
        
        ' This probably shouldn't be done when practicing or making practice files. '
        ' The Vul shown on screen is for NS. '
        If (int_Player(int_Turn) = 1 Or int_Player(int_Turn) = 3) Then
            ' NS : '
            If L_Vul = "Unfav." Then
                v = "U"
            ElseIf L_Vul = "Fav." Then
                v = "F"
            ElseIf L_Vul = "Both" Then
                v = "B"
            End If
        Else
            ' Reverse label for EW: '
            If L_Vul = "Unfav." Then
                v = "F"
            ElseIf L_Vul = "Fav." Then
                v = "U"
            ElseIf L_Vul = "Both" Then
                v = "B"
            End If
        End If
        .Vul = v
        .BidToAnalyze = Trim$(str_BidToAnalyze)
        .GetDisclosures = True
        .SaveDealNum = lng_DealNum
        
        ' Get bid: '
        ' If this doesn't work, quit and check Project - References '
        ' to make sure that BidBaseDb.vbp or .DLL (whichever you are'
        ' using) is checked. '
        b = .Bid
        c = .ErrCode
        d = .Disclosure
        If b <> "" Then
            str_BON = .BON
        End If
        lng_ID = .ID
        
        If b = "" And (c < 4 Or c = 97) Then
            ' If c=1 or 2, no bid was found. '
            b = "P"
            c = 0
        End If
        
        If bln_Abort Then
            str_Call(int_CallNum) = b
            Exit Sub
        End If
        
        If int_CallNum = 0 Then int_CallNum = 1
        If str_Call(int_CallNum) <> "" Then
            int_CallNum = int_CallNum + 1
        End If
        str_Call(int_CallNum) = b
'Stop
        If (bln_SearchForMoreDeals Or _
            bln_FindAuctions Or _
            bln_PracticeFoundBids) And _
           Not bln_FindAuctionsDone _
        Then
            ' See if deal matches any of the specified bids. '
            bln_FindAuctions = True ' If practicing, get the specified bids. '
            ProcessFindBids b, c
            If Timer1.Enabled Then Exit Sub
            If bln_FindAuctions = False Then Exit Sub
                
        ElseIf (int_CallNum = 1 And _
                int_BidRound = 1) Or _
               str_AllCalls = "" _
        Then
            ' AllCalls is the complete bidding sequence. '
            str_AllCalls = b
        
        Else
            str_AllCalls = str_AllCalls & "-" & b
            If Left$(str_AllCalls, 1) = "-" Then
                MsgBox "Stop MakeBid 1"
                ' Haven't tried to track down where '
                ' the leading dash is coming from.  '
                str_AllCalls = Mid$(str_AllCalls, 2)
            End If
        End If
        
        If c > 0 And c <> 2 Then
            t_Disclose = .ErrMsg
        End If
            
        If Not (bln_FindAuctions Or _
                bln_PracticeFoundBids) Or _
               (Len(str_AllCalls) <> int_FoundAuctionLen And _
                bln_FindAuctionsDone) _
        Then
            If Trim$(d) <> "" And _
               b <> "P" And _
               t_Disclose <> d _
            Then ' Don't show disclosure for Pass '
                If Not (bln_FindAuctions And _
                        (Not bln_FindAuctionsDone Or _
                         int_Player(int_Turn) = 1)) _
                Then
                    t_Disclose = d
                Else
                    MsgBox "Stop MakeBid 2" ' Wasn't FindAuctions already screened out? '
                End If
            End If
        End If
                        
        If Not bln_AutoBid And _
           (int_Player(int_Turn) = 1 Or _
            Not bln_HideHands) And _
           Not (bln_FindAuctions And _
                Not bln_FindAuctionsDone) And _
           str_MyBid <> "" _
        Then
            c = Len(d)
            If (Trim$(Right$(str_AllCalls, Len(b))) = b Or b = "P") Then
                b = "       ***** BidBase bid " _
                    & "the same as you. *****"
            Else
                b = "       ***** You bid " _
                    & str_AllCalls _
                    & " and BidBase bid " _
                    & str_Call(int_CallNum) & ". *****"
            End If
                
            If Not (bln_FindAuctions And _
                    Not bln_FindAuctionsDone) _
            Then
                t_Disclose = b & vbCrLf & t_Disclose
            End If
        End If
    End With
    
    ' For regular bidding practice, the first bid was not being shown. '
    DisplayBid 1
    
    If str_BidToAnalyze <> "" Then
        bln_Analysis = False
        str_BidToAnalyze = ""
        b_ShowAnalysis.Caption = "Rebid w/Analysis"
    End If
    b_Ok.Enabled = True
        
End Sub
            
            
Private Sub ProcessFindBids(b As String, c As Integer)
    ' This is called by MakeBid. b = the bid returned by '
    ' the Bidder DLL.  c = Bidder error code, if any.    '
    Dim p As String ' prior calls (e.g.: 1H) '
    Dim p2 As String ' prior calls + current call (e.g.: 1H-P) '
    Dim i As Long
    Dim n As Integer
    Dim X As Integer
    Dim tryAgain As Boolean
    Dim auc As String  ' auction being tested '
    Dim xbid As String
            
    If bln_FindAuctionsDone Or _
       (Not bln_FindAuctions And _
        Not bln_PracticeFoundBids) _
    Then
        Exit Sub
    End If
    
    If int_AuctionNumber > int_NumAuctionSpecs Or str_FindAuctions(int_AuctionNumber) = "" Then
        int_AuctionNumber = 1
    End If
    
    ' Say specified auctions to find are as follows:    '
    '    a. 1C-P-1H-P-1N-P-2D-P                         '
    '    b. 1C-P-1S-P-1N-P-2D-P                         '
    '    c. 1D-P-1H-P-1N-P-2D-P                         '
    '    d. 1D-P-1S-P-1N-P-2D-P                         '
    ' 1. Start with South as dealer & get his 1st bid.  '
    ' 2. Compare his first bid to the first bid in a.   '
    ' 3. If it doesn't match, compare it to b, c, d.    '
    '    If no match is found for any spec, skip to 7.  '
    ' 4. If a match is found to the first bid of, say,  '
    '       c. 1D..., get the bid for West and see if   '
    '       South & West's bids match the first 2 bids  '
    '       (1D-P) in c and if not, then try spec d.    '
    '       If no matches to the first two bids of b-d  '
    '       are found, then this deal will not match.   '
    '       Go to step 7 only North is the next dealer. '
    ' 5. If the first 2 bids match a spec, keep the spec'
    '       number (or c in this example) and get the   '
    '       next bid (North) and repeat step 4 for each '
    '       player in turn (continuing with South when  '
    '       it is his turn again) until all the current '
    '       bids match a specified bid, then...         '
    ' 6. Display all the bids and wait for the user to  '
    '       click "Bid 's OK".                          '
    ' 7. If South's first bid (say, 1H). does not match '
    '       the first bid of any of the auctions, make  '
    '       West the dealer and repeat the above steps  '
    '       to see if North's first bid matches any of  '
    '       the first bids, and if not, then try East.  '
    ' 8. If none of the sequence of bids using any of   '
    '       the players as dealer match the auctions... '
    ' 9. Get the next deal and repeat the above steps.  '
    
    ' This routine is also used when practicing a conv. '
    ' to get the bidding up to the end of a specified   '
    ' auction, at which point the user/program does the '
    ' bidding and this routine is not used again until  '
    ' the next saved deal number is called up. '
    ' The dealer is saved in each .Deals file entry and '
    ' we could start this sub with that dealer instead  '
    ' of with south, but that would complicate this sub.'

    p = str_AllCalls

'If p = "1N" And b <> "p" Then
'    For i = 1 To 4: DisplayHand i: Next
'    DisplayBid 1
'    Stop
'End If
'If p = "1N" And b <> "P" Then Stop
    
    If p <> "" And (p <> b Or b = "P") Then
        p2 = p & "-" & b
    Else
        If 1 = 2 And Left$(p & b & " ", 2) <> "1N" Then
            If int_Dealer = 4 Then ' Or Left$(str_FindAuctions(int_AuctionNumber), 2) = "1N" Then
                'Stop
                int_Dealer = 1
                int_Turn = 1
                str_TimerJob = "Go to next deal"
                GoTo StartNextJob
            Else
                GoTo NextDealer
            End If
        End If
        ' Leave the next line
        p2 = b
    End If

' The following is just for making quizzes. '
If Right$(str_FindAuctions(int_AuctionNumber), 2) <> "-A" And _
   InStr(str_FindAuctions(int_AuctionNumber), "|") = 0 _
Then
    str_FindAuctions(int_AuctionNumber) = str_FindAuctions(int_AuctionNumber) & "-A"
End If

    If str_AllCalls = "" Then
        If Left$(str_FindAuctions(int_AuctionNumber), Len(b)) = b Then
            ' First bid matches the auction. '
            str_AllCalls = b
            str_Call(int_Turn) = b
            str_TimerCallingSub = "ProcessFindBids2"
            str_TimerJob = "Get next bid"
            Timer1.Enabled = True
            Exit Sub
        End If
        ' First bid for the auction doesn't match. '
        GoTo NextAuction
    End If

    If str_AllCalls = "" And int_Turn = 1 Then
        int_AuctionNumber = 1
    Else
        ' Say that auction #1 is 1H-P and '
        ' #2 is 1S-P and str_AllCalls is  '
        ' 1S and this is the 2nd bid made,'
        ' we don't need to test 1H-P again.
        ' int_AuctionNumber would be 2 and'
        ' int_Turn would be 2. '
    End If
    
    Do While Not bln_FindAuctionsDone
        ' Look for an "x" as in 1C-x2S to exclude 2S: '
        ' but not in 1N-P-
        n = InStr(str_FindAuctions(int_AuctionNumber), "x")
        If n > 0 And n <= Len(p2) + 1 Then
            ' Example: b=1H, 1C-x2S then Len("1C")=4-2  '
            ' So the x is on the next bid in the spec.  '
            ' If b <> the x-bid, then replace the x-bid '
            ' in the spec with the current bid: 1C-1H   '
            xbid = RTrim$(Mid$(str_FindAuctions(int_AuctionNumber) & "  ", n + 1, 2))
            If b <> Replace(xbid, "-", "") Then
                ' b is not the call specified NOT to be bid.  '
                ' Example: "1H-xP" & b="2N", replace P w/b 2N.'
                auc = Left$(str_FindAuctions(int_AuctionNumber), n - 1) & b
            Else ' b is the "x" bid;  go to next auction. '
                GoTo NextAuction
            End If
        ElseIf Left$(str_FindAuctions(int_AuctionNumber), Len(p & "-A")) = p & "-A" Then
            ' Accept any bid for "A". '
            auc = Replace(str_FindAuctions(int_AuctionNumber), "A", b)
        Else
            auc = str_FindAuctions(int_AuctionNumber)
        End If
        
        auc = Replace(auc, "|", "")
        ' The vertical bar says not to display the bid(s) after it, but '
        ' when searching for an auction, that bid has to be included.   '
        
        Do
        n = InStr(auc, "A")
        If n > 0 And Len(p2) >= Len(auc) Then
            If Mid$(p2, n, 1) < "1" Then
                Mid$(auc, n, 1) = Mid$(p2, n, 1)
            ElseIf Mid$(p2, n + 1, 1) = "-" Then
                auc = Left$(auc, n - 1) & Mid$(p2, n, 2) & Mid$(auc, n + 2)
            Else
                auc = Left$(auc, n - 1) & Mid$(p2, n, 3) & Mid$(auc, n + 3)
            End If
        Else
            Exit Do
        End If
        Loop While InStr(auc, "A") > 0
        
        If auc = p2 And str_AllCalls <> "" Then
            ' Added the last part because it was taking "1H" = "1H" '
            ' when the original auction
            
            'For i = 1 To 4: DisplayHand i: Next
            'DisplayBid 1 ' Single-step this or it won't display the bidding. '
            'If b = "P" Then Stop
            If p2 = "1N-P-2C" Then
                ' Only use 2 in every 10 Stayman bids. '
                'Stop
                stayman = stayman + 1
                'Stop
                If stayman = 4 Or stayman = 9 Then
                    ' Use it. '
                Else
                    If stayman = 10 Then
                        stayman = 0
                    End If
                    int_Dealer = 1
                    int_Turn = 1
                    str_TimerJob = "Go to next deal"
                    GoTo StartNextJob
                End If
            End If
            MakePracticeWebPage
            
            'If lng_DealNum = 4060 Then Stop
            If bln_UpdateCnvFiles Then
                GoTo Updating ' This is for creating quizzes from existing .Deals files. '
            End If
            
            bln_FindAuctionsDone = True
            int_FoundAuctionLen = Len(p2)
            If InStr(str_FindAuctions(int_AuctionNumber), "|") > 0 Then
                ' When excluding the last bid after looking up a bid for it, '
                ' we don't want to display the bid because that is for the   '
                ' user to figure out. See the docs. '
                str_Call(int_Turn) = ""
                t_Disclose = ""
            Else
                str_AllCalls = p2
            End If
            
            If bln_SearchForMoreDeals Then
                lng_DealNum2Practice = lng_DealNum
                int_DealFileEntryNum = int_DealFileEntryNum + 1
                If int_DealFileEntryNum < 2 Then
                    int_DealFileEntryNum = 2
                End If
            'DisplayBid 1
            'For i = 1 To 4: DisplayHand i: Next
                PutPracticeDeal
                lng_NumPracticeEntries = lng_NumPracticeEntries + 1
                If lng_NumPracticeEntries < 2 Then lng_NumPracticeEntries = 2

                ExportPracticeDealList
                
                If lng_NumPracticeEntries > 100 And bln_FindAuctions Then
                    DisplayBid 1
                    bln_FindAuctionsDone = True
                    bln_FindAuctions = False
                    mFindDeals.Checked = False
                    b_Edit.Enabled = True
                    
                    ' Save # entries & "2" as the first entry. '
                    Dim fiNum As Integer
                    Dim lastDeal As lastDeal
                    fiNum = FreeFile
                    Open str_PracticeDealsFile For Random As fiNum Len = 6
                        lastDeal.NumEntries = lng_NumPracticeEntries
                        lastDeal.LastNumUsed = 2
                        Put #fiNum, 1, lastDeal
                    Close #fiNum
                    
                    'ExportPracticeDealList ' Deals are already added as found. '
                    
                    MsgBox "100 new entries have been saved." & Chr$(13) & _
                           "Click Practice in the Conventions menu to use them."
                    Me.Height = 9330
                    Exit Sub
                End If
                
                'n = int_Turn + 1
                'If n = 5 Then n = 1
                'Debug.Print lng_NumPracticeEntries; "  "; lng_DealNum; "  "; str_Hands(int_Player(int_Turn)); "  "; b
                'Beep
Updating:
                int_AuctionNumber = 1
                int_Dealer = 1
                str_TimerCallingSub = "ProcessFindBids"
                str_TimerJob = "Go to next deal"
                Timer1.Enabled = True
                Exit Sub
            End If
            
            If bln_PracticeFoundBids Then
                ' With the specified auction complete, '
                ' we want the next bidder to be South. '
                ' int_Player(int_Turn) is the last one '
                ' to bid. We want that to be East so   '
                ' that South will be next to bid.      '
                ' If last was 1, rotate CCW 1 time.    '
                ' If last was 2, rotate CCW 2 times.   '
                ' if last was 3, rotate CCW 3 times.   '
                ' if last was 4, don't rotate.         '
                ' If the auction has "|" in it, the    '
                ' last bid will be blanked out so we   '
                ' only want to rotate the hands CCW    '
                ' int_Player(int_Turn) - 1.  Example:  '
                ' If 1H-P-|3C, it will be shown as     '
                ' 1H-P and South will have to make the '
                ' 3C bid, so if West was the dealer,   '
                ' North will have to be made the dealer'
                ' so that South = int_Turn=3. Normally,'
                ' if W:1H-N:P-E:3C with S being next   '
                ' to bid, no rotation would be needed. '
                ' With E:3C erased, we want N to be the'
                ' one to bid 1H instead of W, which    '
                ' requires 3 CCW rotations: W->S->E->N.'
                ' This adjustment is made in mRotate.  '
                
                ' We want int_LastBidder to be 4 so that 1 is next to bid. '
                X = InStr(str_FindAuctions(int_AuctionNumber), "|")
                If X > 0 Then
                    int_LastBidder = int_Player(int_Turn) - 1
                    If int_LastBidder = 0 Then
                        int_LastBidder = 4
                    End If
                Else
                    int_LastBidder = int_Player(int_Turn)
                End If
                If int_LastBidder < 4 Then
                    mRotate_Click int_LastBidder 'int_Player(int_Turn)
                End If
                ShowDDA False
                bln_PracDealIsSetup = True
            End If
            Exit Sub
            ' A specified auction was found. '
            ' From here on out the user or   '
            ' the program will make the bids '
            ' in the current deal.  This sub '
            ' will not be used until the next'
            ' practice deal is brought up.   '
        End If
        
        If Left$(auc, Len(p2)) = p2 Then ' #5 '
            ' If p2 matches the first part of the auction   '
            ' like p2="1H" and str_FindAuctions="1H-P" then '
            ' let the computer make the next bid. '
            str_AllCalls = p2
            str_TimerCallingSub = "ProcessFindBids2"
            str_TimerJob = "Get next bid"
            Timer1.Enabled = True
            Exit Sub
        End If

NextAuction:
        ' At this point, the current bid does '
        ' not match the current auction.  If  '
        ' this is the last auction, go to the '
        ' next player as dealer. If this was  '
        ' the last player go to the next deal.'
        
        int_AuctionNumber = int_AuctionNumber + 1

        If int_AuctionNumber > int_NumAuctionSpecs Then
            ' No matches found for any of the specified bids. '
            If bln_PracticeFoundBids Then
                ' Practice for a deal should match the  '
                ' dealer saved in the Deals file. Don't '
                ' bother trying other players. '
                MsgBox "Bidding for deal #" & lng_DealNum & " does not match" & vbCrLf & _
                       "the specified auctions." & vbCrLf & _
                       "Most likely bid specifications in the" & vbCrLf & _
                       "database have changed." & vbCrLf & _
                       "Skip to the next saved deal? (No to exit.)"
                'Debug.Print lng_DealNum; " no longer matches."
                DelPracticeDeal
                str_TimerJob = "Go to next deal"
            
            ElseIf int_Turn = 4 Or (bln_FindAuctions And int_Dealer = 4) Then
                ' Each player already tried the specified bids as dealer. '
                ' Go to the next deal. '
                'If int_Dealer < 4 Then
                '    int_Dealer = int_Dealer + 1
                'Else
                    int_Dealer = 1
                'End If
                str_TimerJob = "Go to next deal"
            Else
NextDealer:
                ' Advance to the next player as dealer. '
                int_Dealer = int_Dealer + 1
                If int_Dealer = 5 Then int_Dealer = 1
                str_TimerJob = "Restart same deal"
            End If
StartNextJob:
            int_DealerSpecified = int_Dealer
            bln_FindAuctionsDone = False
            str_AllCalls = ""
            ' int_Turn = 1 - This is done in FinishDealSetup '
            ' SetPlayerNumbers - Ditto '
            ' int_AuctionNumber = 1 - Ditto '
            str_TimerCallingSub = "ProcessFindBids3"
            If str_TimerJob = "" Then MsgBox "Stop ProcessFindBids 2"
            Timer1.Enabled = True
            Exit Sub
        End If
    Loop
    
    'int_AuctionNumber = i ' This should be getting the next bids spec. '

    If bln_FindAuctions And Not bln_FindAuctionsDone Then
        If Left$(str_FindAuctions(int_AuctionNumber), Len(p2)) = p2 Then
            ProcessBid
            Exit Sub
        Else
            MsgBox "Stop ProcessFindBids 3"
            ' get next deal '
        End If
    End If
End Sub


Private Sub DisplayBid(ByVal h As Integer)
    ' The Bid Display grid is set up like this:  '
    ' Col. 1=W, 2=N, 3=E, 4=S to follow standard '
    ' bridge publication standards.              '
    '                                            '
    ' W deals, all 1st-round bids are on 1st row.'
    ' N deals, N/E/S are on the 1st row, but     '
    '    W's first call is on the 2nd row.       '
    ' E deals, E/S = 1st row on 1st round, but   '
    '    W/N = 2nd row on 1st round.             '
    ' S deals, S is on 1st row, others on 2nd    '
    '                                            '
    ' The h parameter is the player # (S=1, etc.)'
    Dim i As Integer
    Dim c As Integer
    Dim r As Integer
    Dim n As Integer
    Dim b As String
    
    'If bln_FindAuctions And Not bln_FindAuctionsDone Then Exit Sub
    
    SetUpBidDisplayGrid
    
With g_Bidding
    i = int_Dealer - 1 ' S is column 0 but Dealer 1, etc. '
    .row = 1
    n = 0
    For r = 1 To 7
        For c = 0 To 3
            If c = i Then
                .col = c
                n = n + 1
                If InStr(str_Call(n), "@") = 3 Then
                    .Text = Left$(str_Call(n), 2)
                Else
                    .Text = str_Call(n)
                End If
                If str_Call(n) = "" Then
                    GoTo ExitDisplayBid
                End If
                i = i + 1
                If i > 3 Then
                    i = 0
                    .row = .row + 1
                    If .row > 6 Then MsgBox "Stop DisplasyhBid 1"
                    Exit For
                End If
            End If
        Next
    Next

    ' Set column: '
    If h > 1 Then
        .col = h - 2 ' Columns are 0-3
    Else
        .col = 3 ' South is last col.
    End If

    ' Set row:    '
    If int_Player(1) = 2 Then ' West dealt
        .row = int_BidRound
    
    ElseIf int_Player(1) = 3 Then ' North dealt
        If h = 2 Then
            .row = int_BidRound + 1
        Else
            .row = int_BidRound
        End If
    
    ElseIf int_Player(1) = 4 Then ' East dealt
        If h = 2 Or h = 3 Then
            .row = int_BidRound + 1
        Else
            .row = int_BidRound
        End If
    
    Else  ' South dealt
        If h > 1 Then
            .row = int_BidRound + 1
        Else
            .row = int_BidRound
        End If
    End If
    
    b = str_Call(int_CallNum)
    i = InStr(b, "@")
    If i > 0 Then
        b = Left$(b, i - 1)
    End If
    .Text = b
    If .Text <> "P" Then
        int_NumPasses = 0
    End If

End With
  
ExitDisplayBid:
    If bln_Analysis Then
        b_ShowAnalysis.Caption = "Rebid w/ &Analysis"
        b_Ok.Caption = "&BB Bid"
        bln_Analysis = False
    End If

End Sub


Private Sub SetPlayerNumbers()
    ' The int_Plyaer() variable lets us keep '
    ' track of which actual player (1-4) is  '
    ' in which bidding position.             '
    
    If int_Dealer = 0 Then
        int_Dealer = InStr("SWNE", Left$(L_Dealer & "S", 1))
        int_DealerSpecified = int_Dealer
    End If
    
    If ck_KeepDealer.Value = 0 Then
    
        If int_DealerSpecified > 0 Then
            int_Player(1) = int_DealerSpecified
            'If Not bln_PbnEvent Then int_DealerSpecified = 0
        Else
            int_Dealer = int_Dealer + 1
            If int_Dealer > 4 Then int_Dealer = 1
            int_Player(1) = int_Dealer
        End If
  
        If int_Player(1) = 0 Then
            int_Player(1) = 1
        End If
    Else
        int_Player(1) = int_Dealer
    End If

    If Not bln_PbnEvent Then
        L_Dealer.Visible = True
        If int_Dealer < 1 Then
            If Val(L_Dealer) = 0 Then
                int_Dealer = 1
            Else
                int_Dealer = Val(L_Dealer) + 1
                If int_Dealer = 5 Then int_Dealer = 1
            End If
        End If
        L_Dealer = Mid$("SWNE", "" & int_Dealer, 1)
        ck_KeepDealer.Visible = True
    End If
    
    Select Case int_Player(1)
        Case 1
            int_Player(2) = 2
            int_Player(3) = 3
            int_Player(4) = 4
  
        Case 2
            int_Player(2) = 3
            int_Player(3) = 4
            int_Player(4) = 1
  
        Case 3
            int_Player(2) = 4
            int_Player(3) = 1
            int_Player(4) = 2
        
        Case 4
            int_Player(2) = 1
            int_Player(3) = 2
            int_Player(4) = 3
    End Select
    
    Dim i As Long
    For i = 1 To 4
        If int_Player(i) < 1 Then int_Player(i) = 1
        str_Player(i) = Trim$(Mid$("South West  North East ", _
                                   (int_Player(i) - 1) * 6 + 1, 5))
    Next
End Sub


Private Sub LoadCfg()
    ' Load the configuration file which '
    ' is saved when the program ends.   '
    Dim f As String
    Dim p As String
    Dim fiNum As Integer
    Dim lng_PlayScrLeft As Long
    Dim lng_PlayScrTop As Long
    Dim i As Long
    Dim findBids As String
    Dim b As String
    
    str_Suits(1) = "Clubs"
    str_Suits(2) = "Diamonds"
    str_Suits(3) = "Hearts"
    str_Suits(4) = "Spades"
    
    ChDir str_BBpath
    
    If Dir$(str_BBpath & "\BBPractice.cfg") = "" Then
        f = str_BBpath & "\BBPractice new.cfg"
        If Dir$(f) = "" Then
            MsgBox str_BBpath & "\BBPractice.cfg is missing." & vbCr & _
                   "Default settings will be used."
            Exit Sub
        End If
        ' Load the settings from the initial file. '
        ' A Practice.cfg file will be created when '
        ' quitting the program. '
    Else
        f = str_BBpath & "\BBPractice.cfg"
    End If
    
    fiNum = FreeFile
    Open f For Input As #fiNum
        Input #fiNum, lng_DealNum
        lng_DealNum = lng_DealNum - 1
        Input #fiNum, lng_PlayScrLeft, lng_PlayScrTop
        Input #fiNum, bln_AutoBid
        Input #fiNum, bln_HideHands
        Input #fiNum, int_DealerSpecified
            If int_DealerSpecified = 0 Then int_DealerSpecified = 1
            L_Dealer = Mid$("SWNE", int_DealerSpecified, 1)
            L_Dealer.Visible = True
            int_Dealer = int_DealerSpecified
        Input #fiNum, i
            ck_KeepDealer.Value = i
            ck_KeepDealer.Visible = True
        'Input #fiNum, str_PbnEventFile
        'Input #fiNum, int_PbnDealNum
    Close #fiNum
    
    ' Other initializations: '
    Randomize
    CARDWIDTH = 100
    CARDHEIGHT = 200
    ' Set ScaleMode to Pixel(3) so the form and '
    ' card-display DLL use the same coordinates '
    Me.ScaleMode = 3
    bln_NoPaint = True
    Me.Show
    Me.Move lng_PlayScrLeft, lng_PlayScrTop
    
    'mAutoBid_Click
    bln_NoPaint = True
    mHide_Click
    
    Set NewBid = New BidBaseDB.Bidder
    
    str_Dealer(1) = "South"
    str_Dealer(2) = "West"
    str_Dealer(3) = "North"
    str_Dealer(4) = "East"
    
    SetUpBidDisplayGrid
    
    'If str_PbnEventFile <> "" Then
    '    bln_PbnEvent = True
    '    If int_PbnDealNum = 0 Then int_PbnDealNum = 1
    '    GetPbnDeal int_PbnDealNum
    '    mLoadPbnDealNum.Enabled = True
    'End If
End Sub


'Private Sub CheckBattery()
    ' When I'm running on battery, I usually '
    ' don't have room to see a charge-left   '
    ' display, so I like having it in my     '
    ' program. That's what this routine does.'
    
    'If SysInfo.ACStatus = 1 Then 'AC power
        'lblBatt.Visible = False
        'lblBattStat.Visible = False
    '    lblBattTime.Visible = False
    '    lblTime.Visible = False
    '    Exit Sub
    'End If
  
    'lblBatt.Visible = True
    'lblBattStat.Visible = True
    'lblBattTime.Visible = True
    'lblTime.Visible = True
  
    'Dim PerCentLeft As String
    
    ' I don't have much room on this form, '
    ' so I just show the time left. This   '
    ' code shows the % remaining, and I    '
    ' left it here in case someone else    '
    ' wants to use it.                     '
    'If SysInfo.BatteryLifePercent <> 255 Then
        'PerCentLeft = SysInfo.BatteryLifePercent
        'If PerCentLeft > 49 Then
        '    lblBatt.ForeColor = vbGreen
        'Else
        '    lblBatt.ForeColor = vbRed
        'End If
        'lblBatt.Caption = PerCentLeft & "%"
    'Else
        'lblBatt.Caption = "Unknown"
    'End If
  
    'If SysInfo.BatteryLifeTime <> &HFFFFFFFF Then
    '    lblTime.Caption = Format((TimeSerial(0, 0, _
    '    SysInfo.BatteryLifeTime)), "h:mm")
    'Else
    '    lblTime.Caption = "Unknown"
    'End If
'End Sub


Private Sub Timer1_Timer()
    ' When doing Find Bids, the program has to '
    ' keep going in circles until it finds a   '
    ' deal which matches the bidding specs.    '
    ' This can lead to "Out of stack space" if '
    ' it loops too many times, so I'm using the'
    ' timer to start the next deal without     '
    ' looping. The timer interval should give  '
    ' it time to exit whatever Sub it's in. '
    Timer1.Enabled = False
    If str_TimerJob = "Cancel entering hands" Then
        str_TimerJob = ""
        Dim l As String
        l = L_Deal
        If InStr(l, "User") = 1 Then
            l = RTrim$(Mid$(l & Space$(10), 8)) & ".txt"
            LoadDeal l
        Else
            lng_DealNumSpecified = lng_DealNum
            FillSuits
            bln_DontAsk4DealNum = True
            int_DealerSpecified = int_Dealer
            b_Deal_Click
        End If
        
    ElseIf str_TimerJob = "Get next bid" Then
        str_TimerJob = ""
        ProcessBid
    ElseIf str_TimerJob = "Restart same deal" Then
        str_TimerJob = ""
        b_RestartDeal_Click
    ElseIf str_TimerJob = "Go to next deal" Then
        str_TimerJob = ""
        bln_HandsDealt = False
        b_DealBackFwd_Click 1
    End If
End Sub


Private Sub cmb_FindDeals_KeyPress(KeyAscii As Integer)
    Dim c As String
    Dim bids As String
    Dim suit As String
    Dim priorChar As String
    
    If KeyAscii = 13 Then 'Enter'
        cmb_FindDeals.AddItem cmb_FindDeals
        cmb_FindDeals.Text = ""
        
        If int_NumAuctionSpecs < 6 Then
            cmb_FindDeals.Height = 95
            Me.Height = 10770
        ElseIf int_NumAuctionSpecs = 6 Then
            cmb_FindDeals.Height = 110
            Me.Height = 10970
        Else
            cmb_FindDeals.Height = 125
            Me.Height = 11140
        End If
        Exit Sub
    End If
    
    c = UCase$(Chr$(KeyAscii))
    If c = "X" Then c = "x"
    KeyAscii = Asc(c)
    
    If KeyAscii = 32 Then ' convert space to - '
        KeyAscii = Asc("-")
        Exit Sub
    End If
    
    'bids = cmb_FindDeals  Load specs into array.
    priorChar = Trim$(Right$(" " & cmb_FindDeals.Text, 1))
    
    suit = InStr("CDHSNPDR", c) ' Okay, so PassDoubleRedouble is not really a "suit"... '
    If suit = 0 And InStr("1234567-| , xA", c) = 0 And KeyAscii <> 8 Then
        MsgBox "'" & c & "' is not a valid bidding character."
        KeyAscii = 0
    ElseIf c = " " And priorChar <> "," Then
        ' A space is allowed only after a comma. '
        KeyAscii = 0
    ElseIf c = "," Then
        If InStr("CDHSNPDR", priorChar) = 0 Then
            KeyAscii = 0 ' Comma only allowed after "CDHSNPDR". '
        Else
            cmb_FindDeals = cmb_FindDeals & ", "
            KeyAscii = 0
            cmb_FindDeals.SelStart = Len(cmb_FindDeals) + 1
        End If
    ElseIf Val(priorChar) = 0 And suit > 0 And suit < 6 And c <> "D" Then
        ' If priorChar not a number and '
        ' this key is a suit, then BAD! '
        KeyAscii = 0
    End If
    ' Anything that is not KeyAscii = 0 by this point is Okay by me. '
End Sub


Public Sub ShowDDA(DDA_Only As Boolean)
    Dim pbn As ddTableDealPBN
    Dim thePBN As String
    Dim X() As Byte
    Dim p As Integer ' player '
    Dim s As Integer ' suit '
    Dim c As Integer ' card '
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim res As ddTableResults
    Dim suit As Integer
    Dim player As Integer
    'Dim sOutput(3) As String
    Dim turn(3) As Integer
    Dim s1 As String ' For splitting DDA. '
    Dim s2() As String ' ditto '
    Dim opt As String
        
    Dim sStart As Date
    Dim sEnd As Date
    Dim nSecs As Long
    sStart = DateTime.Now
    
    Erase str_DDA()
    Erase str_DDAOutput()
    If mAverageDDA.Checked = False Then Erase int_DDA()

    If bln_SearchForMoreDeals And Not DDA_Only Then Exit Sub
    
    Select Case int_Dealer
        Case 1
            turn(0) = 1
            turn(1) = 2
            turn(2) = 3
            turn(3) = 4
        Case 2
            turn(0) = 2
            turn(1) = 3
            turn(2) = 4
            turn(3) = 1
        Case 3
            turn(0) = 3
            turn(1) = 4
            turn(2) = 1
            turn(3) = 2
        Case 4
            turn(0) = 4
            turn(1) = 1
            turn(2) = 2
            turn(3) = 3
    End Select
    
    'If Not bln_PbnEvent Then
    'thePBN = "E:AT5.AJT.A632.KJ7 Q763.KQ9.KQJ94.T 942.87653..98653 KJ8.42.T875.AQ42"
   
    thePBN = L_Dealer & ":" & _
             str_Hands(turn(0)) & " " & _
             str_Hands(turn(1)) & " " & _
             str_Hands(turn(2)) & " " & _
             str_Hands(turn(3))
    thePBN = Replace(thePBN, "-", ".", , 99)
    X = StrConv(thePBN, vbFromUnicode)
    
'If int_AvgDDANum > 0 Then str_Hands(1); "  "; str_Hands(3); "  "; str_Hands(2); "  "; str_Hands(4)
    
    For i = 0 To UBound(X)
        If i > 79 Then
            MsgBox "String too long"
            Exit Sub
        End If
        pbn.cards(i) = X(i)
    Next

    res = CalcDDTablePBN(pbn)

'If int_AvgDDANum > 0 Then Print #fiNum, "''''''''''''''''''" & int_AvgDDANum
    For player = 0 To 3
        ' For str_DDAOutput(), player is 0-3 but for '
        ' str_DDA and int_DDA, player is 1-4.  '
        For suit = 3 To 0 Step -1 ' Spade down to Club '
          'If player = 1 And suit = 3 Then Stop
            str_DDAOutput(player) = _
                str_DDAOutput(player) & _
                Right$(" " & Str(res.resTable(suit, player)), 2) & "  "
            'Debug.Print str_DDAOutput(player);
            If str_DDAOutput(player) = " 7   7   12   12" Then Stop
        Next
        str_DDAOutput(player) = _
            str_DDAOutput(player) & _
            Right$(" " & Str(res.resTable(4, player)), 2)
        'Debug.Print str_DDAOutput(player)
        ' Redo str_DDAOutput to make results equally spaced: '
        s1 = Replace(str_DDAOutput(player), "  ", ", ")

        s2() = Split(s1, ",")
        
'    Debug.Print thePBN: Debug.Print s1
        
        For s = 0 To 4
            int_OptByStrain(player + 1, s + 1, 0) = s2(s) ' 6 etc. '
        Next
        
        ' Keep track of DDA by player and suit for averaging. '
        ' int_DDA is loaded from PBN file when showing PBNs.  '
        If int_AvgDDANum > 0 Then ' Set to 1-nIt in AverageDDA.
            For s = 0 To 4
                j = Val(s2(s))
                int_DDA(player, s, j) = int_DDA(player, s, j) + 1
                'Debug.Print "p="; player; "s="; s; "j="; j
            Next
        End If
        
'If int_AvgDDANum > 0 Then Print #fiNum, ""
        
        str_DDAOutput(player) = Right$(" " & s2(0), 2)
        For i = 1 To 4
            str_DDAOutput(player) = _
                str_DDAOutput(player) & _
                Right$("  " & s2(i), 3)
        Next
        If player = int_Player(int_Turn) - 1 Then
            For s = 0 To 4 ' For C,D,H,S,N. '
                str_DDA(s) = "  " & Mid$("CDHSN", s + 1, 1) & "=" & LTrim$(s2(s))
            Next
        End If
    Next
    'End If ' DDA for PBN is loaded from PBN file - [not now]. '

'    Close #fiNum

'sEnd = DateTime.Now
'Debug.Print " "; int_AvgDDANum & "     " & DateDiff("s", sStart, sEnd); "     ";
    
    If DDA_Only Then
        ' mScramble and mAverageDDA only need the DDA, not the screen output. '
        Exit Sub
    End If
    
    '''''''''''''''''
    
    ' Compute Losing Trick Count: '
    Dim ltc(4) As Single
    Dim HCP(4, 4) As Integer ' HCP by player & suit. hcp(p, 0) = total hand. '
    Dim qty(4, 4)
    Dim hon(4, 4) As Integer ' # honors per suit (Ace-10) '
    Dim crd As String
    Dim fitNS As Integer
    Dim fitEW As Integer
    Dim lottNS As Integer
    Dim lottEW As Integer
    Dim ltcCap(2) As String
    Dim ltcSp(4) As String
    
    ' Add up HCP and # honors by player: '
    For player = 1 To 4
        suit = 4
        For i = 1 To 16
            crd = Mid$(str_Hands(player), i, 1)
            If crd = "-" Then
                HCP(player, 0) = HCP(player, 0) + HCP(player, suit)
                suit = suit - 1
            ElseIf crd > "9" Then
                HCP(player, suit) = HCP(player, suit) + InStr("JQKA", crd)
                hon(player, suit) = hon(player, suit) + 1
            End If
        Next
        HCP(player, 0) = HCP(player, 0) + HCP(player, suit) ' clubs '
    Next
    
    On Local Error GoTo SuitQtyErr
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
           int_SuitQty(2, suit) + HCP(2, suit) > 9 Or _
           int_SuitQty(4, suit) + HCP(4, suit) > 9 _
        Then
            fitEW = suit
        End If
    Next
    On Local Error GoTo 0
    
    For player = 1 To 4
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
                    
                    Else ' A? = 1 '
                        ltc(player) = ltc(player) + 1
                    End If
                
                ElseIf int_Suit(player, suit, King) > 0 Then  ' K? '
                    ltc(player) = ltc(player) + 1
                
                Else ' ?? = 2 '
                    ltc(player) = ltc(player) + 2
                End If
            
            ElseIf int_SuitQty(player, suit) > 2 Then  ' ??? '
                
                If int_Suit(player, suit, Ace) > 0 Then  ' A?? '
                    
                    If int_Suit(player, suit, King) > 0 Then ' AK? '
                        
                        If int_Suit(player, suit, 12) > 0 Then ' AKQ = 0 '
                            ' no losers. '
                        
                        ElseIf int_Suit(player, suit, 11) > 0 Then ' AKJ = .5 '
                            ltc(player) = ltc(player) + 0.5
                        
                        Else ' AK? = 1 loser '
                            ltc(player) = ltc(player) + 1
                        End If
                    
                    ElseIf int_Suit(player, suit, 12) > 0 Then ' AQ? '
                        
                        If int_Suit(player, suit, 11) > 0 Then 'AQJ = 1 '
                            ltc(player) = ltc(player) + 1
                        
                        Else ' AQ? = 1.5 '
                            ltc(player) = ltc(player) + 1.5
                        End If
                    Else ' Axx = 1 '
                        ltc(player) = ltc(player) + 2
                    End If
                
                ElseIf int_Suit(player, suit, King) > 0 Then  ' K?? '
                    
                    If int_Suit(player, suit, 12) > 0 Then ' KQ?  '
                        ltc(player) = ltc(player) + 1
                    
                    Else  ' K?? '
                        ltc(player) = ltc(player) + 2
                    End If
                
                ElseIf int_Suit(player, suit, 12) > 0 Then

                    If int_Suit(player, suit, 11) > 0 Or _
                       int_Suit(player, suit, 10) > 0 _
                    Then
                        ' QJ? or QT? = 2 '
                        ltc(player) = ltc(player) + 2
                    
                    Else ' Q?? '
                        ltc(player) = ltc(player) + 2.5
                    End If
                
                Else ' ??? = 3 '
                    ltc(player) = ltc(player) + 3
                End If
            End If
        Next
        If player = 3 Then
            ltcCap(1) = "-> 24 - " & _
                        ltc(3) & " - " & _
                        ltc(1) & " = " & _
                        (24 - ltc(3) - ltc(1))
                        
        ElseIf player = 4 Then
            ltcCap(2) = "-> 24 - " & _
                        ltc(4) & " - " & _
                        ltc(2) & " = " & _
                        (24 - ltc(4) - ltc(2))
        End If
        If ((player = 1 Or player = 3) And fitNS = 0) Or _
           ((player = 2 Or player = 4) And fitEW = 0) _
        Then
            If player = 3 Then
                ltcCap(1) = "-> No 8-card fit."
            ElseIf player = 4 Then
                ltcCap(2) = "-> No 8-card fit."
            End If
        End If
    Next
    
DisplayAnalysis:
    cmb_DDA.Locked = False
    
    For i = 0 To 4
        str_DDAOutput(i) = Trim$(str_DDAOutput(i))
        If Len(str_DDAOutput(i)) = 13 Then
            str_DDAOutput(i) = " " & str_DDAOutput(i)
        End If
    Next
    
    For i = 0 To 3
        If i = 0 Then
            p = 3
        ElseIf i = 1 Then
            p = 4
        ElseIf i = 2 Then
            p = 1
        Else
            p = 2
        End If
        str_OptByStrain(p, 0, 0) = str_DDAOutput(i)
    Next
    
    If cmb_DDA.ListCount > 0 Then
        cmb_DDA.Clear
        cmb_DDA.Text = "DDA = Maximum tricks that can be taken " & _
                       "by each player in each suit/NT."
    End If
    
    For i = 1 To 4
        If i = 1 Then
            ltcSp(i) = ltc(3)
        ElseIf i = 3 Then
            ltcSp(i) = ltc(1)
        ElseIf i = 2 Then
            ltcSp(i) = ltc(4)
        Else
            ltcSp(i) = ltc(2)
        End If
        
        If Len(ltcSp(i)) = 1 Then
            ltcSp(i) = " " & ltcSp(i) & " "
        ElseIf Len(ltcSp(i)) = 2 Then
            ltcSp(i) = ltcSp(i) & " "
        End If
    Next
          
    cmb_DDA.AddItem "DDA:C  D  H  S  N | HCP | LoTT | LTC: Calculate Tricks"
    cmb_DDA.AddItem "N: " & str_DDAOutput(0) & " | " & Right$("  " & HCP(3, 0), 3) & _
                    " |      | " & ltcSp(1) & "   Pairs Can Take:"
    cmb_DDA.AddItem "S: " & str_DDAOutput(2) & " |  " & Right$(" " & HCP(1, 0), 2) & _
                    " | " & Right$("  " & lottNS, 3) & _
                    "  | " & ltcSp(3) & ltcCap(1)
    cmb_DDA.AddItem "E: " & str_DDAOutput(1) & " |  " & Right$(" " & HCP(4, 0), 2) & _
                    " |      | " & ltcSp(2)
    cmb_DDA.AddItem "W: " & str_DDAOutput(3) & " |  " & Right$(" " & HCP(2, 0), 2) & _
                    " | " & Right$("  " & lottEW, 3) & "  | " & ltcSp(4) & ltcCap(2)
    cmb_DDA.Locked = True ' Don't allow changes to cmb_DDA by user. '
    
    Dim fi As Integer
    fi = FreeFile
    Close #fi
    Open "DDAresults.txt" For Output As #fi
    Print #fi, str_DDAOutput(2)
    Print #fi, str_DDAOutput(1)
    Print #fi, str_DDAOutput(0)
    Print #fi, str_DDAOutput(3)
    Close #fi
    
    CalcOpt (0)
    
    If mAverageDDA.Checked Or m100DealAvg.Checked Then
        AverageDDA 11
    End If
    
Exit Sub

SuitQtyErr:
    If Err.Number Then
        'Erase int_SuitQty()
        'ReDim int_SuitQty(4, 4)
        ' int_SuitQty() not filled in. '
        For i = 1 To 4
            For j = 1 To 4
                For k = 1 To 13
                    If int_Suit(i, j, k) > 0 Then
                        int_SuitQty(i, j) = int_SuitQty(i, j) + 1
                    End If
                Next
            Next
        Next
        Resume
    End If
End Sub


Sub AverageDDA(nIt As Integer)
    ' To get a more realistic DDA result,  '
    ' this Sub calculates the average DDA  '
    ' for each pair by mixing the cards for'
    ' the other pair and recalculating DDA.'
    ' nIt = number of iterations. '
    
    Dim h As Integer
    Dim s As Integer
    Dim p As Integer
    Dim n As Integer
    Dim n2 As Integer
    Dim i As Integer
    Dim j As Integer
    Dim trks As Integer
    Dim cards() As Integer
    Dim suitQty() As Integer  ' h, s '
    Dim suit() As Integer ' h, s, c '
    Dim handStr() As String ' h '
    Dim handInt As Variant ' h, c '
    Dim d As String ' deal name '
    
    Dim ddaTotal(4, 5, 13)
    ' Total DDA for player 1-4, strains 1-5 (C,D,H,S,N), # tricks 1-13. '

    ' ddaTotal is for averaging.      '
    ' Example: ddaTotal(1, 4, 10) = 6 '
    '          ddaTotal(1, 5,  9) = 5 '
    '          ddaTotal(1, 5, 10) = 2 '
    ' means Spades(4) took 10 tricks  '
    ' on 6 out of n deals while NT(5) '
    ' took 9+ tricks on 7 deals.      '
    ' ddaTotal(1, 3, 10) = 4 and      '
    ' ddaTotal(1, 3, 11) = 3 means    '
    ' that Hearts(3) took 4+ tricks 7 '
    ' times, so the optimum contract  '
    ' is Hearts over 10 deals. Note   '
    ' that on any 1 of those deals,   '
    ' either H, S or NT could be opt. '
    
    Erase int_DDA()
    
    Dim sStart As Date
    Dim sEnd As Date
    Dim nSecs As Long
    sStart = DateTime.Now
    
    If bln_Averaging Then
        Exit Sub
    End If
    bln_Averaging = True
    t_Calculating.Visible = True
    t_Calculating.ZOrder 0
        
    ' Save the original deal's data: '
    suitQty = int_SuitQty
    suit = int_Suit
    cards = int_Cards
    handStr = str_Hands
    handInt = int_Hand
    
    If cmb_DDA.ListCount < 11 Then
        Me.Height = 10510
        For n = 1 To 6: cmb_DDA.AddItem "": Next
    End If
    
    If cmb_DDA.Height < 130 Then
        bln_NoPaint = True
        cmb_DDA.Height = 180
        cmb_FindDeals.Visible = False
    End If
    If nIt = 11 Then
        cmb_DDA.List(6) = "Calculating averages for 11 deals."
    End If
    If nIt = 101 Then
        t_Calculating.Visible = True
        t_Calculating.ZOrder 0
        cmb_DDA.List(6) = "Calculating averages for 101 deals."
    End If
    cmb_DDA.Text = "DDA = Maximum tricks that can be taken by each player in each suit/NT."
    DoEvents
    
'For j = 1 To 4: Debug.Print str_Hands(i); "  ";: Next: Debug.Print ""
    ' First keep N-S and scramble E-W,   '
    ' then restore E-W and scramble N-S. '
    For p = 1 To 2
        int_NumPickedHands = 2 ' # of hands to scramble; used elsewhere. '
        
        If p = 1 Then ' Scramble hands 2 and 4: '
            bln_PickedHand(2) = True
            bln_PickedHand(4) = True
            bln_PickedHand(1) = False
            bln_PickedHand(3) = False
        Else ' Scramble hands 1 and 3: '
            bln_PickedHand(1) = True
            bln_PickedHand(3) = True
            bln_PickedHand(2) = False
            bln_PickedHand(4) = False
        End If

        For i = 1 To nIt
            ScrambleHands
            SortHands ' Sort hands within int_Cards().   '
            FillSuits ' Assign int_Cards() to int_Suit().'
            DoEvents
                
            If bln_PickedHand(2) Then
                SetUpHands 2
                SetUpHands 4
            End If
            int_AvgDDANum = i
            ShowDDA True ' True = just do DDA, not the other specs. '
'sEnd = DateTime.Now
'nSecs = DateDiff("s", sStart, sEnd)
'Debug.Print str_Hands(1); " "; str_Hands(2); " "; str_Hands(3); " "; str_Hands(4); " secs: "; nSecs
        Next
'Debug.Print ""

        ' Need to compute the optimums here for  '
        ' the non-picked hands since they are the'
        ' ones not mixed. '
        Dim div As Single
        If nIt = 11 Then
            div = 5.5 ' ?? This is supposed to be half nIt '
        Else
            div = 55
        End If
        For h = 1 To 4
            If Not bln_PickedHand(h) Then
                For s = 1 To 5
                    i = 0
                    For n = 13 To 1 Step -1
                      'If h = 4 And s > 2 Then Stop
                        i = i + int_DDA(h - 1, s - 1, n)
                        If nIt = 11 Then j = 1 Else j = 2
                        If i > div Then
                            ' n starts with 13, the highest   '
                            ' possible level of bidding.      '
                            ' i is the running total of times '
                            ' level n or higher was bid. When '
                            ' i is more than half the number  '
                            ' of iterations (nIt), then the   '
                            ' level (n) is the median and the '
                            ' optimum level for suit s. '
                            int_OptByStrain(h, s, j) = n
                            'If h = 4 Then
                            '    Debug.Print "h="; h; "  s="; s; "  j="; j; "  n="; n
                            'End If
                            'ddaTotal(h, s, 0) = n
                            Exit For
                        Else
                            int_OptByStrain(h, s, j) = 0
                            'ddaTotal(h, s, 0) = 0
                        End If
                    Next
                Next
            End If
        Next
        
        ' We need to restore original deal. '
        ' If viewing PBNs, we can reload it;'
        ' otherwise, we can redo the deal   '
        ' using the deal number in L_Deal.  '
        int_Cards = cards
        int_Suit = suit
        int_SuitQty = suitQty
        str_Hands = handStr
        int_Hand = handInt
'sEnd = DateTime.Now
'nSecs = DateDiff("s", sStart, sEnd)
'Debug.Print str_Hands(1); " "; str_Hands(2); " "; str_Hands(3); " "; str_Hands(4); " secs: "; nSecs
    Next

    With cmb_DDA
        .Locked = False
        If .ListCount < 10 Then
            For n = .ListCount To 10
                .AddItem ""
            Next
        End If
        
        sEnd = DateTime.Now
        nSecs = DateDiff("s", sStart, sEnd)
        .List(6) = "Median DDA: " & div ' & nSecs & " secs."
        
        If Left$(.List(7) & "  ", 2) <> "N:" Or nIt = 11 Then
            .List(7) = "N:"
            .List(8) = "S:"
            .List(9) = "E:"
            .List(10) = "W:"
        Else
            .List(7) = .List(7) & "  "
            .List(8) = .List(8) & "  "
            .List(9) = .List(9) & "  "
            .List(10) = .List(10) & "  "
        End If
        
        If nIt = 11 Then
            i = 1
        ElseIf nIt > 99 Then
            i = 2
            sEnd = DateTime.Now
            nSecs = DateDiff("s", sStart, sEnd)
            .List(6) = "Median DDA (11)      C  D  H  S  N  (101) " & div  ' & nSecs & " secs."
        Else
            i = 0
        End If
        
        .List(7) = .List(7) & Right$("  " & int_OptByStrain(3, 1, i), 3) & _
                              Right$("  " & int_OptByStrain(3, 2, i), 3) & _
                              Right$("  " & int_OptByStrain(3, 3, i), 3) & _
                              Right$("  " & int_OptByStrain(3, 4, i), 3) & _
                              Right$("  " & int_OptByStrain(3, 5, i), 3)
        .List(8) = .List(8) & Right$("  " & int_OptByStrain(1, 1, i), 3) & _
                              Right$("  " & int_OptByStrain(1, 2, i), 3) & _
                              Right$("  " & int_OptByStrain(1, 3, i), 3) & _
                              Right$("  " & int_OptByStrain(1, 4, i), 3) & _
                              Right$("  " & int_OptByStrain(1, 5, i), 3)
        .List(9) = .List(9) & Right$("  " & int_OptByStrain(4, 1, i), 3) & _
                              Right$("  " & int_OptByStrain(4, 2, i), 3) & _
                              Right$("  " & int_OptByStrain(4, 3, i), 3) & _
                              Right$("  " & int_OptByStrain(4, 4, i), 3) & _
                              Right$("  " & int_OptByStrain(4, 5, i), 3)
        .List(10) = .List(10) & Right$("  " & int_OptByStrain(2, 1, i), 3) & _
                              Right$("  " & int_OptByStrain(2, 2, i), 3) & _
                              Right$("  " & int_OptByStrain(2, 3, i), 3) & _
                              Right$("  " & int_OptByStrain(2, 4, i), 3) & _
                              Right$("  " & int_OptByStrain(2, 5, i), 3)
      
        Dim fi As Integer
    fi = FreeFile
    Close #fi
    Open "DDAresults.txt" For Append As #fi
    If nIt = 11 Then
        Print #fi, Mid$(.List(8), 4)  ' S '
        Print #fi, Mid$(.List(10), 4) ' W '
        Print #fi, Mid$(.List(7), 4)  ' N '
        Print #fi, Mid$(.List(9), 4)  ' E '
    Else
        Print #fi, Mid$(.List(8), 21)
        Print #fi, Mid$(.List(10), 21)
        Print #fi, Mid$(.List(7), 21)
        Print #fi, Mid$(.List(9), 21)
    End If
    Close #fi
      
      .Text = "DDA = Maximum tricks that can be taken by each player in each suit/NT."
      .Locked = True ' Don't allow changes to cmb_DDA by user. '
      int_AvgDDANum = 0
    End With
    
    ' Calculate the Optimum contract for each player: '
    ' If trks < 7 then no bid can be made; points = 0.'
    ' If C/D, <11 then points = 50 + (trks-6) * 20.   '
    '         11 = 400 (N.V.).  12 = 920.  12 = 1440. '
    ' If H/S, <10 = 50 + (trks-6) * 30. (2H = 110)    '
    '         10 = 420. 11 = 450. 12 = 980. 13 = 1510.'
    ' If NT, <9 = 50 + (trks-6) * 30 + 10. (2N = 120) '
    '  9=400. 10 = 430. 11 = 460. 12 = 990. 13 = 1520 '
    ' Vul doesn't matter, just # of tricks taken.     '
    ' Tricks are converted to points for comparison.  '
    ' E.g.: 5C/5D = 500.  3N = 500, 5N = 560.         '
    '                                                 '
    ' If DDA shows C and D take the same # of tricks, '
    ' or H and S take the same, then Optimum is the   '
    ' suit which has the most trumps. If they have the'
    ' same number of trumps, then the higher suit is  '
    ' the optimum (D over C, S over H). E.g.: hearts  '
    ' has a 5-3 fit and spades has 4-3 and both take  '
    ' 10 tricks per DDA, the hearts is the Opt. suit  '
    ' because it normally would be the easiest/safest.'
    ' If both have 5-3 fit, then spades is the Opt.   '
    
    ' Edit: This can create a problem. If 2 suits have'
    ' the same score, then it should be noted that we '
    ' have 2 possible Opts when computing bids via DDA.
        Dim k As Integer
        For h = 1 To 4
            j = 0 ' Keep track of high score for suit for player h. '
            n = 0
            For s = 1 To 5
                trks = int_OptByStrain(h, s, i) ' ddaTotal(h, s, 0)
                                           ' i is 11, 99, or 0 where 0 is no avg. '
                If trks < 7 Then
                    n = 0
                    
                ElseIf trks = 13 Then
                    If s < 3 Then ' C/D '
                        n = 1440
                    ElseIf s < 5 Then ' H/S '
                        n = 1510
                    Else ' NT '
                        n = 1520
                    End If
                        
                ElseIf trks = 12 Then
                    If s < 3 Then ' C/D '
                        n = 920
                    ElseIf s < 5 Then ' H/S '
                        n = 980
                    Else ' NT '
                        n = 990
                    End If
                
                ElseIf trks = 11 Then
                    If s < 3 Then ' C/D '
                        n = 400
                    ElseIf s < 5 Then ' H/S '
                        n = 450
                    Else ' NT '
                        n = 460
                    End If
                
                ElseIf trks = 10 Then
                    If s < 3 Then ' C/D '
                        n = 130
                    ElseIf s < 5 Then ' H/S '
                        n = 420
                    Else ' NT '
                        n = 430
                    End If
                    
                ElseIf trks = 9 And s = 5 Then ' NT '
                    n = 400
                    
                Else ' <9 tricks '
                    If s < 3 Then ' C/D '
                        n = 50 + (trks - 6) * 20
                    ElseIf s < 5 Then ' H/S '
                        n = 50 + (trks - 6) * 30
                    Else ' NT '
                        n = 50 + (trks - 6) * 30 + 10
                    End If
                End If

                If n >= j And n + j > 0 And h > 2 Then ' j = previous high '
                    ' Current suit is >= previous high. '
                    If n = j And (s = 2 Or s = 4) Then
                        ' If D=C or S=H then the suit with the most trumps is Opt. '
                        ' If both have the same # of trumps, the higher suit = Opt.'
                        If suitQty(h, s) + suitQty(h - 2, s) >= suitQty(h, s - 1) + suitQty(h - 2, s - 1) Then
                            j = n
                            k = s ' j = Optimum suit number. '
                        End If
                    Else
                        j = n
                        k = s
                    End If
                ElseIf n >= j Then
                    j = n
                    k = s ' k = Optimum suit number. '
                End If
'If s = 2 Then Stop
            Next
            n = Fix(int_OptByStrain(h, k, i)) - 6 'Fix(ddaTotal(h, k, 0)) - 6
            If n < 1 Then
                int_OptAvg(h) = 0
            Else
                int_OptAvg(h) = n * 10 + k
                ' Optimum. Examples: 32=3D, 34=3S, 25=2N '
            End If
        Next
        
    bln_Averaging = False
    t_Calculating.Visible = False
    For i = 1 To 4
        DisplayHand i
    Next

sEnd = DateTime.Now
nSecs = DateDiff("s", sStart, sEnd)
'Debug.Print " secs: "; nSecs

End Sub


Private Sub CalcOpt(avg As Integer)
    ' avg is # iterations: 0=original, 1=11, 2=101    '
    
    ' Calculate the Optimum contract for each player: '
    ' If trks < 7 then no bid can be made; points = 0.'
    ' If C/D, <11 then points = 50 + (trks-6) * 20.   '
    '         11 = 400 (N.V.).  12 = 920.  12 = 1440. '
    ' If H/S, <10 = 50 + (trks-6) * 30. (2H = 110)    '
    '         10 = 420. 11 = 450. 12 = 980. 13 = 1510.'
    ' If NT, <9 = 50 + (trks-6) * 30 + 10. (2N = 120) '
    '  9=400. 10 = 430. 11 = 460. 12 = 990. 13 = 1520 '
    ' Vul doesn't matter, just # of tricks taken.     '
    ' Tricks are converted to points for comparison.  '
    ' E.g.: 5C/5D = 500.  3N = 500, 5N = 560.         '
    '                                                 '
    ' If DDA shows C and D take the same # of tricks, '
    ' or H and S take the same, then Optimum is the   '
    ' suit which has the most trumps. If they have the'
    ' same number of trumps, then the higher suit is  '
    ' the optimum (D over C, S over H). E.g.: hearts  '
    ' has a 5-3 fit and spades has 4-3 and both take  '
    ' 10 tricks per DDA, the hearts is the Opt. suit  '
    ' because it normally would be the easiest/safest.'
    ' If both have 5-3 fit, then spades is the Opt.   '
    
    ' Edit: This can create a problem. If 2 suits have'
    ' the same score, then it should be noted that we '
    ' have 2 possible Opts when computing bids via DDA.
        Dim k As Integer
        Dim h As Integer
        Dim j As Integer
        Dim n As Integer
        Dim s As Integer
        Dim trks As Integer
        Dim suitQty() As Integer
    
    suitQty = int_SuitQty
        
        For h = 1 To 4
            j = 0 ' Keep track of high score for suit for player h. '
            n = 0
            For s = 1 To 5
                trks = int_OptByStrain(h, s, avg)
                        ' Where h = player #: S=1 E=2 N=3 W=4     '
                
                If trks < 7 Then
                    n = 0
                    
                ElseIf trks = 13 Then
                    If s < 3 Then ' C/D '
                        n = 1440
                    ElseIf s < 5 Then ' H/S '
                        n = 1510
                    Else ' NT '
                        n = 1520
                    End If
                        
                ElseIf trks = 12 Then
                    If s < 3 Then ' C/D '
                        n = 920
                    ElseIf s < 5 Then ' H/S '
                        n = 980
                    Else ' NT '
                        n = 990
                    End If
                
                ElseIf trks = 11 Then
                    If s < 3 Then ' C/D '
                        n = 400
                    ElseIf s < 5 Then ' H/S '
                        n = 450
                    Else ' NT '
                        n = 460
                    End If
                
                ElseIf trks = 10 Then
                    If s < 3 Then ' C/D '
                        n = 130
                    ElseIf s < 5 Then ' H/S '
                        n = 420
                    Else ' NT '
                        n = 430
                    End If
                    
                ElseIf trks = 9 And s = 5 Then ' NT '
                    n = 400
                    
                Else ' <9 tricks '
                    If s < 3 Then ' C/D '
                        n = 50 + (trks - 6) * 20
                    ElseIf s < 5 Then ' H/S '
                        n = 50 + (trks - 6) * 30
                    Else ' NT '
                        n = 50 + (trks - 6) * 30 + 10
                    End If
                End If

                If n >= j And n + j > 0 And h > 2 Then ' j = previous high '
                    ' Current suit is >= previous high. '
                    If n = j And (s = 2 Or s = 4) Then
                        ' If D=C or S=H then the suit with the most trumps is Opt. '
                        ' If both have the same # of trumps, the higher suit = Opt.'
                        If suitQty(h, s) + suitQty(h - 2, s) >= suitQty(h, s - 1) + suitQty(h - 2, s - 1) Then
                            j = n
                            k = s ' j = Optimum suit number. '
                        End If
                    Else
                        j = n
                        k = s
                    End If
                ElseIf n >= j Then
                    j = n
                    k = s ' k = Optimum suit number. '
                End If
            Next
            
            n = Fix(int_OptByStrain(h, k, avg)) - 6
            
            If avg = 0 Then
                If n < 1 Then
                    int_OptOrig(h) = 0
                Else
                    int_OptOrig(h) = n * 10 + k
                    ' Optimum. Examples: 32=3D, 34=3S, 25=2N '
                End If
            Else
                If n < 1 Then
                    int_OptAvg(h) = 0
                Else
                    int_OptAvg(h) = n * 10 + k
                    ' Optimum. Examples: 32=3D, 34=3S, 25=2N '
                End If
            End If
        Next

End Sub


Private Sub ExportPracticeDealList()
    Dim fiNum As Integer
    Dim fiNum2 As Integer
    Dim found As FoundDeal
    Dim lastDeal As lastDeal
    Dim j As Integer
    Dim i As Integer
    Dim txt As String

    txt = Replace(str_PracticeDealsFile, ".Deals", ".Txt")
    fiNum2 = FreeFile
    Open txt For Append As fiNum2
    
    fiNum = FreeFile
    Open str_PracticeDealsFile For Random As fiNum Len = 6
        'Get #fiNum, 1, found
        'j = found.DealNum ' Number of entries in the file. '
        'For i = 1 To j
        i = lng_NumPracticeEntries
            Get #fiNum, i, found
            Write #fiNum2, i, found.dealNum, found.Dealer, str_Hands(int_Dealer)
            'Debug.Print found.DealNum, found.Dealer
        'Next
    Close fiNum2
    Close fiNum
End Sub


Private Sub MakePracticeWebPage()
    Dim fiNum As Integer
    Dim fiNum3 As Integer
    Dim j As Integer
    Dim i As Integer
    Dim k As Integer
    Dim txt As String
    Dim hand As String
    Dim id1 As Long
    Dim crlf As String
    Dim qt As String
    Dim s As String
    Dim s1 As String
    Dim t As String
    Dim calls As String
    Dim getAnotherBid As Boolean
    Dim cn As Integer ' Call Number '
    Dim myTurn As Integer ' int_CallNum '
    Dim pb As String ' prior bids '
    
    qt = Chr$(34)
    crlf = Chr$(13) & Chr$(10)
    int_QuizQuestionNum = int_QuizQuestionNum + 1
    
    'Stop ' make sure this works: '
    pb = "(" & str_AllCalls & ")" ' e.g.: (1C - D - P) '
    i = InStr(pb, "-")
    pb = Left$(pb, i - 1) & ")" & Mid$(pb, i)
    i = InStrRev(pb, "-")
    pb = Left$(pb, i) & "(" & Mid$(pb, i + 1)
    
    id1 = int_QuizQuestionNum
    hand = str_Hands(int_Player(int_Turn))
    
    txt = str_BBpath & "\Notes\Practice_" & str_PracticeSpecsFile & ".txt"
    txt = Replace(txt, " ", "_")
    
    s = "<BR>" & id1 & ". &nbsp;Deal #" & lng_DealNum & crlf
    s = s & "<UL>Your hand: <B>" & hand & "</B>" & crlf
    s = s & "<BR>Prior bids: &nbsp;<B>" & pb & "</B>" & crlf
    s = s & "<SPAN ID=" & qt & "prompt" & id1 & qt & " STYLE=''" & qt & " COLOR=BLACK>" & crlf
    s = s & "    <SPAN OnClick=" & qt & "prompt" & id1 & ".style.display='none'" & qt & ">" & crlf
    s = s & "        <SPAN OnClick=" & qt & "hidden" & id1 & ".style.display=''" & qt & ">" & crlf
    s = s & "            <SPAN id=" & qt & "GreenHand" & qt & "><BR>" & crlf
    s = s & "                 <U><B>Your bid</B></U>" & crlf
    s = s & "           </SPAN>" & crlf
    s = s & "       </SPAN>" & crlf
    s = s & "   </SPAN>" & crlf
    s = s & "</SPAN>" & crlf
    s = s & "<SPAN ID=" & qt & "hidden" & id1 & qt & " STYLE=" & qt & "display:none" & qt & " COLOR=BLACK>" & crlf
    
    fiNum3 = FreeFile
    Open txt For Append As fiNum3 ' Paste this into .htm file. '
    
    ShowDDA True ' Calculates DDA but not other stats. '
    ' str_DDA(h) is clubs thru NT. '
    ' All dda's are added to clubs.   '
    ' Ex: "C=9, D=10, H=8, S=10, N=9" '
    MsgBox "Stop MakePracticeWebPage 1" ' make sure this is right: '
    For j = 1 To 4
        str_DDA(0) = str_DDA(0) & " " & str_DDA(j)
    Next
    str_DDA(0) = Replace(str_DDA(0), "  ", " &nbsp;")
    myTurn = int_Turn
        
    If Not bln_Get2ndBid Then
        s = s & "<BR>BidBase: &nbsp;<SPAN style=""color:#FF0000""><U><B>" & _
                str_Call(int_CallNum) & "</B></U><SPAN style=""color:#000000""> &nbsp;(DDA: " & str_DDA(0) & ")</SPAN>" & _
                "</SPAN>" & _
                crlf
        s = s & "</SPAN>" & crlf & "</UL>" & crlf & crlf
        Print #fiNum3, s
    Else
        ' This next code gets partner's next bid & this player's next bid. '
        ' All MakeBid needs is player number and str_AllCalls to be right. '
    'MsgBox "Stop MakePracticeWebPage 1" ' For MOSTD, want to get everyone's bid for the next round. ' 1C-D-A-Me, then LHO-CHO-RHO-Me2 '
        calls = str_AllCalls ' Keep track of str_AllCalls before getting the extra bids. '
        cn = int_CallNum
        
        str_AllCalls = str_AllCalls & "-" & str_Call(int_CallNum)
                        
        ' Get next 4 bids: '
        For j = 1 To 4
            bln_Abort = True ' To get an immediate return from MakeBid
            int_Turn = int_Turn + 1
            If int_Turn > 4 Then int_Turn = int_Turn - 4
            If int_Turn = 0 Then int_Turn = 1
            int_CallNum = int_CallNum + 1
            MakeBid int_Player(int_Turn) ' The bid is returned in str_Call(int_CallNum). '
            'If j = 1 Or j = 3 Then - don't know why not doing j=2 or 4. '
                str_AllCalls = str_AllCalls & "-" & str_Call(int_CallNum)
            'End If
        Next
        
        ' Get partner's hand: (Should use my_Turn?) '
        If int_Turn - 2 > 0 Then
            i = int_Turn - 2
        Else
            i = int_Turn + 2
        End If
        hand = str_Hands(int_Player(i))
        
        ' cn is my call number, so cn+1=LHO, cn+2=CHO, cn+3=RHO '
        ' pb is the 2nd round of bidding excluding my 2nd bid.  '
        pb = "(" & str_Call(cn + 1) & ")-" & str_Call(cn + 2) & "-(" & str_Call(cn + 3) & ")"
        
        t = t & s & crlf & "    <BR>Your 1st bid: &nbsp;<B>" & str_Call(cn) & "</B>" & _
                   crlf & "    <BR>Next 3 bids: &nbsp;<B>" & pb & "</B>" & crlf & crlf
                                                             ''
        t = t & "<SPAN ID=" & qt & "prompt" & id1 & "a" & qt & " STYLE=''" & qt & " COLOR=BLACK>" & crlf
      
      If pb <> "(P)-P-(P)" Then
        t = t & "    <SPAN OnClick=" & qt & "prompt" & id1 & "a.style.display='none'" & qt & ">" & crlf
        t = t & "        <SPAN OnClick=" & qt & "hidden" & id1 & "a.style.display=''" & qt & ">" & crlf
        t = t & "            <SPAN id=" & qt & "GreenHand" & qt & "><BR>" & crlf
        t = t & "                <U><B>Your 2nd bid</B></U>" & crlf
        t = t & "            </SPAN>" & crlf
        t = t & "         </SPAN>" & crlf
        t = t & "     </SPAN>" & crlf
        t = t & "</SPAN>" & crlf
        t = t & "<SPAN ID=" & qt & "hidden" & id1 & "a" & qt & " STYLE=" & qt & "display:none" & qt & " COLOR=BLACK>" & crlf
        t = t & "    <BR>Your 2nd bid: &nbsp;<B>" & str_Call(int_CallNum) & "</B>" & crlf
                                                    '''''''''''''''''''''
      End If
        
        t = t & "    <BR>Pard's hand: &nbsp;<B>" & hand & "</B>" & crlf
                                                   ''''
        t = t & "    <BR>DDA: &nbsp;<B>" & str_DDA(0) & "</B>" & crlf
                                           ''''''''''
        ' The next 2 lines result in an extra </SPAN> when pb<>(P)-P-(P), '
        ' but it doesn't seem to hurt. '
        t = t & "</SPAN>" & crlf & "<BR>"
        t = t & "</SPAN></UL>" & crlf & crlf
        Print #fiNum3, t
    End If
    
        str_AllCalls = calls ' Restore the original str_AllCalls. Why? We go to next deal now. '
        int_CallNum = cn     ' All these variables should start over with next deal. '
        int_Turn = myTurn
        
    Close fiNum3
End Sub


Function MyMsgBox(txt As String, numLines As Integer, _
                  response As String, boxCaption As String) As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim n As Integer
    Dim X As String
    Dim useOk As Integer
    Dim lenTxt As Long
    Dim msgWidth As Integer
    Dim tl As Long
    Dim longestLine As String
    
    If f_Msg.Visible Then Exit Function
    f_Msg.Width = 500 ' Shortened below. '
    
    If Not bln_SkipIt Then
        t_Calculating.Visible = True
        t_Calculating.ZOrder 0
    End If
    f_Msg.ZOrder 0
    
    If boxCaption = "" Then boxCaption = "CardShark BidBase"
    L_Caption = "  " & boxCaption
    If numLines = 0 Then numLines = 1

    L_MsgBox = txt
    ' Find longest line: '
    j = InStr(txt, vbCr)
    If numLines < 2 Or j = 0 Then
        msgWidth = Len(txt)
        longestLine = txt
    Else
        longestLine = Left$(txt, j - 1)
        msgWidth = Len(longestLine)
        k = j
        For n = 2 To numLines
            j = InStr(j + 1, txt, vbCr)
            If j = 0 Then
                X = Mid$(txt, k + 1)
            Else
                X = Mid$(txt, k + 1, j - k)
                k = j
            End If
            'Debug.Print n; j; "  "; x
            If Len(X) > msgWidth Then
                msgWidth = Len(X)
                longestLine = X
            End If
        Next
    End If
    
    msgWidth = TextPixelWidth(longestLine)
    
    L_MsgBox.Width = msgWidth * 15.8
        
    t_MsgInput.Width = L_MsgBox.Width - t_MsgInput.Left + 200
    f_Msg.Width = (L_MsgBox.Width + 400) / 15.3 ' Margins of 200 on each side. '
    f_Msg.Left = (Me.Width / 15.3 - f_Msg.Width) / 2
    L_Caption.Width = 5000
f_Msg.Visible = True
    L_MsgBox.Height = 300 * numLines  ' Set L_MsgBox.BackStyle=1 to see it. '
    f_Msg.Height = 20 * numLines + Me.ScaleY(b_MsgInput(0).Height, vbTwips, vbPixels) + 45
    b_MsgInput(0).Top = L_MsgBox.Top + L_MsgBox.Height + 50
    L_MsgBackground.Height = f_Msg.Height * 14.67 - 385
    L_MsgBackground.Width = L_MsgBox.Width + 200
    t_MsgInput.Visible = False

    response = UCase(response)

    b_MsgInput(1).Visible = False  ' No '
    b_MsgInput(2).Visible = False  ' Cancel '

    b_MsgInput(0).Default = True

    If response = "OK" Then ' Wait for Ok press. '
        b_MsgInput(0).Caption = "OK"

    ElseIf Left$(response & "   ", 2) = "YN" Then
        ' Show message, wait for Yes, No. '
        b_MsgInput(0).Caption = "Yes"
        b_MsgInput(1).Top = b_MsgInput(0).Top
        b_MsgInput(1).Visible = True
        If Right$(response, 1) = "C" Then
            b_MsgInput(2).Top = b_MsgInput(0).Top
            b_MsgInput(2).Visible = True
        End If

    Else ' Input box uses response as default input. '
        t_MsgInput.Visible = True
        b_MsgInput(0).Caption = "Ok"
        t_MsgInput.Top = b_MsgInput(0).Top
        t_MsgInput.Text = response
        f_Msg.Visible = True ' In order to set focus '
        t_MsgInput.SetFocus
    End If
    
    f_Msg.Visible = True
    f_Msg.ZOrder 0
    
    Do While f_Msg.Visible
        X = DoEvents()
    Loop
    
    If response = "OK" Then
        ' No value is returned. '
    ElseIf Left$(response, 1) = "Y" Then
        ' 0=Yes, 1=No, 2=Cancel '
        If int_MsgClick = 2 Then
            MyMsgBox = "cancel"
        ElseIf int_MsgClick = 1 Then
            MyMsgBox = "no"
        Else
            MyMsgBox = "yes"
        End If
    Else ' Input box '
        MyMsgBox = t_MsgInput.Text
    End If
    b_MsgInput(0).Default = False
    t_Calculating.Visible = False
End Function


Function TextPixelWidth(txt As String) As Long
    Dim c As String ' character being measured '
    Dim i As Integer
    Dim t As Double ' width of a character in twips '
    Dim tt As Double ' total pixels width '

    ' This sub calculates the screen width of text for MyMsgBox. '
    ' Seems excessive, but what the heck. Neatness counts. '
    
    txt = Trim$(txt)
    tt = 0
    t = 0
    For i = 1 To Len(txt)
        c = Mid$(txt, i, 1)
        Select Case c
        Case "i", "j", "l", "'"
            t = t + 4
        Case "/", "\", " ", ".", ",", ";", ":", "[", "]"
            t = t + 5
        Case "f", "r", "t", "-", "{", "}", "(", ")", "I", "|"
            t = t + 6
        Case Chr$(34)
            t = t + 7.33
        Case "v", "x", "y"
            t = t + 8
        Case "c", "k", "s", "z", "J"
            t = t + 9
        Case "a", "b", "d", "e", "g", "h", "n", "o", "p", "q", "u", "L", "<", ">", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "="
            t = t + 10
        Case "F", "K", "P", "Z"
            t = t + 11
        Case "A", "B", "C", "E", "S", "T", "V", "X", "Y", "w", "#"
            t = t + 12
        Case "D", "G", "H", "N", "O", "Q", "R", "U"
            t = t + 13.333
        Case "M", "m"
            t = t + 14.133
        Case "W"
            t = t + 16.333
        Case Else
            t = t + 1
        End Select
'        Debug.Print c; t - tt; t
        tt = t
    Next
    TextPixelWidth = t
End Function

