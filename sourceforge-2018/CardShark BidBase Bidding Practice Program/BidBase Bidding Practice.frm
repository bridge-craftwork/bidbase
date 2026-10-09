VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form PlayScr 
   Appearance      =   0  'Flat
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CardShark BidBase Practice Dealer"
   ClientHeight    =   7125
   ClientLeft      =   1320
   ClientTop       =   2070
   ClientWidth     =   9825
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
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   475
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   655
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4170
      Picture         =   "BidBase Bidding Practice.frx":0000
      ScaleHeight     =   300
      ScaleWidth      =   990
      TabIndex        =   94
      Top             =   3270
      Width           =   990
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2085
      Left            =   2655
      TabIndex        =   1
      Top             =   1560
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   3678
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
   Begin VB.TextBox t_EditData 
      Height          =   840
      Left            =   1230
      TabIndex        =   93
      Top             =   120
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Frame f_CardDeck 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3570
      Left            =   1005
      TabIndex        =   20
      Top             =   1365
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
         Top             =   2700
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
         Top             =   2700
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
         Top             =   2700
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
         Top             =   2700
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
         Top             =   2700
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
         Top             =   2700
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
         Top             =   2700
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
         Top             =   2700
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
         Top             =   2700
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
         Top             =   2700
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
         Top             =   2700
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
         Top             =   2700
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
         Top             =   2700
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
         Top             =   1800
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
         Top             =   1800
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
         Top             =   1800
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
         Top             =   1800
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
         Top             =   1800
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
         Top             =   1800
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
         Top             =   1800
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
         Top             =   1800
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
         Top             =   1800
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
         Top             =   1800
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
         Top             =   1800
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
         Top             =   1800
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
         Top             =   1800
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
         Top             =   900
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
         Top             =   900
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
         Top             =   900
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
         Top             =   900
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
         Top             =   900
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
         Top             =   900
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
         Top             =   900
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
         Top             =   900
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
         Top             =   900
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
         Top             =   900
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
         Top             =   900
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
         Top             =   900
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
         Top             =   900
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
         Top             =   0
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
         Top             =   0
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
         Top             =   0
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
         Top             =   0
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
         Top             =   0
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
         Top             =   0
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
         Top             =   0
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
         Top             =   0
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
         Top             =   0
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
         Top             =   0
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
         Top             =   0
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
         Top             =   0
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
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.Frame f_Inbox 
      Height          =   1920
      Left            =   3690
      TabIndex        =   88
      Top             =   2400
      Visible         =   0   'False
      Width           =   2415
      Begin VB.CommandButton b_DoneDealNum 
         Caption         =   "Done"
         Height          =   360
         Left            =   795
         TabIndex        =   92
         Top             =   1450
         Width           =   790
      End
      Begin VB.TextBox t_Inbox 
         Height          =   330
         Left            =   595
         TabIndex        =   91
         Top             =   950
         Width           =   1190
      End
      Begin VB.Label L_InboxLbl 
         Alignment       =   2  'Center
         Caption         =   "Enter a deal number or ""user"" to enter hands:"
         Height          =   380
         Left            =   10
         TabIndex        =   90
         Top             =   450
         Width           =   2340
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
         TabIndex        =   89
         Top             =   0
         Width           =   2400
      End
   End
   Begin VB.TextBox t_Disclose 
      ForeColor       =   &H00004000&
      Height          =   1275
      Left            =   2625
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   3630
      Width           =   4875
   End
   Begin VB.CheckBox ck_KeepDealer 
      BackColor       =   &H0000C000&
      Caption         =   "Keep"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1140
      TabIndex        =   86
      Top             =   6780
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Frame f_FindBidDeals 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   1200
      TabIndex        =   84
      Top             =   5625
      Width           =   1305
      Begin VB.CheckBox ck_FindBids 
         BackColor       =   &H0000C000&
         Caption         =   "Start with"
         Height          =   200
         Left            =   0
         TabIndex        =   87
         Top             =   -15
         Width           =   1170
      End
      Begin VB.TextBox t_FindBids 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   85
         Text            =   "BidBase Bidding Practice.frx":3514
         Top             =   200
         Width           =   1300
      End
   End
   Begin VB.Frame f_NSEW 
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   900
      Left            =   7650
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
      Begin VB.CommandButton b_NSEW 
         Caption         =   "N"
         Height          =   210
         Index           =   3
         Left            =   200
         TabIndex        =   81
         Top             =   8
         Width           =   400
      End
      Begin VB.CommandButton b_NSEW 
         Caption         =   "E"
         Height          =   210
         Index           =   2
         Left            =   398
         TabIndex        =   80
         Top             =   218
         Width           =   400
      End
      Begin VB.CommandButton b_NSEW 
         Caption         =   "S"
         Height          =   210
         Index           =   1
         Left            =   200
         TabIndex        =   79
         Top             =   428
         Width           =   400
      End
      Begin VB.CommandButton b_NSEW 
         Caption         =   "W"
         Height          =   210
         Index           =   4
         Left            =   7
         TabIndex        =   78
         Top             =   218
         Width           =   400
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
   Begin SysInfoLib.SysInfo SysInfo 
      Left            =   7560
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CmDia 
      Left            =   7560
      Top             =   765
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton b_RestartDeal 
      Caption         =   "&Repeat"
      Height          =   315
      Left            =   2535
      TabIndex        =   16
      Top             =   6810
      Width           =   810
   End
   Begin VB.CommandButton b_ShowAnalysis 
      Caption         =   "Bid w/ &Analysis"
      Height          =   645
      Left            =   8775
      TabIndex        =   15
      Top             =   6450
      Width           =   975
   End
   Begin VB.PictureBox p_Hand 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1470
      Index           =   1
      Left            =   2640
      ScaleHeight     =   1470
      ScaleWidth      =   4920
      TabIndex        =   7
      Top             =   4935
      Width           =   4920
   End
   Begin VB.PictureBox p_Hand 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6360
      Index           =   4
      Left            =   8685
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
      Left            =   2670
      ScaleHeight     =   1470
      ScaleWidth      =   4875
      TabIndex        =   5
      Top             =   15
      Width           =   4875
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
   Begin VB.PictureBox p_Bottom 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   30
      ScaleHeight     =   1065
      ScaleWidth      =   9750
      TabIndex        =   2
      Top             =   6390
      Width           =   9750
      Begin VB.CommandButton b_DealBackFwd 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Bowlby One SC"
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
         TabIndex        =   83
         Top             =   405
         Width           =   330
      End
      Begin VB.CommandButton b_DealBackFwd 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Bowlby One SC"
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
         TabIndex        =   82
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
         Left            =   6330
         TabIndex        =   8
         Top             =   405
         Width           =   900
      End
      Begin VB.CommandButton b_Ok 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "&Bid's OK"
         Default         =   -1  'True
         Height          =   315
         Left            =   6315
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   930
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
         Width           =   1440
      End
      Begin VB.Label L_Dealer 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   860
         TabIndex        =   10
         Top             =   435
         Width           =   200
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g_Bidding 
      Height          =   2070
      Left            =   5280
      TabIndex        =   0
      Top             =   1560
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   3651
      _Version        =   393216
      Rows            =   8
      Cols            =   4
      FixedCols       =   0
      BackColor       =   49152
      BackColorFixed  =   8454016
      AllowBigSelection=   0   'False
      Enabled         =   0   'False
      TextStyleFixed  =   2
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      ScrollBars      =   0
      FormatString    =   "^   S  |^   W  |^   N  |^    E"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9225
      Top             =   495
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
      Begin VB.Menu mSaveDeal 
         Caption         =   "&Save Deal"
         Shortcut        =   ^S
      End
      Begin VB.Menu zF2 
         Caption         =   "-"
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
         Shortcut        =   ^A
      End
      Begin VB.Menu mHide 
         Caption         =   "&Hide Hands"
         Shortcut        =   ^H
      End
      Begin VB.Menu zP1 
         Caption         =   "-"
      End
      Begin VB.Menu mEnterHands 
         Caption         =   "&Enter Hands"
         Shortcut        =   ^E
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
      Begin VB.Menu mAceAsking 
         Caption         =   "Practice Ace Asking"
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
         Shortcut        =   {F4}
      End
   End
End
Attribute VB_Name = "PlayScr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' CardShark BidBase Bidding Practice Program         '
' 09-27-2018                                         '
'                                                    '
' Project - References:                              '
'        OLE Automation                              '
'        BidBaseDB.vbp or .DLL                       '
' Project - Components:                              '
'        Common Dialog                               '
'        FlexGrid                                    '
'        SysInfo                                     '
'                                                    '
' BidBaseDB.vbp and .DLL are in a separate folder    '
' on the BidBase page on SourceForge since they must '
' be used with most of the BidBase programs.         '
' If you have any BidBase questions, email me at     '
' nfordbb982@aeyec.com                               '

Option Explicit

Public NewBid As BidBaseDB.Bidder ' This is Bidder.cls in BidBaseDB.vbp '

' Global and module variables have a 3-letter abbreviation '
' indicating their variable type.  Local variables in Subs '
' do not have these and their names always start lowercase.'

Private str_AllCalls As String
Private bln_Analysis As Boolean
Private bln_AutoBid As Boolean
Private int_BidRound As Integer
Private bln_BidsFound As Boolean
Private int_BidsFoundNum As Integer
Private str_BidToAnalyze As String
Private str_Call(50) As String
Private int_CallNum As Integer
Private bln_Continue As Boolean
Private bln_DontAsk4DealNum As Boolean
Private lng_ID As Long
Private bln_NoPaint As Boolean
Private int_NumPasses As Integer
Private lng_PaintTimer As Long
Dim int_Player(4) As Integer
Dim str_Player(4) As String
Private str_Source As String  ' Magazine, book, etc.
Private str_StartWith(10) As String


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
        End If
    Next
End Sub


Private Function AllCalls() As String
    ' Computes the Prior Bids string. '
    ' Ex.: P-1C-D-P-1H-P-2H, etc.     '
    
    Dim i As Integer
    Dim a As String
    
    a = ""
    
    For i = 1 To int_CallNum - 1
        
        a = a & str_Call(i)
        
        If i < int_CallNum - 1 Then
            a = a & "-"
        End If
    
    Next
    
    AllCalls = a
End Function


Private Sub SetUpBidDisplayGrid()
    ' Clear the grid that displays '
    ' the biding and rewrite the   '
    ' column headings.             '
    With g_Bidding
        .Clear
        .Row = 0
        .RowHeight(0) = 260
        .RowHeight(7) = 280
        .Col = 0: .Text = "S"
        .Col = 1: .Text = "W"
        .Col = 2: .Text = "N"
        .Col = 3: .Text = "E"
        .ColWidth(0) = 540
        .ColWidth(1) = 540
        .ColWidth(2) = 540
        .ColWidth(3) = 540
    End With
End Sub


Private Sub SetUpBiddingGrid()
    ' Fill out the grid on which '
    ' bids are selected.         '
    
    Dim c As Integer
    Dim r As Integer
  
    With Grid1
        For c = 0 To 4
            .ColWidth(c) = 500
        Next
    
        .RowHeight(7) = 300
    
        For r = 1 To 7
            .Row = r - 1
            For c = 0 To 4
                .Col = c
                .Text = r & Mid$("CDHSN", c + 1, 1) & " "
            Next
        Next
        .Row = 7
        .Col = 0
        .Text = "Pass"
        .Col = 1
        .Text = " Dbl"
        .Col = 2
        .Text = "Rdbl"
        .Col = 3
        .Text = " BID"
        .Col = 4
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


Private Sub b_CancelHands_Click()
    ClearVariables
    b_Deal.Enabled = True
    b_RestartDeal.Enabled = True
    b_Ok.Enabled = True
    b_ShowAnalysis.Enabled = True
    f_FindBidDeals.Visible = True
    
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
    Erase int_SuitQty()
    Erase str_Suits()
    Erase bln_PickedHand()
    Erase str_Hands()
    Erase int_Cards()
   
    b_EnterHands.Visible = True
    L_EnterHandsHelp.Visible = False
    f_CardDeck.Visible = False
    b_CancelHands.Visible = False
    b_DealRest.Visible = False
End Sub

Private Sub b_Deal_Click()
    If lng_DealNumSpecified > 0 Then
        InitiateBidding
        Exit Sub
    End If
    mScramble(1).Enabled = False ' Turn off Remix menu item. '
    Erase str_Call()
    str_AllCalls = ""
    str_Source = ""
    bln_SameDeal = False
    bln_Rebid = False
    SetUpBidDisplayGrid
    b_Ok.Caption = "&Bid's OK"
    
    If lng_DealNum = Val(L_Deal) + 1 Then
        lng_DealNumSpecified = lng_DealNum
        L_Deal = lng_DealNum
        bln_DontAsk4DealNum = True
    End If
    
    If Not bln_DontAsk4DealNum And lng_DealNumSpecified = 0 And lng_DealNum <> Val(Mid$(L_Deal & " ", 2)) + 1 Then
        GetDealNum ' in DealerRoutines module '
    End If
    
    Deal ' in DealerRoutines module '
    lng_DealNumSpecified = 0
    If bln_Abort Then
        bln_Abort = False
    Else
        InitiateBidding
    End If
End Sub


Private Sub b_DealBackFwd_Click(Index As Integer)
    ' Decrement or increment the last deal number.'
    ' Since lng_DealNum will be incremented for   '
    ' the next deal, we decrement by 2 and don't  '
    ' increment. The difference between increment '
    ' here and when clicking Deal is that b_Deal  '
    ' displays the deal number and to just incr., '
    ' you have to click Okay. This saves the click.
    If Index = 0 Then
        lng_DealNum = lng_DealNum - 1
        If ck_KeepDealer.Value = 0 Then
            int_Dealer = int_Dealer - 2 ' It will be increased by 1. '
            int_DealerSpecified = int_Dealer
            SetPlayerNumbers
        End If
    Else
        lng_DealNum = lng_DealNum + 1
        If ck_KeepDealer.Value = 0 Then
            int_Dealer = int_Dealer
            int_DealerSpecified = int_Dealer
            SetPlayerNumbers
        End If
    End If
    bln_DontAsk4DealNum = True
    Timer1.Enabled = True
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
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim tempi As Integer
    Dim numToDeal As Integer
    Dim d(4) As Integer
    Dim s As Integer
    
    For i = 1 To 4
        If int_HandQty(i) = 0 Then
            numToDeal = numToDeal + 1
            j = j + 1
            d(j) = i ' d(j) is an empty hand. '
            bln_PickedHand(i) = True
        End If
    Next
    
    If numToDeal = 1 Then
        ' Given the remaining cards to the only empty hand. '
        For i = 1 To 52
            If Not bln_SelectedCards(i) Then
                k = k + 1
                s = SuitOf(i)
                int_Suit(d(1), s, Ace) = i
            End If
        Next
                
    Else
        j = 0
        For i = 1 To 52
            If bln_SelectedCards(i) Then
                j = j + 1
                int_Cards(j) = i
            End If
        Next
    
        ' Scramble ("shuffle") the array:  '
        For j = 1 To 10 + Int(5 * Rnd) ' # of times to shuffle '
            For i = 1 To j
                ' Get a random array # '
                k = Int(j * Rnd + 1)
                ' Swap with current array #: '
                tempi = int_Cards(i)
                int_Cards(i) = int_Cards(k)
                int_Cards(k) = tempi
            Next
        Next
        
        ' Deal the cards to the empty hands: '
        
    End If
    
    
End Sub


Private Sub b_DoneDealNum_Click()
    f_Inbox.Visible = False
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
End Sub


Private Sub b_Edit_Click()
    Dim xxx As Long
    Dim pc As String
    Dim i As Integer
    
    For i = 1 To int_CallNum - 1
        If i < int_CallNum - 1 Then
            pc = pc & str_Call(i) & "-"
        Else
            pc = pc & str_Call(i)
        End If
    Next
    'If pc = "" Then pc = "none"

    t_EditData = pc & "/" & int_Turn & "/" & str_Hands(int_Player(int_Turn))

    MsgBox ("Deal info has been sent to the BidBase Editor." & vbCr & "You can edit the bidding there.")

    'i = int_Player(int_Turn) - 1 ' It was sending the hand for the next player. '
    'If i = 0 Then i = 1          ' "pc" below is Prior Calls (or Prior Bids).   '
    
    'xxx = Shell(App.Path _
            & "\BidBase Editor.exe " _
            & str_Hands(i) _
            & ", " _
            & pc, _
            vbNormalFocus)
End Sub


Private Sub b_EnterHands_Click()
    Dim i As Integer
    Dim j As Integer
    
    Erase bln_PickedHand()
    int_NumPickedHands = 0
    Erase int_Suit()
    Erase int_SuitQty()
    Erase int_Cards()
    Erase bln_PickedHand()
    
    b_Deal.Enabled = False
    b_RestartDeal.Enabled = False
    b_Ok.Enabled = False
    b_ShowAnalysis.Enabled = False
    t_Disclose.Text = ""
    
    f_FindBidDeals.Visible = False
    b_EnterHands.Visible = False
    f_CardDeck.Visible = True
    f_CardDeck.ZOrder 0
    p_Hand(1).Visible = False
    L_EnterHandsHelp.Visible = True
    SetUpBidDisplayGrid ' Clear the grid that shows the bids. '
    LayOutDeck ' Sub in DealerRoutines '
    For i = 1 To 4
        j = DoEvents()
        p_Hand(i).Cls
        p_Hand(i).Refresh
        j = DoEvents()
    Next
End Sub


Private Sub b_NSEW_Click(Index As Integer)
    If int_HandQty(int_Turn) > 0 And int_HandQty(int_Turn) < 13 Then
        MsgBox "Finishing picking 13 cards for " & str_Player(int_Turn) & "first."
        Exit Sub
    End If
    
    int_Turn = Index ' 1=S, 2=W, 3=N, 4=E '
    If str_Player(1) <> "South" Then
        str_Player(1) = "South"
        str_Player(2) = "East"
        str_Player(3) = "North"
        str_Player(4) = "West"
    End If
    
    L_EnterHandsHelp = vbCrLf & Space$(14) & "Click on 13 cards" & _
                       vbCrLf & Space$(14) & " to give to " & str_Player(Index)
End Sub


Private Sub b_Ok_Click()
    
    If b_Ok.Caption = "Deal" Then
        
        b_Ok.Caption = "&Bid's OK"
        Timer1.Enabled = True
  
    ElseIf b_Ok.Caption = "&Bid's OK" Then
        ProcessBid
        
    ElseIf b_Ok.Caption = "&Continue" Then
        ' Was waiting for user to '
        ' finish editing an entry.'
        b_Ok.Caption = "&Bid's OK"
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
    SetUpBidDisplayGrid
    bln_SameDeal = True
    int_DealerSpecified = InStr("SWNE", Left$(L_Dealer, 1))
    int_Dealer = int_DealerSpecified
    SetPlayerNumbers
    int_CallNum = 0
    int_NumPasses = 0
    b_Ok.Caption = "&Bid's OK"
    bln_Abort = True
    InitiateBidding
End Sub


Private Sub b_ShowAnalysis_Click()
    Dim n As Integer
    
    If bln_Analysis Then
        b_ShowAnalysis.Caption = "Rebid w/ &Analysis"
        b_Ok.Caption = "&Bid's OK"
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
    If L_Vul.Caption = "" Or _
       L_Vul.Caption = "both" _
    Then
        L_Vul.Caption = "none"
    
    ElseIf L_Vul.Caption = "none" Then
        L_Vul.Caption = "N-S"
  
    ElseIf L_Vul.Caption = "N-S" Then
        L_Vul.Caption = "E-W"
  
    Else
        L_Vul.Caption = "both"
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
    
    n = 0 ' number of cards added to Hand. '
    
    p_Hand(1).Visible = True
    
    cardNumInDeck = Index ' KC=13, AC=1, AD=14, KD=26... '
    
    If int_Turn = 0 Then
        MsgBox "Click on N, S, E or W first."
        Exit Sub
    End If
    
    If int_HandQty(int_Turn) = 13 Then
        MsgBox "You've already selected 13 cards."
        Exit Sub
    End If
    
    If bln_SelectedCards(cardNumInDeck) = True Then
        Exit Sub
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    bln_SelectedCards(cardNumInDeck) = True
    
    DrawBack Card(cardNumInDeck).hwnd, 3, 0, 0
    
    ' Go through player's hand to find where the new card goes. '
    ' Cards are numbered 1-52 in C, D, H, S order where A=1 and '
    ' 2-K are 2-13 and suits are 1-4.
    suit = Fix((cardNumInDeck - 1) / 13) + 1
    cardNumInHand = cardNumInDeck - (suit - 1) * 13 ' card # 1-13 '
    
    ' e.g.: Diamond King = 26 - (2-1) * 13 = 13 = King '
    
    ' Put cardNum in player's hand (int_Turn) at position x in Suit: '
    int_Suit(int_Turn, suit, cardNumInHand) = cardNumInDeck
    int_SuitQty(int_Turn, suit) = int_SuitQty(int_Turn, suit) + 1
    int_HandQty(int_Turn) = int_HandQty(int_Turn) + 1
        
    ' Rather than try to figure out where to insert'
    ' new card, it's easier just to redo the hand: '
    If int_HandQty(int_Turn) = 1 Then
        Erase int_Hand()
        int_Hand(int_Turn, 1) = cardNumInDeck
    Else
        s = Spd
        Do
            If int_Suit(int_Turn, s, Ace) Then
                n = n + 1
                int_Hand(4, n) = int_Hand(4, n) + Index
            End If
            For c = int_SuitQty(int_Turn, s) To 2 Step -1
                If int_Suit(int_Turn, s, c) > 0 Then
                    n = n + 1
                    int_Hand(int_Turn, n) = Index
                End If
            Next
            
            If s = Spd Then s = Hrt Else If s = Hrt Then s = Clb Else If s = Clb Then s = Dia Else Exit Do
            ' In hand display, we want to alternate colors: S, H, C, D. '
        Loop
    End If
                
    'For i = 1 To 13: Debug.Print int_Hand(1, i);: Next: Debug.Print ""
    
    DisplayHand int_Turn
    If int_HandQty(int_Turn) = 13 Then
        b_DealRest.Visible = True
        L_EnterHandsHelp = Space$(10) & "Click another player" & vbCrLf & Space$(17) & "to deal to or" & Chr$(13) & Space$(15) & "click Cancel or" & vbCrLf & Space$(11) & "click Deal The Rest."
    End If
End Sub


Private Sub Form_Load()
    Dim n As Integer
    
    ' Only let the program load once: '
    If App.PrevInstance Then
        MsgBox "Program is already running."
        End
    End If
    
    ' Initialize DLL that displays cards: '
    On Local Error GoTo CantLoadDLL
    n = InitializeDeck(PlayScr.hwnd)
    On Local Error GoTo 0
    If n = False Then
        GoTo CantLoadDLL
    End If
    
    LoadCfg
    
    SetUpBiddingGrid
    
    Deal
    
    If b_CancelHands.Visible Then Exit Sub
    
    InitiateBidding

Exit Sub
  
CantLoadDLL:
    MsgBox "Problem loading QCard32.DLL." & vbCr & _
           "Make sure it is in the BidBase directory." & vbCr & _
           "If it is, then try restarting this program." & vbCr & _
           "If it still won't load, restart Windows."
    
    Unload PlayScr
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim FiNum As Integer
    
    Set NewBid = Nothing
    
    FiNum = FreeFile
    Open App.Path & "\BBPractice.cfg" For Output As #FiNum
        Print #FiNum, lng_DealNum
        Print #FiNum, PlayScr.Left, PlayScr.Top
        Write #FiNum, Not bln_AutoBid
        Write #FiNum, Not bln_HideHands
        Print #FiNum, int_DealerSpecified
        Print #FiNum, ck_KeepDealer.Value
        Print #FiNum, ck_FindBids.Value
        Write #FiNum, t_FindBids
    Close #FiNum
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

    If bln_NoPaint Or bln_Abort Then Exit Sub
    bln_NoPaint = True
    
    For h = 1 To 4
        DisplayHand h
    Next

    bln_NoPaint = False
End Sub


Private Sub Form_Resize()
    Form_Paint
End Sub


Private Sub Grid1_Click()
    Dim b As String
    Dim q As String
    Dim X As Long
    
    If Grid1.Row = 7 Then
        ' Pass, Dbl, Redbl - just use 1st letter: '
        b = Left$(Trim$(Grid1.Text), 1)
    Else ' Other bids, use whatever's there: '
        b = Trim$(Grid1.Text)
    End If
    
    If bln_AutoBid Then
        q = "Did you want to change BidBase's bid of " & str_Call(int_CallNum)
        q = q & vbCrLf & " to " & b & " in the database?"
        X = MsgBox(q, vbYesNo)
        If X = vbYes Then
            b_Edit_Click
        End If
    End If
    
    b_Ok.Enabled = True
    BidNextHand ' Let the program bid the same hand. '
End Sub


Private Sub mAceAsking_Click()
    ' 1. Tell the user to select an Ace-Asking convention. '
    ' 2. Randomly select suit lengths (0-6) for each suit. '
    ' 3. Randomly choose 1-4 Aces and the trump K and Q.   '
    ' 4. Put them in the str_Hands(1) array and fill array.'
    ' 5. Put Prior Bids (4N-P or 5N-P or 4N-D or 4N-5suit) '
    ' 6. Randomly assign a trump suit to int-TrumpSuit.    '
    ' 7. Call the MakeBid sub.                             '
    Stop
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
    
    For h = 1 To 4
      
        Do
            hnd = Trim$(UCase$(InputBox$( _
                   "Enter " _
                   & str_Dealer(h) _
                   & "'s hand or leave it blank:" & vbCr _
                   & "(Enter hands as: " & vbCr _
                   & " AKT3-T92--QJT532" & vbCr _
                   & " in suit order S-H-D-C.)" _
                   , "Enter Hands" _
                   , str_Hands(h))))
                   ' Default is last hand used.
                        
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
    
        If int_NumPickedHands = 0 And h = 3 Then
            ' User has entered 3 hands. '
            ' We can fill in the fourth.'
            bln_PickedHand(4) = True
            Exit For
        End If
    
    Next
  
    FillInTheDeal ' Puts remaining cards for the '
                  ' random hands into int_Cards()'
                  
    ScrambleHands ' Mix up cards in int_Cards()  '
                  ' for the random hands.        '
                  
    SortHands ' Sort hands within int_Cards().   '
    
    FillSuits ' Assign int_Cards() to int_Suit().'
        
    For h = 1 To 4
        DisplayHand h
    Next
    
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
        s = SuitOf(i) ' e.g.: #3 is C3, '
        c = i Mod 13  '       #16 is D3 '
        If c = 0 Then c = 13
        ' If not already used... '
        If int_Suit(0, s, c) = 0 Then
            ' Keep track of cards used: '
            int_Suit(0, s, c) = i
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


Private Sub mHelp_Click()
    Dim f1 As String
    Dim f2 As String
    Dim xx As Long
    
    f1 = App.Path & "\Notes\CardShark_BidBase_Practice_Program.htm"
    f2 = App.Path & "\Notes\CardShark_BidBase_Editor_Menus.htm"
    xx = Shell(App.Path & "\BidBase FileViewer.exe " & _
            f1 & " " + f2, _
            vbNormalFocus)
End Sub


Private Sub mHide_Click()
    Dim h As Integer
    
    If bln_AutoBid Then
        bln_HideHands = False
        mHide.Checked = False
    Else
        bln_HideHands = Not bln_HideHands
        mHide.Checked = bln_HideHands
    End If
    
    For h = 1 To 4
        DisplayHand h
    Next
End Sub

Private Sub mSaveDeal_Click()
    ' Appends the current deal to the '
    ' SavedDeals.txt file in App.Path.'
    
    Dim f As String
    Dim FiNum As Integer
    Dim i As Integer
    Dim xx As Long
    
    f = App.Path & "\Notes\SavedDeals.txt"
    
    FiNum = FreeFile
    Open f For Append As #FiNum
        Print #FiNum, ""
        Print #FiNum, ""
        Print #FiNum, "Deal #"; L_Deal
        Print #FiNum, "N:  "; str_Hands(3)
        Print #FiNum, "S:  "; str_Hands(1)
        Print #FiNum, "W:  "; str_Hands(2)
        Print #FiNum, "E:  "; str_Hands(4)
        Print #FiNum, ""
        Print #FiNum, Left$(L_Dealer, 1); " dealt."
        Print #FiNum, "Bidding:"
        For i = 1 To int_CallNum
            Print #FiNum, str_Call(i);
            If i < int_CallNum And i Mod 4 > 0 Then
                Print #FiNum, "-";
            Else
                Print #FiNum, ""
            End If
        Next
        Print #FiNum, ""
    Close #FiNum
    ' Display file in Notepad: '
    xx = Shell("c:\windows\notepad.exe " + f, vbNormalFocus)
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
    
    int_NumPickedHands = 0
    
    If Index = 1 Then
        
        ' Rescramble previously selected hands: '
        For h = 1 To 4
            If bln_PickedHand(h) Then
                int_NumPickedHands = _
                int_NumPickedHands + 1
            End If
        Next
    
        If int_NumPickedHands = 0 Then
            MsgBox "No record of hands selected." & vbCr & _
                   "Please enter again."
            mScramble_Click 0
            Exit Sub
        End If
    
    Else
    
        rh = Trim$(UCase$(InputBox$( _
            "Enter the hands (N,S,E,W) to scramble.")))
    
        If rh = "" Then Exit Sub
    
        Erase bln_PickedHand()
        For h = 1 To 4
            If InStr(rh, Mid$("SWNE", h, 1)) Then
                bln_PickedHand(h) = True
                int_NumPickedHands = _
                int_NumPickedHands + 1
            End If
        Next
    End If
  
    If int_NumPickedHands < 2 Then
        MsgBox "You must specify at least 2 hands to Scramble."
        mScramble_Click 0
        Exit Sub
    
    ElseIf int_NumPickedHands > 3 Then
        MsgBox "You cannot Scramble more than 3 hands." _
            & vbCr & "To Scramble all 4, just click on Deal."
        mScramble_Click 0
        Exit Sub
    End If

    ScrambleHands
    SortHands ' Sort hands within int_Cards().   '
    FillSuits ' Assign int_Cards() to int_Suit().'
    
    For h = 1 To 4
        DisplayHand h
    Next
    b_RestartDeal_Click
    mScramble(1).Enabled = True

End Sub


Private Sub mTutorial_Click()
    mTutorial.Checked = True
    Welcome
    mTutorial.Checked = False
End Sub


Private Sub p_Hand_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim lastCard As Long
    Dim cardNumInHand As Long
    Dim cardNumInDeck As Long
    Dim CardNumInSuit As Long
    Dim numPix As Long
    Dim suit As Integer

    lastCard = int_HandQty(Index)
    
    ' The cards are just drawn on 1 picture control. '
    ' Calculate which card was clicked on: '
    If Index = 1 Or Index = 3 Then
        ' S and N cards go left to right 300 pix each. '
        numPix = lastCard * 300
        If X > numPix Then
            cardNumInHand = lastCard
        Else
            cardNumInHand = Fix(X / 300) + 1
        End If
    Else ' E and W card go top to bottom 400 pix each. '
        numPix = lastCard * 400
        If Y > numPix Then
            cardNumInHand = lastCard
        Else
            cardNumInHand = Fix(Y / 400) '+ 1
        End If
    End If
        
    cardNumInDeck = int_Hand(Index, cardNumInHand)
        
    bln_SelectedCards(cardNumInDeck) = False
    
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
    int_SuitQty(Index, suit) = int_SuitQty(Index, suit) - 1
    int_HandQty(Index) = int_HandQty(Index) - 1
    DisplayHand Index
End Sub


Private Sub t_FindBids_LostFocus()
    Dim t As String
    Dim i As Long
    Dim n As Long
    
    t = Trim$(t_FindBids)
    If t = "" Then Exit Sub
    Erase str_StartWith()
        
    ' e.g.: "1H, 1S-P" - split into arrays. '
    Do
        i = InStr(t, ",")
        n = n + 1
        If i > 0 Then
            str_StartWith(n) = Left$(t, i - 1)
            t = Trim$(Mid$(t, i + 1))
        Else
            str_StartWith(n) = t
            Exit Sub
        End If
    Loop
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
    
    If Dir$(App.Path & "\Notes\" & f & ".htm") <> "" Then
        xx = Shell(App.Path & "\BidBase FileViewer.exe " _
             & App.Path & "\Notes\" & f & ".htm", 1)
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

    SetPlayerNumbers
    
    Erase str_Call()
    str_AllCalls = ""
    bln_SameDeal = False
    
    ' If not a user-entered hand or if '
    ' user-entered but no prior calls, '
    ' initialize the bidding variables:'
    If L_Deal <> "User" Or int_CallNum = 0 Then
        int_Turn = 0
        int_CallNum = 1
        int_BidRound = 1
        int_NumPasses = 0
    End If
    
    b_Edit.Enabled = True
    b_RestartDeal.Enabled = True
    Deal
    BidNextHand
    
End Sub

Sub BidNextHand()
    ' Prepares for and calls MakeBid.     '
    '                                     '
    ' This routine is called first by the '
    ' InitiateBidding routine and then by '
    ' the ProcessBid routine to continue  '
    ' the bidding until it is passed out. '
    
    Dim n As Integer
    
    If bln_Rebid Then
        str_Call(int_CallNum) = ""
        str_AllCalls = AllCalls
    Else
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
        And Not bln_AutoBid _
    Then
        ' Let user bid first. If '
        ' other hands are hidden,'
        ' user just bids South;  '
        ' if hands aren't hidden,'
        ' user bids all hands.   '
        ' After the user picks a '
        ' bid, the program then  '
        ' bids and compares bids.'
        b_Ok.Enabled = False
        b_Edit.Enabled = False
        Grid1.Visible = True
        bln_Continue = False
        Exit Sub
    End If
            
    If int_Player(int_Turn) = 0 Then
        SetPlayerNumbers
    End If
    
    ' Let BidBase bid: '
    MakeBid int_Player(int_Turn)
        
    If Timer1.Enabled Then
        ' Bid found for deal did not match Find Bids. '
        Exit Sub
    End If
        
    If ck_FindBids.Value = 0 Then
        For n = 1 To 4
            DisplayHand int_Player(n)
        Next
    End If
        
    b_Ok.Enabled = True
    
End Sub


Sub ProcessBid()
    ' After the DLL makes the bid, it is  '
    ' shown to the user who can click any '
    ' of the buttons, menu options, etc.  '
    ' If he clicks "Bid's OK", then this  '
    ' routine is called to accept/process '
    ' the bid, then this ends by calling  '
    ' BidNextHand unless it's passed out. '
    
    Dim n As Integer
        
        If bln_Abort Then
            GoTo ExitBidHands
        End If
        
        If Not bln_Rebid Then
            If str_Call(int_CallNum) = "P" Then
                int_NumPasses = int_NumPasses + 1
            Else
                int_NumPasses = 0
            End If
            
            int_CallNum = int_CallNum + 1
                
        ElseIf bln_Rebid Then
            str_AllCalls = ""
            str_Call(int_CallNum) = ""
            For n = 1 To int_CallNum - 1
                str_AllCalls = str_AllCalls _
                & str_Call(n) & "-"
            Next
            If Len(str_AllCalls) > 0 Then
                str_AllCalls = _
                    Left$(str_AllCalls, _
                    Len(str_AllCalls) - 1)
            End If
        End If
        
    If int_NumPasses < 3 Or int_CallNum < 5 Then
        BidNextHand
    End If
    
ExitBidHands:
    If int_NumPasses = 3 Then
        b_Edit.Enabled = False
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
    Dim p As String
    Dim pDash As String
    Dim fb As String
    Dim i As Integer
    Dim int_TrumpSuit As Integer
    
    t_Disclose = ""
    str_AllCalls = Replace(str_AllCalls, "?", "")
    'bln_BidsFound = False ' This is done below if no specified bid is found. '
    
    'Submit specs to DLL:
    With NewBid
        b_Ok.Enabled = False
        .DBPathAndName = "BidBase.accdb"
        .PlayersHand = str_Hands(h)
        .PriorCalls = str_AllCalls
        .Vul = Trim$(Left$(L_Vul.Caption & " ", 1))
        .BidToAnalyze = Trim$(str_BidToAnalyze)
        .GetDisclosures = True
        .TrumpSuit = 0 ' When the bidder needs to know the trump suit, it asks. '
        
        ' Get bid: '
        ' If this doesn't work, quit and check Project - References '
        ' to make sure that BidBaseDb.vbp or .DLL (whichever you are'
        ' using) is checked. '
        b = .Bid
        c = .ErrCode
        d = .Disclosure
        
        If ck_FindBids.Value = 1 Then
            ' t_FindBids could be something like "1H-P, 1S-P"     '
            ' where a deal with either of those bidding sequences '
            ' would be acceptable. The sequences are placed into  '
            ' the array str_StartWith() such as (1)="1H-P" and    '
            ' str_StartWith(2)="1S-P". The user could change the  '
            ' specs at any time and ck_FindBids_LostFocus would   '
            ' redo the array. '
            If b = "" Then b = "P"
            ' b can be any of 1C to 7N or P, D, R '
            p = str_AllCalls
            pDash = p & "-"
            
            'DisplayHand 3
            i = 0
            Do While Not bln_BidsFound And str_StartWith(i + 1) <> ""
                i = i + 1
                If str_StartWith(i) = p & b Then
                    ' Could be ["" & 1S] or [1S- & P], etc. '
                    bln_BidsFound = True
                    int_BidsFoundNum = i
                    Exit Do
                End If
            Loop
            
            'Stop ' needs more work '
            If Left$(p, Len(str_StartWith(int_BidsFoundNum))) <> str_StartWith(int_BidsFoundNum) And str_StartWith(int_BidsFoundNum) <> p & b Then
                bln_BidsFound = False
            End If
            
            If Not bln_BidsFound Then
                ' If "1H-P" then see if PriorBids & b = specification.
                ' e.g.: str_AllCalls = "1H" & "-" and bid is "P"   '
                ' and t_FindBids is "1S-P, 1H-P" then bid is found '
                ' otherwise, get next deal. '
                lng_DealNum = lng_DealNum + 1
                lng_DealNumSpecified = lng_DealNum
                bln_BidsFound = False
                Timer1.Enabled = True
                Exit Sub
            End If
        End If
        
        If b = "" And c < 3 Then
            ' If c=1 or 2, no bid was found. '
            b = "P?"
            c = 0
        End If
        
        If c > 0 And c <> 2 Then
            t_Disclose = .ErrMsg
        
        Else
            str_Call(int_CallNum) = b
            
            If int_Turn = 1 And int_BidRound = 1 Then
                str_AllCalls = b
            Else
                str_AllCalls = str_AllCalls & "-" & b
            End If
            
            If ck_FindBids.Value = 1 Then
                If InStr(fb, str_AllCalls & "-") > 0 Then
                    ProcessBid
                    Exit Sub
                Else
                    bln_BidsFound = True
                    For i = 1 To 4
                        DisplayHand i
                    Next
                End If
            End If
                
            DisplayBid int_Player(int_Turn)
                If Trim$(d) <> "" And b <> "P" Then
                    ' Don't show disclosure for Pass '
                    t_Disclose = d
                End If
            'End If
            
            If Not bln_AutoBid _
               And (int_Player(int_Turn) = 1 _
                    Or Not bln_HideHands) _
            Then
                c = Len(d)
                If Trim$(str_Call(0)) = b Then
                    b = "       ***** BidBase bid " _
                        & "the same as you. *****'"
                Else
                    b = "       ***** You bid " _
                        & str_Call(0) _
                        & " and BidBase bid " _
                        & b & ". *****"
                End If
                t_Disclose = t_Disclose & vbCrLf & b
                t_Disclose.SelStart = c
                t_Disclose.SelLength = Len(b) + 6
            End If
        End If
    End With
    
    If str_BidToAnalyze <> "" Then
        bln_Analysis = False
        str_BidToAnalyze = ""
        b_ShowAnalysis.Caption = "Rebid w/ &Analysis"
    End If
    b_Ok.Enabled = True
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
    
    SetUpBidDisplayGrid
    
With g_Bidding
    i = int_Dealer - 1 ' S is column 0 but Dealer 1, etc. '
    .Row = 1
    n = 0
    For r = 1 To 7
        For c = 0 To 3
            If c = i Then
                .Col = c
                n = n + 1
                .Text = str_Call(n)
                If str_Call(n) = "" Then Exit Sub
                i = i + 1
                If i > 3 Then
                    i = 0
                    .Row = .Row + 1
                    Exit For
                End If
            End If
        Next
    Next

    ' Set column: '
    If h > 1 Then
        .Col = h - 2 ' Columns are 0-3
    Else
        .Col = 3 ' South is last col.
    End If

    ' Set row:    '
    If int_Player(1) = 2 Then ' West dealt
        .Row = int_BidRound
    
    ElseIf int_Player(1) = 3 Then ' North dealt
        If h = 2 Then
            .Row = int_BidRound + 1
        Else
            .Row = int_BidRound
        End If
    
    ElseIf int_Player(1) = 4 Then ' East dealt
        If h = 2 Or h = 3 Then
            .Row = int_BidRound + 1
        Else
            .Row = int_BidRound
        End If
    
    Else  ' South dealt
        If h > 1 Then
            .Row = int_BidRound + 1
        Else
            .Row = int_BidRound
        End If
    End If
    
    .Text = str_Call(int_CallNum)
    If .Text <> "P" Then
        int_NumPasses = 0
    End If

End With
  
    If bln_Analysis Then
        b_ShowAnalysis.Caption = "Rebid w/ &Analysis"
        b_Ok.Caption = "&Bid's OK"
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
        int_DealerSpecified = 0
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

    L_Dealer.Visible = True
    L_Dealer = Mid$("SWNE", int_Dealer, 1)
    ck_KeepDealer.Visible = True
    
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
End Sub


Private Sub LoadCfg()
    ' Load the configuration file which '
    ' is saved when the program ends.   '
    Dim f As String
    Dim p As String
    Dim FiNum As Integer
    Dim lng_PlayScrLeft As Long
    Dim lng_PlayScrTop As Long
    Dim i As Long
    Dim findBids As String
    
    If Dir$(App.Path & "\BBPractice.cfg") = "" Then
        If Dir$(App.Path & "\BBPractice.cfg") = "" Then
            MsgBox App.Path & "\BBPractice new.cfg is missing." & vbCr & "Default settings will be used."
            Exit Sub
        End If
        f = App.Path & "\BBPractice new.cfg"
        ' Load the settings from the initial file. '
        ' A Practice.cfg file will be created when '
        ' quitting the program. '
    Else
        f = App.Path & "\BBPractice.cfg"
    End If
    
    FiNum = FreeFile
    Open f For Input As #FiNum
        Input #FiNum, lng_DealNum
        Input #FiNum, lng_PlayScrLeft, lng_PlayScrTop
        Input #FiNum, bln_AutoBid
        Input #FiNum, bln_HideHands
        Input #FiNum, int_DealerSpecified
            L_Dealer = Mid$("SWNE", int_DealerSpecified, 1)
            L_Dealer.Visible = True
            int_Dealer = int_DealerSpecified
        Input #FiNum, i
           ck_KeepDealer.Value = i
           ck_KeepDealer.Visible = True
        Input #FiNum, i
           ck_FindBids.Value = i
        Input #FiNum, findBids
           t_FindBids = findBids
           t_FindBids.FontBold = True
           t_FindBids_LostFocus
    Close #FiNum
    
    ' Other initializations: '
    Randomize
    CARDWIDTH = 100
    CARDHEIGHT = 200
    ' Set ScaleMode to Pixel(3) so the form and '
    ' card-display DLL use the same coordinates '
    Me.ScaleMode = 3
    Me.Show
    Me.Move lng_PlayScrLeft, lng_PlayScrTop
    
    mAutoBid_Click
    mHide_Click
    
    Set NewBid = New BidBaseDB.Bidder
    
    str_Dealer(1) = "South"
    str_Dealer(2) = "West"
    str_Dealer(3) = "North"
    str_Dealer(4) = "East"
    
    SetUpBidDisplayGrid
End Sub


Private Sub CheckBattery()
    ' When I'm running on battery, I usually '
    ' don't have room to see a charge-left   '
    ' display, so I like having it in my     '
    ' program. That's what this routine does.'
    
    If SysInfo.ACStatus = 1 Then 'AC power
        'lblBatt.Visible = False
        'lblBattStat.Visible = False
        lblBattTime.Visible = False
        lblTime.Visible = False
        Exit Sub
    End If
  
    'lblBatt.Visible = True
    'lblBattStat.Visible = True
    lblBattTime.Visible = True
    lblTime.Visible = True
  
    Dim PerCentLeft As String
    
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
  
    If SysInfo.BatteryLifeTime <> &HFFFFFFFF Then
        lblTime.Caption = Format((TimeSerial(0, 0, _
        SysInfo.BatteryLifeTime)), "h:mm")
    Else
        lblTime.Caption = "Unknown"
    End If
End Sub


Private Sub t_FindBids_GotFocus()
    If t_FindBids.FontBold = False Then
        ' Remove initial prompt. Don't think it needs to be put back. '
        t_FindBids = ""
        t_FindBids.FontBold = True
    End If
End Sub

Private Sub t_FindBids_KeyPress(KeyAscii As Integer)
    Dim c As String
    Dim bids As String
    Dim suit As String
    Dim priorChar As String
    
    c = UCase$(Chr$(KeyAscii))
    KeyAscii = Asc(c)
    
    bids = t_FindBids
    priorChar = Trim$(Right$(" " & bids, 1))
    
    suit = InStr("CDHSNPDR", c) ' Okay, so PassDoubleRedouble is not really a "suit"... '
    If suit = 0 And InStr("1234567-, ", c) = 0 And KeyAscii <> 8 Then
        MsgBox "'" & c & "' is not a valid bidding character."
        KeyAscii = 0
    ElseIf c = " " And priorChar <> "," Then
        ' A space is allowed only after a comma. '
        KeyAscii = 0
    ElseIf c = "," Then
        If InStr("CDHSNPDR", priorChar) = 0 Then
            KeyAscii = 0 ' Comma only allowed after "CDHSNPDR". '
        Else
            t_FindBids = t_FindBids & ", "
            KeyAscii = 0
            t_FindBids.SelStart = Len(t_FindBids) + 1
        End If
    ElseIf Val(priorChar) = 0 And suit > 0 And suit < 6 Then
        ' If priorChar not a number and '
        ' this key is a suit, then BAD! '
        KeyAscii = 0
    End If
    ' Anything that is not KeyAscii = 0 by this point is Okay by me. '
End Sub


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
    b_Deal_Click
    InitiateBidding
End Sub

