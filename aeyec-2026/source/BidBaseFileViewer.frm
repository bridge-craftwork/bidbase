VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form BidBaseFileViewer 
   BackColor       =   &H8000000A&
   Caption         =   "CardShark BidBase HTML Viewer/Editor 3.0"
   ClientHeight    =   9465
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15525
   Icon            =   "BidBaseFileViewer.frx":0000
   ScaleHeight     =   9465
   ScaleWidth      =   15525
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox t_Back 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5955
      TabIndex        =   34
      Text            =   "<"
      Top             =   20
      Width           =   570
   End
   Begin VB.PictureBox brwWebBrowser 
      Height          =   7185
      Index           =   1
      Left            =   6915
      ScaleHeight     =   7125
      ScaleWidth      =   6375
      TabIndex        =   2
      Top             =   525
      Visible         =   0   'False
      Width           =   6435
   End
   Begin VB.Frame frmColors 
      BackColor       =   &H8000000A&
      Caption         =   "Page Colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3510
      Left            =   6930
      TabIndex        =   7
      Top             =   555
      Visible         =   0   'False
      Width           =   3000
      Begin VB.CommandButton cmdGetColor 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   1860
         TabIndex        =   32
         Top             =   2880
         Width           =   810
      End
      Begin VB.TextBox t_FormColorCode 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   1830
         TabIndex        =   31
         Text            =   "000000"
         Top             =   2400
         Width           =   810
      End
      Begin VB.TextBox t_FormColorCode 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   1830
         TabIndex        =   30
         Text            =   "000000"
         Top             =   1905
         Width           =   810
      End
      Begin VB.TextBox t_FormColorCode 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   1830
         TabIndex        =   29
         Text            =   "000000"
         Top             =   1410
         Width           =   810
      End
      Begin VB.TextBox t_FormColorCode 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   1830
         TabIndex        =   28
         Text            =   "000000"
         Top             =   915
         Width           =   810
      End
      Begin VB.TextBox t_FormColorCode 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   1830
         TabIndex        =   27
         Text            =   "000000"
         Top             =   420
         Width           =   810
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFC0&
         Height          =   3030
         Left            =   105
         ScaleHeight     =   2970
         ScaleWidth      =   1395
         TabIndex        =   8
         Top             =   315
         Width           =   1455
         Begin VB.CommandButton cmdGetColor 
            Caption         =   "Text Color"
            Height          =   345
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   105
            Width           =   1155
         End
         Begin VB.CommandButton cmdGetColor 
            Caption         =   "Background"
            Height          =   345
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   585
            Width           =   1155
         End
         Begin VB.CommandButton cmdGetColor 
            Caption         =   "Link"
            Height          =   345
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   1065
            Width           =   1155
         End
         Begin VB.CommandButton cmdGetColor 
            Caption         =   "Visited Link"
            Height          =   345
            Index           =   3
            Left            =   120
            TabIndex        =   11
            Top             =   1560
            Width           =   1155
         End
         Begin VB.CommandButton cmdGetColor 
            Caption         =   "Active Link"
            Height          =   345
            Index           =   4
            Left            =   120
            TabIndex        =   10
            Top             =   2040
            Width           =   1155
         End
         Begin VB.CommandButton cmdGetColor 
            Caption         =   "Close"
            Height          =   345
            Index           =   5
            Left            =   300
            TabIndex        =   9
            Top             =   2520
            Width           =   840
         End
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   13920
      Top             =   2715
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CmDia 
      Left            =   13860
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstSuggestions2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5340
      Left            =   4920
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   8700
      Visible         =   0   'False
      Width           =   4020
   End
   Begin VB.TextBox LinkMsgBox 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   1770
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "BidBaseFileViewer.frx":0CCA
      Top             =   1785
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.Frame frmSuggestions 
      Caption         =   "Spellings:"
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
      Height          =   5370
      Left            =   6840
      TabIndex        =   6
      Top             =   495
      Visible         =   0   'False
      Width           =   1560
      Begin VB.TextBox t_Word 
         Height          =   314
         Left            =   30
         TabIndex        =   23
         Top             =   450
         Width           =   1470
      End
      Begin VB.CommandButton cmdFixSpelling 
         Caption         =   "&Replace with:"
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
         Left            =   45
         MaskColor       =   &H000000FF&
         TabIndex        =   18
         Top             =   1665
         Width           =   1500
      End
      Begin VB.CommandButton cmdSkipWord 
         Caption         =   "Ignore word - all"
         Height          =   300
         Index           =   1
         Left            =   45
         TabIndex        =   17
         Top             =   1380
         Width           =   1500
      End
      Begin VB.TextBox t_Replace 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   314
         Left            =   45
         TabIndex        =   19
         Top             =   2025
         Width           =   1470
      End
      Begin VB.CommandButton cmdCloseSpellCk 
         Caption         =   "&Close"
         Height          =   315
         Left            =   30
         TabIndex        =   21
         Top             =   5010
         Width           =   1500
      End
      Begin VB.CommandButton cmdSkipWord 
         Caption         =   "&Ignore word - once"
         Height          =   300
         Index           =   0
         Left            =   45
         TabIndex        =   16
         Top             =   1095
         Width           =   1500
      End
      Begin VB.CommandButton cmdAddWord 
         Caption         =   "&Add word"
         Height          =   300
         Left            =   45
         TabIndex        =   15
         Top             =   810
         Width           =   1500
      End
      Begin VB.ListBox lstSuggestions 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         ItemData        =   "BidBaseFileViewer.frx":0D5B
         Left            =   45
         List            =   "BidBaseFileViewer.frx":0D5D
         TabIndex        =   20
         Top             =   2370
         Width           =   1470
      End
      Begin VB.Label Label2 
         Caption         =   "Word:"
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
         Left            =   75
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox L_Name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   6840
      TabIndex        =   4
      Top             =   165
      Width           =   6315
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13935
      Top             =   3585
   End
   Begin VB.PictureBox pSplitterBar 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7635
      Left            =   6540
      ScaleHeight     =   7635
      ScaleWidth      =   345
      TabIndex        =   5
      Top             =   525
      Width           =   345
   End
   Begin VB.PictureBox brwWebBrowser 
      Height          =   7665
      Index           =   0
      Left            =   285
      ScaleHeight     =   7605
      ScaleWidth      =   6360
      TabIndex        =   1
      Top             =   615
      Width           =   6420
   End
   Begin VB.TextBox L_Name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   180
      TabIndex        =   24
      Top             =   165
      Width           =   6315
   End
   Begin VB.PictureBox imlToolbarIcons 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   13785
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   33
      Top             =   1215
      Width           =   1200
   End
   Begin VB.PictureBox Text2 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7680
      Index           =   0
      Left            =   195
      ScaleHeight     =   7620
      ScaleWidth      =   6345
      TabIndex        =   25
      Top             =   525
      Visible         =   0   'False
      Width           =   6405
   End
   Begin VB.PictureBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7635
      Index           =   1
      Left            =   6825
      ScaleHeight     =   7575
      ScaleWidth      =   6345
      TabIndex        =   26
      Top             =   495
      Visible         =   0   'False
      Width           =   6405
   End
   Begin VB.Menu mF 
      Caption         =   "&File"
      Begin VB.Menu mNewFile 
         Caption         =   "&New File"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mOpen 
         Caption         =   "&Open File"
         Shortcut        =   ^O
      End
      Begin VB.Menu mOpenRight 
         Caption         =   "Open Links/Files in &Right Window"
      End
      Begin VB.Menu zF 
         Caption         =   "-"
      End
      Begin VB.Menu mCloseRight 
         Caption         =   "&Close Right Panel"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mSwap 
         Caption         =   "S&wap Panels"
         Shortcut        =   {F7}
      End
      Begin VB.Menu zF0 
         Caption         =   "-"
      End
      Begin VB.Menu mChangeMonitors 
         Caption         =   "Change &Monitors"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mResetMon 
         Caption         =   "Return to Main Monitor"
      End
      Begin VB.Menu zF1 
         Caption         =   "-"
      End
      Begin VB.Menu mDeleteFile 
         Caption         =   "&Delete File"
      End
      Begin VB.Menu mSave 
         Caption         =   "&Save Source"
         Index           =   0
         Shortcut        =   ^S
      End
      Begin VB.Menu mSave 
         Caption         =   "Save Source &As"
         Index           =   1
      End
      Begin VB.Menu mSaveWebPage 
         Caption         =   "Save Web page"
      End
      Begin VB.Menu z31 
         Caption         =   "-"
      End
      Begin VB.Menu mRestore 
         Caption         =   "Restore Original Text"
         Index           =   0
      End
      Begin VB.Menu mRestore 
         Caption         =   "Restore Before Last Save"
         Index           =   1
      End
      Begin VB.Menu mRestore 
         Caption         =   "Restore At Last Save"
         Index           =   2
      End
      Begin VB.Menu z4 
         Caption         =   "-"
      End
      Begin VB.Menu mPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mPrintPreview 
         Caption         =   "Print Preview"
      End
      Begin VB.Menu z4a 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mRecent 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mRecent 
         Caption         =   "Recent file 1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mRecent 
         Caption         =   "Recent file 2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mRecent 
         Caption         =   "Recent file 3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mRecent 
         Caption         =   "Recent file 4"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mRecent 
         Caption         =   "Recent file 5"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mRecent 
         Caption         =   "Recent file 6"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mRecent 
         Caption         =   "Recent file 7"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mRecent 
         Caption         =   "Recent file 8"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mRecent 
         Caption         =   "Recent file 9"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mRecent 
         Caption         =   "Recent file 10"
         Index           =   10
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mE 
      Caption         =   "&Edit"
      Begin VB.Menu mEditHTML 
         Caption         =   "&Edit HTML"
         Shortcut        =   {F11}
      End
      Begin VB.Menu zE1 
         Caption         =   "-"
      End
      Begin VB.Menu mFind 
         Caption         =   "&Find/Replace Text"
         Index           =   0
         Shortcut        =   ^F
      End
      Begin VB.Menu mFind 
         Caption         =   "Find &Again"
         Index           =   1
         Shortcut        =   ^A
         Visible         =   0   'False
      End
      Begin VB.Menu mFind 
         Caption         =   "&Replace Text"
         Index           =   2
         Shortcut        =   {F9}
         Visible         =   0   'False
      End
      Begin VB.Menu mSpellCheckDocument 
         Caption         =   "SpellCheck &Document"
         Enabled         =   0   'False
         Shortcut        =   {F12}
         Visible         =   0   'False
      End
      Begin VB.Menu mSpellCheckWord 
         Caption         =   "SpellCheck &Word"
         Enabled         =   0   'False
         Shortcut        =   ^W
         Visible         =   0   'False
      End
      Begin VB.Menu z11b 
         Caption         =   "-"
      End
      Begin VB.Menu mTableOfContents 
         Caption         =   "Make Table Of Contents"
      End
   End
   Begin VB.Menu mH 
      Caption         =   "&HTML"
      Begin VB.Menu mHTMLStartEnd 
         Caption         =   "&HTML Start-End"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu z7 
         Caption         =   "-"
      End
      Begin VB.Menu mC 
         Caption         =   "&Colors"
         Begin VB.Menu mColors 
            Caption         =   "&Page Colors"
            Index           =   0
         End
         Begin VB.Menu mColors 
            Caption         =   "&Text Colors"
            Index           =   1
         End
         Begin VB.Menu mColors 
            Caption         =   "Background-Text Colors"
            Index           =   2
         End
         Begin VB.Menu mColors 
            Caption         =   "&Color Code"
            Index           =   3
         End
      End
      Begin VB.Menu z7b 
         Caption         =   "-"
      End
      Begin VB.Menu mHd 
         Caption         =   "Headers"
         Begin VB.Menu mHeader 
            Caption         =   "Header &1"
            Index           =   0
            Shortcut        =   {F1}
         End
         Begin VB.Menu mHeader 
            Caption         =   "Header &2"
            Index           =   1
            Shortcut        =   {F2}
         End
         Begin VB.Menu mHeader 
            Caption         =   "Header &3"
            Index           =   2
            Shortcut        =   {F3}
         End
      End
      Begin VB.Menu z2 
         Caption         =   "-"
      End
      Begin VB.Menu mHL 
         Caption         =   "Hyper&Links"
         Begin VB.Menu mLinkExt 
            Caption         =   "Link &External"
            Shortcut        =   ^E
         End
         Begin VB.Menu mLinkIn 
            Caption         =   "&Link Internal"
            Shortcut        =   ^L
         End
      End
      Begin VB.Menu zHide 
         Caption         =   "-"
      End
      Begin VB.Menu zHT 
         Caption         =   "Hide Text"
         Begin VB.Menu mHideText 
            Caption         =   "&Hide Text"
            Shortcut        =   ^{INSERT}
         End
         Begin VB.Menu mHideDeeper 
            Caption         =   "Hide Text Deeper"
         End
      End
      Begin VB.Menu zBrk 
         Caption         =   "-"
      End
      Begin VB.Menu zLB 
         Caption         =   "Line Breaks"
         Begin VB.Menu mParagraph 
            Caption         =   "&Paragraph"
            Shortcut        =   ^P
         End
         Begin VB.Menu mHalfLineAbove 
            Caption         =   "Half-line Break Above"
         End
         Begin VB.Menu mHalfLineBelow 
            Caption         =   "Half-line Break Below"
         End
         Begin VB.Menu mHalfLineAll 
            Caption         =   "Half-line Break - All"
         End
         Begin VB.Menu mPixelSpacer 
            Caption         =   "1x30-Pixel Spacer"
         End
         Begin VB.Menu mBreak 
            Caption         =   "Brea&k"
            Shortcut        =   ^K
         End
         Begin VB.Menu mNoBreak 
            Caption         =   "No-Break"
         End
      End
      Begin VB.Menu zFlow 
         Caption         =   "-"
      End
      Begin VB.Menu zTF 
         Caption         =   "Text Flow"
         Begin VB.Menu mCenter 
            Caption         =   "Center"
         End
         Begin VB.Menu mPre 
            Caption         =   "Pre(formatted text)"
         End
         Begin VB.Menu mSuperscript 
            Caption         =   "Superscript"
            Shortcut        =   ^{F11}
         End
         Begin VB.Menu mSubscript 
            Caption         =   "Subscript"
            Shortcut        =   ^{F12}
         End
         Begin VB.Menu mWrapAround 
            Caption         =   "Wrap around image"
         End
      End
      Begin VB.Menu zL1 
         Caption         =   "-"
      End
      Begin VB.Menu zLists 
         Caption         =   "Lists"
         Begin VB.Menu mNumberedList 
            Caption         =   "&Numbered List <OL>"
            Shortcut        =   ^N
         End
         Begin VB.Menu mIndent 
            Caption         =   "I&ndent w/ Bullets <UL>"
         End
         Begin VB.Menu mBulletsNoIndent 
            Caption         =   "&Bullet List w/o Indent"
         End
         Begin VB.Menu mIndentNoBreak 
            Caption         =   "I&ndent Without Break"
         End
         Begin VB.Menu mIndentWithoutBreakAll 
            Caption         =   "Indent W/O Break - All"
         End
      End
      Begin VB.Menu z3a 
         Caption         =   "-"
      End
      Begin VB.Menu zSC 
         Caption         =   "&Special Characters"
         Begin VB.Menu zPix 
            Caption         =   "Pixel Spacing"
            Begin VB.Menu mPixelSpacing 
               Caption         =   "2-Pixel Space"
               Index           =   0
            End
            Begin VB.Menu mPixelSpacing 
               Caption         =   "20-Pixel Space"
               Index           =   1
            End
            Begin VB.Menu mPixelSpacing 
               Caption         =   "30-Pixel Space"
               Index           =   2
            End
            Begin VB.Menu mPixelSpacing 
               Caption         =   "60-Pixel Space"
               Index           =   3
            End
         End
         Begin VB.Menu mSpecial 
            Caption         =   "Space (&nbsp;)"
            Index           =   0
         End
         Begin VB.Menu mSpecial 
            Caption         =   "'<'"
            Index           =   1
         End
         Begin VB.Menu mSpecial 
            Caption         =   "'>'"
            Index           =   2
         End
         Begin VB.Menu mHyphen 
            Caption         =   "Hyphen"
            Shortcut        =   ^H
         End
         Begin VB.Menu mMarks 
            Caption         =   "ACBL"
            Index           =   0
            Shortcut        =   +{F9}
         End
         Begin VB.Menu mMarks 
            Caption         =   "Alert"
            Index           =   1
            Shortcut        =   +{F11}
         End
         Begin VB.Menu mMarks 
            Caption         =   "Announce"
            Index           =   2
            Shortcut        =   +{F12}
         End
         Begin VB.Menu mMarks 
            Caption         =   "Spaces3"
            Index           =   4
         End
      End
      Begin VB.Menu zS 
         Caption         =   "-"
      End
      Begin VB.Menu zSS 
         Caption         =   "Suit Symbols"
         Begin VB.Menu mSuitSymbol 
            Caption         =   "Club"
            Index           =   1
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu mSuitSymbol 
            Caption         =   "Diamond"
            Index           =   2
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mSuitSymbol 
            Caption         =   "Heart"
            Index           =   3
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu mSuitSymbol 
            Caption         =   "Spade"
            Index           =   4
            Shortcut        =   ^{F4}
         End
         Begin VB.Menu mSuitSymbol 
            Caption         =   "Club/Diamond"
            Index           =   5
            Shortcut        =   ^{F5}
         End
         Begin VB.Menu mSuitSymbol 
            Caption         =   "Heart/Spade"
            Index           =   6
            Shortcut        =   ^{F6}
         End
         Begin VB.Menu zBStS 
            Caption         =   "-"
         End
         Begin VB.Menu mSuits2Symbols 
            Caption         =   "Change  Suits to Symbols"
         End
         Begin VB.Menu mUndoSuits2Symbols 
            Caption         =   "Undo Suits to Symbols"
         End
         Begin VB.Menu mSymbolsForHands 
            Caption         =   "Make Symbols For Hands"
            Shortcut        =   ^M
         End
         Begin VB.Menu mInsertBidFont 
            Caption         =   "Insert Bidding Font"
            Visible         =   0   'False
         End
         Begin VB.Menu mUndoBidFont 
            Caption         =   "Undo Bidding Font"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu z3b 
         Caption         =   "-"
      End
      Begin VB.Menu mTbl 
         Caption         =   "&Table"
         Begin VB.Menu mTableSpecsWindow 
            Caption         =   "Table &Specs Window"
         End
         Begin VB.Menu mTable 
            Caption         =   "&Table <Table>"
            Shortcut        =   ^T
         End
         Begin VB.Menu mRow 
            Caption         =   "&Row <TR>"
            Shortcut        =   ^Q
         End
         Begin VB.Menu mCell 
            Caption         =   "&Data Cell <TD>"
            Shortcut        =   ^D
         End
         Begin VB.Menu mMergeCellsAcross 
            Caption         =   "Merge Cells Across"
         End
         Begin VB.Menu mMergeCellsDown 
            Caption         =   "Merge Cells Down"
         End
         Begin VB.Menu mAlignTop 
            Caption         =   "&Align Top"
         End
         Begin VB.Menu mWidth 
            Caption         =   "&Width"
         End
      End
      Begin VB.Menu z8 
         Caption         =   "-"
      End
      Begin VB.Menu mTe 
         Caption         =   "Text &Emphasis"
         Begin VB.Menu mFont 
            Caption         =   "&Font"
         End
         Begin VB.Menu mBiddingFont 
            Caption         =   "Bidding Font/Courier"
         End
         Begin VB.Menu mFontSize 
            Caption         =   "Font Size"
         End
         Begin VB.Menu mBold 
            Caption         =   "&Bold"
            Shortcut        =   ^B
         End
         Begin VB.Menu mColorWords 
            Caption         =   "&Color"
         End
         Begin VB.Menu mItalics 
            Caption         =   "&Italics"
            Shortcut        =   ^I
         End
         Begin VB.Menu mUnderline 
            Caption         =   "&Underline"
            Shortcut        =   ^U
         End
         Begin VB.Menu mTextBox 
            Caption         =   "Text-In-A-Box"
         End
         Begin VB.Menu mYellowHighlighter 
            Caption         =   "&Yellow Highlighter"
            Shortcut        =   ^Y
         End
         Begin VB.Menu mJavascript 
            Caption         =   "Javascript Features"
         End
      End
      Begin VB.Menu mZm 
         Caption         =   "Zoom Images"
         Begin VB.Menu mZoom 
            Caption         =   "All Images"
            Index           =   0
            Shortcut        =   ^{F7}
         End
         Begin VB.Menu mZoom 
            Caption         =   "One Image"
            Index           =   1
            Shortcut        =   +{F8}
         End
      End
   End
   Begin VB.Menu mH2 
      Caption         =   "&Help"
      Begin VB.Menu mHelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu mVer 
         Caption         =   "Version 20181023"
      End
   End
End
Attribute VB_Name = "BidBaseFileViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************
'            CardShark BidBase File Viewer             *
'                www.aeyec.com/BidBase                 *
'*******************************************************

Option Explicit

Private Const Version$ = "2022-10-30" ' 2-window display > 1 screen fixed '
Private Const DebugMode = False

Private int_AlreadyPicked  As Integer

Private str_AppPath As String
Private str_DataPath As String
Private str_FiName(1) As String
Private str_FiPath(1) As String
' Form positioning:
Private lng_File1ViewerWidth As Long
Private lng_File2ViewerWidth As Long
Private lng_FileViewerLeft As Long
Private lng_FileViewerWidth As Long
Private lng_FileViewerTop As Long
Private lng_FileViewerHeight As Long
Private int_ActiveMonitor As Integer

Private lng_CursPtr As Long
Private lng_TextLen As Long

Private str_Backup(1) As String ' Backup(0)=original file, (1)=before last save
Private str_CmdLine As String
Private str_FormColors(4) As String
Private str_FormColorCaptions(4) As String
Private qt As String

Private int_IntLinkNum As Integer ' Internal links number. '
Private str_LinkIDs(10) As String ' For going back after an internal link jump. '
Private int_LastLink As Integer
Private bln_NoResize As Boolean
Private bln_OpenInRightWindow As Boolean ' Where to open new files. '
Private str_Orig(1) As String ' Original text; for detecting edit changes. '
Private lng_Panel1Width As Long
Private lng_Panel2Width As Long
Private lng_PanelHeight As Long
Private str_RecentFiles(1, 10) As String ' element 0=path, 1=file name
Private bln_ShowLinksOnRight As Boolean  ' Where to open links. '
Private bln_SkipIt As Boolean
Private int_TxtNum As Integer
Private hideID As Integer
' Text before converting suits '
' to symbols; for Undo:        '
Private str_UndoSuits2Symbols As String
' Text before inserting bidding font: '
Private str_UndoBidFont As String


Private Sub GetColor(HTMLcode As String, ByVal num)
    Dim hx As String
    Dim VBcode As String
    Dim t As String
    Dim i As Integer
    Dim c As String
    
    On Error GoTo ErrHandler
    With CmDia
        .CancelError = True
        .Flags = cdlCCRGBInit Or cdlCCFullOpen
        'HTML colors are RGB,
        'and CmDia's are BGR,
        'so reverse them:
        VBcode = "&H" & _
                    Right$(HTMLcode, 2) & _
                    Mid$(HTMLcode, 4, 2) & _
                    Mid$(HTMLcode, 2, 2)
        .Color = VBcode
        .ShowColor
        hx = Right$("000000" & Hex(.Color), 6)
        HTMLcode = "#" _
                 & Right$(hx$, 2) _
                 & Mid$(hx$, 3, 2) _
                 & Left$(hx$, 2)
    End With
   
    t_FormColorCode(num) = HTMLcode
    str_FormColors(num) = HTMLcode
   
    t = Text2(1).Text
    c = str_FormColorCaptions(num)
    i = InStr(t, c)
    If i > 0 Then
        Mid$(t, _
             i + Len(c) + 2 _
             ) = HTMLcode
        Text2(1).Text = t
        mSave_Click 0
    End If

ExitGetColor:
    On Error GoTo 0
Exit Sub

ErrHandler:
    'User pressed Cancel button.
    Resume ExitGetColor
End Sub


Private Sub brwWebBrowser_BeforeNavigate2( _
                index As Integer, _
                ByVal pDisp As Object, _
                URL As Variant, _
                Flags As Variant, _
                TargetFrameName As Variant, _
                PostData As Variant, _
                Headers As Variant, _
                Cancel As Boolean)
    ' This routine is called every time a  '
    ' new page is displayed in the browser '
    ' but it is only used when a link has  '
    ' been clicked in the left window with '
    ' the intent being to display the link '
    ' in the right window. If such is the  '
    ' actual case, then the the call is    '
    ' Canceled so that the page will not   '
    ' be loaded in the left window, then   '
    ' another call is made to open the same'
    ' file in the right window.            '
    ' If link clicked, Flags = 335.        '
    ' If back arrow clicked, Flags = 272.  '
        
    Dim u As String
    Dim i As Long
    Dim p As String
    Dim F As String
    Dim e As String
    Dim lnk As String
    Dim editHTML As Boolean
    
    If bln_SkipIt Then Exit Sub
    ' When opening links in the right window, '
    ' this sub keeps getting called until out '
    ' of space. '
    
    ' Following are several cases in which '
    ' we do not want to display a file in  '
    ' the right window:                    '
    'Debug.Print Flags, URL
    ' If link is javascript to display a '
    ' picture or file in a pop-up windows:
    If Left$(URL, 4) = "java" Then Exit Sub
'On Local Error Resume Next
    
    ' Some web pages can generate some '
    ' strange stuff here, so ignore it.'
    If InStr(LCase$(URL), ".htm") = 0 And _
       InStr(LCase$(URL), ".txt") = 0 _
    Then
        Exit Sub
    End If
    
    ' If link is a web page or if a web  '
    ' address was entered in the URL line:
    If Left$(URL, 4) = "http" Then
        Exit Sub
    End If
    
    ''''''''''''''''''''''''''''''''''''''''
    
'MsgBox "Ck.Pt. 1"
    ' Clean up URL: '
    u = URL
    u = Replace(u, "%20", " ")
    u = Replace(u, "/", "\")
    If Left$(u, 8) = "file:\\\" Then
        u = Mid$(u, 9)
    End If
        
    ' Extract file name from URL: '
    i = InStrRev(u, "\")
    F = Mid$(u, i + 1) ' Note: may still have Link on end (e.g.: #A).'
    p = Left$(u, i) ' path of local file. '
'MsgBox "ck.pt. 2"
    ' If same path\file, ignore it: '
    If Not Text2(1).Visible And str_FiPath(int_AlreadyPicked) & _
       str_FiName(int_AlreadyPicked) = _
          Left$(u, Len(str_FiPath(int_AlreadyPicked) & _
          str_FiName(int_AlreadyPicked))) _
    Then
        i = InStrRev(URL, ".htm#")
        If i > 0 Then
            str_FiName(int_AlreadyPicked) = _
            str_FiName(int_AlreadyPicked) & Mid$(URL, i + 4)
            FixFileName
            If bln_OpenInRightWindow Then
                str_FiName(1) = str_FiName(0)
                str_FiPath(1) = str_FiPath(0)
                int_TxtNum = 1
                L_Name(1).Visible = False
                Text2(1).Visible = False
                OpenFile p, F
                'Text2(1).Refresh
                bln_SkipIt = True ' Keeps this sub from getting run endlessly. '
                DoEvents
                Cancel = True
            End If
'MsgBox "ck.pt. 3"
            bln_SkipIt = False
            Exit Sub
        End If
    End If
'MsgBox "ck.pt. 4"
    ' If source is in right window and  '
    ' an internal link is clicked, then '
    ' sync the source to browser panel: '
    If Text2(1).Visible Then
        If Flags = 272 Then
            ' Left arrow was clicked to go back. '
            If int_LastLink < 1 Then int_LastLink = 1
            i = InStr(InStr(Text2(1).Text, "</HEAD>"), Text2(1).Text, "#" & str_LinkIDs(int_LastLink))
            If i > 200 Then
                Text2(1).SelStart = i - 200
            Else
                Text2(1).SelStart = i
            End If
            int_LastLink = int_LastLink - 1
            Exit Sub
        
        ElseIf Left$(F, Len(str_FiName(1))) = str_FiName(1) Then
            ' Search back of something like "MyFile.htm#menu8" '
            ' for the "#" where #menu'
            i = InStrRev(F, "#", Len(F) - 1)
            If i > 0 Then
                lnk = Mid$(F, i + 1)
                i = InStr(Text2(1).Text, "NAME=" & Chr$(34) & lnk)
                If i > 0 Then
                    Text2(1).SelStart = i + 2000
                End If
            End If
            L_Name(0).Text = F
            Me.Caption = str_FiName(0) & " - BidBase File Viewer"
            ' This is the caption shown on the Task Bar, '
            ' so we want the file name to come first.    '
            GoTo RecordLink
        Else
            If str_FiName(0) = str_FiName(1) Then editHTML = True
            mEditHTML_Click
            L_Name(0).Text = F
            Me.Caption = str_FiName(0) & " - BidBase File Viewer"
            str_FiPath(0) = p
            str_FiName(0) = F
            FixFileName
            UpdateRecentFiles 0
            'If editHTML Then
            '    MsgBox "Note that the source code window will not change."
            'End If
            mEditHTML_Click
        End If
        GoTo RecordLink
    End If
    
    ' If swapping windows, or if Open '
    ' In Right Window is not checked, '
    ' then we exit. mSwap may also be '
    ' checked as a flag at other times'
    ' when we don't want this routine '
    ' to execute beyond this point.   '
    ' This is not up with the other   '
    ' exit routines because we may    '
    ' want the above sync routine to  '
    ' execute.                        '
'MsgBox "ck.pt. 4"
    If mSwap.Checked Or _
       bln_OpenInRightWindow = False _
    Then
        Exit Sub
    End If
        
    ' See if loading the same file: '
    If Mid$(Headers, 9) = u Then
        Exit Sub
    End If
          
    If index = 1 Then
        ' Didn't click in left window. '
        int_TxtNum = index
        str_FiPath(int_TxtNum) = p
        str_FiName(int_TxtNum) = F
        FixFileName
        UpdateRecentFiles int_TxtNum
        ' browser will continue to navigate '
    
    Else
        ' Stop the original link click and '
        ' open the file in the right window.
        int_TxtNum = 1
        L_Name(1).Visible = False
        Text2(1).Visible = False
'MsgBox "ck.pt. 5"
        OpenFile p, F
        Cancel = True
    End If
    
RecordLink:
    If lnk = "" Then Exit Sub
    If int_LastLink = 10 Then
        For i = 1 To 9
            str_LinkIDs(i) = str_LinkIDs(i + 1)
        Next
    Else
        int_LastLink = int_LastLink + 1
    End If
    str_LinkIDs(int_LastLink) = lnk
End Sub


Private Sub brwWebBrowser_DocumentComplete( _
                index As Integer, _
                ByVal pDisp As Object, _
                URL As Variant)
    ' After the page has loaded,  '
    ' save the path and file name.'
    
    Dim F As String
    Dim i As Integer
    Dim u As String
    Dim n As Long
    
    ' Turn off the swapping flag, if on: '
    mSwap.Checked = False
    
On Local Error Resume Next
    If Left$(URL & "    ", 4) = "http" Then
        ' A page from the web is shown. '
        ' Save the URL in the File Name '
        ' variable:                     '
        str_FiName(index) = brwWebBrowser(index).LocationURL
        'If index > 0 Then MsgBox "DocumentComplete #1"
        FixFileName
        ' Don't need anything in the Path: '
        str_FiPath(index) = ""
    
    ElseIf Trim$(brwWebBrowser(index).LocationName) <> "http:///" Then
        ' Must be a local page: '
        F = brwWebBrowser(index).LocationName
        'f = str_FiName(Index) This was changing the line above. Don't know why. '
        If F = "" Then Exit Sub
        
        u = brwWebBrowser(index).LocationURL
        ' Clean up URL: '
        u = Replace(u, "%20", " ")
        u = Replace(u, "/", "\")
        If Left$(u, 8) <> "http:\\\" Then
            If Left$(u, 8) = "file:\\\" Then
                u = Mid$(u, 9)
            End If
        End If
        u = Replace(u, "%23", "#")
        i = InStr(u, F)
        If i > 0 Then
            str_FiPath(index) = Left$(u, i - 2) & "\"
        End If
        str_FiName(index) = F
        L_Name(index) = F
        If 1 = 2 And Text2(1).Visible Then
        ' I'm pretty sure this is always being done in BeforeNavigate2 '
            n = InStrRev(u, "#")
            If n > 0 Then
                u = "<A NAME=" & qt & Mid$(F, n + 1)
                n = InStr(Text2(1).Text, u)
                If n > 0 Then
                    Text2(1).SelStart = Len(Text2(1).Text)
                    Text2(1).SelStart = n
                End If
            Else
                Text2(1).SelStart = 1
            End If
        End If
        FixFileName
    End If
    ' Update the Recent Files list:
    UpdateRecentFiles 0
End Sub


Private Sub brwWebBrowser_NavigateComplete2(index As Integer, ByVal pDisp As Object, URL As Variant)
'Stop
End Sub

Private Sub brwWebBrowser_NewWindow2(index As Integer, ppDisp As Object, Cancel As Boolean)
    'Stop - This was supposed to allow a user to right-click on a link to open it in a new window - presumbably in Chrome, etc.?
End Sub

Private Sub brwWebBrowser_OnMenuBar(index As Integer, ByVal MenuBar As Boolean)
    'Not using this.
End Sub

Private Sub cmdAddWord_Click()
    ' Add a word to my dictionary. '
    'AddWord t_Word
    'If bln_CheckWord Then
    '    MsgBox qt & t_Word & qt & " has been added."
    '    CloseSpellCheck
    'Else
    '    SkipWord 0
    'End If
End Sub


Private Sub brwWebBrowser_GotFocus(index As Integer)
    ' Flag for user clicking in this browser. '
    int_AlreadyPicked = index
End Sub


Public Sub BringOutSuggestionBox()
    ' For spell checker '
    
    Dim i As Integer
    Dim j As Integer
    Dim stp As Integer
    
    lstSuggestions.Clear
    t_Replace = ""
    t_Word = ""
    
    With frmSuggestions
        .Left = Text2(1).Left
        .Top = Text2(1).Top
        .Visible = True
        stp = .Width / 100
        For i = 1 To 105
            .Left = .Left - stp
            .Refresh
            If i / 5 = i \ 5 Then SlowItDown 5
        Next
    End With
End Sub


Public Sub cmdCloseSpellCk_Click()
    'bln_CheckWord = True 'To stop "completed" message
    'CloseSpellCheck
    'Make sure highlighted text is unhighlighted:
    'If HighlightLen > 0 Then
    '    Text2(1).SelStart = HighlightStart
    '    Text2(1).SelLength = HighlightLen
    '    Text2(1).SelText = vbBlack
    '    HighlightStart = 0
    '    HighlightLen = 0
    'End If
End Sub


Public Sub HideSuggestionBox()
    Dim stp As Integer
    Dim i As Integer
    
    With frmSuggestions
        Text2(1).ZOrder 0
        stp = .Width / 100
        For i = 1 To 105
            .Left = .Left + stp
            If i / 5 = i \ 5 Then SlowItDown 5
        Next
        .Visible = False
    End With
End Sub


Private Sub cmdFixSpelling_Click()
    'ReplaceWord
End Sub


Private Sub cmdGetColor_Click(index As Integer)
    Dim stp As Integer
    Dim i As Integer
    Dim j As Integer
    Dim t As String
    Dim c As String
    
    If index = 6 Then
        ' Reset starting colors: '
        If t_FormColorCode(i).Tag = "" Then
            With Text2(1)
                i = InStr(.Text, "TEXT=")
                If i = 0 Then
                    For j = 0 To 4
                        t_FormColorCode(j) = str_FormColors(j)
                        t_FormColorCode(j).Tag = str_FormColors(j)
                    Next
                    Exit Sub
                End If    ' 12345678      '
                i = i + 6 ' TEXT="#000000 '
                t_FormColorCode(0) = Mid$(.Text, i, 7)
                
                i = InStr(i + 7, .Text, "BGCOLOR=")
                If i = 0 Then
                    t_FormColorCode(1) = str_FormColors(1)
                Else
                    i = i + 9
                    t_FormColorCode(1) = Mid$(.Text, i, 7)
                End If
                    
                i = InStr(i + 6, .Text, "LINK=")
                If i = 0 Then
                    t_FormColorCode(2) = str_FormColors(2)
                Else
                    i = i + 6
                    t_FormColorCode(2) = Mid$(.Text, i, 7)
                End If
                    
                i = InStr(i + 7, .Text, "VLINK=")
                If i = 0 Then
                    t_FormColorCode(3) = str_FormColors(3)
                Else
                    i = i + 7
                    t_FormColorCode(3) = Mid$(.Text, i, 7)
                End If
                    
                i = InStr(i + 7, .Text, "VLINK=")
                If i = 0 Or i > 1000 Then
                    t_FormColorCode(4) = str_FormColors(4)
                Else
                    i = i + 7
                    t_FormColorCode(4) = Mid$(.Text, i, 7)
                End If
            End With
        Else
            For i = 0 To 4
                t_FormColorCode(i) = t_FormColorCode(i).Tag
            Next
        End If
                    
        If frmColors.Visible Then
            t = Text2(1).Text
            For i = 0 To 4
                t_FormColorCode(i).Tag = t_FormColorCode(i)
                ' Put HTML back in source code: '
                c = str_FormColorCaptions(i)
                j = InStr(t, c)
                If j > 0 Then
                    Mid$(t, j + Len(c) + 2) = t_FormColorCode(i)
                End If
            Next
            Text2(1).Text = t
            mSave_Click 0
            t = Empty
        End If
        
    ElseIf index = 5 Then
        With frmColors
            For i = 0 To 4
                t_FormColorCode(i).Tag = t_FormColorCode(i)
            Next
            .Visible = False
        End With
    Else
        GetColor str_FormColors(index), index
    End If
End Sub


Private Sub cmdSkipWord_Click(index As Integer)
    'SkipWord index
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


Private Sub Form_Load()
    'Set AICWordsWS = DBEngine.Workspaces(0)
    'Set AICWordsDB = AICWordsWS.OpenDatabase(str_AppPath & "ai.accdb")
    'Set WordsRS = AICWordsDB.OpenRecordset("Words")
    Dim i As Integer
    
    bln_NoResize = True ' Don't want the form resizing before setup is done. '
'Me.Show
'Me.Left = ScaleX(-1920, vbPixels, vbTwips)

    Me.Caption = Me.Caption & "      ver." & Version$
    Me.Move 0, 0

    'If RunningInIDE Then
    '    str_AppPath = "C:\Programming\VB98\Bridge\"
    'Else ' Running the exe. '
    '    str_AppPath = "C:\Program Files (x86)\CardShark BidBase\"
    'End If ' This was causing problems. '
    
    str_AppPath = App.Path & "\"
    str_DataPath = str_AppPath
    ChDir str_DataPath
    
    ' Other BB programs can call Viewer '
    ' & pass a file name to be viewed:  '
    Dim c As String
    c = Command()
    If c <> "" Then
        L_Name(0) = c
        GetCommandLine
    End If
    
    GetMonSpecs ' Fills these variables:  '
    ' lng_NumMonitors is # of monitors.   '
    ' lng_MonPosition(4,#) is left, top,  '
    '   right, bottom for monitor # 1-... '
    ' str_MonitorID(#) is handle for # 1- '
    ' int_PrimaryMon is not same as Win10.'
    
    ' Convert Pixels to Twips: '
    ' MonitorInfo gives results in pixels but VB6 wants them in twips. '
    For i = 1 To 12
        If lng_MonPosition(4, i) = 0 Then Exit For ' Fewer than 12 monitors. ' Overkill '
        lng_MonPosition(1, i) = Me.ScaleX(lng_MonPosition(1, i), vbPixels, vbTwips)
        lng_MonPosition(2, i) = Me.ScaleX(lng_MonPosition(2, i), vbPixels, vbTwips)
        lng_MonPosition(3, i) = Me.ScaleX(lng_MonPosition(3, i), vbPixels, vbTwips)
        lng_MonPosition(4, i) = Me.ScaleX(lng_MonPosition(4, i), vbPixels, vbTwips)
    Next
    ' Virtual screen = all monitors combined. '
    ' I've commented these lines out because  '
    ' they are not needed right now. '
    'lng_VirtualScrLeft = _
    '    Me.ScaleX(GetSystemMetrics(SM_XVIRTUALSCREEN), vbPixels, vbTwips)
    'lng_VirtualScrTop = _
    '    Me.ScaleY(GetSystemMetrics(SM_YVIRTUALSCREEN), vbPixels, vbTwips)
    'lng_VirtualScrHeight = _
    '    Me.ScaleY(GetSystemMetrics(SM_CYVIRTUALSCREEN), vbPixels, vbTwips)
    'lng_VirtualScrWidth = _
    '    Me.ScaleX(GetSystemMetrics(SM_CXVIRTUALSCREEN), vbPixels, vbTwips)
    'GetDesktopMaximumHeight ' Sets lng_MaxMonitorHeight 's
    'lng_MaxMonitorHeight = Me.ScaleY(lng_MaxMonitorHeight, vbPixels, vbTwips)
    
    GetActiveMonitor
    Me.Height = 16095 ' lng_MonPosition(4, int_ActiveMonitor) ' 4 is .Bottom '
    
    LoadFiles
    LoadSetUp
    bln_NoResize = False
    Form_Resize
    
    If Text2(1).Visible Then
        int_TxtNum = 1
        Text2(1).SetFocus
        Text2(1).SelStart = Len(Text2(1).Text)
    Else
        int_TxtNum = 0
    End If
End Sub


Private Sub GetActiveMonitor()
    If Me.Left < -200 Then
        int_ActiveMonitor = 1 ' Not the same as Windows 10 monitor numbers  '
    ElseIf Me.Left < (lng_MonPosition(3, 2) - 200) Then ' On my system, the left monitor is 3, not 1. '
        int_ActiveMonitor = 2
    Else
        int_ActiveMonitor = 3
    End If
End Sub


Private Sub SetMenuStatus(ByVal Status As Boolean)
    mH.Enabled = Status
    mSpellCheckWord.Enabled = Status
    mSpellCheckDocument.Enabled = Status
    mSuits2Symbols.Enabled = Status
    mUndoSuits2Symbols.Enabled = Status
    mFind(0).Enabled = Status
End Sub


Private Sub LoadHTML()
    Dim ret As Long
    Dim a As String
    Dim i As Integer
  
    ' If loading the source code for an HTML page, it '
    ' is loaded into the Rich Text Editor, Text2(1).  '
  
    If Text2(1).Visible And _
       str_Orig(1) <> Text2(1).Text And _
       L_Name(1) <> "" _
    Then
        ret = MsgBox("Save changes to " & _
                     L_Name(1).Text & "?", _
                     vbYesNoCancel)
        If ret = vbCancel Then
            Exit Sub
    
        ElseIf ret = vbYes Then
            mSave_Click 0
        End If
    End If
  
    bln_NoResize = True
  
    If str_FiName(0) = "" Then
        str_FiName(0) = L_Name(0)
        str_FiName(1) = str_FiName(0)
    End If
  
    'Load file:
    If Left$(str_FiPath(1), 6) = "res:\\" Then
        'I don't know where this comes from '
        'or why, but it has turned up here, '
        'so I'm fixing it here: '
        str_FiPath(0) = str_AppPath
        str_FiPath(1) = str_AppPath
    End If
  
    If Left$(str_FiName(1), 4) = "http" Then
        MsgBox "To edit a Web page's source HTML" & vbCr _
             & "code you must click File - Save Web Page" & vbCr _
             & "to save the source HTML code to disk." & vbCr _
             & "The source code will then be saved to" & vbCr _
             & "disk and displayed in the right window." & vbCr & "Some Web features will not work here."
        Exit Sub
    End If
  
    On Local Error GoTo CantLoadHtml
  
    If Dir$(str_FiPath(1) & str_FiName(1)) = "" Then
        If Dir$(str_FiPath(1) & str_FiName(1) & ".htm") <> "" Then
            str_FiName(1) = str_FiName(1) & ".htm"
        ElseIf Dir$(str_FiPath(1) & str_FiName(1) & ".html") <> "" Then
            str_FiName(1) = str_FiName(1) & ".html"
        ElseIf Dir$(str_FiPath(1) & str_FiName(1) & ".txt") <> "" Then
            str_FiName(1) = str_FiName(1) & ".txt"
        End If
    End If
    
    If Dir$(str_FiPath(1) & str_FiName(1)) = "" Then
        a = InputBox$("The following source file path or name is invalid:" & vbCr & _
               str_FiPath(1) & str_FiName(1) & vbCr & _
               "Enter correction:", , str_FiPath(1) & str_FiName(1))
        If a <> "" And a <> str_FiPath(1) & str_FiName(1) Then
            i = InStrRev(a, "\")
            If i = 0 Then i = InStrRev(a, "/")
            If i > 0 Then
                str_FiPath(1) = Left$(a, i)
                str_FiName(1) = Mid$(a, i + 1)
                LoadHTML
            End If
        End If
        On Local Error GoTo 0
        Exit Sub
    ElseIf Dir$(str_FiPath(1) & str_FiName(1)) <> "" Then
        Text2(1).LoadFile str_FiPath(1) & str_FiName(1)
    End If
    
    ' Hide the right browser, show the right text editor:
    brwWebBrowser(1).Visible = False
    Text2(1).Visible = True
    L_Name(1) = str_FiName(1)
    L_Name(1).Visible = True
  
    Dim curScrWidth As Long
    Dim curScrLeft As Long
  
    ' If the 2-viewer Window width is less than .7 '
    ' of the screen width, then make it the screen '
    ' width:                                       '
    lng_File2ViewerWidth = lng_MonPosition(3, int_ActiveMonitor) - lng_MonPosition(1, int_ActiveMonitor)
    
    ' Do initial panel setup. Resize '
    ' only does changes once set up. '
    If lng_Panel1Width > Me.Width * 0.6 Or _
       lng_Panel2Width < Me.Width * 0.4 _
    Then
        lng_Panel1Width = Me.Width * 0.6
        lng_Panel2Width = Me.Width * 0.4
    End If
    
    SetMenuStatus True
    bln_NoResize = False
    Form_Resize
    str_Orig(1) = Text2(1).Text
    SetMenuStatus True
    bln_NoResize = True
Exit Sub

CantLoadHtml:
    On Local Error GoTo 0
    MsgBox "Error loading " & str_FiPath(1) & str_FiName(1)
    bln_NoResize = True
End Sub

Private Sub Form_Resize()
    '                                                      '
    ' This routine is used ONLY for adjusting the size of  '
    ' everything IN the Viewer window when the user drags  '
    ' the edges of the form or when the form size changes  '
    ' because of switching monitors or changing the number '
    ' of panels shown. All variables which track object &  '
    ' form sizes after the form size has changed are reset '
    ' in this routine.                                     '
    '                                                      '
    ' Other routines simply change the size of the form,   '
    ' then call this routine to resize/move components.    '
    '                                                      '
    ' This routine does NOT change the values of Me.Left,  '
    ' Me.Top, Me.Width or Me.Height.                       '
    '                                                      '
    ' This routine does NOT change any non-sizing or non-  '
    ' positioning variables/properties (eg: L_Name.Visible).
    '                                                      '
    Dim chng As Long
    Dim chng1 As Long
    Dim chng2 As Long
    Dim oldCombinedPanelWidth As Long
    Dim newCombinedPanelWidth As Long
    Dim c As String
    Dim i As Integer
    Dim j As Integer
    
    If bln_NoResize Then
        Exit Sub
    End If
    
    If Me.WindowState = 1 Then
        ' Just show the file name in the Windows Task Bar. '
        Me.Caption = L_Name(0)
        Exit Sub
    ElseIf InStr(Me.Caption, "ver.") = 0 Then
        Me.Caption = str_FiName(0) & " - BidBase File Viewer  ver.:" & Version
    End If
    
    bln_NoResize = True
    On Local Error Resume Next
    Me.WindowState = vbNormal ' Full-screen lock not allowed. It messes up sizing. '
    
    ' Keep track of form measurements, etc., in twips: '
    lng_FileViewerLeft = Me.Left
    lng_FileViewerTop = Me.Top
    lng_FileViewerHeight = Me.Height
    lng_PanelHeight = lng_FileViewerHeight - 1500 ' Work area below the menus, etc. '
    
    If L_Name(1) <> "" Then
        ' Right panel is used. Give it room: '
        If lng_FileViewerWidth = Me.Width Then
            lng_FileViewerWidth = Me.Width / 2
        End If
    End If
    chng = Me.Width - lng_FileViewerWidth
    lng_FileViewerWidth = Me.Width
    ' If 2nd panel is open, changes its size: '
    If Text2(1).Visible Or _
       brwWebBrowser(1).Visible _
    Then ' 2 panels: '
        ' Variable for form width when 2 panels show: '
        lng_File2ViewerWidth = lng_FileViewerWidth
        ' Calculate panel widths :
        If lng_Panel1Width < lng_FileViewerWidth * 0.3 Or _
           lng_Panel2Width < lng_FileViewerWidth * 0.4 _
        Then
            lng_Panel1Width = lng_FileViewerWidth * 0.5 - 200
            lng_Panel2Width = lng_Panel1Width
        End If
      
        ' Leave 1st panel's width as-is if possible: '
        brwWebBrowser(0).Height = lng_PanelHeight
        Text2(0).Height = brwWebBrowser(0).Height
        Text2(0).Width = brwWebBrowser(0).Width
        If InStr(str_CmdLine, ", ") > 0 Then
            ' Command line used, if comma, then 2 files. '
            brwWebBrowser(0).Width = Me.Width \ 2
            ' Now that this is done, clear it so '
            ' that it doesn't keep doing this. '
            str_CmdLine = ""
        End If
        
        Text2(0).Width = brwWebBrowser(0).Width
      
        ' 2nd panel: '
        Text2(1).Left = brwWebBrowser(0).Width + 315 'left margin + center margin
        Text2(1).Width = Me.Width - Text2(1).Left - 315
        Text2(1).Height = Text2(0).Height
        brwWebBrowser(1).Width = Text2(1).Width
        brwWebBrowser(1).Left = Text2(1).Left
        brwWebBrowser(1).Height = Text2(1).Height
        L_Name(1).Left = Text2(1).Left + 50
        L_Name(1).Width = lng_Panel2Width
        pSplitterBar.Left = brwWebBrowser(0).Width + brwWebBrowser(0).Left + 15
        pSplitterBar.Height = Text2(0).Height
        pSplitterBar.Width = Text2(1).Left - pSplitterBar.Left
        pSplitterBar.Visible = True
    
    Else ' 1 panel: '
        'Note: lng_Panel1Width is ONLY '
        ' changed for 2-panel display. '
        lng_Panel1Width = lng_FileViewerWidth - 200
        brwWebBrowser(0).Top = 500
        lng_File1ViewerWidth = lng_FileViewerWidth ' Notice the "1" in first var. '
        Text2(0).Width = Me.Width - 500
        Text2(0).Height = lng_PanelHeight
        brwWebBrowser(0).Left = 100
        brwWebBrowser(0).Width = Text2(0).Width
        brwWebBrowser(0).Height = Text2(0).Height
        pSplitterBar.Visible = False
        L_Name(1).Visible = False
    End If
    
    lng_Panel1Width = brwWebBrowser(0).Width
    
    t_Back.Left = brwWebBrowser(0).Left _
                   + brwWebBrowser(0).Width _
                   - t_Back.Width + 5
    t_Back.Top = brwWebBrowser(0).Top - t_Back.Height
    
    L_Name(0).Width = lng_Panel1Width _
                    - t_Back.Width _
                    - L_Name(0).Left + 80
    L_Name(0).Left = brwWebBrowser(0).Left + 50
    L_Name(0).Top = brwWebBrowser(0).Top - L_Name(0).Height + 50
    
    bln_NoResize = False
    On Local Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim fiNum As Integer
    Dim i As Integer
    
    CheckForSave Cancel
    If Cancel Then Exit Sub
If DebugMode Then MsgBox "Unload 1"
    ' Store configuration data: '
    fiNum = FreeFile
    On Local Error GoTo 0 ' BadPath
    
    Open str_AppPath & "BidBaseFileViewer.cfg" For Output As #fiNum
    On Local Error GoTo 0
        Print #fiNum, Me.Left
        Print #fiNum, Me.Top
        Print #fiNum, Me.Height
        Print #fiNum, Me.Width
        
        If str_FiName(1) <> "" Then
            ' 2nd window is shown. Save variable '
            ' for single-window form width and   '
            ' the actual window width for 2 wins.'
            Print #fiNum, lng_File1ViewerWidth
            Print #fiNum, Me.Width
        Else ' Vice-versa: '
            Print #fiNum, Me.Width
            Print #fiNum, lng_File2ViewerWidth
        End If
        
        Print #fiNum, lng_Panel1Width
        Print #fiNum, lng_Panel2Width
        Print #fiNum, lng_PanelHeight
        ' Recent files (with paths first): '
        For i = 1 To 10
            Print #fiNum, str_RecentFiles(0, i)
            Print #fiNum, str_RecentFiles(1, i)
        Next
        Print #fiNum, mOpenRight.Checked
    Close #fiNum
Exit Sub

BadPath:
MsgBox "Unload2"
    str_AppPath = InputBox$("Bad path: " & str_AppPath & "BidBaseFileViewer.cfg" & Chr$(13) & "Enter another path:")
    If str_AppPath <> "" Then Resume
End Sub


Private Sub CheckForSave(Cancel)
    ' Check if changes to HTML saved: '
  
    Dim i As Integer
    Dim ret As Long
If DebugMode Then MsgBox "dCheckForSave1"
    For i = 0 To 1
        If Text2(i).Visible And _
           str_FiName(i) <> "" _
        Then
            If Text2(i).Text <> str_Orig(i) Then
                ret = MsgBox("Save changes to " & _
                              L_Name(i) & "?", _
                              vbYesNoCancel)
                If ret = vbYes Then
                    int_TxtNum = i
                    mSave_Click 0
        
                ElseIf ret = vbCancel Then
                    Cancel = True
                    Exit Sub
                End If
            End If
        End If
    Next
End Sub


Private Sub L_Name_KeyDown( _
                index As Integer, _
                KeyCode As Integer, _
                Shift As Integer)
    
    ' User can enter a URL and press '
    ' Enter to load it into browser: '
    If KeyCode = 13 Then ' ascii-13 is Enter '
        If L_Name(index) <> str_FiName(index) Then
            str_FiName(index) = L_Name(index)
            ' If just "www", add "http:.//" - '
            If Left$(str_FiName(index), 3) = "www" Then
                str_FiName(index) = "http://" & str_FiName(index)
                L_Name(index) = str_FiName(index)
            End If
            
            If Left$(str_FiName(index), 4) = "http" Then
                str_FiPath(index) = ""
                brwWebBrowser(0).Navigate str_FiName(index)
                UpdateRecentFiles 0
            End If
        End If
    End If
End Sub


Private Sub LinkMsgBox_Click()
    int_IntLinkNum = int_IntLinkNum - 1
    LinkMsgBox.Visible = False
    Text2(1).SetFocus
End Sub


Private Sub lstSuggestions_Click()
    t_Replace = lstSuggestions
End Sub


Private Sub lstSuggestions_DblClick()
    cmdFixSpelling_Click
End Sub


Private Sub mFontSize_Click()
    With Text2(int_TxtNum)
        .SelText = "<FONT SIZE=+1>" & .SelText & "</FONT>"
    End With
End Sub

Private Sub mHalfLineAbove_Click()
    With Text2(int_TxtNum)
        .SelText = "<SUP>&nbsp;</SUP>"
        Clipboard.Clear
        Clipboard.SetText "<SUP>&nbsp;</SUP>"
    End With
End Sub

Private Sub mHalfLineBelow_Click()
    With Text2(int_TxtNum)
        .SelText = "<SUB>&nbsp;</SUB>"
        Clipboard.Clear
        Clipboard.SetText "<SUB>&nbsp;</SUB>"
    End With
End Sub

Private Sub mHalfLineAll_Click()
    Dim x As Integer
    With Text2(int_TxtNum)
        x = MsgBox("This will set all List lines in the file" & Chr$(13) & "with half-line breaks between them." & Chr$(13) & "Continue?", vbYesNo)
        If x = vbNo Then Exit Sub
        x = InStr(LCase(Left$(.Text, 300)), "li (margin-top:")
        If x > 0 Then
            MsgBox "This option has already been set."
            Exit Sub
        End If
        x = InStr(LCase(Left$(.Text, 300)), "</head>")
        If x = 0 Then
            MsgBox "This file does not have an HTML header."
            Exit Sub
        End If
        Text2(int_TxtNum) = Left$(.Text, x - 1) & Chr$(13) & _
                            "<style>" & Chr$(13) & _
                            "li {margin-top: 4px;" & Chr$(13) & _
                            "margin-bottom: 0px;}" & Chr$(13) & _
                            "</style>" & Chr$(13) & _
                            Mid$(.Text, x)
    End With
End Sub


Private Sub mHyphen_Click()
    With Text2(int_TxtNum)
        .SelText = "–"
    End With
End Sub


Private Sub mMarks_Click(index As Integer)
    With Text2(int_TxtNum)
        Clipboard.Clear
        If index = 0 Then
            .SelText = "<span style=""background-color: RED""><B>!!</B></SPAN>"
            Clipboard.SetText "<span style=""background-color: RED""><B>!!</B></SPAN>"
        ElseIf index = 1 Then
            .SelText = "<span style=""color: WHITE; background-COLOR: RED"">!!</span>"
            Clipboard.SetText "<span style=""color: WHITE; background-COLOR: RED"">!!</span>"
        ElseIf index = 2 Then
            .SelText = "<span style=""color: WHITE; background-COLOR: BLUE"">a</span>"
            Clipboard.SetText "<span style=""color: WHITE; background-COLOR: BLUE"">a</span>"
        Else
            .SelText = "&nbsp;&nbsp;&nbsp;"
            Clipboard.SetText "&nbsp;&nbsp;&nbsp;"
        End If
    End With
End Sub


Private Sub mNoBreak_Click()
    With Text2(int_TxtNum)
        .SelText = "<span style=" & qt & "white-space: nowrap" & qt & ">" & .SelText & "</SPAN>"
    End With
End Sub


Private Sub mPixelSpacer_Click()
    With Text2(int_TxtNum)
        .SelText = "<IMG SRC=""1x30 line height.png"">" & .SelText
    End With
End Sub


Private Sub mPixelSpacing_Click(index As Integer)
    With Text2(int_TxtNum)
        If index = 0 Then
            .SelText = "<IMG SRC=""2x4 space.png"">"
        ElseIf index = 1 Then
            .SelText = "<IMG SRC=""20x4 space.png"">"
        ElseIf index = 2 Then
            .SelText = "<IMG SRC=""30x4 space.png"">"
        ElseIf index = 3 Then
            .SelText = "<IMG SRC=""60x4 space.png"">"
        End If
    End With
End Sub



Private Sub mTextBox_Click()
    Dim t As String
    Dim txt As String
    
    txt = Text2(int_TxtNum).SelText
    If txt = "" Then txt = "Your text goes here."
    
    t = "<SPAN style=" & qt & "border:3px; border-style:solid; border-color:GRAY; padding: 1em; width:440px;" & qt & ">"
    Text2(int_TxtNum).SelText = t & "</SPAN>"
Exit Sub
    
    ' This was done with nested tables. '
    ' The code above is much simpler.   '
    t = "<P><Table Border=" & qt & "2" & qt & _
        " Width=" & qt & "800" & qt & ">" & vbCr & _
        "<TR>" & vbCr & _
        "<TD ALIGN=CENTER VALIGN=CENTER>" & vbCr & _
        "<TABLE Width=" & qt & "800" & qt & ">" & vbCr & _
        "<TR>" & vbCr & _
        "<TD ALIGN=CENTER VALIGN=CENTER><BR>" & vbCr & _
        txt & vbCr & _
        "<BR><BR>" & vbCr & _
        "</TD>" & vbCr & _
        "</TR>" & vbCr & _
        "</TABLE>" & _
        "</TD>" & vbCr & _
        "</TR>" & vbCr & _
        "</TABLE>" & vbCr
        
    ' A borderless table with the text is nested inside a '
    ' table with a border to provide a margin all around. '
    ' I spent a lot of time trying other ways and this way'
    ' worked the best, so note to me: Don't waste more time on it. '
    
    With Text2(int_TxtNum)
        .SelText = vbCr & t & vbCr
    End With
End Sub

Private Sub mAlignTop_Click()
    With Text2(int_TxtNum)
        .SelText = " VAlign=Top" & .SelText
    End With
End Sub

Private Sub mBiddingFont_Click()
    With Text2(int_TxtNum)
        .SelText = "<FONT FACE=" & _
                      qt & "Courier New" & qt _
                 & ">" & .SelText _
                 & "</FONT>"
    End With
    mSave_Click 0
End Sub


Private Sub mBold_Click()
    Dim ss As Long
    Dim sl As Long
    
    With Text2(int_TxtNum)
        ss = .SelStart
        sl = .SelLength
        If Right$(.SelText, 1) = " " Then
            ' Double-clicking a word causes a blank space '
            ' next to it to be highlighted too. (Thanks,  '
            ' Microsoft.)                                 '
            .SelText = "<B>" & Trim$(.SelText) & "</B> "
        Else
            .SelText = "<B>" & .SelText & "</B>"
        End If
        .SelStart = ss + 3
        .SelLength = sl
    End With
End Sub


Private Sub mBreak_Click()
    With Text2(int_TxtNum)
        .SelText = "<BR>" & .SelText
    End With
End Sub


Private Sub mBulletsNoIndent_Click()
    Text2(int_TxtNum).SelText = "<DL>" & vbCrLf & _
        "(Any text you want before the list.)" & _
        vbCrLf & "<LI>...</LI>" & vbCrLf & _
        "<LI>...</LI>" & vbCrLf & "</DL>"
End Sub


Private Sub mCenter_Click()
    With Text2(int_TxtNum)
        .SelText = "<CENTER>" & Trim$(.SelText) & "</CENTER> "
    End With
End Sub


Private Sub mCell_Click()
    With Text2(int_TxtNum)
        .SelText = "<TD VAlign=Top>" & .SelText & "</TD>"
    End With
End Sub

Private Sub mChangeMonitors_Click()
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '   <---Virtual Screen Width--->                                   '
  '   -------------- -------------  <-|                              '
  '   |            | |           |    |                              '
  '   |     1      | |    2      |    |Virtual                       '
  '   |            | |           |    |Screen                        '
  '   |            | -------------    |Height                        '
  '   --------------                <-|                              '
  '   <-Scr1Width->  <-Scr2Width->                                   '
  '      3840x800      1920x1080                                     '
  '         4K                      5K = 5120x2880                   '
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
    Dim pctWidth As Single
    Dim pctHeight As Single
    
    ' Monitor numbers aren't necessarily the same  '
    ' as on the user's computer. It doesn't matter '
    ' because we are just rotating the window.     '
    ' Here 1=left, 2=center, 3=right. '
    
    bln_NoResize = True
    
    If int_ActiveMonitor = 1 Then
        ' Move program to monitor 2: '
        int_ActiveMonitor = 2
        Me.Left = 0
    ElseIf int_ActiveMonitor = 2 Then
        int_ActiveMonitor = 3
        Me.Left = lng_MonPosition(1, 3)
    Else
        int_ActiveMonitor = 1
        Me.Left = lng_MonPosition(1, 1)
    End If
    
    bln_NoResize = False
    Form_Resize
    If Dir$("\My Files\*.*") <> "" And Me.Left = 0 Then
        Me.Left = -120
    End If
End Sub


Private Sub mCloseRight_Click()
    ' Close the right window: '
    
    If Text2(1).Visible Or _
       brwWebBrowser(1).Visible _
    Then
        mEditHTML.Checked = False
        ' Save last width for 2 file viewers. '
        lng_File2ViewerWidth = Me.Width
        ' Clear right-window variables: '
        str_FiPath(1) = ""
        str_FiName(1) = ""
        brwWebBrowser(1).Visible = False
        Text2(1).Visible = False
        L_Name(1) = ""
        L_Name(1).Visible = False
        ' Resize form to 1-window width: '
        Me.Width = brwWebBrowser(0).Width + 500
        lng_File1ViewerWidth = Me.Width
        ' Gray-out the HTML menus: '
        SetMenuStatus False
    End If
End Sub


Private Sub mColors_Click(index As Integer)
    ' Set colors in HTML. '
    
    Dim i As Integer
    Dim bk As String
    Dim txt As String
    Dim stp As Integer
    Dim HTMLcodeBack As String
    Dim HTMLcodeText As String
    
    If str_FormColorCaptions(0) = "" Then
        str_FormColorCaptions(0) = "TEXT"
        str_FormColorCaptions(1) = "BGCOLOR"
        str_FormColorCaptions(2) = "LINK"
        str_FormColorCaptions(3) = "VLINK"
        str_FormColorCaptions(4) = "ALINK"
    End If
    
    Select Case index
        Case 0
            cmdGetColor_Click 6 ' Put color codes in boxes. '
            frmColors.Left = pSplitterBar.Left + 3000
            frmColors.Visible = True
            Exit Sub
   
        Case 1
            ' "Text Colors" is the same as the '
            ' "Color" in "Text Emphasis" menu, '
            ' but is included in the "Colors"  '
            ' menu for convenience of the user.'
            mColorWords_Click
            Exit Sub
            
        Case 2
            ' Background and Text Colors '
            With CmDia
                MsgBox "The first color picker box is for background." & vbCr & "The next will be for text color."
                .CancelError = True
                .Flags = cdlCCFullOpen
                .DialogTitle = "Background"
                .ShowColor
                bk = Right$("000000" & Hex(.Color), 6)
            
                .CancelError = True
                .Flags = cdlCCFullOpen
                .DialogTitle = "Text Color"
                .ShowColor
                txt = Right$("000000" & Hex(.Color), 6)
            End With
            
            HTMLcodeBack = "#" & _
                        Right$(bk, 2) & _
                          Mid$(bk, 3, 2) & _
                         Left$(bk, 2)
            
            HTMLcodeText = "#" & _
                        Right$(txt, 2) & _
                          Mid$(txt, 3, 2) & _
                         Left$(txt, 2)
            
            Text2(1).SelText = _
                      "<SPAN style=" & qt & "color: " & HTMLcodeText & "; " & _
                      "background-color: " & HTMLcodeBack & qt & ">" & _
                      Text2(1).SelText & "</SPAN>"
            'Example:  <span style="color: #ffffff; background-color: #ff0000">
            GoTo ExitColors
    End Select
   
    On Error GoTo ColorsErrHandler
   
    With CmDia
        .CancelError = True
        .Flags = cdlCCFullOpen
        .DialogTitle = "Text Color"
        .ShowColor
        bk = Right$("000000" & Hex(.Color), 6)
    End With
   
    ' Switch code from Windows '
    ' format to HTML format:   '
    Text2(1).SelText = Right$(bk, 2) & _
                         Mid$(bk, 3, 2) & _
                        Left$(bk, 2) & " " & Right$(txt, 2) & Mid$(txt, 3, 2) & Left$(txt, 2)
               
ExitColors:
    On Error GoTo 0
Exit Sub

ColorsErrHandler:
    'User pressed Cancel button.
    Resume ExitColors
End Sub


Private Sub mColorWords_Click()
    ' HTML code to change text color. '
    
    Dim hx As String
    Dim HTMLcode As String
    
    On Error GoTo ColorWordsErrHandler
    
    With CmDia
        .CancelError = True
        .Flags = cdlCCFullOpen
        .ShowColor
        hx = Right$("000000" & Hex(.Color), 6)
    End With
   
    HTMLcode = "#" & _
                Right$(hx, 2) & _
                  Mid$(hx, 3, 2) & _
                 Left$(hx, 2)
   
    Text2(1).SelText = _
              "<SPAN style=" & qt & "color:" & _
              HTMLcode & _
              qt & ">" & _
              Text2(1).SelText & "</SPAN>"
              
ExitColorWords:
    On Error GoTo 0
Exit Sub

ColorWordsErrHandler:
    Resume ExitColorWords
End Sub


Private Sub mDeleteFile_Click()
    ' Delete a file. '
    Dim i As Integer
    Dim F As String
    Dim ret As Long
    
    F = L_Name(0)
    ret = MsgBox("Delete " + F + "?", vbYesNoCancel)
  
    If ret <> vbYes Then
        MsgBox "Deletion aborted."
        Exit Sub
    End If
  
    ' Delete the file: '
    Kill str_FiPath(0) & F
    
    ' If the right window was the source file, '
    ' clear its window and variables too:      '
    If L_Name(1) = F Then
        Text2(1).Text = ""
        L_Name(1).Text = ""
        str_FiName(1) = ""
        str_FiPath(1) = ""
        Text2(1).Visible = False
        brwWebBrowser(1).Visible = False
        lng_File2ViewerWidth = Me.Width
        If lng_File1ViewerWidth > 0 Then
            Me.Width = lng_File1ViewerWidth
        End If
        SetMenuStatus False
    End If
  
    ' Clear the variables: '
    If Text2(0).Visible Then
        Text2(0).Visible = False
        Text2(0).Text = ""
        brwWebBrowser(0).Visible = True
    End If
  
    If mOpenRight.Checked Then mOpenRight_Click
    L_Name(0).Text = "BidBase_File_Viewer.htm"
    str_FiPath(0) = str_AppPath & "Notes\"
    str_FiName(0) = "CardShark_BidBase_File_Viewer.htm"
    brwWebBrowser(0).Navigate str_FiPath(0) & str_FiName(0)
    ChangeCaption
End Sub


Private Sub mEditHTML_Click()
    ' Edit HTML code for left browser, '
    ' or if already editing, close it. '
    Dim i As Integer
    If Left$(str_FiName(0), 4) = "http" Then
        MsgBox "To edit a Web page's source HTML" & vbCr _
             & "code you must click File - Save Web Page" & vbCr _
             & "to save the source HTML code to disk." & vbCr _
             & "The source code will then be saved to" & vbCr _
             & "disk and displayed in the right window." & vbCr & "Some online features of Web pages will not work."
        Exit Sub
    End If
    
    bln_NoResize = True
    GetActiveMonitor
    
    If L_Name(0).Text = "" Then
        ' Not editing yet, so get ready: '
        L_Name(0).Text = str_FiName(0)
        Me.Caption = str_FiName(0) & " - BidBase File Viewer"
    End If
    If L_Name(0).Text = L_Name(1).Text Then
        ' Already editing, now close: '
        
        ' Save last width for 2 file viewers. '
        lng_File2ViewerWidth = Me.Width
        ' Clear right-window variables: '
        str_FiPath(1) = ""
        str_FiName(1) = ""
        brwWebBrowser(1).Visible = False
        Text2(1).Visible = False
        L_Name(1) = ""
        L_Name(1).Visible = False
        ' Resize form to 1-window width: '
        Me.Width = brwWebBrowser(0).Width + 500
        lng_File1ViewerWidth = Me.Width
        ' Gray-out the HTML menus: '
        SetMenuStatus False
          
    Else 'Load HTML source file:
        
        ' Enable HTML menus: '
        SetMenuStatus True
        
        ' Save 1-window form width: '
        ' (Should alredy be saved.) '
        If Text2(1).Visible = False And _
           brwWebBrowser(1).Visible = False _
        Then
            lng_File1ViewerWidth = lng_MonPosition(3, int_ActiveMonitor) - lng_MonPosition(1, int_ActiveMonitor)
            Me.Width = 29000 ' lng_File1ViewerWidth ' Me.ScaleX(lng_MonPosition(3, int_ActiveMonitor) - _
                                 lng_MonPosition(1, int_ActiveMonitor), _
                                 vbPixels, vbTwips)
            Me.Left = lng_MonPosition(1, int_ActiveMonitor) ' Me.ScaleX(lng_MonPosition(1, int_ActiveMonitor), _
                                vbPixels, vbTwips)
        End If
        ' Set up right window variables: '
        If Left$(L_Name(0), 8) = "file:///" Then
            FixFileName
        End If
                
        str_FiName(1) = str_FiName(0)
        str_FiPath(1) = str_FiPath(0)
        int_TxtNum = 1
        
        bln_NoResize = False
        LoadHTML 'This also adjusts object sizes on window.
        If Text2(1).Visible Then Text2(1).SetFocus
        str_Backup(0) = Text2(1).Text
    End If
    
    bln_NoResize = False
    Form_Resize
End Sub


Private Sub mExit_Click()
    Form_Unload 0
End Sub


Private Sub mFind_Click(index As Integer)
    ' Find/Replace text. '
    
    Dim defaultText As String
    Dim ii As Long
    Dim j As Long
    Dim K As Long
        
    If Text2(1).SelText <> "" Then
        ' User has selected text: '
        defaultText = Text2(1).SelText
   
    ElseIf Text2(1).SelStart > 0 Then
        ' Find the end of the word at the cursor: '
        ii = Text2(int_TxtNum).SelStart + 1
        j = ii + 1
      
        Do While UCase$(Mid$(Text2(int_TxtNum).Text, j, 1)) >= "A" _
             And UCase$(Mid$(Text2(int_TxtNum).Text, j, 1)) <= "Z" _
             And j < Len(Text2(int_TxtNum).Text)
            
            j = j + 1
            DoEvents
        Loop
      
        If UCase$(Mid$(Text2(int_TxtNum).Text, ii + 1, 1)) < "A" Or _
           UCase$(Mid$(Text2(int_TxtNum).Text, j, 1)) > "Z" _
        Then
            defaultText = Mid$(Text2(int_TxtNum).Text, ii, j - ii)
      
        Else
            ' Find start of word at cursor: '
            K = ii - 1
            
            Do While UCase$(Mid$(Text2(int_TxtNum).Text, K, 1)) >= "A" _
                 And UCase$(Mid$(Text2(int_TxtNum).Text, K, 1)) <= "Z" _
                 And K > 0
                
                K = K - 1
                DoEvents
            Loop
            
            defaultText = Mid$(Text2(int_TxtNum).Text, K + 1, j - K - 1)
        End If
    End If
    ' Show the find/replace form: '
    frmFindReplace.txtFind = defaultText
    frmFindReplace.Move Me.Left + 4000, Me.Top + 4000
    frmFindReplace.Show , Me
End Sub


Private Sub mFont_Click()
    ' Insert font codes. '
    Dim n As String
    Dim f1 As String
    Dim fs As Integer
    Dim fc As String
    Dim f2 As String
    
    ' Set Cancel to True
    CmDia.CancelError = True
    On Local Error GoTo FontErr
    ' Set the Flags property
    CmDia.Flags = cdlCFEffects Or cdlCFBoth
    ' Display the Font dialog box
    CmDia.ShowFont
    n = CmDia.FontName
    f1 = "<FONT FACE=" & qt & n & qt
  
    fs = CmDia.FontSize
    Select Case fs
        Case Is < 10
            fs = 1
        Case Is < 12
            fs = 2
        Case Is < 15
            fs = 3
        Case Is < 18
            fs = 4
        Case Is < 24
            fs = 5
        Case Is < 36
            fs = 6
        Case Is < 48
            fs = 7
        Case Else
            fs = 8
    End Select
    
    f1 = f1 & " SIZE=" & qt & fs & qt

    fc = Right$("00000" & Hex$(CmDia.Color), 6)
    fc = "#" _
          & Right$(fc, 2) _
            & Mid$(fc, 3, 2) _
           & Left$(fc, 2)
           
    f1 = f1 & " COLOR=" & qt & fc & qt & ">"
  
    f2 = "</FONT>"
  
    If CmDia.FontBold Then
        f1 = f1 & "<B>"
        f2 = "</B>" & f2
    End If
  
    If CmDia.FontItalic Then
        f1 = f1 & "<I>"
        f2 = "</I>" & f2
    End If
  
    If CmDia.FontUnderline Then
        f1 = f1 & "<U>"
        f2 = "</U>" & f2
    End If
  
    If CmDia.FontStrikethru Then
        f1 = f1 & "<S>"
        f2 = "</S>" & f2
    End If
  
    With Text2(int_TxtNum)
        .SelText = f1 & .SelText & f2
    End With
    
    mSave_Click 0

FontErr:
    On Local Error GoTo 0
    ' User pressed the Cancel button
End Sub


Private Sub mHeader_Click(index As Integer)
    ' Inserts HTML Header code '
    ' (e.g.: <H1>, <H2>, etc.  '
    
    With Text2(int_TxtNum)
        If index = 0 Then
            .SelText = "<H1 ALIGN=CENTER>" & .SelText & "</H1>"
        ElseIf index = 1 Then
            .SelText = "<H2>" & .SelText & "</H2>"
        ElseIf index = 2 Then
            .SelText = "<H3>" & .SelText & "</H3>"
        End If
    End With
End Sub


Private Sub mHelp_Click()
    Dim ret As Long
    Dim F As String
    
    F = str_FiPath(0) & str_FiName(0)
    
    ret = Shell(str_AppPath & "BidBase File Viewer.exe " & _
                str_FiPath(0) & "CardShark_BidBase_File_Viewer.htm", _
                vbNormalFocus)
End Sub


Private Sub mHideText_Click()
    ' Insert code which will hide text '
    ' until the user does a mouseover  '
    ' of the highlighted prompt.       '
    Dim crlf As String
    Dim id1 As String
    Dim id2 As String
    Dim id3 As String
    Dim i As Long
    Dim sel As String
    Dim s As String
    Dim hideID As Long
    Dim t As String
    Dim selStart1 As Long
    Dim selLength1 As Long
    
    crlf = Chr$(13) & Chr$(10)
    
    selStart1 = Text2(int_TxtNum).SelStart
    selLength1 = Text2(int_TxtNum).SelLength
    sel = Trim$(Text2(int_TxtNum).SelText)
    
    If sel = "" Then
        sel = "<P><B><I>Put the hidden text, etc., here.</I></B>"
    End If
    
    i = 200
    Do
        hideID = hideID + 1
        i = InStr(i + 10, Text2(1).Text, "hidden" & hideID)
        If i = 0 Then Exit Do
        DoEvents
    Loop While i > 0
    
    id1 = "" & hideID
    id2 = "" & (hideID + 1)
    id3 = "" & (hideID + 2)
    
    i = InStr(Left$(Text2(1).Text, 1000), "cursor: pointer;")
    If i = 0 Then
        i = InStr(UCase$(Left$(Text2(1).Text, 1000)), "</HEAD>")
        If i = 0 Then
            i = InStr(990, UCase$(Left$(Text2(1).Text, 1000)), "</HEAD>")
        End If
        If i > 0 Then
            ' Set up the CSS for a Hand cursor with    '
            ' color text for a prompt like Click Here: '
            Text2(1).SelStart = i - 2
            Text2(1).SelLength = 1
            t = "<style>" & vbCr & "#GreenHand {" & vbCr & "    color: green;" & vbCr & "    cursor: pointer;" & vbCr & "}" & vbCr & "#RedHand {" & vbCr & "    color: red;" & vbCr & "    cursor: pointer;" & vbCr & "}" & vbCr & "</style>" & vbCr
            Text2(1).SelText = t
            Text2(1).SelStart = selStart1 + Len(t) - 1
            Text2(1).SelLength = selLength1
        End If
    End If
                
    ' The simple version: '
    s = "<!--- Leave the next line in or this routine won't work with Chrome. -->" & crlf
    s = s & "<P></P>"
    s = s & "<!--- The next line IDs the prompt code. STYLE='' means the prompt is visible. -->" & crlf
    s = s & "<SPAN ID=" & qt & "prompt" & id1 & qt & " STYLE=''" & qt & " COLOR=BLACK>" & crlf
    s = s & "<P><I>Enter any text here you want to show before the prompt but not after it's clicked.<BR>" & crlf
    s = s & "    <!--- The next line hides this section when the user clicks the prompt. -->" & crlf
    s = s & "    <SPAN OnClick=" & qt & "prompt" & id1 & ".style.display='none'" & qt & ">" & crlf
    s = s & "        <!--- The next line unhides the hidden section when the user clicks the prompt. -->" & crlf
    s = s & "        <SPAN OnClick=" & qt & "hidden" & id1 & ".style.display=''" & qt & ">" & crlf
    s = s & "            <SPAN id=" & qt & "GreenHand" & qt & ">" & crlf
    s = s & "                 <U><B>Click here</B> <I>or put whatever prompt you want here</I></U>" & crlf
    s = s & "           </SPAN>" & crlf
    s = s & "       </SPAN>" & crlf
    s = s & "   </SPAN>" & crlf
    s = s & "<P></I>Enter any text you want to appear below the prompt but not after the click.</I>" & crlf
    s = s & "</SPAN>" & crlf
    s = s & crlf  ' Separate the two sections of code with a blank line. '
    
    s = s & "<SPAN ID=" & qt & "hidden" & id1 & qt & " STYLE=" & qt & "display:none" & qt & " COLOR=BLACK>" & crlf
    s = s & crlf & sel & crlf
    s = s & "    <SPAN onClick=" & qt & "prompt" & id1 & ".style.display=''" & qt & ">" & crlf
    s = s & "        <SPAN OnClick=" & qt & "hidden" & id1 & ".style.display='none'" & qt & "><CENTER>" & crlf
    s = s & "            <SPAN id=" & qt & "RedHand" & qt & ">" & crlf
    s = s & "                <BR><B><U>Close</U></B></CENTER>" & crlf
    s = s & "           </SPAN>" & crlf
    s = s & "       </SPAN>" & crlf
    s = s & "   </SPAN>" & crlf
    s = s & "</SPAN>"
    Text2(int_TxtNum).SelText = s
    
Exit Sub
    ' Using multiple displays:
    s = ""
    Text2(int_TxtNum).SelText = _
        s = "<FONT ID=" & qt & id1 & qt & " STYLE=''" & ">" & crlf
        s = s & "<I>Put anything here leading up to the 'Click here' prompt.</I>" & crlf
        s = s & crlf
        s = s & "<!--- PROMPT FOR FIRST HIDDEN SECTION: -->"
        s = s & crlf
        s = s & "   <SPAN onClick=" & qt & id2 & ".style.display=''" & qt & ">" & crlf
        s = s & "      <SPAN onClick=" & qt & id1 & ".style.display='none'" & qt & ">" & crlf
        s = s & "<FONT STYLE=" & qt & "cursor:hand" & qt & " COLOR=" & qt & "#0000EE" & qt & ">" & crlf
        s = s & " <U>Click here<U>. <I>Or change this to whatever prompt you want.</I> " & crlf
        s = s & "</FONT>" & crlf
        s = s & "      </SPAN> " & crlf
        s = s & "   </SPAN>" & crlf
        s = s & "</FONT>" & crlf
        s = s & crlf
        s = s & "<!--- DISPLAY FIRST HIDDEN SECTION. -->"
        s = s & crlf
        s = s & "<FONT ID=" & qt & id2 & qt & " STYLE=" & qt & "display:'none'" & qt & ">" & crlf
        s = s & "   <SPAN onClick=" & id1 & ".style.display=''" & ">" & crlf
        s = s & "      <SPAN onClick=" & id3 & ".style.display='none'" & qt & ">" & crlf
        s = s & sel & crlf
        s = s & "<!--- PROMPT FOR 2ND HIDDEN SECTION: -->"
        s = s & crlf
        s = s & "<FONT STYLE=" & qt & "cursor:hand" & qt & " COLOR=" & qt & "#0000EE" & qt & qt & " > " & crlf
        s = s & "<P><CENTER><U>See next section</U>." & crlf
        s = s & "</FONT>"
        s = s & "      </SPAN>" & crlf
        s = s & "   </SPAN>" & crlf
        s = s & "</FONT>" & crlf
        s = s & crlf
        s = s & "<!--- DISPLAY 2ND HIDDEN SECTION -->"
        s = s & crlf
        s = s & "<FONT ID=" & qt & id3 & qt & " STYLE=" & qt & "display:'none'" & qt & ">" & crlf
        s = s & "   <SPAN onClick=" & id2 & ".style.display=''" & ">" & crlf
        s = s & "      <SPAN onClick=" & id1 & ".style.display='none'" & qt & ">" & crlf
        s = s & crlf
        s = s & "<B><I>Cut and paste the second section into this line.</I></B> & crlf"
        s = s & crlf
        s = s & "<!--- NEXT PROMPT: -->"
        s = s & crlf
        s = s & "<FONT STYLE=" & qt & "cursor:hand" & qt & " COLOR='RED'" & qt & " > " & crlf
        s = s & "<P><CENTER><B><I>Enter '<U>Close</U>' or '<U>See next section</U>." & crlf
        s = s & "</FONT>"
        s = s & "      </SPAN>" & crlf
        s = s & "   </SPAN>" & crlf
        s = s & "</FONT>"
        s = s & crlf
        s = s & "<!--- START OF SECOND HIDDEN TEXT/PIX/ETC. -->"
        s = s & crlf
Exit Sub
    
    ' Using onMouseOver: '
    Text2(int_TxtNum).SelText = _
        "<Span onMouseover=" & qt & _
        "hide" & hideID & "ID.style.display=''" & qt & Chr$(13) & Chr$(10) & _
        " onMouseOut=" & qt & "hide" & hideID & "ID.style.display='none'" & qt & ">" & _
        "<FONT color=" & qt & "green" & qt & ">" & Chr$(13) & Chr$(10) & _
        Text2(int_TxtNum).SelText & Chr$(13) & Chr$(10) & _
        "</FONT></SPAN> " & Chr$(13) & Chr$(10) & _
        "<FONT ID=" & qt & "hide" & hideID & "ID" & qt & _
        " STYLE=" & qt & "display:none" & qt & " COLOR=" & qt & "black" & qt & _
        ">" & _
        ""
End Sub

Private Sub mHideDeeper_Click()
    Text2(1).SelText = "<!---" & Text2(1).SelText + "-->"
End Sub

Private Sub mHTMLStartEnd_Click()
    ' Put the starting and ending '
    ' HTML code in the file.      '
  
    Dim st As Long
    Dim a1 As String
    Dim a2 As String
    Dim qt As String
    
    qt = qt
    
    If UCase$(Left$(Text2(int_TxtNum).Text, 6)) = "<HTML>" Then
        MsgBox "This file already has HTML headers."
        Exit Sub
    End If
  
    st = Text2(int_TxtNum).SelStart
  
    'Start:
    a1 = "<HTML><HEAD>" & vbCrLf _
        & "<TITLE>" & L_Name(int_TxtNum) & "</TITLE>" & vbCrLf _
        & "</HEAD>" & vbCrLf _
        & "<BODY" & vbCrLf _
        & "  TEXT=" & qt & "#000000" & qt & vbCrLf _
        & "  BGCOLOR=" & qt & "#FFFFFF" & qt & vbCrLf _
        & "  LINK=" & qt & "#0000EE" & qt & vbCrLf _
        & "  VLINK=" & qt & "#008000" & qt & vbCrLf _
        & "  ALINK=" & qt & "#FF0000" & qt & vbCrLf _
        & ">" & vbCrLf _
        & "<CENTER>" & vbCrLf _
        & "<TABLE WIDTH=850><TR><TD>" & vbCrLf & vbCrLf _
        & "<H1 ALIGN=CENTER>" & L_Name(int_TxtNum) & "</H1>" _
        & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
  
    ' End of HTML file: '
    a2 = vbCrLf & vbCrLf _
         & "<BR><BR><BR><BR><BR><BR><BR>" _
         & "<A NAME=""EndOfFile"">" & vbCrLf _
         & "</TD></TR></TABLE>" & vbCrLf _
         & "</BODY></HTML>"
    
    ' Add the code: '
    Text2(int_TxtNum).Text = a1 & Text2(int_TxtNum).Text & a2
    Text2(int_TxtNum).SelStart = st + 500
    Text2(int_TxtNum).Refresh
    Text2(int_TxtNum).SelStart = st
  
    ' If file name was *.TXT, change to *.HTM: '
    If LCase$(Right$(L_Name(1), 3)) = "txt" And _
       L_Name(0) = L_Name(1) _
    Then
        str_FiName(0) = Left$(str_FiName(0), Len(str_FiName(0)) - 3) & "htm"
        Name str_FiPath(1) & str_FiName(1) As _
             str_FiPath(0) & str_FiName(0)
        str_FiName(1) = str_FiName(0)
        L_Name(0).Text = str_FiName(0)
        Me.Caption = str_FiName(0) & " - BidBase File Viewer"
        L_Name(1).Text = str_FiName(0)
        brwWebBrowser(0).Visible = True
        brwWebBrowser(0).Navigate str_FiPath(1) + str_FiName(1)
        UpdateRecentFiles 0
    End If
End Sub


Private Sub mIndent_Click()
    With Text2(int_TxtNum)
        .SelText = "<UL><LI>" & .SelText + "</LI></UL>"
    End With
End Sub


Private Sub mIndexNoBreak_Click()
    Text2(int_TxtNum).SelText = "<BLOCKQUOTE PADDING-LEFT: 5px>" _
                                & vbCrLf & Text2(int_TxtNum).SelText & vbCrLf & _
                                "</BLOCKQUOTE>"
End Sub

Private Sub mIndentWithoutBreakAll_Click()
    Dim x As Integer
    With Text2(int_TxtNum)
        x = MsgBox("This will set all indents in the file" & Chr$(13) & "to having no break before them." & Chr$(13) & "Continue?", vbYesNo)
        If x = vbNo Then Exit Sub
        x = InStr(LCase(Left$(.Text, 300)), "margin-top: 0px;")
        If x > 0 Then
            MsgBox "This option has already been set."
            Exit Sub
        End If
        x = InStr(LCase(Left$(.Text, 300)), "</head>")
        If x = 0 Then
            MsgBox "This file does not have an HTML header."
            Exit Sub
        End If
        Text2(int_TxtNum) = Left$(.Text, x - 1) & Chr$(13) & _
                            "<style>" & Chr$(13) & _
                            "ul {margin-top: 0px;" & Chr$(13) & _
                            "margin-bottom: 0px;}" & Chr$(13) & _
                            "</style>" & Chr$(13) & _
                            "<style>" & Chr$(13) & _
                            "ol {margin-top: 0px;" & Chr$(13) & _
                            "margin-bottom: 0px;}" & Chr$(13) & _
                            "</style>" & Chr$(13) & _
                            Mid$(.Text, x)
    End With
End Sub

Private Sub mInsertBidFont_Click()
    
    Dim ret As Long
    Dim t As String
    Dim n As Integer
    Dim b As String
    Dim c As String
    Dim i As Integer
    Dim j As Integer
    
    ret = MsgBox("This function looks for bids in the" & vbCr _
         & "file and puts them in a Courier font." & vbCr & vbCr _
         & "It assumes that you have already" & vbCr _
         & "converted Suits To Symbols and" & vbCr _
         & "that bids are in a format like" & vbCr _
         & " 1C-1D-P-1H, 1S... or " & vbCr _
         & " 1C-(1D)-P-(1H), 1S..." & vbCr & vbCr _
         & "Do it?", vbYesNo)
    
    If ret = vbNo Then Exit Sub
    
    t = Text2(1).Text
    If str_UndoBidFont = "" Then
        ' Keep original text for Undo: '
        str_UndoBidFont = t$
    End If
    
    For n = 1 To 7
        ' Since suits-to-symbols must already be '
        ' done, look for numbers 1-7 followed by '
        ' the image reference, "<IMG..."         '
        b = " " & n & "<IMG"
        i = InStr(i + 1, t, b)
        If i > 0 Then
            ' Find end of bid sequence: '
            j = i + 4
            Do
                j = j + 1
                If j = Len(t) Then
                    Exit Do
                End If
                ' Get the character: '
                c = Mid$(t, j, 1)
                If c$ = "(" Then
                    Do
                        j = j + 1
                        DoEvents
                    Loop While Mid$(t, j, 1) <> ")"
                End If
                DoEvents
            Loop While InStr(">NPDR", c) = 0
        End If
        MsgBox "Stop mInsertBidFont" ' I can't see where the font is going in. '
    Next
End Sub

Private Sub mItalics_Click()
    With Text2(int_TxtNum)
        If Right$(.SelText, 1) = " " Then
            ' Double-clicking a word causes a blank space '
            ' next to it to be highlighted too. (Thanks,  '
            ' Microsoft.)                                 '
            .SelText = "<I>" & Trim$(.SelText) & "</I> "
        Else
            .SelText = "<I>" & .SelText & "</I>"
        End If
    End With
End Sub

Private Sub mLinkExt_Click()
    ' External link: '
    CmDia.DialogTitle = "Select file."
    CmDia.FileName = "C:\*.*"
    On Local Error GoTo ExitLink
    CmDia.ShowOpen
    If Right$(CmDia.FileName, 1) = "*" Then GoTo ExitLink
    Text2(int_TxtNum).SelText = "<A HREF=" & qt & CmDia.FileTitle & qt & _
                                ">" & Text2(int_TxtNum).SelText & "</A>"
    Exit Sub
    
ExitLink:
    On Local Error GoTo 0
    Text2(int_TxtNum).SelText = "<A HREF=" & qt & qt & ">" & Text2(int_TxtNum).SelText & "</A>"
End Sub

Private Sub mLinkIn_Click()
    ' Create an internal link to jump '
    ' to a specific part of the file. '
    ' The program keeps track of the  '
    ' last link tag and increments it '
    ' automatically.                  '
    
    Dim st As Long
    Dim alph As String
    Dim ln As String
    Dim t As String
    Dim iy As Long
    Dim ix As Long
    Dim a1 As String
    
With Text2(int_TxtNum)

    alph = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    st = .SelStart + 1
    ' Note: If you look at Mid$(.Text, .SelStart), '
    ' it will start with the character to the left '
    ' of the cursor.                               '
  
    If LinkMsgBox.Visible = False Then
        lng_CursPtr = st
        lng_TextLen = .SelLength
        
        ' Find the next link tag to use: '
        Do
            int_IntLinkNum = int_IntLinkNum + 1
            ln = Mid$(alph, int_IntLinkNum, 1)
            DoEvents
        Loop While _
            InStr(.Text, "HREF=#" & ln & ">") > 0 _
            And int_IntLinkNum < 27 ' end of alphabet
    
        Do While InStr(.Text, "HREF=#" & ln) > 0 _
                 And int_IntLinkNum < 53 ' end of double-alphabet: "ZZ")
      
            If int_IntLinkNum < 27 Then
                ln = Mid$("ABCDEFGHIJKLMNOPQRSTUVWXYZ", int_IntLinkNum, 1)
            Else
                ln = "A" & Mid$(alph, int_IntLinkNum - 26, 1)
            End If
      
            If InStr(.Text, "HREF=#" & ln) = 0 Then
                ' Next unused link found, so exit Do '
                Exit Do
            End If
      
            int_IntLinkNum = int_IntLinkNum + 1
            DoEvents
        Loop
    
        LinkMsgBox.Visible = True
        LinkMsgBox.Left = .Left + 200
        LinkMsgBox.ZOrder 0
        .SetFocus
  
    Else ' Source tag already inserted. '
         ' Now insert target code:      '
         
        Dim tt As String
        Dim origLN As String
        Dim i As Integer
        Dim ss As Long
        Dim reuseLink As Boolean
        
        LinkMsgBox.Visible = False
            
        tt = UCase$(Mid$(.Text, st - 20, 50))
        i = InStr(tt, "<A NAME=")
        If i > 0 Then
            reuseLink = True
            ln = Mid$(tt, i + 9, 3)
            Do While Right$(ln, 1) = ">" Or _
                     Right$(ln, 1) = qt
                ln = Left$(ln, Len(ln) - 1)
                DoEvents
            Loop
            If int_IntLinkNum > 0 Then
                int_IntLinkNum = int_IntLinkNum - 1
            End If
            
        ElseIf int_IntLinkNum < 27 Then ' one-letter tag '
            ln = Mid$(alph, int_IntLinkNum, 1)
        
        Else ' two-letter tag, starting with "A": '
            ln = "A" + Mid$(alph, int_IntLinkNum - 26, 1)
        End If
    
        ' Insert the Link code to jump to:   '
        ' (We do this first because the user '
        ' just clicked on that point.)       '
        If Not reuseLink Then
            a1 = "<A NAME=" & qt + ln & qt & "></A>"
            .SelText = a1
        End If
        
        ' lng_CursPtr points to the cursor position '
        ' of the link. If the jump-to point comes   '
        ' before the link, then after the code is   '
        ' inserted (above), lng_CursPtr no longer   '
        ' points to the right spot, so increment    '
        ' lng_CursPtr by the length of the code:    '
        If .SelStart < lng_CursPtr Then
            lng_CursPtr = lng_CursPtr + Len(a1)
        End If
        
        ' Insert the Link code to jump from: '
        .SelStart = lng_CursPtr - 1
        .SelLength = lng_TextLen
        .SelText = "<A HREF=#" & ln & ">" & .SelText + "</A>"
            
        ' Put the original text back in view: '
        If lng_CursPtr - 100 > 0 Then
            .SelStart = lng_CursPtr - 100
        Else
            .SelStart = 0
        End If
        
        ' Put the cursor back where it started: '
        .SelStart = lng_CursPtr
    End If
End With

End Sub


Private Sub mMergeCellsAcross_Click()
    With Text2(int_TxtNum)
        .SelText = "<TD ColSpan=" & Chr$(34) & Chr$(34) & ">" & .SelText & "</TD>"
    End With
End Sub


Private Sub mMergeCellsDown_Click()
    With Text2(int_TxtNum)
        .SelText = "<TD RowSpan=" & Chr$(34) & Chr$(34) & ">" & .SelText & "</TD>"
    End With
End Sub


Private Sub mNewFile_Click()
    ' Start a new file. '
    
    Dim fiNum As Integer
    Dim Cancel As Boolean
    Dim F As String
    Dim n As Integer
    Dim t As String
    
    CheckForSave Cancel
    If Cancel Then Exit Sub
  
  
    Dim p As String
    Dim i As Integer
  
    ' Use last path as default: '
    p = str_FiPath(1)
    If p = "" Then p = str_FiPath(0)
    
    On Local Error GoTo NewFileError
    CmDia.DialogTitle = "Select file."
    CmDia.FileName = p & "*.*"
    CmDia.ShowOpen
    On Local Error GoTo 0
    
    F = CmDia.FileTitle
    If F = "" Or F = str_FiName(int_TxtNum) Then Exit Sub
  
    ' Extract path from full file spec: '
    p = CmDia.FileName
    i = InStr(p, F)
    If i = 0 Then i = InStr(p, "*.*")
    p = Left$(p, i - 1)
    If Right$(p, 1) <> "\" Then
        p = p & "\"
    End If
  
    str_FiName(0) = F
    str_FiPath(0) = p
  
    'str_FiName(0) = InputBox$("Give us a filename:", "Start A New HTM File")
    'If str_FiName(0) = "" Then Exit Sub
  
    'str_FiName(0) = Replace(str_FiName(0), " ", "_", 1, 99)
    
    ' Must use editor on right to create '
    ' a new file, so set up the right    '
    ' window variables too.              '
    str_FiName(1) = str_FiName(0)
    str_FiPath(1) = str_FiPath(0)
    L_Name(0).Text = str_FiName(0)
    Me.Caption = str_FiName(0) & " - BidBase File Viewer"
    n = Len(F)
    
    If Mid$("    " & F, n - 4, 1) = "." Then
        F = Left$(F, n - 4) ' xy.htm '
    End If
    
    On Local Error Resume Next
  
    If Right$(F, 4) = ".htm" Then
        t = Left$(F, Len(F) - 4)
    Else
        t = F
    End If
    
    If Dir$(str_FiPath(0) & str_FiName(0)) = "" Then
        fiNum = FreeFile
        Open str_FiPath(0) & str_FiName(0) For Output As #fiNum
            Print #fiNum, "<HTML><HEAD>"
            Print #fiNum, "<TITLE>"; t; "</TITLE>"
            Print #fiNum, "</HEAD>"
            Print #fiNum, "<BODY"
            Print #fiNum, "  TEXT="; qt; str_FormColors(0); qt
            Print #fiNum, "  BGCOLOR="; qt; str_FormColors(1); qt
            Print #fiNum, "  LINK="; qt; str_FormColors(2); qt
            Print #fiNum, "  VLINK="; qt; str_FormColors(3); qt
            Print #fiNum, "  ALINK="; qt; str_FormColors(4); qt
            Print #fiNum, ">"
            Print #fiNum, "<CENTER>"
            Print #fiNum, "<H1>"; t; "</H1>"
            Print #fiNum, ""
            Print #fiNum, "<TABLE WIDTH=90%><TR><TD>"
            Print #fiNum, "<FONT SIZE=+1>"
            Print #fiNum, ""
            Print #fiNum, "<H2></H2>"
            Print #fiNum, "<P>"
            Print #fiNum, ""
            Print #fiNum, "<H2></H2>"
            Print #fiNum, "<P>"
            Print #fiNum, ""
            Print #fiNum, ""
            Print #fiNum, ""
            ' End of HTML file: '
            Print #fiNum, "<BR><BR><BR><BR><BR><BR>"
            Print #fiNum, "</TD></TR></TABLE>"
            Print #fiNum, "<A NAME=""EndOfFile""></A>"
            Print #fiNum, "<P></BODY></HTML>"
        Close #fiNum
    End If
    
    brwWebBrowser(0).Navigate str_FiPath(0) & str_FiName(0)
    UpdateRecentFiles 0
    int_IntLinkNum = 0 'Reset internal link number
    int_TxtNum = 1
    str_UndoSuits2Symbols = Empty
    mUndoSuits2Symbols.Enabled = False
    ' Load the source on right: '
    LoadHTML
    str_Orig(1) = Text2(1).Text
    str_Backup(0) = str_Orig(1)
    str_Backup(1) = str_Orig(1)
    Text2(1).SelStart = 219
Exit Sub

NewFileError:
    'CmdFiName = ""
    Resume
    
End Sub


Private Sub mNumberedList_Click()
    With Text2(int_TxtNum)
        .SelText = "<OL>" & vbCrLf & "<LI> </LI>" _
        & vbCrLf & "<LI> </LI>" & vbCrLf & "</OL>" _
        & .SelText
    End With
End Sub


Private Sub mOpen_Click()
    Dim p As String
    Dim F As String
    Dim i As Integer
  
    ' Use last path as default: '
    p = str_FiPath(1)
    If p = "" Then p = str_FiPath(0)
    If p = "" Or Left$(p, 4) = "res:" Then
        p = "C:\"
    End If
    
    CmDia.DialogTitle = "Select file."
    CmDia.InitDir = p
    CmDia.Filter = "Help files (*.htm, *.txt)|*.htm;*.txt|All Files (*.*)|*.*"
    CmDia.ShowOpen
  
    F = CmDia.FileTitle
    If F = "" Or Right$(CmDia.FileName, 3) = "*.*" Then Exit Sub
  
    ' Extract path from full file spec: '
    p = CmDia.FileName
    i = InStr(p, F)
    p = Left$(p, i - 1)
  
    OpenFile p, F
End Sub
  
  
Private Sub OpenFile(ByVal p As String, _
                     ByVal F As String)
    Dim i As Integer
    Dim lastFile As String

    ' Open in right or left window? '
    If Text2(1).Visible _
       Or bln_OpenInRightWindow = False _
    Then
        int_TxtNum = 0
        lastFile = str_FiName(0)
    Else
        int_TxtNum = 1
        mSwap.Enabled = True
        mCloseRight.Enabled = True
    End If
  
    str_FiName(int_TxtNum) = F
    str_FiPath(int_TxtNum) = p
    
    ' Check for internal link (e.g.: mypage.htm#A1)
    i = InStr(Len(F) - 5, F, "#")
    If i > 0 Then
        L_Name(int_TxtNum).Text = Left$(F, i - 1)
    Else
        L_Name(int_TxtNum).Text = F
    End If
  
    L_Name(int_TxtNum).Visible = True
    
    Me.Caption = L_Name(int_TxtNum) & " - BidBase File Viewer"
    
    If str_FiPath(int_TxtNum) = "" And _
       Left$(F, 4) <> "http" _
    Then
        str_FiPath(int_TxtNum) = "C:\"
    End If
  
    brwWebBrowser(int_TxtNum).Navigate _
        str_FiPath(int_TxtNum) & _
        str_FiName(int_TxtNum)
    If DebugMode Then MsgBox "Ck.Pt. OpenFile"
    UpdateRecentFiles int_TxtNum
  
    brwWebBrowser(int_TxtNum).Visible = True
  
    ' Reset variables: '
    str_UndoSuits2Symbols = Empty
    mUndoSuits2Symbols.Enabled = False
    int_IntLinkNum = 0 'Reset internal link number
  
    ' If source code was shown on '
    ' right, show source for new  '
    ' file:                       '
    If int_TxtNum = 0 And _
       Text2(1).Visible And _
       lastFile = L_Name(1) _
    Then
        If Left$(str_FiName(int_TxtNum), 4) <> "http" Then
            int_TxtNum = 1
            str_FiName(1) = str_FiName(0)
            LoadHTML
        End If
    Else
        If int_TxtNum = 1 Then
            bln_NoResize = True
            GetActiveMonitor
            Me.Width = lng_MonPosition(3, int_ActiveMonitor)
            bln_NoResize = False
        End If
        Form_Resize
    End If
End Sub


Private Sub mOpenRight_Click()
    bln_OpenInRightWindow = Not bln_OpenInRightWindow
    mOpenRight.Checked = bln_OpenInRightWindow
End Sub


Private Sub mParagraph_Click()
    Dim paragStart As Long
    Dim paragEnd As Long
    Dim paragMark As Long
    Dim headerMark As Long
    Dim startPoint As Long
    Dim dblCrLf As String
    Dim ctrlMark As Long
  
    With Text2(int_TxtNum)
    ' The user can select multiple paragraphs by highlighting '
    ' them and this routine will add <P> to the start of each '
    ' paragraph. An original paragraph in this routine is a   '
    ' group of lines ending with double CrLf's.               '
    dblCrLf = vbCrLf & vbCrLf
  
    paragStart = .SelStart
    If .SelLength = 0 Then
        paragEnd = InStr(paragStart + 1, .Text, dblCrLf)
    Else
        paragEnd = paragStart + .SelLength
    End If
    
    ' Some text is run together with <P> marks in '
    ' it or with Headings in it. If so, stop there:
    paragMark = InStr(paragStart + 2, .Text, "<P>") - 1
    If paragMark = 0 Then
        paragMark = InStr(paragStart, .Text, "<p>") - 1
    End If
    If paragMark > 0 And paragMark < paragEnd Then
        paragEnd = paragMark
    End If
    
    headerMark = InStr(paragStart, .Text, "<H") - 1
    If headerMark = 0 Then
        headerMark = InStr(paragStart, .Text, "<h") - 1
    End If
    If headerMark > 0 And headerMark < paragEnd Then
        paragEnd = headerMark
    End If
    
    ctrlMark = InStr(paragStart, .Text, "<PRE>") - 1
    If ctrlMark > 0 And ctrlMark < paragEnd Then
        paragEnd = ctrlMark
    End If
    
    ctrlMark = InStr(paragStart, .Text, "<UL>") - 1
    If ctrlMark > 0 And ctrlMark < paragEnd Then
        paragEnd = ctrlMark
    End If
    
    ctrlMark = InStr(paragStart, .Text, "<OL>") - 1
    If ctrlMark > 0 And ctrlMark < paragEnd Then
        paragEnd = ctrlMark
    End If
    
    ctrlMark = InStr(paragStart, .Text, "<DL>") - 1
    If ctrlMark > 0 And ctrlMark < paragEnd Then
        paragEnd = ctrlMark
    End If
    
    startPoint = paragStart ' first starting point
    If .SelLength = 0 Then
        .SelLength = paragEnd - startPoint
    End If
  
    Do
        If Left$(UCase$(.SelText), 3) <> "<P>" Then
            .SelText = "<P>" & .SelText
            paragEnd = paragEnd + 3
        End If
        ' Just to be neat, take out any CrLf's at the end of '
        ' lines in the paragraph (up to the double CrLf's):  '
        paragStart = InStr(paragStart + 3, .Text, vbCrLf)
        Do While paragStart > 0 And paragStart + 2 < paragEnd
            ' Apparently, you cannot just replace the CrLf with a '
            ' blank space. (Without the space, the two lines run  '
            ' together: "we", "saw" becomes "wesaw", instead of   '
            ' "we saw".) So it has to be done in two steps:
            .SelStart = paragStart
            .SelLength = 2
            .SelText = .SelText & " "
            .SelStart = paragStart
            .SelLength = 2
            .SelText = ""
            .SelStart = paragStart - 2
            .SelLength = 2
            Do While .SelText = "  "
                .SelText = " "
                paragEnd = paragEnd - 1
                .SelStart = .SelStart - 1
                .SelLength = 2
                DoEvents
            Loop
            paragEnd = paragEnd - 1 ' Replacing Cr-Lf with a space  '
                        ' reduces the end location by 1.'
            paragStart = InStr(paragStart, .Text, vbCrLf)
            DoEvents
        Loop
        ' End of one paragraph reached, do next '
        ' paragraph if any more were highlighted:
        If InStr(paragStart, .Text, dblCrLf) > InStr(paragStart + 1, .Text, "<P>") Then
            Exit Do
        End If
        
        .SelStart = InStr(paragStart, .Text, dblCrLf) + 3
        
        If paragEnd > .SelStart Then
            .SelLength = paragEnd - .SelStart
        Else
            Exit Do
        End If
        DoEvents
    Loop While paragStart < paragEnd
    
    ' Go back to original starting point: '
    If paragEnd < .SelStart Then paragEnd = .SelStart
    If Mid$(.Text, paragEnd - 2, 2) <> dblCrLf And Mid$(.Text, paragEnd, 2) <> dblCrLf Then
        .Text = Left$(.Text, paragEnd - 1) & Mid$(.Text, paragEnd)
    End If
    
    If paragEnd > startPoint + 5 Then
        .SelStart = paragEnd + 1000
        .SelStart = paragEnd
    Else
        .SelStart = startPoint
    End If
    '      '
    End With
End Sub


Private Sub mPre_Click()
    ' <PRE> causes text to be displayed as formatted. '
    Text2(int_TxtNum).SelText = "<PRE>" & Text2(int_TxtNum).SelText & "</PRE>"
End Sub


Private Sub mPrint_Click()
    ' Print the web page: '
    Dim ret As Long
    Dim p As Integer
    
    If brwWebBrowser(0).Visible And _
       brwWebBrowser(1).Visible _
    Then
        ret = MsgBox("Print left panel?" & vbCr _
                     & "(Click NO for right panel.)", _
                    vbYesNoCancel)
        If ret = vbCancel Then
            Exit Sub
    
        ElseIf ret = vbYes Then
            p = 0
    
        Else
            p = 1
        End If
  
    ElseIf brwWebBrowser(0).Visible Then
        p = 0
  
    Else
        p = 1
    End If
    
    brwWebBrowser(p).ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
End Sub


Private Sub mPrintPreview_Click()
    Dim ret As Long
    Dim p As Integer
    
    If brwWebBrowser(0).Visible And _
       brwWebBrowser(1).Visible _
    Then
        ret = MsgBox("Preview left panel?" & vbCr _
                     & "(Click NO for right panel.)", _
                    vbYesNoCancel)
        If ret = vbCancel Then
            Exit Sub
        ElseIf ret = vbYes Then
            p = 0
        Else
            p = 1
        End If
    ElseIf brwWebBrowser(0).Visible Then
        p = 0
    Else
        p = 1
    End If
  
    brwWebBrowser(p).ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
End Sub


Private Sub mRecent_Click(index As Integer)
    ' Recent Files list in the File menu. '
    
    Dim p As String
    Dim F As String
    
    p = str_RecentFiles(0, index)
    F = str_RecentFiles(1, index)
    
    ' If "file" was a web page, its URL  '
    ' is stored in str_FiName(), so FiPath$ '
    ' should be left blank.              '
    If Left$(LCase$(F), 4) = "http" Then
        p$ = ""
    Else
        If Dir$(p & "*.*") = "" Then
            MsgBox "The path " & p & " is no longer present."
            Exit Sub
        ElseIf Dir$(p & F) = "" Then
            If InStr(F, ".") = 0 Then
                F = F & ".htm"
            End If
            If Dir$(p & F) = "" Then
                MsgBox "The file " & F & " is no longer present."
                Exit Sub
            End If
        End If
    End If
    
    int_TxtNum = 0
    
    OpenFile p, F
    If Right$("   " & str_FiName(0), 4) <> ".htm" And Right$("   " & str_FiName(0), 5) <> ".html" Then
        MsgBox "Viewer only opens .htm and .html files."
        F = ""
    End If
End Sub

Private Sub mResetMon_Click()
    If Dir$("\My Files\*.*") <> "" Then
        Me.Left = -120
    Else
        Me.Left = 0
    End If
    Me.Top = 0
    Me.Width = Screen.Width
    Me.Height = Screen.Height
End Sub

Private Sub mRestore_Click(index As Integer)
    Dim x As Integer
    
    Select Case index
    Case Is = 0
        x = MsgBox("Restore this file to its" & vbCr & "contents when first opened?", vbYesNo, "Restore Original File")
        If x = vbYes Then
            Text2(1) = str_Backup(0)
        End If
    Case Is = 1
        x = MsgBox("Restore this file to its" & vbCr & "contents before the last save?", vbYesNo, "Restore From Last Save")
        If x = vbYes Then
            Text2(1) = str_Backup(1)
        End If
    Case Is = 2
        x = MsgBox("Discard unsaved changes?", vbYesNo)
        If x = vbYes Then
            Text2(1) = str_Orig(1)
        End If
    End Select
    MsgBox "File has been restored in edit window." & vbCr & "Press Ctrl-S to save to disk."
End Sub

Private Sub mSaveWebPage_Click()
    ' When a Web page is shown,  '
    ' get its source, save it to '
    ' disk, display the source   '
    ' file in the left browser   '
    ' and the source in the right.
    
    ' Jan 2021 note: Since FileViewer was originally written, '
    ' web page source has moved way beyond HTML and very few  '
    ' page's source can be displayed in the current MS web    '
    ' control, so this feature has been dropped but this code '
    ' is being left in for future reference. '
    
    Dim b() As Byte
    Dim F As String
    Dim f1 As String
    Dim i As Integer
    
    F = InputBox$("Enter a file name to save as:")
    If F = "" Then Exit Sub
    If Right$(LCase$(F), 4) <> ".htm" Then
        F = F & ".htm"
    End If
    f1 = F
    ' Use the Inet control to '
    ' get the source code:    '
    Inet1.URL = str_FiName(0)
    b() = Inet1.OpenURL(, icByteArray)
    ' Save it to disk: '
    Open f1 For Binary Access Write As #1
        Put #1, , b()
    Close #1
    
    ' Set the variables to the local '
    ' file & load it in the browser: '
    i = InStrRev(F, "\")
    If i > 0 Then
        str_FiPath(0) = Left$(F, i)
        F = Mid$(F, i + 1)
    Else
        str_FiPath(0) = "c:\"
    End If
    str_FiName(0) = F
    L_Name(0) = F
    FixFileName
    brwWebBrowser(int_TxtNum).Navigate str_FiPath(0) & str_FiName(0)
    ' Load the source into the edit window: '
    Text2(1).Visible = False
    brwWebBrowser(1).Visible = False
    mEditHTML_Click
End Sub


Private Sub mSpecial_Click(index As Integer)
    ' Insert special characters: '
    Dim code
    
        Clipboard.Clear
    Select Case index
        Case 0
            code = "&nbsp;"  ' space '
            Clipboard.SetText "&nbsp;"
        Case 1
            code = "&lt;"    ' < '
            Clipboard.SetText "&lt;"
        Case 2
            code = "&gt;"    ' > '
            Clipboard.SetText "&gt;"
    End Select
    
    With Text2(int_TxtNum)
        .SelText = code & .SelText
    End With
End Sub


Private Sub mSubscript_Click()
    Dim ss As Long
    Dim sl As Long
    
    With Text2(int_TxtNum)
        ss = .SelStart
        sl = .SelLength
        If Right$(.SelText, 1) = " " Then
            ' Double-clicking a word causes a blank space '
            ' next to it to be highlighted too.           '
            .SelText = "<SUB>" & Trim$(.SelText) & "</SUB> "
        Else
            .SelText = "<SUB>" & .SelText & "</SUB>"
        End If
        .SelStart = ss + 3
        .SelLength = sl
    End With
End Sub


Private Sub mSuits2Symbols_Click()
    ' Convert suit letters (S,H,D,C) '
    ' into graphic symbols:          '
    
    Dim t As String
    Dim suit As String
    Dim s As Integer
    Dim n As Integer
    Dim startAt As Long
    Dim card As String
    Dim cardPos As Long
    Dim i As Long
    Dim linkEnd As Long
    Dim linkStart As Long
    Dim hdStart As Long
    Dim hdEnd As Long
    Dim repl As String
    Dim lvl As String
    
    With Text2(1)
        If .SelText = "" Then
            .SelStart = 0
            .SelLength = Len(.Text)
        End If
        
        t = .SelText
    End With
    
    If str_UndoSuits2Symbols = "" Then
        ' Keep original text: '
        str_UndoSuits2Symbols = t
    End If
   
    For s = 1 To 4 ' = number of suits '
        suit = Mid$("CDHS", s, 1)
        For n = 1 To 7 ' = number of bidding levels '
                       '   i.e.: 1C to 7S           '
            ' Jump to title heading: '
            startAt = InStr(t, "<H1")
            If startAt = 0 Then
                startAt = 1
            Else
                startAt = startAt + 6
            End If
            
            card = Mid$("1234567", n, 1) & suit ' e.g.: 1C '
            cardPos = InStr(startAt, t, card)
         
            Do While cardPos < Len(t) And cardPos > 0
              If cardPos > 100 Then
                ' If a link starts before the suit and '
                ' ends after it, don't change the suit '
                ' because it is part of a file name:   '
                ' (We arbitrarily go back 100 chars to '
                ' look for the start of a link.)       '
                linkStart = InStr(cardPos - 100, t, "<A HREF")
                ' There may be more than one link in   '
                ' that span, so look for the last link '
                ' after the one found and before the   '
                ' card position:                       '
                i = InStr(linkStart + 10, t, "<A HREF")
                Do While i > 0 And i < cardPos
                    linkStart = i
                    i = InStr(i + 10, t, "<A HREF")
                    DoEvents
                Loop
                ' If link start found, look for end:  '
                If linkStart > 0 And linkStart < cardPos Then
                    ' See if end is before the card '
                    ' position. If it is, then the  '
                    ' card is not part of a file spec
                    linkEnd = cardPos - 4
                    Do While Mid$(t$, linkEnd, 4) <> "</A>" And _
                             linkEnd > linkStart + 5
                        linkEnd = linkEnd - 1
                        DoEvents
                    Loop
                End If
              End If
              
                If linkStart = 0 Or linkEnd < cardPos Then
                    ' Card position is not inside a link. '
                  If cardPos > 100 Then
                    ' Several different size suit symbols  '
                    ' exist for different sizes of text.   '
                    ' If a heading starts before the suit  '
                    ' position and ends after it, then use '
                    ' the correspondingly sized symbol:    '
                    hdStart = InStr(cardPos - 100, t$, "<H")
                    'Find the last heading start before the card:
                    i = InStr(hdStart + 4, t, "<H")
                    Do While i > 0 And i < cardPos
                        hdStart = i
                        i = InStr(i + 4, t$, "<H")
                        DoEvents
                    Loop
                    ' If a is heading found, look for end: '
                    If hdStart > 0 And hdStart < cardPos Then
                        hdEnd = cardPos - 4
                        Do While Mid$(t$, hdEnd, 3) <> "</H" And _
                                 hdEnd > hdStart + 3
                            hdEnd = hdEnd - 1
                            DoEvents
                        Loop
                    End If
                  End If
                  
                    ' HTML code for using the graphic symbol: '
                    repl = "<IMG SRC=" & qt & suit & "3.png" & qt & ">"
                    
                    ' Use a larger symbol if in a heading: '
                    If (hdEnd > cardPos Or hdEnd = hdStart + 4) And hdEnd > 0 Then
                        lvl = Mid$(t, hdStart + 2, 1)
                        repl = Replace(repl, ".", lvl & ".")
                    End If
                    
                    t = Left$(t, cardPos) & repl & Mid$(t, cardPos + 2)
                
                End If
                
                cardPos = InStr(cardPos + 2, t, card)
                ' Show user what's happening: '
                ' (Is this necessary?) '
                'Text2(1).Text = t
                'Text2(1).Refresh
                DoEvents
            Loop
        Next
    Next
    
    ' Put working variable '
    ' into the edit box.   '
    If Len(t) > 100 Then
        mUndoSuits2Symbols.Enabled = True
    End If
    Text2(1).SelText = t
    t = Empty
    
    ' Save changes so that they '
    ' show up in left window:   '
    mSave_Click 0
    
    MsgBox "Done." & vbCr & vbCr _
           & "Look over the file for problems." & vbCr & vbCr _
           & "Edit source to fix small problems" & vbCr _
           & "such as symbols in links or headings." & vbCr & vbCr _
           & "Click " & qt & "Edit - Undo Suits to Symbols" & qt & vbCr _
           & "to restore the original file."
    
End Sub


Private Sub mRow_Click()
    ' Insert a Row code: '
    With Text2(int_TxtNum)
        .SelText = "<TR>" & .SelText & "</TR>"
    End With
End Sub


Public Sub mSave_Click(index As Integer)
    ' Save edited HTML code or text. '
    ' Usually, this will be to save  '
    ' HTML code, which can only be in'
    ' the right window, but it could '
    ' be to save a text file which,  '
    ' again, should be in the right  '
    ' window, but it is possible that'
    ' there could be plain text in   '
    ' the left window.               '
    
    If Text2(int_TxtNum).Visible = False Then
        MsgBox "You are not editing a file," & vbCr _
             & "so there is nothing to save."
        Exit Sub
    End If
    
    If Trim$(L_Name(1).Text) = "" Then
        If Trim$(str_FiName(0)) <> "" Then
            str_FiPath(1) = str_FiPath(0)
            str_FiName(1) = str_FiName(0)
            L_Name(1) = L_Name(0)
        Else
            index = 1
        End If
    End If
  
    Dim origFiName As String
    Dim ret As Long
    Dim fiNum As Integer
    
    If index = 1 Then ' Index=1 is "Save As..." '
        
        ' Get file name: '
        origFiName = str_FiName(int_TxtNum)
        Do
            CmDia.DialogTitle = "Enter file name."
            CmDia.FileName = str_FiPath(int_TxtNum) _
                           & str_FiName(int_TxtNum)
            CmDia.ShowSave
            If CmDia.FileTitle = "" Then
                Exit Sub
            End If
            
            str_FiName(int_TxtNum) = CmDia.FileTitle
            If Dir$(str_FiName(int_TxtNum)) <> "" Then
                ret = MsgBox("File already exists." & vbCr _
                      & "Overwrite it?", _
                      vbYesNoCancel)
            Else
                ret = vbYes
            End If
            DoEvents
        Loop While ret = vbNo
    
        If LCase$(Right$(str_FiName(int_TxtNum), 3)) = "htm" And _
           L_Name(int_TxtNum) <> "" _
        Then ' Doing "Save-As" leaves the old '
             ' file on the disk. See if the   '
             ' user wants to delete it:       '
            ret = MsgBox("Delete the original file:" & vbCr _
                & str_FiPath(0) & str_FiName(0) & "?" _
                , vbYesNoCancel)
      
            If ret = vbCancel Then
                str_FiName(int_TxtNum) = origFiName
                Exit Sub
      
            ElseIf ret = vbYes Then
                Kill str_FiPath(0) & str_FiName(0)
            End If
        End If
    End If
  
    ' Clean up file name: '
    If Left$(str_FiPath(int_TxtNum), 8) = "file:\\\" Then
        str_FiPath(int_TxtNum) = Mid$(str_FiPath(int_TxtNum), 9)
    End If
    str_FiName(int_TxtNum) = Replace$(str_FiName(int_TxtNum), "#EndOfFile", "")
    
    ' Write the file to disk: '
    fiNum = FreeFile
    If InStr(str_FiName(int_TxtNum), ".") = 0 Then
        str_FiName(int_TxtNum) = str_FiName(0)
        L_Name(int_TxtNum) = str_FiName(int_TxtNum)
    End If
    If str_FiPath(int_TxtNum) <> "" Then
        If Right$(str_FiPath(int_TxtNum), 1) <> "\" Then
            str_FiPath(int_TxtNum) = str_FiPath(int_TxtNum) & "\"
        End If
    End If
    If Len(str_FiName(int_TxtNum)) > 2 Then
        If Left$(str_FiName(int_TxtNum), 1) = "\" Then
            str_FiName(int_TxtNum) = Mid$(str_FiName(int_TxtNum), 2)
        End If
    End If
    Open str_FiPath(int_TxtNum) & str_FiName(int_TxtNum) For Output As #fiNum
        Print #fiNum, Text2(int_TxtNum).Text
    Close #fiNum
    ' Log changes: '
    str_Backup(1) = str_Orig(1)
    str_Orig(int_TxtNum) = Text2(int_TxtNum).Text
  
    brwWebBrowser(0).Refresh
    
    With Text2(1)
        ret = .SelStart
        If ret + 1000 > Len(.Text) And 1 = 2 Then
            ' When saving a change, the following was supposed to '
            ' keep the display where it was, but instead it moved '
            ' it back to the top. '
            brwWebBrowser(0).Navigate str_FiPath(0) & str_FiName(0)
            DoEvents
            .SelStart = Len(.Text)
            .SelStart = ret
        End If
    End With
  
    If index = 0 Then Exit Sub ' 0 = plain Save '
  
    ' Adjust for "Save-As" name change: '
    If L_Name(0) = L_Name(1) Then
        L_Name(0) = str_FiName(int_TxtNum)
        L_Name(1) = str_FiName(int_TxtNum)
        If int_TxtNum = 1 Then
            str_FiName(0) = str_FiName(int_TxtNum)
            brwWebBrowser(0).Navigate str_FiPath(0) & str_FiName(0)
        Else
            str_FiName(1) = str_FiName(int_TxtNum)
            brwWebBrowser(1).Navigate str_FiPath(0) & str_FiName(1)
        End If
        UpdateRecentFiles int_TxtNum
    Else
        L_Name(int_TxtNum) = str_FiName(int_TxtNum)
    End If
End Sub


Private Sub mSpellCheckDocument_Click()
    ' This routine uses my home-grown spell checker.
    ' It is faster to load than using Word's for   '
    ' spell-checking, plus some users may not have '
    ' Word, but mine requires 100 MB data files.   '
    ' If the user doesn't have that, this program  '
    ' will try to use Word's.                      '
    
    ' Spell checking has been disabled for now. (01-10-2021)
    On Local Error Resume Next
    'If Dir$(str_AppPath & "aic Words NEEDS CHECKING.accdb") = "" Then
    '    WordSpellcheckDoc
    'Else
    '    OpenMySpellChecker
    '    Text2(1).SelStart = 0
    '    Text2(1).SelLength = 0
    '    SkipWord 0
    'End If
End Sub
   
   
Private Sub mSpellCheckWord_Click()
    ' Spell-check a selected word '
    ' instead of the whole file:  '
    
    'If Trim$(Text2(1).SelText) = "" Then
    '    MsgBox "Highlight a word to spellcheck first."
    
    'ElseIf Dir$("ai.accdb.") = "" Then
    '    MySpellCheckWord
    
    'Else
    '    MySpellCheckWord
    'End If
End Sub
    
    
Private Sub mSuitSymbol_Click(index As Integer)
    Dim code As String
    Dim i As Integer
    Dim L As Integer
    Dim siz As Integer
    Dim strt As Long
    Dim lvl As Integer
    Dim suit As String
    
    ' This sub will insert HTML for a suit symbol into the file.     '
    ' This can be done using the Special Characters menu or by using '
    ' Ctrl-F3 through Ctrl-F5.                                       '
    ' In this file, suit size 3 is used with the regular text.  This '
    ' may need to be adjusted for use with different text font sizes.'
    code = UCase(Text2(1).SelText)
    strt = Text2(1).SelStart
    
    If code = "" Then      '123456789-123456789-
        code = LTrim$(Mid$("   C  D  H  S HS DC", index * 3, 2))
        Text2(1).SelText = "<IMG SRC=" & qt & code & "3.png" & qt & ">"
        Clipboard.Clear
        Clipboard.SetText "<IMG SRC=" & qt & code & "3.png" & qt & ">"
        Exit Sub
    End If
    
    If Val(Right$(" " & code, 1)) > 0 And Len(code) = 3 Then ' e.g.: DC2 '
        lvl = Val(Right$(code, 1))
        suit = Left$(code, 2)
    End If
              '123456789-123456789-
    i = InStr("   C  D  H  S  HS DC", " " & suit) / 3
    If i > 0 Then
        lvl = lvl + 1
        If lvl = 4 Then lvl = 1
        code = Trim$(suit) & lvl
    Else
        code = "<IMG SRC=" & qt & code & ".png" & qt & ">"
    End If
    
    Text2(1).SelText = code
End Sub

Private Sub mSuperscript_Click()
    Dim ss As Long
    Dim sl As Long
    
    With Text2(int_TxtNum)
        ss = .SelStart
        sl = .SelLength
        If Right$(.SelText, 1) = " " Then
            ' Double-clicking a word causes a blank space '
            ' next to it to be highlighted too.           '
            .SelText = "<SUP>" & Trim$(.SelText) & "</SUP> "
        Else
            .SelText = "<SUP>" & .SelText & "</SUP>"
        End If
        .SelStart = ss + 3
        .SelLength = sl
    End With
End Sub


Private Sub mSwap_Click()
    ' Swap contents of windows: '
    
    Dim p As String
    Dim t As String
    Dim text0vis As Boolean
    Dim Text2vis As Boolean
    
    'Screen.MousePointer = 11
    mSwap.Checked = True
    ' The reason for checking the '
    ' menu item is for use as a   '
    ' flag to keep the browser's  '
    ' BeforeNavigate2 routine from'
    ' executing, which screws up  '
    ' things when swapping files. '
    
    ' Swap name and path: '
    t = str_FiName(1)
    str_FiName(1) = str_FiName(0)
    str_FiName(0) = t
    
    p = str_FiPath(1)
    str_FiPath(1) = str_FiPath(0)
    str_FiPath(0) = p
    
    ' I can't swear that Text2(0)  '
    ' will ever be visible, but if '
    ' it is, then put it into the  '
    ' right text box, not into the '
    ' right browser.               '
    If Text2(0).Visible Then
        text0vis = True
        t = Text2(0)
    End If
  
    L_Name(0).Text = str_FiName(0)
    L_Name(1).Text = str_FiName(1)
    Me.Caption = str_FiName(0) & " - BidBase File Viewer"
      
    If Text2(1).Visible Then
        ' Put right text box into left: '
        Text2(0) = Text2(1)
        Text2(0).Visible = True
        brwWebBrowser(0).Visible = False
    Else
        ' Put right browser into left: '
        brwWebBrowser(0).Navigate str_FiPath(0) & str_FiName(0)
        Text2(0).Visible = False
        brwWebBrowser(0).Visible = True
        ' Wait for the browser   '
        ' navigation to be done: '
        Do While mSwap.Checked
            DoEvents
        Loop
        mSwap.Checked = True
    End If
    'ChangeCaption
  
    If text0vis Then
        ' Put left text box into right: '
        Text2(1) = t
        Text2(1).Visible = True
        brwWebBrowser(1).Visible = False
        Text2vis = True
  
    Else
        ' Put left browser into right: '
        brwWebBrowser(1).Navigate str_FiPath(1) & str_FiName(1)
        Text2(1).Visible = False
        brwWebBrowser(1).Visible = True
        ' Wait for the browser   '
        ' navigation to be done: '
        Do While mSwap.Checked
            DoEvents
        Loop
    End If
    
    Screen.MousePointer = 0
End Sub


Private Sub mSymbolsForHands_Click()
    ' Add suit symbols to hands  '
    ' in user-selected text.     '
    ' (Because finding a hand    '
    ' like K432-AJ53-T874-3 in   '
    ' a text file is not that    '
    ' easy to do, we require the '
    ' user to highlight the hand '
    ' to be converted.)          '
    
    Dim h As String
    Dim h1 As String
    Dim h2 As String
    Dim i As Integer
    
    On Local Error GoTo HandError
    
    ' Set up working variable: '
    h = Trim$(Text2(1).SelText)
    If h = "" Then Exit Sub
    
    ' Set up the code: '
    h1 = "<IMG SRC=" & qt
    h2 = qt & ">"
    ' Start with Spade symbol: '
    h = h1 & "S.gif" & h2 & h
    ' Add Heart symbol: '
    i = InStr(h, "-")
    h = Left$(h, i - 1) & " " & h1 & "h.gif" & h2 & Mid$(h, i + 1)
    ' Diamond symbol: '
    i = InStr(i + 1, h, "-")
    h = Left$(h, i - 1) & " " & h1 & "D.gif" & h2 & Mid$(h, i + 1)
    ' Club symblo: '
    i = InStr(i + 1, h, "-")
    h = Left$(h, i - 1) & " " & h1 & "c.gif" & h2 & Mid$(h, i + 1)
    ' Put working variable back into Text box: '
    Text2(1).SelText = h
    ' Save so that user can see it: '
    mSave_Click 0
    
    On Local Error GoTo 0
Exit Sub
HandError:
    MsgBox "A problem occurred in the hand " & Text2(1).SelText
    On Local Error GoTo 0
End Sub


Private Sub mTable_Click()
    ' Insert a table into the text: '
    Text2(int_TxtNum).SelText = _
        "<TABLE>" & _
         Text2(int_TxtNum).SelText & _
        "</TABLE>"
End Sub

Private Sub mTableOfContents_Click()
    ' Create a Table Of Contents '
    ' from the files Headers.    '
    
    Dim t As String
    Dim t1 As String
    Dim i As Long
    Dim j As Long
    Dim j2 As Long
    Dim j3 As Long
    Dim q As String
    Dim lvl As Integer
    Dim lvlNow As Integer
    Dim hd As String
    Dim alph As String
    Dim ln As String
    Dim subLinkNum As Integer
    Dim t2 As String
    
    q = qt
    alph = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

    ' Set up working variable: '
    t = Text2(1).Text
    
    ' In our files, only the document title '
    ' is in an <H1> header, so by finding   '
    ' the </H1>, we get past the doc title. '
    i = InStr(t, "/H1>")
    If i = 0 Then
        i = InStr(t, "/h1>")
    End If
    If i = 0 Then
        MsgBox "This function requires that the" & vbCr _
               & "document heading be a '<H1>'" & vbCr _
               & "heading. The TOC will be inserted" & vbCr _
               & "after it."
        Exit Sub
    End If
    
    ' Contents heading: '
    i = i + 6 ' Get past the title's Cr-Lf.
    t1 = vbCrLf & "<H2>Contents:</H2>" & vbCrLf _
       & "<UL>" & vbCrLf _
       & " " & vbCrLf _
       & "</UL></UL>" & vbCrLf _
       & "<P><HR>" & vbCrLf
    t = Left(t, i - 1) & t1 & Mid$(t, i)
    
    int_IntLinkNum = 0 ' Internal link number.      '
    lvlNow = 2         ' Level of indentation.      '
    i = i + 26         ' Get past Contents heading. '
    
    ' Index the <LI> entries instead of <H2>:  '
    'j = InStr(i, t, "<LI><B><I>") + 6  ' This <LI> stuff was bombing the sub. See note further down. '
    j = InStr(i, t, "<H2")
    lvl = 2
    subLinkNum = 0
    
GoTo SkipH2 ' This skip was from another version. '
    ' Find next heading (either <H2> or <H3>): '
    j2 = InStr(i, t, "<H2")
    j3 = InStr(i, t, "<H3")
    
    ' If there is a <H3> before  '
    ' an <H2> then put it first: '
    If (j3 < j2 Or j2 = 0) And j3 > 0 Then
        j = j3
        lvl = 3
    Else
        j = j2
        lvl = 2
        subLinkNum = 0
    End If

SkipH2:
    ' Find end of heading: '
    j2 = InStr(j + 5, t, "</H")
    
    Do
        ' Get heading: '
        hd = Mid$(t, j + 4, j2 - j - 4)
        
        'Calculate link ID letter:
        If lvl = 2 Then
            int_IntLinkNum = int_IntLinkNum + 1
        End If
        
        If int_IntLinkNum < 27 Then
            ' First use the 26 letter '
            ' of the alphabet:        '
            ln = Mid$(alph, int_IntLinkNum, 1)
        Else ' Then start AA, AB, etc.'
            ln = "A" & Mid$(alph, int_IntLinkNum - 26, 1)
        End If
        
        ' All <H3>'s under the same <H2> '
        ' get incremented links, such as '
        ' when <H2> link is "F", then the'
        ' <H3> links under it are F1, F2,'
        ' F3, etc.                       '
        If lvl = 3 Then
            subLinkNum = subLinkNum + 1
            ln = ln & subLinkNum
        End If
        
        ' Insert target ID code first:  '
        t2 = "<A NAME=" & q & ln & q & "></A>"
        t = Left(t, j + 3) & t2 & Mid$(t, j + 4)
        
        ' Insert hyperlink code in TOC: '
        t1 = "<A HREF=#" & ln & ">"
        If lvl = 2 And lvlNow = 3 Then
            t1 = "</UL>" & vbCrLf & t1
            lvlNow = 2
        ElseIf lvl = 3 And lvlNow = 2 Then
            t1 = "<UL>" & vbCrLf & t1
            lvlNow = 3
        End If
        t1 = t1 & hd$ & "</A><BR>" & vbCrLf
        t = Left(t, i) & t1 & Mid$(t, i + 1)
        
        Text2(1).Text = t
        
        'Increment pointers for next TOC line:
        i = i + Len(t1)
        j = j + Len(t1)
        If i > Len(t) Or j > Len(t) Then Exit Do
        
        ' 12-29-18:  This sub was bombing because of the following '
        '            code about finding <LI>.  I have no idea what '
        '            that is for but I commented it out and the    '
        '            sub now works. I'm leaving it in just in case '
        '            whatever reason put it in before comes again. '
        'Find next <LI>: '
        'j = InStr(j + 1, t, "<LI><B><I>") + 6
        'j2 = InStr(j + 5, t, "</I></B>")
'GoTo SkipH2Again ' Also from another version. '
        'Find next heading (either <H2> or <H3>):
        j2 = InStr(j + 1, t, "<H2")
        j3 = InStr(j + 1, t, "<H3")
    
        If (j3 < j2 Or j2 = 0) And j3 > 0 Then
            j = j3
            lvl = 3
        Else
            j = j2
            lvl = 2
            subLinkNum = 0
        End If
        j2 = InStr(j + 5, t, "</H")
SkipH2Again:
        DoEvents
    Loop While j > 0 And j2 > 0
    Text2(1).Text = t
    mSave_Click 0
End Sub


Private Sub mTableSpecsWindow_Click()
    TableSpecs.Show 1, Me
End Sub


Private Sub mUnderline_Click()
    ' HTML to underline selected text: '
    Text2(int_TxtNum).SelText = _
        "<U>" _
         & Text2(int_TxtNum).SelText & _
        "</U>"
End Sub


Private Sub mUndoSuits2Symbols_Click()
    ' When suits are converted to    '
    ' symbols, the original text is  '
    ' saved to str_UndoSuits2Symbols '
    ' so that we can restore it:     '
    If str_UndoSuits2Symbols = "" Then
        MsgBox "The Suits2Symbols Undo has nothing in it." & vbCr & "Are you sure you didn't Make Symbols For Hands?"
        Exit Sub
    End If
    Text2(1).Text = str_UndoSuits2Symbols
    ' Save and show the user: '
    Dim x As Integer
    x = MsgBox("To save the Undo, click Ok.", vbOKCancel, "Undo Suits-To-Symbols")
    If x = vbCancel Then Exit Sub
    mSave_Click 0
    ' Free up the memory used by the '
    ' variable:                      '
    str_UndoSuits2Symbols = Empty
    ' Nothing left to undo, '
    ' so gray out the menu. '
    mUndoSuits2Symbols.Enabled = False
End Sub


Private Sub mWidth_Click()
    ' Insert HTML "Width" code: '
    Text2(int_TxtNum).SelText = " WIDTH=30" & Text2(int_TxtNum).SelText
End Sub


Private Sub mWrapAround_Click()
    Text2(1).SelText = "<img src=" & qt & "(Image file name)" & qt & _
                       " height=" & qt & "xxx" & qt & _
                       " width=" & qt & "xxx" & qt & _
                       " alt=" & qt & "Image" & qt & " style=" & qt & _
                       "float: right; margin: 4px 10px 0px 10px; border: 1px solid #000000;" & qt & ">" & _
                       Text2(1).SelText & "<div style=""clear:both;"">"
End Sub


Private Sub mYellowHighlighter_Click()
    ' Alternative: '
    Text2(1).SelText = _
              "<span style=" & qt & "background-COLOR: #FFFF00" & qt & ">" & _
              Text2(1).SelText & "</span>"
End Sub


Private Sub mZoom_Click(index As Integer)
    Dim x As Integer
    
    With Text2(int_TxtNum)
    ''''''''''''''''''''''
    If index = 0 Then
        x = MsgBox("This will set all images to zoom +50%" & vbCr & _
                   "when the mouse cursor hovers them." & vbCr & _
                   "Continue?", vbYesNo)
        If x = vbNo Then Exit Sub
        x = InStr(LCase(Left$(.Text, 300)), "img:hover")
        If x > 0 Then
            MsgBox "This option has already been set."
            Exit Sub
        End If
        x = InStr(LCase(Left$(.Text, 800)), "</head>")
        If x = 0 Then
            MsgBox "This file does not have an HTML header."
            Exit Sub
        End If
        Text2(int_TxtNum) = Left$(.Text, x - 1) & Chr$(13) & _
                            "<style>" & Chr$(13) & _
                            "img {" & Chr$(13) & _
                            "    transition: -webkit-transform 0.25s ease;" & Chr$(13) & _
                            "    transition: transform 0.25s ease;" & Chr$(13) & _
                            "}" & Chr$(13) & _
                            "img:hover {" & Chr$(13) & _
                            "    -webkit-transform: scale(1.25);" & Chr$(13) & _
                            "    transform: scale(1.25);" & Chr$(13) & _
                            "}" & Chr$(13) & _
                            "</STYLE>" & Chr$(13) & _
                            Mid$(.Text, x)
        MsgBox "Source code must be saved."
    Else
        x = MsgBox("This will zoom in on an image" & vbCr & _
                   "when the cursor moves over it." & vbCr & _
                   "and will un-zoom when the mouse" & vbCr & _
                   "moves away or the image is clicked." & vbCr & _
                   "Continue?", vbYesNo)
        If x = vbNo Then Exit Sub
        
    Dim i As Long
    Dim sel As String
    Dim s As String
    Dim zoomID As Long
    Dim selStart1 As Long
    Dim selLength1 As Long
        
    selStart1 = Text2(int_TxtNum).SelStart
    selLength1 = Text2(int_TxtNum).SelLength
    sel = Trim$(Text2(int_TxtNum).SelText)
    
    If sel = "" Then
        sel = "<P><B><I>Put the hidden text, etc., here.</I></B>"
    End If
    
    ' Get the highest zoomID used.
    i = 200
    zoomID = 0
    Do
        zoomID = zoomID + 1
        i = InStr(i + 10, Text2(1).Text, "small" & zoomID)
        If i = 0 Then Exit Do ' No entry found using this zoomID, so stop looking. '
        DoEvents
    Loop While i > 0
    
        s = "" & _
            "<SPAN ID=" & qt & "small" & zoomID & qt & " STYLE=''" & qt & "> " & vbCr & _
            "    <SPAN onmouseover=" & qt & "small" & zoomID & ".style.display='none'" & qt & ">" & vbCr & _
            "        <SPAN onmouseover=" & qt & "zoom" & zoomID & ".style.display=''" & qt & ">" & vbCr & _
            "            Put the first image wrap code here with smaller dimensions." & vbCr & _
            "        </SPAN>" & vbCr & _
            "    </SPAN>" & vbCr & _
            "</SPAN>" & vbCr & vbCr & _
            "<SPAN ID=" & qt & "zoom" & zoomID & qt & " STYLE='display:none'" & qt & ">" & vbCr & _
            "    <SPAN onmouseout=" & qt & "small" & zoomID & ".style.display=''" & qt & ">" & vbCr & _
            "        <SPAN onmouseout=" & qt & "zoom" & zoomID & ".style.display='none'" & qt & ">" & vbCr & _
            "            Put the second image wrap code here with larger dimensions." & vbCr & _
            "        </SPAN>" & vbCr & _
            "    </SPAN>" & vbCr & _
            "</SPAN>"
            Text2(int_TxtNum).SelText = s
            
        MsgBox "The first image link should look like this:" & vbCr & _
               "  <img src='image.jpg' width='400' height='300' " & vbCr & _
               "  alt='Image' style='float: right;" & vbCr & "  margin: 0px 0px 0px 0px;" & vbCr & "  border: 1px solid #000000;'>" & vbCr & _
               "The dimensions of the second link should be about 25% greater (500 and 375," & _
               "in this case) though you can adjust it to taste." & _
               "This puts a lot more code in your web page than setting zoom for ALL images," & _
               "but it gives you more control over individual images. In fact, you can use a different image file if you want to swap pictures. You can also use OnClick instead of OnMouseOver. "

    End If
    ''''''''
    End With
End Sub


Private Sub pSplitterBar_MouseDown(Button As Integer, _
                                   Shift As Integer, _
                                   x As Single, y As Single)
    ' Remember starting location '
    ' for pSplitterBar_MouseMove.'
    If Button = 1 Then pSplitterBar.Tag = x
End Sub


Private Sub pSplitterBar_MouseMove( _
                Button As Integer, _
                Shift As Integer, _
                x As Single, _
                y As Single)
  
    ' Don't let the Form_Resize code do '
    ' its thing yet:                    '
    bln_NoResize = True
    ' Show the left-right arrow cursor: '
    pSplitterBar.MousePointer = vbSizeWE
    
    If Button = 1 Then
        ' There is a Text box and browser on each side. '
        ' Only one of these is visible on each side,    '
        ' depending on the type of file being displayed.'
        ' Normally, the browser will be on the left and '
        ' if editing HTML, the Text box will be on the  '
        ' right. However, to make it easier to resize   '
        ' them and to switch from Browser to Text box,  '
        ' all of them are resized whether visible or not.
        If (Text2(0).Width + (x - Val(pSplitterBar.Tag))) > 1000 And _
           (Text2(1).Width - (x - Val(pSplitterBar.Tag))) > 1000 _
        Then
            pSplitterBar.Left = _
                 pSplitterBar.Left _
                 + x - Val(pSplitterBar.Tag)
            
            Text2(0).Width = _
                Text2(0).Width _
                + x - Val(pSplitterBar.Tag)
                
            Text2(1).Left = _
                Text2(1).Left _
                + x - Val(pSplitterBar.Tag)
                
            Text2(1).Width = _
                Text2(1).Width _
                - (x - Val(pSplitterBar.Tag))
        End If
    End If
  
    brwWebBrowser(0).Width = Text2(0).Width
    brwWebBrowser(1).Width = Text2(1).Width
    brwWebBrowser(1).Left = Text2(1).Left
    
    L_Name(1).Left = Text2(1).Left
    t_Back.Left = L_Name(1).Left - 150 - t_Back.Width
  
    lng_Panel1Width = Text2(0).Width
    lng_Panel2Width = Text2(1).Width
    
    bln_NoResize = False
End Sub


Private Sub t_Back_Click()
    brwWebBrowser(int_AlreadyPicked).GoBack
    brwWebBrowser(0).SetFocus
End Sub


Private Sub Text2_Click(index As Integer)
    ' Keep track of last window '
    ' in which the user clicked:'
    int_TxtNum = index
End Sub


Private Sub Text2_GotFocus(index As Integer)
    ' Keep track of last window clicked: '
    int_TxtNum = index
End Sub


Private Sub Text2_KeyDown(index As Integer, _
                          KeyCode As Integer, _
                          Shift As Integer)
    Dim c As String
    Dim i As Long
    Dim j As Long
    Dim K As Long
    Dim w As String
    Dim w2 As String
    
    c = LCase$(Chr$(KeyCode))
    
    If KeyCode = 69 And Shift = 2 Then
        ' Trap Ctrl-E, which is supposed to trigger the  '
        ' External Link menu option, but which the Rich  '
        ' Text box interprets as a command to center text.
        KeyCode = 0
        Shift = 0
        mLinkExt_Click
    ElseIf KeyCode = 77 And Shift = 2 Then ' Ctrl-M '
        KeyCode = 0
        Shift = 0
        mSymbolsForHands_Click
    ElseIf KeyCode = 27 And Shift = 0 Then 'Esc
        If frmFindReplace.Visible Then
            Unload frmFindReplace
        End If
    ElseIf KeyCode = 46 And Shift = 0 Then ' Del '
        Exit Sub
    ElseIf InStr("abcdefghijklmnopqrstuvwxyz'-", c) = 0 And Shift = 0 Then
        ' Non-alpha character. '
        If KeyCode = 13 Or KeyCode = 8 Then Exit Sub
        If KeyCode = 37 Then ' left arrow '
            Exit Sub
        End If
        
        Exit Sub ' Spell checking is disabled for now. '
        
        i = Text2(index).SelStart
        If i = 0 Then Exit Sub
        c = LCase$(Mid$(Text2(index).Text, i, 1))
        If InStr("abcdefghijklmnopqrstuvwxyz", c) = 0 Then Exit Sub ' No word on left. '
        
        ' Look for the end if the word if other than space key pressed: '
        If KeyCode = 38 Or KeyCode = 40 Then
            Do While InStr("abcdefghijklmnopqrstuvwxyz'-", LCase$(Mid$(Text2(index).Text, i + 1, 1))) > 0 And i < 40
                i = i + 1
                DoEvents
            Loop
        End If
        
        ' Look for beginning of the word: '
        j = i - 1
        Do While InStr("abcdefghijklmnopqrstuvwxyz'-", LCase$(Mid$(Text2(index).Text, j, 1))) > 0 And j > 1 And j < 40
            j = j - 1
            DoEvents
        Loop
        
        w = Mid$(Text2(index).Text, j + 1, i - j)
        
        'Set CortexWS = DBEngine.Workspaces(0)
        'Set CortexDB = CortexWS.OpenDatabase("C:\Users\Public\WebPages\AIC\ai.accdb")
        'Set CortexRS = CortexDB.OpenRecordset("Cortex")
        ''Set WordsRS = CortexDB.OpenRecordset("Words")
        'Set AICWordsWS = DBEngine.Workspaces(0)
        'Set AICWordsDB = AICWordsWS.OpenDatabase("C:\Users\Public\WebPages\AIC\aic words.accdb")
        'Set WordsRS = AICWordsDB.OpenRecordset("Words")
        
        'With WordsRS
        '    .index = "Text"
        '    .Seek "=", w
        '    If .NoMatch Then
        '        If Right$(w, 2) = "'s" Then
        '            w = Left$(w, Len(w) - 2)
        '            .Seek "=", w
        '        End If
        '        If .NoMatch Then
        '            ' Check hyphenated words: '
        '            K = InStr(w, "-")
        '            If K > 0 Then
        '                w2 = Mid$(w, K + 1)
        '                w = Left$(w, K - 1)
        '                .Seek "=", w
        '                If Not .NoMatch Then .Seek "=", w2
        '            End If
        '        End If
        '    End If
        '    ' Case is ignored. '
        '    Text2(index).SelStart = j
        '    Text2(index).SelLength = i - j
        '    If .NoMatch Then
        '        Text2(index).SelColor = vbRed
        '        Text2(index).SelBold = True
        '        Text2(index).SelStart = i
        '        Text2(index).SelLength = 1
        '    Else
        '        ' Make sure the text in the Words table is good. '
        '        CortexRS.index = "WordID"
        '        CortexRS.Seek "=", !Id
        '        If Not CortexRS.NoMatch Then
        '            Do While CortexRS!WordID = !Id And (CortexRS!LinkID < 30000 Or CortexRS!LinkID > 30210)
        '                CortexRS.MoveNext
        '                If CortexRS.EOF Then
        '                    CortexRS.Seek "=", !Id
        '                    Exit Do
        '                End If
        '                DoEvents
        '            Loop
        '            If CortexRS!WordID <> !Id Or CortexRS!LinkID < 30000 Or CortexRS!LinkID > 30210 Then
        '                ' Word is in the Words table, '
        '                ' but is not in the Cortex.   '
        '                Text2(index).SelColor = vbBlue
        '                Text2(index).SelBold = True
        '                Text2(index).SelStart = i
        '                Text2(index).SelLength = 1
        '            End If
        '        End If
        '    End If
        '    ' The following stops the red, bold text '
        '    ' if no match was found and it puts the  '
        '    ' text in black if it was found. This is '
        '    ' needed if the user fixes a typo by hand'
        '    ' instead of using the spell corrector.  '
        '    Text2(index).SelColor = vbBlack
        '    Text2(index).SelBold = False
        '    Text2(index).SelStart = i
        '    Text2(index).SelLength = 0
        'End With
            
    End If
End Sub


Private Sub Text2_MouseDown(index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            x As Single, y As Single)
    ' PopupMenu is a VB method. See VB Help file.      '
    ' mH is the name of the HTML menu on the menu bar. '
    ' The below command says to display the HTML menu  '
    ' in a floating, pop-up menu at the cursor location.
    If Button = 2 Then
        PopupMenu mH
    End If
End Sub


Private Sub GetCommandLine()
    ' File path\names can be passed to FileViewer '
    ' on the command line or by dragging files to '
    ' the FileViewer icon.                        '
    ' If a file name is double-clicked in a file  '
    ' manager and BidBase FileViewer is the       '
    ' default for loading that file type then the '
    ' command line will look something like this: '
    ' C:\Programming\VB98\Bridge\index.htm        '
    ' If FileViewer is run by double-clicking on  '
    ' the program's directory listing, Command()  '
    ' will be blank and the FileViewer's doc file '
    ' will be loaded
  
    Dim cmdLine As String
    Dim cmdLnLen As Integer
    Dim i As Integer
    Dim p As String
    Dim n As Integer
    Dim x As Integer
    Dim F As String
    
    qt = Chr$(34)
    DoEvents
    'Get command line arguments.
    cmdLine = Trim$(Command())
    str_CmdLine = cmdLine ' Keep track of original command line, if any. '
    'cmdLine = "C:\Programming\VB98\Bridge\Notes\Kokish_game_tries.htm"
    If DebugMode Then MsgBox Dir$(cmdLine) ' This is for debugging use. '
    If cmdLine = "" Then Exit Sub

    If InStr(cmdLine, "http://www.google.com/search") > 0 Then
        str_FiName(0) = cmdLine
        str_FiPath(0) = ""
        Exit Sub
    End If
  
    If Left$(cmdLine, 1) = qt Then
        ' In WinXP, if you drag a file to the   '
        ' FileViewer icon on the desktop, WinXP '
        ' puts quotes around the path-file name.'
        ' So we take them off here:             '
        cmdLine = Mid$(cmdLine, 2, Len(cmdLine) - 2)
        i = InStr(cmdLine, qt)
        Do While i > 0
            cmdLine = Left$(cmdLine, i - 1) _
                    & Mid$(cmdLine, i + 1)
            i = InStr(i, cmdLine, qt)
            DoEvents
        Loop
    End If
  
    cmdLnLen = Len(cmdLine)
    If cmdLnLen = 0 Then Exit Sub
  
    ' Extract path and file name(s): '
    n = 0
If Left$(cmdLine, 7) <> "http://" Then
    Do
        x = InStr(LCase$(cmdLine), ".htm")
        If x < 1 Then
            x = InStr(LCase$(cmdLine), ".txt")
            If x = 0 Then Exit Sub
        End If
    
        If x < 2 Then
            F$ = cmdLine
            cmdLine = ""
        ElseIf InStr(x + 4, cmdLine, " ") = 0 And _
               Mid$(cmdLine, x + 4, 1) = "#" _
        Then
            F$ = cmdLine
            cmdLine = ""
        
        Else ' Extract first path\file, if more than 1:
            x = InStr(x + 4, cmdLine, ", ")
            If x > 0 Then
                F = Left$(cmdLine, x - 1)
                cmdLine = Trim$(Mid$(cmdLine, x + 1))
            Else
                F = cmdLine
                cmdLine = ""
            End If
        End If
    
        'Find last "\" to separate path from filename:
        i = InStrRev(F, "\")

        If i > 1 Then
            str_FiName(n) = Mid$(F, i + 1)
            str_FiPath(n) = Left$(F, i)
        Else
            str_FiName(n) = F
            str_FiPath(n) = ""
        End If
    
        L_Name(n) = str_FiName(n)
    
        n = n + 1
        DoEvents
    Loop While cmdLine <> ""

Else
    str_FiName(n) = cmdLine
    L_Name(n) = cmdLine
End If

End Sub


Private Sub LoadDocs()
        Dim p As String
        
        If Dir$(str_AppPath & "Notes\..") <> "" Then
            str_FiPath(0) = str_AppPath & "Notes\"
        Else
            str_FiPath(0) = str_AppPath ' Note: str_AppPath ends with "\". '
        End If
        
        If Dir$(str_FiPath(0) & "CardShark_BidBase_File_Viewer.htm") <> "" Then
            str_FiName(0) = "CardShark_BidBase_File_Viewer.htm"
        Else
                Open str_FiPath(0) & "temp98765.htm" For Output As #1
                Print #1, "The Help file (CardShark_BidBase_File_Viewer.htm) was<BR>"
                Print #1, "not found in the folder where this program is located.<P>"
                Print #1, "If it is not on your system, it can be found at<BR>"
                Print #1, "<A HREF="
                Print #1, Chr$(34); "http://www.aeyec.com\BidBase\CardShark_BidBase_File_Viewer.htm";
                Print #1, Chr$(34); ">www.aeyec.com/BidBase</A>"
                Close #1
            str_FiName(0) = "temp98765.htm"
        End If
        
        On Local Error GoTo ErrorLoadingDocs
        L_Name(0) = str_FiPath(0) & str_FiName(0)
        brwWebBrowser(0).Navigate str_FiPath(0) & str_FiName(0)
        On Local Error GoTo 0
        
    Exit Sub
    
ErrorLoadingDocs:
    MsgBox "Error loading the file:" & vbCrLf & _
            str_FiPath(0) & str_FiName(0)
End Sub


Private Sub LoadFiles()
    Dim F As String
    Dim i As Integer
    Dim lnk As String
    Dim b As Boolean
    
    If str_FiName(0) = "" Then
        ' Normally, when the file viewer   '
        ' was started without a file name. '
        LoadDocs
        bln_NoResize = False
        Form_Resize
        GoTo ExitWithoutMenuUpdate
    End If
  
    On Local Error Resume Next
    F = str_FiName(0)
    
    ' Look for internal link in file spec: '
    i = InStr(Len(F) - 6, F, "#")
    If i > 0 Then
        lnk = Mid$(F, i)
        F = Left$(F, i - 1)
    End If
    
    ' Add a slash to path if not already there: '
    If Right$(str_FiPath(0), 1) <> "\" And _
       Left$(str_FiName(0), 7) <> "http://" _
    Then
        str_FiPath(0) = str_FiPath(0) & "\"
    End If
    
    If InStr(str_FiName(0), "google") = 0 And _
       Left$(str_FiName(0), 7) <> "http://" _
    Then ' Check to see if local path and filename are valid: '
        If Dir$(str_FiPath(0) & F) = "" Then
            If Dir$(str_AppPath & "Notes\..") <> "" Then
                If Dir$(str_AppPath & "notes\" & F) = "" Then
                    MsgBox "The following path or file name is invalid:" & vbCr & _
                            str_FiPath(0) & F
                    ' Load doc file, unless that was what was being loaded. '
                    If F <> "CardShark_BidBase_File_Viewer.htm" Then
                        LoadDocs
                    End If
                
                    GoTo ExitWithoutMenuUpdate
                End If
            End If
        End If
    End If
    
    If Left$(str_FiName(0), 7) = "http://" Then str_FiPath(0) = ""
    
    ' This is where the file is loaded: '
    brwWebBrowser(0).Navigate str_FiPath(0) & str_FiName(0)
    
    On Local Error GoTo 0
    
    UpdateRecentFiles 0
    L_Name(0).Text = str_FiName(0)
    Me.Caption = str_FiName(0) & " - BidBase File Viewer"
    
    'Load second file, if any:
    bln_NoResize = True
    If str_FiName(1) <> "" Then
        If Dir$(str_FiPath(1) & str_FiName(1)) = "" Then
            MsgBox "The following path and/or file name is invalid:" & vbCr & _
                   str_FiPath(1) & str_FiName(1)
            GoTo ExitWithoutMenuUpdate
        End If
        
        If lng_File2ViewerWidth > 0 Then
            Me.Width = lng_File2ViewerWidth
        End If
    
        If Me.Left < -200 Then
            Me.Width = lng_MonPosition(3, 1)
        ElseIf Me.Left < lng_MonPosition(1, 3) - 200 Then
            Me.Width = lng_MonPosition(3, 2)
        Else
            Me.Width = lng_MonPosition(3, 3)
        End If
    
        If LCase$(Right$(str_FiName(1), 3)) = "txt" Then
            Text2(1).Visible = True
            Text2(1).LoadFile str_FiPath(1) & str_FiName(1)
            str_Orig(1) = Text2(1).Text 'For testing for changes.
    
        ElseIf str_FiPath(0) & str_FiName(0) <> str_FiPath(1) & str_FiName(1) Then
            ' Not the same filespec for edit'
            ' Not a text file; use browser: '
            brwWebBrowser(1).Visible = True
            ' Browse the file: '
            brwWebBrowser(1).Navigate str_FiPath(1) & str_FiName(1)
            
            If lng_Panel1Width > lng_FileViewerWidth \ 2 Or _
               lng_Panel1Width = 0 _
            Then
                lng_Panel1Width = lng_FileViewerWidth \ 2 - 315
                lng_Panel2Width = lng_FileViewerWidth - lng_Panel1Width - 265
            End If
        
        Else
            L_Name(1) = ""
            mEditHTML_Click
        
        End If
        L_Name(1).Text = str_FiName(1)
        L_Name(1).Visible = True
    Else
        If lng_File1ViewerWidth > 0 Then
            Me.Width = lng_File1ViewerWidth
        End If
    End If
    
ExitLoadFiles:
    b = (L_Name(0) = L_Name(1))
    SetMenuStatus b
ExitWithoutMenuUpdate:
    bln_NoResize = False
End Sub


Private Sub LoadSetUp()
    Dim i As Integer
    Dim fiNum As Integer
    Dim x As String
       
    If Dir$(str_AppPath & "BidBaseFileViewer.cfg") <> "" Then
        fiNum = FreeFile
        On Local Error GoTo LoadSetUp_Error
        Open str_AppPath & "BidBaseFileViewer.cfg" For Input As #fiNum
            Input #fiNum, lng_FileViewerLeft
            If lng_FileViewerLeft < 0 Then lng_FileViewerLeft = 0
            Input #fiNum, lng_FileViewerTop
            If lng_FileViewerTop < 0 Then lng_FileViewerTop = 0
            Input #fiNum, lng_FileViewerHeight
            Input #fiNum, lng_FileViewerWidth
            Input #fiNum, lng_File1ViewerWidth
            Input #fiNum, lng_File2ViewerWidth
            Input #fiNum, lng_Panel1Width
            Input #fiNum, lng_Panel2Width
            Input #fiNum, lng_PanelHeight
            For i = 1 To 10
                Line Input #fiNum, str_RecentFiles(0, i)
                Line Input #fiNum, str_RecentFiles(1, i)
                If str_RecentFiles(1, i) <> "" Then
                    mRecent(i).Visible = True
                    mRecent(i).Caption = str_RecentFiles(1, i) & " (" & str_RecentFiles(0, i) & ")"
                    If str_RecentFiles(0, i) = "" Then
                        str_RecentFiles(0, i) = "C:\"
                    End If
                End If
            Next
            mRecent(0).Visible = mRecent(1).Visible
            Input #fiNum, x
        Close #fiNum
        
        If lng_File1ViewerWidth > 20000 Then
            lng_File1ViewerWidth = lng_FileViewerWidth / 2
            lng_Panel1Width = lng_File1ViewerWidth - 580
        End If
        bln_OpenInRightWindow = (x = "False")
        mOpenRight_Click
    End If

LoadSetUp_Resume:
    Close #fiNum
    ' Default colors: '
    For i = 0 To 4
        If str_FormColors(i) = "" Then
            Select Case i
                Case 1
                    str_FormColors(i) = "#000000"
                Case 2
                    str_FormColors(i) = "#FFFFFF"
                Case 3
                    str_FormColors(i) = "#0000EE"
                Case 4
                    str_FormColors(i) = "#008000"
                Case 5
                    str_FormColors(i) = "#FF0000"
            End Select
        End If
    Next
    bln_NoResize = False
Exit Sub

LoadSetUp_Error:
    Resume LoadSetUp_Resume
End Sub


Private Sub UpdateRecentFiles(ByVal n)
    Dim i As Integer
    Dim alreadyListed As Integer
    Dim F As String
    Dim e As String
    Dim j As Integer
'MsgBox "Ck.Pt. UpdateRecentFiles"
    F = str_FiName(n)
    If Len(F) < 5 Then Exit Sub
    
    ' Remove link, if any: '
    i = InStrRev(F, "#")
    If i > 0 Then
        F = Left$(F, i - 1)
    End If
    
    ' Check for file name already in list: '
    For i = 1 To 10
        If str_RecentFiles(1, i) = F Then
            alreadyListed = i
            Exit For
        End If
    Next
    
    If alreadyListed > 0 Then
        ' If already in list, move up to top '
        ' by moving all files down which are '
        ' higher up in the list:             '
        For i = alreadyListed To 2 Step -1
            str_RecentFiles(0, i) = str_RecentFiles(0, i - 1)
            str_RecentFiles(1, i) = str_RecentFiles(1, i - 1)
        Next
    Else
        ' If not in list, move all files down. '
        ' If already 10 in list, #10 will drop '
        ' off the list (by being overwritten). '
        For i = 10 To 2 Step -1
            str_RecentFiles(0, i) = str_RecentFiles(0, i - 1)
            str_RecentFiles(1, i) = str_RecentFiles(1, i - 1)
        Next
    End If
    
    ' Save the file path & name at top of list: '
    str_RecentFiles(0, 1) = str_FiPath(n)
    str_RecentFiles(1, 1) = F
    
    ' Rewrite "Recent Files" in File menu: '
    mRecent(0).Visible = True
    For i = 1 To 10
        If str_RecentFiles(1, i) <> "" Then
            mRecent(i).Caption = str_RecentFiles(1, i) & " (" & str_RecentFiles(0, i) & ")"
            mRecent(i).Visible = True
        End If
    Next
    If n = 0 Then
        'ChangeCaption
    End If
End Sub


Private Sub ChangeCaption()
    Dim c As String
    Dim i As Integer
    
    c = L_Name(0)
    If Left$(c & Space$(18), 18) = "BidBase_" Then
        c = Trim$(Mid$(c, 19))
        i = InStr(LCase$(c), ".htm")
        If i < 2 Then i = InStr(LCase$(c), ".txt")
        If i > 1 Then c = Left$(c, i - 1)
        c = Replace(c, "_", " ", 1, 99)
    ElseIf Left$(LCase$(c) & Space$(7), 7) = "http://" Then
        c = Mid$(c, 8)
    End If
    Me.Caption = str_FiName(0) & " - BidBase File Viewer"
End Sub


Public Sub SlowItDown(n)
    ' Slows down sliding in/out '
    ' of the spell-check window.'
    Dim t1 As Double
    Dim t2 As Double
    
    t1 = Timer
    Do
        DoEvents
        t2 = Timer
    Loop While t2 - t1 < 0.02
End Sub


Private Sub FixFileName()
    Dim i As Integer
    
    If Left$(L_Name(0), 8) = "file:///" Then
        L_Name(0) = Mid$(L_Name(0), 9)
    ElseIf Left$(str_FiName(0), 8) = "file:///" Then
        str_FiName(0) = Mid$(str_FiName(0), 9)
        L_Name(0) = str_FiName(0)
    End If
    
    i = InStrRev(L_Name(0), "/")
    If i = 0 Then
        i = InStrRev(L_Name(0), "\")
    End If
    If i > 0 Then
        str_FiName(0) = Mid$(L_Name(0), i + 1)
        str_FiPath(0) = Left$(L_Name(0), i)
        L_Name(0) = str_FiName(0)
    End If
    
    i = InStr(str_FiName(0), "#")
    If i > 0 Then
        L_Name(0) = Left$(str_FiName(0), i - 1)
        str_FiName(0) = L_Name(0)
    ElseIf L_Name(0) <> str_FiName(0) Then
        L_Name(0) = str_FiName(0)
    End If
        
    ' Get rid of the trailing slash.'
    If Right$(str_FiName(0), 1) = "/" Then
        str_FiName(0) = _
            Left$(str_FiName(0), _
                  Len(str_FiName(0)) - 1)
    End If

    ' Check for an internal link: '
    ' (e.g.: myfile.htm#A1)       '
    'i = Len(str_FiName(0))
    'Do While Mid$(str_FiName(0), i, 1) <> "." And _
    '         Mid$(str_FiName(0), i, 1) <> "#" And _
    '         i > 1
    '    i = i - 1
    'Loop
    'If i > 1 And Mid$(str_FiName(0), i, 1) = "#" Then
    '    L_Name(0) = Left$(str_FiName(0), i - 1)
    'End If
End Sub

