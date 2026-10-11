VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblMonitors 
      Caption         =   "Label1"
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'1. Start a new project, set the border style of Form1 to Fixed Single.
'2. Place a Label on the form. Set its name to lblMonitors and make it a control array by setting its Index to 0.
'3. Insert the following code in Form1.__

Private Sub Form_Load()
    EnumMonitors Me
    LoadPosition hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SavePosition hwnd
End Sub

' 4. Add a module and insert the following code and test the program with different monitor configurations.

