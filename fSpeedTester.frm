VERSION 5.00
Begin VB.Form fSpeedTester 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SpeedTester"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "fSpeedTester.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHTML 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   3600
      Width           =   4815
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Height          =   375
      Left            =   3840
      TabIndex        =   20
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtRepeats 
      Height          =   285
      Left            =   3960
      TabIndex        =   16
      Text            =   "1000"
      Top             =   2280
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.CheckBox chkToClip 
         Caption         =   "Auto HTML to clipboard"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox txtProcCount 
         Height          =   285
         Left            =   3840
         TabIndex        =   23
         Text            =   "5"
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtGroup 
         Height          =   285
         Left            =   1080
         TabIndex        =   19
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblProcCount 
         Caption         =   "Procedures to Test"
         Height          =   255
         Left            =   2400
         TabIndex        =   22
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Group Name"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4680
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label lblRepeats 
         Caption         =   "Repeats"
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblName 
         Caption         =   "Procedure 5:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblTime 
         Caption         =   "Time 5"
         Height          =   255
         Index           =   4
         Left            =   3600
         TabIndex        =   14
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblBar 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   13
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblBar 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblBar 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblBar 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblBar 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblTime 
         Caption         =   "Time 4"
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblTime 
         Caption         =   "Time 3"
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblTime 
         Caption         =   "Time 2"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblTime 
         Caption         =   "Time 1"
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblName 
         Caption         =   "Procedure 4:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblName 
         Caption         =   "Procedure 3:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblName 
         Caption         =   "Procedure 2:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblName 
         Caption         =   "Procedure 1:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Results as HTML"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3360
      Width           =   1335
   End
End
Attribute VB_Name = "fSpeedTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' file    : fSpeedTester.frm
' revised : 2003-10-31
' author  : redbird77
' email   : redbird77@earthlink.net
' www     : http://home.earthlink.net/~redbird77

' about   : see Related Document - README.txt

' idea : since one would probably code as many procedures as they would like to test
' nix the ProcToTest textbox and just run as many procedures as coded

Option Explicit

Private G As tGroup

Private Sub cmdRun_Click()

    SpeedTester_Init G, txtGroup.Text, Val(txtProcCount.Text), Val(txtRepeats.Text)
    SpeedTester_Run G
    SpeedTester_Graph G, lblBar, lblTime
    
    txtHTML.Text = SpeedTester_ToHTML(G, CBool(chkToClip.Value))

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set fSpeedTester = Nothing
    
End Sub

