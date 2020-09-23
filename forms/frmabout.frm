VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   360
      Left            =   3555
      TabIndex        =   5
      Top             =   1800
      Width           =   1005
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   0
      ScaleHeight     =   870
      ScaleWidth      =   4680
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Create and design your own colofull scrollbar for Internet Explorer."
         Height          =   405
         Left            =   720
         TabIndex        =   2
         Top             =   390
         Width           =   3675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DM CSS Scrollbar Designer Plus"
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
         Left            =   690
         TabIndex        =   1
         Top             =   150
         Width           =   2790
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   75
         Picture         =   "frmabout.frx":0000
         Top             =   90
         Width           =   480
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Designed by Ben Jones"
      Height          =   195
      Index           =   1
      Left            =   1650
      TabIndex        =   4
      Top             =   1245
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "THIS PROGRAM IS FREEWARE"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   990
      Width           =   2400
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   1695
      Y1              =   870
      Y2              =   870
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()
    Unload frmabout
End Sub

Private Sub Form_Load()
    Line1.X2 = frmabout.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmabout = Nothing
End Sub
