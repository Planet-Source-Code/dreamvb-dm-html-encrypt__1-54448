VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About...."
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdregister 
      Caption         =   "Register"
      Height          =   375
      Left            =   3705
      TabIndex        =   8
      Top             =   2280
      Width           =   960
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2580
      TabIndex        =   7
      Top             =   2295
      Width           =   960
   End
   Begin VB.PictureBox Picture1 
      Height          =   795
      Left            =   240
      ScaleHeight     =   735
      ScaleWidth      =   4365
      TabIndex        =   3
      Top             =   1305
      Width           =   4425
      Begin VB.Label lblsernum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1620
         TabIndex        =   6
         Top             =   390
         Width           =   60
      End
      Begin VB.Label lblser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serial number:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   390
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unregistered User"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   90
         Width           =   1545
      End
   End
   Begin Project1.Line3D Line3D1 
      Height          =   30
      Left            =   195
      TabIndex        =   1
      Top             =   870
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   53
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2003-2004 Ben Jones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1740
      TabIndex        =   9
      Top             =   2910
      Width           =   2970
   End
   Begin VB.Label lblreg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This program is registered to:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   255
      TabIndex        =   2
      Top             =   1080
      Width           =   2565
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DM HTML Locker"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   885
      TabIndex        =   0
      Top             =   210
      Width           =   1575
   End
   Begin VB.Image imgicon 
      Height          =   570
      Left            =   105
      Top             =   120
      Width           =   585
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
    frmabout.Icon = Nothing
    imgicon.Picture = frmmain.Icon
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmabout = Nothing
End Sub
