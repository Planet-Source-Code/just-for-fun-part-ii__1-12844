VERSION 5.00
Begin VB.Form frmTrickYou 
   Caption         =   "Thank You for Your VOTE."
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1560
      Left            =   360
      Picture         =   "frmTrickYou.frx":0000
      ScaleHeight     =   1500
      ScaleWidth      =   1350
      TabIndex        =   1
      Top             =   360
      Width           =   1410
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   495
      Left            =   1913
      TabIndex        =   0
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   $"frmTrickYou.frx":166A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frmTrickYou"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    Unload Me
End Sub
