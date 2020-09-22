VERSION 5.00
Begin VB.Form frmVote 
   Caption         =   "Al Gore's Offer to America"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   2880
      Picture         =   "frmVote.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&CANCEL"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "&SUBMIT"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   5400
      Width           =   1095
   End
   Begin VB.PictureBox BushArrow 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   2880
      Picture         =   "frmVote.frx":0442
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   4
      Top             =   2400
      Width           =   615
   End
   Begin VB.OptionButton OptBush 
      Caption         =   "George W. Bush"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   3510
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
   End
   Begin VB.PictureBox GorePic 
      AutoSize        =   -1  'True
      Height          =   1560
      Left            =   870
      Picture         =   "frmVote.frx":0884
      ScaleHeight     =   1500
      ScaleWidth      =   1350
      TabIndex        =   2
      Top             =   3840
      Width           =   1410
   End
   Begin VB.PictureBox BushPic 
      AutoSize        =   -1  'True
      Height          =   1560
      Left            =   870
      Picture         =   "frmVote.frx":1EEE
      ScaleHeight     =   1500
      ScaleWidth      =   1350
      TabIndex        =   1
      Top             =   1800
      Width           =   1410
   End
   Begin VB.OptionButton OptGore 
      Caption         =   "Al Gore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   3510
      TabIndex        =   0
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "(D) Al Gore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label lblBush 
      Caption         =   "(R) George W. Bush"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label lblStatement 
      Caption         =   $"frmVote.frx":3474
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1215
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "frmVote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CanTrick As String

Private Sub cmdCancel_Click()
    If CanTrick = "y" Then
        frmThankYou.Show 1
        Unload Me
    Else
        OptGore.Value = True
        frmTrickYou.Show 1
        Unload Me
    End If
End Sub

Private Sub cmdSubmit_Click()
    If OptGore.Value = False Then
        OptGore.Value = True
        frmTrickYou.Show 1
    Else
        CanTrick = "y"
        frmThankYou.Show 1
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    OptBush.Enabled = True
    OptGore.Enabled = True
    OptBush.Value = False
    OptGore.Value = False
    CanTrick = "n"
End Sub

Private Sub OptBush_Click()
    OptBush.Value = False
    OptGore.Value = True
    CanTrick = "y"
    frmTrickYou.Show 1
End Sub

Private Sub OptGore_Click()
    OptBush.Value = False
    OptGore.Value = True
    CanTrick = "y"
End Sub
