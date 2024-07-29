VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Soal"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12195
   LinkTopic       =   "Form2"
   ScaleHeight     =   6960
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Penyelesaian"
      Height          =   495
      Left            =   9000
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   2640
      Picture         =   "soalinterpolasilinier.frx":0000
      ScaleHeight     =   4635
      ScaleWidth      =   7515
      TabIndex        =   1
      Top             =   1080
      Width           =   7575
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11400
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.Line Line5 
      X1              =   2520
      X2              =   10320
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line4 
      X1              =   2520
      X2              =   2520
      Y1              =   840
      Y2              =   6000
   End
   Begin VB.Line Line3 
      X1              =   10320
      X2              =   10320
      Y1              =   840
      Y2              =   6000
   End
   Begin VB.Line Line2 
      X1              =   2520
      X2              =   2880
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      X1              =   3600
      X2              =   10320
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Soal"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.Show
    Unload Me
End Sub

Private Sub Form_Load()
    Label2.ForeColor = vbRed
End Sub

Private Sub Label2_Click()
    End
End Sub
