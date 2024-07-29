VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Penyelesaian"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14220
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   14220
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   2775
      Left            =   480
      Picture         =   "interpolasilinear.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   5115
      TabIndex        =   22
      Top             =   5160
      Width           =   5175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Review Soal"
      Height          =   495
      Left            =   840
      TabIndex        =   20
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Grafik"
      Height          =   495
      Left            =   3000
      TabIndex        =   18
      Top             =   3960
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   6480
      Picture         =   "interpolasilinear.frx":4E44
      ScaleHeight     =   4215
      ScaleWidth      =   7095
      TabIndex        =   17
      Top             =   3720
      Width           =   7095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Keluar"
      Height          =   495
      Left            =   4320
      TabIndex        =   16
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      Height          =   495
      Left            =   1680
      TabIndex        =   15
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hitung"
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   9360
      TabIndex        =   13
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   4920
      TabIndex        =   9
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4920
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Line Line15 
      X1              =   3360
      X2              =   5760
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line14 
      X1              =   5760
      X2              =   5760
      Y1              =   4920
      Y2              =   8040
   End
   Begin VB.Line Line13 
      X1              =   360
      X2              =   5760
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line Line12 
      X1              =   360
      X2              =   360
      Y1              =   4920
      Y2              =   8040
   End
   Begin VB.Line Line11 
      X1              =   360
      X2              =   600
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label6 
      Caption         =   "Tabel Penyelesaian :"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   21
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Line Line10 
      X1              =   13680
      X2              =   13680
      Y1              =   8040
      Y2              =   3360
   End
   Begin VB.Line Line9 
      X1              =   10560
      X2              =   13680
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line8 
      X1              =   6360
      X2              =   6720
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label5 
      Caption         =   "Grafik Metode Interpolasi Linier"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   19
      Top             =   3240
      Width           =   3975
   End
   Begin VB.Line Line7 
      X1              =   6360
      X2              =   13680
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line Line6 
      X1              =   6360
      X2              =   6360
      Y1              =   3360
      Y2              =   8040
   End
   Begin VB.Label Label14 
      Caption         =   "JARAK HENTI (y) ="
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
      Height          =   495
      Left            =   6960
      TabIndex        =   12
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "Kecepatan Sebuah Kendaraan (x) ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3360
      Width           =   3855
   End
   Begin VB.Label Label9 
      Caption         =   "Jarak Henti 2"
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Jarak Henti 1"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Line Line5 
      X1              =   360
      X2              =   6240
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label4 
      Caption         =   "Kecepatan 2"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Kecepatan 1"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Line Line4 
      X1              =   6240
      X2              =   6240
      Y1              =   1080
      Y2              =   3120
   End
   Begin VB.Line Line3 
      X1              =   4440
      X2              =   6240
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   600
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      Caption         =   "Data Hubungan Antara Kecepatan dan Jarak Henti Yang Dibutuhkan"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   3855
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   360
      Y1              =   1080
      Y2              =   3120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Metode Interpolasi Linier, Untuk Menghitung Jarak Henti"
      BeginProperty Font 
         Name            =   "Niagara Engraved"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   12975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim x, x1, x2, y, y1, y2, hasil As Double
    
    x1 = Val(Text1)
    x2 = Val(Text2)
    y1 = Val(Text3)
    y2 = Val(Text4)
    x = Val(Text5)
    y = Val(Text6)
    
    hasil = ((y2 - y1) / (x2 - x1)) * (x - x1) + y1
    If Text5.Text = Empty Then
        MsgBox "Kecepatan (x) harus diisi"
        Text5.SetFocus
    Else
        Text6.Text = hasil
        Picture2.Visible = True
    End If
End Sub

Private Sub Command2_Click()
    Text1 = Empty
    Text2 = Empty
    Text3 = Empty
    Text4 = Empty
    Text5 = Empty
    Text6 = Empty
    Text1.SetFocus
    Picture1.Visible = False
    Picture2.Visible = False
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub Command4_Click()
    Picture1.Visible = True
End Sub

Private Sub Command5_Click()
    Form2.Show
    Unload Me
End Sub

Private Sub Form_Load()
    Command1.Enabled = False
    Command4.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = False
End Sub

Private Sub Text2_GotFocus()
    If Text3.Text = Empty Then
        MsgBox "Jarak Henti 1 harus diisi"
        Text3.SetFocus
    End If
End Sub

Private Sub Text3_GotFocus()
    If Text1.Text = Empty Then
        MsgBox "Kecepatan 1 harus diisi"
        Text1.SetFocus
    End If
End Sub

Private Sub Text4_GotFocus()
    If Text2.Text = Empty Then
        MsgBox "Kecepatan 2 harus diisi"
        Text2.SetFocus
    End If
End Sub

Private Sub Text5_GotFocus()
    If Text4.Text = Empty Then
        MsgBox "Jarak Henti 2 harus diisi"
        Text4.SetFocus
    End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text3.SetFocus
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text4.SetFocus
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text2.SetFocus
    End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text5.SetFocus
    End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1.Enabled = True
        Command4.Enabled = True
        Command1.SetFocus
    End If
End Sub
