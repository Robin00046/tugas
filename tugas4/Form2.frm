VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5925
   LinkTopic       =   "Form2"
   ScaleHeight     =   4485
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Tugas 4B"
      Height          =   735
      Left            =   3720
      TabIndex        =   10
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Selesai 
      Caption         =   "Selesai"
      Height          =   735
      Left            =   1920
      TabIndex        =   9
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Ulangi 
      Caption         =   "Ulangi"
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Diterima 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CheckBox Percobaan 
      Caption         =   "Percobaan"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CheckBox Kontrak 
      Caption         =   "Kontrak"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.CheckBox Tetap 
      Caption         =   "Tetap"
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Gaji 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Gaji Diterima"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label label2 
      Caption         =   "Option"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label text 
      Caption         =   "Gaji"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
Form1.Show
Form2.Hide
End Sub

Private Sub Gaji_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
 KeyAscii = 0
 End If
End Sub

Private Sub Kontrak_Click()
 Me.Percobaan.Value = False
 Me.Tetap.Value = False
Me.Diterima.text = Me.Gaji * 90 / 100
End Sub

Private Sub Percobaan_Click()
 Me.Kontrak.Value = False
 Me.Tetap.Value = False
Me.Diterima.text = Me.Gaji * 75 / 100
End Sub

Private Sub Selesai_Click()
End
End Sub

Private Sub Tetap_Click()
Me.Percobaan.Value = False
Me.Kontrak.Value = False
Me.Diterima.text = Me.Gaji * 100 / 100
End Sub

Private Sub Ulangi_Click()
 Me.Kontrak.Value = False
 Me.Percobaan.Value = False
 Me.Tetap.Value = False
End Sub
