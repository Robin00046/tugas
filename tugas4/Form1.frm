VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tugas 4B"
   ClientHeight    =   2925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox EsTeh 
      Caption         =   "EsTeh"
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox NasiPutih 
      Caption         =   "NasiPutih"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.CheckBox AyamGoreng 
      Caption         =   "AyamGoreng"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Selesai 
      Caption         =   "Selesai"
      Height          =   615
      Left            =   3600
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Tugas 4A"
      Height          =   615
      Left            =   5400
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ulangi"
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox TxtBiaya 
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Jumlah Bayar"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Pilihan Menu"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AyamGoreng_Click()
If Me.AyamGoreng.Value = 1 And Me.NasiPutih.Value = 1 And Me.EsTeh.Value = 1 Then
Me.TxtBiaya.text = 13000
ElseIf Me.AyamGoreng.Value = 1 And Me.NasiPutih.Value = 1 Then
Me.TxtBiaya.text = 11000
ElseIf Me.AyamGoreng.Value = 1 And Me.EsTeh.Value = 1 Then
Me.TxtBiaya.text = 10000
ElseIf Me.NasiPutih.Value = 1 And Me.EsTeh.Value = 1 Then
Me.TxtBiaya.text = 6000
ElseIf Me.AyamGoreng.Value = 1 Then
Me.TxtBiaya.text = 8000
ElseIf Me.NasiPutih.Value = 1 Then
Me.TxtBiaya.text = 4000
ElseIf Me.EsTeh.Value = 1 Then
Me.TxtBiaya.text = 2500
Else
Me.TxtBiaya.text = 0
End If
End Sub

Private Sub Command1_Click()
 Me.AyamGoreng.Value = False
 Me.NasiPutih.Value = False
 Me.EsTeh.Value = False
End Sub

Private Sub Command2_Click()
Form2.Show
Form1.Hide
End Sub

Private Sub EsTeh_Click()
Call AyamGoreng_Click
End Sub

Private Sub NasiPutih_Click()
Call AyamGoreng_Click
End Sub

Private Sub Option1_Click()

End Sub

Private Sub Option2_Click()

End Sub

Private Sub Selesai_Click()
End
End Sub
