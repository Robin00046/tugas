VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Selesai 
      Caption         =   "Selesai"
      Height          =   855
      Left            =   2520
      TabIndex        =   15
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Ulangi 
      Caption         =   "Ulangi"
      Height          =   855
      Left            =   360
      TabIndex        =   14
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox GajiPokok 
      Height          =   495
      Left            =   1800
      TabIndex        =   13
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox Agama 
      Height          =   495
      Left            =   1800
      TabIndex        =   12
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox JK 
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox Alamat 
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox Nama 
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox Bagian 
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox NIP 
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "Gaji Pokok"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Agama"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Jenis Kelamin"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Alamat"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Nama"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Bagian"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "NIP"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Agama_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 'validasi isian jurusan
 xnip = Left(Me.Agama.Text, 1)
 isi = LTrim(xnip)
 If xnip = "I" Then
 Me.Agama.Text = isi & " - Islam"
 ElseIf xnip = "K" Then
  Me.Agama.Text = isi & " - Kristen"
   ElseIf xnip = "T" Then
  Me.Agama.Text = isi & " - Katolik"
   ElseIf xnip = "B" Then
  Me.Agama.Text = isi & " - Budha "
   ElseIf xnip = "H" Then
  Me.Agama.Text = isi & " - Hindu"
 Else
 MsgBox "Agama Harus Diisi I/K/T/B/H .. Ok", vbOKOnly, "Pesan"
  Me.Agama.SetFocus
 Exit Sub
 End If
 Me.Ulangi.SetFocus
 Exit Sub
 End If
End Sub

Private Sub Alamat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Me.JK.SetFocus
End If
End Sub

Private Sub JK_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 'validasi isian jurusan
 xnip = Left(Me.JK.Text, 1)
 isi = LTrim(xnip)
 If xnip = "L" Then
 Me.JK.Text = isi & " - Laki-Laki"
 ElseIf xnip = "P" Then
  Me.JK.Text = isi & " - Perempuan"
 Else
 MsgBox "Jenis Kelamin Harus L/P .. Ok", vbOKOnly, "Pesan"
  Me.JK.SetFocus
 Exit Sub
 End If
 Me.Agama.SetFocus
 Exit Sub
 End If
End Sub


Private Sub Nama_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
Me.Alamat.SetFocus
End If
End Sub

Private Sub NIP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If Len(Trim(Me.NIP.Text)) > 7 Then
 MsgBox "Jumlah Kareakter kurang dari 7 .. Ok", vbOKOnly, "Pesan"
 Me.NIP.SetFocus
 Exit Sub
 End If
 'validasi isian jurusan
 xnip = Left(Me.NIP.Text, 2)
 isi = LTrim(xnip)
 If xnip = "DR" Then
 Me.Bagian.Text = isi & " - Direktur"
 Me.GajiPokok.Text = isi & " - 1000000"
 ElseIf xnip = "MN" Then
  Me.Bagian.Text = isi & " - Manager"
 Me.GajiPokok.Text = isi & " - 800000"
  ElseIf xnip = "ST" Then
  Me.Bagian.Text = isi & " - Staff"
 Me.GajiPokok.Text = isi & " - 700000"
  ElseIf xnip = "PB" Then
  Me.Bagian.Text = isi & " - Pembantu Umum"
 Me.GajiPokok.Text = isi & " - 500000"
 Else
 MsgBox "Kode Jurusan Harus DR/MN/ST/PB .. Ok", vbOKOnly, "Pesan"
  Me.NIP.SetFocus
 Exit Sub
 End If
 Me.Nama.SetFocus
 Exit Sub
 End If

 End Sub

Private Sub Selesai_Click()
End
End Sub

Private Sub Ulangi_Click()
NIP.Text = ""
Bagian.Text = ""
Nama.Text = ""
Alamat.Text = ""
JK.Text = ""
Agama.Text = ""
GajiPokok.Text = ""
End Sub
