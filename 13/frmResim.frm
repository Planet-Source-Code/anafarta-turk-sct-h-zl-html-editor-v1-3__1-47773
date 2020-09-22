VERSION 5.00
Begin VB.Form frmResim 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resim Editörü"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmResim.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar hscerceve 
      Height          =   300
      Left            =   2760
      Max             =   30
      TabIndex        =   9
      Top             =   1560
      Width           =   1815
   End
   Begin HizliHtml.Command cmdKodEkle 
      Height          =   330
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      Caption         =   "Kodu Ekle"
   End
   Begin VB.TextBox txtResimLink 
      Enabled         =   0   'False
      Height          =   325
      Left            =   120
      TabIndex        =   6
      Text            =   "Resmin Linki"
      Top             =   2280
      Width           =   4455
   End
   Begin VB.CheckBox chkResimLink 
      Caption         =   "Resime Link Ekle"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox txtCerceve 
      Height          =   325
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Resmin Çerçeve Büyüklüðü"
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtAlt 
      Height          =   325
      Left            =   120
      TabIndex        =   3
      Text            =   "Resime ALT Yazýsý Ekle"
      Top             =   1080
      Width           =   4455
   End
   Begin HizliHtml.Command cmdResimYolu 
      Height          =   330
      Left            =   3960
      TabIndex        =   2
      Top             =   600
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      Caption         =   "..."
   End
   Begin VB.TextBox txtResimYolu 
      Height          =   325
      Left            =   120
      TabIndex        =   1
      Text            =   "Resmin Yolu"
      Top             =   600
      Width           =   3735
   End
   Begin HizliHtml.Command cmdKapat 
      Height          =   330
      Left            =   3120
      TabIndex        =   8
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      Caption         =   "Kapat"
   End
   Begin VB.Label lblGiris 
      Caption         =   "Resim Editörü ile html sayfanýza resimler ve resimli linkler ekleyebilirsiniz."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmResim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkResimLink_Click()
If chkResimLink.Value = 1 Then
    txtResimLink.Enabled = True
Else
    txtResimLink.Enabled = False
    txtResimLink.Text = ""
End If
End Sub

Private Sub cmdKapat_Click()
Unload Me
End Sub

Private Sub cmdKodEkle_Click()
If txtCerceve.Text = "Resmin Çerçeve Büyüklüðü" Then
    If chkResimLink.Value = 0 Then
        frmAna.rchHtml.SelText = "<img src=" & Chr(34) & txtResimYolu.Text & Chr(34) & " alt=" & Chr(34) & txtAlt.Text & Chr(34) & ">"
    ElseIf chkResimLink.Value = 1 Then
        frmAna.rchHtml.SelText = "<a href=" & Chr(34) & txtResimLink.Text & Chr(34) & "><img src=" & Chr(34) & txtResimYolu.Text & Chr(34) & " alt=" & Chr(34) & txtAlt.Text & Chr(34) & "></a>"
    End If
Else
    If chkResimLink.Value = 0 Then
        frmAna.rchHtml.SelText = "<img src=" & Chr(34) & txtResimYolu.Text & Chr(34) & " alt=" & Chr(34) & txtAlt.Text & Chr(34) & " border=" & Chr(34) & txtCerceve.Text & Chr(34) & ">"
    ElseIf chkResimLink.Value = 1 Then
        frmAna.rchHtml.SelText = "<a href=" & Chr(34) & txtResimLink.Text & Chr(34) & "><img src=" & Chr(34) & txtResimYolu.Text & Chr(34) & " alt=" & Chr(34) & txtAlt.Text & Chr(34) & " border=" & Chr(34) & txtCerceve.Text & Chr(34) & "></a>"
    End If
End If
End Sub

Private Sub cmdResimYolu_Click()
frmAna.cd1.Filter = "Tüm Resim Dosyalarý|*.jpg;*.jpe;*.bmp;*.ico;*.png;*.pic;*.jpeg;*.emf;*.gif;*.tga|Tüm Dosyalar|*.*"
frmAna.cd1.ShowOpen
On Local Error Resume Next
txtResimYolu.Text = "file://" + frmAna.cd1.FileName
End Sub

Private Sub hscerceve_Change()
txtCerceve.Text = hscerceve.Value
End Sub
