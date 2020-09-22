VERSION 5.00
Begin VB.Form frmBaglanti 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Link Editörü"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
   Icon            =   "fmBaglanti.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin HizliHtml.Command cmdKodOlustur 
      Height          =   325
      Left            =   360
      TabIndex        =   15
      Top             =   4200
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      Caption         =   "Kodu Oluþtur"
   End
   Begin VB.TextBox txtAdres 
      Height          =   325
      Left            =   120
      TabIndex        =   14
      Text            =   "www.sct.tr.cx"
      Top             =   2880
      Width           =   3735
   End
   Begin HizliHtml.Command cmdKapat 
      Height          =   325
      Left            =   2040
      TabIndex        =   10
      Top             =   4800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      Caption         =   "Kapat"
   End
   Begin HizliHtml.Command cmdKodEkle 
      Height          =   325
      Left            =   360
      TabIndex        =   9
      Top             =   4800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      Caption         =   "Kodu Ekle"
   End
   Begin VB.TextBox txtADDR 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "fmBaglanti.frx":1042
      Top             =   3360
      Width           =   3735
   End
   Begin VB.ComboBox cmbTarget 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Text            =   "Target"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CheckBox chkTarget 
      Caption         =   "Target"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.OptionButton optProtokol 
      Caption         =   "telnet://"
      Height          =   195
      Index           =   4
      Left            =   2400
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton optProtokol 
      Caption         =   "gopher://"
      Height          =   195
      Index           =   3
      Left            =   2400
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton optProtokol 
      Caption         =   "ftp://"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton optProtokol 
      Caption         =   "http://"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton optProtokol 
      Caption         =   "Protokol Yok"
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblAdres 
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label lblProtokol 
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label lblTarget 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label lblGiris 
      Caption         =   "Link olmayan bir sayfa düþünülemez her halde :) Link Editörü size daha kolay linkler hazýrlayabilmenizi saðlayacak.."
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmBaglanti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkTarget_Click()
If chkTarget.Value = 1 Then
    cmbTarget.Enabled = True
ElseIf chkTarget.Value = 0 Then
    cmbTarget.Enabled = False
    lblTarget.Caption = ""
End If
End Sub

Private Sub cmbTarget_Click()
If cmbTarget.Text = "_blank" Then
    lblTarget.Caption = "target=" & Chr(34) & "_blank" & Chr(34) & ""
ElseIf cmbTarget.Text = "_self" Then
    lblTarget.Caption = "target=" & Chr(34) & "_self" & Chr(34) & ""
ElseIf cmbTarget.Text = "_main" Then
    lblTarget.Caption = "target=" & Chr(34) & "_main" & Chr(34) & ""
End If
End Sub

Private Sub cmdKapat_Click()
Unload Me
End Sub

Private Sub cmdKodEkle_Click()
EtiketEkle txtADDR.Text, True
frmAna.ImlecYerlestir frmAna.rchHtml.SelText, 4
End Sub

Private Sub cmdKodOlustur_Click()
txtADDR.Text = "<a href=" & Chr(34) & lblProtokol.Caption + lblAdres.Caption & Chr(34) & Chr(32) & lblTarget.Caption & Chr(62) & "Gidilecek Adres</a>"
End Sub

Private Sub Form_Load()
cmbTarget.AddItem "_blank"
cmbTarget.AddItem "_self"
cmbTarget.AddItem "_main"
End Sub

Private Sub optProtokol_Click(Index As Integer)
If optProtokol(0).Value = True Then
    lblProtokol.Caption = ""
ElseIf optProtokol(1).Value = True Then
    lblProtokol.Caption = "http://"
ElseIf optProtokol(2).Value = True Then
    lblProtokol.Caption = "ftp://"
ElseIf optProtokol(3).Value = True Then
    lblProtokol.Caption = "gopher://"
ElseIf optProtokol(4).Value = True Then
    lblProtokol.Caption = "telnet://"
End If
End Sub

Private Sub txtAdres_Change()
lblAdres.Caption = txtAdres.Text
End Sub
