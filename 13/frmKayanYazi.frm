VERSION 5.00
Begin VB.Form frmKayanYazi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kayan Kazý Editörü"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   Icon            =   "frmKayanYazi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optBehavior 
      Caption         =   "Scroll"
      Enabled         =   0   'False
      Height          =   195
      Index           =   0
      Left            =   4200
      TabIndex        =   35
      Top             =   6240
      Width           =   855
   End
   Begin VB.OptionButton optAlign 
      Caption         =   "Top"
      Enabled         =   0   'False
      Height          =   195
      Index           =   0
      Left            =   4200
      TabIndex        =   32
      Top             =   6360
      Width           =   855
   End
   Begin VB.OptionButton optYon 
      Caption         =   "Sol"
      Enabled         =   0   'False
      Height          =   195
      Index           =   0
      Left            =   4200
      TabIndex        =   30
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox txtMar 
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Text            =   "frmKayanYazi.frx":1042
      Top             =   3840
      Width           =   4935
   End
   Begin HizliHtml.Command cmdKapat 
      Height          =   325
      Left            =   3720
      TabIndex        =   18
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      Caption         =   "Kapat"
   End
   Begin HizliHtml.Command cmdKodEkle 
      Height          =   325
      Left            =   3720
      TabIndex        =   17
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      Caption         =   "Kodu Ekle"
   End
   Begin HizliHtml.Command cmdKodOlustur 
      Height          =   325
      Left            =   3720
      TabIndex        =   16
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      Caption         =   "Kodu Oluþtur"
   End
   Begin VB.ComboBox cmbRenk 
      Height          =   315
      Left            =   3360
      TabIndex        =   15
      Text            =   "Renk Seç"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Frame frmeYon 
      Caption         =   "Yön: Sol mu Sað mý?"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   3375
      Begin VB.OptionButton optYon 
         Caption         =   "Sol"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optYon 
         Caption         =   "Sað"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame frmeAlign 
      Caption         =   "Align"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   3375
      Begin VB.OptionButton optAlign 
         Caption         =   "Top"
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optAlign 
         Caption         =   "Bottom"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   2160
         TabIndex        =   34
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optAlign 
         Caption         =   "Middle"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   33
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame frmeBehavior 
      Caption         =   "Behavior"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   3375
      Begin VB.OptionButton optBehavior 
         Caption         =   "Scroll"
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optBehavior 
         Caption         =   "Alternate"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   2160
         TabIndex        =   37
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optBehavior 
         Caption         =   "Slide"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CheckBox chkRenkYok 
      Caption         =   "Arkaplan Rengi Yok"
      Height          =   195
      Left            =   3360
      TabIndex        =   11
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CheckBox chkHiz 
      Caption         =   "Hýz"
      Height          =   195
      Left            =   3360
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox chkDelay 
      Caption         =   "Delay"
      Height          =   195
      Left            =   3360
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.CheckBox chkTekrarSayisi 
      Caption         =   "Tekrar Sayýsý"
      Height          =   195
      Left            =   1800
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CheckBox chkDikeyM 
      Caption         =   "Dikey Margin"
      Height          =   195
      Left            =   1800
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CheckBox chkYatayM 
      Caption         =   "Yatay Margin"
      Height          =   195
      Left            =   1800
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CheckBox chkGenislik 
      Caption         =   "Geniþlik"
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.CheckBox chkYukseklik 
      Caption         =   "Yükseklik"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CheckBox chkSagSol 
      Caption         =   "Saða mý Sola mý?"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CheckBox chkBehavior 
      Caption         =   "Behavior"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox chkAlign 
      Caption         =   "Align"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblRenk 
      Height          =   255
      Left            =   3480
      TabIndex        =   41
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label lblHiz 
      Height          =   255
      Left            =   3480
      TabIndex        =   29
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lblDelay 
      Height          =   255
      Left            =   3480
      TabIndex        =   28
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label lblTekrarSayisi 
      Height          =   255
      Left            =   1800
      TabIndex        =   27
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label lblDikeyM 
      Height          =   255
      Left            =   1800
      TabIndex        =   26
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label lblYatayM 
      Height          =   255
      Left            =   1800
      TabIndex        =   25
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lblGenislik 
      Height          =   255
      Left            =   1800
      TabIndex        =   24
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label lblYukseklik 
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label lblYon 
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label lblBehavior 
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lblAlign 
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label lblGiris 
      Caption         =   $"frmKayanYazi.frx":1056
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmKayanYazi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAlign_Click()
If chkAlign.Value = 1 Then
    frmeAlign.Enabled = True
    optAlign(1).Enabled = True
    optAlign(2).Enabled = True
    optAlign(3).Enabled = True
ElseIf chkAlign.Value = 0 Then
    lblAlign.Caption = ""
    frmeAlign.Enabled = False
    optAlign(1).Enabled = False
    optAlign(2).Enabled = False
    optAlign(3).Enabled = False
    optAlign(1).Value = False
    optAlign(2).Value = False
    optAlign(3).Value = False
End If
End Sub

Private Sub chkBehavior_Click()
If chkBehavior.Value = 1 Then
    frmeBehavior.Enabled = True
    optBehavior(1).Enabled = True
    optBehavior(2).Enabled = True
    optBehavior(3).Enabled = True
ElseIf chkBehavior.Value = 0 Then
    lblBehavior.Caption = ""
    frmeBehavior.Enabled = False
    optBehavior(1).Enabled = False
    optBehavior(2).Enabled = False
    optBehavior(3).Enabled = False
    optBehavior(1).Value = False
    optBehavior(2).Value = False
    optBehavior(3).Value = False
End If
End Sub

Private Sub chkDelay_Click()
If chkDelay.Value = 1 Then
    lblDelay.Caption = "scrolldelay=" & Chr(34) & 400 & Chr(34) & ""
ElseIf chkDelay.Value = 0 Then
    lblDelay.Caption = ""
End If
End Sub

Private Sub chkDikeyM_Click()
If chkDikeyM.Value = 1 Then
    lblDikeyM.Caption = "vspace=" & Chr(34) & 10 & Chr(34) & ""
ElseIf chkDikeyM.Value = 0 Then
    lblDikeyM.Caption = ""
End If
End Sub

Private Sub chkGenislik_Click()
If chkGenislik.Value = 1 Then
    lblGenislik.Caption = "width=" & Chr(34) & 100 & Chr(34) & ""
ElseIf chkGenislik.Value = 0 Then
    lblGenislik.Caption = ""
End If
End Sub

Private Sub chkHiz_Click()
If chkHiz.Value = 1 Then
    lblHiz.Caption = "scrollamount=" & Chr(34) & 4 & Chr(34) & ""
ElseIf chkHiz.Value = 0 Then
    lblHiz.Caption = ""
End If
End Sub

Private Sub chkRenkYok_Click()
If chkRenkYok.Value = 1 Then
cmbRenk.Enabled = False
lblRenk.Caption = ""
lblRenk.Enabled = False
ElseIf chkRenkYok.Value = 0 Then
cmbRenk.Enabled = True
lblRenk.Caption = "bgcolor="
End If
End Sub

Private Sub chkSagSol_Click()
If chkSagSol.Value = 1 Then
    optYon(1).Enabled = True
    optYon(2).Enabled = True
    frmeYon.Enabled = True
ElseIf chkSagSol.Value = 0 Then
    lblYon.Caption = ""
    optYon(1).Enabled = False
    optYon(2).Enabled = False
    optYon(1).Value = False
    optYon(2).Value = False
    frmeYon.Enabled = False
End If
End Sub

Private Sub chkTekrarSayisi_Click()
If chkTekrarSayisi.Value = 1 Then
    lblTekrarSayisi.Caption = "loop=" & Chr(34) & 1 & Chr(34) & ""
ElseIf chkTekrarSayisi.Value = 0 Then
    lblTekrarSayisi.Caption = ""
End If
End Sub

Private Sub chkYatayM_Click()
If chkYatayM.Value = 1 Then
    lblYatayM.Caption = "hspace=" & Chr(34) & 10 & Chr(34) & ""
ElseIf chkYatayM.Value = 0 Then
    lblYatayM.Caption = ""
End If
End Sub

Private Sub chkYukseklik_Click()
If chkYukseklik.Value = 1 Then
    lblYukseklik.Caption = "height=" & Chr(34) & 40 & Chr(34) & ""
ElseIf chkYukseklik.Value = 0 Then
    lblYukseklik.Caption = ""
End If
End Sub

Private Sub cmbRenk_Click()
'Renklerin türkçelerinin yanýna ingilizcelerinide yazdým
'türkçe isimler ile kodlar tutmuyorsa lütfen beni
'haberdar edin :)
If cmbRenk.Text = "Siyah" Then 'Black
    lblRenk.Caption = "bgcolor=" & Chr(34) & Chr(35) & "000000" & Chr(34) & ""
ElseIf cmbRenk.Text = "Gri" Then 'Gray
    lblRenk.Caption = "bgcolor=" & Chr(34) & Chr(35) & "808080" & Chr(34) & ""
ElseIf cmbRenk.Text = "Gümüþ" Then 'Silver
    lblRenk.Caption = "bgcolor=" & Chr(34) & Chr(35) & "C0C0C0" & Chr(34) & ""
ElseIf cmbRenk.Text = "Beyaz" Then 'White
    lblRenk.Caption = "bgcolor=" & Chr(34) & Chr(35) & "FFFFFF" & Chr(34) & ""
ElseIf cmbRenk.Text = "Kýrmýzý" Then 'Red
    lblRenk.Caption = "bgcolor=" & Chr(34) & Chr(35) & "FF0000" & Chr(34) & ""
ElseIf cmbRenk.Text = "Mor" Then 'Purple
    lblRenk.Caption = "bgcolor=" & Chr(34) & Chr(35) & "800080" & Chr(34) & ""
ElseIf cmbRenk.Text = "Parlak Mor" Then 'Fucsia
    lblRenk.Caption = "bgcolor=" & Chr(34) & Chr(35) & "FF00FF" & Chr(34) & ""
ElseIf cmbRenk.Text = "Sarý" Then 'Yellow
    lblRenk.Caption = "bgcolor=" & Chr(34) & Chr(35) & "FFFF00" & Chr(34) & ""
ElseIf cmbRenk.Text = "Çuha Yeþili" Then 'Teal
    lblRenk.Caption = "bgcolor=" & Chr(34) & Chr(35) & "008080" & Chr(34) & ""
ElseIf cmbRenk.Text = "Mavi" Then 'Blue
    lblRenk.Caption = "bgcolor=" & Chr(34) & Chr(35) & "0000FF" & Chr(34) & ""
ElseIf cmbRenk.Text = "Deniz Mavisi" Then 'Navy
    lblRenk.Caption = "bgcolor=" & Chr(34) & Chr(35) & "000080" & Chr(34) & ""
ElseIf cmbRenk.Text = "Parlak Mavi" Then 'Cyan
    lblRenk.Caption = "bgcolor=" & Chr(34) & Chr(35) & "00FFFF" & Chr(34) & ""
ElseIf cmbRenk.Text = "Açýk Yeþil" Then 'Lime
    lblRenk.Caption = "bgcolor=" & Chr(34) & Chr(35) & "008000" & Chr(34) & ""
ElseIf cmbRenk.Text = "Yeþil" Then 'Green
    lblRenk.Caption = "bgcolor=" & Chr(34) & Chr(35) & "008080" & Chr(34) & ""
ElseIf cmbRenk.Text = "Zeytuni" Then 'Olive
    lblRenk.Caption = "bgcolor=" & Chr(34) & Chr(35) & "808000" & Chr(34) & ""
ElseIf cmbRenk.Text = "Bordo" Then 'Maroon
    lblRenk.Caption = "bgcolor=" & Chr(34) & Chr(35) & "800000" & Chr(34) & ""
End If
End Sub

Private Sub cmdKapat_Click()
Unload Me
End Sub

Private Sub cmdKodEkle_Click()
EtiketEkle txtMar.Text, True
frmAna.ImlecYerlestir frmAna.rchHtml.SelText, 10
End Sub

Private Sub cmdKodOlustur_Click()
On Local Error Resume Next
txtMar.Text = "<marquee" & Chr(32) & lblYukseklik.Caption & Chr(32) & lblGenislik.Caption & Chr(32) & lblAlign.Caption & Chr(32) & lblBehavior & Chr(32) & lblDelay.Caption & Chr(32) & lblDikeyM.Caption & Chr(32) & lblYatayM.Caption & Chr(32) & lblHiz.Caption & Chr(32) & lblTekrarSayisi.Caption & Chr(32) & lblYon.Caption & Chr(32) & lblRenk.Caption & Chr(32) & ">Kayan Yazi Metni</marquee>"
End Sub

Private Sub cmdRenk_Click()
frmAna.cd1.ShowColor

End Sub

Private Sub Form_Load()
cmbRenk.AddItem "Siyah"
cmbRenk.AddItem "Gri"
cmbRenk.AddItem "Gümüþ"
cmbRenk.AddItem "Beyaz"
cmbRenk.AddItem "Kýrmýzý"
cmbRenk.AddItem "Mor"
cmbRenk.AddItem "Koyu Mor"
cmbRenk.AddItem "Sarý"
cmbRenk.AddItem "Açýk Mavi"
cmbRenk.AddItem "Mavi"
cmbRenk.AddItem "Deniz Mavisi"
cmbRenk.AddItem "Parlak Mavi"
cmbRenk.AddItem "Açýk Yeþil"
cmbRenk.AddItem "Yeþil"
cmbRenk.AddItem "Zeytuni"
cmbRenk.AddItem "Bordo"
End Sub

Private Sub optAlign_Click(Index As Integer)
If optAlign(1).Value = True Then
    lblAlign.Caption = "align=" & Chr(34) & "middle" & Chr(34) & ""
ElseIf optAlign(2).Value = True Then
    lblAlign.Caption = "align=" & Chr(34) & "bottom" & Chr(34) & ""
ElseIf optAlign(3).Value = True Then
    lblAlign.Caption = "align=" & Chr(34) & "top" & Chr(34) & ""
ElseIf optAlign(1).Value = False And optAlign(2).Value = False And optAlign(3).Value = False Then
    lblAlign.Caption = ""
End If
End Sub

Private Sub optBehavior_Click(Index As Integer)
If optBehavior(1).Value = True Then
    lblBehavior.Caption = "behavior=" & Chr(34) & "slide" & Chr(34) & ""
ElseIf optBehavior(2).Value = True Then
    lblBehavior.Caption = "behavior=" & Chr(34) & "alternate" & Chr(34) & ""
ElseIf optBehavior(3).Value = True Then
    lblBehavior.Caption = "behavior=" & Chr(34) & "scroll" & Chr(34) & ""
ElseIf optBehavior(1).Value = False And optBehavior(2).Value = False And optBehavior(3).Value = False Then
    lblBehavior.Caption = ""
End If
End Sub

Private Sub optYon_Click(Index As Integer)
If optYon(1).Value = True Then
    lblYon.Caption = "direction=" & Chr(34) & "right" & Chr(34) & ""
ElseIf optYon(2).Value = True Then
    lblYon.Caption = "direction=" & Chr(34) & "left" & Chr(34) & ""
ElseIf optYon(1).Value = False And optYon(2).Value = False Then
    lblYon.Caption = ""
End If
End Sub
