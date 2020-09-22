VERSION 5.00
Begin VB.Form frmListe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liste Editörü"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmListe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin HizliHtml.Command cmdKodEkle 
      Height          =   325
      Left            =   2760
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      Caption         =   "Kodu Ekle"
   End
   Begin VB.HScrollBar hsOgeSayisi 
      Height          =   255
      Left            =   2400
      Max             =   50
      TabIndex        =   8
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtOgeSayisi 
      Height          =   300
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "1"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.OptionButton optSirali 
      Caption         =   "Roma Rakamlý Sýralý Metin"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   2175
   End
   Begin VB.OptionButton optSirali 
      Caption         =   "Sayý Sýralý Metin"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.OptionButton optSirali 
      Caption         =   "Kare Noktalý Metin"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.OptionButton optSirali 
      Caption         =   "Alfabetik Sýralý Metin"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.OptionButton optSirali 
      Caption         =   "Noktalý Metin"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblSayi 
      Caption         =   "Liste öðelerinin sayýsýný belirtiniz:"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblGiris 
      Caption         =   "Liste editörü html sayfalarýnýz için oluþturacaðýnýz sýralý metinleri için  size kolaylýk saðlar."
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmListe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdKodEkle_Click()
On Local Error Resume Next
Dim X As Long, Y
    'Noktalý Metnimizin kodlarýný oluþturalým:
    If optSirali(0).Value = True Then
        frmAna.rchHtml.SelText = "<ul>" + vbCrLf
            Do
                X = txtOgeSayisi.Text
                frmAna.rchHtml.SelText = "<li>Nokta Sýralý Metin</li>" + vbCrLf
                Y = Y + 1
            Loop While Y < X
        frmAna.rchHtml.SelText = "</ul>" + vbCrLf
    End If
    '
    'Sayý sýralý Metnimizin kodlarýný oluþturalým:
    If optSirali(3).Value = True Then
        frmAna.rchHtml.SelText = "<ol>" + vbCrLf
            Do
                X = txtOgeSayisi.Text
                frmAna.rchHtml.SelText = "<li>Sayý Sýralý Metin</li>" + vbCrLf
                Y = Y + 1
            Loop While Y < X
        frmAna.rchHtml.SelText = "</ol>" + vbCrLf
    End If
    '
    'Kare noktalý sýralý metnimizin kodlarýný oluþturalým:
    If optSirali(2).Value = True Then
        frmAna.rchHtml.SelText = "<ul type=" & Chr(34) & "square" & Chr(34) & ">" + vbCrLf
            Do
                X = txtOgeSayisi.Text
                frmAna.rchHtml.SelText = "<li>Kare Noktalý Sýrali Metin</li>" + vbCrLf
                Y = Y + 1
            Loop While Y < X
        frmAna.rchHtml.SelText = "</ul>" + vbCrLf
    End If
    '
    'Alfabetik sýralý metnimizn kodlarýný hazýrlayalým:
    If optSirali(1).Value = True Then
        frmAna.rchHtml.SelText = "<ol type=" & Chr(34) & "a" & Chr(34) & ">" + vbCrLf
            Do
                X = txtOgeSayisi.Text
                frmAna.rchHtml.SelText = "<li>Alfabetik Sýrali Metin</li>" + vbCrLf
                Y = Y + 1
            Loop While Y < X
        frmAna.rchHtml.SelText = "</ol>" + vbCrLf
    End If
    '
    'Roma rakamlý siralý metnimizi oluþturalým:
    If optSirali(4).Value = True Then
        frmAna.rchHtml.SelText = "<ol type=" & Chr(34) & "i" & Chr(34) & ">" + vbCrLf
            Do
                X = txtOgeSayisi.Text
                frmAna.rchHtml.SelText = "<li>Roma Rakamý Sýrali Metin</li>" + vbCrLf
                Y = Y + 1
            Loop While Y < X
        frmAna.rchHtml.SelText = "</ol>" + vbCrLf
    End If
Unload Me
End Sub

Private Sub Form_Load()
txtOgeSayisi.Text = hsOgeSayisi.Value
End Sub

Private Sub hsOgeSayisi_Change()
txtOgeSayisi.Text = hsOgeSayisi.Value
End Sub
