VERSION 5.00
Begin VB.Form frmBul 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bul.."
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   ControlBox      =   0   'False
   Icon            =   "frmBul.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin HizliHtml.Command cmdTumunuDegistir 
      Height          =   325
      Left            =   4440
      TabIndex        =   10
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      Caption         =   "T�m�n� De�i�tir"
   End
   Begin HizliHtml.Command cmdDegistir 
      Height          =   325
      Left            =   4440
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      Caption         =   "&De�i�tir..."
   End
   Begin VB.Frame frmeArama 
      Caption         =   "Arama Se�enekleri"
      Height          =   980
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   3255
      Begin VB.CheckBox chkDuyarli 
         Caption         =   "B�y�k k���k harf duyarl�"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2535
      End
      Begin VB.CheckBox chkTumKelime 
         Caption         =   "Sadece T�m Kelime"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.ComboBox cmoDegistir 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   600
      Width           =   2775
   End
   Begin HizliHtml.Command cmdIptal 
      Height          =   325
      Left            =   4440
      TabIndex        =   3
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      Caption         =   "�ptal"
   End
   Begin VB.ComboBox cmoAranacakKelime 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin HizliHtml.Command cmdBul 
      Height          =   325
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      Caption         =   "Bul"
   End
   Begin VB.Label lblDegistir 
      Caption         =   "Bununla De�i�tir:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   650
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Aranacak Kelime:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   160
      Width           =   1335
   End
End
Attribute VB_Name = "frmBul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdIptal_Click()
Unload Me
End Sub

Private Sub Form_Load()
lblDegistir.Visible = False
cmoDegistir.Visible = False
cmdTumunuDegistir.Visible = False

cmoAranacakKelime.Text = frmAna.rchHtml.SelText 'metni cboya yerle�tir
End Sub

Private Sub cmdbul_Click()
    On Local Error Resume Next
    Dim lngSonuc As Long
    Dim lngYer As Long
    Dim intSecenekler As Integer
    ' arama se�eneklerini ayarla
    If chkTumKelime.Value = 1 Then intSecenekler = intSecenekler + 2
    If chkDuyarli.Value = 1 Then intSecenekler = intSecenekler + 4

    If cmdBul.Caption = "&Bul" Then 'e�er ilk kez aran�l�yorsa
        ' aranan kelimenin terini bul
        lngSonuc = frmAna.rchHtml.Find(cmoAranacakKelime.Text, 0, , intSecenekler)

        If lngSonuc = -1 Then 'aran�lan metin bulunamad�
            MsgBox "Arad���n�z kelime bulunamad�.", vbExclamation + vbOKOnly, "[SCT] H�zl� HTML v1.3" 'mesaj� g�zterenzi
            cmdBul.Caption = "&Bul" 'ba�l��� ayarla
        Else 'aran�lan metin bulundu
            frmAna.rchHtml.SetFocus 'rchHtmlye odaklan
            cmdDegistir.Enabled = True 'de�i�tir tu�unu etkinle�tir
            cmdTumunuDegistir.Enabled = True 't�m�n� de�i�tir tu�unu etkinle�tir
            cmdBul.Caption = "&Sonrakini Bul" 'ba�l��� ayarla
        End If
    Else 'Sonrakini bul
        lngYer = frmAna.rchHtml.SelStart + frmAna.rchHtml.SelLength
        lngSonuc = frmAna.rchHtml.Find(cmoAranacakKelime.Text, lngYer, , intSecenekler)

        If lngSonuc = -1 Then 'aran�lan metin bulunamad�
            MsgBox "Arad���n�z kelime bulunamad�.", vbExclamation + vbOKOnly, "[SCT] H�zl� HTML v1.3" 'mesaj� g�zterenzi
            cmdBul.Caption = "&Bul" 'ba�l��� ayarla
            cmdDegistir.Enabled = False 'de�i�tir tu�unu etkinle�tirme
            cmdTumunuDegistir.Enabled = False 't�m�n� de�i�tir tu�unu etkinle�tirme
        Else 'metin bulunanzi
            frmAna.rchHtml.SetFocus 'odaklan
        End If
    End If
End Sub

Private Sub cmddegistir_Click()
    On Local Error Resume Next
    Dim lngSonuc As Long
    Dim lngYer As Long
    Dim intSecenekler As Integer
    
    If cmdDegistir.Caption = "&De�i�tir..." Then 'de�i�tiri g�ster
        cmdDegistir.Caption = "&De�i�tir" 'ba�l��� ayarla
        lblDegistir.Visible = True 'lblDe�i�tiri g�ster
        cmoDegistir.Visible = True 'cmodegistiri g�ster
        cmdTumunuDegistir.Visible = True 'cmdtumunudegistiri g�ster
        Exit Sub
    End If

    ' arama se�eneklerini ayarla
    If chkTumKelime.Value = 1 Then intSecenekler = intSecenekler + 2
    If chkDuyarli.Value = 1 Then intSecenekler = intSecenekler + 4
    
    With frmAna
        .rchHtml.SelText = cmoDegistir.Text 'Metni de�i�tir
        ' sonrakini bul
        lngYer = .rchHtml.SelStart + .rchHtml.SelLength
        ' aran�lan kelimenin yerini bul
        lngSonuc = .rchHtml.Find(cmoAranacakKelime.Text, lngYer, , intSecenekler)

        If lngSonuc = -1 Then 'aran�lan metin bulunamad�
            MsgBox "Arad���n�z kelime bulunamad�.", vbExclamation + vbOKOnly, "[SCT] H�zl� HTML v1.3" 'mesaj� g�zterenzi
            cmdBul.Caption = "&Bul" 'ba�l��� ayarla
            cmdDegistir.Enabled = False 'de�i�tir tu�unu etkinle�tirme
            cmdTumunuDegistir.Enabled = False 't�m�n� de�i�tir tu�unu etkinle�tirme
        Else 'metin bulunanzi
            .rchHtml.SetFocus 'odaklan
        End If
    End With
End Sub

Private Sub cmdtumunudegistir_Click()
    On Local Error Resume Next
    Dim intSay As Integer
    Dim lngYer As Long
    Dim intSecenekler As Integer
    ' arama se�eneklerini ayarla
    If chkTumKelime.Value = 1 Then intSecenekler = intSecenekler + 2
    If chkDuyarli.Value = 1 Then intSecenekler = intSecenekler + 4

    intSay = 0
    lngYer = 0
    With frmAna
        Do
            If .rchHtml.Find(cmoAranacakKelime.Text, lngYer, , intSecenekler) = -1 Then 'metin bulunamad�
                If intSay > 0 Then 'ka� tane yer de�i�tirilmesi yap�ld���n� g�ster
                    MsgBox "Belirti�iniz alan tarand�." & intSay & " de�i�tirilme yap�ld�.", vbInformation + vbOKOnly, "[SCT] H�zl� HTML v1.3"
                End If
                cmdBul.Caption = "&Bul" 'ba�l��� ayarla
                cmdDegistir.Enabled = False 'de�i�tir tu�unu etkinle�tirme
                cmdTumunuDegistir.Enabled = False 't�m�n� de�i�tir tu�unu etkinle�tirme
                Exit Do
            Else 'metin bulunanzi
                lngYer = .rchHtml.SelStart + .rchHtml.SelLength
                intSay = intSay + 1 'sayac� birer artt�r
                .rchHtml.SelText = cmoDegistir.Text 'Metni de�i�tir
            End If
        Loop
    End With
End Sub

