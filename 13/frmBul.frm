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
      Caption         =   "Tümünü Deðiþtir"
   End
   Begin HizliHtml.Command cmdDegistir 
      Height          =   325
      Left            =   4440
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      Caption         =   "&Deðiþtir..."
   End
   Begin VB.Frame frmeArama 
      Caption         =   "Arama Seçenekleri"
      Height          =   980
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   3255
      Begin VB.CheckBox chkDuyarli 
         Caption         =   "Büyük küçük harf duyarlý"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2535
      End
      Begin VB.CheckBox chkTumKelime 
         Caption         =   "Sadece Tüm Kelime"
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
      Caption         =   "Ýptal"
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
      Caption         =   "Bununla Deðiþtir:"
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

cmoAranacakKelime.Text = frmAna.rchHtml.SelText 'metni cboya yerleþtir
End Sub

Private Sub cmdbul_Click()
    On Local Error Resume Next
    Dim lngSonuc As Long
    Dim lngYer As Long
    Dim intSecenekler As Integer
    ' arama seçeneklerini ayarla
    If chkTumKelime.Value = 1 Then intSecenekler = intSecenekler + 2
    If chkDuyarli.Value = 1 Then intSecenekler = intSecenekler + 4

    If cmdBul.Caption = "&Bul" Then 'eðer ilk kez aranýlýyorsa
        ' aranan kelimenin terini bul
        lngSonuc = frmAna.rchHtml.Find(cmoAranacakKelime.Text, 0, , intSecenekler)

        If lngSonuc = -1 Then 'aranýlan metin bulunamadý
            MsgBox "Aradýðýnýz kelime bulunamadý.", vbExclamation + vbOKOnly, "[SCT] Hýzlý HTML v1.3" 'mesajý gözterenzi
            cmdBul.Caption = "&Bul" 'baþlýðý ayarla
        Else 'aranýlan metin bulundu
            frmAna.rchHtml.SetFocus 'rchHtmlye odaklan
            cmdDegistir.Enabled = True 'deðiþtir tuþunu etkinleþtir
            cmdTumunuDegistir.Enabled = True 'tümünü deðiþtir tuþunu etkinleþtir
            cmdBul.Caption = "&Sonrakini Bul" 'baþlýðý ayarla
        End If
    Else 'Sonrakini bul
        lngYer = frmAna.rchHtml.SelStart + frmAna.rchHtml.SelLength
        lngSonuc = frmAna.rchHtml.Find(cmoAranacakKelime.Text, lngYer, , intSecenekler)

        If lngSonuc = -1 Then 'aranýlan metin bulunamadý
            MsgBox "Aradýðýnýz kelime bulunamadý.", vbExclamation + vbOKOnly, "[SCT] Hýzlý HTML v1.3" 'mesajý gözterenzi
            cmdBul.Caption = "&Bul" 'baþlýðý ayarla
            cmdDegistir.Enabled = False 'deðiþtir tuþunu etkinleþtirme
            cmdTumunuDegistir.Enabled = False 'tümünü deðiþtir tuþunu etkinleþtirme
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
    
    If cmdDegistir.Caption = "&Deðiþtir..." Then 'deðiþtiri göster
        cmdDegistir.Caption = "&Deðiþtir" 'baþlýðý ayarla
        lblDegistir.Visible = True 'lblDeðiþtiri göster
        cmoDegistir.Visible = True 'cmodegistiri göster
        cmdTumunuDegistir.Visible = True 'cmdtumunudegistiri göster
        Exit Sub
    End If

    ' arama seçeneklerini ayarla
    If chkTumKelime.Value = 1 Then intSecenekler = intSecenekler + 2
    If chkDuyarli.Value = 1 Then intSecenekler = intSecenekler + 4
    
    With frmAna
        .rchHtml.SelText = cmoDegistir.Text 'Metni deðiþtir
        ' sonrakini bul
        lngYer = .rchHtml.SelStart + .rchHtml.SelLength
        ' aranýlan kelimenin yerini bul
        lngSonuc = .rchHtml.Find(cmoAranacakKelime.Text, lngYer, , intSecenekler)

        If lngSonuc = -1 Then 'aranýlan metin bulunamadý
            MsgBox "Aradýðýnýz kelime bulunamadý.", vbExclamation + vbOKOnly, "[SCT] Hýzlý HTML v1.3" 'mesajý gözterenzi
            cmdBul.Caption = "&Bul" 'baþlýðý ayarla
            cmdDegistir.Enabled = False 'deðiþtir tuþunu etkinleþtirme
            cmdTumunuDegistir.Enabled = False 'tümünü deðiþtir tuþunu etkinleþtirme
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
    ' arama seçeneklerini ayarla
    If chkTumKelime.Value = 1 Then intSecenekler = intSecenekler + 2
    If chkDuyarli.Value = 1 Then intSecenekler = intSecenekler + 4

    intSay = 0
    lngYer = 0
    With frmAna
        Do
            If .rchHtml.Find(cmoAranacakKelime.Text, lngYer, , intSecenekler) = -1 Then 'metin bulunamadý
                If intSay > 0 Then 'kaç tane yer deðiþtirilmesi yapýldýðýný göster
                    MsgBox "Belirtiðiniz alan tarandý." & intSay & " deðiþtirilme yapýldý.", vbInformation + vbOKOnly, "[SCT] Hýzlý HTML v1.3"
                End If
                cmdBul.Caption = "&Bul" 'baþlýðý ayarla
                cmdDegistir.Enabled = False 'deðiþtir tuþunu etkinleþtirme
                cmdTumunuDegistir.Enabled = False 'tümünü deðiþtir tuþunu etkinleþtirme
                Exit Do
            Else 'metin bulunanzi
                lngYer = .rchHtml.SelStart + .rchHtml.SelLength
                intSay = intSay + 1 'sayacý birer arttýr
                .rchHtml.SelText = cmoDegistir.Text 'Metni deðiþtir
            End If
        Loop
    End With
End Sub

