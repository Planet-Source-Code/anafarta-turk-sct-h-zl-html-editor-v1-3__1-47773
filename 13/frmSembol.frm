VERSION 5.00
Begin VB.Form frmSembol 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Özel Karakter Editörü"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "frmSembol.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtKodHTML 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin HizliHtml.Command cmdKapat 
      Height          =   325
      Left            =   5040
      TabIndex        =   3
      Top             =   3105
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      Caption         =   "Kapat"
   End
   Begin HizliHtml.Command cmdKodEkle 
      Height          =   325
      Left            =   120
      TabIndex        =   2
      Top             =   3105
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      Caption         =   "Kodu Ekle"
   End
   Begin VB.ListBox lstSembol 
      Columns         =   20
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      ItemData        =   "frmSembol.frx":1042
      Left            =   240
      List            =   "frmSembol.frx":1044
      TabIndex        =   1
      Top             =   480
      Width           =   6255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Kullanmak Ýstediðiniz Özel Karaktere Çift Týklayýnýz"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmSembol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdKapat_Click()
Unload Me
End Sub

Private Sub cmdKodEkle_Click()
EtiketEkle txtKodHTML.Text, True
frmAna.ImlecYerlestir frmAna.rchHtml.SelText, 0
End Sub

Private Sub Form_Load()
lstSembol.AddItem "" '&nbsp;
lstSembol.AddItem "¡" '&iexcl;
lstSembol.AddItem "¢" '&cent;
lstSembol.AddItem "£" '&pound;
lstSembol.AddItem "¤" '&curren;
lstSembol.AddItem "¥" '&yen;
lstSembol.AddItem "¦" '&brvbar;
lstSembol.AddItem "§" '&sect;
lstSembol.AddItem "¨" '&uml;
lstSembol.AddItem "©" '&copy;
lstSembol.AddItem "ª" '&ordf;
lstSembol.AddItem "«" '&laquo;
lstSembol.AddItem "¬" '&not;
lstSembol.AddItem "­" '&shy;
lstSembol.AddItem "®" '&reg;
lstSembol.AddItem "¯" '&macr;
lstSembol.AddItem "º" '&deg;
lstSembol.AddItem "±" '&plusmn;
lstSembol.AddItem "²" '&sup2;
lstSembol.AddItem "³" '&sup3;
lstSembol.AddItem "´" '&acute;
lstSembol.AddItem "µ" '&micro;
lstSembol.AddItem "¶" '&para;
lstSembol.AddItem "·" '&middot;
lstSembol.AddItem "¸" '&cedil;
lstSembol.AddItem "¹" '&sup1;
lstSembol.AddItem "º" '&ordm;
lstSembol.AddItem Chr(34) '&quot;
lstSembol.AddItem "<" '&lt;
lstSembol.AddItem ">" '&gt;
lstSembol.AddItem "&" '&amp;
lstSembol.AddItem "¿" '&iquest;
lstSembol.AddItem "¾" '&frac34;
lstSembol.AddItem "½" '&frac12;
lstSembol.AddItem "¼" '&frac14;
lstSembol.AddItem "»" '&raquo;
lstSembol.AddItem "À" '&Agrave;
lstSembol.AddItem "Á" '&Aacute;
lstSembol.AddItem "Â" '&Acirc;
lstSembol.AddItem "Ã" '&Atilde;
lstSembol.AddItem "Ä" '&Auml;
lstSembol.AddItem "Å" '&Aring;
lstSembol.AddItem "Æ" '&AElig;
lstSembol.AddItem "Ç" '&Ccedil;
lstSembol.AddItem "È" '&Egrave;
lstSembol.AddItem "É" '&Eacute;
lstSembol.AddItem "Ê" '&Ecirc;
lstSembol.AddItem "Ë" '&Euml;
lstSembol.AddItem "Ì" '&Igrave;
lstSembol.AddItem "Í" '&Iacute;
lstSembol.AddItem "Î" '&Icirc;
lstSembol.AddItem "Ï" '&Iuml;
lstSembol.AddItem "Ð" '&ETH;
lstSembol.AddItem "Ñ" '&Ntilde;
lstSembol.AddItem "Ò" '&Ograve;
lstSembol.AddItem "Ó" '&Oacute;
lstSembol.AddItem "Ô" '&Ocirc;
lstSembol.AddItem "Õ" '&Otilde;
lstSembol.AddItem "Ö" '&Ouml;
lstSembol.AddItem "Œ" '&OElig;
lstSembol.AddItem "×" '&times;
lstSembol.AddItem "Ø" '&Oslash;
lstSembol.AddItem "Ù" '&Ugrave;
lstSembol.AddItem "Ú" '&Uacute;
lstSembol.AddItem "Û" '&Ucirc;
lstSembol.AddItem "Ü" '&Uuml;
lstSembol.AddItem "Ý" '&Yacute;
lstSembol.AddItem "Þ" '&THORN;
lstSembol.AddItem "ß" '&szlig;
lstSembol.AddItem "à" '&agrave;
lstSembol.AddItem "á" '&aacute;
lstSembol.AddItem "â" '&acirc;
lstSembol.AddItem "ã" '&atilde;
lstSembol.AddItem "ä" '&auml;
lstSembol.AddItem "å" '&aring;
lstSembol.AddItem "æ" '&aelig;
lstSembol.AddItem "ç" '&ccedil;
lstSembol.AddItem "è" '&egrave;
lstSembol.AddItem "é" '&eacute;
lstSembol.AddItem "ê" '&ecirc;
lstSembol.AddItem "ë" '&euml;
lstSembol.AddItem "ì" '&igrave;
lstSembol.AddItem "í" '&iacute;
lstSembol.AddItem "î" '&icirc;
lstSembol.AddItem "ï" '&iuml;
lstSembol.AddItem "÷" '&divide;
lstSembol.AddItem "ö" '&ouml;
lstSembol.AddItem "õ" '&otilde;
lstSembol.AddItem "ô" '&ocirc;
lstSembol.AddItem "ó" '&oacute;
lstSembol.AddItem "ò" '&ograve;
lstSembol.AddItem "œ" '&oelig;
lstSembol.AddItem "ñ" '&ntilde;
lstSembol.AddItem "ð" '&eth;
lstSembol.AddItem "ø" '&oslash;
lstSembol.AddItem "ù" '&ugrave;
lstSembol.AddItem "ú" '&uacute;
lstSembol.AddItem "û" '&ucirc;
lstSembol.AddItem "ü" '&uuml;
lstSembol.AddItem "ý" '&yacute;
lstSembol.AddItem "þ" '&thorn;
lstSembol.AddItem "ÿ" '&yuml;
lstSembol.AddItem "™" '&trade;
End Sub

Private Sub lstSembol_Click()
If lstSembol.Text = "" Then
    txtKodHTML.Text = "&nbsp;"
ElseIf lstSembol.Text = "¡" Then
    txtKodHTML.Text = "&iexcl;"
ElseIf lstSembol.Text = "¢" Then
    txtKodHTML.Text = "&cent;"
ElseIf lstSembol.Text = "£" Then
    txtKodHTML.Text = "&pound;"
ElseIf lstSembol.Text = "¤" Then
    txtKodHTML.Text = "&curren;"
ElseIf lstSembol.Text = "¥" Then
    txtKodHTML.Text = "&yen;"
ElseIf lstSembol.Text = "¦" Then
    txtKodHTML.Text = "&brvbar;"
ElseIf lstSembol.Text = "§" Then
    txtKodHTML.Text = "&sect;"
ElseIf lstSembol.Text = "¨" Then
    txtKodHTML.Text = "&uml;"
ElseIf lstSembol.Text = "©" Then
    txtKodHTML.Text = "&copy;"
ElseIf lstSembol.Text = "ª" Then
    txtKodHTML.Text = "&ordf;"
ElseIf lstSembol.Text = "«" Then
    txtKodHTML.Text = "&laquo;"
ElseIf lstSembol.Text = "¬" Then
    txtKodHTML.Text = "&not;"
ElseIf lstSembol.Text = "­" Then
    txtKodHTML.Text = "&shy;"
ElseIf lstSembol.Text = "®" Then
    txtKodHTML.Text = "&reg;"
ElseIf lstSembol.Text = "¯" Then
    txtKodHTML.Text = "&macr;"
ElseIf lstSembol.Text = "º" Then
    txtKodHTML.Text = "&deg;"
ElseIf lstSembol.Text = "±" Then
    txtKodHTML.Text = "&plusmn;"
ElseIf lstSembol.Text = "²" Then
    txtKodHTML.Text = "&sup2;"
ElseIf lstSembol.Text = "³" Then
    txtKodHTML.Text = "&sup3;"
ElseIf lstSembol.Text = "´" Then
    txtKodHTML.Text = "&acute;"
ElseIf lstSembol.Text = "µ" Then
    txtKodHTML.Text = "&micro;"
ElseIf lstSembol.Text = "¶" Then
    txtKodHTML.Text = "&para;"
ElseIf lstSembol.Text = "·" Then
    txtKodHTML.Text = "&middot;"
ElseIf lstSembol.Text = "¸" Then
    txtKodHTML.Text = "&cedil;"
ElseIf lstSembol.Text = "¹" Then
    txtKodHTML.Text = "&sup1;"
ElseIf lstSembol.Text = "º" Then
    txtKodHTML.Text = "&ordm;"
ElseIf lstSembol.Text = Chr(34) Then
    txtKodHTML.Text = "&quot;"
ElseIf lstSembol.Text = "<" Then
    txtKodHTML.Text = "&lt;"
ElseIf lstSembol.Text = ">" Then
    txtKodHTML.Text = "&gt;"
ElseIf lstSembol.Text = "&" Then
    txtKodHTML.Text = "&amp;"
ElseIf lstSembol.Text = "¿" Then
    txtKodHTML.Text = "&iquest;"
ElseIf lstSembol.Text = "¾" Then
    txtKodHTML.Text = "&frac34;"
ElseIf lstSembol.Text = "½" Then
    txtKodHTML.Text = "&frac12;"
ElseIf lstSembol.Text = "¼" Then
    txtKodHTML.Text = "&frac14;"
ElseIf lstSembol.Text = "»" Then
    txtKodHTML.Text = "&raquo;"
ElseIf lstSembol.Text = "À" Then
    txtKodHTML.Text = "&Agrave;"
ElseIf lstSembol.Text = "Á" Then
    txtKodHTML.Text = "&Aacute;"
ElseIf lstSembol.Text = "Â" Then
    txtKodHTML.Text = "&Acirc;"
ElseIf lstSembol.Text = "Ã" Then
    txtKodHTML.Text = "&Atilde;"
ElseIf lstSembol.Text = "Ä" Then
    txtKodHTML.Text = "&Auml;"
ElseIf lstSembol.Text = "Å" Then
    txtKodHTML.Text = "&Aring;"
ElseIf lstSembol.Text = "Æ" Then
    txtKodHTML.Text = "&AElig;"
ElseIf lstSembol.Text = "Ç" Then
    txtKodHTML.Text = "&Ccedil;"
ElseIf lstSembol.Text = "È" Then
    txtKodHTML.Text = "&Egrave;"
ElseIf lstSembol.Text = "É" Then
    txtKodHTML.Text = "&Eacute;"
ElseIf lstSembol.Text = "Ê" Then
    txtKodHTML.Text = "&Ecirc;"
ElseIf lstSembol.Text = "Ë" Then
    txtKodHTML.Text = "&Euml;"
ElseIf lstSembol.Text = "Ì" Then
    txtKodHTML.Text = "&Igrave;"
ElseIf lstSembol.Text = "Í" Then
    txtKodHTML.Text = "&Iacute;"
ElseIf lstSembol.Text = "Î" Then
    txtKodHTML.Text = "&Icirc;"
ElseIf lstSembol.Text = "Ï" Then
    txtKodHTML.Text = "&Iuml;"
ElseIf lstSembol.Text = "Ð" Then
    txtKodHTML.Text = "&ETH;"
ElseIf lstSembol.Text = "Ñ" Then
    txtKodHTML.Text = "&Ntilde;"
ElseIf lstSembol.Text = "Ò" Then
    txtKodHTML.Text = "&Ograve;"
ElseIf lstSembol.Text = "Ó" Then
    txtKodHTML.Text = "&Oacute;"
ElseIf lstSembol.Text = "Ô" Then
    txtKodHTML.Text = "&Ocirc;"
ElseIf lstSembol.Text = "Õ" Then
    txtKodHTML.Text = "&Otilde;"
ElseIf lstSembol.Text = "Ö" Then
    txtKodHTML.Text = "&Ouml;"
ElseIf lstSembol.Text = "Œ" Then
    txtKodHTML.Text = "&OElig;"
ElseIf lstSembol.Text = "×" Then
    txtKodHTML.Text = "&times;"
ElseIf lstSembol.Text = "Ø" Then
    txtKodHTML.Text = "&Oslash;"
ElseIf lstSembol.Text = "Ù" Then
    txtKodHTML.Text = "&Ugrave;"
ElseIf lstSembol.Text = "Ú" Then
    txtKodHTML.Text = "&Uacute;"
ElseIf lstSembol.Text = "Û" Then
    txtKodHTML.Text = "&Ucirc;"
ElseIf lstSembol.Text = "Ü" Then
    txtKodHTML.Text = "&Uuml;"
ElseIf lstSembol.Text = "Ý" Then
    txtKodHTML.Text = "&Yacute;"
ElseIf lstSembol.Text = "Þ" Then
    txtKodHTML.Text = "&THORN;"
ElseIf lstSembol.Text = "ß" Then
    txtKodHTML.Text = "&szlig;"
ElseIf lstSembol.Text = "à" Then
    txtKodHTML.Text = "&agrave;"
ElseIf lstSembol.Text = "á" Then
    txtKodHTML.Text = "&aacute;"
ElseIf lstSembol.Text = "â" Then
    txtKodHTML.Text = "&acirc;"
ElseIf lstSembol.Text = "ã" Then
    txtKodHTML.Text = "&atilde;"
ElseIf lstSembol.Text = "ä" Then
    txtKodHTML.Text = "&auml;"
ElseIf lstSembol.Text = "å" Then
    txtKodHTML.Text = "&aring;"
ElseIf lstSembol.Text = "æ" Then
    txtKodHTML.Text = "&aelig;"
ElseIf lstSembol.Text = "ç" Then
    txtKodHTML.Text = "&ccedil;"
ElseIf lstSembol.Text = "è" Then
    txtKodHTML.Text = "&egrave;"
ElseIf lstSembol.Text = "é" Then
    txtKodHTML.Text = "&eacute;"
ElseIf lstSembol.Text = "ê" Then
    txtKodHTML.Text = "&ecirc;"
ElseIf lstSembol.Text = "ë" Then
    txtKodHTML.Text = "&euml;"
ElseIf lstSembol.Text = "ì" Then
    txtKodHTML.Text = "&igrave;"
ElseIf lstSembol.Text = "í" Then
    txtKodHTML.Text = "&iacute;"
ElseIf lstSembol.Text = "î" Then
    txtKodHTML.Text = "&icirc;"
ElseIf lstSembol.Text = "ï" Then
    txtKodHTML.Text = "&iuml;"
ElseIf lstSembol.Text = "÷" Then
    txtKodHTML.Text = "&divide;"
ElseIf lstSembol.Text = "ö" Then
    txtKodHTML.Text = "&ouml;"
ElseIf lstSembol.Text = "õ" Then
    txtKodHTML.Text = "&otilde;"
ElseIf lstSembol.Text = "ô" Then
    txtKodHTML.Text = "&ocirc;"
ElseIf lstSembol.Text = "ó" Then
    txtKodHTML.Text = "&oacute;"
ElseIf lstSembol.Text = "ò" Then
    txtKodHTML.Text = "&ograve;"
ElseIf lstSembol.Text = "œ" Then
    txtKodHTML.Text = "&oelig;"
ElseIf lstSembol.Text = "ñ" Then
    txtKodHTML.Text = "&ntilde;"
ElseIf lstSembol.Text = "ð" Then
    txtKodHTML.Text = "&eth;"
ElseIf lstSembol.Text = "ø" Then
    txtKodHTML.Text = "&oslash;"
ElseIf lstSembol.Text = "ù" Then
    txtKodHTML.Text = "&ugrave;"
ElseIf lstSembol.Text = "ú" Then
    txtKodHTML.Text = "&uacute;"
ElseIf lstSembol.Text = "û" Then
    txtKodHTML.Text = "&ucirc;"
ElseIf lstSembol.Text = "ü" Then
    txtKodHTML.Text = "&uuml;"
ElseIf lstSembol.Text = "ý" Then
    txtKodHTML.Text = "&yacute;"
ElseIf lstSembol.Text = "þ" Then
    txtKodHTML.Text = "&thorn;"
ElseIf lstSembol.Text = "ÿ" Then
    txtKodHTML.Text = "&yuml;"
ElseIf lstSembol.Text = "™" Then
    txtKodHTML.Text = "&trade;"
End If
End Sub
