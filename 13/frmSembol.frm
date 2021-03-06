VERSION 5.00
Begin VB.Form frmSembol 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�zel Karakter Edit�r�"
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
      Caption         =   "Kullanmak �stedi�iniz �zel Karaktere �ift T�klay�n�z"
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
lstSembol.AddItem "�" '&iexcl;
lstSembol.AddItem "�" '&cent;
lstSembol.AddItem "�" '&pound;
lstSembol.AddItem "�" '&curren;
lstSembol.AddItem "�" '&yen;
lstSembol.AddItem "�" '&brvbar;
lstSembol.AddItem "�" '&sect;
lstSembol.AddItem "�" '&uml;
lstSembol.AddItem "�" '&copy;
lstSembol.AddItem "�" '&ordf;
lstSembol.AddItem "�" '&laquo;
lstSembol.AddItem "�" '&not;
lstSembol.AddItem "�" '&shy;
lstSembol.AddItem "�" '&reg;
lstSembol.AddItem "�" '&macr;
lstSembol.AddItem "�" '&deg;
lstSembol.AddItem "�" '&plusmn;
lstSembol.AddItem "�" '&sup2;
lstSembol.AddItem "�" '&sup3;
lstSembol.AddItem "�" '&acute;
lstSembol.AddItem "�" '&micro;
lstSembol.AddItem "�" '&para;
lstSembol.AddItem "�" '&middot;
lstSembol.AddItem "�" '&cedil;
lstSembol.AddItem "�" '&sup1;
lstSembol.AddItem "�" '&ordm;
lstSembol.AddItem Chr(34) '&quot;
lstSembol.AddItem "<" '&lt;
lstSembol.AddItem ">" '&gt;
lstSembol.AddItem "&" '&amp;
lstSembol.AddItem "�" '&iquest;
lstSembol.AddItem "�" '&frac34;
lstSembol.AddItem "�" '&frac12;
lstSembol.AddItem "�" '&frac14;
lstSembol.AddItem "�" '&raquo;
lstSembol.AddItem "�" '&Agrave;
lstSembol.AddItem "�" '&Aacute;
lstSembol.AddItem "�" '&Acirc;
lstSembol.AddItem "�" '&Atilde;
lstSembol.AddItem "�" '&Auml;
lstSembol.AddItem "�" '&Aring;
lstSembol.AddItem "�" '&AElig;
lstSembol.AddItem "�" '&Ccedil;
lstSembol.AddItem "�" '&Egrave;
lstSembol.AddItem "�" '&Eacute;
lstSembol.AddItem "�" '&Ecirc;
lstSembol.AddItem "�" '&Euml;
lstSembol.AddItem "�" '&Igrave;
lstSembol.AddItem "�" '&Iacute;
lstSembol.AddItem "�" '&Icirc;
lstSembol.AddItem "�" '&Iuml;
lstSembol.AddItem "�" '&ETH;
lstSembol.AddItem "�" '&Ntilde;
lstSembol.AddItem "�" '&Ograve;
lstSembol.AddItem "�" '&Oacute;
lstSembol.AddItem "�" '&Ocirc;
lstSembol.AddItem "�" '&Otilde;
lstSembol.AddItem "�" '&Ouml;
lstSembol.AddItem "�" '&OElig;
lstSembol.AddItem "�" '&times;
lstSembol.AddItem "�" '&Oslash;
lstSembol.AddItem "�" '&Ugrave;
lstSembol.AddItem "�" '&Uacute;
lstSembol.AddItem "�" '&Ucirc;
lstSembol.AddItem "�" '&Uuml;
lstSembol.AddItem "�" '&Yacute;
lstSembol.AddItem "�" '&THORN;
lstSembol.AddItem "�" '&szlig;
lstSembol.AddItem "�" '&agrave;
lstSembol.AddItem "�" '&aacute;
lstSembol.AddItem "�" '&acirc;
lstSembol.AddItem "�" '&atilde;
lstSembol.AddItem "�" '&auml;
lstSembol.AddItem "�" '&aring;
lstSembol.AddItem "�" '&aelig;
lstSembol.AddItem "�" '&ccedil;
lstSembol.AddItem "�" '&egrave;
lstSembol.AddItem "�" '&eacute;
lstSembol.AddItem "�" '&ecirc;
lstSembol.AddItem "�" '&euml;
lstSembol.AddItem "�" '&igrave;
lstSembol.AddItem "�" '&iacute;
lstSembol.AddItem "�" '&icirc;
lstSembol.AddItem "�" '&iuml;
lstSembol.AddItem "�" '&divide;
lstSembol.AddItem "�" '&ouml;
lstSembol.AddItem "�" '&otilde;
lstSembol.AddItem "�" '&ocirc;
lstSembol.AddItem "�" '&oacute;
lstSembol.AddItem "�" '&ograve;
lstSembol.AddItem "�" '&oelig;
lstSembol.AddItem "�" '&ntilde;
lstSembol.AddItem "�" '&eth;
lstSembol.AddItem "�" '&oslash;
lstSembol.AddItem "�" '&ugrave;
lstSembol.AddItem "�" '&uacute;
lstSembol.AddItem "�" '&ucirc;
lstSembol.AddItem "�" '&uuml;
lstSembol.AddItem "�" '&yacute;
lstSembol.AddItem "�" '&thorn;
lstSembol.AddItem "�" '&yuml;
lstSembol.AddItem "�" '&trade;
End Sub

Private Sub lstSembol_Click()
If lstSembol.Text = "" Then
    txtKodHTML.Text = "&nbsp;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&iexcl;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&cent;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&pound;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&curren;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&yen;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&brvbar;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&sect;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&uml;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&copy;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&ordf;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&laquo;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&not;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&shy;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&reg;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&macr;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&deg;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&plusmn;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&sup2;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&sup3;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&acute;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&micro;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&para;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&middot;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&cedil;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&sup1;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&ordm;"
ElseIf lstSembol.Text = Chr(34) Then
    txtKodHTML.Text = "&quot;"
ElseIf lstSembol.Text = "<" Then
    txtKodHTML.Text = "&lt;"
ElseIf lstSembol.Text = ">" Then
    txtKodHTML.Text = "&gt;"
ElseIf lstSembol.Text = "&" Then
    txtKodHTML.Text = "&amp;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&iquest;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&frac34;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&frac12;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&frac14;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&raquo;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Agrave;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Aacute;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Acirc;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Atilde;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Auml;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Aring;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&AElig;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Ccedil;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Egrave;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Eacute;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Ecirc;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Euml;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Igrave;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Iacute;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Icirc;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Iuml;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&ETH;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Ntilde;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Ograve;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Oacute;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Ocirc;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Otilde;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Ouml;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&OElig;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&times;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Oslash;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Ugrave;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Uacute;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Ucirc;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Uuml;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&Yacute;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&THORN;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&szlig;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&agrave;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&aacute;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&acirc;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&atilde;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&auml;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&aring;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&aelig;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&ccedil;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&egrave;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&eacute;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&ecirc;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&euml;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&igrave;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&iacute;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&icirc;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&iuml;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&divide;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&ouml;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&otilde;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&ocirc;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&oacute;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&ograve;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&oelig;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&ntilde;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&eth;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&oslash;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&ugrave;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&uacute;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&ucirc;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&uuml;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&yacute;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&thorn;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&yuml;"
ElseIf lstSembol.Text = "�" Then
    txtKodHTML.Text = "&trade;"
End If
End Sub
