VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAna 
   Caption         =   "[SCT] Hýzlý HTML Editörü v1.3 : Baþlýksýz"
   ClientHeight    =   7515
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAna.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImgLstArac 
      Left            =   120
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":1042
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":15DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":1B7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":2116
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":26B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":2DAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":34AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":3BA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":42A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":499E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":4F3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":54D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":5A72
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":674E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":742A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgLstBaslik 
      Left            =   120
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":8106
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":86A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":8C3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":91DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":9776
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":9D12
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar durum1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   4
      Top             =   7245
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13582
            MinWidth        =   13582
            Text            =   "Yardým için F1 'e basýn"
            TextSave        =   "Yardým için F1 'e basýn"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1535
            MinWidth        =   1534
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1535
            MinWidth        =   1534
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2119
            MinWidth        =   2119
            TextSave        =   "23:30"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdDummy 
      Height          =   255
      Left            =   -840
      TabIndex        =   1
      Top             =   5640
      Width           =   615
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   600
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rchHtml 
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"frmAna.frx":A2AE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   1296
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Genel"
      TabPicture(0)   =   "frmAna.frx":A37C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tbrAna"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ImgLstGenTool"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Metin"
      TabPicture(1)   =   "frmAna.frx":A398
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tbrMetin"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Baþlýklar"
      TabPicture(2)   =   "frmAna.frx":A3B4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tbrBaslik"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Araçlar"
      TabPicture(3)   =   "frmAna.frx":A3D0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "tbrArac"
      Tab(3).ControlCount=   1
      Begin VB.PictureBox ImgLstGenTool 
         BackColor       =   &H000000FF&
         Height          =   1000
         Left            =   4080
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   0
         Top             =   3840
         Width           =   1000
      End
      Begin MSComctlLib.Toolbar tbrAna 
         Height          =   390
         Left            =   60
         TabIndex        =   5
         Top             =   15
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "imgStandard"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   16
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Yeni"
               Object.ToolTipText     =   "Yeni"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Ac"
               Object.ToolTipText     =   "Aç"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Kaydet"
               Object.ToolTipText     =   "Kaydet"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   4
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Yazdir"
               Object.ToolTipText     =   "Yazdýr"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "GeriAl"
               Object.ToolTipText     =   "Geri Al"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "IleriAl"
               Object.ToolTipText     =   "Ýleri Al"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Kes"
               Object.ToolTipText     =   "Kes"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Kopyala"
               Object.ToolTipText     =   "Kopyala"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Yapistir"
               Object.ToolTipText     =   "Yapýþtýr"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bul"
               Object.ToolTipText     =   "Bul"
               ImageIndex      =   16
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Yardim"
               Object.ToolTipText     =   "[SCT] Hýzlý HTML v1.3 Yardýmý"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Hakkinda"
               Object.ToolTipText     =   "[SCT] Hýzlý HTML v1.3 Hakkýnda"
               ImageIndex      =   15
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbrMetin 
         Height          =   390
         Left            =   -74940
         TabIndex        =   6
         Top             =   15
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "ImgLstMetin"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   17
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Kalin"
               Object.ToolTipText     =   "Kalýn"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Egik"
               Object.ToolTipText     =   "Eðik"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "AltCizgili"
               Object.ToolTipText     =   "Alt Çizgili"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "YaziTipi"
               Object.ToolTipText     =   "Yazý Tipi"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "YTipEksi1"
               Object.ToolTipText     =   "Yazý Tipi -1"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "YTipArti1"
               Object.ToolTipText     =   "Yazý Tipi +1"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SuperScript"
               Object.ToolTipText     =   "Süper Script"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SubScript"
               Object.ToolTipText     =   "Sub Script"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Small"
               Object.ToolTipText     =   "Small Etiketi"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Big"
               Object.ToolTipText     =   "Big Etiketi"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "KayanYazi"
               Object.ToolTipText     =   "Kayan Yazý Editörü"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "TTFont"
               Object.ToolTipText     =   "TeleType Yazý Tipi"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Vurgu"
               Object.ToolTipText     =   "Vurgu Etiketi <em>"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Strong"
               Object.ToolTipText     =   "Strong Etiketi"
               ImageIndex      =   13
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbrBaslik 
         Height          =   390
         Left            =   -74940
         TabIndex        =   7
         Top             =   15
         Width           =   11880
         _ExtentX        =   20955
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "ImgLstBaslik"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B1"
               Object.ToolTipText     =   "Baþlýk 1 (En Büyük Baþlýk)"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B2"
               Object.ToolTipText     =   "Baþlýk 2"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B3"
               Object.ToolTipText     =   "Baþlýk 3"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B4"
               Object.ToolTipText     =   "Baþlýk 4"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B5"
               Object.ToolTipText     =   "Baþlýk 5"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "B6"
               Object.ToolTipText     =   "Baþlýk 6 (En Küçük Baþlýk)"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbrArac 
         Height          =   390
         Left            =   -74940
         TabIndex        =   8
         Top             =   15
         Width           =   11880
         _ExtentX        =   20955
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "ImgLstArac"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   18
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Sola"
               Object.ToolTipText     =   "Sola Daya"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Orta"
               Object.ToolTipText     =   "Ortala"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Saga"
               Object.ToolTipText     =   "Saða Daya"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "IkiYana"
               Object.ToolTipText     =   "Ýki Yana Dayalý"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bosluk"
               Object.ToolTipText     =   "Boþluk ( &nbsp; )"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Break"
               Object.ToolTipText     =   "Bir Satýr Aþaðýya (Break)"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paragraf"
               Object.ToolTipText     =   "Paragraf"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Div"
               Object.ToolTipText     =   "Bölüm Yarat (Division)"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cizik"
               Object.ToolTipText     =   "Yatay Çizgi"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "UstCizgili"
               Object.ToolTipText     =   "Üst Çizgili Metin"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   12
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bul"
               Object.ToolTipText     =   "Bul"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Yorum"
               Object.ToolTipText     =   "Yorum Ekle"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Resim"
               Object.ToolTipText     =   "Resim Ekle"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EPosta"
               Object.ToolTipText     =   "E-Posta Linki"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Link"
               Object.ToolTipText     =   "Link Ekle"
               ImageIndex      =   15
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImgLstMetin 
      Left            =   120
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":A3EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":A70C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":AA40
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":B194
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":B544
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":B920
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":BDA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":C0FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":C21C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":C33C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":C45C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":C858
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":CCAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":D100
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":D44C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":D7C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":DB3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":DE78
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":E23C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":E648
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":EA54
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgStandard 
      Left            =   120
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":EE88
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":EFE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":F140
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":F29C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":F3F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":F554
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":F8F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":FA4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":FBA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":FD04
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":FE60
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":FFBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":10198
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":10374
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":104D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAna.frx":11524
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuDosya 
      Caption         =   "&Dosya"
      Begin VB.Menu mnuYeni 
         Caption         =   "Yeni"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuAc 
         Caption         =   "Aç"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnukesme1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKaydet 
         Caption         =   "Kaydet"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFakliKaydet 
         Caption         =   "Farklý Kaydet"
      End
      Begin VB.Menu mnukesme2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuYazdir 
         Caption         =   "Yazdýr"
      End
      Begin VB.Menu mnuKesme6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOnizleme 
         Caption         =   "Önizleme"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuKesme3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCikis 
         Caption         =   "Çýkýþ"
      End
   End
   Begin VB.Menu mnuDuzen 
      Caption         =   "Dü&zen"
      Begin VB.Menu mnuGeriAl 
         Caption         =   "Geri Al"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuIleriAl 
         Caption         =   "Ýleri Al"
      End
      Begin VB.Menu mnuKesme5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKes 
         Caption         =   "Kes"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuKopyala 
         Caption         =   "Kopyala"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuYapistir 
         Caption         =   "Yapýþtýr"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSil 
         Caption         =   "Sil"
      End
      Begin VB.Menu mnuKesme4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTumunuSec 
         Caption         =   "Tümünü Seç"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuBuul 
      Caption         =   "&Bul"
      Begin VB.Menu mnuBul 
         Caption         =   "Bul"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuDegistir 
         Caption         =   "Deðiþtir"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuArac 
      Caption         =   "&Araçlar"
      Begin VB.Menu mnuKayanYazi 
         Caption         =   "Kayan Yazý Editörü"
      End
      Begin VB.Menu mnuSembol 
         Caption         =   "Sembol Editörü"
      End
      Begin VB.Menu mnuResimEditoru 
         Caption         =   "Resim Editörü"
      End
      Begin VB.Menu mnuListeEditoru 
         Caption         =   "Liste Editörü"
      End
      Begin VB.Menu mnuLinkEdit 
         Caption         =   "Link Editörü"
      End
   End
   Begin VB.Menu mnuYardim 
      Caption         =   "&Yardým"
      Begin VB.Menu mnuHHtmlYardim 
         Caption         =   "Hýzlý Html Yardýmý"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHakkinda 
         Caption         =   "Hýzlý Html Hakkýnda"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmAna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DosyaDegisirse As Boolean
Public AcilanDosya As String
'geri al ileri al için gerekli olan deðiþkenler
Dim gdysFarkiOnemseme As Boolean
Dim gsylIcerik As Integer
Dim gyzlYigin(1000) As String

Private Sub Form_Load()
'Form açýldýðýnda yeni sayfa açýlmýþ olur
YeniSayfa
'bununlada rchHtml'ye yazýlan html kodlarýnýn
'yeni bir sayfada olduðunu belirtiyoruz.
DosyaDegisirse = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Local Error Resume Next
    Dim cvp
    If DosyaDegisirse = True Then
        cvp = MsgBox("Dosya içeriði deðiþti." & vbCrLf & "Kaydetmek istiyor musunuz?", vbInformation + vbYesNoCancel, "[SCT] Hýzlý HTML v1.3")
            If cvp = vbCancel Then
                Cancel = 1
            ElseIf cvp = vbNo Then
                Cancel = 0
            Else
                mnuKaydet_Click
                Unload Me
            End If
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
SSTab1.Move 30, 50, Me.Width - 230, 735
rchHtml.Move 30, SSTab1.Top + SSTab1.Height, Me.Width - 230, Me.Height - (SSTab1.Top + SSTab1.Height + 950)
End Sub

Private Sub mnuAc_Click()
On Local Error Resume Next
    Dim cvp
    If DosyaDegisirse = True Then
        cvp = MsgBox("Dosya içeriði deðiþti." & vbCrLf & "Kaydetmek istiyor musunuz?", vbInformation + vbYesNoCancel, "[SCT] Hýzlý HTML v1.3")
            If cvp = vbYes Then
                mnuKaydet_Click
                VarolanDosyayiAc
            ElseIf cvp = vbNo Then
                VarolanDosyayiAc
            Else
                Exit Sub
            End If
    Else
        VarolanDosyayiAc
    End If
End Sub

Private Sub VarolanDosyayiAc()
'CD1'den dosya ismini alýr..
    cd1.DialogTitle = "Aç"
    cd1.Filter = " Web Dosyalarý ( *.htm, *.html, *.css) | *.htm; *.html; *.css; |"
    cd1.ShowOpen
If Err <> 32755 Then 'kullanýcýnýn Cancel tuþuna baþmadýðýndan emin olalým:)
    'þimdi yeni dosyayý açalým:
    rchHtml.LoadFile cd1.FileName, rtfText
    AcilanDosya = cd1.FileName
    'þimdide baþlýðýmýzý koyalým
    Me.Caption = "[SCT] Hýzlý HTML Editörü v1.3 : " & AcilanDosya
    'Gerialý sýfýrlayalým
    'ResetGeriAl
End If
DosyaDegisirse = False
rchHtml.SetFocus
End Sub

Private Sub mnuBul_Click()
    frmBul.Show 1
End Sub

Private Sub mnuCikis_Click()
Unload Me ' programý kapatýr
End Sub

Private Sub mnuDegistir_Click()
    frmBul.cmdDegistir.Caption = "&Deðiþtir" 'baþlýðý ayarla
    frmBul.lblDegistir.Visible = True 'lblDeðiþtiri göster
    frmBul.cmoDegistir.Visible = True 'cmodegistiri göster
    frmBul.cmdTumunuDegistir.Visible = True 'cmdtumunudegistiri göster
    frmBul.Caption = "Deðiþtir.."
    frmBul.Show 1
End Sub

Private Sub mnuFakliKaydet_Click()
'Dosyayý farklý kaydedebiliriz:)
    cd1.Filter = " Web Dosyalarý ( *.htm, *.html, *.css) | *.htm; *.html; *.css; | Html Dosyalarý ( *.html ) | *.html; | Htm Dosyalarý ( *.htm ) | *.htm; | CSS Dosyalarý ( *.css ) | *.css"
    cd1.DialogTitle = "Farklý Kaydet"
    cd1.ShowSave
    rchHtml.SaveFile cd1.FileName, rtfText
    Me.Caption = "[SCT] Hýzlý HTML Editörü v1.3 : " & cd1.FileName
    DosyaDegisirse = False
    rchHtml.SetFocus
End Sub

Private Sub mnuIleriAl_Click()
    'basit bir ileri al olayý
    gdysFarkiOnemseme = True
    gsylIcerik = gsylIcerik + 1
    On Error Resume Next
    rchHtml.TextRTF = gyzlYigin(gsylIcerik)
    gdysFarkiOnemseme = False
End Sub

Private Sub mnuHakkinda_Click()
frmHakkinda.Show 1
End Sub

Private Sub mnuGeriAl_Click()
    'index 0 sa daha fazla geri al yapamazsýn
    If gsylIcerik = 0 Then Exit Sub
    
    'basit bir geri al olayý.
    gdysFarkiOnemseme = True
    gsylIcerik = gsylIcerik - 1
    On Error Resume Next
    rchHtml.TextRTF = gyzlYigin(gsylIcerik)
    gdysFarkiOnemseme = False
End Sub

Private Sub mnuKayanYazi_Click()
    frmKayanYazi.Show 1
End Sub

Private Sub mnuKaydet_Click()
'Dosyayý kaydedelim
If AcilanDosya = "" Then
    '
    cd1.Filter = " Web Dosyalarý ( *.htm, *.html, *.css) | *.htm; *.html; *.css; | Html Dosyalarý ( *.html ) | *.html; | Htm Dosyalarý ( *.htm ) | *.htm; | CSS Dosyalarý ( *.css ) | *.css"
    cd1.DialogTitle = "Kaydet"
    cd1.ShowSave
    rchHtml.SaveFile cd1.FileName, rtfText
Else
    '
    rchHtml.SaveFile AcilanDosya, rtfText
End If
DosyaDegisirse = False
rchHtml.SetFocus
End Sub

Private Sub mnuKes_Click()
'Gayet basit bir kodla Kes, Kopyala, Yapýþtýr iþlemlerimizi hallediyoruz
'Clipboardu kullanarak bu üç iþlemi gerçekleþtiriyoruz
If rchHtml.SelText = "" Then
        Exit Sub
    Else
        Clipboard.Clear
        Clipboard.SetText rchHtml.SelText
        rchHtml.SelText = ""
    End If
End Sub

Private Sub mnuKopyala_Click()
If rchHtml.SelText = "" Then
        Exit Sub
    Else
        Clipboard.Clear
        Clipboard.SetText rchHtml.SelText
    End If
End Sub

Private Sub mnuLinkEdit_Click()
    frmBaglanti.Show 1
End Sub

Private Sub mnuListeEditoru_Click()
    frmListe.Show 1
End Sub

Private Sub mnuResimEditoru_Click()
    frmResim.Show 1
End Sub

Private Sub mnuSembol_Click()
    frmSembol.Show 1
End Sub

Private Sub mnuSil_Click()
If rchHtml.SelText = "" Then
    Exit Sub
Else
    rchHtml.SelText = ""
End If
End Sub

Private Sub mnuTumunuSec_Click()
 'imlecý pozisyonunu sýfýrla
 rchHtml.SelStart = 0
 'metnin tüm uzunluðunu alýr
 rchHtml.SelLength = Len(rchHtml.Text)
 'daha sonra tekrar rchHtml ye odaklanýr
 rchHtml.SetFocus
End Sub

Private Sub mnuYapistir_Click()
 rchHtml.SelText = Clipboard.GetText
End Sub

Private Sub mnuYazdir_Click()
    MsgBox "Yazdýr [SCT]Hýzlý HTML v1.3 sürümünde kullanýmda deðildir", vbInformation + vbOKOnly, "[SCT]Hýzlý HTML v1.3"
End Sub

Private Sub mnuYeni_Click()
On Local Error Resume Next
Dim cvp
If DosyaDegisirse = True Then
    cvp = MsgBox("Dosya içeriði deðiþti." & vbCrLf & "Kaydetmek istiyor musunuz?", vbInformation + vbYesNoCancel, "[SCT] Hýzlý HTML v1.3")
        If cvp = vbYes Then
            mnuKaydet_Click
            YeniDosyaAc
        ElseIf cvp = vbNo Then
            YeniDosyaAc
        Else
            Exit Sub
        End If
Else
    YeniDosyaAc
End If
End Sub

Private Sub YeniDosyaAc()
    rchHtml.SelStart = 0
    rchHtml.SelLength = Len(rchHtml.Text)
    rchHtml.SelColor = vbBlack
    rchHtml.SelStart = 0
    rchHtml.Text = ""
    AcilanDosya = ""
    Me.Caption = "[SCT] Hýzlý HTML Editörü v1.3 : Baþlýksýz"
    'Yeni Sayfa dendiðinde boþ sayfa gözükmeyeceðine göre
    'yeni sayfa diye bi alt prosedür hazýrladým
    YeniSayfa
    rchHtml.SetFocus
    DosyaDegisirse = False 'yeni dosya açýldýðýnda dosya içeriði deðiþmemiþ olacak
End Sub

Sub YeniSayfa()
'Yeni sayfa açýlmak istendiðinde
'rchHtml'de bu html kodlarý gözükecek
On Error Resume Next
rchHtml.Text = "<html>" & _
vbCrLf & "<head>" & _
vbCrLf & "<title>Yeni Sayfa</title>" & _
vbCrLf & "<!--Öldürmeyen Her Darbe Güce Güç Katar-->" & _
vbCrLf & "<meta name=GENERATOR content=[SCT]Hýzlý HTML Editörü v1.3>" & _
vbCrLf & "<meta http-equiv=Content-Type content=text/html; charset=windows-1254>" & _
vbCrLf & "</head>" & _
vbCrLf & "<body>" & _
vbCrLf & "" & _
vbCrLf & "</body>" & _
vbCrLf & "</html>"
rchHtml.SetFocus
End Sub

Private Sub rchHtml_Change()
'metnin deðiþtiðinden emin olalým
DosyaDegisirse = True
    'temelde geri ve ileri almayý bu günceller
    If Not gdysFarkiOnemseme Then
        gsylIcerik = gsylIcerik + 1
        gyzlYigin(gsylIcerik) = rchHtml.TextRTF
    End If
End Sub

Private Sub rchHtml_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
durum1.Panels(1).Text = "Yardým için F1 'e basýn"
End Sub

Private Sub tbrAna_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Yeni"
            durum1.Panels(1).Text = "Yeni sayfa açýldý"
            mnuYeni_Click
        Case "Ac"
            durum1.Panels(1).Text = "Varolan dosyayý açar"
            mnuAc_Click
        Case "Kaydet"
            durum1.Panels(1).Text = "Dosyayý kaydeder"
            mnuKaydet_Click
        Case "Yazdir"
            durum1.Panels(1).Text = "HTML kodlarýný yazdýrýr (Henüz aktif deðil)"
            mnuYazdir_Click
        Case "GeriAl"
            mnuGeriAl_Click
        Case "IleriAl"
            mnuIleriAl_Click
        Case "Bul"
            durum1.Panels(1).Text = "HTML kodlarý içinde arama yapmanýzý saðlar"
            frmBul.Show 1
        Case "Kes"
            mnuKes_Click
        Case "Kopyala"
            mnuKopyala_Click
        Case "Yapistir"
            mnuYapistir_Click
        Case "Sil"
            mnuSil_Click
        Case "Hakkinda"
            durum1.Panels(1).Text = "[SCT] Hýzlý HTML Hakkýnda"
            mnuHakkinda_Click
    End Select
End Sub

Public Sub ImlecYerlestir(Metin$, Imlec As Long)
Dim T As Long
    T = rchHtml.SelStart
    rchHtml.SelStart = (T + Len(Tag)) - Imlec
End Sub

Private Sub tbrArac_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Sola"
            EtiketEkle "<p align=" & Chr(34) & "left" & Chr(34) & ">&nbsp;</p>", True
            ImlecYerlestir rchHtml.SelText, 4
        Case "Orta"
            EtiketEkle "<p align=" & Chr(34) & "center" & Chr(34) & ">&nbsp;</p>", True
            ImlecYerlestir rchHtml.SelText, 4
        Case "Saga"
            EtiketEkle "<p align=" & Chr(34) & "right" & Chr(34) & ">&nbsp;</p>", True
            ImlecYerlestir rchHtml.SelText, 4
        Case "IkiYana"
            EtiketEkle "<p align=" & Chr(34) & "justify" & Chr(34) & ">&nbsp;</p>", True
            ImlecYerlestir rchHtml.SelText, 4
        Case "Bosluk"
            EtiketEkle "&nbsp;", True
            ImlecYerlestir rchHtml.SelText, 0
        Case "Break"
            EtiketEkle "<br>", True
            ImlecYerlestir rchHtml.SelText, 0
        Case "Paragraf"
            EtiketEkle "<p></p>", True
            ImlecYerlestir rchHtml.SelText, 4
        Case "Div"
            EtiketEkle "<div></div>", True
            ImlecYerlestir rchHtml.SelText, 6
        Case "Cizik"
            EtiketEkle "<hr width=" & Chr(34) & "100%" & Chr(34) & " size=" & Chr(34) & "1" & Chr(34) & ">", True
            ImlecYerlestir rchHtml.SelText, 0
        Case "UstCizgili"
            EtiketEkle "<s></s>", True
            ImlecYerlestir rchHtml.SelText, 4
        Case "Bul"
            frmBul.Show 1
        Case "Yorum"
            EtiketEkle "<!-- &nbsp; -->", True
            ImlecYerlestir rchHtml.SelText, 4
        Case "Resim"
            frmResim.Show 1
        Case "EPosta"
            EtiketEkle "<a href=" & Chr(34) & "mailto:sctposta@hotmail.com" & Chr(34) & ">sctposta@hotmail.com</a>", True
            ImlecYerlestir rchHtml.SelText, 4
        Case "Link"
            frmBaglanti.Show 1
    End Select
End Sub

Private Sub tbrBaslik_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "B1"
            EtiketEkle "<h1></h1>", True
            ImlecYerlestir rchHtml.SelText, 5
        Case "B2"
            EtiketEkle "<h2></h2>", True
            ImlecYerlestir rchHtml.SelText, 5
        Case "B3"
            EtiketEkle "<h3></h3>", True
            ImlecYerlestir rchHtml.SelText, 5
        Case "B4"
            EtiketEkle "<h4></h4>", True
            ImlecYerlestir rchHtml.SelText, 5
        Case "B5"
            EtiketEkle "<h5></h5>", True
            ImlecYerlestir rchHtml.SelText, 5
        Case "B6"
            EtiketEkle "<h6></h6>", True
            ImlecYerlestir rchHtml.SelText, 5
    End Select
End Sub

Private Sub tbrMetin_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Kalin"
            EtiketEkle "<b></b>", True
            ImlecYerlestir rchHtml.SelText, 4
        Case "Egik"
            EtiketEkle "<i></i>", True
            ImlecYerlestir rchHtml.SelText, 4
        Case "AltCizgili"
            EtiketEkle "<u></u>", True
            ImlecYerlestir rchHtml.SelText, 4
        Case "YaziTipi"
            EtiketEkle "<font color=" & Chr(34) & Chr(34) & " size=" & Chr(34) & Chr(34) & ">&nbsp;</font>", True
            ImlecYerlestir rchHtml.SelStart, 7
        Case "YTipEksi1"
            EtiketEkle "<font size=" & Chr(34) & "-1" & Chr(34) & ">&nbsp;</font>", True
            ImlecYerlestir rchHtml.SelStart, 7
        Case "YTipArti1"
            EtiketEkle "<font size=" & Chr(34) & "+1" & Chr(34) & ">&nbsp;</font>", True
            ImlecYerlestir rchHtml.SelStart, 7
        Case "SuperScript"
            EtiketEkle "<sup></sup>", True
            ImlecYerlestir rchHtml.SelText, 6
        Case "SubScript"
            EtiketEkle "<sub></sub>", True
            ImlecYerlestir rchHtml.SelText, 6
        Case "Small"
            EtiketEkle "<small></small>", True
            ImlecYerlestir rchHtml.SelText, 8
        Case "Big"
            EtiketEkle "<big></big>", True
            ImlecYerlestir rchHtml.SelText, 6
        Case "KayanYazi"
            frmKayanYazi.Show 1
        Case "TTFont"
            EtiketEkle "<tt></tt>", True
            ImlecYerlestir rchHtml.SelText, 5
        Case "Vurgu"
            EtiketEkle "<em></em>", True
            ImlecYerlestir rchHtml.SelText, 5
        Case "Strong"
            EtiketEkle "<strong></strong>", True
            ImlecYerlestir rchHtml.SelText, 9
        Case "B1"
            EtiketEkle "<h1></h1>", True
            ImlecYerlestir rchHtml.SelText, 5
        Case "B2"
            EtiketEkle "<h2></h2>", True
            ImlecYerlestir rchHtml.SelText, 5
        Case "B3"
            EtiketEkle "<h3></h3>", True
            ImlecYerlestir rchHtml.SelText, 5
        Case "B4"
            EtiketEkle "<h4></h4>", True
            ImlecYerlestir rchHtml.SelText, 5
        Case "B5"
            EtiketEkle "<h5></h5>", True
            ImlecYerlestir rchHtml.SelText, 5
        Case "B6"
            EtiketEkle "<h6></h6>", True
            ImlecYerlestir rchHtml.SelText, 5
    End Select
End Sub
