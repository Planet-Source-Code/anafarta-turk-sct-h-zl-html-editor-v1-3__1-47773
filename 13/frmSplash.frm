VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5010
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":1042
   ScaleHeight     =   2490
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer 
      Interval        =   3000
      Left            =   -240
      Top             =   3120
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2460
      Left            =   20
      Top             =   30
      Width           =   4990
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer_Timer()
    Timer.Enabled = False
    Unload Me
    frmAna.Show
End Sub

