VERSION 5.00
Begin VB.Form frmSpla 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrPause 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2625
      Top             =   2835
   End
   Begin VB.PictureBox picMainSkin 
      Height          =   2940
      Left            =   0
      MousePointer    =   99  'Custom
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   2880
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   0
      Width           =   4635
   End
End
Attribute VB_Name = "frmSpla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************
'internet dialer program developed by Harish
'for use to connect to the internet
'project : Dial Net
' the code for translucence in this for was
'developed by someone else.
'******************************************
Option Explicit
Private Sub Form_Load()
    Dim WindowRegion As Long
    ' I set all these settings here so you won't forget
    ' them and have a non-working demo... Set them in
    ' design time
    picMainSkin.ScaleMode = vbPixels
    picMainSkin.AutoRedraw = True
    picMainSkin.AutoSize = True
    picMainSkin.BorderStyle = vbBSNone
    Me.BorderStyle = vbBSNone
        
    'Set picMainSkin.Picture = LoadPicture(App.Path & "\bigsqueel.bmp")
    
    Me.Width = picMainSkin.Width
    Me.Height = picMainSkin.Height
    
    WindowRegion = MakeRegion(picMainSkin)
    SetWindowRgn Me.hWnd, WindowRegion, True
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2
    DoEvents
    'Me.Show
    tmrPause.Enabled = True
    End Sub
Private Sub tmrPause_Timer()
Load frmdial
frmdial.Show
DoEvents
Unload Me
End Sub


