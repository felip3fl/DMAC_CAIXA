VERSION 5.00
Begin VB.Form frmPortal 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   1035
   ClientTop       =   1320
   ClientWidth     =   14940
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   14940
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   -15
      ScaleHeight     =   390
      ScaleWidth      =   15465
      TabIndex        =   0
      Top             =   -135
      Width           =   15465
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "(Pressione as teclas alt + espaço + x para maximiza a janela)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   330
         Left            =   930
         TabIndex        =   2
         Top             =   150
         Width           =   6360
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Portal NFe "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   15
         TabIndex        =   1
         Top             =   150
         Width           =   1005
      End
   End
   Begin VB.PictureBox frameLimitadorNavegador 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8820
      Left            =   390
      ScaleHeight     =   8790
      ScaleWidth      =   16050
      TabIndex        =   3
      Top             =   45
      Width           =   16080
      Begin Balcao2010.IEalt webNavegador 
         Left            =   1110
         Top             =   1275
         _ExtentX        =   847
         _ExtentY        =   847
      End
   End
End
Attribute VB_Name = "frmPortal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    'Dim margem As Integer
    
   ' margem = 500
    'frameLimitadorNavegador.top = (margem - (margem * 2))
    frameLimitadorNavegador.top = -100
    frameLimitadorNavegador.left = 0
    frameLimitadorNavegador.Height = (Me.Height) + 100
    frameLimitadorNavegador.Width = Me.Width
    
End Sub

Private Sub Form_Load()
    'AjustaTela Me
    'frmControlaCaixa.Enabled = True
    
   'webNavegador.bControlInDevelopmentMode = True
    
   'webNavegador.Nav GLB_EnderecoPortal
   'webNavegador.EmbedIE frameLimitadorNavegador.Hwnd
   
End Sub


Private Sub Form_LostFocus()
    Unload Me
End Sub

