VERSION 5.00
Begin VB.Form frmEntradaDeValores 
   BackColor       =   &H00232323&
   BorderStyle     =   0  'None
   Caption         =   "Entrada de Valores"
   ClientHeight    =   3600
   ClientLeft      =   7515
   ClientTop       =   3285
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtEntrada 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   200
      TabIndex        =   0
      Top             =   2310
      Width           =   4905
   End
   Begin Balcao2010.chameleonButton cmdCancelar 
      Height          =   400
      Left            =   3495
      TabIndex        =   3
      Top             =   3015
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Cancelar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   2500134
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "frmEntradaDeValores.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton cmdOK 
      Height          =   400
      Left            =   195
      TabIndex        =   4
      Top             =   3015
      Width           =   1600
      _ExtentX        =   2831
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Entrar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   2500134
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "frmEntradaDeValores.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton cmdRetornoOperacaoTEF 
      Height          =   400
      Left            =   1845
      TabIndex        =   5
      Top             =   3015
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   714
      BTYPE           =   14
      TX              =   "Retornar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   2500134
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "frmEntradaDeValores.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblMensagem 
      BackColor       =   &H00AE7411&
      BackStyle       =   0  'Transparent
      Caption         =   "Mensagem"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1710
      Left            =   195
      TabIndex        =   2
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00AE7411&
      BackStyle       =   0  'Transparent
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2520
      TabIndex        =   1
      Top             =   150
      Width           =   555
   End
End
Attribute VB_Name = "frmEntradaDeValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public tamanhoMaximo As Integer
Public tamanhoMinino As Integer

Private Sub cmdAlterarModalidade_Click(Index As Integer)

End Sub


Private Sub cmdCancelar_Click()
    retornoEntradaDeValores = "-1"
    Unload Me
End Sub

Private Sub cmdOK_Click()
    validaValores
End Sub

Private Sub cmdRetornoOperacaoTEF_Click()
    retornaOperacaoTEF = True
    retornoEntradaDeValores = (txtEntrada.text)
    Me.Visible = False
    Unload Me
End Sub

Private Sub Form_Load()
    
    Screen.MousePointer = 0
    
    Me.top = 4700
    Me.left = 5050
    Me.Width = 5320
    Me.Height = 3700
    
    txtEntrada.text = ""
    lblTitulo.Caption = ""
    lblMensagem.Caption = ""
    
    retornoEntradaDeValores = "-1"
    
End Sub


Private Sub validaValores()

    If Len(txtEntrada.text) > tamanhoMaximo Or Len(txtEntrada.text) < tamanhoMinino Then
        MsgBox "Atenção! a quantidade de caracteres deve ser entre " & tamanhoMinino & " e " & tamanhoMaximo & "", vbExclamation, "Entrada de Valores"
        txtEntrada.SetFocus
    Else
        retornoEntradaDeValores = (txtEntrada.text)
        Me.Visible = False
        Unload Me
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.MousePointer = 0
End Sub

Private Sub txtEntrada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        validaValores
    ElseIf KeyAscii = 27 Then
        cmdCancelar_Click
    End If
End Sub
