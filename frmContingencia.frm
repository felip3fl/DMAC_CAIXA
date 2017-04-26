VERSION 5.00
Begin VB.Form frmContingencia 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Modo de Contingencia"
   ClientHeight    =   6105
   ClientLeft      =   2820
   ClientTop       =   3255
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmeSenhaDesbloqueioCont 
      Interval        =   1000
      Left            =   5985
      Top             =   4260
   End
   Begin VB.Frame fraSenha 
      BackColor       =   &H00000000&
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   1815
      TabIndex        =   0
      Top             =   2820
      Width           =   1920
      Begin VB.TextBox txtSenha 
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
         Left            =   120
         MaxLength       =   4
         TabIndex        =   1
         Top             =   270
         Width           =   1680
      End
   End
   Begin Balcao2010.chameleonButton cmdDesabilitarContingencia 
      Height          =   930
      Left            =   465
      TabIndex        =   3
      Top             =   4650
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   1640
      BTYPE           =   14
      TX              =   "DESABILITAR modo de Contingencia"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmContingencia.frx":0000
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
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmContingencia.frx":001C
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2505
      Left            =   -435
      TabIndex        =   2
      Top             =   495
      Width           =   9195
   End
End
Attribute VB_Name = "frmContingencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mensagemAtencao(contigenciaAtiva As Boolean)

    lblMensagem.Caption = "Atenção!" & vbNewLine & _
    "Antes de habilitar o modo contingência, certifique que o FORMULARIO DE SEGURANÇA" & vbNewLine & _
    "esteja na impressora. Alerte os colaboradores para não enviar impressão(cotação, documentos) " & vbNewLine & _
    "enquanto este modo estiver habilitado. Habilite a contingência apenas no computador que emite nota (no Caixa)" & vbNewLine & vbNewLine & vbNewLine & _
    "STATUS CONTIGENCIA: "
    
    If contigenciaAtiva Then
        lblMensagem.Caption = lblMensagem.Caption & "HABILITADO"
    Else
        lblMensagem.Caption = lblMensagem.Caption & "DESABILITADO"
    End If
    
End Sub

Private Sub Form_Load()
    
    Call AjustaTela(Me)
    verificaModoEmissaoAtual
    
    lblMensagem.left = (Me.Width / 2) - (lblMensagem.Width / 2)
    fraSenha.left = (Me.Width / 2) - (fraSenha.Width / 2)
    cmdDesabilitarContingencia.left = (Me.Width / 2) - (cmdDesabilitarContingencia.Width / 2)
    
End Sub

Public Function verificaModoEmissaoAtual() As Byte

    Dim RsControleCont As New ADODB.Recordset
    
    sql = "select CTS_Contingencia as Contingencia from controleSistema"
    RsControleCont.CursorLocation = adUseClient
    RsControleCont.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If (RsControleCont("Contingencia") <> 0) Then
        habilitaContingencia RsControleCont("Contingencia")
    Else
        desabilitaContingencia
    End If
    
    verificaModoEmissaoAtual = RsControleCont("Contingencia")
    
    RsControleCont.Close

End Function

Private Sub criaTXTContingencia(habilita As Byte, justificativa As String)
    Dim corpoMensagem As String
    
On Error GoTo TrataErro
    
    corpoMensagem = "modalidade=" & habilita
    If justificativa <> "" Then corpoMensagem = corpoMensagem & vbNewLine & justificativa
    
    criaTXTtemporario = GLB_EnderecoPastaFIL & "contingencia" & "#" & wCGC & ".txt"
    Open criaTXTtemporario For Output As #1
         Print #1, corpoMensagem
    Close #1
    
    Exit Sub
    
TrataErro:
    Select Case Err.Number
    Case Else
        'mensagemErroDesconhecido Err, "Erro na criação do arquivo"
    End Select
End Sub

Private Sub habilitaContingencia(cod As Byte)
    '0 - normal
    '2 - contingencia
    
    criaTXTContingencia cod, "justificativa=Problemas Tecnicos por conexao com a internet"
    cmdDesabilitarContingencia.Visible = True
    fraSenha.Visible = False
    txtSenha.text = ""
    
    rdoCNLoja.Execute "update ControleSistema set CTS_Contingencia = '" & cod & "'"
    
    mensagemAtencao True
    
End Sub

Private Sub desabilitaContingencia()

    criaTXTContingencia "0", ""
    
    cmdDesabilitarContingencia.Visible = False
    fraSenha.Visible = True
    txtSenha.text = ""
    
    rdoCNLoja.Execute "update controlesistema set CTS_Contingencia = '0'"
    
    mensagemAtencao False
    
End Sub



