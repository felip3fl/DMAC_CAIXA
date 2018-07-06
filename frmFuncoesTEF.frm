VERSION 5.00
Begin VB.Form frmFuncoesTEF 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10455
   ClientLeft      =   495
   ClientTop       =   1890
   ClientWidth     =   17955
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10455
   ScaleWidth      =   17955
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmAdministrador 
      BackColor       =   &H00000000&
      Height          =   3960
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   4035
      Begin Balcao2010.chameleonButton cmdLimparComprovantes 
         Height          =   555
         Left            =   285
         TabIndex        =   1
         Top             =   645
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   979
         BTYPE           =   14
         TX              =   "Limpar comprovantes"
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
         MICON           =   "frmFuncoesTEF.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Balcao2010.chameleonButton cmdConsultarSaldoCartao 
         Height          =   555
         Left            =   285
         TabIndex        =   2
         Top             =   1200
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   979
         BTYPE           =   14
         TX              =   "Consultar Saldo"
         ENAB            =   0   'False
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
         MICON           =   "frmFuncoesTEF.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Balcao2010.chameleonButton cmdRetornar 
         Height          =   555
         Left            =   285
         TabIndex        =   4
         Top             =   3120
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   979
         BTYPE           =   14
         TX              =   "Retornar"
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
         MICON           =   "frmFuncoesTEF.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Balcao2010.chameleonButton cmdFuncao110 
         Height          =   555
         Left            =   285
         TabIndex        =   5
         Top             =   1755
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   979
         BTYPE           =   14
         TX              =   "Funções Especiais"
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
         MICON           =   "frmFuncoesTEF.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Funções TEF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   315
         TabIndex        =   3
         Top             =   255
         Width           =   1380
      End
   End
   Begin VB.Label lblMensagensTEF 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "teste teste teste teste teste teste teste teste teste teste teste teste teste teste teste "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   870
      Left            =   75
      TabIndex        =   6
      Top             =   6900
      Width           =   15180
   End
End
Attribute VB_Name = "frmFuncoesTEF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFuncao110_Click()
    funcao110
End Sub

Private Sub cmdLimparComprovantes_Click()

    If limparArquivosImpressaoTEF() Then
        MsgBox "Limpeza realizada com sucesso" & vbNewLine & _
        GLB_ENDERECOCOMPROVANTETEF, vbInformation
    End If

End Sub

Private Sub cmdRetornar_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    Call AjustaTela(Me)
    
    lblMensagensTEF.Caption = ""
    
End Sub

Private Sub funcao110()
    'If GLB_Administrador Then
        
        Dim nf As notaFiscalTEF
        PegaNumeroPedido
        
        nf.pedido = pedido
        
        Call EfetuaOperacaoTEF("110", nf, lblMensagensTEF, lblMensagensTEF)
        ImprimeComprovanteTEF nf.pedido
        finalizarTransacaoTEF nf.pedido, nf.serie, False
        
    'Else
    '    MsgBox "Você não tem permissão para executar essa função", vbInformation
    'End If
End Sub

Private Sub lblMensagensTEF_Click()
    cancelarOperacaoTEF = True
End Sub
