VERSION 5.00
Begin VB.Form frmFuncoesTEF 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   495
   ClientTop       =   1890
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmAdministrador 
      BackColor       =   &H00000000&
      Height          =   3045
      Left            =   195
      TabIndex        =   0
      Top             =   90
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
         Top             =   2100
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
End
Attribute VB_Name = "frmFuncoesTEF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    
End Sub
