VERSION 5.00
Object = "{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7u.ocx"
Begin VB.Form frmEmissaoNFe 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Emiss�o NFe"
   ClientHeight    =   9240
   ClientLeft      =   555
   ClientTop       =   1140
   ClientWidth     =   19035
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   19035
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame frmAdministrador 
      BackColor       =   &H00000000&
      Height          =   1995
      Left            =   420
      TabIndex        =   24
      Top             =   5490
      Visible         =   0   'False
      Width           =   3825
      Begin Balcao2010.chameleonButton cmdCancelar 
         Height          =   555
         Left            =   165
         TabIndex        =   26
         Top             =   645
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   979
         BTYPE           =   14
         TX              =   "Cancelar NFe / CFe"
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
         MICON           =   "frmEmissaoNFe.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Balcao2010.chameleonButton cmdLiberar 
         Height          =   555
         Left            =   180
         TabIndex        =   27
         Top             =   1185
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   979
         BTYPE           =   14
         TX              =   "Atualizar TM "
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
         MICON           =   "frmEmissaoNFe.frx":001C
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
         Caption         =   "Administrador"
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
         Left            =   195
         TabIndex        =   25
         Top             =   255
         Width           =   1455
      End
   End
   Begin VB.Timer timerExibirMSG 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   705
      Top             =   7245
   End
   Begin VB.Frame frmNCM 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   6015
      TabIndex        =   20
      Top             =   3750
      Visible         =   0   'False
      Width           =   3450
      Begin VB.TextBox txtNCM 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   200
         MaxLength       =   8
         TabIndex        =   21
         Top             =   600
         Width           =   3000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Digite o NCM:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   200
         TabIndex        =   22
         Top             =   200
         Width           =   1230
      End
   End
   Begin VB.Frame frameNFE 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      ForeColor       =   &H80000008&
      Height          =   7950
      Left            =   4385
      TabIndex        =   11
      Top             =   375
      Width           =   10725
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Height          =   675
         Left            =   315
         TabIndex        =   12
         Top             =   0
         Width           =   9990
         Begin VB.CheckBox chkMostraLogErro 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "&Erro"
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   3360
            TabIndex        =   14
            Top             =   255
            Value           =   1  'Checked
            Width           =   600
         End
         Begin VB.CheckBox chkMostraLogSucesso 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "&Sucesso"
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   4395
            TabIndex        =   13
            Top             =   255
            Value           =   1  'Checked
            Width           =   915
         End
         Begin Balcao2010.chameleonButton cmdAtualizar 
            Height          =   315
            Left            =   7740
            TabIndex        =   15
            Top             =   210
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   556
            BTYPE           =   14
            TX              =   "Atualizar"
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
            MICON           =   "frmEmissaoNFe.frx":0038
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filtro de mensagem de log:"
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
            Left            =   195
            TabIndex        =   16
            Top             =   255
            Width           =   2850
         End
      End
      Begin VSFlex7UCtl.VSFlexGrid grdLogSig 
         Height          =   2835
         Left            =   315
         TabIndex        =   17
         Top             =   1110
         Width           =   9990
         _cx             =   17621
         _cy             =   5001
         _ConvInfo       =   1
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   14737632
         ForeColor       =   4210752
         BackColorFixed  =   0
         ForeColorFixed  =   16777215
         BackColorSel    =   3421236
         ForeColorSel    =   16777215
         BackColorBkg    =   0
         BackColorAlternate=   12632256
         GridColor       =   14737632
         GridColorFixed  =   8421504
         TreeColor       =   8421504
         FloodColor      =   16777215
         SheetBorder     =   8421504
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   5
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEmissaoNFe.frx":0054
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   5
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   -2147483633
         ForeColorFrozen =   4210752
         WallPaperAlignment=   4
         Begin VB.Timer timerVerificaResposta 
            Enabled         =   0   'False
            Interval        =   3000
            Left            =   0
            Top             =   0
         End
      End
      Begin VSFlex7UCtl.VSFlexGrid grdLogSigSAT 
         Height          =   2805
         Left            =   315
         TabIndex        =   19
         Top             =   4320
         Width           =   9990
         _cx             =   17621
         _cy             =   4948
         _ConvInfo       =   1
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   14737632
         ForeColor       =   4210752
         BackColorFixed  =   0
         ForeColorFixed  =   16777215
         BackColorSel    =   3421236
         ForeColorSel    =   16777215
         BackColorBkg    =   0
         BackColorAlternate=   12632256
         GridColor       =   14737632
         GridColorFixed  =   8421504
         TreeColor       =   8421504
         FloodColor      =   16777215
         SheetBorder     =   8421504
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEmissaoNFe.frx":0172
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   5
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   -2147483633
         ForeColorFrozen =   4210752
         WallPaperAlignment=   4
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Left            =   450
            Top             =   3630
         End
      End
   End
   Begin VB.Timer timerSairSistema 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1845
      Top             =   7965
   End
   Begin VB.Frame frameDadosNotaFiscal 
      BackColor       =   &H00000000&
      Height          =   4575
      Left            =   400
      TabIndex        =   2
      Top             =   400
      Width           =   3825
      Begin VB.OptionButton optPesquisaNumero 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "N�mero da Nota"
         ForeColor       =   &H8000000F&
         Height          =   585
         Left            =   1935
         TabIndex        =   5
         Top             =   870
         Width           =   1740
      End
      Begin VB.OptionButton optPesquisaPedido 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Pedido"
         ForeColor       =   &H8000000F&
         Height          =   585
         Left            =   195
         TabIndex        =   4
         Top             =   870
         Value           =   -1  'True
         Width           =   1440
      End
      Begin VB.TextBox txtNFe 
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
         TabIndex        =   3
         Top             =   1485
         Width           =   3435
      End
      Begin Balcao2010.chameleonButton cmdTransmitir 
         Height          =   720
         Left            =   200
         TabIndex        =   9
         Top             =   2250
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   1270
         BTYPE           =   14
         TX              =   "Transmitir"
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
         MICON           =   "frmEmissaoNFe.frx":0295
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Balcao2010.chameleonButton cmdImprimir 
         Height          =   720
         Left            =   200
         TabIndex        =   10
         Top             =   2970
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   1270
         BTYPE           =   14
         TX              =   "Imprimir"
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
         MICON           =   "frmEmissaoNFe.frx":02B1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Balcao2010.chameleonButton cmdEmail 
         Height          =   720
         Left            =   195
         TabIndex        =   23
         Top             =   3690
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   1270
         BTYPE           =   14
         TX              =   "Enviar Email"
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
         MICON           =   "frmEmissaoNFe.frx":02CD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de pesquisa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   195
         TabIndex        =   7
         Top             =   630
         Width           =   1635
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nota Fiscal Eletr�nica"
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
         Left            =   195
         TabIndex        =   6
         Top             =   255
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   375
      TabIndex        =   1
      Top             =   780
      Width           =   3735
   End
   Begin VB.Label cmdIgnorarResultado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click aqui para ignorar o resultado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   210
      TabIndex        =   18
      Top             =   9165
      Visible         =   0   'False
      Width           =   15210
   End
   Begin VB.Label lblMSGNota 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota Fiscal n�o encontrada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   405
      TabIndex        =   8
      Top             =   5085
      Width           =   3765
   End
   Begin VB.Label lblStatusImpressao 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Preparando para emitir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   210
      TabIndex        =   0
      Top             =   8520
      Visible         =   0   'False
      Width           =   15210
   End
End
Attribute VB_Name = "frmEmissaoNFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vetCampos() As String
Dim sql As String
Dim tiponota As String
Private whereNotaFiscal As String
Const insertTabelaNFLojas = "insert into NFE_NFLojas " & vbNewLine & _
                            "(nfl_sequencia, nfl_descricao, nfl_dados, nfl_loja, nfl_nroNFE, nfl_dataEmissao) " & vbNewLine & _
                            "values ('"
                            
Dim nf As notaFiscal
Dim Tempo As Byte
Dim mensagemStatus As String
Dim qtdeLinhaAnterior As Integer
Dim abrirAqruivo As Boolean
Public lojasWhere As String
Dim tempoVerificacaoResposta As Long
Dim endArquivoResposta As String

Dim icms41 As Boolean
Dim icms50 As Boolean


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub chameleonButton1_Click()

End Sub

Private Sub chkMostraLogErro_Click()
    carregaGrdLogSig
End Sub

Private Sub chkMostraLogSucesso_Click()
    carregaGrdLogSig
End Sub

Private Sub cmdAtualizar_Click()
    qtdeLinhaAnterior = -1
    carregaArquivo
End Sub



Private Sub cmdCancelar_Click()
    Dim Arquivo As String

    cancelaNota = False
    
    If nf.eSerie = "NE" Then
        'Arquivo = Dir(GLB_EnderecoPastaRESP & "*" & nf.numero & "#" & nf.CNPJ & ".txt", vbDirectory)
        'If Arquivo <> "" Then
            deletaArquivo GLB_EnderecoPastaRESP & "*" & nf.numero & "*"
        'End If
        
        finalizaProcesso "Cancelando Nota Fiscal Eletr�nico " & nf.numero, True
        cancelaNE nf
        
    ElseIf nf.eSerie Like "CE*" Then

        'Arquivo = Dir(GLB_EnderecoPastaRESP & "*" & nf.pedido & "#" & nf.CNPJ & ".txt", vbDirectory)
        'If Arquivo <> "" Then
            deletaArquivo GLB_EnderecoPastaRESP & "*" & nf.numero & "*"
        'End If
        
        finalizaProcesso "Cancelando Cupom Fiscal Eletr�nico " & nf.numero, True
        cancelaSAT nf
    Else
        MsgBox "Nota n�o valida para esse tipo de cancelamento", vbCritical, "Cancelamento de NE ou CE"
    End If
    

End Sub


Private Sub cmdContingencia_Click()

End Sub

Private Sub cmdEmail_Click()

    Dim rsNFE As New ADODB.Recordset
    Dim Arquivo As String
    
    sql = "select top 1 nf as nf, " & vbNewLine _
        & "ChaveNFe as chave," & vbNewLine _
        & "ce_email as email," & vbNewLine _
        & "ce_razao as nome" & vbNewLine _
        & "from nfcapa, fin_cliente" & vbNewLine _
        & "where " & vbNewLine _
        & "lojaorigem in " & lojasWhere & " " & "" & vbNewLine _
        & "and ce_codigoCliente = cliente" & vbNewLine _
        & "and tiponota in ('V','S','E','R')" & vbNewLine _
        & "and numeroped = " & nf.pedido
    
    rsNFE.CursorLocation = adUseClient
    rsNFE.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
        If Not rsNFE.EOF Then
            If rsNFE("chave") = Empty Then
                lblMSGNota.Caption = "Chave de acesso n�o encontrada"
            ElseIf rsNFE("email") = Empty Then
                lblMSGNota.Caption = "Email n�o cadastrado ou v�lido"
                cmdEmail.Enabled = False
            Else
                lblMSGNota.Caption = ""
                criarArquivorEmail nf, rsNFE("chave"), rsNFE("email"), rsNFE("nome")
                Tempo = 56
                finalizaProcesso "Enviando XML e DANFE para " & rsNFE("email"), False
            End If
        Else
            lblMSGNota.Caption = "Nota n�o encontrada ou opera��o de nf inv�lido"
        End If
    
    rsNFE.Close
    
End Sub

Private Sub cmdIgnorarResultado_Click()
    Tempo = 200
End Sub

Private Sub cmdLiberar_Click()
    Screen.MousePointer = 11
    sql = "update nfcapa set tm = 100"
    rdoCNLoja.Execute sql
    Screen.MousePointer = 0
    MsgBox "Atualiza��o do TM realizada com sucesso", vbInformation, "DMAC Caixa"
End Sub

'Public Sub deletaResposta(pedido As String, nf As notaFiscal)

 '   arquivo = Dir(GLB_EnderecoPastaRESP & "*" & pedido & "#" & nf.cnpj & ".txt", vbDirectory)
    
'End Sub

Public Sub cmdTransmitir_Click()
    
    Dim Arquivo As String
    
    emitiNota = False
    nf.loja = wLoja
    If optPesquisaPedido.Value = True Then nf.pedido = txtNFe.text
     
    If nf.eSerie = "NE" Then
        Arquivo = Dir(GLB_EnderecoPastaRESP & "*" & nf.numero & "#" & nf.CNPJ & ".txt", vbDirectory)
    ElseIf nf.eSerie Like "CE*" Then
        Arquivo = Dir(GLB_EnderecoPastaRESP & "*" & nf.pedido & "#" & nf.CNPJ & ".txt", vbDirectory)
    End If
    
    If Arquivo <> "" Then
        deletaArquivo GLB_EnderecoPastaRESP & Arquivo
    End If
     
    If nf.eSerie = "NE" Then
    
            
        finalizaProcesso "Emitindo Nota Fiscal Eletr�nica " & nf.numero, True
        exibirMensagemTEF "  Emitindo NFe" & vbNewLine & "   Aguarde..."
        
        
        criaDuplicataBanco
    
        sql = "exec sp_vda_cria_nfe '" & nf.loja & "', '" & nf.numero & "', 'NE', ''"
        rdoCNLoja.Execute sql
    
        Dim i As Byte
    
        For i = 0 To UBound(vetCampos)
            If vetCampos(i) <> "" Then
                leituraEstrutura vetCampos(i)
            End If
        Next i
    
        numeroCopiaImpressao
    
        criaTXT "nota", nf
        atualizaNota "IDE"
        
    ElseIf nf.eSerie Like "CE*" Then

        finalizaProcesso "Emitindo Cupom Fiscal Eletr�nico", True
        exibirMensagemTEF " Emitindo Cupom" & vbNewLine & "   Aguarde..."
        
        criaTXTSAT "sat", nf
             
        
    End If

End Sub

Private Sub cmdImprimir_Click()
    Tempo = 56
    If nf.eSerie = "NE" Then
        finalizaProcesso "Imprimindo Nota Fiscal Eletr�nico " & nf.numero, False
        Call ImprimirNota(nf, "NOTA")
    ElseIf nf.eSerie Like "CE*" Then
        finalizaProcesso "Imprimindo Cupom Fiscal Eletr�nico " & nf.numero, False
        Call ImprimirNota(nf, "SAT")
    End If
End Sub


Private Sub finalizaProcesso(mensagem As String, esperaResposta As Boolean)
            
    mensagemStatus = mensagem
    frameDadosNotaFiscal.Visible = False
    frmAdministrador.Visible = False
    frameNFE.Visible = False
    lblStatusImpressao.Width = Me.Width
    cmdIgnorarResultado.Width = Me.Width
    
    timerSairSistema.Enabled = esperaResposta
    timerExibirMSG.Enabled = Not (esperaResposta)
    
    lblStatusImpressao.Visible = True
    cmdIgnorarResultado.Visible = True
    
End Sub

Private Sub notaPedentes()

    Dim ado_estrutura As New ADODB.Recordset
    Dim i As Integer
    Dim add As Boolean
    'Dim dataPesquisa As String
    Dim tiponota As String
    
    'dataPesquisa = Format(DateAdd("m", -1, Date), "YYYY/MM/DD")

    sql = "select HORA, DATAEMI, lojaorigem, NUMEROPED, nf, tm, serie, tiponota " & vbNewLine & _
          "from nfcapa " & vbNewLine & _
          "where tm not in (4012,4016,9016,100,101,9005,4005,9012,204,124,4014,4017)   " & vbNewLine & _
          "and tiponota in ('V','T','E','S','R') " & vbNewLine & _
          "and (serie in ('NE') " & vbNewLine & _
          "or serie like 'CE%') " & vbNewLine & _
          "and dataemi >= '" & Format(GLB_DataInicial, "YYYY/MM") & "/01'"

    ado_estrutura.CursorLocation = adUseClient
    ado_estrutura.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
        Do While Not ado_estrutura.EOF
            
            If ado_estrutura("serie") = "NE" Then
            
                If ado_estrutura("tiponota") = "V" Then
                    tiponota = "Venda"
                ElseIf ado_estrutura("tiponota") = "T" Then
                    tiponota = "Transfer�ncia"
                ElseIf ado_estrutura("tiponota") = "E" Then
                    tiponota = "Devolu��o"
                Else
                    tiponota = "NF de Outras Opera��es"
                End If
            
                mensagemLOG2 grdLogSig, Format(ado_estrutura("DATAEMI"), "YYYY/MM/DD") & " " & Format(ado_estrutura("HORA"), "HH:MM"), _
                ado_estrutura("tm"), ado_estrutura("lojaorigem"), ado_estrutura("NF"), ado_estrutura("NUMEROPED"), ado_estrutura("tm") & " - [DMAC] " & tiponota & " n�o sicronizada com a retaguarda"
                
            Else

                'If add Then
                Call mensagemLOG2(grdLogSigSAT, Format(ado_estrutura("DATAEMI"), "YYYY/MM/DD") & " " & Format(ado_estrutura("HORA"), "HH:MM"), _
                Val(ado_estrutura("tm")), ado_estrutura("lojaorigem"), ado_estrutura("NF"), ado_estrutura("NUMEROPED"), ado_estrutura("tm") & " - [DMAC] N�o sicronizada com a retaguarda")
                'End If
                
            End If
        
            ado_estrutura.MoveNext
            
        Loop
    
    ado_estrutura.Close
    
''    'grdLogSig.MergeRow(0) = True
''    'grdLogSig.MergeRow(1) = True
''    'grdLogSig.MergeRow(2) = True
''    'grdLogSig.MergeRow(3) = True
''    grdLogSig.MergeCol(0) = False
''    grdLogSig.MergeCol(1) = True
''    grdLogSig.MergeCol(2) = True
''    grdLogSig.MergeCol(3) = True
''    grdLogSig.MergeCol(4) = False
''    grdLogSig.MergeCol(5) = False
''
''
''    'grdLogSigSAT.MergeRow(0) = True
''    'grdLogSigSAT.MergeRow(1) = True
''    'grdLogSigSAT.MergeRow(2) = True
''    'grdLogSigSAT.MergeRow(3) = True
''    grdLogSigSAT.MergeCol(0) = False
''    grdLogSigSAT.MergeCol(1) = True
''    grdLogSigSAT.MergeCol(2) = True
''    grdLogSigSAT.MergeCol(3) = True
''    grdLogSigSAT.MergeCol(4) = False
''    grdLogSigSAT.MergeCol(5) = False

End Sub

Private Sub Form_Activate()
    
   frmAdministrador.Visible = GLB_Administrador
   
   qtdeLinhaAnterior = 0
  
   grdLogSig.MergeRow(0) = True
   grdLogSig.MergeCol(0) = True
   grdLogSig.MergeCol(1) = True
   grdLogSig.MergeCol(2) = True
   grdLogSig.MergeCol(3) = True
   grdLogSig.MergeCol(4) = True
   'grdLogSig.MergeCol(5) = True
   
   grdLogSigSAT.MergeRow(0) = True
   grdLogSigSAT.MergeCol(0) = True
   grdLogSigSAT.MergeCol(1) = True
   grdLogSigSAT.MergeCol(2) = True
   grdLogSigSAT.MergeCol(3) = True
   grdLogSigSAT.MergeCol(4) = True
'   grdLogSigSAT.MergeCol(5) = True
     
   endArquivoResposta = GLB_EnderecoPastaRESP & "resp-*" & wCGC & ".txt"
  
   cmdIgnorarResultado.top = (frmEmissaoNFe.Height - cmdIgnorarResultado.Height) - 200
   lblStatusImpressao.top = (cmdIgnorarResultado.top - cmdIgnorarResultado.Height) - 200
  
    lblMSGNota.Caption = ""
    lblStatusImpressao.Visible = False
    cmdIgnorarResultado.Visible = False
    cmdTransmitir.Enabled = False
    cmdImprimir.Enabled = False
    cmdEmail.Enabled = False
    timerSairSistema.Enabled = False
    cmdImprimir.ToolTipText = "Nota Fiscal Eletr�nica n�o transmitida"
    Tempo = 0
    cancelaNotaResultado = False
    
    Me.Picture = LoadPicture(endIMG("FundoProcessa"))
    Me.Visible = True
    
    If emitiNota = True Then
        cmdTransmitir_Click
    ElseIf cancelaNota = True Then
        cmdCancelar_Click
    Else
        
        timerVerificaResposta_Timer
        timerVerificaResposta.Enabled = True
        
    End If
    
    
    
End Sub

Public Sub statusFuncionamento(mensagem As String)
    
   ' mensagem = "Imprimindo Garantia Estendida" & " "
    If lblStatusImpressao.Caption = mensagem & " " & "  . . . ." Then
        lblStatusImpressao.Caption = mensagem & " " & ".   . . ."
    ElseIf lblStatusImpressao.Caption = mensagem & " " & ".   . . ." Then
        lblStatusImpressao.Caption = mensagem & " " & ". .   . ."
    ElseIf lblStatusImpressao.Caption = mensagem & " " & ". .   . ." Then
        lblStatusImpressao.Caption = mensagem & " " & ". . .   ."
    ElseIf lblStatusImpressao.Caption = mensagem & " " & ". . .   ." Then
        lblStatusImpressao.Caption = mensagem & " " & ". . . .  "
    ElseIf lblStatusImpressao.Caption = mensagem & " " & ". . . .  " Then
        lblStatusImpressao.Caption = mensagem & " " & "  . . . ."
    Else
        lblStatusImpressao.Caption = mensagem & " " & "  . . . ."
    End If
    
End Sub


Private Sub Form_Load()
    Call AjustaTela(Me)
    carregaInfoLoja
    montaCamposRotulo
    limpaTela
    montaComboLoja
    'carregaArquivo
    
    If emitiNota = True Or cancelaNota = True Then
        txtNFe.text = wPedido
        txtNFe_KeyPress 13
        optPesquisaPedido.Value = True
    End If
    
End Sub

Private Sub limpaTela()
    grdLogSig.Rows = 2
    grdLogSigSAT.Rows = 2
        frameNFE.Visible = Not (emitiNota)
        frameDadosNotaFiscal.Visible = Not (emitiNota)
        frmAdministrador.Visible = Not (emitiNota)
End Sub

Private Sub carregaInfoLoja()

    Dim sql As String
    Dim rsNotaFiscal As New ADODB.Recordset
    
    sql = "select lo_cgc from loja where lo_loja = '" & GLB_Loja & "'"
    rsNotaFiscal.CursorLocation = adUseClient
    rsNotaFiscal.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
        nf.CNPJ = rsNotaFiscal("lo_cgc")
        nf.loja = wLoja
        
    rsNotaFiscal.Close
    
End Sub

Private Function obterNumeroNota(pedido As String, numeroNFE As String) As String

    Dim rsNotaFiscal As New ADODB.Recordset
    Dim campo As String
    Dim valor As String

    If pedido = Empty Then
        campo = "nf"
        valor = numeroNFE
    ElseIf numeroNFE = Empty Then
        campo = "numeroPED"
        valor = pedido
    Else
        Exit Function
    End If

    sql = "select nf " & vbNewLine & _
          "from nfcapa " & vbNewLine & _
          "where " & campo & " = '" & valor & " and serie = 'NE' and lojaorigem = " & wLoja

    'rsNotaFiscal.CursorLocation = adUseClient
    'rsNotaFiscal.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

    'If Not rsNotaFiscal.EOF Then
        'obterNumeroNota = rsno
    'Else
        'obterNumeroNota
    'End If

    rsNotaFiscal.Close

End Function

Private Sub Form_Unload(Cancel As Integer)
    exibirMensagemPadraoTEF
End Sub

Private Sub grdLogSig_Click()
    If grdLogSig.CellForeColor = vbRed Then
        grdLogSig.BackColorSel = vbRed
    Else
        grdLogSig.BackColorSel = &H343434
    End If
    grdLogSigSAT.Row = 0
    abrirAqruivo = False
End Sub

Private Sub acaoDblGrid(grid, tiponota As String)
    If linhaSelecionaValida(grid) = True Then
    
        Dim nf As notaFiscal
        
        lblMSGNota.Caption = ""
        
        nf.pedido = grid.TextMatrix(grid.Row, 2)
        If nf.pedido = Empty Then
            lblMSGNota.Caption = "N�mero de pedido n�o encontrado"
            Exit Sub
        End If
        
        nf.loja = grid.TextMatrix(grid.Row, 1)
        If nf.loja = Empty Then
            lblMSGNota.Caption = "Loja n�o encontrada"
            Exit Sub
        End If
        
        nf.CNPJ = obterCNPJloja
        If nf.CNPJ = Empty Then
            
            lblMSGNota.Caption = "CNPJ n�o encontrado"
            Exit Sub
        End If
        
        If abrirAqruivo = True Then
            abrirTXT nf, tiponota
        Else
        
            If grid.CellForeColor <> vbRed And Not grid.TextMatrix(grid.Row, 4) Like "*Cancelamento*" Then
                ImprimirNota nf, tiponota
            Else
                abrirArquivoResposta nf, tiponota
            End If
        
        End If
        
    End If
End Sub

Private Sub grdLogSig_DblClick()
    acaoDblGrid grdLogSig, "NOTA"
End Sub



Private Sub grdLogSig_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 17 Then
        abrirAqruivo = True
        grdLogSig.BackColorSel = vbBlue
    End If
End Sub

Public Function mensagemExluir(NomeCampo As String) As Boolean

    If MsgBox("Deseja exluir " & NomeCampo & "?", vbQuestion + vbYesNo, "Excluir") = vbYes Then
            mensagemExluir = True
    Else
            mensagemExluir = False
    End If
    
End Function

Private Sub grdLogSig_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then txtNFe_KeyPress 27
End Sub

Private Sub grdLogSig_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then
        If linhaSelecionaValida(grdLogSig) Then
            Dim nomeArquivo As String
            nomeArquivo = procuraArquivo(grdLogSig.TextMatrix(grdLogSig.Row, 2), grdLogSig.TextMatrix(grdLogSig.Row, 1))
            If mensagemExluir(nomeArquivo) = True Then
                deletaArquivo GLB_EnderecoPastaRESP & nomeArquivo
                qtdeLinhaAnterior = 0
                cmdAtualizar_Click
            End If
        End If
    End If

    If KeyCode = 17 Then
        abrirAqruivo = False
        grdLogSig.BackColorSel = &H343434
    End If
    
End Sub

Private Function procuraArquivo(pedido As String, loja As String) As String
    Dim sFile As String
    Dim nomeArquivoPesquisa As String
    Dim arq As File
    
    sFile = Dir(GLB_EnderecoPastaRESP & "*.txt", vbDirectory)
    
    Do While sFile <> ""
        If InStr(sFile, ".txt") > 0 Then
            If sFile Like "*" & pedido & "*" And sFile Like "*" & obterCNPJloja & "*" Then
                procuraArquivo = sFile
            End If
        End If
    sFile = Dir
    Loop
End Function

Private Sub Label2_Click()
    Tempo = 10
    timerSairSistema_Timer
End Sub

Private Sub grdLogSig_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mostraPopup grdLogSig
End Sub

Private Sub grdLogSigSAT_Click()
    If grdLogSigSAT.CellForeColor = vbRed Then
        grdLogSigSAT.BackColorSel = vbRed
    Else
        grdLogSigSAT.BackColorSel = &H343434
    End If
    grdLogSig.Row = 0
    abrirAqruivo = False
End Sub

Private Sub grdLogSigSAT_DblClick()
    acaoDblGrid grdLogSigSAT, "SAT"
End Sub

Private Sub abrirArquivoResposta(nf As notaFiscal, tiponota As String)
    
    Dim Arquivo As String
    
    Dim informacaoArquivo As String
    Dim mensagemArquivoTXT As TextStream
    Dim Resultado As String
    Dim fso As New FileSystemObject
    
    If tiponota = "NOTA" Then
        Arquivo = Dir(GLB_EnderecoPastaRESP & "*" & nf.numero & "#" & nf.CNPJ & ".txt", vbDirectory)
    Else
        Arquivo = Dir(GLB_EnderecoPastaRESP & "*" & nf.pedido & "#" & nf.CNPJ & ".txt", vbDirectory)
    End If
    
     If Arquivo <> "" Then
        Screen.MousePointer = 11
    
        Tempo = 200
        
        Set mensagemArquivoTXT = fso.OpenTextFile(GLB_EnderecoPastaRESP & Arquivo)
        informacaoArquivo = mensagemArquivoTXT.ReadAll
        mensagemArquivoTXT.Close
        
        MsgBox informacaoArquivo, vbInformation, Arquivo
        
        Screen.MousePointer = 0
    Else
        lblMSGNota.Caption = "N�o foi encontrado o arquivo resp"
    End If
    
End Sub

Private Sub grdLogSigSAT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 17 Then
        abrirAqruivo = True
        grdLogSigSAT.BackColorSel = vbBlue
    End If
End Sub

Private Sub grdLogSigSAT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then txtNFe_KeyPress 27
End Sub

Private Sub grdLogSigSAT_KeyUp(KeyCode As Integer, Shift As Integer)

'    If KeyCode = vbKeyDelete Then
'        If linhaSelecionaValida(grdLogSigSAT) Then
'            Dim nomeArquivo As String
'            nomeArquivo = procuraArquivo(grdLogSigSAT.TextMatrix(grdLogSigSAT.Row, 2), grdLogSigSAT.TextMatrix(grdLogSig.Row, 1))
'            If mensagemExluir(nomeArquivo) = True Then
'                deletaArquivo GLB_EnderecoPastaRESP & nomeArquivo
'                qtdeLinhaAnterior = 0
'                cmdAtualizar_Click
'            End If
'        End If
'    End If

    If KeyCode = 17 Then
        abrirAqruivo = False
        grdLogSigSAT.BackColorSel = &H343434
    End If
    
End Sub

Private Sub grdLogSigSAT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mostraPopup grdLogSigSAT
End Sub

Private Sub Image1_Click()

End Sub

Private Sub lblStatusImpressao_Click()
    Tempo = 200
End Sub

Private Sub optPesquisaNumero_Click()
    txtNFe.SetFocus
End Sub

Private Sub optPesquisaPedido_Click()
    txtNFe.SetFocus
End Sub


Private Sub timerExibirMSG_Timer()
    Tempo = Tempo + 1
    statusFuncionamento mensagemStatus
    If Tempo > 60 Then
        timerSairSistema.Enabled = False
        Unload Me
    End If
End Sub

Private Sub timerSairSistema_Timer()
    Tempo = Tempo + 1
    statusFuncionamento mensagemStatus
    carregaArquivoUnico
    If Tempo > 60 Then 'felipetef
        timerExibirMSG.Enabled = False
        'ImprimeComprovanteTEF ComprovantePagamentoFila
        Unload Me
    End If
End Sub

Private Sub timerVerificaResposta_Timer()
    'PrintForm
    carregaArquivo
    timerVerificaResposta.Interval = tempoVerificacaoResposta
End Sub

Public Sub txtNFe_KeyPress(KeyAscii As Integer)

    lblMSGNota.Caption = ""
    cmdImprimir.Enabled = False
    cmdEmail.Enabled = False
    
    If KeyAscii = 13 Then
        
        Dim rsNFE As New ADODB.Recordset
        
        txtNFe.text = Val(txtNFe.text)
        
        If optPesquisaPedido.Value Then sql = " and numeroped = "
        If optPesquisaNumero.Value Then sql = " and nf = "
        sql = "select top 1 nf as nf, " & vbNewLine _
              & "ChaveNFe as chave, " & vbNewLine _
              & "serie as serie, " & vbNewLine _
              & "lo_cgc as cgc, " & vbNewLine _
              & "numeroped as pedido, " & vbNewLine _
              & "codoper as cfop, " & vbNewLine _
              & "dataprocesso as data " & vbNewLine _
              & "from nfcapa, loja " & vbNewLine _
              & "where tiponota in ('V','T','E','S','R') " & vbNewLine _
              & "and lo_loja = lojaorigem " & vbNewLine _
              & "and lojaorigem = '" & wLoja & "' " & sql & txtNFe.text & vbNewLine _
              & "order by dataemi desc"
        
        rsNFE.CursorLocation = adUseClient
        rsNFE.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            
        If rsNFE.EOF Then
            cmdImprimir.Enabled = False
            cmdEmail.Enabled = False
            cmdTransmitir.Enabled = False
            optPesquisaPedido.Value = True
            txtNFe.text = Empty
            lblMSGNota.Caption = "Nota Fiscal n�o encontrada"
            Exit Sub
        Else
            
            nf.numero = RTrim(rsNFE("NF"))
            nf.chave = RTrim(rsNFE("chave"))
            nf.eSerie = RTrim(rsNFE("serie"))
            nf.CNPJ = RTrim(rsNFE("cgc"))
            nf.pedido = RTrim(rsNFE("pedido"))
            nf.cfop = RTrim(rsNFE("cfop"))
            wPedido = RTrim(rsNFE("pedido"))
            pedido = RTrim(rsNFE("pedido"))
            If Format(rsNFE("data"), "YYYY/MM/DD") < GLB_DataInicial And GLB_Administrador = False Then
                MsgBox "N�o � permitido emitir NFe/Cupom fora da data do movimento", vbExclamation, "Emiss�o de NFe/Cupom"
                cmdTransmitir.Enabled = False
            Else
                cmdTransmitir.Enabled = True
            End If
            
            If nf.chave <> "" Then
                cmdImprimir.Enabled = True
                cmdImprimir.ToolTipText = ""
                If nf.eSerie = "NE" Then
                    cmdEmail.Enabled = True
                End If
            End If
            
        End If
        
        rsNFE.Close

    End If

    If KeyAscii = 27 Then
        Unload Me
    End If
    
End Sub

Private Sub limpaTabelaArquivos()

    Dim sql As String
    
    sql = "delete NFE_NFLojas where NFL_DataEmissao < dateadd(d,-6,GETDATE())"
    rdoCNLoja.Execute sql
    
End Sub


Public Sub leituraEstrutura(campo As String)
    Dim ado_estrutura As New ADODB.Recordset
    
    Call montaWhereFiscal

    sql = "select etr_sequencia, etr_campo, etr_tabela_de, " & vbNewLine & _
          "etr_campo_de, ETR_ROTULO " & vbNewLine & _
          "from NFE_Estrutura " & vbNewLine & _
          "where etr_rotulo = '" & campo & "' and etr_tabela_de <> '' AND etr_campo_de <> ''"

    ado_estrutura.CursorLocation = adUseClient
    ado_estrutura.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

    If campo = "PROD" Or campo = "DET" Or campo = "PISALIQ" Or campo = "COFINSALIQ" Or campo = "IPI" Or campo = "IPITRIB" Or campo = "ICMSUFDEST" Or campo = "ICMSSN102" Then
        
        Dim ado_campo As New ADODB.Recordset
            
            sql = "select h_nItem item " & _
                  "from " & ado_estrutura("etr_tabela_de") & _
                  " where " & whereNotaFiscal & " order by h_nItem"
    
            ado_campo.CursorLocation = adUseClient
            ado_campo.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic


        If campo = "PROD" Then
            sql = insertTabelaNFLojas & Trim(ado_estrutura("etr_sequencia")) - 2 & "','--','','" & _
                      nf.loja & "','" & nf.numero & "','" & Format(Date, "YYYY/MM/DD") & "')"
            rdoCNLoja.Execute sql
        End If


        Do While Not ado_campo.EOF
        
            If campo = "PROD" Then
                sql = insertTabelaNFLojas & Trim(ado_estrutura("etr_sequencia")) + (500 * (Trim(ado_campo("item")) - 1)) - 2 & "','[DET]','','" & _
                nf.loja & "','" & nf.numero & "','" & Format(Date, "YYYY/MM/DD") & "')"
                rdoCNLoja.Execute sql
            End If
            
            sql = insertTabelaNFLojas & Trim(ado_estrutura("etr_sequencia")) + (500 * (Trim(ado_campo("item")) - 1)) - 1 & "','[" & campo & "]','','" & _
                  nf.loja & "','" & nf.numero & "','" & Format(Date, "YYYY/MM/DD") & "')"
            
            rdoCNLoja.Execute sql
            ado_campo.MoveNext
        Loop

        ado_campo.Close
    ElseIf campo = "ICMS00" Or campo = "ICMS10" Or campo = "ICMS20" Or campo = "ICMS30" Or campo = "ICMS40" Or campo = "ICMS51" Or campo = "ICMS60" Or campo = "ICMS70" Or campo = "ICMS90" Or campo = "DUP" Or campo = "ICMSSN102" Or campo = "DETPAG" Then

    
    Else
        sql = insertTabelaNFLojas & Trim(ado_estrutura("etr_sequencia")) - 1 & "','[" & campo & "]','','" & _
              nf.loja & "','" & nf.numero & "','" & Format(Date, "YYYY/MM/DD") & "')"
        rdoCNLoja.Execute sql
    End If

    
    Do While Not ado_estrutura.EOF
        If campo = "PROD" Or campo = "DET" Or campo = "ICMS00" Or campo = "ICMS10" Or campo = "ICMS20" Or campo = "ICMS30" Or campo = "ICMS40" Or campo = "ICMS51" Or campo = "ICMS60" Or campo = "ICMS70" Or campo = "ICMS90" Or campo = "PISALIQ" Or campo = "COFINSALIQ" Or campo = "IPI" Or campo = "IPITRIB" Or campo = "ICMSUFDEST" Or campo = "ICMSSN102" Then
            gravaVariosDado campo, ado_estrutura
        ElseIf campo = "DUP" Then
            gravaDadosDUP campo, ado_estrutura
        ElseIf campo = "DETPAG" Then
            gravaDadosPAG campo, ado_estrutura
        Else
            gravaDados campo, ado_estrutura
        End If
    Loop
    
    sql = "delete NFE_NFLojas where NFL_Descricao = '    voutro' and NFL_Dados = '0.00' AND NFL_Sequencia > 200"
    rdoCNLoja.Execute (sql)
    
    If nf.cfop = "5602" Or nf.cfop = "5605" Then
        zerarValoresNota (nf.numero)
    End If
    
    ado_estrutura.Close
    
'frmLogNotaFiscal.mensagemLOG2 frmLogNotaFiscal.grdLog, Now, 100, "181", nf.numero, "Rotulo '" & campo & "' inserido com sucesso"

End Sub

Private Sub zerarValoresNota(numeroNF As String)
    Dim sql As String
    
    sql = "update nfe_nflojas " & vbNewLine & _
          "set nfl_dados = '0.00'" & vbNewLine & _
          "where nfl_nroNFE = '" & numeroNF & "'" & vbNewLine & _
          "and NFL_Descricao like '%VPROD%'" & vbNewLine & _
          "or NFL_Descricao like '%VUNCOM%'" & vbNewLine & _
          "or NFL_Descricao like '%VUNTRIB%'" & vbNewLine & _
          "or NFL_Descricao like '%VNF%'"
    
    rdoCNLoja.Execute (sql)

End Sub

Private Function montaWhereFiscal()
    whereNotaFiscal = "eLoja = '" & nf.loja & "' and eNF = '" & nf.numero & _
                      "' and eSerie = '" & nf.eSerie & "'"
End Function

Public Sub atualizaNota(campo As String)
    Dim ado_estrutura As New ADODB.Recordset

    sql = "select top 1 etr_rotulo, etr_tabela_de " & _
          "from NFE_Estrutura " & _
          "where etr_rotulo = '" & campo & "' and etr_tabela_de <> '' and etr_campo_de <> ''"
    ado_estrutura.CursorLocation = adUseClient
    ado_estrutura.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

        sql = "update nfe_ide " & _
              " set situacao = 'P' " & _
              "where " & whereNotaFiscal

        rdoCNLoja.Execute sql

    ado_estrutura.Close
End Sub




Private Sub montaCamposRotulo()

    ReDim vetCampos(34)
    
    vetCampos(0) = "IDE":           vetCampos(1) = "DANFE":         vetCampos(2) = "EMAIL":
    vetCampos(3) = "NFREF":         vetCampos(4) = "EMIT":          vetCampos(5) = "ENDEREMIT"
    vetCampos(6) = "DEST":          vetCampos(7) = "ENDERDEST":     vetCampos(8) = "ICMSTOT"
    vetCampos(9) = "TRANSP":        vetCampos(10) = "TRANSPORTA":   vetCampos(11) = "VEICTRANSP"
    vetCampos(12) = "VOL":
    vetCampos(13) = "INFADIC":      vetCampos(14) = "OBSCONT"
    vetCampos(15) = "FAT":          vetCampos(16) = "DUP":
    vetCampos(17) = "PAG":          vetCampos(18) = "DETPAG":
    vetCampos(19) = "PROD"
    vetCampos(20) = "ICMS00":       vetCampos(21) = "ICMS10":       vetCampos(22) = "ICMS20":
    vetCampos(23) = "ICMS30":       vetCampos(24) = "ICMS40":       vetCampos(25) = "ICMS51":

    vetCampos(26) = "ICMS60":       vetCampos(27) = "ICMS70":       vetCampos(28) = "ICMS90":
    vetCampos(29) = "ICMSSN102":
    vetCampos(30) = "IPI":          vetCampos(31) = "IPITRIB":      vetCampos(32) = "PISALIQ":
    vetCampos(33) = "COFINSALIQ":   vetCampos(34) = "ICMSUFDEST"

End Sub

Private Sub gravaVariosDado(campo As String, ado_estrutura As ADODB.Recordset)
    Dim ado_campo As New ADODB.Recordset
    Dim informacao As String
  
    sql = "select " & Trim(ado_estrutura("etr_campo_de")) & " informacao, h_nItem item, N_cstICMS CST " & _
          "from " & ado_estrutura("etr_tabela_de") & " " & _
          "where " & whereNotaFiscal & " and " & Trim(ado_estrutura("etr_campo_de")) & " is not null " & _
          "order by h_nItem"
    
    Debug.Print ado_estrutura("etr_campo_de")
    
    ado_campo.CursorLocation = adUseClient
    ado_campo.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    Do While Not ado_campo.EOF
    
        If Mid(ado_estrutura("ETR_CAMPO"), 5, 1) = "V" Or Mid(ado_estrutura("ETR_CAMPO"), 5, 1) = "Q" Or Mid(ado_estrutura("ETR_CAMPO"), 5, 1) = "P" Then
            If Trim(ado_estrutura("etr_campo")) = "VFRETE" Or Trim(ado_estrutura("etr_campo")) = "VSEG" Or Trim(ado_estrutura("etr_campo")) = "VDESC" Then
                If ado_campo("informacao") = "0" And campo = "PROD" Then
                    informacao = "''"
                Else
                    informacao = Replace(Format(Trim(ado_campo("informacao")), "0.00"), ",", ".")
                End If
            Else
                informacao = Replace(Format(Trim(ado_campo("informacao")), "0.00"), ",", ".")
            End If
            sql = insertTabelaNFLojas & _
                  (Trim(ado_estrutura("etr_sequencia")) + (500 * (Trim(ado_campo("item")) - 1))) & "', '" & _
                  ado_estrutura("etr_campo") & "', " & informacao & ", '" & _
                  nf.loja & "', '" & nf.numero & "', '" & Format(Date, "YYYY/MM/DD") & "')"
        Else
            If (Trim(ado_estrutura("ETR_CAMPO")) = "CST" And Mid(campo, 1, 4) = "ICMS") Or (Trim(ado_estrutura("ETR_CAMPO")) = "ORIG" And Mid(campo, 1, 6) = "ICMSSN") Then
            'If (Trim(ado_estrutura("ETR_CAMPO")) = "CST" And Mid(campo, 1, 4) = "ICMS") Or (Trim(ado_estrutura("ETR_CAMPO")) = "ORIG" And Mid(campo, 1, 6) = "ICMSSN") Then
                'SQL = "update NFE_NFLojas set nfl_descricao = '[ICMS" & Trim(ado_campo("informacao")) & "]' " & _
                      "where nfl_loja = " & nf.loja & " and nfl_nroNFE = " & nf.numero & " and nfl_sequencia = " & (Trim(ado_estrutura("etr_sequencia")) + (54 * (Trim(ado_campo("item")) - 1))) - 1
                      If Trim(ado_estrutura("ETR_CAMPO")) = "CST" Then
                        
                      End If
                      
                sql = insertTabelaNFLojas & _
                      (Trim(ado_estrutura("etr_sequencia")) + (500 * (Trim(ado_campo("item")) - 1))) - 2 & "', '" & _
                      "[IMPOSTO]', '" & " " & "', '" & _
                      nf.loja & "', '" & nf.numero & "', '" & Format(Date, "YYYY/MM/DD") & "')"
                If Not (Trim(ado_estrutura("ETR_CAMPO")) = "ORIG" And Mid(campo, 1, 6) = "ICMSSN") Then
                    Dim cst As String
                    cst = Format(Trim(ado_campo("informacao")), "00")
                    If cst = "41" Or cst = "50" Then cst = "40"

                    
                    
                      sql = sql & vbNewLine & insertTabelaNFLojas & _
                            (Trim(ado_estrutura("etr_sequencia")) + (500 * (Trim(ado_campo("item")) - 1))) - 1 & "', '" & _
                            "[ICMS" & cst & "]', '" & " " & "', '" & _
                            nf.loja & "', '" & nf.numero & "', '" & Format(Date, "YYYY/MM/DD") & "')"
                            
                End If
                'FELIPE AQUI 2017
                sql = sql & vbNewLine & insertTabelaNFLojas & _
                      (Trim(ado_estrutura("etr_sequencia")) + (500 * (Trim(ado_campo("item")) - 1))) - 0 & "', '" & _
                      ado_estrutura("etr_campo") & "', '" & Trim(ado_campo("informacao")) & "', '" & _
                      nf.loja & "', '" & nf.numero & "', '" & Format(Date, "YYYY/MM/DD") & "')"
                
            Else
                sql = insertTabelaNFLojas & _
                      (Trim(ado_estrutura("etr_sequencia")) + (500 * (Trim(ado_campo("item")) - 1))) + 1 & "', '" & _
                      ado_estrutura("etr_campo") & "', '" & Replace(Trim(ado_campo("informacao")), ",", ".") & "', '" & _
                      nf.loja & "', '" & nf.numero & "', '" & Format(Date, "YYYY/MM/DD") & "')"
            End If
        End If
        
        If Mid(ado_estrutura("ETR_ROTULO"), 5, 2) = "SN" Then
            Debug.Print "ICMS SN"
        End If
        
        If Mid(campo, 1, 4) = "ICMS" And Format(ado_campo("CST"), "00") = "41" And icms41 = False Then
            If LTrim(ado_estrutura("etr_campo")) = "ORIG" Then
                icms41 = True
                rdoCNLoja.Execute sql
                Exit Sub
            End If
            rdoCNLoja.Execute sql
            
        ElseIf Mid(campo, 1, 4) = "ICMS" And Format(ado_campo("CST"), "00") = "50" And icms50 = False Then
            If LTrim(ado_estrutura("etr_campo")) = "ORIG" Then
                icms50 = True
                rdoCNLoja.Execute sql
                Exit Sub
            End If
            rdoCNLoja.Execute sql
            
        ElseIf Mid(campo, 1, 4) = "ICMS" And Mid(ado_estrutura("ETR_ROTULO"), 5, 2) = "SN" And Trim(ado_campo("CST")) = "2" Then
        'If Mid(campo, 1, 4) = "ICMS" And Mid(ado_estrutura("ETR_ROTULO"), 5, 2) = Format(Trim(ado_campo("CST")), "00") Then
           ' MsgBox "campo 1"
            rdoCNLoja.Execute sql
        ElseIf Mid(campo, 1, 4) = "ICMS" And Mid(ado_estrutura("ETR_ROTULO"), 5, 2) = Format(Trim(ado_campo("CST")), "00") Then
            'MsgBox "campo 2"
            rdoCNLoja.Execute sql
        ElseIf campo = "ICMSUFDEST" And informacao <> "0.00" Then
'            MsgBox "campo 3"
            rdoCNLoja.Execute sql
        ElseIf Mid(campo, 1, 4) = "ICMS" And Mid(ado_estrutura("ETR_ROTULO"), 5, 2) <> Format(Trim(ado_campo("CST")), "00") Then
            Debug.Print "ICMS SN"

        Else
 '           MsgBox "campo OUTROS"
            rdoCNLoja.Execute sql
        End If
        
        ado_campo.MoveNext
    Loop
    
    ado_campo.Close
    ado_estrutura.MoveNext
End Sub



Private Sub gravaDados(campo As String, ado_estrutura As ADODB.Recordset)

    Dim ado_campo As New ADODB.Recordset
    Dim informacao As String
    
    sql = "select top 1 " & RTrim(ado_estrutura("etr_campo_de")) & " as Informacao " & vbNewLine & _
          "from " & ado_estrutura("etr_tabela_de") & " " & vbNewLine & _
          "where " & whereNotaFiscal & " and " & ado_estrutura("etr_campo_de") & " is not null"
    
    ado_campo.CursorLocation = adUseClient
    ado_campo.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    Do While Not ado_campo.EOF
        If Mid(ado_estrutura("ETR_CAMPO"), 5, 1) = "V" Or Mid(ado_estrutura("ETR_CAMPO"), 5, 1) = "Q" Then
            If Trim(ado_estrutura("ETR_CAMPO")) = "VERPROC" Then
                informacao = "nfe3g"
            Else
                informacao = ado_campo("informacao")
            End If
            If Trim(ado_estrutura("ETR_CAMPO")) <> "QVOL" Then
                informacao = Replace(Format(informacao, "0.00"), ",", ".")
            End If
        ElseIf Trim(Mid(ado_estrutura("ETR_CAMPO"), 5, 10)) = "DHEMI" Or Trim(Mid(ado_estrutura("ETR_CAMPO"), 5, 10)) = "DHSAIENT" Then
            informacao = Format(ado_campo("informacao"), "YYYY-MM-DD") & "T" & Format(Time, "hh:mm:ss")
        ElseIf Mid(ado_estrutura("ETR_CAMPO"), 5, 1) = "D" Then
            informacao = Format(ado_campo("informacao"), "YYYY-MM-DD")
        Else
            informacao = ado_campo("informacao")
        End If
        
        'Tratamento de erro do IE = 0 e isento
        If Trim(ado_estrutura("ETR_CAMPO")) = "IE" Then
            If Trim(informacao) = "0" Then
                informacao = ""
            ElseIf UCase(Trim(informacao)) = "ISENTO" Then
                informacao = "ISENTO"
            End If
        End If
        
        'Tratamento de erro do TRANSP = NULL
        If Trim(ado_estrutura("ETR_CAMPO")) = "[TRANSP]" Then
            If IsNull(Trim(informacao)) Then
                informacao = "1"
            'ElseIf Trim(informacao) = "isento" Then
                'informacao = "ISENTO"
            End If
        End If
        
        sql = insertTabelaNFLojas & _
              Trim(ado_estrutura("etr_sequencia")) & "', '" & ado_estrutura("etr_campo") & _
              "', '" & RTrim(informacao) & "', '" & _
              nf.loja & "', '" & nf.numero & "', '" & Format(Date, "YYYY/MM/DD") & "')"
              
        rdoCNLoja.Execute sql
              
        ado_campo.MoveNext
    Loop
    ado_campo.Close
    ado_estrutura.MoveNext
    
    
End Sub

Private Sub criarArquivorDACTE(nf As notaFiscal, chaveAcesso As String)

    Open GLB_EnderecoPastaFIL & _
    "dacfesat" & (Format(nf.pedido, "000000000")) & "#" & _
    nf.CNPJ & ".txt" For Output As #1
            
        Print #1, "CHAVESAT     = " & chaveAcesso
        Print #1, "IMPRESSORA   = " & GLB_Impressora00
    
    Close #1

End Sub

Private Sub gravaDadosPAG(campo As String, ado_estrutura As ADODB.Recordset)

    Dim ado_campo As New ADODB.Recordset
    Dim informacao As String
    Dim i As Byte
    
    i = 0
    
    sql = "select " & RTrim(ado_estrutura("etr_campo_de")) & " as Informacao " & vbNewLine & _
          "from " & ado_estrutura("etr_tabela_de") & " " & vbNewLine & _
          "where " & whereNotaFiscal & " and " & ado_estrutura("etr_campo_de") & " is not null"
    
    ado_campo.CursorLocation = adUseClient
    ado_campo.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    Do While Not ado_campo.EOF
        If Mid(ado_estrutura("ETR_CAMPO"), 5, 1) = "V" Or Mid(ado_estrutura("ETR_CAMPO"), 5, 1) = "Q" Then
            informacao = Replace(Format(ado_campo("informacao"), "0.00"), ",", ".")
        ElseIf Mid(ado_estrutura("ETR_CAMPO"), 5, 1) = "D" Then
            informacao = Format(ado_campo("informacao"), "YYYY-MM-DD")
        Else
            informacao = ado_campo("informacao")
        End If
        
        sql = insertTabelaNFLojas & _
              Trim(ado_estrutura("etr_sequencia") + (i)) & "', '" & ado_estrutura("etr_campo") & _
              "', '" & RTrim(informacao) & "', '" & _
              nf.loja & "', '" & nf.numero & "', '" & Format(Date, "YYYY/MM/DD") & "')"
              
        rdoCNLoja.Execute sql
        
        If ado_estrutura("etr_campo") = "    INDPAG" Then
            sql = insertTabelaNFLojas & _
            Trim(ado_estrutura("etr_sequencia") + (i) - 1) & "', '[" & RTrim(ado_estrutura("etr_ROTULO")) & _
            "]', '" & "" & "', '" & _
            nf.loja & "', '" & nf.numero & "', '" & Format(Date, "YYYY/MM/DD") & "')"
              
            rdoCNLoja.Execute sql
            i = i + 1
        End If
              
        ado_campo.MoveNext
        i = i + 5
    Loop
    ado_campo.Close
    ado_estrutura.MoveNext
    
    
End Sub

Private Sub gravaDadosDUP(campo As String, ado_estrutura As ADODB.Recordset)

    Dim ado_campo As New ADODB.Recordset
    Dim informacao As String
    Dim i As Byte
    
    i = 0
    
    sql = "select " & RTrim(ado_estrutura("etr_campo_de")) & " as Informacao " & vbNewLine & _
          "from " & ado_estrutura("etr_tabela_de") & " " & vbNewLine & _
          "where " & whereNotaFiscal & " and " & ado_estrutura("etr_campo_de") & " is not null"
    
    ado_campo.CursorLocation = adUseClient
    ado_campo.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    Do While Not ado_campo.EOF
        If Mid(ado_estrutura("ETR_CAMPO"), 5, 1) = "V" Or Mid(ado_estrutura("ETR_CAMPO"), 5, 1) = "Q" Then
            If Trim(ado_estrutura("ETR_CAMPO")) = "VERPROC" Then
                informacao = "nfe3g"
            Else
                informacao = ado_campo("informacao")
            End If
            If Trim(ado_estrutura("ETR_CAMPO")) <> "QVOL" Then
                informacao = Replace(Format(informacao, "0.00"), ",", ".")
            End If
        ElseIf Trim(Mid(ado_estrutura("ETR_CAMPO"), 5, 10)) = "DHEMI" Or Trim(Mid(ado_estrutura("ETR_CAMPO"), 5, 10)) = "DHSAIENT" Then
            informacao = Format(ado_campo("informacao"), "YYYY-MM-DD") & "T" & Format(Time, "hh:mm:ss")
        ElseIf Mid(ado_estrutura("ETR_CAMPO"), 5, 1) = "D" Then
            informacao = Format(ado_campo("informacao"), "YYYY-MM-DD")
        Else
            informacao = ado_campo("informacao")
        End If
        
        
        sql = insertTabelaNFLojas & _
              Trim(ado_estrutura("etr_sequencia") + (i)) & "', '" & ado_estrutura("etr_campo") & _
              "', '" & RTrim(informacao) & "', '" & _
              nf.loja & "', '" & nf.numero & "', '" & Format(Date, "YYYY/MM/DD") & "')"
              
        rdoCNLoja.Execute sql
        
        If ado_estrutura("etr_campo") = "    NDUP" Then
            sql = insertTabelaNFLojas & _
            Trim(ado_estrutura("etr_sequencia") + (i) - 1) & "', '[" & RTrim(ado_estrutura("etr_ROTULO")) & _
            "]', '" & "" & "', '" & _
            nf.loja & "', '" & nf.numero & "', '" & Format(Date, "YYYY/MM/DD") & "')"
              
            rdoCNLoja.Execute sql
            'i = i + 1
        End If
              
        ado_campo.MoveNext
        i = i + 5
    Loop
    ado_campo.Close
    ado_estrutura.MoveNext
    
    
End Sub

Private Sub criarArquivorEmail(nf As notaFiscal, chaveAcesso As String, email As String, destinatario As String)

    Open GLB_EnderecoPastaFIL & _
    "email" & (Format(nf.numero, "000000000")) & "#" & _
    nf.CNPJ & ".txt" For Output As #1
            
        Print #1, "CHAVENFE     = " & chaveAcesso
        Print #1, "DESTINATARIO   = " & email
        Print #1, "ASSUNTO   = Nota Fiscal Eletr�nica " & wRazao
        Print #1, "MENSAGEM   = Ol�" & destinatario & ", voc� est� recebendo uma c�pia da DANFE e o arquivo XML"
        Print #1, "NOMEEMITENTE = " & wRazao
        Print #1, "ANEXOPDF = sim"
        Print #1, "ANEXOXML = sim"
    
    Close #1

End Sub

Private Sub criarArquivorDanfe(nf As notaFiscal, chaveAcesso As String)

    Open GLB_EnderecoPastaFIL & _
    "danfe" & (Format(nf.numero, "000000000")) & "#" & _
    nf.CNPJ & ".txt" For Output As #1
            
        Print #1, "CHAVENFE     = " & chaveAcesso
        Print #1, "IMPRESSORA   = " & Glb_ImpNotaFiscal
    
    Close #1

End Sub

Private Sub cancelaNE(nf As notaFiscal)

    Dim ado_NFe As New ADODB.Recordset
    Dim sql As String
    Dim xJust As String
      
    If nf.chave <> "" Then
        sql = "select xJust from nfe_ide " & vbNewLine & "where enf = '" & nf.numero & "' and eloja = '" & nf.loja & "'"
        ado_NFe.CursorLocation = adUseClient
        ado_NFe.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
        xJust = RTrim(ado_NFe("xJust"))
        ado_NFe.Close
          
        Open GLB_EnderecoPastaFIL & _
        "cancel" & (Format(nf.numero, "000000000")) & "#" & _
        nf.CNPJ & ".txt" For Output As #1
                
            Print #1, "CHAVENFE      = " & nf.chave
            Print #1, "JUSTIFICATIVA = " & xJust
        
        Close #1
    End If

End Sub

Private Sub cancelaSAT(nf As notaFiscal)
    
    Dim rsNFE As New ADODB.Recordset
    
    sql = "select top 1 nf as nf, " & vbNewLine _
        & "ChaveNFe as chave" & vbNewLine _
        & "from nfcapa" & vbNewLine _
        & "where lojaorigem in " & lojasWhere & " " & "" & vbNewLine _
        & "and numeroped = " & nf.pedido
    
    rsNFE.CursorLocation = adUseClient
    rsNFE.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
        If Not rsNFE.EOF Then
            If rsNFE("chave") = Empty Then
                
            Else
                lblMSGNota.Caption = ""
                      
                Open GLB_EnderecoPastaFIL & _
                "cancelsat" & (Format(nf.pedido, "000000000")) & "#" & _
                nf.CNPJ & ".txt" For Output As #1
                
                Print #1, "CHAVESAT     = " & nf.chave
                Print #1, "IMPRESSORA   = " & GLB_Impressora00
                Print #1, "RETORNARESP   = " & "1"
                
                Close #1
          
            End If
        Else
            
        
            lblMSGNota.Caption = "Nota Fiscal n�o encontrada"
        End If
    
    rsNFE.Close

End Sub

Private Sub ImprimirNota(nf As notaFiscal, tiponota As String)
    
    Dim rsNFE As New ADODB.Recordset
    
    sql = "select top 1 nf as nf, " & vbNewLine _
        & "ChaveNFe as chave" & vbNewLine _
        & "from nfcapa" & vbNewLine _
        & "where " & vbNewLine _
        & "lojaorigem in " & lojasWhere & " " & "" & vbNewLine _
        & "and numeroped = " & nf.pedido
    
    rsNFE.CursorLocation = adUseClient
    rsNFE.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
        If Not rsNFE.EOF Then
            If rsNFE("chave") = Empty Then
                lblMSGNota.Caption = "Chave de acesso n�o encontrada"
            Else
                lblMSGNota.Caption = ""
                
                If frameDadosNotaFiscal.Visible = False Then
                    If tiponota = "NOTA" Then criarArquivorDanfe nf, rsNFE("chave")
                    If tiponota = "SAT" Then criarArquivorDACTE nf, rsNFE("chave")
                ElseIf MsgBox("Deseja imprimir a nota " & nf.numero & "? ", vbQuestion + vbYesNo, "Impress�o de Nota") = vbYes Then
                    If tiponota = "NOTA" Then criarArquivorDanfe nf, rsNFE("chave")
                    If tiponota = "SAT" Then criarArquivorDACTE nf, rsNFE("chave")
                End If
            End If
        Else
        
            lblMSGNota.Caption = "Nota n�o encontrada"
        End If
    
    rsNFE.Close

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function verificaNovoArquivo() As Boolean

    Dim Arquivo As String
    Dim qtdeArquivos As Integer
    'Static qtdeLinhaAnterior As Integer
    On Error GoTo TrataErro
    
    Arquivo = Dir(endArquivoResposta, vbDirectory)
    If Arquivo <> "" Then qtdeArquivos = 1
    
    
    Do While Dir <> ""
        qtdeArquivos = qtdeArquivos + 1
    Loop
    
    If qtdeLinhaAnterior <> qtdeArquivos Then
        qtdeLinhaAnterior = qtdeArquivos
        verificaNovoArquivo = True
    Else
        If qtdeArquivos = 0 Then
            verificaNovoArquivo = True
        Else
            verificaNovoArquivo = False
        End If
    End If
    
    Exit Function
TrataErro:
    Select Case Err.Number
        Case 5
            verificaNovoArquivo = False
        Case Else
            mensagemErroDesconhecido Err, "Verifica��o de novos arquivo"
    End Select
End Function





Private Function lerCampo(informacoes As String, campo As String) As String
    
    If informacoes Like "*" & campo & "=*" Then
        Dim inicioCampo, fimCampo As Integer
    
        inicioCampo = (InStr(informacoes, campo & "=")) + (Len(campo)) + 1
        fimCampo = (InStr(inicioCampo, informacoes, Chr(10))) - inicioCampo - 1
    
        If inicioCampo + fimCampo <> 0 Then
            lerCampo = Mid$(informacoes, inicioCampo, fimCampo)
        End If
    Else
        lerCampo = ""
    End If
    
End Function

Private Function obterCNJPArquivo(Arquivo As String) As String
    obterCNJPArquivo = Mid(Arquivo, InStr(Arquivo, "#") + 1, 14)
End Function

Private Function obterNumNFArquivo(Arquivo, nf As notaFiscal) As String
    Dim numNF As String
    Dim ado_loja As New ADODB.Recordset
    
    If nf.eSerie Like "CE*" Then

        numNF = Mid(nf.chave, 32, 6)
        
        If numNF = "" Or numNF = "0" Then
            
            sql = "select top 1 nf as nf " & vbNewLine & _
            "from nfcapa " & vbNewLine & _
            "where numeroped = '" & nf.pedido & "'"
            
            ado_loja.CursorLocation = adUseClient
            ado_loja.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                
            If Not ado_loja.EOF Then
                obterNumNFArquivo = ado_loja("nf")
            Else
                obterNumNFArquivo = ""
            End If
                
            ado_loja.Close
        Else
            obterNumNFArquivo = numNF
        End If
    Else
        obterNumNFArquivo = Val(Mid(Arquivo, InStr(Arquivo, "#") - 6, 6))
    End If
    
End Function

Private Function obterNumPedidoArquivo(Arquivo As String, nf As notaFiscal) As String

    If nf.eSerie Like "CE*" Then

        obterNumPedidoArquivo = Val(Mid(Arquivo, InStr(Arquivo, "#") - 6, 6))
    Else
    
        Dim ado_loja As New ADODB.Recordset
        
        sql = "select numeroped " & vbNewLine & _
        "from nfcapa " & vbNewLine & _
        "where nf = '" & Val(Mid(Arquivo, InStr(Arquivo, "#") - 6, 6)) & "'" & vbNewLine & _
        "and serie = 'NE'"

        ado_loja.CursorLocation = adUseClient
        ado_loja.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
        
        If Not ado_loja.EOF Then
            obterNumPedidoArquivo = RTrim(ado_loja("numeroped"))
        Else
            obterNumPedidoArquivo = 0
        End If
        ado_loja.Close
        
    End If
    
End Function

Public Function obterLoja(CNPJ As String) As String
On Error GoTo TrataErro
    Dim ado_loja As New ADODB.Recordset
    
    With ado_loja
        sql = "select lo_loja as loja " & vbNewLine & _
        "from loja " & vbNewLine & _
        "where lo_cgc like '%" & CNPJ & "%'"
        
        .CursorLocation = adUseClient
        .Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
        If Not ado_loja.EOF Then
            obterLoja = RTrim(ado_loja("loja"))
        End If
        .Close
    End With
    
    Exit Function
TrataErro:
    Select Case Err.Number
        Case Else
            mensagemErroDesconhecido Err, "Erro ao obter a loja"
    End Select
End Function

Private Sub atualizaNumeroNF(numeropedido, numeroNF)
    Dim sql As String
    
    If numeroNF <> "" Then
    
        sql = "update nfCapa" & vbNewLine & _
              "set nf = '" & numeroNF & "'" & vbNewLine & _
              "where numeroped = '" & numeropedido & "'" & vbNewLine & _
              "and serie like 'CE%'" & vbNewLine & _
              "and NF = '0' " & vbNewLine & _
              "and lojaOrigem in " & lojasWhere & "" & vbNewLine & vbNewLine
    
        rdoCNLoja.Execute sql
        
        sql = "update nfitens" & vbNewLine & _
              "set nf = '" & numeroNF & "'" & vbNewLine & _
              "where numeroped = '" & numeropedido & "'" & vbNewLine & _
              "and serie like 'CE%'" & vbNewLine & _
              "and NF = '0' " & vbNewLine & _
              "and lojaOrigem in " & lojasWhere & "" & vbNewLine & vbNewLine
    
        rdoCNLoja.Execute sql
        
        sql = "update movimentocaixa" & vbNewLine & _
              "set mc_documento = '" & numeroNF & "'" & vbNewLine & _
              "where mc_pedido = '" & numeropedido & "'" & vbNewLine & _
              "and MC_serie like 'CE%'" & vbNewLine & _
              "and MC_documento = '0' " & vbNewLine & _
              "and mc_loja in " & lojasWhere & "" & vbNewLine & vbNewLine
    
        rdoCNLoja.Execute sql
        
        sql = "update CarimboNotaFiscal" & vbNewLine & _
              "set CNF_nf = '" & numeroNF & "'" & vbNewLine & _
              "where CNF_NumeroPed = '" & numeropedido & "'" & vbNewLine & _
              "and cnf_serie like 'CE%'" & vbNewLine & _
              "and CNF_nf = '0' " & vbNewLine & _
              "and CNF_loja in " & lojasWhere & "" & vbNewLine & vbNewLine
    
        rdoCNLoja.Execute sql
        
    End If
    
End Sub

Private Sub atualizaCodigoNF(numeropedido, codigo, lojaNF)
    Dim sql As String
    
    sql = "update nfCapa" & vbNewLine & _
          "set tm = '" & codigo & "'" & vbNewLine & _
          "where numeroped = '" & numeropedido & "'" & vbNewLine & _
          "and lojaOrigem in " & lojasWhere & "" & vbNewLine & vbNewLine

    'Debug.Print sql
    rdoCNLoja.Execute sql
    
    'atualizaArquivo GLB_EnderecoPastaRESP, arquivo, informacaoArquivo, "DMAC=atualizaCodigoNF"
    
End Sub

Private Sub atualizaChaveNF(numeropedido, chaveNF, lojaNF)
    Dim sql As String
    
    If chaveNF <> "" Then
    
        sql = "update nfCapa" & vbNewLine & _
              "set ChaveNFe = '" & chaveNF & "'" & vbNewLine & _
              "where numeroped = '" & numeropedido & "'" & vbNewLine & _
              "and ChaveNFe = ''" & vbNewLine & _
              "and lojaOrigem in " & lojasWhere & "" & vbNewLine & vbNewLine
    
        rdoCNLoja.Execute sql
        
    End If
    
    'atualizaArquivo GLB_EnderecoPastaRESP, arquivo, informacaoArquivo, "DMAC=atualizaChaveNF"
    
End Sub



Public Function carregaArquivoUnico()

    Static Arquivo As String
    On Error GoTo TrataErro
    
    Dim informacaoArquivo As String
    Dim mensagemArquivoTXT As TextStream
    Dim Resultado As String
    Dim fso As New FileSystemObject
    
    
    If nf.eSerie = "NE" Then
        Arquivo = Dir(GLB_EnderecoPastaRESP & "*" & nf.numero & "#" & nf.CNPJ & ".txt", vbDirectory)
    Else
        Arquivo = Dir(GLB_EnderecoPastaRESP & "*" & pedido & "#" & nf.CNPJ & ".txt", vbDirectory)
    End If
    
    If Arquivo <> "" Then
        Screen.MousePointer = 11
    
        Tempo = 200
        
        Set mensagemArquivoTXT = fso.OpenTextFile(GLB_EnderecoPastaRESP & Arquivo)
        informacaoArquivo = mensagemArquivoTXT.ReadAll
        mensagemArquivoTXT.Close
        
    If UCase(Arquivo) Like "*RESP-*" = False Then
          deletaArquivo GLB_EnderecoPastaRESP & Arquivo
    Else
    
        Resultado = lerCampo(informacaoArquivo, "Resultado")
        
        nf.CNPJ = obterCNJPArquivo(Arquivo)
        nf.loja = obterLoja(nf.CNPJ)
        
        If nf.eSerie Like "CE*" Then

            nf.pedido = obterNumPedidoArquivo(Arquivo, nf)
            nf.chave = lerCampo(informacaoArquivo, "ChaveSAT")
            nf.numero = obterNumNFArquivo(Arquivo, nf)
        End If
        
        If nf.eSerie = "NE" Then
            nf.numero = obterNumNFArquivo(Arquivo, nf)
            nf.pedido = obterNumPedidoArquivo(Arquivo, nf)
            nf.chave = lerCampo(informacaoArquivo, "ChaveNFe")
        End If
             
        If Resultado <> "4014" Then
            atualizaArquivoDestalhesNF nf, Arquivo, informacaoArquivo
            atualizaCodigoNF nf.pedido, Resultado, nf.loja
            atualizaChaveNF nf.pedido, nf.chave, nf.loja

        End If
                     
        If Resultado = 4014 Then
             statusFuncionamento "Email enviado com sucesso"
             Esperar 4
        ElseIf Resultado = 100 Or Resultado = 4012 Or Resultado = 9016 Or Resultado = 124 Or Resultado = 4017 Then

        
             statusFuncionamento "Nota emitida e autorizada com sucesso"
             
             atualizaChaveNF nf.pedido, nf.chave, nf.loja
             If nf.eSerie Like "CE*" Then atualizaNumeroNF nf.pedido, nf.numero
             atualizaArquivo GLB_EnderecoPastaRESP, Arquivo, informacaoArquivo, "DMAC=Atualizado"
             
             Call Devolucao(nf.pedido)
             
             If ReImpressao_Dev Then
                Call CriaNotaCredito1(Nf_Dev, Serie_Dev, NfDev_Dev, SerieDev_Dev, DataDev_Dev, ValorNotaCredito_Dev, NotaCredito_Dev, ReImpressao_Dev)
             End If
             
             Esperar 2
             
        ElseIf Resultado = 101 Or Resultado = 9005 Or Resultado = 4005 Then 'Para cancelamentos
             
             cancelaNotaResultado = True
             statusFuncionamento "Nota cancelada com sucesso"
             Esperar 4
             
        ElseIf Resultado = 9012 Then 'Para cancelamentos
             
             statusFuncionamento "Impress�o concluida com sucesso"
             Esperar 4
             
        ElseIf Resultado = 695 Or Resultado = 521 Then 'Erro de ICMS irregular
             statusFuncionamento "Nota Rejeitada. Tentado transmitir novamente"
             Esperar 3
             cmdTransmitir_Click
             Tempo = 0
             timerSairSistema.Enabled = True
             timerSairSistema_Timer
             
        ElseIf Resultado = 778 Then 'ERRO NCM
            
             Dim itemErro As String
             
             itemErro = obterNumeroItem(lerCampo(informacaoArquivo, "Mensagem"))
             
             statusFuncionamento "Erro de NCM na refer�ncia " & obterReferenciaPorItem(nf.pedido, itemErro) & "; Contate a �rea Fiscal"
             Esperar 7
             
        ElseIf Resultado = 4016 Then
             statusFuncionamento "Nota " & nf.numero & " j� autorizada"
             Esperar 4
        Else
             MsgBox informacaoArquivo, vbCritical, "Nota Fiscal Eletronica"
        End If

        
        Screen.MousePointer = 0
        End If
       ' arquivo = Dir
    
    End If
    
    
    Exit Function
TrataErro:
    Select Case Err.Number
        Case 62, 13
            mensagemArquivoTXT.Close
            deletaArquivo GLB_EnderecoPastaRESP & Arquivo
        Case Else
            mensagemErroDesconhecido Err, "Verifica��o de pasta no arquivo unico"
    End Select
End Function




Private Sub atualizaArquivoDestalhesNF(nf As notaFiscal, Arquivo As String, informacaoArquivo As String)

    Dim ado_loja As New ADODB.Recordset
    Dim informacaoSistema As String
      
    sql = "select top 1 " & vbNewLine & _
    "totalnota as totalnota," & vbNewLine & _
    "numeroped as pedido," & vbNewLine & _
    "tiponota as tipo,  " & vbNewLine & _
    "nf as nf  " & vbNewLine & _
    "from nfcapa " & vbNewLine & _
    "where numeroped = '" & nf.pedido & "'"
    
    ado_loja.CursorLocation = adUseClient
    ado_loja.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not ado_loja.EOF Then
        informacaoSistema = vbNewLine & "Pedido=" & ado_loja("pedido") & vbNewLine & _
        "Total Nota=" & Format(ado_loja("totalnota"), "0.00") & vbNewLine & _
        "Numero Nota=" & ado_loja("nf") & vbNewLine & _
        "Tipo Nota=" & ado_loja("tipo") & vbNewLine
    Else
        informacaoSistema = vbNewLine & "(Nenhuma informa��o sobre essa nota foi encontrada)" & vbNewLine
    End If
    
    atualizaArquivo GLB_EnderecoPastaRESP, Arquivo, informacaoArquivo, informacaoSistema
    'informacaoArquivo = informacaoArquivo & informacaoSistema
    ado_loja.Close
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    
    sql = "select top 1 vendedor as vendedor," & vbNewLine & _
    "ve_nome as nomeVendedor  " & vbNewLine & _
    "from nfcapa, vende " & vbNewLine & _
    "where numeroped = '" & nf.pedido & "'" & vbNewLine & _
    "and ve_codigo = vendedor"
    
    ado_loja.CursorLocation = adUseClient
    ado_loja.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not ado_loja.EOF Then
        informacaoSistema = vbNewLine & "Vendedor=" & ado_loja("vendedor") & " - " & ado_loja("nomeVendedor") & vbNewLine
    Else
        informacaoSistema = vbNewLine & "(Nenhuma informa��o sobre o vendedor foi encontrada)" & vbNewLine
    End If
    atualizaArquivo GLB_EnderecoPastaRESP, Arquivo, informacaoArquivo, informacaoSistema
    'informacaoArquivo = informacaoArquivo & informacaoSistema
    ado_loja.Close
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    
    sql = "select top 1 " & vbNewLine & _
    "cliente as codigoCliente,  " & vbNewLine & _
    "ce_razao as nomeCliente, " & vbNewLine & _
    "ce_cgc as cgcCliente " & vbNewLine & _
    "from nfcapa, fin_cliente, vende " & vbNewLine & _
    "where ce_codigoCliente = cliente " & vbNewLine & _
    "and numeroped = '" & nf.pedido & "'"
    
    ado_loja.CursorLocation = adUseClient
    ado_loja.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
        
    If Not ado_loja.EOF Then
        informacaoSistema = vbNewLine & "C�digo Cliente=" & ado_loja("codigoCliente") & vbNewLine & _
        "Nome Cliente=" & ado_loja("nomecliente") & vbNewLine & _
        "CPF/CNPJ Cliente=" & ado_loja("cgcCliente") & vbNewLine
    Else
        informacaoSistema = vbNewLine & "(Nenhuma informa��o sobre o cliente foi encontrada)" & vbNewLine
    End If
    atualizaArquivo GLB_EnderecoPastaRESP, Arquivo, informacaoArquivo, informacaoSistema
    'informacaoArquivo = informacaoArquivo & informacaoSistema
    ado_loja.Close
    
    informacaoSistema = vbNewLine & vbNewLine & "DMAC=Atualizado " & Date & "-" & Time
    atualizaArquivo GLB_EnderecoPastaRESP, Arquivo, informacaoArquivo, informacaoSistema

End Sub

Public Function carregaArquivo()

    Static Arquivo As String
    On Error GoTo TrataErro
    
    If verificaNovoArquivo = True Then
    
        Dim informacaoArquivo As String
        Dim mensagemArquivoTXT As TextStream
        Dim fso As New FileSystemObject
        Dim arq As File
        Dim Resultado As String
        
        limpaGrid grdLogSig
        limpaGrid grdLogSigSAT
               
               
        Arquivo = Dir(endArquivoResposta, vbDirectory)
                
        notaPedentes
        

        
        Do While Arquivo <> ""
            
            Screen.MousePointer = 11
        
            Set mensagemArquivoTXT = fso.OpenTextFile(GLB_EnderecoPastaRESP & Arquivo)
            informacaoArquivo = mensagemArquivoTXT.ReadAll
            mensagemArquivoTXT.Close
                
            If Not UCase(Arquivo) Like "*RESP-CANCEL*" And _
               Not UCase(Arquivo) Like "*RESP-NOTA*" And _
               Not UCase(Arquivo) Like "*RESP-SAT*" Then
               deletaArquivo GLB_EnderecoPastaRESP & Arquivo
            Else
        
                Resultado = lerCampo(informacaoArquivo, "Resultado")
        
                Set arq = fso.GetFile(GLB_EnderecoPastaRESP & Arquivo)
            
                If CDate(left(arq.DateCreated, 10)) <> GLB_DataInicial Then
'                    (resultado = 100 Or _
'                    resultado = 101 Or _
'                    resultado = 204 Or _
'                    resultado = 4016 Or _
'                    resultado = 9005 Or _
'                    resultado = 9016 Or _
'                    resultado = 5000 Or _
'                    resultado = 4012) Then
               
                    deletaArquivo GLB_EnderecoPastaRESP & Arquivo
                    
                Else
              
                    nf.CNPJ = obterCNJPArquivo(Arquivo)
                    nf.loja = obterLoja(nf.CNPJ)
                    
                    If UCase(Arquivo) Like "*SAT*" Then
                        nf.eSerie = GLB_SerieCF
                        nf.pedido = obterNumPedidoArquivo(Arquivo, nf)
                        nf.chave = lerCampo(informacaoArquivo, "ChaveSAT")
                        nf.numero = obterNumNFArquivo(Arquivo, nf)
                    End If
                    
                    If UCase(Arquivo) Like "*NOTA*" Then
                        nf.eSerie = "NE"
                        nf.numero = obterNumNFArquivo(Arquivo, nf)
                        nf.pedido = obterNumPedidoArquivo(Arquivo, nf)
                        nf.chave = lerCampo(informacaoArquivo, "ChaveNFe")
                    End If
                    
                    If lerCampo(informacaoArquivo, "DMAC") = "" And Resultado <> 101 Then
                    
                        atualizaCodigoNF nf.pedido, Resultado, nf.loja
                        atualizaChaveNF nf.pedido, nf.chave, nf.loja
                        If nf.eSerie Like "CE*" Then atualizaNumeroNF nf.pedido, nf.numero
                        
                        atualizaArquivoDestalhesNF nf, Arquivo, informacaoArquivo
                        'atualizaArquivo GLB_EnderecoPastaRESP, ARQUIVO, informacaoArquivo, "DMAC=Atualizado BD pelo segundo metodo"
                        
                    End If
               
                
                    If UCase(Arquivo) Like "*SAT*" = True Then
                    
                        mensagemLOG2 grdLogSigSAT, _
                                arq.DateCreated, _
                                lerCampo(informacaoArquivo, "Resultado"), _
                                nf.loja, _
                                nf.numero, _
                                nf.pedido, _
                                Resultado & " - " & lerCampo(informacaoArquivo, "Mensagem")
                                
                    Else
                    
                        mensagemLOG2 grdLogSig, _
                                arq.DateCreated, _
                                lerCampo(informacaoArquivo, "Resultado"), _
                                nf.loja, _
                                nf.numero, _
                                nf.pedido, _
                                Resultado & " - " & lerCampo(informacaoArquivo, "Mensagem")
                    End If
                    

                    
                End If
                

                    Screen.MousePointer = 0
                    
                End If
            Arquivo = Dir
            
        Loop
    End If
    
    Exit Function
TrataErro:
    Select Case Err.Number
        Case 62, 13
            mensagemArquivoTXT.Close
            deletaArquivo GLB_EnderecoPastaRESP & Arquivo
        Case Else
            mensagemErroDesconhecido Err, "Verifica��o de pasta"
    End Select
End Function

Private Sub atualizaArquivo(ByRef enderecoArquivo As String, Arquivo As String, ByRef InformacaoTXT As String, ByRef Info As String)
    'open
    'deletaArquivo enderecoArquivo & arquivo
    Open enderecoArquivo & Arquivo For Output As #1
            
        Print #1, InformacaoTXT & _
        Info
        
    Close #1
    
    InformacaoTXT = InformacaoTXT & Info
    
End Sub

Public Function mensagemLOG2(grid, data As Date, tipoStatus As Integer, loja As String, numeroNotaFiscal As String, pedido As String, mensagem As String)

    Dim status As String
    Dim corLinha As ColorConstants
    Dim i As Byte
    
    
    Select Case tipoStatus
        Case 100, 4012, 4016, 124, 4014, 101, 9016, 9005, 4005, 4017
            status = "Sucesso"
        Case Else
            status = "Erro"
            corLinha = vbRed
    End Select
               
    If chkMostraLogErro.Value = 1 And chkMostraLogSucesso = 1 Then
        grid.AddItem loja & Chr(9) & data & Chr(9) & Format(pedido, "##") & Chr(9) & numeroNotaFiscal & Chr(9) & mensagem
    ElseIf chkMostraLogErro = 1 And status = "Erro" Then
        grid.AddItem loja & Chr(9) & data & Chr(9) & Format(pedido, "##") & Chr(9) & numeroNotaFiscal & Chr(9) & mensagem
    ElseIf chkMostraLogSucesso = 1 And status <> "Erro" Then
        grid.AddItem loja & Chr(9) & data & Chr(9) & Format(pedido, "##") & Chr(9) & numeroNotaFiscal & Chr(9) & mensagem
    End If
        
    
    If status = "Erro" Then
        pintaLinha grid, corLinha, (grid.Rows - 1)
    End If
    
    grid.TopRow = grid.Row
    
    grid.Row = 0
    grid.Col = 1
    'If grid.Name <> grdLog.Name Then grid.Sort = flexSortStringAscending
    
    grid.Sort = flexSortStringDescending
    grid.Refresh
                   
End Function

Public Sub pintaLinha(grid, Cor, Linha As Integer)
    grid.Row = Linha
    For i = 0 To grid.Cols - 1
        grid.Col = i
        grid.CellForeColor = Cor
    Next i
End Sub

Public Sub pintaFonteLinha(grid, Cor, Linha As Integer)
    grid.Row = Linha
    For i = 0 To grid.Cols - 1
        grid.Col = i
        grid.s = Cor
    Next i
End Sub




Public Sub montaComboLoja()
On Error GoTo TrataErro
    Dim ado_loja As New ADODB.Recordset
    Dim ado_loja2 As New ADODB.Recordset

    With ado_loja
        sql = "select cts_loja as lojas From ControleSistema"
        .CursorLocation = adUseClient
        .Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
        
            sql = "select LO_Loja AS LOJAS from loja where LO_CGC IN (select LO_CGC from loja where LO_Loja = '" & Trim(ado_loja("lojas")) & "') AND LO_Situacao = 'A' ORDER BY LO_Regiao"
            ado_loja2.CursorLocation = adUseClient
            ado_loja2.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            
            lojasWhere = "("
            
            Do While Not ado_loja2.EOF
            
                'comboLojas.AddItem Trim(ado_loja2("lojas"))
                lojasWhere = lojasWhere & "'" & Trim(ado_loja2("lojas")) & "'" & ","
                ado_loja2.MoveNext
                
            Loop
            
            lojasWhere = left(lojasWhere, (Len(lojasWhere) - 1)) & ")"
            ado_loja2.Close
            
        .Close
    End With

    Exit Sub
TrataErro:
    Select Case Err.Number
        Case Else
            mensagemErroDesconhecido Err, "Erro na leitura de lista de lojas"
    End Select
End Sub


Private Sub carregaGrdLogSig()
    qtdeLinhaAnterior = 0
    timerVerificaResposta_Timer
End Sub


Private Sub abrirTXT(nf As notaFiscal, tiponota As String)

    Dim enderecoArquivoTXT As String
     
    Screen.MousePointer = 11
    
    enderecoArquivoTXT = criaTXTtemporario(GLB_EnderecoPastaFIL, tiponota, nf.pedido, nf.CNPJ, nf.loja)
    If enderecoArquivoTXT <> "" Then
        ShellExecute Hwnd, "open", (enderecoArquivoTXT), "", "", 1
        Shell "explorer " & GLB_EnderecoPastaFIL, vbHide
    Else
        lblMSGNota.Caption = "N�o foi possivel abrir o TXT"
    End If
    
    Screen.MousePointer = 0
    
End Sub

Public Function montaTXT(pedido As String, loja As String) As String
    Dim ado_estrutura As New ADODB.Recordset

    sql = "select nfl_descricao, nfl_dados " & _
          "from NFE_NFLojas, nfcapa " & _
          "where nfl_nroNFE = nf " & _
          "and numeroped = '" & pedido & "' " & _
          "order by NFL_sequencia, nfl_NROnfe, nfl_dados desc"
    
    ado_estrutura.CursorLocation = adUseClient
    ado_estrutura.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
        
    Do While Not ado_estrutura.EOF
        If left(ado_estrutura("nfl_descricao"), 1) = "[" Or left(ado_estrutura("nfl_descricao"), 2) = "--" Then
            montaTXT = montaTXT & vbNewLine & vbNewLine & ado_estrutura("nfl_descricao")
        Else
            montaTXT = montaTXT & vbNewLine & ado_estrutura("nfl_descricao") & "= " & Trim(ado_estrutura("nfl_dados"))
        End If
        
        ado_estrutura.MoveNext
    Loop
        
    ado_estrutura.Close
End Function



Public Function criaTXTtemporario(endereco As String, tiponota As String, pedido As String, CNPJ As String, loja As String) As String

    Dim corpoMensagem As String
    Dim nota As notaFiscal
    
On Error GoTo TrataErro
    
    If tiponota = "NOTA" Then corpoMensagem = montaTXT(pedido, loja)
    If tiponota = "SAT" Then corpoMensagem = montaTXTSAT(pedido)
    
    If corpoMensagem <> Empty Then
        criaTXTtemporario = endereco & LCase(tiponota) & (Format(pedido, "000000000")) & "#" & CNPJ & ".txt"
        Open criaTXTtemporario For Output As #1
             Print #1, corpoMensagem
        Close #1
    End If
    
    Exit Function
    
TrataErro:
    Select Case Err.Number
    Case Else
        mensagemErroDesconhecido Err, "Erro na cria��o do arquivo"
    End Select
End Function

Private Function linhaSelecionaValida(ByRef grid) As Boolean
    linhaSelecionaValida = False
    If grid.Row >= grid.FixedRows And grid.Row <= grid.Rows And _
       grid.Col >= grid.FixedCols And grid.Col <= grid.Cols Then
       linhaSelecionaValida = True
    End If
End Function


Public Function obterCNPJloja() As String
On Error GoTo TrataErro
    Dim ado_loja As New ADODB.Recordset
    With ado_loja
        sql = "select top 1 lo_cgc as cnpj from loja where lo_loja in " & lojasWhere & " group by lo_cgc"
        .CursorLocation = adUseClient
        .Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
        
        If Not ado_loja.EOF Then obterCNPJloja = ado_loja("cnpj")
        
        .Close
    End With
    
    Exit Function
TrataErro:
    Select Case Err.Number
        'Case Else
            'mensagemErroDesconhecido Err, "Erro ao obter o CNPJ da loja"
    End Select
End Function



Function imprimirEspaco()




    
   ' Printer.Print Space(80)
  ''  Printer.Print Space(80)
 '   Printer.Print Space(80)
    
'    Printer.EndDoc

End Function

Private Sub mostraPopup(grid)
    With grid
        If (.MouseCol = 0 Or .MouseCol = 5) Then
            If .MouseRow >= .FixedRows And .MouseRow < .Rows Then
                .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
            End If
        ElseIf .MouseCol <> 0 Or .MouseCol <> 5 Then
            .ToolTipText = ""
        End If
    End With
End Sub


Private Sub numeroCopiaImpressao()

    Dim SQLLinhaImpressora As String
    Dim ado_rotulo As New ADODB.Recordset

    SQLLinhaImpressora = "INSERT INTO NFE_NFLOJAS " & vbNewLine & _
                         "SELECT TOP 1 * " & vbNewLine & _
                         "FROM NFE_NFLOJAS " & vbNewLine & _
                         "WHERE LTRIM(RTRIM(NFL_DESCRICAO)) = 'IMPRESSORA' " & vbNewLine & _
                         "AND NFL_NRONFE = '" & nf.numero & "'" & vbNewLine & _
                         "AND NFL_Loja = '" & nf.loja & "'"

    sql = "Select condpag as condpag " & vbNewLine & _
          "from nfcapa" & vbNewLine & _
          "where LojaOrigem = '" & nf.loja & "'" & vbNewLine & _
          "and nf = '" & nf.numero & "'" & vbNewLine & _
          "and serie = '" & nf.eSerie & "'"
    
    ado_rotulo.CursorLocation = adUseClient
    ado_rotulo.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
        
    If Val(ado_rotulo("condpag")) > 3 Then
        rdoCNLoja.Execute SQLLinhaImpressora
    End If
    
    For i = 2 To WQtdeCopiaNE
        rdoCNLoja.Execute SQLLinhaImpressora
    Next i

    ado_rotulo.Close
    
    'rdoCNLoja.Execute SQLLinhaImpressora
    
End Sub

Private Function obterNumeroItem(informacoes As String) As String
    
    informacoes = Replace(informacoes, "[nItem:", "((")
    informacoes = Replace(informacoes, "]", "))")
    
    If informacoes Like "*((*" Then
        Dim inicioCampo, fimCampo As Integer
    
        inicioCampo = (InStr(informacoes, "((")) + (Len("(("))
        fimCampo = (InStr(inicioCampo, informacoes, "))")) - inicioCampo
    
        If inicioCampo + fimCampo <> 0 Then
            obterNumeroItem = Mid$(informacoes, inicioCampo, fimCampo)
        End If
    Else
        obterNumeroItem = ""
    End If
    
End Function

