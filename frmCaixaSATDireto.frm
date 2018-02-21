VERSION 5.00
Object = "{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7u.ocx"
Begin VB.Form frmCaixaSATDireto 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Caixa SAT"
   ClientHeight    =   8250
   ClientLeft      =   2040
   ClientTop       =   1755
   ClientWidth     =   15120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCEP 
      BackColor       =   &H80000007&
      Caption         =   "CEP Cliente"
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
      Height          =   945
      Left            =   8310
      TabIndex        =   17
      Top             =   8985
      Visible         =   0   'False
      Width           =   7110
      Begin VB.TextBox txtCEP 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   105
         MaxLength       =   14
         TabIndex        =   18
         Top             =   285
         Width           =   6885
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   6540
      Left            =   180
      TabIndex        =   8
      Top             =   195
      Width           =   7080
      Begin VB.PictureBox picCabGride 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   45
         ScaleHeight     =   480
         ScaleWidth      =   6960
         TabIndex        =   9
         Top             =   135
         Width           =   6990
         Begin VB.Label lblCabecGride 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "  Itens         Código             Qtde.                Preço Unitário        Total do Itens"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   -15
            TabIndex        =   11
            Top             =   15
            Width           =   5880
         End
         Begin VB.Label lblCabecGride2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   30
            TabIndex        =   10
            Top             =   210
            Width           =   840
         End
      End
      Begin VSFlex7UCtl.VSFlexGrid grdItens 
         Height          =   5280
         Left            =   45
         TabIndex        =   12
         Top             =   600
         Width           =   6990
         _cx             =   12330
         _cy             =   9313
         _ConvInfo       =   1
         Appearance      =   2
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   12632256
         ForeColor       =   0
         BackColorFixed  =   4210752
         ForeColorFixed  =   16777215
         BackColorSel    =   3421236
         ForeColorSel    =   16777215
         BackColorBkg    =   12632256
         BackColorAlternate=   8421504
         GridColor       =   14737632
         GridColorFixed  =   8421504
         TreeColor       =   16777215
         FloodColor      =   16777215
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmCaixaSATDireto.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
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
         BackColorFrozen =   8421504
         ForeColorFrozen =   0
         WallPaperAlignment=   4
      End
      Begin VB.Label lblDescricaoProduto 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "ABCDEFGHIJKLMNOPQRSTUVXYZW12345678912"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   90
         TabIndex        =   13
         Top             =   6150
         Width           =   11700
      End
   End
   Begin Balcao2010.chameleonButton cmdTotalVenda 
      Height          =   0
      Left            =   12645
      TabIndex        =   7
      Top             =   1050
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   0
      BTYPE           =   11
      TX              =   "999"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCaixaSATDireto.frx":0029
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton cmdItens 
      Height          =   0
      Left            =   12015
      TabIndex        =   4
      Top             =   1050
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   0
      BTYPE           =   11
      TX              =   "999"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCaixaSATDireto.frx":0045
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraNFP 
      BackColor       =   &H80000007&
      Caption         =   "CNPJ / CPF"
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
      Height          =   945
      Left            =   7635
      TabIndex        =   3
      Top             =   6720
      Visible         =   0   'False
      Width           =   7110
      Begin VB.TextBox txtCGC_CPF 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   105
         MaxLength       =   14
         TabIndex        =   0
         Top             =   285
         Width           =   6885
      End
   End
   Begin VB.Frame fraProduto 
      BackColor       =   &H80000007&
      Caption         =   "Código do Produto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   180
      TabIndex        =   2
      Top             =   6720
      Visible         =   0   'False
      Width           =   7110
      Begin VB.TextBox txtCodigoProduto 
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
         Left            =   90
         TabIndex        =   1
         Top             =   315
         Width           =   6930
      End
   End
   Begin Balcao2010.chameleonButton chameleonButton1 
      Height          =   0
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   0
      BTYPE           =   11
      TX              =   "999"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCaixaSATDireto.frx":0061
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton chameleonButton2 
      Height          =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   0
      BTYPE           =   11
      TX              =   "999"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCaixaSATDireto.frx":007D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblTotalGarantia 
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   9630
      TabIndex        =   16
      Top             =   1725
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Label lblTotalvenda 
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   8970
      TabIndex        =   15
      Top             =   1005
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Label lblTotalItens 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   9000
      TabIndex        =   14
      Top             =   525
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      Index           =   2
      X1              =   -210
      X2              =   -210
      Y1              =   0
      Y2              =   195
   End
End
Attribute VB_Name = "frmCaixaSATDireto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Asteristico As Boolean
Dim AbreTela As Long
Dim wControlacor As Long
Dim wConfigCor As Long
Dim RefCod As Long
Dim wContCodigoZero As Long
Dim wQuantidade As Integer
Dim wReferencia As String
Dim wEspaco As String * 40
Dim Cor As String
Dim Cor1 As String
Dim Cor2 As String
Dim Cor3 As String
Dim Sql As String
Dim SomaTotalVenda As Double
Dim PrecoVenda As Double
Dim NroPedido As Long
Dim NumeroECF As String
Dim ContadorItens As Long
Dim wConfgCor As Long
Dim L As Long
Dim Ind As Integer
Dim wLen As Integer
Dim NumInicial As Long
Dim NroNotaFiscal As Long
Dim OriginalcmdFotowidth As Long
Dim OriginalcmdFotoHeight As Long
Dim OriginalcmdFotoTop As Long
Dim OriginalcmdFotoleft As Long
Dim OriginalgrdItensHeight As Long
Dim OriginalgrdItenstop As Long
Dim OriginalgrdItensleft As Long
Dim OriginalgrdItenswidth As Long
Dim OriginalcmdLogoMarcaHeight As Long
Dim OriginalcmdLogoMarcaTop As Long
Dim OriginalcmdLogoMarcaLeft As Long
Dim OriginalcmdLogoMarcawidth As Long
Dim OriginallblTotalitensHeight As Long
Dim OriginallblTotalitenstop As Long
Dim OriginallblTotalitensleft As Long
Dim OriginallblTotalvendaHeight As Long
Dim OriginallblTotalvendatop As Long
Dim OriginallblTotalvendaleft As Long
Dim OriginalcmdCabec01Height As Long
Dim OriginalcmdCabec01Left As Long
Dim OriginalcmdCabec01Top As Long
Dim OriginalcmdCabec01Width As Long
Dim Componente As String
Dim Sequencia As String
Dim LeftPagamento As Long
Dim template As Integer
Dim wValorDados As String
Dim wSequencia As Integer
Dim wCodigo As Integer
Dim wlinhaGrid As String
'Dim wItens As Integer
Dim contgrid As Integer
Dim wTipoQuantidade As String * 1
Dim wCasaDecimais As Integer
Dim wTipoDesconto As String * 1
Dim wDescricao As String * 29
Dim wAliquota As String * 5
Dim wAliqIPI As Double
Dim wPrecoVenda As String * 8
Dim wPrecoVenda2 As String * 8
Dim wItemPrecoVenda As Double
'Dim wItemPrecoVenda2 As Double
Dim wDesconto As String * 8
Dim wCodigoProduto As String * 13
Dim wQtde As String * 4
Dim wDescricao38 As String * 38

Private Sub chbSair_Click()
Unload Me
End Sub

Private Sub chSair_Click()

End Sub

Private Sub Form_Load()

Call AjustaTela(frmCaixaSATDireto)

 grdItens.BackColorBkg = &H80000006
 grdItens.ColWidth(0) = 6500
 lblTotalItens.Caption = ""
 lblTotalvenda.Caption = ""
 txtCodigoProduto.text = ""
Call GravaValorCarrinho(frmCaixaSATDireto, lblTotalItens.Caption)
 lblDescricaoProduto.Caption = ""
 txtCodigoProduto.text = ""
 wTipoQuantidade = "I"
 wCasaDecimais = 2
 wTipoDesconto = "$"
 wDesconto = 0
  NroItens = 0
  wTotalVenda = 0
  wtotalitens = 0
  grdItens.BackColorBkg = &H0&
  
  fraNFP.top = fraProduto.top
  fraNFP.left = fraProduto.left
  fraNFP.Visible = True
  
  fraCEP.top = fraProduto.top
  fraCEP.left = fraProduto.left
  fraCEP.Visible = False
  
  GetAsyncKeyState (vbKeyTab)
  wItens = 0
  wNumeroCupom = 0

wlblloja = GLB_Loja

pedido = frmControlaCaixa.txtPedido.text



End Sub



Private Sub grdMovimentoCaixa_Click()

End Sub

Private Sub lblNroCaixa_Click()
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    exibirMensagemPadraoTEF
End Sub

Private Sub grdItens_KeyPress(KeyAscii As Integer)
   
    If KeyAscii = 27 Then
       lblTotalvenda.Caption = ""
       lblTotalItens.Caption = ""
       Call GravaValorCarrinho(frmCaixaSATDireto, lblTotalItens.Caption)
       Unload Me
   End If
End Sub

Private Sub txtCEP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtCEP.text = "" Then
            MsgBox "Informe o CEP do cliente", vbExclamation, "CEP Cliente"
        ElseIf Len(txtCEP.text) <> 8 Then
            MsgBox "CEP Inválido", vbExclamation, "CEP Cliente"
        ElseIf IsNumeric(txtCEP.text) = False Then
            MsgBox "Informe apenas números", vbExclamation, "CEP Cliente"
        ElseIf txtCEP.text = "11111111" Or _
        txtCEP.text = "11111111" Or _
        txtCEP.text = "99999999" Or _
        txtCEP.text = "88888888" Or _
        txtCEP.text = "77777777" Or _
        txtCEP.text = "66666666" Or _
        txtCEP.text = "55555555" Or _
        txtCEP.text = "44444444" Or _
        txtCEP.text = "33333333" Or _
        txtCEP.text = "22222222" Or _
        txtCEP.text = "00000000" Or _
        txtCEP.text = "12345678" Then
            MsgBox "CEP Inválido!", vbExclamation, "CEP Cliente"
        Else
            fraCEP.Visible = False
            fraProduto.Visible = True
            txtCodigoProduto.SetFocus
            Call GravaNumeroCupomCgcCpf
        End If
    ElseIf KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub txtCGC_CPF_Change()
 
If IsNumeric(txtCGC_CPF.text) = False Then
   txtCGC_CPF.text = ""
   txtCGC_CPF.SelStart = 0
   txtCGC_CPF.SelLength = Len(txtCGC_CPF.text)
   
End If
End Sub

Private Sub txtCGC_CPF_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   txtCGC_CPF.text = ""
End If
End Sub

Private Sub txtCGC_CPF_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then
        lblTotalvenda.Caption = ""
       lblTotalItens.Caption = ""
       Call GravaValorCarrinho(frmCaixaSATDireto, lblTotalItens.Caption)
       Unload Me
   End If
   
    If KeyAscii = 27 Then
       lblTotalvenda.Caption = ""
       lblTotalItens.Caption = ""
       Call GravaValorCarrinho(frmCaixaSATDireto, lblTotalItens.Caption)
       Unload Me
   End If
   
   If KeyAscii = vbKeyReturn Then
      If Trim(txtCGC_CPF.text) <> "" Then
        If Len(txtCGC_CPF.text) = 11 Then
           If FU_ValidaCPF(txtCGC_CPF.text) = False Then
              MsgBox "CPF INVALIDO", vbCritical, "Atenção"
              txtCGC_CPF.SetFocus
              txtCGC_CPF.SelStart = 0
              txtCGC_CPF.SelLength = Len(txtCGC_CPF.text)
              Exit Sub
           End If
        End If
 
 
        If Len(txtCGC_CPF.text) = 14 Then
           If FU_ValidaCGC(txtCGC_CPF.text) = False Then
              MsgBox "CNPJ INVALIDO", vbCritical, "Atenção"
              txtCGC_CPF.SetFocus
              txtCGC_CPF.SelStart = 0
              txtCGC_CPF.SelLength = Len(txtCGC_CPF.text)
           Exit Sub
           End If
        Else
           If Len(txtCGC_CPF.text) <> 11 Then
              MsgBox "CNPJ/CPF INVALIDO", vbCritical, "Atenção"
              txtCGC_CPF.SetFocus
              txtCGC_CPF.SelStart = 0
              txtCGC_CPF.SelLength = Len(txtCGC_CPF.text)
           Exit Sub
           End If
        End If
      End If
      
      PegaNumeroPedido

      Call exibirMensagemPedidoTEF(Str(pedido), 1)

      '***** ROTINA ECF (NAO APAGAR)
      'wCupomAberto = False
      'RotinadeAberturadoCupom
      'If wCupomAberto = False And wNumeroCupom <> 0 Then
         'fraNFP.Visible = False
      'Else
        'Exit Sub
      'End If
     
     
         fraNFP.Visible = False
         
         'CEP Cliente consumidor '''''''''''''''''''''''''''''''''''''''''''
         
         'fraCEP.Visible = True
         'txtCEP.text = ""
         'txtCEP.SetFocus

         
         Call GravaNumeroCupomCgcCpf
         
            fraCEP.Visible = False
            fraProduto.Visible = True
            txtCodigoProduto.SetFocus
            
         
         
         ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         
         
           
  '***** ROTINA ECF (NAO APAGAR)
'     Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
     
   
     



   End If
End Sub




Private Sub txtCGC_CPF_LostFocus()

If GetAsyncKeyState(vbKeyTab) <> 0 Then
   txtCGC_CPF.Enabled = True
   
   txtCGC_CPF.SetFocus
End If
End Sub

Private Sub txtCodigoProduto_GotFocus()
 txtCodigoProduto.SelStart = 0
 txtCodigoProduto.SelLength = Len(txtCodigoProduto.text)

End Sub

Private Sub txtCodigoProduto_KeyDown(KeyCode As Integer, Shift As Integer)
    
   If KeyCode = vbKeyF1 Then
      
       wDesconto = 0
   
        Sql = ""
        Sql = "Select count(referencia) as NumeroItem from NfItens " _
          & "where NumeroPed=" & NroPedido & ""
          
          rdoContaItens.CursorLocation = adUseClient
          rdoContaItens.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
             
          
          rdoCNLoja.BeginTrans
       
          Sql = "Update nfcapa set qtditem = " & rdoContaItens("Numeroitem") & ", CPFNFP = '" & txtCGC_CPF & _
                "', cliente = 999999, CEPCLI = '" & txtCEP.text & "'" _
                & " where nf = " & wNumeroCupom & " and serie = '" & GLB_SerieCF & "' and numeroped = " & NroPedido
          rdoCNLoja.Execute Sql, rdExecDirect
    
          rdoCNLoja.CommitTrans
          
           'josi
          '************************ Gravando Valores NFCapa
       Sql = ""
       Sql = "Exec SP_Totaliza_Capa_Nota_Fiscal_Loja " & NroPedido
       rdoCNLoja.Execute Sql
       
                  
       Sql = ""
       Sql = "Select count(referencia) as NumeroItem from NFItens " _
           & "where NumeroPed=" & NroPedido & ""
          
            rsComplementoVenda.CursorLocation = adUseClient
            rsComplementoVenda.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
       
       Sql = ""
       Sql = "Update NFCapa set TipoNota = 'PA', qtditem = " & rsComplementoVenda("NumeroItem") & "" _
             & " Where NumeroPed = " & NroPedido
       rdoCNLoja.Execute Sql
       
       rsComplementoVenda.Close
       
'************************ Gravando TipoNota NFItens
       
       Sql = "Update NFItens Set TipoNota = 'PA' Where NumeroPed = " & NroPedido
       
       rdoCNLoja.Execute Sql
       
      ' Call LimpaForm

   
    ' rsItensVenda.Close
     Screen.MousePointer = vbNormal
     
          rdoContaItens.Close
          
      'fim josi
         
       If lblTotalvenda.Caption = "" Then
          Exit Sub
       Else

          wValoraPagarNORMAL = Format(frmCaixaSATDireto.lblTotalvenda.Caption, "###,###,##0.00")
          frmFormaPagamento.chbValoraPagar.Caption = Format(lblTotalvenda.Caption, "###,###,##0.00")
          frmFormaPagamento.txtIdentificadequeTelaqueveio.text = "frmCaixaSATDireto"
          wtxtCGC_CPF = txtCGC_CPF
          frmFormaPagamento.txtSerie.text = GLB_SerieCF
          frmFormaPagamento.txtPedido = NroPedido
          frmFormaPagamento.txtTipoNota.text = "SAT"
          frmFormaPagamento.Show vbModal
 '        frmFormaPagamento.ZOrder
          
       End If
    End If
    

End Sub


Private Sub txtCodigoProduto_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
            
        
        Ind = 0
        wReferencia = ""
        wQuantidade = 1
         
     
     
     
     For Ind = 1 To Len(txtCodigoProduto.text)
            
            If Mid(txtCodigoProduto.text, Ind, 1) = "*" Then
               If Ind = 1 Then
                  lblDescricaoProduto.Caption = wEspaco & "Código ou Quantidade Invalido"
                  txtCodigoProduto.SelStart = 0
                  txtCodigoProduto.SelLength = Len(txtCodigoProduto.text)
                  txtCodigoProduto.SetFocus
                  Exit Sub
               Else
                  If Asteristico = False Then
                     wReferencia = Mid(txtCodigoProduto.text, 1, (Ind - 1))
                     wQuantidade = Mid(txtCodigoProduto.text, (Ind + 1), 7)
                     Asteristico = True
                     Exit For
                  End If
               End If
            Else
                wReferencia = txtCodigoProduto.text
            End If

        Next Ind
        Asteristico = False
        Sql = ""
          
        Sql = "Select PR_precovenda1,PR_icmPdv,pr_substituicaotributaria," _
             & "PR_Referencia,PR_descricao,* From Produtoloja ,ProdutoBarras " _
             & "Where PRB_Referencia = PR_Referencia and PRB_CodigoBarras='" & wReferencia & "'"
            
                 
             RsDadosTef.CursorLocation = adUseClient
             RsDadosTef.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
         
         If RsDadosTef.EOF Then
              lblDescricaoProduto.Caption = "Código não Cadastrado"
              txtCodigoProduto.SelStart = 0
              txtCodigoProduto.SelLength = Len(txtCodigoProduto.text)
              txtCodigoProduto.SetFocus
              RsDadosTef.Close
        
         ElseIf RsDadosTef("PR_PrecoVenda1") + wTotalVenda > 9999 Then
             MsgBox "Você não pode vende mais de R$10.000,00 no Cupom Fiscal", vbCritical, "Valor limite SAT"
             RsDadosTef.Close
         Else
                  
            
                  
             wReferencia = RsDadosTef("PR_referencia")

         
              lblDescricaoProduto.Caption = RsDadosTef("PR_Descricao")
                wItens = wItens + 1
             
                                                  
                grdItens.AddItem " " & left(Format(wItens, "000") & Space(5), 5) _
                           & "         " & left(wReferencia & Space(7), 7) _
                           & "           " & right(Space(6) & Format(wQuantidade, "000"), 6) _
                           & "                   " & "" & right(Space(11) & Format(RsDadosTef("PR_PrecoVenda1"), "###,##0.00"), 11) & "" _
                           & "                   " & right(Space(11) & Format((RsDadosTef("PR_PrecoVenda1") * wQuantidade), "###,##0.00"), 11)
                                                                                                                
                                                                    
              grdItens.AddItem RsDadosTef("PR_Descricao")
                                                                                                          
              txtCodigoProduto.SelStart = 0
              txtCodigoProduto.SelLength = Len(txtCodigoProduto.text)
              txtCodigoProduto.SetFocus

              wCodigoProduto = wReferencia
              wDescricao = RsDadosTef("PR_Descricao")
              wDescricao38 = RsDadosTef("PR_Descricao")
              wQtde = Format(wQuantidade, "000")
              wPrecoVenda = Format(RsDadosTef("PR_PrecoVenda1"), "###,###,##0.00")
              wItemPrecoVenda = Format(RsDadosTef("PR_PrecoVenda1"), "###,###,##0.00")
              wPLISTA = Format(RsDadosTef("PR_PrecoVenda1"), "###,###,##0.00")
              'wItemPrecoVenda2 = Format(RsDadosTef("PR_PrecoVenda1"), "###,###,##0.00")
              wICMS = Format(RsDadosTef("PR_IcmsSaida"), "###,###,##0.00")
              wVlTotItem = wPrecoVenda * wQuantidade
              
              If RsDadosTef("pr_substituicaotributaria") = "S" Then
                wAliquota = "FF"
              Else
                wAliquota = Replace(Format(RsDadosTef("PR_icmpdv"), "00.00"), ",", "")
                
                If Trim(wAliquota) = "0000" Then
                       wAliquota = "FF"
                ElseIf Trim(wAliquota) <> "0560" And Trim(wAliquota) <> "0700" And Trim(wAliquota) <> "0880" And _
                       Trim(wAliquota) <> "1200" And Trim(wAliquota) <> "1800" And Trim(wAliquota) <> "2500" Then
                       wAliquota = "1200"
                End If
                
              End If
              
              grdItens.TopRow = grdItens.Rows - 1

              
              'Retorno = Bematech_FI_AumentaDescricaoItem(wDescricao38)

'              Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
                  
 '             Retorno = Bematech_FI_VendeItem(wCodigoProduto, wDescricao, _
                  Trim(wAliquota), wTipoQuantidade, wQtde, wCasaDecimais, _
                  (wPrecoVenda * 100), wTipoDesconto, 0)
            
  '            Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
             
             
             
             'If Retorno = 1 Then
                wtotalitens = (wtotalitens + 1)
                NroItens = NroItens + 1
                wTotalVenda = _
                 (wTotalVenda + Format((RsDadosTef("PR_PrecoVenda1") * wQuantidade), "###,##0.00"))
                 lblTotalvenda.Caption = Format(wTotalVenda, "###,##0.00")
                 lblTotalItens.Caption = Format(wtotalitens, "##0")
                 lblTotalGarantia.Caption = "+ G.E " & "0,00"
                 Call GravaValorCarrinho(frmCaixaSATDireto, lblTotalItens.Caption)
                 '    Call PegaNumeroPedido
                 If NroItens = 1 Then
                      CriaCapaPedido NroPedido
                 End If
                 GravaItensPedido NroPedido, 11, 725
              'Else
              'grdItens.Rows = grdItens.Rows - 2
              'lblDescricaoProduto.Caption = ""
              'End If
              RsDadosTef.Close
         
         End If


     End If

   
   If KeyAscii = 27 Then
     If MsgBox("Deseja cancelar essa venda?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        'Retorno = Bematech_FI_CancelaCupom()
       'Função que analisa o retorno da impressora
       ' Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
       lblTotalvenda.Caption = ""
       lblTotalItens.Caption = ""
       Call GravaValorCarrinho(frmCaixaSATDireto, lblTotalItens.Caption)
       Unload Me
     End If
   End If
   
End Sub

Sub PegaNumeroPedido()
 Screen.MousePointer = 11
' If NroItens = 1 Then
 
    Sql = "Select CTs_NumeroPedido from Controlesistema "
    
    rdocontrole.CursorLocation = adUseClient
    rdocontrole.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    NroPedido = rdocontrole("CTs_NumeroPedido")
    frmControlaCaixa.txtPedido.text = rdocontrole("CTs_NumeroPedido")
    pedido = rdocontrole("CTs_NumeroPedido")
    rdocontrole.Close
    rdoCNLoja.BeginTrans
       
    Sql = "Update Controlesistema set CTs_NumeroPedido = " & pedido & " + 1"
    rdoCNLoja.Execute Sql, rdExecDirect
  
    rdoCNLoja.CommitTrans
    
'    CriaCapaPedido NroPedido

    
    'End If
  
' GravaItensPedido NroPedido, 11, 725
    


Screen.MousePointer = vbNormal
End Sub

Function CriaCapaPedido(ByVal numeroPedido As Double)
 
    wLoja = PegaLojaControle
      
      Sql = ""
      Sql = "Select count(referencia) as NumeroItem from NfItens " _
          & "where NumeroPed=" & numeroPedido & ""
          
          rdoContaItens.CursorLocation = adUseClient
          rdoContaItens.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    
    Sql = ""
'''    sql = "Insert into NfCapa (NUMEROPED,Serie, DATAEMI, " _
'''        & "LOJAORIGEM, TIPONOTA, Vendedor, DATAPED, HORA, " _
'''        & " VendedorLojaVenda, LojaVenda,TM,CodOper,CFOAUX,ECF,nf,Condpag,CPFNFP, qtditem,situacaoprocesso,outraloja, dataprocesso, " _
'''        & "OUTROVEND,MODALIDADEVENDA,TIPOFRETE, FRETECOBR,CLIENTE) " _
'''        & "Values (" & numeroPedido & ",'CF' , '" & Format(Date, "yyyy/mm/dd") & "', " _
'''        & "'" & wLoja & "','PA',725, " _
'''        & "'" & Format(Date, "yyyy/mm/dd") & "', '" & Format(Time, "hh:mm:ss") & "', " _
'''        & "725, '" & wLoja & "',0,5012,5012," & GLB_ECF & "," & wNumeroCupom & ",01,'" & txtCGC_CPF.Text & "'," _
'''        & rdoContaItens("Numeroitem") & ",'A','" & wLoja & "','" & Format(Date, "yyyy/mm/dd") & "','725','A Vista',1, 0.00,'999999')"
'''        rdoCNLoja.Execute (sql)
     
      Sql = "Insert into NfCapa (NUMEROPED,Serie, DATAEMI, " _
        & "LOJAORIGEM, TIPONOTA, Vendedor, DATAPED, HORA, " _
        & " VendedorLojaVenda, LojaVenda,TM,CodOper,ECF,nf,Condpag,CPFNFP, qtditem,situacaoprocesso,outraloja, dataprocesso, " _
        & "OUTROVEND,MODALIDADEVENDA,TIPOFRETE, FRETECOBR,CLIENTE) " _
        & "Values (" & numeroPedido & ",'" & GLB_SerieCF & "' , '" & Format(Date, "yyyy/mm/dd") & "', " _
        & "'" & wLoja & "','PA',725, " _
        & "'" & Format(Date, "yyyy/mm/dd") & "', '" & Format(Time, "hh:mm:ss") & "', " _
        & "725, '" & wLoja & "',0,5012," & GLB_ECF & "," & wNumeroCupom & ",01,'" & txtCGC_CPF.text & "'," _
        & rdoContaItens("Numeroitem") & ",'A','" & wLoja & "','" & Format(Date, "yyyy/mm/dd") & "','725','A Vista',1, 0.00,'999999')"
        rdoCNLoja.Execute (Sql)
        
     
     rdoContaItens.Close
     
End Function
Function GravaItensPedido(ByVal numeroPedido As Double, ByVal TipoMovimentacao As Double, ByVal Vendedor As Integer)

    wLoja = PegaLojaControle
          
     Sql = "Insert into NfItens (nf, NUMEROPED,Serie, DATAEMI, REFERENCIA, QTDE, VLUNIT, " _
        & "VLTOTITEM, ICMS, DESCONTO, PLISTA,  " _
        & "LOJAORIGEM,  TIPONOTA,  Item, situacaoprocesso, dataprocesso, baseicms,ICMSAplicado,cest) " _
        & "Values (" & wNumeroCupom & "," & NroPedido & ",'" & GLB_SerieCF & "', '" & Format(Date, "yyyy/mm/dd") & "', '" _
        & wCodigoProduto & "', " & wQtde & ", " _
        & "" & ConverteVirgula(Format(wItemPrecoVenda, "0.00")) & ", " _
        & ConverteVirgula(Format(wVlTotItem, "0.00")) & ", " & ConverteVirgula(Format(wICMS, "0.00")) & ",0, " _
        & "  " & ConverteVirgula(Format(wPLISTA, "0.00")) & ",  " _
        & wLoja & ",'PA'," & NroItens & ",'A','" & Format(Date, "yyyy/mm/dd") & "', 0.00,0,'')"
        rdoCNLoja.Execute (Sql)

'        & ConverteVirgula(Format(wItemPrecoVenda2, "0.00")) & ", "

End Function

'***** ROTINA ECF (NAO APAGAR)
Private Sub RotinadeAberturadoCupom()

      Retorno = 0

      Retorno = Bematech_FI_AbrePortaSerial()

      Call VerificaRetornoImpressora("", "", "BemaFI32")
      
      Retorno = Bematech_FI_AbreCupom(txtCGC_CPF.text)

      If Retorno <> 1 Then
           MsgBox "Verifique se impressora de Cupom Fiscal está Ligada e conectada ao computador!", vbCritical, "Ateção"
           Retorno = Bematech_FI_FechaPortaSerial()

           Exit Sub
        End If

      wValorRetorno = ""
      Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
      Retorno = Bematech_FI_NumeroCupom(wNumeroCupom)

      If Trim(wValorRetorno) = "6, 2, 1" Then

         CancelaCupomFiscal
         wCupomAberto = True
         Exit Sub
         
      End If

End Sub
Private Sub GravaNumeroCupomCgcCpf()

    'If validaNumeroCupom = True Then
    
        Screen.MousePointer = vbHourglass
        rdoCNLoja.BeginTrans
        Sql = "Update nfitens set serie = '" & GLB_SerieCF & "' Where NumeroPed = " & NroPedido
              rdoCNLoja.Execute Sql, rdExecDirect
        Screen.MousePointer = vbNormal
        rdoCNLoja.CommitTrans
        rdoCNLoja.BeginTrans
        Screen.MousePointer = vbHourglass
        
        Sql = "Update nfcapa set serie = '" & GLB_SerieCF & "', ECF = " & GLB_ECF & " , CPFNFP = '" & txtCGC_CPF & "'" _
              & " Where NumeroPed = " & NroPedido
              rdoCNLoja.Execute Sql, rdExecDirect
        Screen.MousePointer = vbNormal
        rdoCNLoja.CommitTrans
        
        Screen.MousePointer = vbNormal
    'Else
        'MsgBox "Falha na gravação do número de Cupom Fiscal." & vbNewLine & "Verifique o Status da Impressora", vbCritical, "Erro Cupom Fiscal"
        'Retorno = Bematech_FI_CancelaCupom()
       ''Função que analisa o retorno da impressora
'        Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
        'lblTotalvenda.Caption = ""
        'lblTotalItens.Caption = ""
        'Call GravaValorCarrinho(frmCaixaSATDireto, lblTotalItens.Caption)
        'Unload Me
    'End If
    
End Sub


