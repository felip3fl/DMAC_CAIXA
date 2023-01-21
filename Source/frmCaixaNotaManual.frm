VERSION 5.00
Object = "{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7u.ocx"
Begin VB.Form frmCaixaNotaManual 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Nota Manual"
   ClientHeight    =   8715
   ClientLeft      =   1935
   ClientTop       =   1290
   ClientWidth     =   14925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   14925
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraNFP 
      BackColor       =   &H80000007&
      Caption         =   "CGC/CPF"
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
      Left            =   7545
      TabIndex        =   13
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
         TabIndex        =   14
         Top             =   285
         Width           =   6885
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   6540
      Left            =   195
      TabIndex        =   3
      Top             =   165
      Width           =   7080
      Begin VB.PictureBox picCabGride 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   45
         ScaleHeight     =   480
         ScaleWidth      =   6960
         TabIndex        =   4
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
            TabIndex        =   6
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
            TabIndex        =   5
            Top             =   210
            Width           =   840
         End
      End
      Begin VSFlex7UCtl.VSFlexGrid grdItens 
         Height          =   5865
         Left            =   60
         TabIndex        =   7
         Top             =   630
         Width           =   6990
         _cx             =   12330
         _cy             =   10345
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
         FormatString    =   $"frmCaixaNotaManual.frx":0000
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
         BackColorFrozen =   -2147483633
         ForeColorFrozen =   4210752
         WallPaperAlignment=   4
      End
   End
   Begin VB.Frame fraProduto 
      BackColor       =   &H80000007&
      Caption         =   "Informações NF Manual"
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
      TabIndex        =   1
      Top             =   6720
      Width           =   7110
      Begin VB.TextBox txtNota 
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
         Left            =   4695
         TabIndex        =   0
         Top             =   315
         Width           =   2040
      End
      Begin VB.TextBox txtSerie 
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
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   315
         Width           =   525
      End
      Begin VB.Label lblNotaFiscal 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Nota Fiscal:"
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
         Left            =   2685
         TabIndex        =   17
         Top             =   435
         Width           =   1995
      End
      Begin VB.Label lblSerie 
         BackStyle       =   0  'Transparent
         Caption         =   "Série:"
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
         Left            =   435
         TabIndex        =   16
         Top             =   435
         Width           =   720
      End
   End
   Begin Balcao2010.chameleonButton cmdTotalVenda 
      Height          =   0
      Left            =   12870
      TabIndex        =   8
      Top             =   1050
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   0
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCaixaNotaManual.frx":0029
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
      Left            =   12240
      TabIndex        =   9
      Top             =   1050
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   0
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCaixaNotaManual.frx":0045
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton chameleonButton1 
      Height          =   0
      Left            =   225
      TabIndex        =   10
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   0
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCaixaNotaManual.frx":0061
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
      Left            =   225
      TabIndex        =   11
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   0
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCaixaNotaManual.frx":007D
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
      Left            =   9015
      TabIndex        =   18
      Top             =   1605
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
      Left            =   9150
      TabIndex        =   15
      Top             =   975
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
      Left            =   9225
      TabIndex        =   12
      Top             =   525
      Visible         =   0   'False
      Width           =   900
   End
End
Attribute VB_Name = "frmCaixaNotaManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wQuantidade As Integer
Dim sql As String
Dim wValorVenda As Double
Dim SomaTotalVenda As Double
Dim PrecoVenda As Double
Dim NroPedido As Long
Dim ContadorItens As Long
Dim L As Long
Dim Ind As Integer
Dim wLen As Integer
Dim NumInicial As Long
Dim NroNotaFiscal As Long
Dim Componente As String
Dim Sequencia As String
Dim LeftPagamento As Long
Dim template As Integer
Dim wValorDados As String
Dim wSequencia As Integer
Dim wCodigo As Integer
Dim wlinhaGrid As String
Dim contgrid As Integer
Dim wValorItem  As String * 10
Dim wValorTotalItem As String * 10

Dim wTipoQuantidade As String * 1
Dim wCasaDecimais As Integer
Dim wTipoDesconto As String * 1
Dim wDescricao As String * 29
Dim wAliquota As String * 5
Dim wPrecoVenda As String * 8
Dim wDesconto As String * 8

Dim wCodigoProduto As String * 13
Dim wQtde As String * 4
Dim wDescricao38 As String * 38
Dim wDescontoECF As String

Private Sub Label1_Click()

End Sub

Private Sub Form_Load()
 lblTotalvenda.Caption = ""
 lblTotalItens.Caption = ""
 fraNFP.top = fraProduto.top
 fraNFP.left = fraProduto.left
 Call GravaValorCarrinho(frmCaixaNotaManual, lblTotalItens.Caption)
wTotalVenda = 0

' webInternet1.Movie = "C:\sistemas\Trader Caixa 2010\Imagens\barrapequenavermelhacomCarrinho.swf"
' webInternet1.Play

Call AjustaTela(frmCaixaNotaManual)

If RsDados.State = 1 Then
  RsDados.Close
End If

 
sql = "Select * from ParametroCaixa where PAR_NroCaixa = " & GLB_Caixa

 RsDados.CursorLocation = adUseClient
 RsDados.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

If RsDados.EOF Then
   RsDados.Close
   MsgBox "Problema com os Parametros avise ao CPD", vbCritical, "Aviso"
   Unload Me
   Exit Sub
End If

'lblNroCaixa.Caption = GLB_Caixa

'lblloja.Caption = Trim(RsDados("PAR_Loja"))
wlblloja = Trim(RsDados("PAR_Loja"))

RsDados.Close

sql = "Select ControleCaixa.*,USU_Codigo,USU_Nome from ControleCaixa,UsuarioCaixa" _
            & " Where CTR_Operador = USU_Codigo and CTR_Supervisor <> '99' and CTR_SituacaoCaixa='A' and CTR_NumeroCaixa = " & GLB_Caixa
          
             RsDados.CursorLocation = adUseClient
             RsDados.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            
            If RsDados.EOF = False Then
               GLB_USU_Nome = RsDados("USU_Nome")
               GLB_USU_Codigo = RsDados("USU_Codigo")
               GLB_CTR_Protocolo = RsDados("CTR_Protocolo")
          
            End If
            RsDados.Close

            frmControlaCaixa.txtPedido.text = frmControlaCaixa.txtPedido.text

            Call CarregaGrid

            sql = "Select cliente from nfcapa where NumeroPed = " & frmControlaCaixa.txtPedido.text & " and tiponota = 'PA'"
            RsDados.CursorLocation = adUseClient
            RsDados.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            
            If Not RsDados.EOF Then
                If RsDados("cliente") = "999999" Then
                    txtSerie = "D1"
                Else
                    txtSerie = "S1"
                End If
            Else
                MsgBox "Nota não encontrada"
            End If

End Sub


Private Sub CarregaGrid()

        wTotalVenda = 0
        wtotalitens = 0
        grdItens.Rows = 0
        grdItens.Visible = True
        
        
        sql = "Select NFItens.*,PR_Descricao " _
             & "From NFItens,Produtoloja " _
             & "Where Referencia = pr_Referencia and NumeroPed = " _
             & frmControlaCaixa.txtPedido.text & " and tiponota = 'PA' order by Item"
               
               RsDados.CursorLocation = adUseClient
               RsDados.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            
        pedido = frmControlaCaixa.txtPedido
            
        sql = "Select * From Nfcapa Where NumeroPed = " _
             & frmControlaCaixa.txtPedido.text & " and tiponota = 'PA' "
               
                RsDadosCapa.CursorLocation = adUseClient
                RsDadosCapa.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
             
             
             If Not RsDadosCapa.EOF Then
                wPegaDesconto = RsDadosCapa("Desconto")
                wPegaFrete = RsDadosCapa("FreteCobr")
                
             End If
                
              RsDadosCapa.Close
             
             If Not RsDados.EOF Then
                  
                  Do While Not RsDados.EOF
                        wValorVenda = (RsDados("vlunit") * RsDados("Qtde"))
                        wValorItem = right(Trim(Format(RsDados("vlunit"), "###,##0.00")), 10)
                        wValorTotalItem = Format((RsDados("vlunit") * RsDados("Qtde")), "###,##0.00")
                               
                         grdItens.AddItem " " & left(Format(RsDados("Item"), "000") & Space(5), 5) _
                           & "         " & left(RsDados("Referencia") & Space(7), 7) _
                           & "           " & right(Space(6) & Format(RsDados("Qtde"), "000"), 6) _
                           & "                   " & "" & right(Space(11) & Format(RsDados("vlunit"), "###,##0.00"), 11) & "" _
                           & "                   " & right(Space(11) & Format(wValorTotalItem, "###,##0.00"), 11)
                                                                    
                                                                    
                        grdItens.AddItem " " & Trim(RsDados("pr_Descricao"))
                        
                        frmControlaCaixa.txtPedido.SelStart = 0
                        
                        
                        wtotalitens = (wtotalitens + 1)
                        wTotalVenda = (wTotalVenda + (Format((RsDados("vlunit") * RsDados("Qtde")), "###,##0.00")))
                                               
                        grdItens.TopRow = grdItens.Rows - 1
                        'grdItens.ZOrder
                        RsDados.MoveNext
                  Loop
                        lblTotalvenda.Caption = Format(wTotalVenda, "###,##0.00") - Format(wPegaDesconto, "###,##0.00")
                        lblTotalvenda.Caption = Format((lblTotalvenda.Caption + wPegaFrete), "###,##0.00")
                        lblTotalItens.Caption = Format(wtotalitens, "#,##0")
                        lblTotalGarantia.Caption = "+ G.E " & "0,00"
                        Call GravaValorCarrinho(frmCaixaNotaManual, lblTotalItens.Caption)

     
             Else
                        MsgBox "Pedido não Existe ou Nota Fiscal já foi emitida.", vbCritical, "Aviso"
                        frmControlaCaixa.txtPedido.SelStart = 0
                        frmControlaCaixa.txtPedido.SelLength = Len(frmControlaCaixa.txtPedido.text)
                         
             End If
    RsDados.Close



End Sub


Private Sub txtNota_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
      
     
      If IsNumeric(txtNota.text) = False Or txtNota.text = "" Or txtNota.text = "'" Then
           txtNota.SelStart = 0
            txtNota.SelLength = Len(txtNota.text)
            Exit Sub
      Else

      
         If txtSerie.text = "D1" Then
           frmControlaCaixa.txtPedido.text = frmControlaCaixa.txtPedido.text
           wValoraPagarNORMAL = Format(lblTotalvenda.Caption, "###,###,##0.00")
           frmFormaPagamento.txtIdentificadequeTelaqueveio.text = "FRMCAIXANOTAMANUAL"
           frmFormaPagamento.txtPedido.text = frmControlaCaixa.txtPedido
           frmFormaPagamento.txtTipoNota.text = "D1"
           frmFormaPagamento.Show vbModal
          Else
           fraNFP.Visible = True
          End If
      End If
      
End If




End Sub

Private Sub txtCGC_CPF_GotFocus()
txtCGC_CPF.SelStart = 0
txtCGC_CPF.SelLength = Len(txtCGC_CPF.text)
End Sub


Private Sub txtCGC_CPF_KeyPress(KeyAscii As Integer)
'   If KeyAscii = vbKeyEscape Then
'      lblTotalvenda.Caption = ""
'      lblTotalItens.Caption = ""
'      Call GravaValorCarrinho(frmCaixaTEFPedido, lblTotalItens.Caption)
'      Unload Me
'      Exit Sub
'   End If

   If KeyAscii = 27 Then
       txtCGC_CPF.text = ""
       fraNFP.Visible = False
       txtNota.SetFocus
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
           
          frmControlaCaixa.txtPedido.text = frmControlaCaixa.txtPedido.text
          wValoraPagarNORMAL = Format(lblTotalvenda.Caption, "###,###,##0.00")
          frmFormaPagamento.txtIdentificadequeTelaqueveio.text = "FRMCAIXANOTAMANUAL"
          frmFormaPagamento.txtPedido.text = frmControlaCaixa.txtPedido
          frmFormaPagamento.txtTipoNota.text = "S1"
          frmFormaPagamento.Show vbModal
     End If
End Sub

Private Sub txtNota_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
       lblTotalvenda.Caption = ""
       lblTotalItens.Caption = ""
        Call GravaValorCarrinho(frmCaixaNotaManual, lblTotalItens.Caption)
       Unload Me
   End If
End Sub

