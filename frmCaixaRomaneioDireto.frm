VERSION 5.00
Object = "{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7u.ocx"
Begin VB.Form frmCaixaRomaneioDireto 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Romaneio"
   ClientHeight    =   8235
   ClientLeft      =   2880
   ClientTop       =   1680
   ClientWidth     =   12360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
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
      Left            =   165
      TabIndex        =   9
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
         TabIndex        =   0
         Text            =   "1165454"
         Top             =   315
         Width           =   6930
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   6540
      Left            =   180
      TabIndex        =   1
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
         TabIndex        =   2
         Top             =   135
         Width           =   6990
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
            TabIndex        =   4
            Top             =   210
            Width           =   840
         End
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
            TabIndex        =   3
            Top             =   15
            Width           =   5880
         End
      End
      Begin VSFlex7UCtl.VSFlexGrid grdItens 
         Height          =   5280
         Left            =   60
         TabIndex        =   5
         Top             =   645
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
         FormatString    =   $"frmCaixaRomaneioDireto.frx":0000
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
         TabIndex        =   6
         Top             =   6150
         Width           =   11700
      End
   End
   Begin Balcao2010.chameleonButton cmdTotalVenda 
      Height          =   0
      Left            =   12855
      TabIndex        =   7
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
      MICON           =   "frmCaixaRomaneioDireto.frx":0029
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
      Left            =   12225
      TabIndex        =   8
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
      MICON           =   "frmCaixaRomaneioDireto.frx":0045
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
      Left            =   210
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
      MICON           =   "frmCaixaRomaneioDireto.frx":0061
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
      Left            =   210
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
      MICON           =   "frmCaixaRomaneioDireto.frx":007D
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
      Left            =   9420
      TabIndex        =   14
      Top             =   1485
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
      Left            =   9210
      TabIndex        =   13
      Top             =   525
      Visible         =   0   'False
      Width           =   900
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
      Left            =   9180
      TabIndex        =   12
      Top             =   1005
      Visible         =   0   'False
      Width           =   2010
   End
End
Attribute VB_Name = "frmCaixaRomaneioDireto"
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
Dim rsComplementoVenda As New ADODB.Recordset

Private Sub chbSair_Click()

Unload Me
End Sub

Private Sub chSair_Click()

End Sub

Private Sub Form_Load()


Call AjustaTela(frmCaixaRomaneioDireto)

 grdItens.BackColorBkg = &H80000006
 grdItens.ColWidth(0) = 6500
 lblTotalItens.Caption = ""
 lblTotalvenda.Caption = ""
Call GravaValorCarrinho(frmCaixaRomaneioDireto, lblTotalItens.Caption)
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
  fraProduto.Visible = True
  
 ' txtCodigoProduto.SetFocus
  GetAsyncKeyState (vbKeyTab)
  wItens = 0
  wNumeroCupom = 0

wlblloja = GLB_Loja

'Pedido = frmControlaCaixa.txtPedido.Text



End Sub


Private Sub grdItens_KeyPress(KeyAscii As Integer)
   
    If KeyAscii = 27 Then
       lblTotalvenda.Caption = ""
       lblTotalItens.Caption = ""
       Call GravaValorCarrinho(frmCaixaRomaneioDireto, lblTotalItens.Caption)
       Unload Me
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
       
          Sql = "Update nfcapa set qtditem = " & rdoContaItens("Numeroitem") & ",Cliente = 999999" _
                & " where nf = " & wNumeroCupom & " and serie = '00' and numeroped = " & NroPedido
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
         '' Call ImprimeCarimbo
          wValoraPagarNORMAL = Format(frmCaixaRomaneioDireto.lblTotalvenda.Caption, "###,###,##0.00")
          frmFormaPagamento.chbValoraPagar.Caption = Format(lblTotalvenda.Caption, "###,###,##0.00")
          frmFormaPagamento.txtIdentificadequeTelaqueveio.text = "frmCaixaRomaneioDireto"
          frmFormaPagamento.txtSerie.text = "00"
          frmFormaPagamento.txtPedido = NroPedido
          frmFormaPagamento.txtTipoNota.text = "RomaneioDireto"
     
          'frmGarantiaEstendida.Show 0
          frmFormaPagamento.Show vbModal

          
       End If
    End If
    
    If KeyCode = 27 Then
        Unload Me
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

                wtotalitens = (wtotalitens + 1)
                NroItens = NroItens + 1
                wTotalVenda = _
                 (wTotalVenda + Format((RsDadosTef("PR_PrecoVenda1") * wQuantidade), "###,##0.00"))
                 lblTotalvenda.Caption = Format(wTotalVenda, "###,##0.00")
                 lblTotalItens.Caption = Format(wtotalitens, "##0")
                 lblTotalGarantia.Caption = "+ G.E " & "0,00"
                 Call GravaValorCarrinho(frmCaixaRomaneioDireto, lblTotalItens.Caption)
                 Call PegaNumeroPedido
                               RsDadosTef.Close
 '             grdItens.Rows = grdItens.Rows - 2
  '            lblDescricaoProduto.Caption = ""
           End If
              

         
         End If
   
End Sub


Sub PegaNumeroPedido()
 Screen.MousePointer = 11
 If NroItens = 1 Then
 
    Sql = "Select * from Controlesistema "
    
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
    
    CriaCapaPedido NroPedido

    
 End If
  
 GravaItensPedido NroPedido, 11, 725
Screen.MousePointer = vbNormal
End Sub
Function CriaCapaPedido(ByVal NumeroPedido As Double)
 
    wLoja = PegaLojaControle
      
      Sql = ""
      Sql = "Select count(referencia) as NumeroItem from NfItens " _
          & "where NumeroPed=" & NroPedido & ""
          
          rdoContaItens.CursorLocation = adUseClient
          rdoContaItens.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    
'''    sql = "Insert into NfCapa (NUMEROPED,Serie, DATAEMI,LOJAORIGEM, TIPONOTA, Vendedor, DATAPED, HORA, " _
'''        & " VendedorLojaVenda, LojaVenda,TM,CodOper,CFOAUX,ECF,nf,Condpag,qtditem,situacaoprocesso,outraloja, dataprocesso, " _
'''        & "OUTROVEND,MODALIDADEVENDA,TIPOFRETE, FRETECOBR) " _
'''        & "Values (" & numeroPedido & ",'00' , '" & Format(Date, "yyyy/mm/dd") & "', " _
'''        & "'" & wLoja & "','PA',725, " _
'''        & "'" & Format(Date, "yyyy/mm/dd") & "', '" & Format(Time, "hh:mm:ss") & "', " _
'''        & "725, '" & wLoja & "',0,512,512," & GLB_ECF & "," & wNumeroCupom & ",01," _
'''        & rdoContaItens("Numeroitem") & ",'A','" & wLoja & "','" & Format(Date, "yyyy/mm/dd") & "','725','A Vista',1,0.00)"
'''        rdoCNLoja.Execute (sql)

    Sql = "Insert into NfCapa (NUMEROPED,Serie, DATAEMI,LOJAORIGEM, TIPONOTA, Vendedor, DATAPED, HORA, " _
        & " VendedorLojaVenda, LojaVenda,TM,CodOper,ECF,nf,Condpag,qtditem,situacaoprocesso,outraloja, dataprocesso, " _
        & "OUTROVEND,MODALIDADEVENDA,TIPOFRETE, FRETECOBR) " _
        & "Values (" & NumeroPedido & ",'00' , '" & Format(Date, "yyyy/mm/dd") & "', " _
        & "'" & wLoja & "','PA',725, " _
        & "'" & Format(Date, "yyyy/mm/dd") & "', '" & Format(Time, "hh:mm:ss") & "', " _
        & "725, '" & wLoja & "',0,512," & GLB_ECF & "," & wNumeroCupom & ",01," _
        & rdoContaItens("Numeroitem") & ",'A','" & wLoja & "','" & Format(Date, "yyyy/mm/dd") & "','725','A Vista',1,0.00)"
        rdoCNLoja.Execute (Sql)

     
     
     rdoContaItens.Close
     
End Function
Function GravaItensPedido(ByVal NumeroPedido As Double, ByVal TipoMovimentacao As Double, ByVal Vendedor As Integer)

    wLoja = PegaLojaControle
          
     Sql = "Insert into NfItens (nf, NUMEROPED,Serie, DATAEMI, REFERENCIA, QTDE, VLUNIT, " _
        & "VLTOTITEM, ICMS, DESCONTO, PLISTA,  " _
        & "LOJAORIGEM,  TIPONOTA,  Item, situacaoprocesso, dataprocesso, baseicms,ICMSAplicado,Cest) " _
        & "Values (" & wNumeroCupom & "," & NroPedido & ",'00', '" & Format(Date, "yyyy/mm/dd") & "', '" _
        & wCodigoProduto & "', " & wQtde & ", " _
        & "" & ConverteVirgula(Format(wItemPrecoVenda, "0.00")) & ", " _
        & ConverteVirgula(Format(wVlTotItem, "0.00")) & ", " & ConverteVirgula(Format(wICMS, "0.00")) & ",0, " _
        & "  " & ConverteVirgula(Format(wPLISTA, "0.00")) & ",  " _
        & wLoja & ",'PA'," & NroItens & ",'A','" & Format(Date, "yyyy/mm/dd") & "', 0.00,0,'')"
        rdoCNLoja.Execute (Sql)

End Function












