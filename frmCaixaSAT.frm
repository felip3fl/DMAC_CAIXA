VERSION 5.00
Object = "{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7u.ocx"
Begin VB.Form frmCaixaSAT 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmCaixaSAT"
   ClientHeight    =   8745
   ClientLeft      =   1845
   ClientTop       =   1800
   ClientWidth     =   12345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10575
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   6510
      Left            =   195
      TabIndex        =   5
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
         TabIndex        =   6
         Top             =   150
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
            Left            =   -30
            TabIndex        =   8
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
            TabIndex        =   7
            Top             =   210
            Width           =   840
         End
      End
      Begin VSFlex7UCtl.VSFlexGrid grdItens 
         Height          =   5790
         Left            =   45
         TabIndex        =   4
         Top             =   675
         Width           =   6990
         _cx             =   12330
         _cy             =   10213
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
         FloodColor      =   -1
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
         FormatString    =   $"frmCaixaSAT.frx":0000
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
   Begin VB.Frame fraPedido 
      BackColor       =   &H00000000&
      Caption         =   "Pedido"
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
      Height          =   885
      Left            =   180
      TabIndex        =   3
      Top             =   6720
      Width           =   1890
      Begin VB.TextBox txtPedido 
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
         MaxLength       =   6
         TabIndex        =   0
         Top             =   270
         Width           =   1680
      End
   End
   Begin VB.Frame fraNFP 
      BackColor       =   &H00000000&
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
      Height          =   885
      Left            =   2190
      TabIndex        =   2
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
         Left            =   75
         MaxLength       =   29
         TabIndex        =   1
         Top             =   270
         Width           =   6915
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SAT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   165
      TabIndex        =   12
      Top             =   -60
      Visible         =   0   'False
      Width           =   1635
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
      Left            =   8970
      TabIndex        =   11
      Top             =   2700
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
      Left            =   14055
      TabIndex        =   10
      Top             =   3780
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
      Left            =   14085
      TabIndex        =   9
      Top             =   3300
      Visible         =   0   'False
      Width           =   900
   End
End
Attribute VB_Name = "frmCaixaSAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ATUAL

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
'Dim wDescontoECF As String
Dim wDescontoECF As Double

Private Sub chameleonButton1_Click()
lblTotalvenda.Caption = ""
 lblTotalItens.Caption = ""
  Call GravaValorCarrinho(frmCaixaSAT, lblTotalItens.Caption)
Unload Me
End Sub

Private Sub Form_Load()
 grdItens.BackColorBkg = &H80000006
 grdItens.ColWidth(0) = 6500


 Call AjustaTela(Me)

 lblTotalvenda.Caption = ""
 lblTotalItens.Caption = ""
  Call GravaValorCarrinho(frmCaixaSAT, lblTotalItens.Caption)
 wTipoQuantidade = "I"
 wCasaDecimais = 2
 wTipoDesconto = "$"
 wDesconto = 0

  wTotalVenda = 0
  wtotalitens = 0
  LeftPagamento = 9
  grdItens.BackColorBkg = &H0&
    
 fraNFP.top = fraPedido.top
 fraNFP.left = fraPedido.left
 
 grdItens.BackColorBkg = &H80000006
 grdItens.ColWidth(0) = 6500
 
 wTipoQuantidade = "I"
 wCasaDecimais = 2
 wTipoDesconto = "$"
 wDesconto = 0
  wTotalVenda = 0
  wtotalitens = 0
  grdItens.BackColorBkg = &H0&
 
  GetAsyncKeyState (vbKeyTab)

wlblloja = GLB_Loja
wNumeroCupom = 0

txtPedido.text = frmControlaCaixa.txtPedido.text
Call CarregaGrid

End Sub

'ROTINA ECF(NAO APAGAR)
Private Sub VerificaSeExisteCupomAberto()
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
         ' Quando por algum motivo o cupom ficau aberto na execução anterior,
         ' a rotina abaixo cancelará o cupom que ficou aberto para poder emitir o proximo cumpom.
         MsgBox "Cupom que se encontrava aberto na Impressora fiscal"
         CancelaCupomFiscal

         Exit Sub
      End If

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Retorno = Bematech_FI_FechaPortaSerial()
End Sub

Private Sub lblDisplayTotal_Click()
End Sub

Private Sub grdItens_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF1 Then
      If fraNFP.Visible = False And fraPedido.Visible = False Then
         frmFormaPagamento.chbValoraPagar.Caption = Format(lblTotalvenda.Caption, "###,###,##0.00")
         wValoraPagarNORMAL = Format(lblTotalvenda.Caption, "###,###,##0.00")
         frmFormaPagamento.txtSerie = GLB_SerieCF
         frmFormaPagamento.txtPedido = txtPedido
         frmFormaPagamento.txtTipoNota.text = "SAT"
         frmFormaPagamento.txtIdentificadequeTelaqueveio.text = "frmCaixaSAT"
         frmFormaPagamento.Show vbModal
      End If
   End If
    
   
End Sub

Private Sub grdItens_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
     If MsgBox("Deseja cancelar essa venda?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        'Retorno = Bematech_FI_CancelaCupom()
       'Função que analisa o retorno da impressora
        'Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
       lblTotalvenda.Caption = ""
       lblTotalItens.Caption = ""
       Call GravaValorCarrinho(frmCaixaSAT, lblTotalItens.Caption)
       Unload Me
     End If
   End If
End Sub

Private Sub txtCGC_CPF_GotFocus()
txtCGC_CPF.SelStart = 0
txtCGC_CPF.SelLength = Len(txtPedido.text)
End Sub


Private Sub txtCGC_CPF_KeyPress(KeyAscii As Integer)

   If KeyAscii = 27 Then
       lblTotalvenda.Caption = ""
       lblTotalItens.Caption = ""
        Call GravaValorCarrinho(frmCaixaSAT, lblTotalItens.Caption)
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
      
      
      
      
       '***** ROTINA ECF (NAO APAGAR)
      'wCupomAberto = False
      'RotinadeAberturadoCupom
      'If wCupomAberto = False And Int(wNumeroCupom) <> 0 Then
         fraNFP.Visible = False
      'Else
        'Exit Sub
      'End If
     
  '***** ROTINA ECF (NAO APAGAR)
     'Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
      
      If GravaNumeroCupomCgcCpf = False Then Exit Sub
    

      sql = "Select nfitens.Referencia,nfitens.QTDE,nfitens.VLUnit,PR_Descricao,PR_icmpdv,pr_substituicaotributaria " _
          & "From nfitens,Produtoloja  " _
          & "Where PR_referencia = Referencia and NumeroPed = " _
          & txtPedido.text & " and Tiponota = 'PA' order by Item"
           RsDadosTef.CursorLocation = adUseClient
           RsDadosTef.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
         
          If Not RsDadosTef.EOF Then
             Do While Not RsDadosTef.EOF
                wCodigoProduto = RsDadosTef("Referencia")
                wDescricao = RsDadosTef("PR_Descricao")
                wQtde = Format(RsDadosTef("QTDE"), "000")
                wPrecoVenda = Format(RsDadosTef("VLUnit"), "###,###,##0.00")
               
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
              
              wTotalVenda = _
                (wTotalVenda + Format((wPrecoVenda * wQuantidade), "###,##0.00"))
              
'              grdItens.TopRow = grdItens.Rows - 1

               '***** ROTINA ECF (NAO APAGAR)
               
                'Retorno = Bematech_FI_AumentaDescricaoItem(wDescricao)
               'Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
                'Retorno = Bematech_FI_VendeItem(wCodigoProduto, wDescricao, _
                  Trim(wAliquota), wTipoQuantidade, wQtde, wCasaDecimais, _
                  (wPrecoVenda * 100), wTipoDesconto, (RsDadosTef("Desconto") * 100))
                'Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
                
             RsDadosTef.MoveNext
            Loop
            lblTotalvenda.Caption = Format(wTotalVenda, "###,##0.00") - Format(wPegaDesconto, "###,##0.00")
            lblTotalvenda.Caption = Format((lblTotalvenda.Caption + wPegaFrete), "###,##0.00")
            lblTotalItens.Caption = Format(wtotalitens, "#,##0")
            lblTotalGarantia.Caption = "+ G.E " & "0,00"
            Call GravaValorCarrinho(frmCaixaSAT, lblTotalItens.Caption)
          Else
              MsgBox "Pedido Não Encontrado", vbCritical, "Aviso"
              Exit Sub
          End If
          RsDadosTef.Close
          
         'wDescontoECF = wPegaDesconto * 100
'         wDescontoECF = Right("00000000000000" & wDescontoECF, 14)
        
        '***** ROTINA ECF (NAO APAGAR)
        'Retorno = Bematech_FI_IniciaFechamentoCupom("D", "$", wDescontoECF)
        'Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
        
      frmFormaPagamento.chbValoraPagar.Caption = Format(lblTotalvenda.Caption, "###,###,##0.00")
      wValoraPagarNORMAL = Format(lblTotalvenda.Caption, "###,###,##0.00")
      frmFormaPagamento.txtSerie = GLB_SerieCF
      frmFormaPagamento.txtPedido = txtPedido
      frmFormaPagamento.txtTipoNota.text = "SAT"
      frmFormaPagamento.txtIdentificadequeTelaqueveio.text = "frmCaixaSAT"
      frmFormaPagamento.Show vbModal

   End If
End Sub

          

          




Private Sub txtPedido_GotFocus()
 txtPedido.SelStart = 0
 txtPedido.SelLength = Len(txtPedido.text)
 grdItens.BackColorBkg = &H80000006
 grdItens.ColWidth(0) = 6500
 lblTotalvenda.Caption = ""
 lblTotalItens.Caption = ""
  Call GravaValorCarrinho(frmCaixaSAT, lblTotalItens.Caption)
 wTipoQuantidade = "I"
 wCasaDecimais = 2
 wTipoDesconto = "$"
 wDesconto = 0
 wTotalVenda = 0
 wtotalitens = 0
 grdItens.BackColorBkg = &H0&
End Sub

Private Sub txtPedido_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
   
        If MsgBox("Deseja cancelar essa venda?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        Retorno = Bematech_FI_CancelaCupom()
       'Função que analisa o retorno da impressora
        Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
   
       lblTotalvenda.Caption = ""
       lblTotalItens.Caption = ""
        Call GravaValorCarrinho(frmCaixaSAT, lblTotalItens.Caption)
       Unload Me
    End If
   

   End If
   If KeyAscii = 13 Then
   txtPedido.text = frmControlaCaixa.txtPedido.text
 End If

End Sub
Private Function GravaNumeroCupomCgcCpf() As Boolean

    GravaNumeroCupomCgcCpf = False
    
    'If validaNumeroCupom = True Then
        Screen.MousePointer = vbHourglass
        rdoCNLoja.BeginTrans
        sql = "Update nfitens set nf = " & "0" & ",Serie = '" & GLB_SerieCF & "' Where NumeroPed = " & txtPedido.text
              rdoCNLoja.Execute sql, rdExecDirect
        Screen.MousePointer = vbNormal
        rdoCNLoja.CommitTrans
        rdoCNLoja.BeginTrans
        Screen.MousePointer = vbHourglass
        
        sql = "Update nfcapa set nf = " & "0" & ",Serie = '" & GLB_SerieCF & "', ECF = " & GLB_ECF & " , CPFNFP = '" & txtCGC_CPF & "'" _
              & " Where NumeroPed = " & txtPedido.text
              rdoCNLoja.Execute sql, rdExecDirect
        Screen.MousePointer = vbNormal
        rdoCNLoja.CommitTrans
        Screen.MousePointer = vbNormal
        GravaNumeroCupomCgcCpf = True
    'Else
        'MsgBox "Falha na gravação do número de Cupom Fiscal." & vbNewLine & "Verifique o Status da Impressora", vbCritical, "Erro Cupom Fiscal"
        'Retorno = Bematech_FI_CancelaCupom()
       'Função que analisa o retorno da impressora
        'Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
        'lblTotalvenda.Caption = ""
        'lblTotalItens.Caption = ""
        'Call GravaValorCarrinho(frmCaixaTEF, lblTotalItens.Caption)
        'Unload Me
    'End If
    
    
End Function

Private Sub CarregaGrid()

        
        wTotalVenda = 0
        wtotalitens = 0
        grdItens.Rows = 0
        grdItens.Visible = True
        
        pedido = txtPedido
            
        sql = "Select Desconto,Fretecobr From Nfcapa Where NumeroPed = " _
             & txtPedido.text & " and tiponota = 'PA' "
               
        RsDadosCapa.CursorLocation = adUseClient
        RsDadosCapa.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
          
             
        If Not RsDadosCapa.EOF Then
           wPegaDesconto = RsDadosCapa("Desconto")
           wPegaFrete = RsDadosCapa("Fretecobr")
                
        End If
           
        RsDadosCapa.Close
            
       
        sql = "Select nfitens.VLUnit,nfitens.QTDE,nfitens.Item,nfitens.Referencia,PR_Descricao " _
             & "From nfitens,Produtoloja " _
             & "Where referencia = PR_Referencia and NumeroPed = " _
             & txtPedido.text & " and Tiponota = 'PA' order by Item"
            RsDadosTef.CursorLocation = adUseClient
            RsDadosTef.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
           
             pedido = txtPedido
             If Not RsDadosTef.EOF Then
               
                  fraNFP.Visible = True
'                  txtCGC_CPF.SetFocus
                  fraPedido.Visible = False
                  Do While Not RsDadosTef.EOF
                        wValorVenda = (RsDadosTef("VLUnit") * RsDadosTef("QTDE"))
                        wValorItem = right(Trim(Format(RsDadosTef("VLUnit"), "###,##0.00")), 10)
                        wValorTotalItem = Format((RsDadosTef("VLUnit") * RsDadosTef("QTDE")), "###,##0.00")
 
                                                         
                         grdItens.AddItem " " & left(Format(RsDadosTef("Item"), "000") & Space(5), 5) _
                           & "         " & left(RsDadosTef("Referencia") & Space(7), 7) _
                           & "           " & right(Space(6) & Format(RsDadosTef("Qtde"), "000"), 6) _
                           & "                   " & "" & right(Space(11) & Format(RsDadosTef("vlunit"), "###,##0.00"), 11) & "" _
                           & "                   " & right(Space(11) & Format(wValorTotalItem, "###,##0.00"), 11)
                                            
                        grdItens.AddItem " " & RsDadosTef("PR_Descricao")
                        
                        txtPedido.SelStart = 0
                        
                        
                        
                        wtotalitens = (wtotalitens + 1)
                        wTotalVenda = (wTotalVenda + (Format((RsDadosTef("VLUnit") * RsDadosTef("QTDE")), "###,##0.00")))
                        lblTotalvenda.Caption = Format(wTotalVenda, "###,##0.00")
                        lblTotalItens.Caption = Format(wtotalitens, "#,##0")
                        lblTotalGarantia.Caption = "+ G.E " & "0,00"
                        grdItens.TopRow = grdItens.Rows - 1
                        grdItens.ZOrder
                        RsDadosTef.MoveNext
                  Loop
                        lblTotalvenda.Caption = Format(wTotalVenda, "###,##0.00") - Format(wPegaDesconto, "###,##0.00")
                        lblTotalvenda.Caption = Format((lblTotalvenda.Caption + wPegaFrete), "###,##0.00")
                        lblTotalItens.Caption = Format(wtotalitens, "#,##0")
                        lblTotalGarantia.Caption = "+ G.E " & "0,00"
                        Call GravaValorCarrinho(frmCaixaSAT, lblTotalItens.Caption)
             Else
                        MsgBox "Pedido Não Encontrado", vbCritical, "Aviso"
             End If
    RsDadosTef.Close

End Sub


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


