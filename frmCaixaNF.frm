VERSION 5.00
Object = "{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7u.ocx"
Begin VB.Form frmCaixaNF 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Caixa NF"
   ClientHeight    =   8010
   ClientLeft      =   1185
   ClientTop       =   1830
   ClientWidth     =   8115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
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
      Left            =   2250
      TabIndex        =   10
      Top             =   6720
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
         Left            =   90
         MaxLength       =   29
         TabIndex        =   0
         Top             =   300
         Width           =   6915
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
      Left            =   165
      TabIndex        =   9
      Top             =   6720
      Width           =   1920
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
         TabIndex        =   1
         Top             =   270
         Width           =   1680
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   6480
      Left            =   200
      TabIndex        =   2
      Top             =   100
      Width           =   7080
      Begin VB.PictureBox picCabGride 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   30
         ScaleHeight     =   480
         ScaleWidth      =   6990
         TabIndex        =   3
         Top             =   135
         Width           =   7020
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
         Begin VB.Label lblCabecGride 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   " Itens         Código             Qtde.                Preço Unitário        Total do Itens"
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
            TabIndex        =   4
            Top             =   15
            Width           =   5835
         End
      End
      Begin VSFlex7UCtl.VSFlexGrid grdItens 
         Height          =   5805
         Left            =   45
         TabIndex        =   6
         Top             =   630
         Width           =   6990
         _cx             =   12330
         _cy             =   10239
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
         FormatString    =   $"frmCaixaNF.frx":0000
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
   Begin VB.Label lblCliente 
      BackColor       =   &H00AE7411&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   840
      Left            =   8940
      TabIndex        =   12
      Top             =   4545
      Width           =   5025
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
      Left            =   12960
      TabIndex        =   11
      Top             =   1500
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
      Left            =   13335
      TabIndex        =   8
      Top             =   360
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
      Left            =   13335
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   2010
   End
End
Attribute VB_Name = "frmCaixaNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim wQuantidade As Integer
Dim Sql As String
Dim wValorVenda As Double
Dim SomaTotalVenda As Double
Dim PrecoVenda As Double
Dim NroPedido As Long
Dim ContadorItens As Long
Dim L As Long
Dim Ind As Integer
Dim wLen As Integer
Dim NumInicial As Long

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

Dim ValorDesconto As Double
Dim SubTotal As Double
Dim TotalVenda As Double

Private Sub cmdFechar_Click()

End Sub

Private Sub cmdRomaneio_Click()

 txtPedido.SetFocus
  
End Sub

Private Sub cmdProtocolo_Click()

End Sub



Private Sub Form_Load()

 grdItens.BackColorBkg = &H80000006
 grdItens.ColWidth(0) = 6500
 lblTotalvenda.Caption = ""
 lblTotalItens.Caption = ""
 
 lblCliente.top = fraPedido.top + 140
 lblCliente.left = fraNFP.left + 170
 
 'Call GravaValorCarrinho(frmCaixaNF, lblTotalItens.Caption)
 'wTotalVenda = 0
 
 'webInternet1.Movie = "C:\sistemas\Trader Caixa 2010\Imagens\barrapequenavermelhacomCarrinho.swf"
 'webInternet1.Play

Call AjustaTela(frmCaixaNF)




If RsDados.State = 1 Then
  RsDados.Close
End If

 
Sql = "Select * from ParametroCaixa where PAR_NroCaixa = " & GLB_Caixa

 RsDados.CursorLocation = adUseClient
 RsDados.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

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

Sql = "Select ControleCaixa.*,USU_Codigo,USU_Nome from ControleCaixa,UsuarioCaixa" _
            & " Where CTR_Operador = USU_Codigo and CTR_SituacaoCaixa='A' AND CTR_Supervisor <> '99' and CTR_NumeroCaixa = " & GLB_Caixa
          
             RsDados.CursorLocation = adUseClient
             RsDados.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            
            If RsDados.EOF = False Then
               GLB_USU_Nome = RsDados("USU_Nome")
               GLB_USU_Codigo = RsDados("USU_Codigo")
               GLB_CTR_Protocolo = RsDados("CTR_Protocolo")
5
            End If
RsDados.Close

txtPedido.text = frmControlaCaixa.txtPedido.text

Call CarregaGrid
Call ClienteConsumidor

End Sub



Private Sub Label1_Click()

End Sub

Private Sub lblNroCaixa_Click()

End Sub

Public Sub Form_Unload(Cancel As Integer)
    exibirMensagemPadraoTEF
End Sub

Private Sub grdItens_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
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
        Call GravaValorCarrinho(frmCaixaNF, lblTotalItens.Caption)
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
        
        Sql = "Update nfcapa set cpfnfp =  " & txtCGC_CPF & " where NumeroPed = " _
              & txtPedido.text & " and tiponota = 'PA'"
        rdoCNLoja.Execute Sql
        
      End If
      txtPedido.Enabled = True
      txtCGC_CPF.text = ""
      fraNFP.Visible = False
      txtPedido.SetFocus
      txtPedido.SetFocus
           
    End If
End Sub

Private Sub txtPedido_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Then
    Unload Me
End If
    
If KeyCode = vbKeyF2 Then
    txtPedido.text = frmControlaCaixa.txtPedido.text
    If txtPedido.text = "" Then
       MsgBox "Informe o Numero do Pedido referente a NF", vbCritical, "Atenção"
       txtPedido.SelStart = 0
       txtPedido.SelLength = Len(txtPedido.text)
       Exit Sub
    End If

     Sql = "Select * From Nfitens " _
             & "Where numeroped  = " & txtPedido.text & "" _
             & " order by Item"

             RsDados.CursorLocation = adUseClient
             RsDados.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

             If RsDados.EOF = False Then
                 wNotaFiscalReemissao = RsDados("NF")
                 wSerieReemissao = Trim(RsDados("Serie"))

                 If wNotaFiscalReemissao <> 0 Then
                    If MsgBox("Deseja Reemitir a Nota Fiscal Nº " & wNotaFiscalReemissao, vbYesNo + vbQuestion, "Atenção") = vbYes Then
                       RsDados.Close
                       EmiteNotafiscal wNotaFiscalReemissao, wSerieReemissao
                    Exit Sub
                    End If
                 Else
                    MsgBox "Nota fiscal não encontrada Nº " & wNotaFiscalReemissao, vbCritical, "Atenção"
                    RsDados.Close
                    Exit Sub
                 End If
             Else
             MsgBox "Nota fiscal não encontrada Nº " & wNotaFiscalReemissao, vbCritical, "Atenção"

             RsDados.Close
             Exit Sub
             End If

             RsDados.Close
 End If

If KeyCode = vbKeyF1 Then

       'carregaProdutoGarantia txtPedido.text
       
       If aceitaGarantia = False Then
       
            If lblTotalvenda.Caption = "" Then
                Exit Sub
            Else
       
            txtPedido.text = frmControlaCaixa.txtPedido.text
            wValoraPagarNORMAL = Format(lblTotalvenda.Caption, "###,###,##0.00")
            frmFormaPagamento.txtIdentificadequeTelaqueveio.text = "FRMCAIXANF"
            frmFormaPagamento.txtPedido.text = frmCaixaNF.txtPedido
            frmFormaPagamento.txtTipoNota.text = "NF"
            frmFormaPagamento.Show vbModal
            
            End If
       Else
            'Unload Me
            Form_Load
            'txtPedido_KeyPress 13
            'frmControlaCaixa.txtPedido.Text = txtPedido.Text
            'frmControlaCaixa.txtPedido_KeyPress 13
       End If
       
       'frmFormaPagamento.ZOrder 1
       
      
       'frmFormaPagamento.ZOrder

    
End If

End Sub





Private Sub CarregaGrid()

    Dim wParcelas As Byte

    wTotalVenda = 0
    wtotalitens = 0
    wtotalGarantia = 0
    grdItens.Rows = 0
    grdItens.Visible = True
    
    Sql = "Select NFItens.*,PR_Descricao " _
    & "From NFItens,Produtoloja " _
    & "Where Referencia = pr_Referencia and NumeroPed = " _
    & txtPedido.text & " and tiponota = 'PA' order by Item"
    
    RsDados.CursorLocation = adUseClient
    RsDados.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    pedido = txtPedido
    
    Sql = "Select * From Nfcapa Where NumeroPed = " _
        & txtPedido.text & " and tiponota = 'PA' "
    
    RsDadosCapa.CursorLocation = adUseClient
    RsDadosCapa.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    
    If Not RsDadosCapa.EOF Then
        wPegaDesconto = RsDadosCapa("Desconto")
        wPegaFrete = RsDadosCapa("FreteCobr")
        wtotalGarantia = RsDadosCapa("TotalGarantia")
    End If
    
    wParcelas = RsDadosCapa("Parcelas")
    
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
            
            txtPedido.SelStart = 0
            
            
            wtotalitens = (wtotalitens + 1)
            wTotalVenda = (wTotalVenda + (Format((RsDados("vlunit") * RsDados("Qtde")), "###,##0.00")))
                        
            grdItens.TopRow = grdItens.Rows - 1
            'grdItens.ZOrder
            RsDados.MoveNext
        Loop
        
        lblTotalvenda.Caption = Format(wTotalVenda, "###,##0.00") - Format(wPegaDesconto, "###,##0.00")
        lblTotalvenda.Caption = Format((lblTotalvenda.Caption + wPegaFrete), "###,##0.00")
        lblTotalItens.Caption = Format(wtotalitens, "#,##0")
        lblTotalGarantia.Caption = "+ G.E " & Format(wtotalGarantia, "###,##0.00")
        Call GravaValorCarrinho(frmCaixaNF, lblTotalItens.Caption)
        
        exibirMensagemPedidoTEF txtPedido.text, wParcelas
        
    
    Else
    
        MsgBox "Pedido não Existe ou Nota Fiscal já foi emitida.", vbCritical, "Aviso"
        txtPedido.SelStart = 0
        txtPedido.SelLength = Len(txtPedido.text)
        
    End If
    
    
    
    RsDados.Close
    


End Sub




Public Sub txtPedido_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
   lblTotalvenda.Caption = ""
   lblTotalItens.Caption = ""
   Call GravaValorCarrinho(frmCaixaNF, frmCaixaNF.lblTotalItens.Caption)
   Unload Me
 End If
 
 If KeyAscii = 13 Then
   txtPedido.text = frmControlaCaixa.txtPedido.text
 End If
 
End Sub

Private Sub carregaCliente(codigo As String, NomeVendedor As String, CodigoVendedor As String)
    
    Dim rsCliente As New ADODB.Recordset
    Dim Cliente As String
    
    Sql = "Select ce_razao " & vbNewLine & _
    "from fin_cliente " & vbNewLine & _
    "where ce_codigocliente = '" & codigo & "'"
                   
    rsCliente.CursorLocation = adUseClient
    rsCliente.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                   
    If Not rsCliente.EOF Then
        Cliente = rsCliente("ce_razao")
        If Len(Cliente) > 26 Then
            Cliente = Mid(Cliente, 1, 26)
            Cliente = Cliente & "..."
        End If
        lblCliente.Caption = "CLIENTE:  " & Cliente
        lblCliente.Caption = lblCliente.Caption & vbNewLine
        lblCliente.Caption = lblCliente.Caption & vbNewLine
        lblCliente.Caption = lblCliente.Caption & "VENDEDOR:  " & CodigoVendedor & " - " & NomeVendedor
    Else
        MsgBox "Erro ao carregar informações sobre o Cliente", vbCritical, "Erro"
    End If
    
End Sub

Private Sub ClienteConsumidor()
        
        
        Sql = "Select top 1 cliente,ve_nome as nomeVendedor,ve_codigo as condigoVendedor from nfcapa,vende where NumeroPed = " _
             & txtPedido.text & " and tiponota = 'PA' and vendedor = ve_codigo"
               
               RsDadosCapa.CursorLocation = adUseClient
               RsDadosCapa.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
               
               If Not RsDadosCapa.EOF Then
                   If RsDadosCapa("cliente") = "999999" Then
                          txtCGC_CPF.Enabled = True
                          txtPedido.Enabled = False
                          fraNFP.Visible = True
                          fraNFP.left = fraPedido.left
                          fraNFP.top = fraPedido.top
                   Else
                          carregaCliente RsDadosCapa("cliente"), RsDadosCapa("nomeVendedor"), RsDadosCapa("condigoVendedor")
                          fraNFP.Visible = False
                   End If
                End If
                RsDadosCapa.Close

End Sub

