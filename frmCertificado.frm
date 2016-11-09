VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmCertificado 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7545
   ClientLeft      =   540
   ClientTop       =   2115
   ClientWidth     =   15120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox imgOcultaNavegador 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   165
      ScaleHeight     =   915
      ScaleWidth      =   1215
      TabIndex        =   3
      Top             =   210
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer timerImprimir 
      Left            =   2895
      Top             =   6825
   End
   Begin VB.Frame fraPedido 
      BackColor       =   &H00000000&
      Caption         =   "Pedido G.E."
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
      Left            =   300
      TabIndex        =   1
      Top             =   6360
      Width           =   1890
      Begin VB.TextBox txtNumeroPedido 
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
         TabIndex        =   2
         Top             =   285
         Width           =   1680
      End
   End
   Begin SHDocVwCtl.WebBrowser webNavegador 
      Height          =   5985
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   14400
      ExtentX         =   25400
      ExtentY         =   10557
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin Balcao2010.chameleonButton cmdImprimirTodos 
      Height          =   585
      Left            =   12810
      TabIndex        =   4
      Top             =   6420
      Visible         =   0   'False
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   1032
      BTYPE           =   11
      TX              =   "Imprimir"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   14737632
      FCOLO           =   14737632
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCertificado.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblStatusImpressao 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Status Impressao"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   300
      TabIndex        =   5
      Top             =   6495
      Visible         =   0   'False
      Width           =   14490
   End
End
Attribute VB_Name = "frmCertificado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''' TODOS OS CAMPOS CERTIFICADO ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim pedido As String

Dim apoliceNumero As String
Dim codigoEstipulante As String
Dim codigoInternoProduto As String
Dim estipulante As String

Dim numeroDoCertificado As String * 18
Dim telefoneParaContato As String

Dim objetivoSeguro As String

Dim segurado As String
Dim CPF As String
Dim outroDoc As String
Dim tipoDoc As String
Dim Endereco As String
Dim Numero As String
Dim Complemento As String
Dim Bairro As String
Dim Cidade As String
Dim uf As String
Dim CEP As String
Dim Telefone As String
Dim email As String

Dim dataContratacaoSeguro As Date
Dim dataCompraBem As Date
Dim perildoGarantiaOriginalInicio As Date
Dim perildoGarantiaOriginalFim As Date
Dim perildoCoberturaSeguroInicio As Date
Dim perildoCoberturaSeguroFim As Date

Dim bemSegurado As String
Dim Marca As String
Dim Modelo As String
Dim valorBemSegurado As String

Dim PremioLiquido As Double
Dim iof As Double
Dim premioTotal As Double

Dim percentualRemuneracao As Double
Dim valorRemuneracao As Double

Dim cidadeAtual As String
Dim Data As String

''''' OUTRAS VARIAVEIS '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim referencia As String '''''''''''''''''''''''
Dim Item As String
Dim quantidadeGarantia As Integer

Dim itemImpressao As Integer
Dim numeroDeCopia As Integer
Dim NumeroItem As Integer
Dim pocentagemItemImpressao As Double

Dim itemExibido As Integer
Dim quantidadeTotalItens As Integer

Dim rsItensGarantia As New ADODB.Recordset
Dim rsProdutoGarantia As New ADODB.Recordset

Dim mensagemImpressao As String
Dim statusImpressao As String

Private Sub Form_Load()

    mensagemImpressao = "         Imprimindo Garantia Estendida"
    statusImpressao = ". . . . ."
    
    limpaTodasVariaveis
    ocultaNavegador True
    Call AjustaTela(Me)
    verificaNumeroPedido
    
End Sub

Public Function statusFuncionamento(campo As String) As String
    If campo = "  . . . ." Then
        statusFuncionamento = ".   . . ."
    ElseIf campo = ".   . . ." Then
        statusFuncionamento = ". .   . ."
    ElseIf campo = ". .   . ." Then
        statusFuncionamento = ". . .   ."
    ElseIf campo = ". . .   ." Then
        statusFuncionamento = ". . . .  "
    ElseIf campo = ". . . .  " Then
        statusFuncionamento = "  . . . ."
    Else
        statusFuncionamento = "  . . . ."
    End If
End Function

Private Sub Form_Activate()
    If frmFormaPagamento.txtPedido.Text <> "" Then
        fraPedido.Visible = False
    End If
End Sub

Private Sub Form_GotFocus()
    txtNumeroPedido.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub ocultaNavegador(ativa As Boolean)
    imgOcultaNavegador.left = webNavegador.left
    imgOcultaNavegador.top = webNavegador.top
    imgOcultaNavegador.Width = webNavegador.Width
    imgOcultaNavegador.Height = webNavegador.Height
    imgOcultaNavegador.Visible = ativa
End Sub

Private Sub cmdEncerrar_Click()
    Screen.MousePointer = 0
    rsItensGarantia.Close
    
    txtNumeroPedido.Text = Empty
    lblStatusImpressao.Visible = False
    cmdImprimirTodos.Visible = True
    fraPedido.Visible = True
    'rsItensGarantia.Close
    
    Unload Me
End Sub

Private Sub cmdImprimirTodos_Click()
    Screen.MousePointer = 11

    lblStatusImpressao.Caption = mensagemImpressao & " " & statusImpressao
    lblStatusImpressao.Visible = True
    cmdImprimirTodos.Visible = False
    fraPedido.Visible = False
    
    ocultaNavegador True

    Dim i As Byte
    For i = 0 To 1
        
    Next i

End Sub

'Private Sub cmdImprimir_Click()
'    timerImprimir2_Timer
'End Sub

'Private Sub cmdItemAnterior_Click()
'    controleNavegacao -1
'End Sub
'
'Private Sub cmdItemProximo_Click()
'    controleNavegacao 1
'End Sub

'Private Sub controleNavegacao(proximoItem As Integer)
'    If (itemExibido + proximoItem) > 0 And (itemExibido + proximoItem) <= rsItensGarantia.RowCount Then
'        If proximoItem >= 0 Then
'            rsItensGarantia.MoveNext
'        Else
'            rsItensGarantia.MovePrevious
'        End If
'        carregaItem
'        criaHTML
'        itemExibido = itemExibido + proximoItem
'        montaCombo rsItensGarantia("CertificadoInicio"), rsItensGarantia("lojaOrigem"), rsItensGarantia("qtdeGarantia")
'        informacaoQuantidadeItem (itemExibido)
'    End If
'End Sub

Public Sub verificaNumeroPedido()
    If frmFormaPagamento.txtPedido.Text <> "" Then
        frmCertificado.Visible = False
        
        txtNumeroPedido.Text = frmFormaPagamento.txtPedido.Text
        txtNumeroPedido_KeyPress (13)
        'cmdImprimir.Enabled = False
        'frmCaixa.WindowState = 1
        
        fraPedido.Visible = False
        cmdImprimirTodos_Click
        
    Else
        
        'criaHTML
        frmCertificado.Visible = True
        txtNumeroPedido.Enabled = True
        txtNumeroPedido.SetFocus
        'Unload frmCaixa
    End If
End Sub

Private Function carregaListaItem(numeroPedido As String) As Boolean
    Dim sql As String
    'Dim rsItensGarantia
    
    carregaListaItem = True
    
    sql = "SP_FIN_Garantia_Estendida_Monta_Campos " & numeroPedido
    
    rsItensGarantia.CursorLocation = adUseClient
    rsItensGarantia.Open sql, rdoCNLoja, adOpenDynamic, adLockPessimistic
    
    'Set rsItensGarantia = rdoCNLoja.OpenResultset(SQL, rdOpenKeyset)
    
    If rsItensGarantia.EOF Then
        carregaListaItem = False
        rsItensGarantia.Close
    End If
    
End Function

Private Sub atualizaNumeroCertificado(numeroPedido As String, ByRef rsItens)
    Dim sql As String
    
    Do While Not rsItens.EOF
        sql = sql & "update nfitens set " & vbNewLine & _
              "certificadoInicio = (select ct_certificado from controle)+1, " & vbNewLine & _
              "certificadoFim = ((select ct_certificado from controle)+qtdeGarantia) " & vbNewLine & _
              "where numeroPed = " & numeroPedido & " and garantiaEstendida = 'S' and " & vbNewLine & _
              "referencia = " & rsItens("Referencia") & vbNewLine & vbNewLine
              
        sql = sql & "update controle set " & vbNewLine & _
              "ct_certificado = (select certificadoFIM from nfitens " & vbNewLine & _
              "where numeroPed = " & numeroPedido & " and garantiaEstendida = 'S' and " & vbNewLine & _
              "referencia = " & rsItens("Referencia") & ")" & vbNewLine & vbNewLine
        rsItens.MoveNext
    Loop

    rdoCNLoja.Execute (sql)
    'rsItensGarantia.Close
    ''carregaListaItem numeroPedido '''''''''''''''''''''''''''''''''''''''''''''''

End Sub

Private Function possuiNumeroCertificado(numeroPedido As String) As Boolean
    Dim sql As String
    
    possuiNumeroCertificado = True

          
    Set rsProdutoGarantia = rdoCNLoja.OpenResultset(sql)
        If IsNull(rsProdutoGarantia("certificadoInicio")) Then
            possuiNumeroCertificado = False
        End If
    rsProdutoGarantia.Close
End Function


Private Sub obterQuantidadeTotalItem(numeroPedido As String)
    Dim sql As String
    
    sql = "select count(*) quantidade from nfitens " & _
          "where numeroPed = " & numeroPedido & " and garantiaEstendida = 'S'"
    
    rsProdutoGarantia.CursorLocation = adUseClient
    rsProdutoGarantia.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

    'Set rsProdutoGarantia = rdoCNLoja.OpenResultset(SQL)
        quantidadeTotalItens = rsProdutoGarantia("quantidade")
    rsProdutoGarantia.Close
End Sub

Private Sub informacaoQuantidadeItem(Item As Integer)
    'lblItemGarantia.Caption = "Certificado de seguro " & Item & " de " & quantidadeTotalItens
End Sub

Private Sub obterTotalItemImpressao(numeroPedido As String)
    Dim sql As String
    sql = "select sum(qtdeGarantia) quantidade from nfitens where numeroPed = " & numeroPedido
    
    rsProdutoGarantia.CursorLocation = adUseClient
    rsProdutoGarantia.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    'Set rsProdutoGarantia = rdoCNLoja.OpenResultset(SQL)
        pocentagemItemImpressao = ((100 / (rsProdutoGarantia("quantidade"))) / 2) - 0.1
    rsProdutoGarantia.Close
End Sub

Private Sub timerImprimir_Timer()
'    statusImpressao = statusFuncionamento(statusImpressao)
'    lblStatusImpressao.Caption = mensagemImpressao & " " & statusImpressao
'    lblStatusImpressao.Refresh
'
'    If Not rsItensGarantia.EOF Then
'        If itemImpressao < rsItensGarantia("qtdeGarantia") Then
'            If numeroDeCopia < 2 Then
'                If numeroDeCopia = 0 Then
'                    criaHTML True
'                End If
'                timerImprimir2.Interval = ((timerImprimir.Interval) / 2) + 100
'                timerImprimir2.Enabled = True
'                numeroDeCopia = numeroDeCopia + 1
'            Else
'                itemImpressao = itemImpressao + 1
'                numeroDeCopia = 0
'            End If
'        Else
'            rsItensGarantia.MoveNext
'            numeroDeCopia = 0
'            itemImpressao = 0
'            If Not rsItensGarantia.EOF Then carregaItem
'        End If
'    Else
'        timerImprimir.Enabled = False
'        'rsItensGarantia.Close
'        cmdEncerrar_Click
'    End If
End Sub

Private Sub timerImprimir2_Timer()
'    timerImprimir2.Enabled = False
 '   webNavegador.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub txtNumeroPedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
        
    If KeyAscii = 13 Then
        Screen.MousePointer = 11
        
        If IsNumeric(txtNumeroPedido) Then
            pedido = txtNumeroPedido
            'txtNumeroPedido.Text = Empty
            possuiNumCertificado (pedido)
            
            If carregaListaItem(pedido) Then
                carregaInformacoes pedido
                carregaItem
                
                criaHTML True
                imgOcultaNavegador.Visible = False
                cmdImprimirTodos.Visible = True
                
                obterQuantidadeTotalItem pedido
                obterTotalItemImpressao pedido
                itemExibido = 1
                
            Else
                MsgBox "Não há contrato de Garantia Estendida para o pedido " & pedido, vbExclamation, Me.Caption
                txtNumeroPedido.Text = Empty
            End If
        Else
            txtNumeroPedido.Text = Empty
        End If
        
        Screen.MousePointer = 0
        
    End If
End Sub

Private Sub possuiNumCertificado(numeroPedido)

    Dim sql As String
    Dim rsItens As New ADODB.Recordset
    
    sql = "select itens.referencia as referencia " & vbNewLine & _
          "from nfitens itens, nfcapa capa " & vbNewLine & _
          "where itens.garantiaEstendida = 'S' AND itens.certificadoInicio is null and " & vbNewLine & _
          "capa.numeroPed = itens.numeroPed and capa.tipoNota = 'V' and " & vbNewLine & _
          "capa.garantiaEstendida = 'S' and capa.numeroPed = " & numeroPedido & "" & vbNewLine & _
          "order by itens.item"

    rsItens.CursorLocation = adUseClient
    rsItens.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

    'Set rsItens = rdoCNLoja.OpenResultset(SQL)
        If Not rsItens.EOF Then
            atualizaNumeroCertificado pedido, rsItens
        End If
    rsItens.Close
    
End Sub

Private Sub montaCombo(numeroDoCertificado As String, Loja As String, quantidadeItem As Integer)
    Dim i As Integer
    'cmbCertificado.Clear
''    For i = 0 To quantidadeItem - 1
''        If i = 0 Then cmbCertificado.Text = formatNumCertificado(Loja, numeroDoCertificado + i, _
''                                            codigoEstipulante, codigoInternoProduto)
''        cmbCertificado.AddItem formatNumCertificado(Loja, numeroDoCertificado + i, _
''                               codigoEstipulante, codigoInternoProduto)
''
''    Next i
    'cmbCertificado.ListIndex = 0
End Sub

Public Function formatNumCertificado(Loja As String, Numero As Integer, _
codEstipulante As String, codInternoProduto As String) As String
    
    formatNumCertificado = codEstipulante & codInternoProduto & _
                           Format(Loja, "000000") & _
                           Format(Numero, "000000")
End Function


Private Sub limpaTodasVariaveis()
    pedido = ""
    apoliceNumero = ""
    estipulante = ""
    numeroDoCertificado = ""
    telefoneParaContato = ""
    objetivoSeguro = ""
    segurado = ""
    CPF = ""
    outroDoc = ""
    tipoDoc = ""
    Endereco = ""
    Numero = ""
    Complemento = ""
    Bairro = ""
    Cidade = ""
    uf = ""
    CEP = ""
    Telefone = ""
    email = ""
    dataContratacaoSeguro = Date
    dataCompraBem = Date
    perildoGarantiaOriginalInicio = Date
    perildoGarantiaOriginalFim = Date
    perildoCoberturaSeguroInicio = Date
    perildoCoberturaSeguroFim = Date
    bemSegurado = ""
    Marca = ""
    Modelo = ""
    valorBemSegurado = ""
    PremioLiquido = 0
    iof = 0
    premioTotal = 0
    percentualRemuneracao = 0
    valorRemuneracao = 0
    cidadeAtual = ""
    Data = ""
    
    referencia = ""
    Item = ""
    quantidadeGarantia = 0
'    quantidadeTotalItem = 0
    itemImpressao = 0
    numeroDeCopia = 0
    NumeroItem = 0
    pocentagemItemImpressao = 0

End Sub

Private Function formatValorExibicao(Valor As String) As String
    formatValorExibicao = Format(Valor, "##,#0.00")
End Function

Private Function formatCampoExibir(Valor As Double) As String
    formatCampoExibir = Format(Valor, "##,#0.00")
End Function

Private Sub carregaInformacoes(numeroPedido As String)
    montaCamposApolice
    montaCamposCliente numeroPedido
    montaCamposCidadeData numeroPedido
End Sub

Private Sub montaCamposApolice()
    Dim sql As String
    sql = "select ct_razao,CT_codigoEstipulanteGE, CT_codigoInternoProdutoGE, ct_apolice " & _
          "from controle"
          
rsProdutoGarantia.CursorLocation = adUseClient
rsProdutoGarantia.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    'Set rsProdutoGarantia = rdoCNLoja.OpenResultset(SQL)
        apoliceNumero = rsProdutoGarantia("ct_apolice")
        codigoEstipulante = rsProdutoGarantia("CT_codigoEstipulanteGE")
        codigoInternoProduto = rsProdutoGarantia("CT_codigoInternoProdutoGE")
        estipulante = rsProdutoGarantia("ct_razao")
    rsProdutoGarantia.Close
End Sub

Private Sub montaCamposCliente(numeroPedido As String)

    On Error GoTo erroLeituraCliente
    
    Dim sql As String
    sql = "select CE_Razao, ce_cgc,ce_inscricaoEstadual,ce_endereco, ce_numero,ce_complemento, " & _
          "ce_bairro, ce_municipio, ce_estado,ce_cep,ce_telefone,ce_email from nfcapa, fin_cliente " & _
          "where numeroPed = " & numeroPedido & " and cliente = ce_codigoCliente"
          
    rsProdutoGarantia.CursorLocation = adUseClient
    rsProdutoGarantia.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    'Set rsProdutoGarantia = rdoCNLoja.OpenResultset(SQL)
        segurado = rsProdutoGarantia("CE_Razao")
        CPF = rsProdutoGarantia("ce_cgc")
        outroDoc = rsProdutoGarantia("ce_inscricaoEstadual")
        tipoDoc = ""
        Endereco = rsProdutoGarantia("ce_endereco")
        Numero = rsProdutoGarantia("ce_numero")
        Complemento = rsProdutoGarantia("ce_complemento")
        Bairro = rsProdutoGarantia("ce_bairro")
        Cidade = rsProdutoGarantia("ce_municipio")
        uf = rsProdutoGarantia("ce_estado")
        CEP = rsProdutoGarantia("ce_cep")
        Telefone = Format(rsProdutoGarantia("ce_telefone"), "(##) 0000 0000")
        email = rsProdutoGarantia("ce_email")
    rsProdutoGarantia.Close
    
erroLeituraCliente:
    Select Case Err.Number
        Case 40009
            MsgBox "Não foi possível verificar as informações sobre o cliente" & vbNewLine _
            & "Verifique se o pedido " & numeroPedido & " possui o código do cliente", vbCritical, frmCertificado.Caption
    End Select
End Sub

Private Sub montaCamposCidadeData(numeroPedido As String)
    Dim sql As String
    sql = "select lo_municipio municipio, dataemi dataEmissao from lojas, nfcapa where numeroPed = " & numeroPedido & " and lojaorigem = lo_loja"

rsProdutoGarantia.CursorLocation = adUseClient
rsProdutoGarantia.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

'    Set rsProdutoGarantia = rdoCNLoja.OpenResultset(SQL)
        cidadeAtual = StrConv(rsProdutoGarantia("municipio"), 3)
        Data = Format(rsProdutoGarantia("dataEmissao"), "D MMMM") & " de " & Format(rsProdutoGarantia("dataEmissao"), "YYYY")
    rsProdutoGarantia.Close
End Sub

''''' CAMPOS ITENS ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub carregaItem()
    Dim calcPis As Double
    Dim calcConfis As Double
    Dim calcNET As Double

    numeroDoCertificado = formatNumCertificado(rsItensGarantia("lojaOrigem"), _
                      rsItensGarantia("certificadoInicio") + itemImpressao, _
                      codigoEstipulante, codigoInternoProduto)
    objetivoSeguro = "R"
    dataContratacaoSeguro = rsItensGarantia("dataemissao")
    dataCompraBem = rsItensGarantia("dataemissao")
    perildoGarantiaOriginalInicio = dataCompraBem
    perildoGarantiaOriginalFim = DateAdd("m", rsItensGarantia("garantiaFabricante"), dataCompraBem)
    perildoCoberturaSeguroInicio = DateAdd("D", 1, perildoGarantiaOriginalFim)
    perildoCoberturaSeguroFim = DateAdd("M", rsItensGarantia("planoGarantia"), dataCompraBem)
    
    bemSegurado = RTrim(rsItensGarantia("pr_descricao"))
    Marca = rsItensGarantia("Marca")
    Modelo = rsItensGarantia("referencia")
    valorBemSegurado = rsItensGarantia("VLUNIT")
    
    PremioLiquido = rsItensGarantia("premioLiquido")
    iof = rsItensGarantia("IOF")
    premioTotal = rsItensGarantia("premioTotal")
    
    calcPis = (PremioLiquido * 0.65) / 100
    calcConfis = (PremioLiquido * 4) / 100
    calcNET = rsItensGarantia("CustoDaSegurandora")
    
    percentualRemuneracao = premioTotal - iof - calcPis - calcConfis - calcNET
    valorRemuneracao = (percentualRemuneracao / premioTotal) * 100
    
End Sub


'///////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////// CODIGO HTML /////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////


Private Sub criaHTML(corParaImpressao As Boolean)

    Dim enderecoPadraoHTML As String
    Dim codigoHTML As String
    Dim corFundo As String
    Dim corFonte As String
    
    If corParaImpressao = True Then
        corFundo = "White"
        corFonte = "Black"
    Else
        corFundo = "Black"
        corFonte = "White"
    End If
    
    enderecoPadraoHTML = Environ("TEMP") & "\" & numeroDeCopia & numeroDoCertificado & ".html"
    'enderecoPadraoHTML = Environ("TEMP") & "\certificado" & rsItensGarantia("Referencia") & ".html"
    
    codigoHTML = "<html xmlns:v=" & Chr(34) & "urn:schemas-microsoft-com:vml" & Chr(34) & "" & vbNewLine & _
    "xmlns:o=" & Chr(34) & "urn:schemas-microsoft-com:office:office" & Chr(34) & "" & vbNewLine & _
    "xmlns:w=" & Chr(34) & "urn:schemas-microsoft-com:office:word" & Chr(34) & "" & vbNewLine & _
    "xmlns:m=" & Chr(34) & "http://schemas.microsoft.com/office/2004/12/omml" & Chr(34) & "" & vbNewLine & _
    "xmlns=" & Chr(34) & "http://www.w3.org/TR/REC-html40" & Chr(34) & "> " & vbNewLine & _
    "<head>" & vbNewLine & _
    "<meta http-equiv=Content-Type content=" & Chr(34) & "text/html; charset=windows-1252" & Chr(34) & ">" & vbNewLine & _
    "<meta name=ProgId content=Word.Document>" & vbNewLine & _
    "<meta name=Generator content=" & Chr(34) & "Microsoft Word 15" & Chr(34) & ">" & vbNewLine & _
    "<meta name=Originator content=" & Chr(34) & "Microsoft Word 15" & Chr(34) & ">" & vbNewLine & _
    "<link rel=File-List" & vbNewLine & _
    "href=" & Chr(34) & "IMP%209%20-LayoutCertEletronicoPDV_GE_BW_Mar13_v_16042013_arquivos/filelist.xml" & Chr(34) & ">" & vbNewLine & _
    "<title>ANEXO 2</title>" & vbNewLine & _
    "<link rel=themeData" & vbNewLine & _
    "href=" & Chr(34) & "IMP%209%20-LayoutCertEletronicoPDV_GE_BW_Mar13_v_16042013_arquivos/themedata.thmx" & Chr(34) & ">" & vbNewLine & _
    "<link rel=colorSchemeMapping" & vbNewLine & _
    "href=" & Chr(34) & "IMP%209%20-LayoutCertEletronicoPDV_GE_BW_Mar13_v_16042013_arquivos/colorschememapping.xml" & Chr(34) & ">"
    
    codigoHTML = codigoHTML & vbNewLine & _
    "<style>" & vbNewLine & _
    "<!--" & vbNewLine & _
     "/* Font Definitions */" & vbNewLine & _
     "@font-face" & vbNewLine & _
        "{font-family:" & Chr(34) & "Cambria Math" & Chr(34) & ";" & vbNewLine & _
        "panose-1:2 4 5 3 5 4 6 3 2 4;" & vbNewLine & _
        "mso-font-charset:1;" & vbNewLine & _
        "mso-generic-font-family:roman;" & vbNewLine & _
        "mso-font-format:other;" & vbNewLine & _
        "mso-font-pitch:variable;" & vbNewLine & _
        "mso-font-signature:0 0 0 0 0 0;}" & vbNewLine & _
    "@font-face" & vbNewLine & _
        "{font-family:Garamond;" & vbNewLine & _
        "panose-1:2 2 4 4 3 3 1 1 8 3;" & vbNewLine & _
        "mso-font-charset:0;" & vbNewLine & _
        "mso-generic-font-family:roman;" & vbNewLine & _
        "mso-font-pitch:variable;" & vbNewLine & _
        "mso-font-signature:647 0 0 0 159 0;}"
        
    codigoHTML = codigoHTML & vbNewLine & _
    "@font-face" & vbNewLine & _
        "{font-family:Tahoma;" & vbNewLine & _
        "panose-1:2 11 6 4 3 5 4 4 2 4;" & vbNewLine & _
        "mso-font-charset:0;" & vbNewLine & _
        "mso-generic-font-family:swiss;" & vbNewLine & _
        "mso-font-pitch:variable;" & vbNewLine & _
        "mso-font-signature:-520081665 -1073717157 41 0 66047 0;}" & vbNewLine & _
    "@font-face" & vbNewLine & _
        "{font-family:" & Chr(34) & "Arial Narrow" & Chr(34) & ";" & vbNewLine & _
        "panose-1:2 11 6 6 2 2 2 3 2 4;" & vbNewLine & _
        "mso-font-charset:0;" & vbNewLine & _
        "mso-generic-font-family:swiss;" & vbNewLine & _
        "mso-font-pitch:variable;" & vbNewLine & _
        "mso-font-signature:647 2048 0 0 159 0;}" & vbNewLine & _
     "/* Style Definitions */"
     
     codigoHTML = codigoHTML & vbNewLine & _
     "P.MsoNormal , li.MsoNormal, div.MsoNormal" & vbNewLine & _
        "{mso-style-unhide:no;" & vbNewLine & _
        "mso-style-qformat:yes;" & vbNewLine & _
        "mso-style-parent:" & Chr(34) & "" & Chr(34) & ";" & vbNewLine & _
        "margin:0cm;" & vbNewLine & _
        "margin-bottom:.0001pt;" & vbNewLine & _
        "mso-pagination:widow-orphan;" & vbNewLine & _
        "font-size:10.0pt;" & vbNewLine & _
        "font-family:" & Chr(34) & "Times New Roman" & Chr(34) & "," & Chr(34) & "serif" & Chr(34) & ";" & vbNewLine & _
        "mso-fareast-font-family:" & Chr(34) & "Times New Roman" & Chr(34) & ";" & vbNewLine & _
        "mso-no-proof:yes;}"
        
    codigoHTML = codigoHTML & vbNewLine & _
    "P.MsoBodyText , li.MsoBodyText, div.MsoBodyText" & vbNewLine & _
        "{mso-style-unhide:no;" & vbNewLine & _
        "margin-top:0cm;" & vbNewLine & _
        "margin-right:0cm;" & vbNewLine & _
        "margin-bottom:12.0pt;" & vbNewLine & _
        "margin-left:0cm;" & vbNewLine & _
        "text-align:justify;" & vbNewLine & _
        "mso-pagination:widow-orphan;" & vbNewLine & _
        "font-size:12.0pt;" & vbNewLine & _
        "mso-bidi-font-size:10.0pt;" & vbNewLine & _
        "font-family:" & Chr(34) & "Garamond" & Chr(34) & "," & Chr(34) & "serif" & Chr(34) & ";" & vbNewLine & _
        "mso-fareast-font-family:" & Chr(34) & "Times New Roman" & Chr(34) & ";" & vbNewLine & _
        "mso-bidi-font-family:" & Chr(34) & "Times New Roman" & Chr(34) & ";" & vbNewLine & _
        "letter-spacing:-.25pt;" & vbNewLine & _
        "mso-no-proof:yes;}"
        
    codigoHTML = codigoHTML & vbNewLine & _
    "a: link , span.MsoHyperlink" & vbNewLine & _
        "{mso-style-unhide:no;" & vbNewLine & _
        "color:blue;" & vbNewLine & _
        "mso-themecolor:hyperlink;" & vbNewLine & _
        "text-decoration:underline;" & vbNewLine & _
        "text-underline:single;}" & vbNewLine & _
    "a: visited , span.MsoHyperlinkFollowed" & vbNewLine & _
        "{mso-style-noshow:yes;" & vbNewLine & _
        "color:purple;" & vbNewLine & _
        "mso-themecolor:followedhyperlink;" & vbNewLine & _
        "text-decoration:underline;" & vbNewLine & _
        "text-underline:single;}"
    
    codigoHTML = codigoHTML & vbNewLine & _
    "P.Endereodoremetente , li.Endereodoremetente, div.Endereodoremetente" & vbNewLine & _
        "{mso-style-name:" & Chr(34) & "Endereço do remetente" & Chr(34) & ";" & vbNewLine & _
        "mso-style-unhide:no;" & vbNewLine & _
        "margin:0cm;" & vbNewLine & _
        "margin-bottom:.0001pt;" & vbNewLine & _
        "text-align:center;" & vbNewLine & _
        "mso-pagination:widow-orphan;" & vbNewLine & _
        "font-size:10.0pt;" & vbNewLine & _
        "font-family:" & Chr(34) & "Garamond" & Chr(34) & "," & Chr(34) & "serif" & Chr(34) & ";" & vbNewLine & _
        "mso-fareast-font-family:" & Chr(34) & "Times New Roman" & Chr(34) & ";" & vbNewLine & _
        "mso-bidi-font-family:" & Chr(34) & "Times New Roman" & Chr(34) & ";" & vbNewLine & _
        "letter-spacing:-.15pt;" & vbNewLine & _
        "mso-no-proof:yes;}" & vbNewLine & _
    ".MsoChpDefault" & vbNewLine & _
        "{mso-style-type:export-only;" & vbNewLine & _
        "mso-default-props:yes;" & vbNewLine & _
        "font-size:10.0pt;" & vbNewLine & _
        "mso-ansi-font-size:10.0pt;" & vbNewLine & _
        "mso-bidi-font-size:10.0pt;}"
    
    codigoHTML = codigoHTML & vbNewLine & _
    "@page WordSection1" & vbNewLine & _
        "{size:21.0cm 842.0pt;" & vbNewLine & _
        "margin:1.0cm 42.55pt 1.0cm 39.7pt;" & vbNewLine & _
        "mso-header-margin:5.65pt;" & vbNewLine & _
        "mso-footer-margin:5.65pt;" & vbNewLine & _
        "mso-gutter-margin:1.0cm;" & vbNewLine & _
        "mso-paper-source:0;}" & vbNewLine & _
    "div.WordSection1" & vbNewLine & _
        "{page:WordSection1;}" & vbNewLine & _
    "-->" & vbNewLine & _
    "</style>"
    
    codigoHTML = codigoHTML & "</head>" & vbNewLine & _
    "<body style=" & Chr(34) & "background:" & corFundo & Chr(34) & " lang=PT-BR link=blue vlink=purple style='tab-interval:35.4pt'>" & vbNewLine & _
    "<span style=" & Chr(34) & "color:" & corFonte & Chr(34) & ";>" & vbNewLine & _
    "<div class=WordSection1>" & vbNewLine & _
    "<p class=MsoBodyText align=left style='margin-bottom:0cm;margin-bottom:.0001pt;" & vbNewLine & _
    "Text -Align: Left '><b style='mso-bidi-font-weight:normal'><span style='font-size:" & vbNewLine & _
    "11.0pt;font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & ";letter-spacing:0pt'>CERTIFICADO DE " & vbNewLine & _
    "SEGURO EXTENSÃO DE GARANTIA DIFERENCIADA</><o:p></o:p></span></b></p>" & vbNewLine & _
    "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Virginia Surety Cia. de Seguros do " & vbNewLine & _
    "Brasil - CNPJ 03.505.295/0001-46<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Garantia Estendida Processo SUSEP No." & vbNewLine & _
    "15.414.003.197/2006-48 - www.virginiasurety.com.br<i style='mso-bidi-font-style:normal'><o:p></o:p></i></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><i style='mso-bidi-font-style:" & vbNewLine & _
    "Normal '><span style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>" & Chr(34) & "O registro deste plano " & vbNewLine & _
    "na Susep não<span style='mso-spacerun:yes'>  </span>implica, por parte da" & vbNewLine & _
    "Autarquia, incentivo ou recomendação a sua comercialização." & Chr(34) & "<o:p></o:p></span></i></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'><o:p>&nbsp;</o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal style='text-align:justify'><b style='mso-bidi-font-weight:" & vbNewLine & _
    "Normal '><span style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Apólice nº: " & vbNewLine & _
    apoliceNumero & "<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Eletroeletrônicos / Móveis<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Estipulante: " & estipulante & "<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><span style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'><o:p>&nbsp;</o:p></span></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal style='text-align:justify'><b style='mso-bidi-font-weight:" & vbNewLine & _
    "Normal '><span style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>NÚMERO DO CERTIFICADO: " & vbNewLine & _
    numeroDoCertificado & "</span></b><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Arial Narrow" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & ";mso-bidi-font-family:Arial'><o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>TELEFONE PARA CONTATO: 0800 19 89 15</span></b><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'><o:p></o:p></span></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Atendimento a deficientes auditivos" & vbNewLine & _
    "ou de fala: 0800 728 0837.<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'><o:p>&nbsp;</o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Objetivo do Seguro: "
    If objetivoSeguro = "R" Then
        codigoHTML = codigoHTML & "(X) Reparação<span style='mso-tab-count:1'>  </span>( ) Troca<o:p></o:p></span></b></p>"
    ElseIf objetivoSeguro = "T" Then
        codigoHTML = codigoHTML & "( ) Reparação<span style='mso-tab-count:1'>  </span>(X) Troca<o:p></o:p></span></b></p>"
    Else
        codigoHTML = codigoHTML & "( ) Reparação<span style='mso-tab-count:1'>  </span>( ) Troca<o:p></o:p></span></b></p>"
    End If

    codigoHTML = codigoHTML & "<p class=MsoNormal><span style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'><o:p>&nbsp;</o:p></span></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Segurado: " & segurado & "<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'></b>CPF: " & CPF & "<span" & vbNewLine & _
    "style='mso-tab-count:1'> </span></span><span style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'><span" & vbNewLine & _
    "style='mso-tab-count:1'>            </span>Outro Doc: " & outroDoc & " " & vbNewLine & _
    "Tipo Doc: " & tipoDoc & "<o:p></o:p></span></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><span style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Endereço: " & vbNewLine & _
    Endereco & " Número: " & Numero & " <span style='mso-spacerun:yes'>" & vbNewLine & _
    "</span>Compl: " & Complemento & "<o:p></o:p></span></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><span style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Bairro: " & vbNewLine & _
    Bairro & "<span style='mso-tab-count:2'>                </span><o:p></o:p></span></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><span style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Cidade: " & vbNewLine & _
    Cidade & "<span style='mso-tab-count:1'> </span>UF: " & uf & "<o:p></o:p></span></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><span style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>CEP: " & vbNewLine & _
    CEP & "<o:p></o:p></span></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><span style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Telefone de contato: " & vbNewLine & _
    Telefone & "<o:p></o:p></span></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><span style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>E-mail: " & vbNewLine & _
    email & "<o:p></o:p></span></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><span style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'><o:p>&nbsp;</o:p></span></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Data Contratação Seguro: " & dataContratacaoSeguro & "<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Data Compra do Bem: " & dataCompraBem & "<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>PERÍODO DA GARANTIA ORIGINAL: " & vbNewLine & _
    perildoGarantiaOriginalInicio & " a " & perildoGarantiaOriginalFim & "<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>PERÍODO DE COBERTURA DO SEGURO: " & vbNewLine & _
    perildoCoberturaSeguroInicio & " a " & perildoCoberturaSeguroFim & "<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Limite Máximo de Indenização:" & vbNewLine & _
    "conforme Condições Gerais<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><span style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'><o:p>&nbsp;</o:p></span></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Bem Segurado: " & bemSegurado & "<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Marca: " & Marca & "<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Modelo: " & Modelo & "<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Valor do Bem Segurado: R$ " & Format(valorBemSegurado, "##,#0.00") & "<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'><o:p>&nbsp;</o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Prêmio Líquido: " & formatCampoExibir(PremioLiquido) & "<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>IOF: R$ " & formatCampoExibir(iof) & "<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Prêmio Total: R$ " & formatCampoExibir(premioTotal) & "<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'><o:p>&nbsp;</o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal style='text-align:justify'><b style='mso-bidi-font-weight:" & vbNewLine & _
    "Normal '><span style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Este certificado de " & vbNewLine & _
    "Seguro Extensão de Garantia Diferenciada será considerado válido caso a seguradora não " & vbNewLine & _
    "se manifeste formalmente em contrário, conforme disposto nas Condições Gerais e " & vbNewLine & _
    "Especiais deste seguro. O percentual de remuneração do estipulante é " & Format(valorRemuneracao, "00") & "%<sup> </sup>" & vbNewLine & _
    "e valor é R$ " & formatCampoExibir(percentualRemuneracao) & "<sup></sup>. A contratação do seguro é opcional, sendo possível a desistência do contrato em até 7 (sete) dias com a devolução integral do valor pago do seguro. O não repasse do prêmio do seguro, pelo estipulante à seguradora, pode gerar o cancelamento do seguro. <o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal style='text-align:justify'><b style='mso-bidi-font-weight:" & vbNewLine & _
    "Normal '><span style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'><o:p>&nbsp;</o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b><span style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Caso o" & vbNewLine & _
    "segurado já tenha recorrido à Central de Atendimento e não tenha se sentido " & vbNewLine & _
    "satisfeito com a solução apresentada, poderá recorrer à Ouvidoria pelo email: " & vbNewLine & _
    "ouvidoria@thewarrantygroup.com ou pelo telefone 0800 7274586. </span></b><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'><o:p></o:p></span></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><span style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'><o:p>&nbsp;</o:p></span></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal style='text-align:justify'><b style='mso-bidi-font-weight:" & vbNewLine & _
    "Normal '><span style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Declaro que: as " & vbNewLine & _
    "informações contidas neste certificado são verdadeiras, tive conhecimento " & vbNewLine & _
    "prévio da íntegra das Condições Gerais e Especiais do Seguro de Extensão de " & vbNewLine & _
    "Garantia Diferenciada, recebi o respectivo Resumo dessas Condições e concordo " & vbNewLine & _
    "com todos os termos neles contidos. <o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'><o:p>&nbsp;</o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>" & cidadeAtual & ", " & Data & "<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'><o:p>&nbsp;</o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>___________________________<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Assinatura do Segurado<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'><o:p>&nbsp;</o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'>Autenticação/Comprovante de Pagamento" & vbNewLine & _
    "<o:p></o:p></span></b></p>"
    
    codigoHTML = codigoHTML & "<p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span" & vbNewLine & _
    "style='font-family:" & Chr(34) & "Tahoma" & Chr(34) & "," & Chr(34) & "sans-serif" & Chr(34) & "'><o:p>&nbsp;</o:p></span></b></p>"
    
    codigoHTML = codigoHTML & " "
    
    codigoHTML = codigoHTML & " " & vbNewLine & _
    "</div>" & vbNewLine & _
    "</body>" & vbNewLine & _
    "</html>"

    Open enderecoPadraoHTML For Output As #1
         Print #1, codigoHTML
    Close #1
    
    webNavegador.Navigate enderecoPadraoHTML

End Sub

Private Sub webNavegador_GotFocus()
    frmCertificado.SetFocus
End Sub

