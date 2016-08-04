VERSION 5.00
Begin VB.Form frmTEFVariosCartoes 
   BackColor       =   &H00F1D1A7&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exemplo de Venda com Transação TEF, usando a BemaFI32.dll"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9465
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCupom 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   120
      Width           =   6255
   End
   Begin VB.CommandButton cmdFinalizarCupom 
      Caption         =   "&Terminar Fechamento Cupom"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Frame fraFormaPagto 
      BackColor       =   &H00F1D1A7&
      Caption         =   "Forma de Pagamento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   2895
      Begin VB.CommandButton cmdConfirmar 
         BackColor       =   &H8000000A&
         Caption         =   "&Confirmar"
         Default         =   -1  'True
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtValor 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """R$""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox cboFormaPagto 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label lblValor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F1D1A7&
         Caption         =   "Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1110
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdVenderItem 
      Caption         =   "&Vender Item"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   2895
   End
   Begin VB.CommandButton cmdFecharCupom 
      Caption         =   "&Iniciar Fechamento Cupom"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2895
   End
   Begin VB.CommandButton cmdAbrirCupom 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Abrir Cupom Fiscal"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2895
   End
   Begin VB.Image imgLogoBema 
      Height          =   855
      Index           =   0
      Left            =   0
      Picture         =   "TEFVariosCartoes.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblResta 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F8E8D3&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7680
      TabIndex        =   15
      Top             =   5805
      Width           =   1575
   End
   Begin VB.Label lblSubtotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F8E8D3&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4920
      TabIndex        =   14
      Top             =   5805
      Width           =   1575
   End
   Begin VB.Label lblR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F1D1A7&
      Caption         =   "Resta:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   13
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label lblS 
      BackColor       =   &H00F1D1A7&
      Caption         =   "Subtotal:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00F8E8D3&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   6360
      Width           =   9255
   End
   Begin VB.Label lblMensagem 
      BackColor       =   &H00F1D1A7&
      Caption         =   "Mensagem:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6120
      Width           =   1095
   End
End
Attribute VB_Name = "frmTEFVariosCartoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cFormaPGTO As String
Dim cNumeroCupom As String
Dim ValorPago() As String
Dim hora As Date
Dim pagtoCartao As Boolean
Dim acumulado As Double
Public i As Integer
Public j As Integer
Dim iConta As Integer
Dim curSubTotal As Currency
Dim curValorRestante As Currency
Private Sub cboFormaPagto_LostFocus()
    With txtValor
        If .Enabled And fraFormaPagto.Enabled Then
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            
        End If
    End With
End Sub

Private Sub cmdAbrirCupom_Click()

 
    Dim iRetorno As Integer
    iRetorno = Bematech_FI_CancelaCupom()
    
    Me.MousePointer = vbHourglass
    curSubTotal = 0
    iRetorno = Bematech_FI_AbreCupom("")
    
    'Verifica qual foi o retorno que a impressora deu após a execução da função
    If (VerificaRetornoFuncaoImpressora(iRetorno)) Then
        With txtCupom
            .Text = .Text & "-------------------------------------------------" & vbCrLf
            .Text = .Text & "                 Abrindo Cupom...                " & vbCrLf
            .Text = .Text & "-------------------------------------------------" & vbCrLf
            .Text = .Text & "                     ÍTENS                       " & vbCrLf
            .Text = .Text & "-------------------------------------------------" & vbCrLf
            .Text = .Text & " COD          DESCRICAO                   VALOR  " & vbCrLf
            .Text = .Text & "-------------------------------------------------" & vbCrLf
        End With
        cmdAbrirCupom.Enabled = False
        cmdVenderItem.Enabled = True
        cmdVenderItem.SetFocus
    End If
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdVenderItem_Click()
    Dim cCodigoProduto As String
    Dim cDescricaoProduto As String
    Dim cValorItem As String
    Dim iRetorno As Integer
        
    Me.MousePointer = vbHourglass
    cCodigoProduto = "1234567890123"
    cDescricaoProduto = "Teste de Venda de Item..."
    cValorItem = "1,50"
    
    'Função utilizada para realizar a venda de um item.
    iRetorno = Bematech_FI_VendeItem(cCodigoProduto, cDescricaoProduto, "II", "I", "1", 2, cValorItem, "%", "00,00")
    
    'Verifica se o retorno da função de vender ítem foi positivo
    If (VerificaRetornoFuncaoImpressora(iRetorno)) Then
    
        'Exibição dos valores na tela
        txtCupom.Text = txtCupom.Text & cCodigoProduto & "  " & cDescricaoProduto & "  " & cValorItem & vbCrLf
        curSubTotal = curSubTotal + CCur(cValorItem)
        lblSubtotal.Caption = Format((curSubTotal), "###,###,##0.00")
        txtValor.Text = lblSubtotal.Caption
        cmdFecharCupom.Enabled = True
    End If
    Me.MousePointer = vbDefault

End Sub
Private Sub cmdFecharCupom_Click()

    Dim iRetorno As Integer
    
    Me.MousePointer = vbHourglass
    
    'Função que realiza o início do fechamento do cupom fiscal.
    iRetorno = Bematech_FI_IniciaFechamentoCupom("A", "%", "00,00")
    
    If (VerificaRetornoFuncaoImpressora(iRetorno)) Then
        With txtCupom
            .Text = .Text & "-------------------------------------------------" & vbCrLf
            .Text = .Text & "           Iniciando Fechamento do Cupom...      " & vbCrLf
            .Text = .Text & "-------------------------------------------------" & vbCrLf
        End With
        cmdVenderItem.Enabled = False
        cmdFecharCupom.Enabled = False
        cmdConfirmar.Enabled = True
        fraFormaPagto.Enabled = True
        curValorRestante = curSubTotal
        lblResta.Caption = lblSubtotal.Caption
        cboFormaPagto.ListIndex = 1
        cboFormaPagto.SetFocus
        Call SendMessage(cboFormaPagto.hwnd, CB_SHOWDROPDOWN, 1, 0&)
    End If
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdConfirmar_Click()

    Dim iRetorno As Integer 'Retorno da função VerificaRetornoFuncaoImpressora
    Dim cValorPago As String
    Dim i As Integer
    Dim iArquivo As Integer
    Dim strLeitura As String
    Dim iRetRealiza As Integer
        
    cmdConfirmar.Enabled = False
    If Not IsNumeric(txtValor.Text) Then
        MsgBox "O valor " & txtValor.Text & " não é válido. Insira um valor no formato 000.000,00", vbInformation + vbOKOnly, "Valor inválido"
        cmdConfirmar.Enabled = True
        Exit Sub
    End If
    
    'Não deixa pagar no cartão um valor maior do que o restante
    If (cFormaPGTO = "Cartao Credito") Then
       If CDec(txtValor.Text) > CDec(lblResta.Caption) Then
           MsgBox "Não pode ser pago um valor maior que o restante.", vbOKOnly + vbInformation, "TEF - Vários Cartões"
           cmdConfirmar.Enabled = True
           Exit Sub
       End If
    End If
    Me.MousePointer = vbHourglass
    lblMsg.Caption = ""
    cFormaPGTO = ""
    cFormaPGTO = Trim(cboFormaPagto.Text)
        
    'Se foi escolhido cartão de crédito...
    If (cFormaPGTO = "Cartao Credito") Then
        pagtoCartao = True
        
        'Se já existe mais de um pagamento, deve-se confirmar a transação anterior
        If Dir(App.Path & "\PENDENTE.TXT") <> "" Then
            iArquivo = FreeFile
            Open App.Path & "\PENDENTE.TXT" For Input As iArquivo
            Line Input #iArquivo, strLeitura
            If IsNumeric(CInt(Trim(strLeitura))) Then
                Close iArquivo
                ConfirmaTransacao (CInt(Trim(strLeitura)))
                MataArquivo (App.Path & "\PENDENTE.TXT")
            End If
        End If
        
        'Pega a hora atual
        hora = Time
        cNumeroCupom = Space(6)
        
        'Busca o número do cupom atual
        iRetorno = Bematech_FI_NumeroCupom(cNumeroCupom)
        VerificaRetornoFuncaoImpressora (iRetorno)
        'Verifica se a chave Retorno do BemaFI32.ini está habilitada ou não.
        If Trim(cNumeroCupom) = "" Then
            MsgBox "Não foi possível obter o número do cupom. " _
                    & "Verifique a chave Retorno do arquivo BemaFI32.ini"
            cmdConfirmar.Enabled = True
            Me.MousePointer = vbDefault
            Exit Sub
        End If
  
        cValorPago = txtValor.Text
        cValorPago = Replace(cValorPago, ",", "", , , vbTextCompare)
        iRetRealiza = RealizaTransacao(hora, cNumeroCupom, cValorPago, iConta)
        If iRetRealiza = 1 Then
            iRetorno = Bematech_FI_EfetuaFormaPagamento(cFormaPGTO, txtValor.Text)
            If (VerificaRetornoFuncaoImpressora(iRetorno)) Then
                ReDim Preserve ValorPago(iConta - 1)
                ValorPago(iConta - 1) = cValorPago
                iConta = iConta + 1
                iTransacao = iTransacao + 1
                curValorRestante = curValorRestante - CCur(txtValor.Text)
                lblResta.Caption = ""
                acumulado = acumulado + txtValor.Text
                lblResta.Caption = Format(CDec(lblSubtotal.Caption) - acumulado, "##,##0.00")
                txtCupom.Text = txtCupom.Text & cboFormaPagto.Text & "          " & txtValor.Text & vbCrLf
                txtValor.Text = lblResta.Caption
                If curValorRestante = 0 Then
                    cmdFinalizarCupom.Enabled = True
                    cmdConfirmar.Enabled = False
                    fraFormaPagto.Enabled = False
                Else
                    cmdConfirmar.Enabled = True
                End If
            End If
        Else
            cboFormaPagto.SetFocus
            cmdConfirmar.Enabled = True
        End If
    Else
        iRetorno = Bematech_FI_EfetuaFormaPagamento(cFormaPGTO, txtValor.Text)
        If (VerificaRetornoFuncaoImpressora(iRetorno)) Then
            curValorRestante = curValorRestante - CCur(txtValor.Text)
            lblResta.Caption = ""
            acumulado = acumulado + txtValor.Text
            lblResta.Caption = Format(lblSubtotal.Caption - acumulado, "##,##0.00")
            txtCupom.Text = txtCupom.Text & cboFormaPagto.Text & "          " & txtValor.Text & vbCrLf
            txtValor.Text = lblResta.Caption
            If curValorRestante = 0 Then
                cmdFinalizarCupom.Enabled = True
                cmdConfirmar.Enabled = False
                fraFormaPagto.Enabled = False
            Else
                cmdConfirmar.Enabled = True
            End If
        End If
    End If
    Me.MousePointer = vbDefault
End Sub


Private Sub cmdFinalizarCupom_Click()

    Dim mensagemPromocional As String
    Dim iRetorno As Integer
    Dim iQuantasTransacoes As Integer
    Dim iVezes As Integer
    Dim VlrPago As Variant
    Dim Gerencial As Boolean
    
    cmdFinalizarCupom.Enabled = False
    
    Me.MousePointer = vbHourglass
    If CDec(lblResta.Caption) > 0 Then
        Exit Sub
    End If
    
    
    cmdConfirmar.Enabled = False
    
    ''''''''''''FECHANDO O CUPOM''''''''''''
    mensagemPromocional = "Obrigado, volte sempre !!!"
    iRetorno = Bematech_FI_TerminaFechamentoCupom(mensagemPromocional)
    VerificaRetornoFuncaoImpressora (iRetorno)

    ''''''''''''ZERANDO VALORES''''''''''''
    txtCupom.Text = ""
    txtValor.Text = ""
    lblSubtotal.Caption = "0,00"
    lblResta.Caption = "0,00"
    
    If (pagtoCartao = True) Then
        Gerencial = False
        ''''''''''''IMPRIMINDO TRANSAÇÕES''''''''''''
        For iQuantasTransacoes = 1 To iTransacao Step 1
                If Not (ImprimeTransacao(cFormaPGTO, ValorPago(iQuantasTransacoes - 1), cNumeroCupom, hora, iQuantasTransacoes, Gerencial)) Then
                    Me.MousePointer = vbDefault
                    If naoConfirmado = True Then
                        Exit For
                    End If
                    If MsgBox("A impressora não responde!" & vbCrLf & "Deseja imprimir novamente?", vbYesNo + vbInformation, "Atenção") = vbYes Then
                        Gerencial = True
                        iQuantasTransacoes = 0
                    Else
                    ''''''''''''SE OPTAR POR NÃO IMPRIMIR AS TRANSAÇÕES NOVAMENTE,
                    ''''''''''''SERÁ FEITA A NÃO CONFIRMAÇÃO DELAS
                        Me.MousePointer = vbHourglass
                        NaoConfirmaTransacao (iTransacao)
                    End If
                End If
        Next iQuantasTransacoes
        
        ''''''''''''ZERANDO VARIÁVEIS''''''''''''
        For i = 0 To iConta - 2 Step 1
            ValorPago(i) = ""
        Next i
        Me.MousePointer = vbHourglass
        ''''''''''''CONFIRMANDO A ÚLTIMA TRANSAÇÃO (EM CASO DE NÃO TER SIDO FEITA
        ''''''''''''A NÃO CONFIRMAÇÃO)
        If ((iQuantasTransacoes - 1) = iTransacao) Then
            ConfirmaTransacao (iTransacao)
        End If
        MataArquivo (App.Path & "\PENDENTE.TXT")
        ''''''''''''MATANDO OS ARQUIVOS RESTANTES''''''''''''
        For iVezes = 1 To (iQuantasTransacoes - 1) Step 1
            If Dir("C:\TEF_DIAL\RESP\INTPOS" & CStr(iVezes) & ".001") <> "" Then
               MataArquivo ("C:\TEF_DIAL\RESP\INTPOS" & CStr(iVezes) & ".001")
               MataArquivo (App.Path & "\IMPRIME" & CStr(iVezes) & ".TXT")
            End If
        Next iVezes
     End If
    MataArquivo (App.Path & "\TEF.TXT")
    
    ''''''''''''INICIALIZANDO VARIÁVEIS''''''''''''
    cmdAbrirCupom.Enabled = True
    cmdVenderItem.Enabled = False
    cmdFecharCupom.Enabled = False
    fraFormaPagto.Enabled = False
    cboFormaPagto.Text = ""
    cmdFinalizarCupom.Enabled = False
    iConta = 1
    iTransacao = 0
    curValorRestante = 0
    curSubTotal = 0
    cFormaPGTO = ""
    cNumeroCupom = 0
    hora = "00:00:00"
    acumulado = 0
    i = 0
    j = 0
    naoConfirmado = False
    lblMsg.Caption = ""
    cmdAbrirCupom.SetFocus
    Me.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
    Dim iRetorno As Integer
    
    cboFormaPagto.AddItem "Dinheiro"
    cboFormaPagto.AddItem "Cartao Credito"
    cboFormaPagto.AddItem "Ticket"
    cboFormaPagto.AddItem "A prazo"
    cboFormaPagto.AddItem "Duplicata"
    
    If Dir(App.Path + "\TEF.TXT") <> "" Then
        'Função que realiza o fechamento do comprovante não fiscal vinculado,
        'caso esteja aberto.
        iRetorno = Bematech_FI_FechaComprovanteNaoFiscalVinculado
        'Abre o arquivo da última transação (TEF.TXT) para leitura
        If (VerificaRetornoFuncaoImpressora(iRetorno)) Then
           CancelarTransacoesPendentes
        End If
        MataArquivo (App.Path + "\TEF.TXT")
        
    End If
    
    'O conteúdo da combo forma de pagamento deve ser configurado
    'conforme as formas de pagamento existentes na impressora
    iConta = 1
    iTransacao = 0
    curValorRestante = 0
    cmdVenderItem.Enabled = False
    cmdConfirmar.Enabled = False
    fraFormaPagto.Enabled = False
    cmdFinalizarCupom.Enabled = False
    cmdFecharCupom.Enabled = False
    naoConfirmado = False
    Me.MousePointer = vbDefault
    
    MataArquivo (App.Path & "\PENDENTE.TXT")
   
End Sub
