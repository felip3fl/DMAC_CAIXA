VERSION 5.00
Object = "{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7u.ocx"
Begin VB.Form frmFechaCaixa 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Fechamento do Caixa"
   ClientHeight    =   8385
   ClientLeft      =   12600
   ClientTop       =   1605
   ClientWidth     =   5970
   FillColor       =   &H00808080&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox ChkFechamento 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Fechamento Geral "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   520
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   7
      Top             =   6450
      Width           =   3735
   End
   Begin VSFlex7UCtl.VSFlexGrid grdMovimentoCaixa 
      Height          =   5265
      Left            =   300
      TabIndex        =   4
      Top             =   600
      Width           =   4950
      _cx             =   8731
      _cy             =   9287
      _ConvInfo       =   1
      Appearance      =   0
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
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFechaCaixa.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
   Begin VB.TextBox txtSenha 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3690
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   7035
      Width           =   1300
   End
   Begin VB.TextBox txtOperador 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1500
      TabIndex        =   2
      Top             =   7035
      Width           =   1300
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   450
      Left            =   300
      Top             =   6360
      Width           =   4980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldos > 0 Serão retidos como Saldo de Caixa  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   525
      TabIndex        =   6
      Top             =   6030
      Width           =   4065
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   300
      Top             =   5970
      Width           =   4980
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   645
      Left            =   300
      Top             =   6885
      Width           =   4980
   End
   Begin VB.Label lblCabec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fechamento do Caixa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   300
      TabIndex        =   5
      Top             =   200
      Width           =   2280
   End
   Begin VB.Label lblSenha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3015
      TabIndex        =   1
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label lblOperador 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   525
      TabIndex        =   0
      Top             =   7080
      Width           =   900
   End
End
Attribute VB_Name = "frmFechaCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim wSubTotal As Double
Dim wSubTotal_S As Double
Dim wTotalEntrada As Double
Dim wSubTotal2 As Double
Dim wTotalSaldo As Double
Dim wTotalSaldo_S As Double
Dim wTotalQtde As Long
Dim wTotalNf As Double
Dim wTotalFatFin As Double
Dim wControlacor As Long
Dim wConfigCor As Long
Dim GuardaSequencia As Long
Dim wSaldoFinal As Double
Dim wSubTotalEntfin As Double
Dim wSangria As Double
Dim wMovimentoDoPeriodo As Double
Dim wEntradanoCaixa As Double
Dim wPegaImpressora As String
Dim wQuantidade As Double
Dim wDesconto As Double
Dim wPrecoUnitario As Double
Dim wTotalTipoNota As Double
Dim wControlaSaldoCaixa As Double
Dim wNomeSupervisor As String
Dim wVenda As Double
Dim wCancelamento As Double
Dim wDevolucao As Double
Dim wTR As Double
Dim wSubTotalEntFat As Double
Dim wSaldoNovo As Double
Dim wSaldoAnterior As Double
Dim wSaldoFinalDinheiro As Double
Dim wSaldoFinalCheque As Double
Dim wSaldoFinalAVR As Double
Dim wMovimentoPeriodo As Double
Dim wTotalTransferencia As Double
Dim wTNDinheiro As Double
Dim wTNCheque As Double
Dim wTNVisa As Double
Dim wTNRedecard As Double
Dim wTNAmex As Double
Dim wTNHiperCard As Double
Dim wTNBNDES As Double
Dim wTNVisaEletron As Double
Dim wTNRedeShop As Double
Dim wTNDeposito As Double
Dim wTNNotaCredito As Double
Dim wTNConducao As Double
Dim wTNDespLoja As Double
Dim wTNOutros As Double
Dim wTNTotal As Double
Dim wTNFinanciado As Double
Dim wTNFaturado As Double
Dim Linhas As Long
Dim datageral As Date
Dim wQtdeGrid As Integer
Dim Idx As Long
Dim sql As String
Dim Cor As String
Dim Cor1 As String
Dim Cor2 As String
Dim Cor3 As String
Dim notaPedente As Boolean

Private Sub cmbSair_Click()
 Unload Me
End Sub

Private Sub FechaCaixaOK()

  For Idx = 1 To grdMovimentoCaixa.Rows - 1 Step 1
    If Idx < 18 Then
       If (Idx = 1) Or (Idx = 2) Or (Idx = 9) Or (Idx = 17) Then
          If Idx = 1 Then
             wSaldoNovo = Format((grdMovimentoCaixa.TextMatrix(Idx, 3)), "###,###,###.00")
             wSaldoNovo = wSaldoNovo + Format(grdMovimentoCaixa.TextMatrix(17, 3), "###,###,###.00") 'Soma reforco com dinheiro
             wSaldoAnterior = Format(grdMovimentoCaixa.TextMatrix(40, 3), "###,###,###.00")
          ElseIf Idx = 2 Then
             wSaldoNovo = Format(grdMovimentoCaixa.TextMatrix(Idx, 3), "###,###,###.00")
             wSaldoAnterior = Format(grdMovimentoCaixa.TextMatrix(41, 3), "###,###,###.00")
          'ElseIf Idx = 9 Then
             'wSaldoNovo = Format(grdMovimentoCaixa.TextMatrix(Idx, 3), "###,###,###.00")
             'wSaldoAnterior = Format(grdMovimentoCaixa.TextMatrix(15, 3), "###,###,###.00")
 '         ElseIf Idx = 12 Then
 '            wSaldoNovo = Format(grdMovimentoCaixa.TextMatrix(Idx, 3), "###,###,###.00")
 '            wSaldoAnterior = Format(grdMovimentoCaixa.TextMatrix(16, 3), "###,###,###.00")
          End If
          
          wSaldoNovo = (wSaldoNovo + wSaldoAnterior)
          If wSaldoNovo > 0 Then
             Call GravaSaldoCaixa
          End If
       End If
    End If
  Next Idx
  Call FechaCaixa
  Call CarregaValoresTransfNumerario(GLB_CTR_Protocolo)
  
  wQdteViasImpressao = 1
  Call BuscaQtdeViaImpressaoMovimento
  
  For i = 1 To wQdteViasImpressao
      
      'Call ImprimeMovimento
        Call NOVO_ImprimeMovimento(grdMovimentoCaixa, "FECHAMENTO DE CAIXA", txtOperador.text, GLB_Caixa, _
                                   Format(GLB_DataInicial, "DD/MM/YYYY"), GLB_HoraInicial, _
                                   Date, Time, _
                                   CStr(GLB_CTR_Protocolo))
                                   
        Call NOVO_ImprimeTransfNumerario(grdMovimentoCaixa, "TRANSFERENCIA DE NUMERARIO", txtOperador.text, GLB_Caixa, _
                                         Format(GLB_DataInicial, "DD/MM/YYYY"), GLB_HoraInicial, _
                                         Date, Time, _
                                         CStr(GLB_CTR_Protocolo))
                                 
      'Call ImprimeMovimento00
      'Call ImprimeTransfNumerario
  Next i
  
  
  
  If ChkFechamento.Value = 0 Then
    Call AlterarResolucao(resolucaoOriginal.Colunas, resolucaoOriginal.Linhas)
    Unload Me
    Unload frmControlaCaixa
  End If
End Sub

Private Sub cmdRetornar_Click()
 Unload Me
End Sub

Private Sub cmdGravar_Click()

End Sub

Private Sub ChkFechamento_Click()
    
    
    If ChkFechamento.Value = 1 Then
        sql = "     SELECT * FROM  ControleCaixa WHERE  ctr_SITUACAOcaixa='A' AND ctr_supervisor <>99"
                rdoFechamentoGeral.CursorLocation = adUseClient
                rdoFechamentoGeral.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
             Linhas = rdoFechamentoGeral.RecordCount
        
        If Linhas = 1 Then
            rdoFechamentoGeral.MoveFirst
            datageral = Format(rdoFechamentoGeral("ctr_datainicial"), "yyyy/mm/dd")
        Else
            ChkFechamento.Value = 0
            MsgBox "Não é possivel fazer o  Fechamento Geral. Há Mais de Um Caixa em  Aberto !", vbInformation, "Atenção"
        End If
        
        rdoFechamentoGeral.Close

    End If

End Sub

Private Sub ChkFechamento_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ChkFechamento.Value = 1 Then
        If MsgBox("Após efetuar o Fechamento Geral não será possível emitir Nota Fiscal no dia. Deseja Continuar?", vbYesNo, "Atenção") = vbYes Then
            ChkFechamento.Value = 1
        Else
            ChkFechamento.Value = 0
        End If
    End If
End Sub

Private Sub Form_Activate()

    defineImpressora

    wControlaSaldoCaixa = 0

    wControlaSaldoCaixa = wControlaSaldoCaixa + (Format(grdMovimentoCaixa.TextMatrix(4, 3), "0.00"))
    wControlaSaldoCaixa = wControlaSaldoCaixa + (Format(grdMovimentoCaixa.TextMatrix(5, 3), "0.00"))
    wControlaSaldoCaixa = wControlaSaldoCaixa + (Format(grdMovimentoCaixa.TextMatrix(6, 3), "0.00"))
    wControlaSaldoCaixa = wControlaSaldoCaixa + (Format(grdMovimentoCaixa.TextMatrix(7, 3), "0.00"))
    wControlaSaldoCaixa = wControlaSaldoCaixa + (Format(grdMovimentoCaixa.TextMatrix(8, 3), "0.00"))
    wControlaSaldoCaixa = wControlaSaldoCaixa + (Format(grdMovimentoCaixa.TextMatrix(9, 3), "0.00"))
    wControlaSaldoCaixa = wControlaSaldoCaixa + (Format(grdMovimentoCaixa.TextMatrix(10, 3), "0.00"))
    wControlaSaldoCaixa = wControlaSaldoCaixa + (Format(grdMovimentoCaixa.TextMatrix(11, 3), "0.00"))
    wControlaSaldoCaixa = wControlaSaldoCaixa + (Format(grdMovimentoCaixa.TextMatrix(15, 3), "0.00"))
    wControlaSaldoCaixa = wControlaSaldoCaixa + (Format(grdMovimentoCaixa.TextMatrix(16, 3), "0.00"))
    wControlaSaldoCaixa = wControlaSaldoCaixa + (Format(grdMovimentoCaixa.TextMatrix(21, 3), "0.00"))
    wControlaSaldoCaixa = wControlaSaldoCaixa + (Format(grdMovimentoCaixa.TextMatrix(22, 3), "0.00"))
    wControlaSaldoCaixa = wControlaSaldoCaixa + (Format(grdMovimentoCaixa.TextMatrix(23, 3), "0.00"))
                        
    controlaNotasEmitidas
                        
' If wControlaSaldoCaixa > 0 Then
'    MsgBox "Existe Saldo em Modalidade Não Permitida para o Fechamento.", vbOKOnly
'    Unload Me
'    frmSangria.Show vbModal
' End If

End Sub

Private Sub controlaNotasEmitidas()

    Dim sql As String
    Dim rdoNotas As New ADODB.Recordset

    notaPedente = False

    sql = ("select count(*) as NFpedente from nfcapa where protocolo = " & GLB_CTR_Protocolo & " and nf = 0 and tiponota in ('V','T','E','S')")
    
    rdoNotas.CursorLocation = adUseClient
    rdoNotas.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
        
        If rdoNotas("NFpedente") > 0 Then
            notaPedente = True
        End If
    
    rdoNotas.Close
  
End Sub

Private Sub Form_Load()

  

  'left = 100
  'top = 2900
  Call AjustaTela(frmFechaCaixa)
  
  wControlaSaldoCaixa = 0
  
  sql = ("Select Max(CTr_DataInicial)as DataMov,Max(Ctr_Protocolo) as Seq " _
       & "from ControleCaixa where CTR_Supervisor <> 99 and CTr_NumeroCaixa = " & GLB_Caixa)
        rdoFechamentoGeral.CursorLocation = adUseClient
        rdoFechamentoGeral.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
  lblCabec.Caption = lblCabec & " " & Format(rdoFechamentoGeral("DataMov"), "dd/mm/yyyy")
  
  rdoFechamentoGeral.Close
  
  CarregaMovimento grdMovimentoCaixa, CStr(GLB_CTR_Protocolo)
  
'  grdMovimentoCaixa.Rows = 1
'  grdMovimentoCaixa.Rows = 1
'  grdMovimentoCaixa.AddItem "Dinheiro" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "70101"                    '1
'  grdMovimentoCaixa.AddItem "Cheque" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "70201"                '2                                                                              '2
'  grdMovimentoCaixa.AddItem "Cartões >>"                                                                           '3
'  grdMovimentoCaixa.AddItem "  Visa" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50301"                      '4
'  grdMovimentoCaixa.AddItem "  MasterCard" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50302"                '5
'  grdMovimentoCaixa.AddItem "  Amex" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50303"                      '6
'  grdMovimentoCaixa.AddItem "  BNDES" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50304"                     '7
'  grdMovimentoCaixa.AddItem "  T E F" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50203"                     '8
'  grdMovimentoCaixa.AddItem "Nota de Credito" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50701"             '9
'  grdMovimentoCaixa.AddItem ""                                                                                      '10
'  grdMovimentoCaixa.AddItem "*** TOTAL CAIXA" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "70204"             '11
'  grdMovimentoCaixa.AddItem ""                                                                                      '12
'  grdMovimentoCaixa.AddItem "Faturado" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                       '13'14
'  grdMovimentoCaixa.AddItem "Financiada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                     '14             '15
'  grdMovimentoCaixa.AddItem "Entrada Faturada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50502"            '15
'  grdMovimentoCaixa.AddItem "Entrada Financiada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50602"          '16
'  grdMovimentoCaixa.AddItem "Reforco de Caixa" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50801"            '17
'  grdMovimentoCaixa.AddItem "Garantia Estendida" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                             '18
'  grdMovimentoCaixa.AddItem ""                                                                                      '19
'  grdMovimentoCaixa.AddItem "*** TOTAL GERAL" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                '20              '17
'  grdMovimentoCaixa.AddItem ""                                                                                      '21
'  grdMovimentoCaixa.AddItem "*** Movimento NF"                                                                      '22
'  grdMovimentoCaixa.AddItem ""                                                                                      '23
'  grdMovimentoCaixa.AddItem "CF" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                             '24               '21
'  grdMovimentoCaixa.AddItem "NE" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                             '25                    '22
'  grdMovimentoCaixa.AddItem "D1" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                             '26                  '23
'  grdMovimentoCaixa.AddItem "S1" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                             '27                 '24
'  grdMovimentoCaixa.AddItem ""                                                                                      '28
'  grdMovimentoCaixa.AddItem "*** TOTAL NF" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                   '29                   '26
'  grdMovimentoCaixa.AddItem ""                                                                                      '30
'  grdMovimentoCaixa.AddItem "Transferencia Saida" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                            '31               '28
'  grdMovimentoCaixa.AddItem "Remessa Saida" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                  '32                  '29
'  grdMovimentoCaixa.AddItem "Devolucao Entrada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                              '33                 '30
'  grdMovimentoCaixa.AddItem ""                                                                                      '34
'  grdMovimentoCaixa.AddItem "CF Cancelada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                   '35
'  grdMovimentoCaixa.AddItem "NE Cancelada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                   '36
'  grdMovimentoCaixa.AddItem "D1 Cancelada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                   '37
'  grdMovimentoCaixa.AddItem "S1 Cancelada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                   '38
'  grdMovimentoCaixa.AddItem ""                                                                                      '39                                                                         '46
'  grdMovimentoCaixa.AddItem "** Saldo Anterior**"                                                                   '40
'  grdMovimentoCaixa.AddItem "Dinheiro" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "70101"                    '41
'  grdMovimentoCaixa.AddItem "Cheque" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "70201"                      '42
'  grdMovimentoCaixa.AddItem "Total do Saldo" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "00000"              '43
'  grdMovimentoCaixa.AddItem ""
'  wTotalSaldo = 0
'  wTotalSaldo_S = 0
'  Call CarregaMovimentocaixa
'  rdoFechamentoGeral.Close
'
'
  
  
  
End Sub


Private Sub ImprimeMovimento()
 
Screen.MousePointer = 11
Retorno = Bematech_FI_AbreRelatorioGerencialMFD("01")
     
     wSaldoFinalDinheiro = 0
     wSaldoFinalCheque = 0
     wSaldoFinalAVR = 0
     
     'wDataFinalFechamento = Format(grdMovimentosDisponiveis.TextMatrix(grdMovimentosDisponiveis.Row, 1), "DD/MM/YYYY")
     
     Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
               "                    IMPRESSAO                   " & _
               " Fechamento do Caixa " & Format(right(lblCabec.Caption, 10), "dd/mm/yyyy") & "         Loja " & Format(GLB_Loja, "000") & _
               "________________________________________________")
     
     For Idx = 1 To grdMovimentoCaixa.Rows - 6 Step 1
         If (Idx = 25) Or (Idx = 26) Or (Idx = 27) Or (Idx = 28) Or (Idx = 29) Or (Idx = 32) Or (Idx = 33) Or (Idx = 34) Or _
         (Idx = 37) Or (Idx = 38) Or (Idx = 39) Or (Idx = 40) Or (Idx = 41) Then
                          
         Retorno = Bematech_FI_UsaRelatorioGerencialMFD(left(grdMovimentoCaixa.TextMatrix(Idx, 0) & Space(23), 23) & _
               right(Space(20) & Format(grdMovimentoCaixa.TextMatrix(Idx, 1), "###,###,##0.00"), 20) & _
               right(Space(5) & grdMovimentoCaixa.TextMatrix(Idx, 2), 5))
         Else
         
          Retorno = Bematech_FI_UsaRelatorioGerencialMFD(left(grdMovimentoCaixa.TextMatrix(Idx, 0) & Space(23), 23) & _
               right(Space(20) & Format(grdMovimentoCaixa.TextMatrix(Idx, 1), "###,###,##0.00"), 20) & "     ")
         End If
     Next Idx
     
     
     wMovimentoPeriodo = Format(grdMovimentoCaixa.TextMatrix(13, 1), "###,###,##0.00")
     'wMovimentoPeriodo = wMovimentoPeriodo - wtotalGarantia
     'wMovimentoPeriodo = Format(wMovimentoPeriodo + CDbl(grdMovimentoCaixa.TextMatrix(46, 2)), "###,###,##0.00")
     wSaldoAnterior = Format(grdMovimentoCaixa.TextMatrix(45, 1), "###,###,##0.00")
     'wMovimentoPeriodo = (wMovimentoPeriodo - wSaldoAnterior)
     wSaldoFinalDinheiro = Format(grdMovimentoCaixa.TextMatrix(1, 3), "###,###,##0.00")
     wSaldoFinalDinheiro = (wSaldoFinalDinheiro + Format(grdMovimentoCaixa.TextMatrix(44, 3), "###,###,##0.00"))
     wSaldoFinalCheque = Format(grdMovimentoCaixa.TextMatrix(2, 3))
     wSaldoFinalCheque = (wSaldoFinalCheque + Format(grdMovimentoCaixa.TextMatrix(45, 3), "###,###,##0.00"))
     
     wTotalTransferencia = CDbl(grdMovimentoCaixa.TextMatrix(13, 2)) + CDbl(grdMovimentoCaixa.TextMatrix(46, 2))
     wTotalTransferencia = wTotalTransferencia + CDbl(grdMovimentoCaixa.TextMatrix(17, 2))
     wTotalTransferencia = wTotalTransferencia - wtotalGarantia
     
'     wMovimentoPeriodo = Format(grdMovimentoCaixa.TextMatrix(11, 1), "###,###,##0.00")
'     wSaldoAnterior = Format(grdMovimentoCaixa.TextMatrix(41, 1), "###,###,##0.00")
'     wMovimentoPeriodo = (wMovimentoPeriodo - wSaldoAnterior)
'     wSaldoFinalDinheiro = Format(grdMovimentoCaixa.TextMatrix(1, 3), "###,###,##0.00")
'     wSaldoFinalDinheiro = (wSaldoFinalDinheiro + Format(grdMovimentoCaixa.TextMatrix(41, 3), "###,###,##0.00"))
'     wSaldoFinalCheque = Format(grdMovimentoCaixa.TextMatrix(2, 3))
'     wSaldoFinalCheque = (wSaldoFinalCheque + Format(grdMovimentoCaixa.TextMatrix(42, 3), "###,###,##0.00"))
     'wSaldoFinalAVR = Format(grdMovimentoCaixa.TextMatrix(9, 3))
     'wSaldoFinalAVR = (wSaldoFinalAVR + Format(grdMovimentoCaixa.TextMatrix(15, 3), "###,###,##0.00"))
     
'        Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
'        "               MOVIMENTO DE CAIXA               " & _
'        "________________________________________________" & _
'        "SALDO ANTERIOR >> " & right(Space(30) & Format("", ""), 30) & _
'        "  DINHEIRO        " & right(Space(30) & Format(grdMovimentoCaixa.TextMatrix(41, 1), "###,###,##0.00"), 30) & _
'        "  CHEQUE          " & right(Space(30) & Format(grdMovimentoCaixa.TextMatrix(42, 1), "###,###,##0.00"), 30) & _
'        "  TOTAL           " & right(Space(30) & Format(wTotalSaldo, "###,###,##0.00"), 30) & _
'        "MOVIMENTO PERIODO " & right(Space(30) & Format(wMovimentoPeriodo, "###,###,##0.00"), 30) & _
'        "REFORCO           " & right(Space(30) & Format(grdMovimentoCaixa.TextMatrix(17, 1), "###,###,##0.00"), 30) & _
'        "TRANSFERENCIA NUM." & right(Space(30) & Format(grdMovimentoCaixa.TextMatrix(11, 2), "###,###,##0.00"), 30))
'
'
'      Retorno = Bematech_FI_UsaRelatorioGerencialMFD("SALDO DO CAIXA >>                               " & _
'        "  Dinheiro        " & right(Space(30) & Format(wSaldoFinalDinheiro, "###,###,##0.00"), 30) & _
'        "  Cheque          " & right(Space(30) & Format(wSaldoFinalCheque, "###,###,##0.00"), 30) & _
'        "  SALDO FINAL     " & right(Space(30) & Format((wSaldoFinalDinheiro + wSaldoFinalCheque), "###,###,##0.00"), 30) & _
'        "________________________________________________")
'
'
'      Retorno = Bematech_FI_UsaRelatorioGerencialMFD(left("Caixa Nro.   " & Trim(frmControlaCaixa.lblNroCaixa.Caption) & Space(48), 48) & _
'         left("Operador     " & Trim(txtOperador.Text) & Space(48), 48) & _
'         left("Data Inicial " & Format(Date, "dd/mm/yyyy") & Space(48), 48) & _
'         left("Data Final   " & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS") & Space(48), 48) & _
'         left("Protocolo    " & GLB_CTR_Protocolo & Space(48), 48))
'
'
'       Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
'         "                                                " & _
'         "                                                " & _
'         "                                                " & _
'         " ______________________  ______________________ " & _
'         " " & left(Trim(txtOperador.Text) & Space(22), 22) & "  " & left(Trim(txtSupervisor.Text) & Space(23), 23) & _
'         "                                                " & _
'         "                                                ")
     
     
     Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
        "               MOVIMENTO DE CAIXA               " & _
        "________________________________________________" & _
        "SALDO ANTERIOR >> " & right(Space(30) & Format("", ""), 30) & _
        "  DINHEIRO        " & right(Space(30) & Format(grdMovimentoCaixa.TextMatrix(44, 1), "###,###,##0.00"), 30) & _
        "  CHEQUE          " & right(Space(30) & Format(grdMovimentoCaixa.TextMatrix(45, 1), "###,###,##0.00"), 30) & _
        "  TOTAL           " & right(Space(30) & Format(wTotalSaldo, "###,###,##0.00"), 30) & _
        "MOVIMENTO PERIODO " & right(Space(30) & Format(wMovimentoPeriodo, "###,###,##0.00"), 30) & _
        "REFORCO           " & right(Space(30) & Format(grdMovimentoCaixa.TextMatrix(17, 1), "###,###,##0.00"), 30) & _
        "TRANSFERENCIA NUM." & right(Space(30) & Format(wTotalTransferencia, "###,###,##0.00"), 30) & _
        "GARANTIA ESTEN.   " & right(Space(30) & Format(wtotalGarantia, "###,###,##0.00"), 30))

 
          
      Retorno = Bematech_FI_UsaRelatorioGerencialMFD("SALDO DO CAIXA >>                               " & _
        "  Dinheiro        " & right(Space(30) & Format(wSaldoFinalDinheiro, "###,###,##0.00"), 30) & _
        "  Cheque          " & right(Space(30) & Format(wSaldoFinalCheque, "###,###,##0.00"), 30) & _
        "  SALDO FINAL     " & right(Space(30) & Format((wSaldoFinalDinheiro + wSaldoFinalCheque), "###,###,##0.00"), 30) & _
        "________________________________________________")
        
 
      Retorno = Bematech_FI_UsaRelatorioGerencialMFD(left("Caixa Nro.   " & Trim(frmControlaCaixa.lblNroCaixa.Caption) & Space(48), 48) & _
         left("Operador     " & Trim(txtOperador.text) & Space(48), 48) & _
         left("Data Inicial " & Format(Date, "dd/mm/yyyy") & Space(48), 48) & _
         left("Data Final   " & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS") & Space(48), 48) & _
         left("Protocolo    " & GLB_CTR_Protocolo & Space(48), 48))
                    
                    
       Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
         "                                                " & _
         "                                                " & _
         "                                                " & _
         " ______________________  ______________________ " & _
         " " & left(Trim(txtOperador.text) & Space(22), 22) & "  " & "SUPERVISOR" & _
         "                                                " & _
         "                                                ")
     
 
        Retorno = Bematech_FI_FechaRelatorioGerencial()
 
      Screen.MousePointer = 0
 
     
End Sub


Private Sub ImprimeMovimento00()
Dim wImprime As String
Dim wImprime2 As String
Dim wTotal00 As Double
wTotal00 = 0
 
sql = ("Select Serie, sum(totalnota) as TotalSerieNota, count(Serie) as QtdeSerie from nfcapa Where ecf = " & GLB_ECF & "" _
     & " and  TipoNota = 'V' and Serie = '00' and  DataEmi = '" & Format(Date, "yyyy/mm/dd") _
     & "' " & " and Protocolo = " & GLB_CTR_Protocolo & " group by Serie ")
     rdoCapa.CursorLocation = adUseClient
     rdoCapa.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
     
If Not rdoCapa.EOF Then
   wTotal00 = rdoCapa("TotalSerieNota")
End If
 
rdoCapa.Close
    Screen.MousePointer = 11
         Retorno = Bematech_FI_AbreRelatorioGerencialMFD("01")
    
          Retorno = Bematech_FI_UsaRelatorioGerencialMFD("                                                " & _
                    "                     CONTROLE X                 " & _
                    " Loja " & left(GLB_Loja & Space(4), 4) & Space(28) & Format(Date, "dd/mm/yyyy") & _
                    "                                                " & _
                    " Valor Total:                     " & right(Space(10) & Format(wTotal00, "###,###,##0.00"), 14) & _
                    "________________________________________________" & _
                    "                                                " & _
                    left("Caixa Nro.   " & frmControlaCaixa.lblNroCaixa.Caption & Space(48), 48) & _
                    left("Operador     " & txtOperador.text & Space(48), 48) & _
                    left("Data Inicial " & Format(Date, "dd/mm/yyyy") & Space(48), 48) & _
                    left("Data Final   " & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS") & Space(48), 48) & _
                    left("Protocolo    " & GLB_CTR_Protocolo & Space(48), 48))
                    
                    
    Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
                    "                                                " & _
                    "                                                " & _
                    "                                                " & _
                    " ______________________   _____________________ " & _
                    " " & "       SUPERVISOR             " & "   " & left(txtOperador.text & Space(22), 22) & " " & _
                    "                                                " & _
                    "                                                ")
        Retorno = Bematech_FI_FechaRelatorioGerencial()
 
      Screen.MousePointer = 0
 
 
End Sub
 
 
Private Sub ImprimeTransfNumerario()
     
    Screen.MousePointer = 11
    Retorno = Bematech_FI_AbreRelatorioGerencialMFD("01")
 
    Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
                   "TRANSFERENCIA DE NUMERARIO" & right(Space(22) & "Nro.  " & GLB_CTR_Protocolo, 22) & _
                   left("Loja " & Format(GLB_Loja, "000") & Space(10), 10) & right(Space(38) & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS"), 38) & _
                   "________________________________________________" & _
                   "    USO INTERNO          SEM VALOR COMERCIAL    ")
   
   Retorno = Bematech_FI_UsaRelatorioGerencialMFD("------------------------------------------------" & _
                   "DINHEIRO /P TESOU." & right(Space(30) & Format(wTNDinheiro, "###,###,##0.00"), 30) & _
                   "CHEQUE /P TESOU.  " & right(Space(30) & Format(wTNCheque, "###,###,##0.00"), 30) & _
                   "VISA              " & right(Space(30) & Format(wTNVisa, "###,###,##0.00"), 30) & _
                   "MASTERCARD        " & right(Space(30) & Format(wTNRedecard, "###,###,##0.00"), 30) & _
                   "AMEX              " & right(Space(30) & Format(wTNAmex, "###,###,##0.00"), 30) & _
                   "BNDS              " & right(Space(30) & Format(wTNBNDES, "###,###,##0.00"), 30) & _
                   "REDE SHOP         " & right(Space(30) & Format(wTNRedeShop, "###,###,##0.00"), 30) & _
                   "VISA ELEC.        " & right(Space(30) & Format(wTNVisaEletron, "###,###,##0.00"), 30))
                   
   Retorno = Bematech_FI_UsaRelatorioGerencialMFD( _
                   "HIPERCARD         " & right(Space(30) & Format(wTNHiperCard, "###,###,##0.00"), 30) & _
                   "DEPOSITO          " & right(Space(30) & Format(wTNDeposito, "###,###,##0.00"), 30) & _
                   "NOTA CREDITO      " & right(Space(30) & Format(wTNNotaCredito, "###,###,##0.00"), 30) & _
                   "DESPESA CONDUCAO  " & right(Space(30) & Format(wTNConducao, "###,###,##0.00"), 30) & _
                   "DESPESA LOJA      " & right(Space(30) & Format(wTNDespLoja, "###,###,##0.00"), 30) & _
                   "OUTRAS DESPESAS   " & right(Space(30) & Format(wTNOutros, "###,###,##0.00"), 30) & _
                   "                                                " & _
                   "TOTAL             " & right(Space(30) & Format(wTNTotal, "###,###,##0.00"), 30))
                   
    Retorno = Bematech_FI_UsaRelatorioGerencialMFD( _
                   "                  " & right(Space(30) & "", 30) & _
                   "GARANTIA ESTEN.   " & right(Space(30) & Format(wtotalGarantia, "###,###,##0.00"), 30) & _
                   "ENTRADA FINANCIADA" & right(Space(30) & Format(wTNFinanciado, "###,###,##0.00"), 30) & _
                   "ENTRADA FATURADA  " & right(Space(30) & Format(wTNFaturado, "###,###,##0.00"), 30))
                   
    Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
                    "                                                " & _
                    "                                                " & _
                    "                                                " & _
                    " ______________________   _____________________ " & _
                    " " & "       SUPERVISOR             " & "   " & left(txtOperador.text & Space(22), 22) & " " & _
                    "                                                " & _
                    "                                                ")
        Retorno = Bematech_FI_FechaRelatorioGerencial()
 
      Screen.MousePointer = 0
 
End Sub


'Private Sub CarregaMovimentocaixa()
' wTNNotaCredito = 0
' wTotalNf = 0
' wTotalFatFin = 0
' grdMovimentoCaixa.Row = 1
'
' sql = ("Select Max(CTr_DataInicial)as DataMov,Max(Ctr_Protocolo) as Seq from ControleCaixa where CTR_Supervisor <> 99 and CTr_NumeroCaixa = " & GLB_Caixa & "")
'       rdoDataFechamentoRetaguarda.CursorLocation = adUseClient
'       rdoDataFechamentoRetaguarda.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
'
' sql = ("select mc_Grupo,sum(MC_Valor) as TotalModalidade,Count(*) as Quantidade from movimentocaixa" _
'       & " Where MC_NumeroEcf = " & GLB_ECF & " and MC_NroCaixa=" & GLB_Caixa & " and MC_Protocolo=" & GLB_CTR_Protocolo _
'       & " and MC_Data >='" & Format(rdoDataFechamentoRetaguarda("DataMov"), "yyyy/mm/dd") & "' and  MC_Serie <> '00' and (MC_Grupo like '10%' or MC_Grupo like '11%'" _
'       & " or MC_Grupo like '50%' or MC_Grupo like '20%') and mc_tiponota  <> 'C' group by mc_grupo")
'       rdoFormaPagamento.CursorLocation = adUseClient
'       rdoFormaPagamento.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
'
'       rdoDataFechamentoRetaguarda.Close
'
'If Not rdoFormaPagamento.EOF Then
'     Do While Not rdoFormaPagamento.EOF
'        If rdoFormaPagamento("MC_Grupo") = "10101" Then
'           grdMovimentoCaixa.TextMatrix(1, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "10201" Then
'           grdMovimentoCaixa.TextMatrix(2, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("mc_grupo") = "10301" Then
'           grdMovimentoCaixa.TextMatrix(4, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "10302" Then
'           grdMovimentoCaixa.TextMatrix(5, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "10303" Then
'           grdMovimentoCaixa.TextMatrix(6, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "10304" Then
'           grdMovimentoCaixa.TextMatrix(7, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "10203" Then
'           grdMovimentoCaixa.TextMatrix(8, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "10701" Then
'           grdMovimentoCaixa.TextMatrix(9, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
'           'wTNNotaCredito = wTNNotaCredito + rdoFormaPagamento("TotalModalidade")
'        'ElseIf rdoFormaPagamento("MC_Grupo") = "10204" Then 'AVR
'           'grdMovimentoCaixa.TextMatrix(9, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           'wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "11004" Then
'           grdMovimentoCaixa.TextMatrix(15, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
'           wTotalEntrada = (wTotalEntrada + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "11005" Then
'           grdMovimentoCaixa.TextMatrix(16, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           'wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
'           wTotalEntrada = (wTotalEntrada + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "10801" Then
'           grdMovimentoCaixa.TextMatrix(17, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           'wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
'        'ElseIf rdoFormaPagamento("MC_Grupo") = "11008" Then 'AVR
'           'grdMovimentoCaixa.TextMatrix(15, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           'wTotalSaldo = (wTotalSaldo + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "11006" Then 'SALDO ANTERIOR
'           grdMovimentoCaixa.TextMatrix(41, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wTotalSaldo = (wTotalSaldo + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "11007" Then 'SALDO ANTERIOR
'           grdMovimentoCaixa.TextMatrix(42, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wTotalSaldo = (wTotalSaldo + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "10501" Then
'           grdMovimentoCaixa.TextMatrix(13, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wTotalFatFin = (wTotalFatFin + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "10601" Then
'           grdMovimentoCaixa.TextMatrix(14, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wTotalFatFin = (wTotalFatFin + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "11009" Then
'           grdMovimentoCaixa.TextMatrix(18, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wtotalGarantia = (wtotalGarantia + rdoFormaPagamento("TotalModalidade"))
'           'wSubTotal = (wSubTotal - rdoFormaPagamento("TotalModalidade"))
'           'wSubTotal_S = (wSubTotal_S - rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "20101" Then
'           grdMovimentoCaixa.TextMatrix(24, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           grdMovimentoCaixa.TextMatrix(24, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
'           wTotalNf = (wTotalNf + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "20102" Then
'           grdMovimentoCaixa.TextMatrix(25, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           grdMovimentoCaixa.TextMatrix(25, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
'           wTotalNf = (wTotalNf + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "20107" Then
'           grdMovimentoCaixa.TextMatrix(26, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           grdMovimentoCaixa.TextMatrix(26, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
'           wTotalNf = (wTotalNf + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "20108" Then
'           grdMovimentoCaixa.TextMatrix(27, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           grdMovimentoCaixa.TextMatrix(27, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
'           wTotalNf = (wTotalNf + rdoFormaPagamento("TotalModalidade"))
'        'ElseIf rdoFormaPagamento("MC_Grupo") = "20111" Then
'        '   grdMovimentoCaixa.TextMatrix(33, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'        '   grdMovimentoCaixa.TextMatrix(33, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
'        '   wTotalNf = (wTotalNf + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "20109" Then
'           grdMovimentoCaixa.TextMatrix(31, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           grdMovimentoCaixa.TextMatrix(31, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
'        ElseIf rdoFormaPagamento("MC_Grupo") = "20110" Then
'           grdMovimentoCaixa.TextMatrix(32, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           grdMovimentoCaixa.TextMatrix(32, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
'        ElseIf rdoFormaPagamento("MC_Grupo") = "20201" Then
'           grdMovimentoCaixa.TextMatrix(33, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           grdMovimentoCaixa.TextMatrix(33, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
'
'        ElseIf rdoFormaPagamento("MC_Grupo") = "50101" Then
'           grdMovimentoCaixa.TextMatrix(1, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "50201" Then
'           grdMovimentoCaixa.TextMatrix(2, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("mc_grupo") = "50301" Then
'           grdMovimentoCaixa.TextMatrix(4, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "50302" Then
'           grdMovimentoCaixa.TextMatrix(5, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "50303" Then
'           grdMovimentoCaixa.TextMatrix(6, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "50304" Then
'           grdMovimentoCaixa.TextMatrix(7, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "50203" Then
'           grdMovimentoCaixa.TextMatrix(8, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "50701" Then
'           grdMovimentoCaixa.TextMatrix(9, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
'        'ElseIf rdoFormaPagamento("MC_Grupo") = "50204" Then 'AVR
'           'grdMovimentoCaixa.TextMatrix(9, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           'wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "50502" Then
'           grdMovimentoCaixa.TextMatrix(15, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "50602" Then
'           grdMovimentoCaixa.TextMatrix(16, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           'wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "50801" Then
'           grdMovimentoCaixa.TextMatrix(17, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
'        'ElseIf rdoFormaPagamento("MC_Grupo") = "50804" Then
'           'grdMovimentoCaixa.TextMatrix(15, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           'wTotalSaldo_S = (wTotalSaldo_S + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "50806" Then
'           grdMovimentoCaixa.TextMatrix(41, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wTotalSaldo_S = (wTotalSaldo_S + rdoFormaPagamento("TotalModalidade"))
'        ElseIf rdoFormaPagamento("MC_Grupo") = "50807" Then
'           grdMovimentoCaixa.TextMatrix(42, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           wTotalSaldo_S = (wTotalSaldo_S + rdoFormaPagamento("TotalModalidade"))
'
'        End If
'       rdoFormaPagamento.MoveNext
'     Loop
'
'
'     grdMovimentoCaixa.TextMatrix(1, 1) = Format((grdMovimentoCaixa.TextMatrix(1, 1) - wtotalGarantia), "###,###,###,##0.00")
'     grdMovimentoCaixa.TextMatrix(1, 2) = Format((grdMovimentoCaixa.TextMatrix(1, 2)), "###,###,###,##0.00")
'
'     grdMovimentoCaixa.TextMatrix(1, 3) = Format((grdMovimentoCaixa.TextMatrix(1, 1) - grdMovimentoCaixa.TextMatrix(1, 2)), "###,###,###,##0.00")
'     grdMovimentoCaixa.TextMatrix(2, 3) = Format((grdMovimentoCaixa.TextMatrix(2, 1) - grdMovimentoCaixa.TextMatrix(2, 2)), "###,###,###,##0.00")
'     grdMovimentoCaixa.TextMatrix(4, 3) = Format((grdMovimentoCaixa.TextMatrix(4, 1) - grdMovimentoCaixa.TextMatrix(4, 2)), "###,###,###,##0.00")
'     grdMovimentoCaixa.TextMatrix(5, 3) = Format((grdMovimentoCaixa.TextMatrix(5, 1) - grdMovimentoCaixa.TextMatrix(5, 2)), "###,###,###,##0.00")
'     grdMovimentoCaixa.TextMatrix(6, 3) = Format((grdMovimentoCaixa.TextMatrix(6, 1) - grdMovimentoCaixa.TextMatrix(6, 2)), "###,###,###,##0.00")
'     grdMovimentoCaixa.TextMatrix(7, 3) = Format((grdMovimentoCaixa.TextMatrix(7, 1) - grdMovimentoCaixa.TextMatrix(7, 2)), "###,###,###,##0.00")
'     grdMovimentoCaixa.TextMatrix(8, 3) = Format((grdMovimentoCaixa.TextMatrix(8, 1) - grdMovimentoCaixa.TextMatrix(8, 2)), "###,###,###,##0.00")
'     grdMovimentoCaixa.TextMatrix(9, 3) = Format((grdMovimentoCaixa.TextMatrix(9, 1) - grdMovimentoCaixa.TextMatrix(9, 2)), "###,###,###,##0.00")
'     'grdMovimentoCaixa.TextMatrix(9, 3) = Format((grdMovimentoCaixa.TextMatrix(9, 1) - grdMovimentoCaixa.TextMatrix(9, 2)), "###,###,###,##0.00")
'     grdMovimentoCaixa.TextMatrix(15, 3) = Format((grdMovimentoCaixa.TextMatrix(15, 1) - grdMovimentoCaixa.TextMatrix(15, 2)), "###,###,###,##0.00")
'     grdMovimentoCaixa.TextMatrix(16, 3) = Format((grdMovimentoCaixa.TextMatrix(16, 1) - grdMovimentoCaixa.TextMatrix(16, 2)), "###,###,###,##0.00")
'     grdMovimentoCaixa.TextMatrix(17, 3) = Format((grdMovimentoCaixa.TextMatrix(17, 1) - grdMovimentoCaixa.TextMatrix(17, 2)), "###,###,###,##0.00")
'     grdMovimentoCaixa.TextMatrix(41, 3) = Format((grdMovimentoCaixa.TextMatrix(41, 1) - grdMovimentoCaixa.TextMatrix(41, 2)), "###,###,###,##0.00")
'     grdMovimentoCaixa.TextMatrix(42, 3) = Format((grdMovimentoCaixa.TextMatrix(42, 1) - grdMovimentoCaixa.TextMatrix(42, 2)), "###,###,###,##0.00")
'
'     'grdMovimentoCaixa.TextMatrix(18, 1) = Format(wTotalSaldo, "###,###,###,##0.00")
'     'grdMovimentoCaixa.TextMatrix(18, 2) = Format(wTotalSaldo_S, "###,###,###,##0.00")
'     'grdMovimentoCaixa.TextMatrix(18, 3) = Format((wTotalSaldo - wTotalSaldo_S), "###,###,###,##0.00")
'
'
'    'wSubTotal = grdMovimentoCaixa.TextMatrix(11, 1)
'
'     grdMovimentoCaixa.TextMatrix(43, 1) = Format(wTotalSaldo, "###,###,###,##0.00")
'     grdMovimentoCaixa.TextMatrix(43, 2) = Format(wTotalSaldo_S, "###,###,###,##0.00")
'     grdMovimentoCaixa.TextMatrix(43, 3) = Format((wTotalSaldo - wTotalSaldo_S), "###,###,###,##0.00")
'
'     grdMovimentoCaixa.TextMatrix(11, 1) = Format(((wSubTotal + wTotalSaldo) - wtotalGarantia), "###,###,###,##0.00")
'     grdMovimentoCaixa.TextMatrix(11, 2) = Format((wSubTotal_S + wTotalSaldo_S), "###,###,###,##0.00")
'     grdMovimentoCaixa.TextMatrix(11, 3) = Format((((wSubTotal + wTotalSaldo)) - ((wSubTotal_S + wTotalSaldo_S))), "###,###,###,##0.00")
'     grdMovimentoCaixa.TextMatrix(29, 1) = Format(wTotalNf, "###,###,###,##0.00")
'     grdMovimentoCaixa.TextMatrix(20, 1) = Format(((wTotalFatFin + wSubTotal)), "###,###,###,##0.00")
'
'  End If
'  rdoFormaPagamento.Close
'
'    sql = ("select mc_Grupo,sum(MC_Valor) as TotalModalidade,Count(*) as Quantidade from movimentocaixa" _
'       & " Where MC_NroCaixa=" & GLB_Caixa & " and MC_Protocolo=" & GLB_CTR_Protocolo _
'       & " and  MC_Serie <> '00' and MC_tiponota = 'C' and MC_Grupo like '20%' group by mc_grupo")
'       rdoFormaPagamento.CursorLocation = adUseClient
'       rdoFormaPagamento.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
'
'  If Not rdoFormaPagamento.EOF Then
'     Do While Not rdoFormaPagamento.EOF
'        If rdoFormaPagamento("MC_Grupo") = "20101" Then
'           grdMovimentoCaixa.TextMatrix(35, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           grdMovimentoCaixa.TextMatrix(35, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
'        ElseIf rdoFormaPagamento("MC_Grupo") = "20102" Then
'           grdMovimentoCaixa.TextMatrix(36, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'           grdMovimentoCaixa.TextMatrix(36, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
'        ElseIf rdoFormaPagamento("MC_Grupo") = "20107" Then
'           grdMovimentoCaixa.TextMatrix(37, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00") '?
'           grdMovimentoCaixa.TextMatrix(37, 2) = Format(rdoFormaPagamento("Quantidade"), "0") '?
'        ElseIf rdoFormaPagamento("MC_Grupo") = "20108" Then
'           grdMovimentoCaixa.TextMatrix(38, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00") '?
'           grdMovimentoCaixa.TextMatrix(38, 2) = Format(rdoFormaPagamento("Quantidade"), "0") '?
'        'ElseIf rdoFormaPagamento("MC_Grupo") = "20111" Then
'        '   grdMovimentoCaixa.TextMatrix(44, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
'        '   grdMovimentoCaixa.TextMatrix(44, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
'        End If
'        rdoFormaPagamento.MoveNext
'     Loop
'  End If
'
'  rdoFormaPagamento.Close
'
'  wSubTotal = 0
'  wSubTotal_S = 0
'
'End Sub

'Private Sub CarregaValoresTransfNumerario()
'
'    Dim wchkfechamento As Boolean
'
'    wTNDinheiro = 0
'    wTNCheque = 0
'    wTNVisa = 0
'    wTNRedecard = 0
'    wTNAmex = 0
'    wTNHiperCard = 0
'    wTNBNDES = 0
'    wTNVisaEletron = 0
'    wTNRedeShop = 0
'    wTNDeposito = 0
'    wTNNotaCredito = 0
'    wTNConducao = 0
'    wTNDespLoja = 0
'    wTNOutros = 0
'    wTNTotal = 0
'
'   If wchkfechamento = True Then
'
'      sql = ("SELECT MC_GrupoAuxiliar,MO_Descricao,SUM(MC_Valor) as Valor FROM MOVIMENTOCAIXA,MODALIDADE WHERE Mo_GRUPO=MC_GrupoAuxiliar" _
'        & " AND MC_GRUPOAUXILIAR LIKE '30%' and MC_DATA between '" & Format(datageral, "yyyy/mm/dd") & " 00:00:00.000' and '" _
'       & Format(datageral, "yyyy/mm/dd") & " 23:59:59:000" _
'        & "' GROUP BY MC_GrupoAuxiliar,MO_DESCRICAO order by MC_GrupoAuxiliar")
'       rdoTransfNumerario.CursorLocation = adUseClient
'       rdoTransfNumerario.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
'
'   Else
'
'   sql = ("SELECT MC_GrupoAuxiliar,MO_Descricao,SUM(MC_Valor) as Valor FROM MOVIMENTOCAIXA,MODALIDADE WHERE Mo_GRUPO=MC_GrupoAuxiliar" _
'        & " AND MC_PROTOCOLO=" & GLB_CTR_Protocolo & " AND MC_GRUPOAUXILIAR LIKE '30%'" _
'        & "GROUP BY MC_GrupoAuxiliar,MO_DESCRICAO order by MC_GrupoAuxiliar")
'       rdoTransfNumerario.CursorLocation = adUseClient
'       rdoTransfNumerario.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
'       End If
'
''  If Not rdoTransfNumerario.EOF Then
''     Do While Not rdoTransfNumerario.EOF
''        If rdoTransfNumerario("MC_GrupoAuxiliar") = 30101 Then
''           wTNDeposito = rdoTransfNumerario("Valor")
''           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
''        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30201 Then
''           wTNDeposito = rdoTransfNumerario("Valor")
''           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
''        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30106 Then
''           wTNDinheiro = rdoTransfNumerario("Valor")
''           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
''        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30107 Then
''           wTNCheque = rdoTransfNumerario("Valor")
''           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
''        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30203 Then
''           wTNVisaEletron = rdoTransfNumerario("Valor")
''           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
''        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30301 Then
''           wTNVisa = rdoTransfNumerario("Valor")
''           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
''        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30302 Then
''           wTNRedecard = rdoTransfNumerario("Valor")
''           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
''        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30303 Then
''           wTNAmex = rdoTransfNumerario("Valor")
''           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
''        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30304 Then
''           wTNBNDES = rdoTransfNumerario("Valor")
''           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
''        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30103 Then
''           wTNConducao = rdoTransfNumerario("Valor")
''           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
''        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30104 Then
''           wTNDespLoja = rdoTransfNumerario("Valor")
''           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
''        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = "30701" Then
''           wTNNotaCredito = (wTNNotaCredito + rdoTransfNumerario("Valor"))
''           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
''        End If
''
''
''       rdoTransfNumerario.MoveNext
''
''     Loop
'' End If
'
'
'  If Not rdoTransfNumerario.EOF Then
'     Do While Not rdoTransfNumerario.EOF
'        If rdoTransfNumerario("MC_GrupoAuxiliar") = 30101 Then
'           wTNDeposito = rdoTransfNumerario("Valor")
'           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
'        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30201 Then
'           wTNDeposito = rdoTransfNumerario("Valor")
'           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
'        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30106 Then
'           wTNDinheiro = rdoTransfNumerario("Valor")
'           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
'        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30107 Then
'           wTNCheque = rdoTransfNumerario("Valor")
'           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
'        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30203 Then
'           wTNRedeShop = rdoTransfNumerario("Valor")
'           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
'        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30206 Then
'           wTNVisaEletron = rdoTransfNumerario("Valor")
'           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
'        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30301 Then
'           wTNVisa = rdoTransfNumerario("Valor")
'           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
'        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30302 Then
'           wTNRedecard = rdoTransfNumerario("Valor")
'           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
'        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30303 Then
'           wTNAmex = rdoTransfNumerario("Valor")
'           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
'        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30304 Then
'           wTNBNDES = rdoTransfNumerario("Valor")
'           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
'        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30103 Then
'           wTNConducao = rdoTransfNumerario("Valor")
'           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
'        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30104 Then
'           wTNDespLoja = rdoTransfNumerario("Valor")
'           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
'        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30205 Then
'           wTNHiperCard = rdoTransfNumerario("Valor")
'           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
'        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = 30009 Then
'            wtotalGarantia = rdoTransfNumerario("Valor")
'           'wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
'        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = "30701" Then
'           wTNNotaCredito = (wTNNotaCredito + rdoTransfNumerario("Valor"))
'           wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
'        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = "30502" Then
'           wTNFaturado = (wTNFaturado + rdoTransfNumerario("Valor"))
'        ElseIf rdoTransfNumerario("MC_GrupoAuxiliar") = "30602" Then
'           wTNFinanciado = (wTNFinanciado + rdoTransfNumerario("Valor"))
'           'wTNTotal = (wTNTotal + rdoTransfNumerario("Valor"))
'        End If
'
'
'       rdoTransfNumerario.MoveNext
'
'     Loop
' End If
'
'  rdoTransfNumerario.Close
'End Sub

Private Sub GravaSaldoCaixa()

  sql = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja, MC_Data, MC_Grupo,MC_Subgrupo, MC_Documento,MC_Serie," _
               & "MC_Valor, MC_banco, MC_Agencia,MC_Contacorrente, MC_bomPara, MC_Parcelas,MC_Remessa,MC_SituacaoEnvio," _
               & "MC_NroCaixa,MC_Protocolo,MC_Pedido,MC_DataProcesso,MC_TipoNota)" _
               & " values(" & GLB_ECF & ",'" & GLB_USU_Codigo & "','" & GLB_Loja & "','" _
               & Format(wData, "yyyy/mm/dd") & "'," & grdMovimentoCaixa.TextMatrix(Idx, 4) _
               & ",'',0,'SC'," _
               & ConverteVirgula(Format(wSaldoNovo, "##,###0.00")) & ", " _
               & "0,0,0,0,0,9,'A','" & GLB_Caixa & "'," & GLB_CTR_Protocolo & ",'0','" & Format(wData, "yyyy/mm/dd") & "','V')"
      rdoCNLoja.Execute (sql)
    
End Sub
Private Sub FechaCaixa()

    If ChkFechamento.Value = 1 Then
        
        sql = "Update ControleCaixa set CTR_SaldoFinal = " & ConverteVirgula(Format(grdMovimentoCaixa.TextMatrix(15, 3), "##,###0.00")) _
        & ",CTR_SituacaoCaixa='F',CTR_DataFinal=GetDate()" _
        & " Where CTR_SituacaoCaixa = 'A'"
        
        sql = sql & vbNewLine & "update estoqueloja set el_estoqueanterior = el_estoque"
        rdoCNLoja.Execute sql, rdExecDirect
        
    Else
    
        sql = "Update ControleCaixa set CTR_SaldoFinal =" & ConverteVirgula(Format(grdMovimentoCaixa.TextMatrix(15, 3), "##,###0.00")) _
        & ",CTR_SituacaoCaixa='F',CTR_DataFinal=GetDate()" _
        & " Where CTR_Supervisor <> 99 and CTR_SituacaoCaixa = 'A' and CTR_NumeroCaixa ='" & GLB_Caixa & "'"
        
    End If

    rdoCNLoja.Execute sql, rdExecDirect

'  sql = ("Select Max(CTr_DataInicial)as DataMov,Max(Ctr_Protocolo) as Seq " _
'       & "from ControleCaixa where CTR_Supervisor <> 99 and CTr_NumeroCaixa = " & GLB_Caixa)
'        rdoDataFechamentoRetaguarda.CursorLocation = adUseClient
'        rdoDataFechamentoRetaguarda.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
   
End Sub
   

Private Sub grdMovimentoCaixa_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
       Unload Me
End If
End Sub

Private Sub txtOperador_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
       Unload Me
End If
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
       Unload Me
End If
If KeyAscii = 13 Then
 
 If wControlaSaldoCaixa > 0 Then
    MsgBox "Existe Saldo em Modalidade Não Permitida para o Fechamento.", vbInformation
    Unload Me
    If RsDados.State = 1 Then
        RsDados.Close
     End If
    frmSangria.Show vbModal
    Exit Sub
 End If
 
 If notaPedente = True Then
    MsgBox "Existe nota(s) fiscal(s) ainda não autorizada(s) que precisa ser emitida(s) antes do fechamento", vbInformation
    Unload Me
    frmEmissaoNFe.Show vbModal
    Exit Sub
 End If
 
 If RsDados.State = 1 Then
    RsDados.Close
 End If

   sql = ("Select * from UsuarioCaixa where USU_Nome ='" & txtOperador.text & "' and USU_TipoUsuario='O'")
   RsDados.CursorLocation = adUseClient
   RsDados.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
   If RsDados.EOF Then
      MsgBox "Operador não Cadastrado", vbCritical, "Aviso"
      RsDados.Close
      Exit Sub
   Else
      GLB_USU_Nome = Trim(RsDados("USU_Nome"))
      GLB_USU_Codigo = Trim(RsDados("USU_Codigo"))

      If RTrim(RsDados("USU_Senha")) <> txtSenha.text Then
         MsgBox "Senha do Operador não Cadastrado", vbCritical, "Aviso"
         RsDados.Close
         Exit Sub
      Else
         RsDados.Close
         'sql = ("Select * from UsuarioCaixa where USU_Nome ='" & txtSupervisor.Text & "' and USU_TipoUsuario='S'")
         'RsDados.CursorLocation = adUseClient
         'RsDados.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
         'If RsDados.EOF Then
         '   MsgBox "Supervisor não Cadastrado", vbCritical, "Aviso"
         '   RsDados.Close
         '   Exit Sub
         'Else
          '  If RTrim(RsDados("USU_Senha")) <> txtSenhaSupervisor.Text Then
          '     MsgBox "Senha do Supervisor não Cadastrado", vbCritical, "Aviso"
          '     RsDados.Close
          '     Exit Sub
          '  Else
               Call FechaCaixaOK
               If ChkFechamento.Value = 1 Then
                    
                     Dim rdoProtocolos As New ADODB.Recordset
                     Dim wProtocolos As String
                     Dim dataInicialGeral As String
                     Dim dataFinalGeral As String
                     Dim horaInicialGeral As String
                     Dim horaFinalGeral As String
                     
                     sql = "select ctr_protocolo as protocolo from controlecaixa " & vbNewLine & _
                           "where CTR_Supervisor <> 99 and convert(char(10),CTR_DataInicial,111) = '" & Format(GLB_DataInicial, "yyyy/mm/dd") & "'"
                     rdoProtocolos.CursorLocation = adUseClient
                     rdoProtocolos.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                         
                     wProtocolos = ""
                     Do While Not rdoProtocolos.EOF
                         wProtocolos = wProtocolos & rdoProtocolos("protocolo") & ", "
                         rdoProtocolos.MoveNext
                     Loop
                     

                     wProtocolos = left(wProtocolos, Len(wProtocolos) - 2)
                     rdoProtocolos.Close
                     CarregaMovimento grdMovimentoCaixa, wProtocolos

                
                     Call NOVO_ImprimeMovimento(grdMovimentoCaixa, "FECHAMENTO DE CAIXA GERAL", _
                                               txtOperador.text, "GERAL", _
                                               Date, Time, _
                                               Date, Time, _
                                               wProtocolos)
                    
                     Call NOVO_ImprimeTransfNumerario(grdMovimentoCaixa, "FECHAMENTO DE CAIXA GERAL", _
                                               txtOperador.text, "GERAL", _
                                               Date, Time, _
                                               Date, Time, _
                                               wProtocolos)
                
                  
                     Unload Me
                     Unload frmControlaCaixa

                End If
                End
      End If
   End If
End If
End Sub

Private Sub txtSenhaSupervisor_KeyPress(KeyAscii As Integer)

End Sub

Private Sub txtSupervisor_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
       Unload Me
End If
End Sub
'----------------Fechamento Geral -----------------------

Private Sub FechamentoGeral()

Limapgrd
 wTNNotaCredito = 0
 wTotalNf = 0
 wTotalFatFin = 0
 grdMovimentoCaixa.Row = 1
 Screen.MousePointer = 11
 sql = ("select mc_Grupo,sum(MC_Valor) as TotalModalidade,Count(*) as Quantidade from movimentocaixa" _
       & " Where MC_Data between '" & Format(datageral, "yyyy/mm/dd") & " 00:00:00.000' and '" _
       & Format(datageral, "yyyy/mm/dd") & " 23:59:59:000" _
       & "' and  MC_Serie <> '00' and (MC_Grupo like '10%' or MC_Grupo like '11%'" _
       & " or MC_Grupo like '50%' or MC_Grupo like '20%') and mc_tiponota  <> 'C' group by mc_grupo")
       rdoFormaPagamento.CursorLocation = adUseClient
       rdoFormaPagamento.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
       
'1999-01-02 00:00:00.000
If Not rdoFormaPagamento.EOF Then
    Do While Not rdoFormaPagamento.EOF
        If rdoFormaPagamento("MC_Grupo") = "10101" Then
           grdMovimentoCaixa.TextMatrix(1, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10201" Then
           grdMovimentoCaixa.TextMatrix(2, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("mc_grupo") = "10301" Then
           grdMovimentoCaixa.TextMatrix(4, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10302" Then
           grdMovimentoCaixa.TextMatrix(5, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10303" Then
           grdMovimentoCaixa.TextMatrix(6, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10304" Then
           grdMovimentoCaixa.TextMatrix(7, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10203" Then
           grdMovimentoCaixa.TextMatrix(8, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10701" Then
           grdMovimentoCaixa.TextMatrix(9, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
           'wTNNotaCredito = wTNNotaCredito + rdoFormaPagamento("TotalModalidade")
        'ElseIf rdoFormaPagamento("MC_Grupo") = "10204" Then 'AVR
           'grdMovimentoCaixa.TextMatrix(9, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           'wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "11004" Then
           grdMovimentoCaixa.TextMatrix(15, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
           wTotalEntrada = (wTotalEntrada + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "11005" Then
           grdMovimentoCaixa.TextMatrix(16, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           'wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
           wTotalEntrada = (wTotalEntrada + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10801" Then
           grdMovimentoCaixa.TextMatrix(17, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           'wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        'ElseIf rdoFormaPagamento("MC_Grupo") = "11008" Then 'AVR
           'grdMovimentoCaixa.TextMatrix(15, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           'wTotalSaldo = (wTotalSaldo + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "11006" Then 'SALDO ANTERIOR
           grdMovimentoCaixa.TextMatrix(43, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wTotalSaldo = (wTotalSaldo + rdoFormaPagamento("TotalModalidade"))
           'wtotalSaldoAnterior = (wtotalSaldoAnterior + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "11007" Then 'SALDO ANTERIOR
           grdMovimentoCaixa.TextMatrix(45, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wTotalSaldo = (wTotalSaldo + rdoFormaPagamento("TotalModalidade"))
           'wtotalSaldoAnterior = (wtotalSaldoAnterior + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10501" Then
           grdMovimentoCaixa.TextMatrix(13, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wTotalFatFin = (wTotalFatFin + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10601" Then
           grdMovimentoCaixa.TextMatrix(14, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wTotalFatFin = (wTotalFatFin + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "11009" Then
           grdMovimentoCaixa.TextMatrix(18, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wtotalGarantia = (wtotalGarantia + rdoFormaPagamento("TotalModalidade"))
           wSubTotal = (wSubTotal - rdoFormaPagamento("TotalModalidade"))
           wSubTotal_S = (wSubTotal_S - rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "20101" Then
           grdMovimentoCaixa.TextMatrix(24, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           grdMovimentoCaixa.TextMatrix(24, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
           wTotalNf = (wTotalNf + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "20102" Then
           grdMovimentoCaixa.TextMatrix(25, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           grdMovimentoCaixa.TextMatrix(25, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
           wTotalNf = (wTotalNf + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "20107" Then
           grdMovimentoCaixa.TextMatrix(27, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           grdMovimentoCaixa.TextMatrix(27, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
           wTotalNf = (wTotalNf + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "20108" Then
           grdMovimentoCaixa.TextMatrix(28, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           grdMovimentoCaixa.TextMatrix(28, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
           wTotalNf = (wTotalNf + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "20111" Then
           grdMovimentoCaixa.TextMatrix(26, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           grdMovimentoCaixa.TextMatrix(26, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
           wTotalNf = (wTotalNf + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "20109" Then
           grdMovimentoCaixa.TextMatrix(32, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           grdMovimentoCaixa.TextMatrix(32, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
        ElseIf rdoFormaPagamento("MC_Grupo") = "20110" Then
           grdMovimentoCaixa.TextMatrix(33, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           grdMovimentoCaixa.TextMatrix(33, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
        ElseIf rdoFormaPagamento("MC_Grupo") = "20201" Then
           grdMovimentoCaixa.TextMatrix(34, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           grdMovimentoCaixa.TextMatrix(34, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
        
        ElseIf rdoFormaPagamento("MC_Grupo") = "50101" Then
           grdMovimentoCaixa.TextMatrix(1, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50201" Then
           grdMovimentoCaixa.TextMatrix(2, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("mc_grupo") = "50301" Then
           grdMovimentoCaixa.TextMatrix(4, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50302" Then
           grdMovimentoCaixa.TextMatrix(5, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50303" Then
           grdMovimentoCaixa.TextMatrix(6, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50304" Then
           grdMovimentoCaixa.TextMatrix(7, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50203" Then
           grdMovimentoCaixa.TextMatrix(8, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50701" Then
           grdMovimentoCaixa.TextMatrix(9, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        'ElseIf rdoFormaPagamento("MC_Grupo") = "50204" Then 'AVR
           'grdMovimentoCaixa.TextMatrix(9, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           'wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50502" Then
           grdMovimentoCaixa.TextMatrix(15, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50602" Then
           grdMovimentoCaixa.TextMatrix(16, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           'wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50801" Then
           grdMovimentoCaixa.TextMatrix(17, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        'ElseIf rdoFormaPagamento("MC_Grupo") = "50804" Then
           'grdMovimentoCaixa.TextMatrix(15, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           'wTotalSaldo_S = (wTotalSaldo_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50806" Then
           grdMovimentoCaixa.TextMatrix(43, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wTotalSaldo_S = (wTotalSaldo_S + rdoFormaPagamento("TotalModalidade"))
           'wtotalSaldoAnterior_S = (wtotalSaldoAnterior_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50807" Then
           grdMovimentoCaixa.TextMatrix(45, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wTotalSaldo_S = (wTotalSaldo_S + rdoFormaPagamento("TotalModalidade"))
           'wtotalSaldoAnterior_S = (wtotalSaldoAnterior_S + rdoFormaPagamento("TotalModalidade"))
        
        End If
       rdoFormaPagamento.MoveNext
     Loop
     
     grdMovimentoCaixa.TextMatrix(1, 1) = Format((grdMovimentoCaixa.TextMatrix(1, 1) - wtotalGarantia), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(1, 2) = Format((grdMovimentoCaixa.TextMatrix(1, 2) - wtotalGarantia), "###,###,###,##0.00")

     grdMovimentoCaixa.TextMatrix(1, 3) = Format((grdMovimentoCaixa.TextMatrix(1, 1) - grdMovimentoCaixa.TextMatrix(1, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(2, 3) = Format((grdMovimentoCaixa.TextMatrix(2, 1) - grdMovimentoCaixa.TextMatrix(2, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(4, 3) = Format((grdMovimentoCaixa.TextMatrix(4, 1) - grdMovimentoCaixa.TextMatrix(4, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(5, 3) = Format((grdMovimentoCaixa.TextMatrix(5, 1) - grdMovimentoCaixa.TextMatrix(5, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(6, 3) = Format((grdMovimentoCaixa.TextMatrix(6, 1) - grdMovimentoCaixa.TextMatrix(6, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(7, 3) = Format((grdMovimentoCaixa.TextMatrix(7, 1) - grdMovimentoCaixa.TextMatrix(7, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(8, 3) = Format((grdMovimentoCaixa.TextMatrix(8, 1) - grdMovimentoCaixa.TextMatrix(8, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(9, 3) = Format((grdMovimentoCaixa.TextMatrix(9, 1) - grdMovimentoCaixa.TextMatrix(9, 2)), "###,###,###,##0.00")
     'grdMovimentoCaixa.TextMatrix(9, 3) = Format((grdMovimentoCaixa.TextMatrix(9, 1) - grdMovimentoCaixa.TextMatrix(9, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(15, 3) = Format((grdMovimentoCaixa.TextMatrix(15, 1) - grdMovimentoCaixa.TextMatrix(15, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(16, 3) = Format((grdMovimentoCaixa.TextMatrix(16, 1) - grdMovimentoCaixa.TextMatrix(16, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(17, 3) = Format((grdMovimentoCaixa.TextMatrix(17, 1) - grdMovimentoCaixa.TextMatrix(17, 2)), "###,###,###,##0.00")
      grdMovimentoCaixa.TextMatrix(43, 3) = Format((grdMovimentoCaixa.TextMatrix(43, 1) - grdMovimentoCaixa.TextMatrix(43, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(45, 3) = Format((grdMovimentoCaixa.TextMatrix(45, 1) - grdMovimentoCaixa.TextMatrix(45, 2)), "###,###,###,##0.00")
     
'ARRUMAR

     rdoFormaPagamento.Close
  
    sql = ("select MC_Grupo,sum(mc_valor)as mc_valor from MovimentoCaixa where MC_Grupo in ('11006','50806') and " _
        & "MC_Data = '" & Format(datageral, "yyyy/mm/dd") _
       & "' group by MC_Grupo")
     rdoFormaPagamento.CursorLocation = adUseClient
     rdoFormaPagamento.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
     
     Do While Not rdoFormaPagamento.EOF
         
        If rdoFormaPagamento("MC_GRUPO") = "11006" Then
            grdMovimentoCaixa.TextMatrix(43, 1) = Format(rdoFormaPagamento("mc_valor"), "###,###,###,##0.00")
            grdMovimentoCaixa.TextMatrix(44, 1) = Format(0, "###,###,###,##0.00")
            'grdMovimentoCaixa.TextMatrix(41, 3) = Format(rdoFormaPagamento("mc_valor"), "###,###,###,##0.00")
        ElseIf rdoFormaPagamento("MC_GRUPO") = "50806" Then
            grdMovimentoCaixa.TextMatrix(43, 2) = Format(rdoFormaPagamento("mc_valor"), "###,###,###,##0.00")
            grdMovimentoCaixa.TextMatrix(44, 2) = Format(0, "###,###,###,##0.00")
        End If
        

        rdoFormaPagamento.MoveNext
     Loop
        
        grdMovimentoCaixa.TextMatrix(43, 3) = Format(CDbl(grdMovimentoCaixa.TextMatrix(43, 1)) - CDbl(grdMovimentoCaixa.TextMatrix(43, 2)), "###,###,###,##0.00")
        grdMovimentoCaixa.TextMatrix(44, 3) = Format(grdMovimentoCaixa.TextMatrix(44, 1) - grdMovimentoCaixa.TextMatrix(44, 2), "###,###,###,##0.00")
        
        grdMovimentoCaixa.TextMatrix(45, 1) = Format(CDbl(grdMovimentoCaixa.TextMatrix(43, 1)) + CDbl(grdMovimentoCaixa.TextMatrix(44, 1)), "###,###,###,##0.00")
        grdMovimentoCaixa.TextMatrix(45, 2) = Format(CDbl(grdMovimentoCaixa.TextMatrix(43, 2)) + CDbl(grdMovimentoCaixa.TextMatrix(44, 2)), "###,###,###,##0.00")
        
        grdMovimentoCaixa.TextMatrix(45, 3) = Format(CDbl(grdMovimentoCaixa.TextMatrix(45, 1)) - CDbl(grdMovimentoCaixa.TextMatrix(45, 2)), "###,###,###,##0.00")

''''''''''''''''''''''''''''''''

        grdMovimentoCaixa.TextMatrix(11, 1) = Format((wSubTotal), "###,###,###,##0.00")
        grdMovimentoCaixa.TextMatrix(11, 2) = Format((wSubTotal_S), "###,###,###,##0.00")
        grdMovimentoCaixa.TextMatrix(11, 3) = Format(((wSubTotal) - (wSubTotal_S)), "###,###,###,##0.00")
        grdMovimentoCaixa.TextMatrix(30, 1) = Format(wTotalNf, "###,###,###,##0.00")
        grdMovimentoCaixa.TextMatrix(20, 1) = Format(((wTotalFatFin + wSubTotal) - wTotalEntrada), "###,###,###,##0.00")
        
  End If
  rdoFormaPagamento.Close
  
    sql = ("select mc_Grupo,sum(MC_Valor) as TotalModalidade,Count(*) as Quantidade from movimentocaixa" _
       & " Where MC_Data between '" & Format(datageral, "yyyy/mm/dd") & " 00:00:00.000' and '" _
       & Format(datageral, "yyyy/mm/dd") & " 23:59:59:000" & _
         "' and MC_Serie <> '00' and MC_tiponota = 'C' and MC_Grupo like '20%' group by mc_grupo")
       rdoFormaPagamento.CursorLocation = adUseClient
       rdoFormaPagamento.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
       
  If Not rdoFormaPagamento.EOF Then
     Do While Not rdoFormaPagamento.EOF
        If rdoFormaPagamento("MC_Grupo") = "20101" Then
           grdMovimentoCaixa.TextMatrix(35, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           grdMovimentoCaixa.TextMatrix(35, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
        ElseIf rdoFormaPagamento("MC_Grupo") = "20102" Then
           grdMovimentoCaixa.TextMatrix(36, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           grdMovimentoCaixa.TextMatrix(36, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
        ElseIf rdoFormaPagamento("MC_Grupo") = "20107" Then
           grdMovimentoCaixa.TextMatrix(37, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           grdMovimentoCaixa.TextMatrix(37, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
        ElseIf rdoFormaPagamento("MC_Grupo") = "20108" Then
           grdMovimentoCaixa.TextMatrix(38, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           grdMovimentoCaixa.TextMatrix(38, 2) = Format(rdoFormaPagamento("Quantidade"), "0")
        End If
        rdoFormaPagamento.MoveNext
     Loop
  End If
 
  rdoFormaPagamento.Close
 
  wSubTotal = 0
  wSubTotal_S = 0
  Screen.MousePointer = 0
  


End Sub


Private Sub Limapgrd()
  grdMovimentoCaixa.Rows = 1
  grdMovimentoCaixa.Rows = 1
  grdMovimentoCaixa.AddItem "Dinheiro" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "70101"                    '1
  grdMovimentoCaixa.AddItem "Cheque" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "70201"                '2                                                                              '2
  grdMovimentoCaixa.AddItem "Cartões >>"                                                                           '3
  grdMovimentoCaixa.AddItem "  Visa" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50301"                      '4
  grdMovimentoCaixa.AddItem "  MasterCard" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50302"                '5
  grdMovimentoCaixa.AddItem "  Amex" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50303"                      '6
  grdMovimentoCaixa.AddItem "  BNDES" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50304"                     '7
  grdMovimentoCaixa.AddItem "  T E F" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50203"                     '8
  grdMovimentoCaixa.AddItem "Nota de Credito" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50701"             '9
  grdMovimentoCaixa.AddItem ""                                                                                      '10
  grdMovimentoCaixa.AddItem "*** TOTAL CAIXA" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "70204"             '11
  grdMovimentoCaixa.AddItem ""                                                                                      '12
  grdMovimentoCaixa.AddItem "Faturado" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                       '13'14
  grdMovimentoCaixa.AddItem "Financiada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                     '14             '15
  grdMovimentoCaixa.AddItem "Entrada Faturada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50502"            '15
  grdMovimentoCaixa.AddItem "Entrada Financiada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50602"          '16
  grdMovimentoCaixa.AddItem "Reforco de Caixa" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50801"            '17
  grdMovimentoCaixa.AddItem "Garantia Estendida" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                             '18
  grdMovimentoCaixa.AddItem ""                                                                                      '19
  grdMovimentoCaixa.AddItem "*** TOTAL GERAL" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                '20              '17
  grdMovimentoCaixa.AddItem ""                                                                                      '21
  grdMovimentoCaixa.AddItem "*** Movimento NF"                                                                      '22
  grdMovimentoCaixa.AddItem ""                                                                                      '23
  grdMovimentoCaixa.AddItem "CF" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                             '24
  grdMovimentoCaixa.AddItem "SN" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                             '25
  grdMovimentoCaixa.AddItem "NE" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                             '26
  grdMovimentoCaixa.AddItem "D1" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                             '27
  grdMovimentoCaixa.AddItem "S1" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                             '28
  grdMovimentoCaixa.AddItem ""                                                                                      '29
  grdMovimentoCaixa.AddItem "*** TOTAL NF" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                   '30
  grdMovimentoCaixa.AddItem ""                                                                                      '31
  grdMovimentoCaixa.AddItem "Transferencia Saida" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                            '32
  grdMovimentoCaixa.AddItem "Remessa Saida" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                  '33
  grdMovimentoCaixa.AddItem "Devolucao Entrada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                              '34
  grdMovimentoCaixa.AddItem ""                                                                                      '35
  grdMovimentoCaixa.AddItem "CF Cancelada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                   '36
  grdMovimentoCaixa.AddItem "SN Cancelada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                   '37
  grdMovimentoCaixa.AddItem "NE Cancelada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                   '38
  grdMovimentoCaixa.AddItem "D1 Cancelada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                   '39
  grdMovimentoCaixa.AddItem "S1 Cancelada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0                                   '40
  grdMovimentoCaixa.AddItem ""                                                                                      '41                                                                            '44
  grdMovimentoCaixa.AddItem "** Saldo Anterior**"                                                                   '42
  grdMovimentoCaixa.AddItem "Dinheiro" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "70101"                    '43
  grdMovimentoCaixa.AddItem "Cheque" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "70201"                      '44
  grdMovimentoCaixa.AddItem "Total do Saldo" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "00000"              '45
  grdMovimentoCaixa.AddItem ""



  wControlaSaldoCaixa = 0
  wTotalSaldo = 0
  wTotalSaldo_S = 0 '45                                                                          '46

 
End Sub


Private Sub Imprimefechamento()
 Dim wImprime As String
Dim wImprime2 As String
Dim wTotal00 As Double
wTotal00 = 0
Screen.MousePointer = 11
Retorno = Bematech_FI_AbreRelatorioGerencialMFD("01")
     
     wSaldoFinalDinheiro = 0
     wSaldoFinalCheque = 0
     wSaldoFinalAVR = 0
     
''     Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
''               left(Trim(lblCabec.Caption) & Space(38), 38) & right(Space(10) & "Loja " & Format(GLB_Loja, "000"), 10) & _
''               "________________________________________________")
               
     Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
               "                    IMPRESSAO                   " & _
               " Fechamento GERAL " & Format(right(lblCabec.Caption, 10), "dd/mm/yyyy") & "            Loja " & Format(GLB_Loja, "000") & _
               "________________________________________________")

    For Idx = 1 To grdMovimentoCaixa.Rows - 6 Step 1
    If (Idx = 23) Or (Idx = 24) Or (Idx = 25) Or (Idx = 26) Or (Idx = 27) Or (Idx = 30) Or (Idx = 31) Or (Idx = 32) Or _
         (Idx = 35) Or (Idx = 36) Or (Idx = 37) Or (Idx = 38) Or (Idx = 39) Then
         

         Retorno = Bematech_FI_UsaRelatorioGerencialMFD(left(grdMovimentoCaixa.TextMatrix(Idx, 0) & Space(23), 23) & _
               right(Space(20) & Format(grdMovimentoCaixa.TextMatrix(Idx, 1), "###,###,##0.00"), 20) & _
               right(Space(5) & grdMovimentoCaixa.TextMatrix(Idx, 2), 5))
         Else

          Retorno = Bematech_FI_UsaRelatorioGerencialMFD(left(grdMovimentoCaixa.TextMatrix(Idx, 0) & Space(23), 23) & _
               right(Space(20) & Format(grdMovimentoCaixa.TextMatrix(Idx, 1), "###,###,##0.00"), 20) & "     ")
         End If
     Next Idx
     
     
     wMovimentoPeriodo = Format(grdMovimentoCaixa.TextMatrix(11, 1), "###,###,##0.00")
     wSaldoAnterior = Format(grdMovimentoCaixa.TextMatrix(43, 1), "###,###,##0.00")
     'wMovimentoPeriodo = (wMovimentoPeriodo - wSaldoAnterior)
     wSaldoFinalDinheiro = Format(grdMovimentoCaixa.TextMatrix(1, 3), "###,###,##0.00")
     wSaldoFinalDinheiro = (wSaldoFinalDinheiro + Format(grdMovimentoCaixa.TextMatrix(43, 3), "###,###,##0.00"))
     wSaldoFinalCheque = Format(grdMovimentoCaixa.TextMatrix(2, 3))
     wSaldoFinalCheque = (wSaldoFinalCheque + Format(grdMovimentoCaixa.TextMatrix(45, 3), "###,###,##0.00"))
     'wSaldoFinalAVR = Format(grdMovimentoCaixa.TextMatrix(9, 3))
     'wSaldoFinalAVR = (wSaldoFinalAVR + Format(grdMovimentoCaixa.TextMatrix(15, 3), "###,###,##0.00"))
 
     Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
        "               MOVIMENTO DE CAIXA               " & _
        "________________________________________________" & _
        "SALDO ANTERIOR >> " & right(Space(30) & Format("", ""), 30) & _
        "  DINHEIRO        " & right(Space(30) & Format(grdMovimentoCaixa.TextMatrix(41, 1), "###,###,##0.00"), 30) & _
        "  CHEQUE          " & right(Space(30) & Format(grdMovimentoCaixa.TextMatrix(42, 1), "###,###,##0.00"), 30) & _
        "  TOTAL           " & right(Space(30) & Format(wTotalSaldo, "###,###,##0.00"), 30) & _
        "MOVIMENTO PERIODO " & right(Space(30) & Format(wMovimentoPeriodo, "###,###,##0.00"), 30) & _
        "REFORCO           " & right(Space(30) & Format(grdMovimentoCaixa.TextMatrix(17, 1), "###,###,##0.00"), 30) & _
        "TRANSFERENCIA NUM." & right(Space(30) & Format(grdMovimentoCaixa.TextMatrix(11, 2), "###,###,##0.00"), 30))
 
     Retorno = Bematech_FI_UsaRelatorioGerencialMFD("SALDO DO CAIXA >>                               " & _
        "  Dinheiro        " & right(Space(30) & Format(wSaldoFinalDinheiro, "###,###,##0.00"), 30) & _
        "  Cheque          " & right(Space(30) & Format(wSaldoFinalCheque, "###,###,##0.00"), 30) & _
        "  SALDO FINAL     " & right(Space(30) & Format((wSaldoFinalDinheiro + wSaldoFinalCheque), "###,###,##0.00"), 30) & _
        "________________________________________________")
          
      Retorno = Bematech_FI_UsaRelatorioGerencialMFD(left("Caixa Nro.   " & Trim(frmControlaCaixa.lblNroCaixa.Caption) & Space(48), 48) & _
         left("Data Inicial " & Format(Date, "dd/mm/yyyy") & Space(48), 48) & _
         left("Data Final   " & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS") & Space(48), 48) & _
         left("Protocolo    " & GLB_CTR_Protocolo & Space(48), 48))
 
 
       Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
         "                                                " & _
         "                                                " & _
         "                                                " & _
         " ______________________                         " & _
         " " & left(Trim(wNomeSupervisor) & Space(47), 47) & _
         "                                                " & _
         "                                                ")
    
 
        Retorno = Bematech_FI_FechaRelatorioGerencial()
 
      Screen.MousePointer = 0
 
     


 
sql = ("Select Serie, sum(totalnota) as TotalSerieNota, count(Serie) as QtdeSerie from nfcapa Where ecf = " & GLB_ECF & "" _
     & " and  TipoNota = 'V' and Serie = '00' and  DataEmi = '" & Format(Date, "yyyy/mm/dd") _
     & "' " & " and Protocolo = " & GLB_CTR_Protocolo & " group by Serie ")
     rdoCapa.CursorLocation = adUseClient
     rdoCapa.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
     
If Not rdoCapa.EOF Then
   wTotal00 = rdoCapa("TotalSerieNota")
End If
 
rdoCapa.Close
    Screen.MousePointer = 11
         Retorno = Bematech_FI_AbreRelatorioGerencialMFD("01")
    
          Retorno = Bematech_FI_UsaRelatorioGerencialMFD("                                                " & _
                    "                     CONTROLE X                 " & _
                    " Loja " & left(GLB_Loja & Space(4), 4) & Space(28) & Format(Date, "dd/mm/yyyy") & _
                    "                                                " & _
                    " Valor Total:                     " & right(Space(10) & Format(wTotal00, "###,###,##0.00"), 14) & _
                    "________________________________________________" & _
                    "                                                " & _
                    left("Caixa Nro.   " & frmControlaCaixa.lblNroCaixa.Caption & Space(48), 48) & _
                    left("Operador     " & wNomeSupervisor & Space(48), 48) & _
                    left("Data Inicial " & Format(Date, "dd/mm/yyyy") & Space(48), 48) & _
                    left("Data Final   " & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS") & Space(48), 48) & _
                    left("Protocolo    " & GLB_CTR_Protocolo & Space(48), 48))
                    
                    
          Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
                    "                                                " & _
                    "                                                " & _
                    "                                                " & _
                    " ______________________                         " & _
                    " " & left(wNomeSupervisor & Space(47), 47) & _
                    "                                                " & _
                    "                                                ")
        Retorno = Bematech_FI_FechaRelatorioGerencial()
 
      Screen.MousePointer = 0
 
    Screen.MousePointer = 11
    Retorno = Bematech_FI_AbreRelatorioGerencialMFD("01")
 
    Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
                   "TRANSFERENCIA DE NUMERARIO GERAL" & right(Space(16) & "Nro.  " & GLB_CTR_Protocolo, 16) & _
                   left("Loja " & Format(GLB_Loja, "000") & Space(10), 10) & right(Space(38) & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS"), 38) & _
                   "________________________________________________" & _
                   "    USO INTERNO          SEM VALOR COMERCIAL    ")
   
   Retorno = Bematech_FI_UsaRelatorioGerencialMFD("------------------------------------------------" & _
                   "VISA              " & right(Space(30) & Format(wTNVisa, "###,###,##0.00"), 30) & _
                   "MASTERCARD        " & right(Space(30) & Format(wTNRedecard, "###,###,##0.00"), 30) & _
                   "AMEX              " & right(Space(30) & Format(wTNAmex, "###,###,##0.00"), 30) & _
                   "BNDS              " & right(Space(30) & Format(wTNBNDES, "###,###,##0.00"), 30) & _
                   "HIPERCARD         " & right(Space(30) & Format(wTNHiperCard, "###,###,##0.00"), 30))
                   
   Retorno = Bematech_FI_UsaRelatorioGerencialMFD( _
                   "TEF               " & right(Space(30) & Format(wTNVisaEletron + wTNRedeShop, "###,###,##0.00"), 30) & _
                   "DEPOSITO          " & right(Space(30) & Format(wTNDeposito, "###,###,##0.00"), 30) & _
                   "NOTA CREDITO      " & right(Space(30) & Format(wTNNotaCredito, "###,###,##0.00"), 30) & _
                   "DESPESA CONDUCAO  " & right(Space(30) & Format(wTNConducao, "###,###,##0.00"), 30) & _
                   "DESPESA LOJA      " & right(Space(30) & Format(wTNDespLoja, "###,###,##0.00"), 30) & _
                   "OUTRAS DESPESAS   " & right(Space(30) & Format(wTNOutros, "###,###,##0.00"), 30) & _
                   "                                                " & _
                   "TOTAL             " & right(Space(30) & Format(wTNTotal, "###,###,##0.00"), 30))
                   
    Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
                    "                                                " & _
                    "                                                " & _
                    "                                                " & _
                    " ______________________                         " & _
                    " " & left(wNomeSupervisor & Space(47), 47) & _
                    "                                                " & _
                    "                                                ")
        Retorno = Bematech_FI_FechaRelatorioGerencial()
 
      Screen.MousePointer = 0
 
End Sub


