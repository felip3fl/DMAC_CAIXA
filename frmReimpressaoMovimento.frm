VERSION 5.00
Object = "{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7u.ocx"
Begin VB.Form frmReimpressaoMovimento 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Fechamento do CaixaFechamento do Caixa"
   ClientHeight    =   8505
   ClientLeft      =   480
   ClientTop       =   1125
   ClientWidth     =   15300
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   15300
   ShowInTaskbar   =   0   'False
   Begin VSFlex7UCtl.VSFlexGrid grdMovimentoCaixa 
      Height          =   1185
      Left            =   8730
      TabIndex        =   1
      Top             =   1680
      Width           =   4830
      _cx             =   8520
      _cy             =   2090
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
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReimpressaoMovimento.frx":0000
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
   Begin VSFlex7UCtl.VSFlexGrid grdMovimentosDisponiveis 
      Height          =   6975
      Left            =   300
      TabIndex        =   0
      Top             =   615
      Width           =   4785
      _cx             =   8440
      _cy             =   12303
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReimpressaoMovimento.frx":00A0
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   7
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
   Begin VSFlex7UCtl.VSFlexGrid grdAnaliticoVenda 
      Height          =   780
      Left            =   8790
      TabIndex        =   3
      Top             =   3150
      Width           =   4665
      _cx             =   8229
      _cy             =   1376
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
      Rows            =   50
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReimpressaoMovimento.frx":0174
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
   Begin VSFlex7UCtl.VSFlexGrid gridImpressao 
      Height          =   5355
      Left            =   5745
      TabIndex        =   7
      Top             =   1650
      Visible         =   0   'False
      Width           =   2940
      _cx             =   5186
      _cy             =   9446
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReimpressaoMovimento.frx":01F8
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   4
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
   Begin VB.CheckBox ChkModoImpressao 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Modo Impressão"
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
      Height          =   390
      Left            =   5745
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   6
      Top             =   735
      Width           =   3090
   End
   Begin VSFlex7UCtl.VSFlexGrid grdMovimentoAdministrador 
      Height          =   1065
      Left            =   14355
      TabIndex        =   8
      Top             =   3915
      Visible         =   0   'False
      Width           =   9765
      _cx             =   17224
      _cy             =   1879
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReimpressaoMovimento.frx":02CF
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   4
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Analítico de Venda"
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
      Left            =   10155
      TabIndex        =   9
      Top             =   5925
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Image fraFechamentoAnterior 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   5475
      Top             =   615
      Width           =   3500
   End
   Begin VB.Label lblMSGImpressao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fila de Impressão"
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
      Height          =   255
      Left            =   5745
      TabIndex        =   5
      Top             =   1305
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Label lblCabec2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Analítico de Venda"
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
      Left            =   2895
      TabIndex        =   4
      Top             =   195
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Label lblCabec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Controle Movimento"
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
      TabIndex        =   2
      Top             =   200
      Width           =   2175
   End
End
Attribute VB_Name = "frmReimpressaoMovimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim wSubTotal As Double
Dim wSubTotal_S As Double
Dim wtotalGarantia As Double
Dim wTNFinanciado As Double
Dim wTNFaturado As Double
Dim wtotalSaldoAnterior As Double
Dim wtotalSaldoAnterior_S As Double

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

Dim wQtdeGrid As Integer
Dim Idx As Long
Dim sql As String
Dim Cor As String
Dim Cor1 As String
Dim Cor2 As String
Dim Cor3 As String

'Dim SQL As String
Dim wProtocoloImpressao As Long
Dim wOperadorImpressao As String
Dim wNroCaixaImpressao As String
Dim wTotalCaixa As Double
Dim wDataInicioFechamento As String
Dim wHoraInicioFechamento As String
Dim wDataFinalFechamento As String
Dim wHoraFinalFechamento As String

Dim wCampoAdminstrador As String
Dim rsMovimento As New ADODB.Recordset
Dim wGrupo As String

Dim wDiasCarregaMov As Integer

Private Sub chbLeituraX_Click()

End Sub


Private Sub ChkModoImpressao_Click()

    If ChkModoImpressao.Value = 1 Then
    
        gridImpressao.Rows = 1
        fraFechamentoAnterior.Height = grdMovimentosDisponiveis.Height
    Else
    
        fraFechamentoAnterior.Height = 615
    End If
    
End Sub

Private Sub cmdImprimir_Click()



    
End Sub

Private Sub Form_Activate()
    
    If GLB_Administrador = True Then
        grdMovimentoAdministrador.Rows = 1
        grdMovimentoAdministrador.Visible = True
    End If

    fraFechamentoAnterior.Height = 615

    Screen.MousePointer = 11
    
    
    grdMovimentoCaixa.top = grdMovimentosDisponiveis.top
    grdMovimentoCaixa.left = grdMovimentosDisponiveis.left + grdMovimentosDisponiveis.Width + 200
    
    
    grdAnaliticoVenda.left = grdMovimentoCaixa.Width + grdMovimentoCaixa.left + 200
    grdAnaliticoVenda.left = grdAnaliticoVenda.left
    grdAnaliticoVenda.top = grdMovimentoCaixa.top
    lblCabec2.left = grdAnaliticoVenda.left
    
    Call CarregaControleMovimento
    
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

    wDiasCarregaMov = 15

    left = 100
    top = 2900
    
    Call AjustaTela(frmReimpressaoMovimento)
    'grdMovimentosDisponiveis.Visible = False
    grdMovimentoCaixa.Visible = False
    'Image1.Width = grdMovimentoCaixa.Width
    'txtSupervisor.Enabled = True
    'txtSenhaSupervisor.Enabled = True
    lblCabec.Caption = "Controle Movimento"
    ''optControleMovimento.Value = True
    'optAnaliticoVenda.Value = False
    grdAnaliticoVenda.Visible = False
    
    'fraAlteraDia.top = Image1.top
    'fraAlteraDia.left = Image1.left
    'fraAlteraDia.Width = Image1.Width
    'fraAlteraDia.Height = Image1.Height
    'fraAlteraDia.Visible = False

    defineImpressora

End Sub

Private Sub grdAnaliticoVenda_DblClick()
 If MsgBox("Deseja imprimir o movimento?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
 
  Call ImprimeAnaliticoVenda

  Screen.MousePointer = 0
  
 End If
End Sub

Private Sub grdAnaliticoVenda_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
     'grdMovimentosDisponiveis.Visible = True
     grdMovimentoCaixa.Visible = False
     grdAnaliticoVenda.Visible = False
     lblCabec2.Visible = False
     'Image1.Width = grdMovimentosDisponiveis.Width
     ChkModoImpressao.Visible = True
'   frmControlaCaixa.txtPedido.SetFocus
End If

End Sub

Private Sub grdMovimentoCaixa_DblClick()



    If MsgBox("Deseja imprimir o movimento?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        Call CarregaValoresTransfNumerario(wProtocoloImpressao)
        
        wQdteViasImpressao = 1
        Call BuscaQtdeViaImpressaoMovimento
        
        For i = 1 To wQdteViasImpressao

        
            Call NOVO_ImprimeMovimento(grdMovimentoCaixa, "FECHAMENTO DE CAIXA (Reimpressao)", wOperadorImpressao, wNroCaixaImpressao, _
                                       wDataInicioFechamento, Format(wHoraInicioFechamento, "HH:MM:SS"), _
                                       wDataFinalFechamento, Format(wHoraFinalFechamento, "HH:MM:SS"), _
                                       CStr(wProtocoloImpressao))
                                       
      
            Call NOVO_ImprimeTransfNumerario(grdMovimentoCaixa, "TRANSFERENCIA DE NUMERARIO (Reimpressao)", wOperadorImpressao, wNroCaixaImpressao, _
                                             wDataInicioFechamento, Format(wHoraInicioFechamento, "HH:MM:SS"), _
                                             wDataFinalFechamento, Format(wHoraFinalFechamento, ""), _
                                             CStr(wProtocoloImpressao))
        Next i
        Screen.MousePointer = 0
    End If
 
''''End If
 
End Sub

Private Sub grdMovimentoCaixa_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    'grdMovimentosDisponiveis.Visible = True
     grdMovimentoCaixa.Visible = False
     grdAnaliticoVenda.Visible = False
     lblCabec2.Visible = False
     'Image1.Width = grdMovimentosDisponiveis.Width
     ChkModoImpressao.Visible = True
'   frmControlaCaixa.txtPedido.SetFocus
End If
End Sub

Private Sub grdMovimentosDisponiveis_Click()

 If ChkModoImpressao.Value = 0 Then
       If grdMovimentosDisponiveis.Row > 1 Then
          Screen.MousePointer = 11
       
           ChkModoImpressao.Visible = False
           wProtocoloImpressao = grdMovimentosDisponiveis.TextMatrix(grdMovimentosDisponiveis.Row, 3)
           wOperadorImpressao = grdMovimentosDisponiveis.TextMatrix(grdMovimentosDisponiveis.Row, 4)
           wNroCaixaImpressao = grdMovimentosDisponiveis.TextMatrix(grdMovimentosDisponiveis.Row, 2)
           wDataInicioFechamento = Format(grdMovimentosDisponiveis.TextMatrix(grdMovimentosDisponiveis.Row, 0), "DD/MM/YYYY")
           wHoraInicioFechamento = Format(grdMovimentosDisponiveis.TextMatrix(grdMovimentosDisponiveis.Row, 0), "HH:MM")
           wDataFinalFechamento = Format(grdMovimentosDisponiveis.TextMatrix(grdMovimentosDisponiveis.Row, 1), "DD/MM/YYYY")
           wHoraFinalFechamento = Format(grdMovimentosDisponiveis.TextMatrix(grdMovimentosDisponiveis.Row, 1), "HH:MM")
           
          'grdMovimentosDisponiveis.Visible = False
          grdMovimentoCaixa.Visible = True
          'grdAnaliticoVenda.Visible = True
           
          sql = ("Select Max(CTr_DataInicial)as DataMov,Max(Ctr_Protocolo) as Seq " _
              & "from ControleCaixa where CTR_Supervisor <> 99 and CTr_NumeroCaixa = " & GLB_Caixa)
             
          Call CarregaAnaliticoVenda(wDataInicioFechamento)
          Call CarregaMovimento(grdMovimentoCaixa, CLng(wProtocoloImpressao))
          'grdMovimentoCaixa.SetFocus
           
           Screen.MousePointer = 0
       ElseIf grdMovimentosDisponiveis.Row = 1 Then
           Screen.MousePointer = 11
           wDiasCarregaMov = wDiasCarregaMov + 15
           Call CarregaControleMovimento
           Screen.MousePointer = 0
       End If
        
    Else
    
        lblMSGImpressao.Visible = True
        gridImpressao.Visible = True
        gridImpressao.AddItem grdMovimentosDisponiveis.TextMatrix(grdMovimentosDisponiveis.Row, 0) & Chr(9) _
                              & grdMovimentosDisponiveis.TextMatrix(grdMovimentosDisponiveis.Row, 1) & Chr(9) _
                              & grdMovimentosDisponiveis.TextMatrix(grdMovimentosDisponiveis.Row, 2) & Chr(9) _
                              & grdMovimentosDisponiveis.TextMatrix(grdMovimentosDisponiveis.Row, 3) & Chr(9) _
                              & grdMovimentosDisponiveis.TextMatrix(grdMovimentosDisponiveis.Row, 4) & Chr(9) _
                              & grdMovimentosDisponiveis.TextMatrix(grdMovimentosDisponiveis.Row, 5)
                                                                             
        'wOrdemImpressao = wOrdemImpressao + 1
    
    End If

End Sub

Private Sub grdMovimentosDisponiveis_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
   Unload Me
   frmControlaCaixa.txtPedido.SetFocus
 End If
End Sub

Private Sub lblCarregarMais_Click()

End Sub

Private Sub gridImpressao_DblClick()
    
    Screen.MousePointer = 11

    Do While gridImpressao.Rows > 1
        
        grdMovimentosDisponiveis.Row = 0
        
        gridImpressao.Row = 1
        gridImpressao.Refresh
    
        wProtocoloImpressao = gridImpressao.TextMatrix(gridImpressao.Row, 3)
        wNroCaixaImpressao = gridImpressao.TextMatrix(gridImpressao.Row, 2)
        wDataInicioFechamento = Format(gridImpressao.TextMatrix(gridImpressao.Row, 0), "DD/MM/YYYY")
        wHoraInicioFechamento = Format(gridImpressao.TextMatrix(gridImpressao.Row, 0), "HH:MM")
        wDataFinalFechamento = Format(gridImpressao.TextMatrix(gridImpressao.Row, 1), "DD/MM/YYYY")
        wHoraFinalFechamento = Format(gridImpressao.TextMatrix(gridImpressao.Row, 1), "HH:MM")
        
        
        sql = ("Select Max(CTr_DataInicial)as DataMov,Max(Ctr_Protocolo) as Seq " _
        & "from ControleCaixa where CTR_Supervisor <> 99 and CTr_NumeroCaixa = " & GLB_Caixa)
        
        Call CarregaMovimento(grdMovimentoCaixa, CLng(wProtocoloImpressao))
        
        Call CarregaValoresTransfNumerario(wProtocoloImpressao)
        
            Call NOVO_ImprimeMovimento(grdMovimentoCaixa, "FECHAMENTO DE CAIXA (Reimpressao)", _
                                       wOperadorImpressao, wNroCaixaImpressao, _
                                       wDataInicioFechamento, Format(wHoraInicioFechamento, "HH:MM:SS"), _
                                       wDataFinalFechamento, Format(wHoraFinalFechamento, "HH:MM:SS"), _
                                       CStr(wProtocoloImpressao))
                                       
            Call NOVO_ImprimeTransfNumerario(grdMovimentoCaixa, "TRANSFERENCIA DE NUMERARIO (Reimpressao)", _
                                             wOperadorImpressao, wNroCaixaImpressao, _
                                             wDataInicioFechamento, Format(wHoraInicioFechamento, "HH:MM:SS"), _
                                             wDataFinalFechamento, Format(wHoraFinalFechamento, "HH:MM:SS"), _
                                             CStr(wProtocoloImpressao))
        
        gridImpressao.RemoveItem 1
          
    Loop
    
    ChkModoImpressao.Value = 0
    gridImpressao.Visible = False
    lblMSGImpressao.Visible = False
    Screen.MousePointer = 0
    
End Sub

Private Sub gridImpressao_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        gridImpressao.Rows = 1
        gridImpressao.Visible = False
        lblMSGImpressao.Visible = False
        ChkModoImpressao.Value = 0
        fraFechamentoAnterior.Height = 615
    End If
    
End Sub

Private Sub mskDataFec_GotFocus()
    'mskDataFec.SelStart = 0
    'mskDataFec.SelLength = Len(mskDataFec.Text)
End Sub

''Private Sub mskDataFec_KeyPress(KeyAscii As Integer)
''
''If KeyAscii = 27 Then
''   Unload Me
''   frmControlaCaixa.txtPedido.SetFocus
''End If
''
''If KeyAscii = 13 Then
''   If Trim(mskDataFec.Text) = "" Or mskDataFec.Text = "__/__/____" Then
''        MsgBox "Informe uma data.", vbInformation, Me.Caption
''        mskDataFec.SetFocus
''        Exit Sub
''    ElseIf IsDate(mskDataFec.Text) = False Then
''        MsgBox "Data inválida.", vbCritical, Me.Caption
''        mskDataFec.SetFocus
''        Exit Sub
''    ElseIf CDate(mskDataFec.Text) > Date Then
''        MsgBox "Data maior que data atual.", vbCritical, Me.Caption
''        mskDataFec.SetFocus
''        Exit Sub
''    End If
''
''    Call CarregaAnaliticoVenda(Format(Trim(mskDataFec.Text), "yyyy/mm/dd"))
''
''End If
''
''
''End Sub

Private Sub optAnaliticoVenda_Click()
    'Call CarregaAnaliticoVenda(Format(Date, "yyyy/mm/dd"))
    'fraAlteraDia.Visible = True
End Sub

'Private Sub picAvancar_Click()
'
'   If CDate(mskDataFec.Text) > Date - 1 Then
'        MsgBox "Data maior que data atual.", vbCritical, Me.Caption
'        mskDataFec.SetFocus
'        Exit Sub
'    End If
'
'    mskDataFec.Text = Format(CDate(mskDataFec.Text) + 1, "DD/MM/YYYY")
'    Call CarregaAnaliticoVenda(Format(Trim(mskDataFec.Text), "yyyy/mm/dd"))
'
'End Sub

Private Sub picAvancar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
   frmControlaCaixa.txtPedido.SetFocus
End If
End Sub

Private Sub picVoltar_Click()
'Format(Date - 2, "YYYY/MM/DD")
'    mskDataFec.Text = Format(CDate(mskDataFec.Text) - 1, "DD/MM/YYYY")
'    Call CarregaAnaliticoVenda(Format(Trim(mskDataFec.Text), "yyyy/mm/dd"))
End Sub

Private Sub picVoltar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
   frmControlaCaixa.txtPedido.SetFocus
End If
End Sub
'
'Private Sub txtSenhaSupervisor_GotFocus()
'      txtSenhaSupervisor.Text = ""
'      txtSenhaSupervisor.SelStart = 0
'      txtSenhaSupervisor.SelLength = Len(txtSupervisor.Text)
'End Sub

'Private Sub txtSenhaSupervisor_KeyPress(KeyAscii As Integer)
'
'If KeyAscii = 27 Then
'   txtSupervisor.SetFocus
'End If
'
'If KeyAscii = 13 Then
'
'
'   If rdoDataFechamentoRetaguarda.State = 1 Then
'      rdoDataFechamentoRetaguarda.Close
'   End If
'
'   sql = "Select USU_Senha from UsuarioCaixa where USU_Nome ='" & Trim(txtSupervisor.Text) & "' and USU_TipoUsuario='S' "
'   rdoDataFechamentoRetaguarda.CursorLocation = adUseClient
'   rdoDataFechamentoRetaguarda.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
'
'   If Not rdoDataFechamentoRetaguarda.EOF Then
'
'      If UCase(Trim(rdoDataFechamentoRetaguarda("USU_Senha"))) = UCase(Trim(txtSenhaSupervisor.Text)) Then
'        'If optControleMovimento.Value = True Then
'
'        'Else
'
''            rdoDataFechamentoRetaguarda.Close
'        'End If
'      Else
'          MsgBox "Senha incorreta"
'          'rdoDataFechamentoRetaguarda.Close
'          txtSupervisor.SetFocus
'      End If
'    Else
'        MsgBox "Usúario incorreto"
'        'rdoDataFechamentoRetaguarda.Close
'        txtSupervisor.SetFocus
'    End If
'    If rdoDataFechamentoRetaguarda.State = 1 Then
'      rdoDataFechamentoRetaguarda.Close
'   End If
'End If
'End Sub
'
'Private Sub txtSupervisor_GotFocus()
'      txtSupervisor.Text = ""
'      txtSupervisor.SelStart = 0
'      txtSupervisor.SelLength = Len(txtSupervisor.Text)
'End Sub

'Private Sub txtSupervisor_KeyPress(KeyAscii As Integer)
'
'If KeyAscii = 27 Then
'   Unload Me
'   frmControlaCaixa.txtPedido.SetFocus
'End If
'
'If KeyAscii = 13 Then
'   txtSenhaSupervisor.SetFocus
'End If
'
'End Sub
Sub CarregaMovimentosDisponiveis()
    
    Dim msgCarregarMaisMovimento As String
    
    'grdMovimentosDisponiveis.Visible = True
    grdMovimentosDisponiveis.Rows = 1
    grdMovimentoCaixa.Height = grdMovimentosDisponiveis.Height
    'grdMovimentoCaixa.left = grdMovimentosDisponiveis.left
    grdMovimentoCaixa.top = grdMovimentosDisponiveis.top
    
    'Image1.Width = grdMovimentosDisponiveis.Width

    sql = "select CTR_DataInicial, CTR_DataFinal, CTR_NumeroCaixa, CTR_Protocolo, USU_Nome from controlecaixa,UsuarioCaixa " _
        & "where CTR_Operador = USU_Codigo and CTR_SituacaoCaixa = 'F' and CTR_Supervisor <> 99 and  USU_TipoUsuario = 'O' AND " _
        & "CTR_DataInicial between '" & Format(Date - wDiasCarregaMov, "YYYY/MM/DD") & "' and '" & Format(Date + 3, "YYYY/MM/DD") & "' " _
        & "order by CTR_Protocolo"
       rdoDataFechamentoRetaguarda.CursorLocation = adUseClient
       rdoDataFechamentoRetaguarda.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
       
       msgCarregarMaisMovimento = ""
       
       grdMovimentosDisponiveis.AddItem msgCarregarMaisMovimento & Chr(9) _
                                            & msgCarregarMaisMovimento & Chr(9) _
                                            & msgCarregarMaisMovimento & Chr(9) _
                                            & msgCarregarMaisMovimento & Chr(9) _
                                            & msgCarregarMaisMovimento & Chr(9) _
                                            & msgCarregarMaisMovimento
                                            
       
     
       If Not rdoDataFechamentoRetaguarda.EOF Then
          Do While Not rdoDataFechamentoRetaguarda.EOF
          
            sql = "select sum(MC_Valor) as TotalCaixa from movimentocaixa " _
                & "Where MC_NroCaixa= " & rdoDataFechamentoRetaguarda("CTR_NumeroCaixa") & " " _
                & "and MC_Serie <> '00' and mc_protocolo = " & rdoDataFechamentoRetaguarda("CTR_Protocolo") & " " _
                & "and MC_Grupo like '10%' group by  MC_Protocolo "
                rdoFechamentoGeral.CursorLocation = adUseClient
                rdoFechamentoGeral.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                
                If Not rdoFechamentoGeral.EOF Then
                  wTotalCaixa = rdoFechamentoGeral("TotalCaixa")
                Else
                  wTotalCaixa = 0
                End If
                  
             grdMovimentosDisponiveis.AddItem Format(rdoDataFechamentoRetaguarda("CTR_DataInicial"), "DD/MM/YY HH:MM") & Chr(9) _
                                            & Format(rdoDataFechamentoRetaguarda("CTR_DataFinal"), "DD/MM/YY HH:MM") & Chr(9) _
                                            & Format(rdoDataFechamentoRetaguarda("CTR_NumeroCaixa"), "###00") & Chr(9) _
                                            & Format(rdoDataFechamentoRetaguarda("CTR_Protocolo"), "###00") & Chr(9) _
                                            & Format(rdoDataFechamentoRetaguarda("USU_Nome"), "###00") & Chr(9) _
                                            & Format(wTotalCaixa, "##,###0.00")
                                            
             rdoDataFechamentoRetaguarda.MoveNext
             rdoFechamentoGeral.Close
           Loop
       End If
       rdoDataFechamentoRetaguarda.Close
End Sub




'' - - - ALTERACAO INICIO AQUI
''Private Sub ImprimeMovimento()
''
''     Dim wTotalTransferencia As Double
''
''     wSaldoFinalDinheiro = 0
''     wSaldoFinalCheque = 0
''     wSaldoFinalAVR = 0
''
''
''Screen.MousePointer = 11
''Retorno = Bematech_FI_AbreRelatorioGerencialMFD("01")
''
''     Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
''               "                   REIMPRESSAO                  " & _
''               " Fechamento do Caixa " & Trim(wDataInicioFechamento) & "         Loja " & Format(GLB_Loja, "000") & _
''               "________________________________________________")
''
''
''     For Idx = 1 To grdMovimentoCaixa.Rows - 6 Step 1
''         If (Idx = 25) Or (Idx = 26) Or (Idx = 27) Or (Idx = 28) Or (Idx = 29) Or (Idx = 32) Or (Idx = 33) Or (Idx = 34) Or _
''         (Idx = 37) Or (Idx = 38) Or (Idx = 39) Or (Idx = 40) Or (Idx = 41) Then
''
''         Retorno = Bematech_FI_UsaRelatorioGerencialMFD(left(grdMovimentoCaixa.TextMatrix(Idx, 0) & Space(23), 23) & _
''               right(Space(20) & Format(grdMovimentoCaixa.TextMatrix(Idx, 1), "###,###,##0.00"), 20) & _
''               right(Space(5) & grdMovimentoCaixa.TextMatrix(Idx, 2), 5))
''         Else
''
''          Retorno = Bematech_FI_UsaRelatorioGerencialMFD(left(grdMovimentoCaixa.TextMatrix(Idx, 0) & Space(23), 23) & _
''               right(Space(20) & Format(grdMovimentoCaixa.TextMatrix(Idx, 1), "###,###,##0.00"), 20) & "     ")
''         End If
''     Next Idx
''
''     wMovimentoPeriodo = Format(grdMovimentoCaixa.TextMatrix(13, 1), "###,###,##0.00")
''     'wMovimentoPeriodo = wMovimentoPeriodo - wtotalGarantia
''     'wMovimentoPeriodo = Format(wMovimentoPeriodo + CDbl(grdMovimentoCaixa.TextMatrix(46, 2)), "###,###,##0.00")
''     wSaldoAnterior = Format(grdMovimentoCaixa.TextMatrix(45, 1), "###,###,##0.00")
''     'wMovimentoPeriodo = (wMovimentoPeriodo - wSaldoAnterior)
''     wSaldoFinalDinheiro = Format(grdMovimentoCaixa.TextMatrix(1, 3), "###,###,##0.00")
''     wSaldoFinalDinheiro = (wSaldoFinalDinheiro + Format(grdMovimentoCaixa.TextMatrix(44, 3), "###,###,##0.00"))
''     wSaldoFinalCheque = Format(grdMovimentoCaixa.TextMatrix(2, 3))
''     wSaldoFinalCheque = (wSaldoFinalCheque + Format(grdMovimentoCaixa.TextMatrix(45, 3), "###,###,##0.00"))
''     'wSaldoFinalAVR = Format(grdMovimentoCaixa.TextMatrix(9, 3))
''     'wSaldoFinalAVR = (wSaldoFinalAVR + Format(grdMovimentoCaixa.TextMatrix(17, 3), "###,###,##0.00"))
''     wTotalTransferencia = CDbl(grdMovimentoCaixa.TextMatrix(13, 2)) + CDbl(grdMovimentoCaixa.TextMatrix(46, 2))
''     wTotalTransferencia = wTotalTransferencia + CDbl(grdMovimentoCaixa.TextMatrix(17, 2))
''     wTotalTransferencia = wTotalTransferencia - wtotalGarantia
''
''
''     Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
''        "               MOVIMENTO DE CAIXA               " & _
''        "________________________________________________" & _
''        "SALDO ANTERIOR >> " & right(Space(30) & Format("", ""), 30) & _
''        "  DINHEIRO        " & right(Space(30) & Format(grdMovimentoCaixa.TextMatrix(44, 1), "###,###,##0.00"), 30) & _
''        "  CHEQUE          " & right(Space(30) & Format(grdMovimentoCaixa.TextMatrix(45, 1), "###,###,##0.00"), 30) & _
''        "  TOTAL           " & right(Space(30) & Format(wTotalSaldo, "###,###,##0.00"), 30) & _
''        "MOVIMENTO PERIODO " & right(Space(30) & Format(wMovimentoPeriodo, "###,###,##0.00"), 30) & _
''        "REFORCO           " & right(Space(30) & Format(grdMovimentoCaixa.TextMatrix(17, 1), "###,###,##0.00"), 30) & _
''        "TRANSFERENCIA NUM." & right(Space(30) & Format(wTotalTransferencia, "###,###,##0.00"), 30) & _
''        "GARANTIA ESTEN.   " & right(Space(30) & Format(wtotalGarantia, "###,###,##0.00"), 30))
''
''
''      Retorno = Bematech_FI_UsaRelatorioGerencialMFD("SALDO DO CAIXA >>                               " & _
''        "  Dinheiro        " & right(Space(30) & Format(wSaldoFinalDinheiro, "###,###,##0.00"), 30) & _
''        "  Cheque          " & right(Space(30) & Format(wSaldoFinalCheque, "###,###,##0.00"), 30) & _
''        "  SALDO FINAL     " & right(Space(30) & Format((wSaldoFinalDinheiro + wSaldoFinalCheque), "###,###,##0.00"), 30) & _
''        "________________________________________________")
''
''
''      Retorno = Bematech_FI_UsaRelatorioGerencialMFD(left("Caixa Nro.   " & wNroCaixaImpressao & Space(48), 48) & _
''         left("Operador     " & Trim("") & Space(48), 48) & _
''         left("Data Inicial " & Trim(wDataInicioFechamento) & " " & Trim(wHoraInicioFechamento) & Space(48), 48) & _
''         left("Data Final   " & Trim(wDataFinalFechamento) & " " & Trim(wHoraFinalFechamento) & Space(48), 48) & _
''         left("Protocolo    " & wProtocoloImpressao & Space(48), 48))
''
''
''      Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
''         "                                                " & _
''         "                                                " & _
''         "                                                " & _
''         " ______________________                         " & _
''         " " & left(Trim("") & Space(48), 48) & _
''         "                                                " & _
''         "                                                ")
''
''
''        Retorno = Bematech_FI_FechaRelatorioGerencial()
''
''      Screen.MousePointer = 0
''
''
''End Sub
 

 

 

 
''Private Sub ImprimeMovimento00()
''Dim wImprime As String
''Dim wImprime2 As String
''Dim wTotal00 As Double
''wTotal00 = 0
''
''sql = ("Select Serie, sum(totalnota) as TotalSerieNota, count(Serie) as QtdeSerie from nfcapa Where ecf = " & GLB_ECF & "" _
''     & " and  TipoNota = 'V' and Serie = '00' and  DataEmi = '" & Format(Date, "yyyy/mm/dd") _
''     & "' " & " and Protocolo = " & GLB_CTR_Protocolo & " group by Serie ")
''     rdoCapa.CursorLocation = adUseClient
''     rdoCapa.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
''
''If Not rdoCapa.EOF Then
''   wTotal00 = rdoCapa("TotalSerieNota")
''End If
''
''rdoCapa.Close
''    Screen.MousePointer = 11
''         Retorno = Bematech_FI_AbreRelatorioGerencialMFD("01")
''
''          Retorno = Bematech_FI_UsaRelatorioGerencialMFD("                                                " & _
''                    "              REIMPRESSAO CONTROLE X            " & _
''                    " Loja " & left(GLB_Loja & Space(4), 4) & Space(28) & Format(Date, "dd/mm/yyyy") & _
''                    "                                                " & _
''                    " Valor Total:                     " & right(Space(10) & Format(wTotal00, "###,###,##0.00"), 14) & _
''                    "________________________________________________" & _
''                    "                                                " & _
''                    left("Caixa Nro.   " & frmControlaCaixa.lblNroCaixa.Caption & Space(48), 48) & _
''                    left("Supervisor   " & Trim("") & Space(48), 48) & _
''                    left("Data Inicial " & Format(Date, "dd/mm/yyyy") & Space(48), 48) & _
''                    left("Data Final   " & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS") & Space(48), 48) & _
''                    left("Protocolo    " & GLB_CTR_Protocolo & Space(48), 48))
''
''
''          Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
''                    "                                                " & _
''                    "                                                " & _
''                    "                                                " & _
''                    " ______________________                         " & _
''                    " " & left(Trim("") & Space(47), 47) & _
''                    "                                                " & _
''                    "                                                ")
''
''
''        Retorno = Bematech_FI_FechaRelatorioGerencial()
''
''      Screen.MousePointer = 0
''
''
''End Sub
 
 
 
''Private Sub ImprimeTransfNumerario()
''
''
''    Screen.MousePointer = 11
''    Retorno = Bematech_FI_AbreRelatorioGerencialMFD("01")
''
''    Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
''                   "                 REIMPRESSAO                    " & _
''                   "TRANSFERENCIA DE NUMERARIO" & right(Space(22) & "Nro.  " & wProtocoloImpressao, 22) & _
''                   left("Loja " & Format(GLB_Loja, "000") & Space(10), 10) & _
''                   right(Space(38) & Trim(wDataFinalFechamento) & " " & Trim(wHoraFinalFechamento) & " " & Format(Time, "HH:MM:SS"), 38) & _
''                   "________________________________________________" & _
''                   "    USO INTERNO          SEM VALOR COMERCIAL    ")
''
''   Retorno = Bematech_FI_UsaRelatorioGerencialMFD("------------------------------------------------" & _
''                   "DINHEIRO /P TESOU." & right(Space(30) & Format(wTNDinheiro, "###,###,##0.00"), 30) & _
''                   "CHEQUE /P TESOU.  " & right(Space(30) & Format(wTNCheque, "###,###,##0.00"), 30) & _
''                   "VISA              " & right(Space(30) & Format(wTNVisa, "###,###,##0.00"), 30) & _
''                   "MASTERCARD        " & right(Space(30) & Format(wTNRedecard, "###,###,##0.00"), 30) & _
''                   "AMEX              " & right(Space(30) & Format(wTNAmex, "###,###,##0.00"), 30) & _
''                   "BNDS              " & right(Space(30) & Format(wTNBNDES, "###,###,##0.00"), 30) & _
''                   "REDE SHOP         " & right(Space(30) & Format(wTNRedeShop, "###,###,##0.00"), 30) & _
''                   "VISA ELEC.        " & right(Space(30) & Format(wTNVisaEletron, "###,###,##0.00"), 30))
''
''    Retorno = Bematech_FI_UsaRelatorioGerencialMFD( _
''                   "HIPERCARD         " & right(Space(30) & Format(wTNHiperCard, "###,###,##0.00"), 30) & _
''                   "DEPOSITO          " & right(Space(30) & Format(wTNDeposito, "###,###,##0.00"), 30) & _
''                   "NOTA CREDITO      " & right(Space(30) & Format(wTNNotaCredito, "###,###,##0.00"), 30) & _
''                   "OUTRAS DESPESAS   " & right(Space(30) & Format(wTNConducao, "###,###,##0.00"), 30) & _
''                   "DESPESA LOJA      " & right(Space(30) & Format(wTNDespLoja, "###,###,##0.00"), 30) & _
''                   "OUTRAS DESPESAS   " & right(Space(30) & Format(wTNOutros, "###,###,##0.00"), 30) & _
''                   "                                                " & _
''                   "TOTAL             " & right(Space(30) & Format(wTNTotal, "###,###,##0.00"), 30))
''
''    Retorno = Bematech_FI_UsaRelatorioGerencialMFD( _
''                   "                  " & right(Space(30) & "", 30) & _
''                   "GARANTIA ESTEN.   " & right(Space(30) & Format(wtotalGarantia, "###,###,##0.00"), 30) & _
''                   "ENTRADA FINANCIADA" & right(Space(30) & Format(wTNFinanciado, "###,###,##0.00"), 30) & _
''                   "ENTRADA FATURADA  " & right(Space(30) & Format(wTNFaturado, "###,###,##0.00"), 30))
''
''    Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
''                    "                                                " & _
''                    "                                                " & _
''                    "                                                " & _
''                    "                                                " & _
''                    " ______________________                         " & _
''                    " " & left("" & Space(47), 47) & _
''                    "                                                " & _
''                    "                                                ")
''        Retorno = Bematech_FI_FechaRelatorioGerencial()
''
''      Screen.MousePointer = 0
''
''End Sub

'' - - - ALTERACAO FIM AQUI

Sub CarregaControleMovimento()

    grdMovimentoCaixa.Visible = False
    grdAnaliticoVenda.Visible = False
    
    Call CarregaMovimentosDisponiveis
          
End Sub

Sub CarregaAnaliticoVenda(ByVal wData As Date)
          
          Dim wUltimaSerie As String
          Dim wSerieAnterior As String
          Dim wNF As Long
          Dim wTotalNf As Double
          Dim wNFAnterior As Long
          Dim wtotalSerie As Double
          Dim wTotalGeral As Double
          
          'fraAlteraDia.Visible = True

          wUltimaSerie = " "
          wNF = 0
          wTotalNf = 0
          wNFAnterior = 0
          wSerieAnterior = " "
          wtotalSerie = 0
          wTotalGeral = 0
          
          
          'mskDataFec = Format(wData, "dd/mm/yyyy")
          
          'lblCabec2.Caption = "Analítico de Venda"
          'txtSupervisor.Enabled = False
          'txtSenhaSupervisor.Enabled = False
          'grdAnaliticoVenda.Visible = True
          grdAnaliticoVenda.Rows = 1
          grdAnaliticoVenda.Height = grdMovimentosDisponiveis.Height
          'grdAnaliticoVenda.left = grdMovimentosDisponiveis.left
          'grdAnaliticoVenda.top = grdMovimentosDisponiveis.top
          'optControleMovimento.Visible = False
          'optAnaliticoVenda.Visible = False
          
          sql = ""
          'sql = "select NF,Serie,TOTALNOTA from nfcapa,movimentocaixa " & _
          '"WHERE NF = MC_Documento and mc_serie = serie and " & _
          '" tiponota = 'V' and dataemi = '" & Format(wData, "yyyy/mm/dd") & "' and MC_Grupo like '10%'  " _
          '    & "order by mc_grupo,Serie,NF"
          sql = "select NF,Serie,TOTALNOTA from nfcapa where tiponota = 'V' and dataemi = '" & Format(wData, "yyyy/mm/dd") & "' " _
              & "order by Serie,NF"
              
          RsCarimbo.CursorLocation = adUseClient
          RsCarimbo.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
          
          If Not RsCarimbo.EOF Then
               Do While Not RsCarimbo.EOF
                  
                   wNF = RsCarimbo("NF")
                   wTotalNf = RsCarimbo("totalnota")
                    
                    sql = ""
                    sql = "select (rtrim(ltrim(MO_Descricao)) + ' ' + rtrim(ltrim(mc_SubGrupo))) as Descricao, " _
                        & "MC_Valor From MovimentoCaixa, Modalidade " _
                        & "where MC_Documento = " & RsCarimbo("nf") & " and " _
                        & "mc_serie = '" & Trim(RsCarimbo("serie")) & "' " _
                        & "and MC_Grupo = mo_grupo and mc_grupo like '10%' order by mo_descricao "
                    rsComplementoVenda.CursorLocation = adUseClient
                    rsComplementoVenda.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                        
                    Do While Not rsComplementoVenda.EOF
                        If wSerieAnterior <> RsCarimbo("serie") Then
                         If wtotalSerie > 0 Then
                        
                           grdAnaliticoVenda.AddItem ""
                           grdAnaliticoVenda.AddItem Chr(9) & Chr(9) & "TOTAL " _
                                               & UCase(wSerieAnterior) & Chr(9) _
                                               & Format(wtotalSerie, "##,###0.00")
                            wtotalSerie = 0
                            grdAnaliticoVenda.AddItem ""
                          End If
                            wSerieAnterior = RsCarimbo("serie")
                            wUltimaSerie = RsCarimbo("serie")
                        End If
                        
                        If wNF <> wNFAnterior Then
                           wNFAnterior = wNF
                           grdAnaliticoVenda.AddItem wNF & Chr(9) & wUltimaSerie & Chr(9) _
                                               & UCase(rsComplementoVenda("Descricao")) & Chr(9) _
                                               & Format(rsComplementoVenda("mc_valor"), "##,###0.00")
                        Else
                           grdAnaliticoVenda.AddItem " " & Chr(9) & " " & Chr(9) _
                                               & UCase(rsComplementoVenda("Descricao")) & Chr(9) _
                                               & Format(rsComplementoVenda("mc_valor"), "##,###0.00")
                        End If
                        wtotalSerie = wtotalSerie + rsComplementoVenda("mc_valor")
                        wTotalGeral = wTotalGeral + rsComplementoVenda("mc_valor")
                        rsComplementoVenda.MoveNext
                    Loop
                    grdAnaliticoVenda.AddItem Chr(9) & Chr(9) & "TOTAL " & Chr(9) & Format(wTotalNf, "##,###0.00")
                    grdAnaliticoVenda.AddItem ""
                    rsComplementoVenda.Close
                    RsCarimbo.MoveNext
                Loop
                
                    If wtotalSerie > 0 Then
                        grdAnaliticoVenda.AddItem Chr(9) & Chr(9) & "TOTAL " _
                                               & UCase(wSerieAnterior) & Chr(9) _
                                               & Format(wtotalSerie, "##,###0.00")
                        wtotalSerie = 0
                        grdAnaliticoVenda.AddItem ""
                     End If
                        grdAnaliticoVenda.AddItem Chr(9) & Chr(9) & "TOTAL GERAL " & Chr(9) _
                                               & Format(wTotalGeral, "##,###0.00")
                lblCabec2.Visible = True
                grdAnaliticoVenda.Visible = True
            Else
                lblCabec2.Visible = False
                grdAnaliticoVenda.Visible = False
            End If
            RsCarimbo.Close
          
End Sub

Sub ImprimeAnaliticoVenda()
 
 
Screen.MousePointer = 11
    Retorno = Bematech_FI_AbreRelatorioGerencialMFD("01")
 
    Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
                   "          RELATORIO ANALITICO DE VENDA          " & _
                   left("Loja " & Format(GLB_Loja, "000") & Space(10), 10) & _
                   right(Space(38) & (Format(Trim(""), "dd/mm/yyyy")), 38) & _
                   "________________________________________________")
    
    Retorno = Bematech_FI_UsaRelatorioGerencialMFD("                                                " & _
                   left("NF " & Space(10), 10) & left("SERIE" & Space(8), 8) & _
                   left("FORMA PAGAMENTO" & Space(20), 20) & left("VALOR " & Space(10), 10) & _
                   "                                                ")
 
     For Idx = 1 To grdAnaliticoVenda.Rows - 1 Step 1
     
     Retorno = Bematech_FI_UsaRelatorioGerencialMFD(left(grdAnaliticoVenda.TextMatrix(Idx, 0) & Space(10), 10) & _
                   left(grdAnaliticoVenda.TextMatrix(Idx, 1) & Space(8), 8) & _
                   left(grdAnaliticoVenda.TextMatrix(Idx, 2) & Space(20), 20) & _
                   right(Space(10) & Format(grdAnaliticoVenda.TextMatrix(Idx, 3), "###,###,##0.00"), 10))
     Next Idx
    
     Retorno = Bematech_FI_FechaRelatorioGerencial()
 
     Screen.MousePointer = 0
     
End Sub


