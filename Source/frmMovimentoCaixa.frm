VERSION 5.00
Object = "{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7u.ocx"
Begin VB.Form frmMovimentoCaixa 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Movimento do Caixa"
   ClientHeight    =   10290
   ClientLeft      =   5490
   ClientTop       =   1275
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10290
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Begin VSFlex7UCtl.VSFlexGrid grdMovimentoCaixa 
      Height          =   6720
      Left            =   255
      TabIndex        =   0
      Top             =   375
      Width           =   4755
      _cx             =   8387
      _cy             =   11853
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmMovimentoCaixa.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
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
   Begin VSFlex7UCtl.VSFlexGrid grdMovimento00 
      Height          =   735
      Left            =   255
      TabIndex        =   4
      Top             =   6870
      Width           =   4755
      _cx             =   8387
      _cy             =   1296
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
      ForeColor       =   4210752
      BackColorFixed  =   0
      ForeColorFixed  =   16777215
      BackColorSel    =   3421236
      ForeColorSel    =   16777215
      BackColorBkg    =   0
      BackColorAlternate=   14737632
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
      Rows            =   3
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmMovimentoCaixa.frx":0070
      ScrollTrack     =   0   'False
      ScrollBars      =   3
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
   Begin VB.Label lblCabec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consulta Caixa"
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
      Height          =   285
      Left            =   540
      TabIndex        =   3
      Top             =   90
      Width           =   1740
   End
   Begin VB.Label sklDataMovimento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99/99/9999"
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
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   960
   End
   Begin VB.Label sklMovimentoCaixa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data do Movimento :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   840
      TabIndex        =   1
      Top             =   420
      Width           =   1965
   End
End
Attribute VB_Name = "frmMovimentoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wSubTotal As Double
Dim wSubTotal2 As Double
Dim wTotalQtde As Long
Dim wControlacor As Long
Dim wConfigCor As Long
Dim GuardaSequencia As Long
Dim wSaldoFinal As Double

Dim wSangria As Double
Dim wMovimentoDoPeriodo As Double
Dim wEntradanoCaixa As Double

Dim wQuantidade As Double
Dim wDesconto As Double
Dim wPrecoUnitario As Double
Dim wTotalTipoNota As Double
Dim wVenda As Double
Dim wCancelamento As Double
Dim wDevolucao As Double
Dim wTR As Double
Dim wSubTotalEntfin As Double
Dim wSubTotalEntFat As Double

Dim I As Long
Dim sql As String
Dim Cor As String
Dim Cor1 As String
Dim Cor2 As String
Dim Cor3 As String
Dim wData As String

Private Sub cmdRetornar_Click()
 Unload Me
End Sub

Private Sub cmbSair_Click()
 Unload Me
End Sub

Private Sub Form_Load()
   left = 100
  top = 2880
  
  Call AjustaTela(frmMovimentoCaixa)
  
 grdMovimento00.Visible = False
  
  
 wSubTotal = 0
 wSubTotal2 = 0
 wSubTotalEntfin = 0
 wSubTotalEntFat = 0
 wControlacor = 0
 wConfigCor = 0
 sql = ""
 Cor = ""
 Cor1 = ""
 Cor2 = ""
 Cor3 = ""

 wTotalTipoNota = 0
 wVenda = 0
 wCancelamento = 0
 wDevolucao = 0
 wTR = 0

  
  Cor1 = &HC0E0FF
  grdMovimento00.Rows = 1
  grdMovimentoCaixa.Rows = 1
  grdMovimentoCaixa.AddItem "Venda Bruta"
  Call PintaGridZebrado
  grdMovimentoCaixa.AddItem "Devolucao"
  Call PintaGridZebrado
  grdMovimentoCaixa.AddItem "Venda Liquida"
  Call PintaGridZebrado
  grdMovimentoCaixa.AddItem "Canceladas"
  Call PintaGridZebrado
  Cor1 = &HC0FFFF
  grdMovimentoCaixa.AddItem "Nota de Credito"
  Call PintaGridZebrado
  grdMovimentoCaixa.AddItem "Dinheiro"
  Call PintaGridZebrado
  grdMovimentoCaixa.AddItem "Cheque"
  Call PintaGridZebrado
  grdMovimentoCaixa.AddItem "Visa"
  Call PintaGridZebrado
  grdMovimentoCaixa.AddItem "MasterCard"
  Call PintaGridZebrado
  grdMovimentoCaixa.AddItem "Amex"
  Call PintaGridZebrado
  grdMovimentoCaixa.AddItem "BNDES"
  Call PintaGridZebrado
  grdMovimentoCaixa.AddItem "T E F"
  Call PintaGridZebrado
  Cor1 = &HC0FFC0
  grdMovimentoCaixa.AddItem "*** Sub Total " & Chr(9) & Chr(vbKeyTab) & "Entrada"
  Call PintaGridZebrado
  Cor1 = &HC0FFFF
  grdMovimentoCaixa.AddItem "Faturado"
  Call PintaGridZebrado
  grdMovimentoCaixa.AddItem "Financiado"
  Call PintaGridZebrado
  grdMovimentoCaixa.AddItem "A Vista Receber"
  Call PintaGridZebrado
  Cor1 = &HC0FFC0
  grdMovimentoCaixa.AddItem "*** T O T A L"
  Call PintaGridZebrado
  Cor1 = &HC0FFFF
  grdMovimentoCaixa.AddItem " "
  Call PintaGridZebrado
  grdMovimentoCaixa.AddItem "CF"
  Call PintaGridZebrado
  grdMovimentoCaixa.AddItem "SN"
  Call PintaGridZebrado
  grdMovimentoCaixa.AddItem "S1"
  Call PintaGridZebrado
  grdMovimentoCaixa.AddItem "D1"
  Call PintaGridZebrado
  grdMovimentoCaixa.AddItem "S2"
  Call PintaGridZebrado
  grdMovimentoCaixa.AddItem "NE"
  Call PintaGridZebrado
  grdMovimentoCaixa.AddItem "*** T O T A L"
  Call PintaGridZebrado
  Cor1 = &HC0FFFF
  grdMovimentoCaixa.AddItem " "
  Call PintaGridZebrado
  grdMovimentoCaixa.AddItem "TR Emitida"
  Call PintaGridZebrado
  
  wVenda = 0
  wCancelamento = 0
  wDevolucao = 0
  wTR = 0
     
     
If rdoDataFechamentoRetaguarda.State = 1 Then
    rdoDataFechamentoRetaguarda.Close
End If
     
     
sql = ("Select Max(CTr_DataInicial)as DataMov,Max(Ctr_Protocolo) as Seq from ControleCaixa where CTR_Supervisor <> 99 and CTr_NumeroCaixa = " & GLB_Caixa & "")
       rdoDataFechamentoRetaguarda.CursorLocation = adUseClient
       rdoDataFechamentoRetaguarda.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
     
    If rdoDataFechamentoRetaguarda.EOF Then
       MsgBox "Não foi possível abrir a tabela FechamentoRetaguarda. Contate o CPD.", vbCritical, "Atenção"
       Exit Sub
    End If
     
sklDataMovimento = Format(rdoDataFechamentoRetaguarda("DataMov"), "dd/mm/yyyy")
sql = ("Select totalnota,Numeroped,* from nfcapa Where ecf = " & GLB_ECF & " and Protocolo = " & rdoDataFechamentoRetaguarda("seq") _
     & " and TipoNota <> 'PA'  and Serie <> '00' and  DataEmi = '" & Format(rdoDataFechamentoRetaguarda("datamov"), "yyyy/mm/dd") & "' ")
     rdoItensVenda.CursorLocation = adUseClient
     rdoItensVenda.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
     
     
  If Not rdoItensVenda.EOF Then
      Do While Not rdoItensVenda.EOF
        wTotalTipoNota = rdoItensVenda("totalnota")
        
           If Trim(rdoItensVenda("TipoNota")) = "V" Then
              wVenda = wVenda + wTotalTipoNota
           ElseIf Trim(rdoItensVenda("TipoNota")) = "E" Then
                  wDevolucao = wDevolucao + wTotalTipoNota
           ElseIf Trim(rdoItensVenda("TipoNota")) = "C" Then
                  wCancelamento = wCancelamento + wTotalTipoNota
           ElseIf Trim(rdoItensVenda("TipoNota")) = "T" Then
                  wTR = wTR + wTotalTipoNota
           End If
   
        rdoItensVenda.MoveNext
      Loop
        grdMovimentoCaixa.TextMatrix(1, 1) = Format((wVenda), "###,###,###,##0.00")
        grdMovimentoCaixa.TextMatrix(2, 1) = Format((wDevolucao), "###,###,###,##0.00")
        grdMovimentoCaixa.TextMatrix(4, 1) = Format((wCancelamento), "###,###,###,##0.00")
        grdMovimentoCaixa.TextMatrix(27, 1) = Format((wTR), "###,###,###,##0.00")
        
        If grdMovimentoCaixa.TextMatrix(2, 1) = "0,00" Then
           grdMovimentoCaixa.TextMatrix(3, 1) = Format(grdMovimentoCaixa.TextMatrix(1, 1), "###,###,###,##0.00")
        Else
           grdMovimentoCaixa.TextMatrix(3, 1) = Format(grdMovimentoCaixa.TextMatrix(1, 1) - grdMovimentoCaixa.TextMatrix(2, 1), "###,###,###,##0.00")
        End If
    End If
    rdoItensVenda.Close
  wSubTotal = 0
  
 
 sql = ("select mc_Grupo,sum(MC_Valor) as TotalModalidade,Count(*) as Quantidade from movimentocaixa" _
       & " Where MC_NumeroEcf = " & GLB_ECF & " and MC_NroCaixa=" & GLB_Caixa & " and MC_Protocolo=" & GLB_CTR_Protocolo _
       & " and MC_Data >='" & Format(rdoDataFechamentoRetaguarda("DataMov"), "yyyy/mm/dd") & "' and  MC_Serie <> '00' and (MC_Grupo like '10%' or MC_Grupo like '11%'" _
       & " or MC_Grupo like '50%' or MC_Grupo like '20%') and MC_TipoNota <> 'C' group by mc_grupo")
       rdoFormaPagamento.CursorLocation = adUseClient
       rdoFormaPagamento.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
       
 wData = Format(rdoDataFechamentoRetaguarda("datamov"), "yyyy/mm/dd")

  If Not rdoFormaPagamento.EOF Then
     Do While Not rdoFormaPagamento.EOF
        If rdoFormaPagamento("MC_Grupo") = "10101" Then
           grdMovimentoCaixa.TextMatrix(6, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10201" Then
           grdMovimentoCaixa.TextMatrix(7, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("mc_grupo") = "10301" Then
           grdMovimentoCaixa.TextMatrix(8, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10302" Then
           grdMovimentoCaixa.TextMatrix(9, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10303" Then
           grdMovimentoCaixa.TextMatrix(10, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10304" Then
           grdMovimentoCaixa.TextMatrix(11, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10203" Then
           grdMovimentoCaixa.TextMatrix(12, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10701" Then
           grdMovimentoCaixa.TextMatrix(5, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10501" Then
           grdMovimentoCaixa.TextMatrix(14, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal2 = (wSubTotal2 + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10601" Then
           grdMovimentoCaixa.TextMatrix(15, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal2 = (wSubTotal2 + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10204" Then
           grdMovimentoCaixa.TextMatrix(16, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal2 = (wSubTotal2 + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "11004" Then
           wSubTotalEntFat = (wSubTotalEntFat + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "11005" Then
           wSubTotalEntfin = (wSubTotalEntfin + rdoFormaPagamento("TotalModalidade"))
        End If
        
        rdoFormaPagamento.MoveNext
        grdMovimentoCaixa.TextMatrix(13, 1) = Format(wSubTotal, "###,###,###,##0.00")
        grdMovimentoCaixa.TextMatrix(14, 2) = Format(wSubTotalEntFat, "###,###,###,##0.00")
        grdMovimentoCaixa.TextMatrix(15, 2) = Format(wSubTotalEntfin, "###,###,###,##0.00")
        grdMovimentoCaixa.TextMatrix(17, 1) = Format((wSubTotal + wSubTotal2), "###,###,###,##0.00")
 
     Loop
  End If
  rdoFormaPagamento.Close
  wSubTotal = 0

 sql = ("Select Serie, sum(totalnota) as TotalSerieNota, count(Serie) as QtdeSerie from nfcapa Where ecf = " & GLB_ECF & "" _
     & " and  TipoNota not in ('PD','PA','C') and Serie <> '00' and  DataEmi = '" & Format(rdoDataFechamentoRetaguarda("datamov"), "yyyy/mm/dd") _
     & "' " & " and Protocolo = " & rdoDataFechamentoRetaguarda("seq") & " group by Serie ")
     rdoCapa.CursorLocation = adUseClient
     rdoCapa.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
     
     

      If Not rdoCapa.EOF Then
         Do While Not rdoCapa.EOF

                  If rdoCapa("Serie") Like GLB_SerieCF & "*" Then
                     grdMovimentoCaixa.TextMatrix(19, 1) = Format(rdoCapa("TotalSerieNota"), "###,###,###,##0.00")
                     wSubTotal = (wSubTotal + rdoCapa("TotalSerieNota"))
                     grdMovimentoCaixa.TextMatrix(19, 2) = Format(rdoCapa("QtdeSerie"), "###0")
                  ElseIf rdoCapa("Serie") = "SN" Then
                         grdMovimentoCaixa.TextMatrix(20, 1) = Format(rdoCapa("TotalSerieNota"), "###,###,###,##0.00")
                         wSubTotal = (wSubTotal + rdoCapa("TotalSerieNota"))
                         grdMovimentoCaixa.TextMatrix(20, 2) = Format(rdoCapa("QtdeSerie"), "###0")
                  ElseIf rdoCapa("Serie") = "S1" Then
                         grdMovimentoCaixa.TextMatrix(21, 1) = Format(rdoCapa("TotalSerieNota"), "###,###,###,##0.00")
                         wSubTotal = (wSubTotal + rdoCapa("TotalSerieNota"))
                         grdMovimentoCaixa.TextMatrix(21, 2) = Format(rdoCapa("QtdeSerie"), "###0")
                  ElseIf rdoCapa("Serie") = "D1" Then
                         grdMovimentoCaixa.TextMatrix(22, 1) = Format(rdoCapa("TotalSerieNota"), "###,###,###,##0.00")
                         wSubTotal = (wSubTotal + rdoCapa("TotalSerieNota"))
                         grdMovimentoCaixa.TextMatrix(22, 2) = Format(rdoCapa("QtdeSerie"), "###0")
                  ElseIf rdoCapa("Serie") = "S2" Then
                         grdMovimentoCaixa.TextMatrix(23, 1) = Format(rdoCapa("TotalSerieNota"), "###,###,###,##0.00")
                         wSubTotal = (wSubTotal + rdoCapa("TotalSerieNota"))
                         grdMovimentoCaixa.TextMatrix(23, 2) = Format(rdoCapa("QtdeSerie"), "###0")
                  ElseIf rdoCapa("Serie") = "NE" Then
                         grdMovimentoCaixa.TextMatrix(24, 1) = Format(rdoCapa("TotalSerieNota"), "###,###,###,##0.00")
                         wSubTotal = (wSubTotal + rdoCapa("TotalSerieNota"))
                         grdMovimentoCaixa.TextMatrix(24, 2) = Format(rdoCapa("QtdeSerie"), "###0")

                  End If
                  
         rdoCapa.MoveNext
         Loop
         grdMovimentoCaixa.TextMatrix(25, 1) = Format(wSubTotal, "###,###,###,##0.00")
      End If
        
  rdoCapa.Close
  
sql = ("Select Serie, sum(totalnota) as TotalSerieNota, count(Serie) as QtdeSerie from nfcapa Where ecf = " & GLB_ECF & "" _
     & " and  TipoNota = 'V' and Serie = '00' and  DataEmi = '" & Format(Date, "yyyy/mm/dd") _
     & "' " & " and Protocolo = " & GLB_CTR_Protocolo & " group by Serie ")
     
'  SQL = ("Select Serie, sum(totalnota) as TotalSerieNota, count(Serie) as QtdeSerie from nfcapa Where ecf = " & GLB_ECF & "" _
'     & " and  TipoNota <> 'PD' and TipoNota <> 'PA' and Serie = '00' and  DataEmi = '" & Format(rdoDataFechamentoRetaguarda("datamov"), "yyyy/mm/dd") _
'     & "' " & " and Protocolo = " & rdoDataFechamentoRetaguarda("seq") & " group by Serie ")
     rdoCapa.CursorLocation = adUseClient
     rdoCapa.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
     grdMovimento00.AddItem "  "
     grdMovimento00.AddItem "00"



  If Not rdoCapa.EOF Then
 '   grdMovimento00.TextMatrix(2, 0) = "00"
     grdMovimento00.TextMatrix(2, 1) = Format(rdoCapa("TotalSerieNota"), "###,###,###,##0.00")
     grdMovimento00.TextMatrix(2, 2) = Format(rdoCapa("QtdeSerie"), "###0")
  Else
     grdMovimento00.TextMatrix(2, 1) = ""
     grdMovimento00.TextMatrix(2, 2) = ""
  End If

  rdoCapa.Close

rdoDataFechamentoRetaguarda.Close
'   wData = Format(Date, "yyyy/mm/dd")
'    sklDataMovimento = wData
'''
'''  SQL = ("Select Max(CTr_DataInicial)as DataMov,Max(Ctr_Protocolo) as Seq " _
'''       & "from ControleCaixa where CTr_NumeroCaixa = " & GLB_Caixa)
'''        rdoDataFechamento.CursorLocation = adUseClient
'''        rdoDataFechamento.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
'''  sklDataMovimento = Format(rdoDataFechamento("DataMov"), "yyyy/mm/dd")
'''
'''  rdoDataFechamento.Close


End Sub




Private Sub grdB_Click()

End Sub



Private Sub grdMovimento00_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
       grdMovimento00.Visible = False
       grdMovimentoCaixa.SetFocus
End If
End Sub

Private Sub grdMovimentoCaixa_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF9 Then
   grdMovimento00.Visible = True
   grdMovimento00.SetFocus
End If

End Sub

Private Sub grdMovimentoCaixa_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
       Unload Me
       frmControlaCaixa.txtPedido.SetFocus
End If
    
End Sub

Private Sub sklDataMovimento_DragDrop(Source As Control, X As Single, Y As Single)
'    wData = Format(Date, "yyyy/mm/dd")
'    sklDataMovimento = wData

  sql = ("Select Max(CTr_DataInicial)as DataMov,Max(Ctr_Protocolo) as Seq " _
       & "from ControleCaixa where CTR_Supervisor <> 99 and CTr_NumeroCaixa = " & GLB_Caixa)
        rdoDataFechamento.CursorLocation = adUseClient
        rdoDataFechamento.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
  sklDataMovimento = Format(rdoDataFechamento("DataMov"), "dd/mm/yyyy")
  
  rdoDataFechamento.Close

End Sub

Sub PintaGridZebrado()
   ' Cor = Cor1
   ' grdMovimentoCaixa.Row = grdMovimentoCaixa.Rows - 1
   ' grdMovimentoCaixa.Col = 0
   ' grdMovimentoCaixa.ColSel = 2
   ' grdMovimentoCaixa.FillStyle = flexFillRepeat
   ' grdMovimentoCaixa.CellBackColor = Cor
   ' grdMovimentoCaixa.FillStyle = flexFillSingle
End Sub
