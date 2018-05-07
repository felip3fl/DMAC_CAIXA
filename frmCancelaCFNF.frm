VERSION 5.00
Object = "{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7u.ocx"
Begin VB.Form frmCancelaCFNF 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cancela CF/NF"
   ClientHeight    =   7455
   ClientLeft      =   1725
   ClientTop       =   2520
   ClientWidth     =   15165
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   15165
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameCancelamentoTEF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   5370
      Left            =   5295
      TabIndex        =   12
      Top             =   30
      Width           =   5625
      Begin VSFlex7UCtl.VSFlexGrid grdNumeroTEF 
         Height          =   2835
         Left            =   570
         TabIndex        =   15
         Top             =   885
         Width           =   4515
         _cx             =   7964
         _cy             =   5001
         _ConvInfo       =   1
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmCancelaCFNF.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   5
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
         Begin VB.Timer timerVerificaResposta 
            Enabled         =   0   'False
            Interval        =   3000
            Left            =   0
            Top             =   0
         End
      End
      Begin VB.Label lblModalidade 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Mensagens TEF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   60
         TabIndex        =   16
         Top             =   3945
         Width           =   5535
      End
      Begin VB.Label lblMensagensTEF 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Mensagens TEF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   885
         Left            =   135
         TabIndex        =   14
         Top             =   4395
         Width           =   5355
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelamento TEF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   0
         TabIndex        =   13
         Top             =   225
         Width           =   5625
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   2430
      Left            =   150
      TabIndex        =   0
      Top             =   30
      Width           =   4890
      Begin VB.TextBox txtNotaFiscal 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   405
         TabIndex        =   1
         Top             =   1110
         Width           =   1065
      End
      Begin VB.TextBox txtSerie 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1530
         TabIndex        =   2
         Top             =   1110
         Width           =   555
      End
      Begin VB.TextBox txtPedido 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2145
         TabIndex        =   3
         Top             =   1110
         Width           =   1035
      End
      Begin VB.TextBox txtValorNF 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3225
         TabIndex        =   4
         Top             =   1110
         Width           =   1260
      End
      Begin VB.TextBox txtSenha 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2250
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1695
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H0081E8FA&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total "
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
         Left            =   3255
         TabIndex        =   11
         Top             =   885
         Width           =   1005
      End
      Begin VB.Label lblnroNFCF 
         AutoSize        =   -1  'True
         BackColor       =   &H0081E8FA&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. NF"
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
         Left            =   420
         TabIndex        =   10
         Top             =   885
         Width           =   675
      End
      Begin VB.Label lblSerie 
         AutoSize        =   -1  'True
         BackColor       =   &H0081E8FA&
         BackStyle       =   0  'Transparent
         Caption         =   "Série"
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
         Left            =   1545
         TabIndex        =   9
         Top             =   885
         Width           =   450
      End
      Begin VB.Label lblPedido 
         AutoSize        =   -1  'True
         BackColor       =   &H0081E8FA&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.Pedido"
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
         Left            =   2160
         TabIndex        =   8
         Top             =   885
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelar Nota"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   0
         TabIndex        =   7
         Top             =   300
         Width           =   4890
      End
      Begin VB.Label lblSenha 
         AutoSize        =   -1  'True
         BackColor       =   &H0081E8FA&
         BackStyle       =   0  'Transparent
         Caption         =   "Senha"
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
         Left            =   1605
         TabIndex        =   6
         Top             =   1800
         Width           =   555
      End
   End
   Begin Balcao2010.chameleonButton cmdCancelarTEFnaoFinalizado 
      Height          =   555
      Left            =   6090
      TabIndex        =   17
      Top             =   5595
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   979
      BTYPE           =   14
      TX              =   "Cancelar TEF(s) não finalizado"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   2500134
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "frmCancelaCFNF.frx":0140
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmCancelaCFNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim wUltimoCupom As Double
Dim WGrupoAtualzado As Double
Dim wNumeroPedido As Double
Dim wWhere As String
Dim wDataEmissao As String

Dim nf As notaFiscalTEF

Private Sub cmbSair_Click()
 Unload Me
End Sub

Private Sub finalizarCancelamento()
        
        If Trim(txtNotaFiscal.text) = "" Then
            MsgBox "Favor digite o Numero NF ", vbInformation, "Aviso"
            txtNotaFiscal.SelStart = 0
            txtNotaFiscal.SelLength = Len(txtNotaFiscal.text)
            txtNotaFiscal.SetFocus
            Exit Sub
            
        ElseIf IsNumeric(txtNotaFiscal.text) = False Then
               MsgBox "Numero NF/CF Inválido", vbCritical, "Atenção"
               txtNotaFiscal.SelStart = 0
               txtNotaFiscal.SelLength = Len(txtNotaFiscal.text)
               txtNotaFiscal.SetFocus
               Exit Sub
                                        
        ElseIf Trim(txtSerie.text) = "" Then
               MsgBox "Favor informe a série", vbInformation, "Atenção"
               txtSerie.SelStart = 0
               txtSerie.SelLength = Len(txtSerie.text)
               txtSerie.SetFocus
               Exit Sub
        
        ElseIf txtSenha.text = "" Then
               MsgBox "Favor digite a senha", vbInformation, "Aviso"
               txtSenha.SelStart = 0
               txtSenha.SelLength = Len(txtSenha.text)
               txtSenha.SetFocus
               Exit Sub
        End If
      
      If Trim(UCase((txtSenha.text))) <> Trim(UCase(wSenhaLiberacao)) Then
         MsgBox "Senha para cancelamento não confere", vbCritical, "Atenção"
         txtSenha.SelStart = 0
         txtSenha.SelLength = Len(txtSenha.text)
         txtSenha.SetFocus
         Exit Sub
      End If
      
      If MsgBox("Deseja realmente Cancelar? NF --> " & txtNotaFiscal.text & ", Serie --> " & txtSerie.text & "," _
         & " Valor --> " & txtValorNF.text & "", vbQuestion + vbYesNo, "Atenção") = vbNo Then
         txtSenha.text = ""
         txtSerie.text = ""
         txtNotaFiscal.text = ""
         txtPedido.text = ""
         txtValorNF.text = ""
         txtNotaFiscal.SelStart = 0
         txtNotaFiscal.SelLength = Len(txtNotaFiscal.text)
         txtNotaFiscal.SetFocus
         Exit Sub
      End If
        
      If Not cancelarTEF Then
         MsgBox "Todos os TEF precisam está cancelado para proceguir com o cancelamento", vbCritical, "Atenção"
         txtSenha.SelStart = 0
         txtSenha.SelLength = Len(txtSenha.text)
         txtSenha.SetFocus
         Exit Sub
      End If
        
        wWhere = " "
  '      End If

        sql = "SELECT CTR_DATAINICIAL, CTR_SITUACAOCAIXA FROM  CONTROLECAIXA " _
            & " WHERE CTR_Supervisor <> 99 and CTR_SITUACAOCAIXA = 'A'"
     
        ADOSituacao.CursorLocation = adUseClient
        ADOSituacao.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                                          
        If Not ADOSituacao.EOF Then
           wData = ADOSituacao("CTR_DATAINICIAL")
        Else
           MsgBox "Caixa Fechado", vbInformation, "Aviso"
           ADOSituacao.Close
           Exit Sub
        End If
       
          
        ADOSituacao.Close
     
        sql = "SELECT TOP 1 TIPONOTA,NumeroPed, SERIE, NF, TOTALNOTA, DATAEMI, rtrim(CHAVENFE) as CHAVENFE " _
            & " FROM NFCAPA WHERE " _
            & " SERIE = '" & txtSerie.text & "' AND " _
            & " NF = " & txtNotaFiscal.text & " " & where
         
        ADOCancela.CursorLocation = adUseClient
        ADOCancela.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
 

        If Not ADOCancela.EOF Then
           If ADOCancela("DATAEMI") <> Date Then
              MsgBox "NF não pode ser cancelado. Somente é permitido cancelar NF do mesmo dia.", vbInformation, "Aviso"
              ADOCancela.Close
              Exit Sub
           End If
                 
            If (ADOCancela("CHAVENFE") = "" Or IsNull(ADOCancela("CHAVENFE")) = True) And ADOCancela("Serie") = "NE" Then
                cancelaNotaResultado = True
            ElseIf ADOCancela("Serie") = "00" Then
                cancelaNotaResultado = True
            Else
                cancelaNota = True
                frmEmissaoNFe.Show vbModal
                wPedido = 0
            End If
              
            If cancelaNotaResultado = True Then
                sql = "exec SP_Cancela_NotaFiscal " & txtNotaFiscal.text & ",'" & txtSerie.text & "'"
                rdoCNLoja.Execute (sql)
            Else
                MsgBox "Cancelamento não realizado", vbCritical, "DMAC Caixa"
            End If
                
         Else
             MsgBox "NF não encontrado", vbInformation, "Aviso"
             ADOCancela.Close
             Exit Sub
         End If
        
         ADOCancela.Close
       
    Call LimpaCampos
    
    txtNotaFiscal.SetFocus
       
End Sub

Private Sub cmdRetorna_Click()
Unload Me
End Sub

Private Sub cmdCancelarTEFnaoFinalizado_Click()
    cancelarTEF
End Sub

Private Sub Form_Activate()
  txtNotaFiscal.SetFocus
End Sub

Private Sub cancelarTEFdeOperacaoNaoConcluida()
    
    If Not GLB_TefHabilidado Then Exit Sub
    
    'Dim ADOCancelaTEF As New ADODB.Recordset
    
    carregaNotasComTEFGrid False
    
    'grdNumeroTEF.Rows = grdNumeroTEF.FixedRows
    
End Sub

Private Sub Form_Load()
  
    Call AjustaTela(frmCancelaCFNF)
    
    frameCancelamentoTEF.Visible = False
    grdNumeroTEF.Rows = grdNumeroTEF.FixedRows
    
    carregaNotasComTEF
    cancelarTEFdeOperacaoNaoConcluida
    
    cmdCancelarTEFnaoFinalizado.Visible = GLB_TEFnaoCancelado
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    wPedido = 0
End Sub

Private Sub Label8_Click()

End Sub

Private Sub Text5_Change()
    
End Sub

Private Sub Label5_Click()
4890
End Sub

Private Sub lblMensagensTEF_Click()
    cancelarOperacaoTEF = True
End Sub

Private Sub txtNotaFiscal_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
       Unload Me
    End If

    If KeyAscii = 13 Then
       txtSerie.SetFocus
    End If
    
End Sub

Private Sub txtNotaFiscal_LostFocus()
If txtNotaFiscal.text = "" Then
    Exit Sub
End If

If IsNumeric(txtNotaFiscal.text) = False Then
    
    txtNotaFiscal.text = ""
    txtNotaFiscal.SetFocus
    Exit Sub
End If
End Sub

Private Function cancelarTEF()

    cancelarTEF = True

    If Not GLB_TefHabilidado Then Exit Function
    
    Dim i As Byte
    Dim codigoOperacaoVenda As String
    Dim codigoOperacaoCancelamento As String
    
    i = grdNumeroTEF.FixedRows
    
    Do While i <= grdNumeroTEF.Rows - 1
        
        codigoOperacao = Mid(grdNumeroTEF.TextMatrix(i, 2), 1, 1)
        
        nf.pedido = grdNumeroTEF.TextMatrix(i, 6)
        nf.numeroTEF = grdNumeroTEF.TextMatrix(i, 0)
        nf.serie = grdNumeroTEF.TextMatrix(i, 5)
        nf.dataEmissao = grdNumeroTEF.TextMatrix(i, 4)
        nf.valor = grdNumeroTEF.TextMatrix(i, 1)
        nf.sequenciaMovimentoCaixa = grdNumeroTEF.TextMatrix(i, 8)
        
        lblModalidade.Caption = "Insira cartão da bandeira " & grdNumeroTEF.TextMatrix(i, 3) & " do valor de " & nf.valor
        lblMensagensTEF.Caption = ""
        
        If codigoOperacao = 2 Then codigoOperacaoCancelamento = 211
        If codigoOperacao = 3 Then codigoOperacaoCancelamento = 210
        
        If EfetuaOperacaoTEF(codigoOperacaoCancelamento, nf, lblModalidade, lblMensagensTEF) Then
            
            ImprimeComprovanteTEF nf.pedido
            cancelaMovimentoCaixaEspecifico nf.sequenciaMovimentoCaixa
            finalizarTransacaoTEF nf.pedido, nf.serie, False
            
        End If
            
        'carregaNotasComTEFGrid
        
        i = i + 1
        
    Loop
    
    carregaNotasComTEFGrid Not (GLB_TEFnaoCancelado)
    
    
    
    If grdNumeroTEF.Rows > grdNumeroTEF.FixedRows Then
        MsgBox "Há TEF(s) não cencelado(s) que impedem o cancelamento da Nota Fiscal", vbExclamation, "TEF"
        cancelarTEF = False
    Else
        GLB_TEFnaoCancelado = False
        cmdCancelarTEFnaoFinalizado.Visible = GLB_TEFnaoCancelado
    End If
    
End Function

Private Sub cancelaMovimentoCaixaEspecifico(sequenciaMovimentoCaixa As String)

    Dim sql As String
    
    sql = "update movimentocaixa " & vbNewLine & _
          "set mc_tiponota = 'C'" & vbNewLine & _
          "where " & vbNewLine & _
          "mc_sequencia = '" & sequenciaMovimentoCaixa & "'"

    rdoCNLoja.Execute sql
    
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  
  
  
  If txtSenha.text <> "" Then
     Call finalizarCancelamento
   End If
   
End If

If KeyAscii = 27 Then
   txtSerie.SelStart = 0
   txtSerie.SelLength = Len(txtNotaFiscal.text)
   txtSerie.SetFocus

End If

End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtSenha.SetFocus
End If

If KeyAscii = 27 Then
   txtNotaFiscal.SelStart = 0
   txtNotaFiscal.SelLength = Len(txtNotaFiscal.text)
   txtNotaFiscal.SetFocus

End If
End Sub

Private Sub txtSerie_LostFocus()

    If txtSerie.text = "" Then
        Exit Sub
    End If
    
    If txtNotaFiscal.text = "" Then
        MsgBox "Preencha todos os campos", vbCritical, "Atenção"
        txtNotaFiscal.SelStart = 0
        txtNotaFiscal.SelLength = Len(txtNotaFiscal.text)
        txtNotaFiscal.SetFocus
        Exit Sub
    End If
    
    If txtNotaFiscal.text = "" Then
        MsgBox "Preencha todos os campos", vbCritical, "Atenção"
        txtNotaFiscal.SelStart = 0
        txtNotaFiscal.SelLength = Len(txtNotaFiscal.text)
        txtNotaFiscal.SetFocus
        Exit Sub
    End If
    
    If Not UCase(txtSerie.text) Like "CE*" Then
        If UCase(txtSerie.text) Like GLB_SerieCF & "*" Then
            MsgBox "Para cancelamento de Cupom Fiscal selecione Operações ECF", vbCritical, "Atenção"
            txtSerie.text = ""
            txtNotaFiscal.SelStart = 0
            txtNotaFiscal.SelLength = Len(txtNotaFiscal.text)
            txtNotaFiscal.SetFocus
            Exit Sub
        End If
    End If
    
    
    txtSerie.text = UCase(txtSerie.text)
    
    
    sql = "SELECT TOTALNOTA, NF, SERIE,TipoNota,numeroped, dataemi FROM NFCAPA WHERE NF = " & txtNotaFiscal.text & " " _
    & "AND SERIE = '" & UCase(Trim(txtSerie.text)) & "' and TIPONOTA <> 'C' --and Dataemi = '" & Format(Date, "yyyy/mm/dd") & "'"
    
    ADOCancela.CursorLocation = adUseClient
    ADOCancela.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not ADOCancela.EOF Then
        txtValorNF.text = Format(ADOCancela("TOTALNOTA"), "0.00")
        txtPedido.text = ADOCancela("numeroped")
        wPedido = ADOCancela("numeroped")
        pedido = ADOCancela("numeroped")
        wTipoNota = ADOCancela("TipoNota")
        wDataEmissao = ADOCancela("dataemi")
        
        carregaNotasComTEF
        carregaNotasComTEFGrid True
        
    Else
        MsgBox "NF não encontrado ou já cancelado", vbInformation, "Aviso"
        txtSerie.text = ""
        txtNotaFiscal.SelStart = 0
        txtNotaFiscal.SelLength = Len(txtNotaFiscal.text)
        txtNotaFiscal.SetFocus
    End If
    
    ADOCancela.Close

End Sub

Private Sub carregaNotasComTEF()

    If Not GLB_TefHabilidado Then Exit Sub

    frameCancelamentoTEF.Visible = True
    lblMensagensTEF.Caption = "As mensagens do TEF serão exibida aqui"
    lblModalidade.Caption = ""
    
    

End Sub

Private Sub carregaNotasComTEFGrid(carregaNotasFinalizadas As Boolean)
    Dim sql As String
    Dim where As String
    Dim orderby As String
    Dim RsDados As New ADODB.Recordset
    Dim descricaoModalidade As String
    Dim tipoModalidade As String
    
    grdNumeroTEF.Rows = grdNumeroTEF.FixedRows
    
    sql = "select MC_SequenciaTEF as TEF, MC_VALOR as valor, " & vbNewLine & _
          "MO_Descricao as descricaoModalidade, MC_Grupo as grupo, " & vbNewLine & _
          "Mc_PEDIDO as PEDIDO, MC_DOCUMENTO as Documento, " & vbNewLine & _
          "MC_Sequencia as Sequencia,  " & vbNewLine & _
          "Mc_DATA as data, MC_SERIE as SERIE " & vbNewLine & _
          "from MovimentoCaixa " & vbNewLine & _
          "FULL OUTER JOIN modalidade " & vbNewLine & _
          "on MC_Grupo = MO_Grupo " & vbNewLine
          
    
    If carregaNotasFinalizadas Then
    
        where = "where mc_pedido = '" & txtPedido.text & "' " & vbNewLine & _
              "and mc_tipoNOTA IN ('V') " & vbNewLine & _
              "and MC_Sequenciatef > 0 " & vbNewLine & _
              "and MC_Grupo < '20000'" & vbNewLine

    Else
    
        where = "where mc_tipoNOTA IN ('PA') " & vbNewLine & _
              "and MC_Sequenciatef > 0 "
    
    End If

    orderby = "order by mc_sequenciaTEF" & vbNewLine

    sql = sql & where & orderby

    RsDados.CursorLocation = adUseClient
    RsDados.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
        Do While Not RsDados.EOF
        
            If IsNull(RsDados("descricaoModalidade")) Then
                descricaoModalidade = "DESCONHECIDO"
            Else
                descricaoModalidade = RsDados("descricaoModalidade")
            End If
        
            tipoModalidade = "3 Crédito"
            Select Case RsDados("grupo")
            Case "10203", "10206"
                tipoModalidade = "2 Débito"
            End Select
        
            grdNumeroTEF.AddItem Format(RsDados("TEF"), "000000") & vbTab & _
                                 Format(RsDados("valor"), "###,###,##0.00") & vbTab & _
                                 tipoModalidade & vbTab & _
                                 descricaoModalidade & vbTab & _
                                 RsDados("data") & vbTab & _
                                 RsDados("serie") & vbTab & _
                                 RsDados("pedido") & vbTab & _
                                 RsDados("documento") & vbTab & _
                                 RsDados("Sequencia")
            RsDados.MoveNext
            
        Loop
    
    RsDados.Close
    
End Sub

Sub LimpaCampos()

        txtSerie.text = ""
        txtValorNF.text = ""
        txtSenha.text = ""
        txtNotaFiscal.text = ""
        txtPedido.text = ""
        Exit Sub

End Sub

