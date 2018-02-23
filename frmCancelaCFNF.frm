VERSION 5.00
Begin VB.Form frmCancelaCFNF 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cancela CF/NF"
   ClientHeight    =   7455
   ClientLeft      =   5415
   ClientTop       =   240
   ClientWidth     =   12765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   12765
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   2430
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   4935
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
         Left            =   3240
         TabIndex        =   4
         Top             =   1110
         Width           =   1260
      End
      Begin VB.TextBox txtSenha 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2130
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1695
         Width           =   1035
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   885
         Width           =   960
      End
      Begin VB.Label lblValorTotal 
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
         TabIndex        =   8
         Top             =   885
         Width           =   1005
      End
      Begin VB.Label Label5 
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
         Left            =   1710
         TabIndex        =   7
         Top             =   300
         Width           =   1500
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
         Left            =   1485
         TabIndex        =   6
         Top             =   1800
         Width           =   555
      End
   End
   Begin VB.Label lblMensagensTEF 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Mensagens TEF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2595
      Left            =   5490
      TabIndex        =   12
      Top             =   2715
      Width           =   5055
   End
End
Attribute VB_Name = "frmCancelaCFNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String
Dim wUltimoCupom As Double
Dim WGrupoAtualzado As Double
Dim wNumeroPedido As Double
Dim wWhere As String

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
        
        wWhere = " "
  '      End If

        Sql = "SELECT CTR_DATAINICIAL, CTR_SITUACAOCAIXA FROM  CONTROLECAIXA " _
            & " WHERE CTR_Supervisor <> 99 and CTR_SITUACAOCAIXA = 'A'"
     
        ADOSituacao.CursorLocation = adUseClient
        ADOSituacao.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                                          
        If Not ADOSituacao.EOF Then
           wData = ADOSituacao("CTR_DATAINICIAL")
        Else
           MsgBox "Caixa Fechado", vbInformation, "Aviso"
           ADOSituacao.Close
           Exit Sub
        End If
       
          
        ADOSituacao.Close
     
        Sql = "SELECT TOP 1 TIPONOTA,NumeroPed, SERIE, NF, TOTALNOTA, DATAEMI, rtrim(CHAVENFE) as CHAVENFE " _
            & " FROM NFCAPA WHERE " _
            & " SERIE = '" & txtSerie.text & "' AND " _
            & " NF = " & txtNotaFiscal.text & " " & Where
         
        ADOCancela.CursorLocation = adUseClient
        ADOCancela.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
 

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
                Sql = "exec SP_Cancela_NotaFiscal " & txtNotaFiscal.text & ",'" & txtSerie.text & "'"
                rdoCNLoja.Execute (Sql)
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

Private Sub Form_Activate()
 txtNotaFiscal.SetFocus
End Sub

Private Sub Form_Load()
  
Call AjustaTela(frmCancelaCFNF)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    wPedido = 0
End Sub

Private Sub Label8_Click()

End Sub

Private Sub Text5_Change()
    
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



Private Sub txtSenha_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  
  EfetuaCancelarTEF "210", ""
  
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
    
Sql = "SELECT TOTALNOTA, NF, SERIE,TipoNota,numeroped FROM NFCAPA WHERE NF = " & txtNotaFiscal.text & " " _
    & "AND SERIE = '" & UCase(Trim(txtSerie.text)) & "' and TIPONOTA <> 'C' and Dataemi = '" & Format(Date, "yyyy/mm/dd") & "'"
    
 ADOCancela.CursorLocation = adUseClient
 ADOCancela.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

If Not ADOCancela.EOF Then
       txtValorNF.text = Format(ADOCancela("TOTALNOTA"), "0.00")
       txtPedido.text = ADOCancela("numeroped")
       wPedido = ADOCancela("numeroped")
       pedido = ADOCancela("numeroped")
       wTipoNota = ADOCancela("TipoNota")
    Else
        MsgBox "NF não encontrado ou já cancelado", vbInformation, "Aviso"
        txtSerie.text = ""
        txtNotaFiscal.SelStart = 0
        txtNotaFiscal.SelLength = Len(txtNotaFiscal.text)
        txtNotaFiscal.SetFocus
End If
ADOCancela.Close

End Sub

Sub LimpaCampos()

        txtSerie.text = ""
        txtValorNF.text = ""
        txtSenha.text = ""
        txtNotaFiscal.text = ""
        txtPedido.text = ""
        Exit Sub

End Sub

Public Function EfetuaCancelarTEF(codigoPagamento As String, valorCobrado As String) As Boolean

  Dim Retorno        As Long
  Dim Buffer         As String * 20000
  Dim ProximoComando As Long
  Dim TipoCampo      As Long
  Dim TamanhoMinimo  As Integer
  Dim TamanhoMaximo  As Integer
  Dim ContinuaNavegacao  As Long
  Dim Mensagem As String
  Dim VARIAVEL As String
  
  valorCobrado = Format(valorCobrado, "###,###,##0.00")

  Screen.MousePointer = vbHourglass
  pedido = "1233456"
  valorCobrado = "1,80"
  Retorno = IniciaFuncaoSiTefInterativo(codigoPagamento, valorCobrado & Chr(0), pedido & Chr(0), Format("2018/02/23", "YYYYMMDD") & Chr(0), Format("09:44:00", "HHMMSS") & Chr(0), Trim(GLB_USU_Nome) & Chr(0), Chr(0))
  Screen.MousePointer = vbDefault

  ProximoComando = 0
  TipoCampo = 0
  TamanhoMinimo = 0
  TamanhoMaximo = 0
  ContinuaNavegacao = 0
  Resultado = 0
  Buffer = String(20000, 0)

    lblMensagensTEF.Caption = ""

  Do

    Screen.MousePointer = vbHourglass
    
    Retorno = ContinuaFuncaoSiTefInterativo(ProximoComando, TipoCampo, TamanhoMinimo, TamanhoMaximo, Buffer, Len(Buffer), Resultado)
    Screen.MousePointer = vbDefault

    If (Retorno = 10000) Then

      If ProximoComando = "1" Or ProximoComando = "2" Or ProximoComando = "3" Then
        Mensagem = lblMensagensTEF.Caption
        lblMensagensTEF.Caption = Trim(Buffer)
        If lblMensagensTEF.Caption = "" Then lblMensagensTEF.Caption = Mensagem
        lblMensagensTEF.Caption = UCase(lblMensagensTEF.Caption)
        lblMensagensTEF.Refresh
      End If
     
      'lblParcelas.Caption = Buffer

      VARIAVEL = VARIAVEL & ProximoComando & " - " & Resultado & " - " & Buffer & vbNewLine

     Select Case ProximoComando
          Case 34
              If Buffer Like "Forneca o valor da transacao a ser canc*" Then
                  Buffer = valorCobrado
                  VARIAVEL = VARIAVEL + Buffer & vbNewLine
              End If

            Case 30
                If Buffer Like "Data da transacao*" Then
                    Buffer = "23022018"
                    VARIAVEL = VARIAVEL + Buffer & vbNewLine
                ElseIf Buffer Like "Forneca o numero do documento a ser*" Then
                    Buffer = "999230140"
                    VARIAVEL = VARIAVEL + Buffer & vbNewLine
                End If
              
            Case 21
                If Buffer Like "*1:Magnetico/Chip;2:Digitado;*" Then
                    Buffer = "1"
                    VARIAVEL = VARIAVEL + Buffer & vbNewLine
                End If
              
            End Select

    End If

  Loop Until Not (Retorno = 10000)

  If (Retorno = 0) Then
    lblMensagensTEF.Caption = "Retorno Ok!"
    EfetuaPagamentoTEF = True
    
  Else
    'Retorno = IniciaFuncaoSiTefInterativo(3, 10, 10, "20180216", "101010", "ACASD", "")
    'Retorno = ContinuaFuncaoSiTefInterativo(ProximoComando, TipoCampo, TamanhoMinimo, TamanhoMaximo, Buffer, Len(Buffer), Resultado)
    'FrmSiTef.TxtDisplay.Text = FrmSiTef.TxtDisplay.Text & Buffer

    'felipetef
    'lblMensagensTEF.Caption = "Erro:" & " " & retornoFuncoesTEF(CStr(Retorno))
  End If

     FinalizaTransacaoSiTefInterativo 1, pedido, Format("2018/02/23", "YYYYMMDD"), Format("09:44:00", "HHMMSS")
     criaLogTef (VARIAVEL)
 
End Function


