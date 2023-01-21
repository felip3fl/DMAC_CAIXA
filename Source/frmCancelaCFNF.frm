VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmCancelaCFNF 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cancela CF/NF"
   ClientHeight    =   2580
   ClientLeft      =   4245
   ClientTop       =   4185
   ClientWidth     =   5085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   2430
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   4935
      Begin VB.CheckBox ChkTef 
         BackColor       =   &H80000012&
         Caption         =   "Somente Cancelar o TEF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000BB&
         Height          =   315
         Left            =   480
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   13
         Top             =   2040
         Visible         =   0   'False
         Width           =   2535
      End
      Begin MSWinsockLib.Winsock wskTef 
         Left            =   4440
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
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
      Begin VB.Label lblDiplay 
         AutoSize        =   -1  'True
         BackColor       =   &H00AE7411&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000BB&
         Height          =   300
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Visible         =   0   'False
         Width           =   4650
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

'-------------------Emerson Tef--------------
Dim tef_cupom_1 As String
Dim tef_cupom_2 As String
Dim tef_modelidade As String
Dim tef_mensssagem As String
Dim tef_sequencia As String
Dim tef_Parcelas As String


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
      If ChkTef.Value = 1 Then
      ChkTef.Visible = False
      tef_sql = "select * from  nfcapa where tiponota='E' and  NfDevolucao=" & Trim(txtNotaFiscal.text) & "" _
       & "and SerieDevolucao='" & txtSerie.text & "' and TOTALNOTA= " & ConverteVirgula(txtValorNF.text) & " "
      'MsgBox tef_sql
        ADOTef_C1.CursorLocation = adUseClient
        ADOTef_C1.Open tef_sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
      If Not ADOTef_C1.EOF Then
             If verifica_tef Then
                  tef_dados = ""
                  Call Cancela_Tef(0)
                      If wskTef.State <> 0 Then
                      ADOTef_C1.Close
                      Exit Sub
                      End If
            ADOTef_C1.Close
            Exit Sub
            End If
        Else
        
        MsgBox "TEF não pode ser Cancelado!!", vbCritical, "ERRO"
        Call LimpaCampos
        ADOTef_C1.Close
        Exit Sub
      End If
   
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
            & " NF = " & txtNotaFiscal.text & " " & Where
         
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
                'Emerson
            If verifica_tef Then
            ChkTef.top = 4040
            tef_dados = ""
            Call Cancela_Tef(0)
                If wskTef.State <> 0 Then
                ADOCancela.Clone
                    Exit Sub
                End If
            End If
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



Private Sub Form_Activate()
 txtNotaFiscal.SetFocus
  'Emerson
 Verifica_Tef_Pos
 If verifica_tef Then
 ChkTef.Visible = True
 End If
 
End Sub

Private Sub Form_Load()
  
Call AjustaTela(frmCancelaCFNF)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    wPedido = 0
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
    
sql = "SELECT TOTALNOTA, NF, SERIE,TipoNota,numeroped FROM NFCAPA WHERE NF = " & txtNotaFiscal.text & " " _
    & "AND SERIE = '" & UCase(Trim(txtSerie.text)) & "' and TIPONOTA <> 'C' and Dataemi = '" & Format(Date, "yyyy/mm/dd") & "'"
    
 ADOCancela.CursorLocation = adUseClient
 ADOCancela.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

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
        If verifica_tef Then
            lblDiplay.Visible = False
            ChkTef.top = 2040
            ChkTef.Value = 0
        End If
        Exit Sub

End Sub
'---------Emerson_Tef_VBI



Private Sub Cancela_Tef(ByVal sequecial As Integer)
sql = "Select * from  MovimentoCaixa where mc_data='" & Format(Date, "yyyy/mm/dd") & "' and mc_pedido=" & txtPedido.text & " and  mc_documento=" & txtNotaFiscal.text & "" _
& " and MC_TipoNota='V' and MC_SequenciaTEF > " & Trim(sequecial) & " and MC_Grupo in (10203,10205,10206,10301,10302,10303) order by MC_SequenciaTEF "

 ADOTef_C.CursorLocation = adUseClient
 ADOTef_C.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
 If Not ADOTef_C.EOF Then
 
   qtdCartao = 1
    lblDiplay.Visible = True
    tef_operacao = "Administracao Cancelar"
            tef_num_doc = Format(ADOTef_C("Mc_SequenciaTef"), "000000")
            tef_nsu_ctf = Format(ADOTef_C("Mc_SequenciaTef"), "000000")
            tef_data_cli = Format(Date, "dd/mm/yy")
            data_tef = Date
            tef_num_trans = Format(qtdCartao, "00")
            tef_valor = Format(ADOTef_C("mc_valor"), "##,##0.00")
            tef_Parcelas = Trim(ADOTef_C("mc_parcelas"))
            If Trim(ADOTef_C("MC_Grupo")) = "10203" Or Trim(ADOTef_C("MC_Grupo")) = "10206" Then
            tef_operacao = "Debito"
            Else
            tef_operacao = "Credito"
            End If
            
            Tef_Confrima = False
            
             If tef_dados = "" Then
             IniciaTEF
             End If
End If
   ADOTef_C.Close

End Sub

'Emerson_Tef_Vbi
Private Sub wskTef_Close()
wskTef.Close
 tef_dados = ""
End Sub

Private Sub wskTef_Connect()
wskTef.SendData tef_dados
End Sub


Private Sub wskTef_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "Erro NO Tef - " & Number & " - " & Description, vbCritical, "ERRO"
Conclui_Tef
wskTef.Close
End Sub
Private Function getMenssagem(ByVal testoInteiro As String, ByVal textoBusca As String, ByVal Maximo As Integer) As String
Dim Texto As String

If InStr(testoInteiro, textoBusca) >= 1 Then
    Texto = Mid$(testoInteiro, InStr(testoInteiro, textoBusca) + Maximo)
    Texto = Mid$(testoInteiro, InStr(testoInteiro, textoBusca) + Maximo, InStr(Texto, """") - 1)
    getMenssagem = Texto
Else
    getMenssagem = ""
End If
End Function

Public Function IniciaTEF()
 tef_sequencia = sequencial_Tef_Vbi
 ususrio_senha_Tef_Vbi
    wskTef.Connect "localhost", 60906
    tef_dados = "versao=""v" & App.Major & "." & App.Minor & "." & App.Revision & """" + vbCrLf
    tef_dados = tef_dados + "sequencial=""" & tef_sequencia + 1 & """" + vbCrLf
    tef_dados = tef_dados + "retorno=""1""" + vbCrLf
    tef_dados = tef_dados + "servico=""iniciar""" + vbCrLf
    tef_dados = tef_dados + "aplicacao="" De Meo """ + vbCrLf
    tef_dados = tef_dados + "aplicacao_tela=""Dmac Caixa"""
    tef_servico = "iniciar"
End Function
Private Sub wskTef_DataArrival(ByVal bytesTotal As Long)
Dim resp As String
Dim resp1 As String
tef_menssagem = ""
wskTef.GetData resp, vbString
resp1 = resp
tef_retorno = getMenssagem(resp, "retorno=", 9)
 Call Grava_Log_Diario(resp1)
If tef_servico = "iniciar" Then
    tef_menssagem = getMenssagem(resp, "estado", 8)
    If tef_menssagem = "7" And tef_retorno = "1" Then
        executarTEF
    ElseIf tef_retorno > 1 Then
        MsgBox "Erro NO Tef - " & getMenssagem(resp, "mensagem", 10), vbCritical, "ERRO"
        tef_servico = ""
        Conclui_Tef
    End If
ElseIf tef_servico = "executar" Then
    tef_retorno = getMenssagem(resp, "retorno=", 9)
    If tef_retorno <= 1 Then
            If InStr(resp, "_sequencial=") >= 1 Then
                    tef_menssagem = getMenssagem(resp, "mensagem", 10)
                    lblDiplay.Caption = tef_menssagem
                    Call Continua(getMenssagem(resp, "_sequencial=", 13))
          ElseIf InStr(resp, "_nsu=") >= 1 Then
                    Call Grava_Campos_Tef(resp)
                    Tef_Confrima = True
                    Call valida
                    lblDiplay.Caption = "Retire o Cartão"
            End If
    ElseIf tef_retorno > 1 Then
        lblDiplay.Caption = getMenssagem(resp, "mensagem", 10)
        tef_servico = ""
          MsgBox "Erro NO Tef - " & getMenssagem(resp, "mensagem", 10), vbCritical, "ERRO"
        lblDiplay.Caption = "Retire o Cartão"
         Call Finalizar_Tef
    End If

ElseIf tef_servico = "confirma" Then
         If InStr(resp, "sequencial=") >= 1 Then
          Call Finalizar_Tef
          Tef_Confrima = True
          ElseIf tef_retorno > 1 Then
          
                MsgBox "Erro NO Tef - " & getMenssagem(resp, "mensagem", 10), vbCritical, "ERRO"
                tef_servico = ""
                Finalizar_Tef
            End If
ElseIf tef_servico = "finalizar" Then
        Call Conclui_Tef
         lblDiplay.Caption = ""
ElseIf tef_retorno > 1 Then
        MsgBox "Erro NO Tef - " & getMenssagem(resp, "mensagem", 10), vbCritical, "ERRO"
        tef_servico = ""
        Conclui_Tef
End If

End Sub
Public Function executarTEF()
    tef_servico = "executar" '
    tef_dados = "sequencial=""" & tef_sequencia + 2 & """" + vbCrLf
    tef_dados = tef_dados + "servico=""executar""" + vbCrLf
    tef_dados = tef_dados + "retorno=""1""" + vbCrLf
    tef_dados = tef_dados + "transacao=""Administracao Cancelar""" + vbCrLf
    tef_dados = tef_dados + "transacao_tipo_cartao=""" & tef_operacao & """"
    wskTef.SendData tef_dados

End Function


Private Sub Continua(ByVal sequecial As String)
'ok
Dim retornoLocal As String
Dim sequencialLocal As String
Dim informacao As String
tef_servico = "executar"
        retornoLocal = "0"
        sequencialLocal = sequecial
        
        
         
        
        If tef_menssagem = "Valor" Or tef_menssagem = "Valor da Transacao" Then
        
            informacao = Replace(Format(tef_valor, "#####.00"), ",", ".")
        ElseIf tef_menssagem = "Produto" Then
            informacao = tef_operacao & "-Stone"
        ElseIf tef_menssagem = "Forma de Pagamento" And tef_operacao = "Debito" Then
            informacao = "A vista"
            tef_Parcelas = 0
             MsgBox "A vista"
        ElseIf tef_menssagem = "Forma de Pagamento" And tef_Parcelas <= 1 Then
            informacao = "A vista"
            MsgBox "A vista"
        ElseIf tef_menssagem = "Forma de Pagamento" And tef_Parcelas >= 2 Then
            informacao = "Parcelado"
            MsgBox "Parcelado"
        ElseIf tef_menssagem = "Financiado pelo" Then
            informacao = "Estabelecimento"
        ElseIf tef_menssagem = "Parcelas" Then
           informacao = tef_Parcelas
        ElseIf tef_menssagem = "Taxa de Embarque" Then
           informacao = 0
        ElseIf tef_menssagem = "Usuario de acesso" Then
           informacao = tef_usuario
        ElseIf tef_menssagem = "Senha de acesso" Then
           informacao = tef_senha
        ElseIf tef_menssagem = "Reimprimir" Then
           informacao = "Todos"
        ElseIf tef_menssagem = "Data Transacao Original" Then
           informacao = Format(Date, "dd/mm/yy")
        ElseIf tef_menssagem = "Numero do Documento" Then
           informacao = tef_nsu_ctf
        ElseIf tef_menssagem = "Quatro ultimos digito" Then
           informacao = InputBox(Trim(tef_menssagem) & ":")
        ElseIf tef_menssagem = "Codigo de Seguranca" Then
           informacao = InputBox(Trim(tef_menssagem) & ":")
        ElseIf tef_menssagem = "Validade do Cartao(MM/AA)" Then
           informacao = InputBox(Trim(tef_menssagem) & ":")
        ElseIf InStr(tef_menssagem, "?") >= 1 Then
           informacao = "Sim"
        
        Else
            informacao = ""
        End If
        
        tef_dados = "automacao_coleta_retorno=""" + retornoLocal + """" + vbCrLf
        tef_dados = tef_dados + "automacao_coleta_sequencial=""" + sequencialLocal + """" + vbCrLf

    If informacao <> "" Then
            tef_dados = tef_dados + "automacao_coleta_informacao=""" + informacao + """" + vbCrLf
            wskTef.SendData tef_dados
        
    Else
            wskTef.SendData tef_dados
    End If
End Sub



Private Sub valida()
tef_servico = "confirma"
    tef_dados = "sequencial=""" & tef_sequencia + 2 & """" + vbCrLf
    tef_dados = tef_dados + "servico=""executar""" + vbCrLf
    tef_dados = tef_dados + "retorno=""0""" + vbCrLf
    tef_dados = tef_dados + "transacao=""Administracao Cancelar"""
    wskTef.SendData tef_dados
End Sub
Private Sub Conclui_Tef()
    wskTef.Close
   Screen.MousePointer = 0
   Fecha_Log_Diario
    If Tef_Confrima = True Then
        tef_dados = ""
        If ChkTef.Value = 0 Then
        
        sql = "update movimentocaixa set mc_tiponota='C',mc_sequenciatef1 = " & Trim(tef_nsu_ctf) & " where mc_sequenciatef=" & Trim(tef_num_doc) & " and mc_serie='" & Trim(txtSerie.text) & "' and mc_documento=" & Trim(txtNotaFiscal.text)
        rdoCNLoja.Execute (sql)
        
        Call Cancela_Tef(tef_num_doc)
        If wskTef.State <> 0 Then
           Exit Sub
        End If
        sql = ""
        sql = "exec SP_Cancela_NotaFiscal " & txtNotaFiscal.text & ",'" & txtSerie.text & "'"
        rdoCNLoja.Execute (sql)
        
        ElseIf ChkTef.Value = 1 Then
             sql = "update movimentocaixa set mc_sequenciatef1 = " & Trim(tef_nsu_ctf) & " where mc_sequenciatef=" & Trim(tef_num_doc) & " and mc_serie='" & Trim(txtSerie.text) & "' and mc_documento=" & Trim(txtNotaFiscal.text)
                rdoCNLoja.Execute (sql)
                 Call Cancela_Tef(tef_nsu_ctf)
                If wskTef.State <> 0 Then
                   Exit Sub
                End If
            
        End If
        LimpaCampos
        Imprimir_Tef
        
 ADOCancela.Close
    
 End If

   
End Sub

Private Sub Grava_Campos_Tef(ByVal resp As String)
    'ok
    tef_nsu_ctf = getMenssagem(resp, "_nsu=", 6)
    tef_bandeira = getMenssagem(resp, "_administradora=", 17)
    tef_operacao = getMenssagem(resp, "_cartao=", 9)
    tef_nome_ac = getMenssagem(resp, "o_rede=", 8)
    tef_cupom_1 = getComprovantes(resp, "transacao_", "comprovante_1via")
    Call Grava_Cupom(tef_cupom_1)
    tef_cupom_2 = getComprovantes(resp, "transacao_", "comprovante_2via")
    Call Grava_Cupom(tef_cupom_2)
End Sub


Private Function getComprovantes(ByVal resp As String, ByVal blc As String, ByVal copum As String) As String
'ok
resp = Mid$(resp, InStr(resp, copum) + 17)
getComprovantes = Mid$(resp, InStr(resp, vbCrLf), InStr(resp, blc) - 42)
getComprovantes = Replace(getComprovantes, vbCrLf, ";")

End Function


Private Sub Finalizar_Tef()
tef_servico = "finalizar"
tef_dados = "sequencial=""" & tef_sequencia + 3 & """" + vbCrLf
tef_dados = tef_dados + "retorno=""0""" + vbCrLf
tef_dados = tef_dados + "servico=""finalizar"""
wskTef.SendData tef_dados
End Sub



