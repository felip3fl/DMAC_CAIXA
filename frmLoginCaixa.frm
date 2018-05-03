VERSION 5.00
Begin VB.Form frmLoginCaixa 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Login do Caixa"
   ClientHeight    =   8760
   ClientLeft      =   1755
   ClientTop       =   1590
   ClientWidth     =   13395
   DrawStyle       =   2  'Dot
   Icon            =   "frmLoginCaixa.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLoginCaixa.frx":23FA
   ScaleHeight     =   8760
   ScaleWidth      =   13395
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtSenha 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   9765
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   5730
      Width           =   735
   End
   Begin VB.TextBox txtOperador 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7635
      TabIndex        =   0
      Top             =   5730
      Width           =   1335
   End
   Begin VB.Label lblMensagem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "mensagem"
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
      Height          =   570
      Left            =   5295
      TabIndex        =   4
      Top             =   9600
      Width           =   4740
   End
   Begin VB.Label lblOperador 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Operador"
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
      Left            =   6390
      TabIndex        =   3
      Top             =   5790
      Width           =   1020
   End
   Begin VB.Label lblSenhaOperador 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
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
      Left            =   9030
      TabIndex        =   2
      Top             =   5790
      Width           =   675
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   990
      Left            =   6300
      Top             =   5430
      Width           =   4320
   End
End
Attribute VB_Name = "frmLoginCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ContadorItens As Long
Dim CodigoOperador As Integer
'Dim CodigoSupervisor As Integer
Dim GuardaSequencia As Double
Dim wGrupo As String
Dim sql As String

Private Sub chSair_Click()
Call AlterarResolucao(resolucaoOriginal.Colunas, resolucaoOriginal.Linhas)
 Unload Me
 Unload frmFundoEscuro
 End
 
End Sub

Private Sub modoAdministrador()

ConectaODBCMatriz
GLB_Administrador = False
    
 sql = "Select us_nome as nome, us_senha as senha " & vbNewLine _
       & "from usuario where " & vbNewLine _
       & "US_Nome ='" & txtOperador.text & "' " & vbNewLine _
       & "and US_Permissao='A'"
       
 RsDados.CursorLocation = adUseClient
 RsDados.Open sql, rdoCNRetaguarda, adOpenForwardOnly, adLockPessimistic
 If Not RsDados.EOF Then
      nroProtocoloADM
      If RTrim(RsDados("senha")) <> txtSenha.text Then
         MsgBox "Senha do ADMINISTRADOR não Cadastrado", vbCritical, "Aviso"
         notificacaoEmail "Falha na tentativa de login como " & RTrim(RsDados("nome")) & " (Senha Incorreta)"
      Else
         MsgBox "Modo ADMINISTRADOR ativado!", vbInformation, "Aviso"
         GLB_Administrador = True
         GLB_ADMNome = RTrim(RsDados("nome"))
         frmControlaCaixa.cmdOperador.ButtonType = 1
         frmControlaCaixa.cmdOperador.Caption = "ADMINISTRADOR: " & GLB_ADMNome
         notificacaoEmail "Login feito com sucesso como " & RTrim(RsDados("nome")) & ""
         Unload Me
      End If
 Else
    MsgBox "Usuario não cadastrado ", vbCritical, "Aviso"
 End If
 RsDados.Close
 rdoCNRetaguarda.Close
 
End Sub

Private Sub nroProtocoloADM()

    Dim RsDados As New ADODB.Recordset

    sql = "select MAX(ame_numero) as protocolo from alerta_movimento_email"
    RsDados.CursorLocation = adUseClient
    RsDados.Open sql, rdoCNRetaguarda, adOpenForwardOnly, adLockPessimistic
    
    If Not RsDados.EOF And IsNull(RsDados("protocolo")) = False Then
        GLB_ADMProtocolo = Val(RsDados("protocolo")) + 1
    Else
        GLB_ADMProtocolo = 1
    End If
    
End Sub

Private Sub EfetuarLogin()
 '''''''''' Fechamento Geral

'**********
       wPermitirVenda = True
       sql = ("Select * from ControleCaixa Where CTR_Supervisor = 99 and CTR_SituacaoCaixa='A' and " _
                    & "CTR_DataInicial < '" & Format(Date, "yyyy/mm/dd") & "'")

       rsTEF.CursorLocation = adUseClient
       rsTEF.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            
       If rsTEF.EOF = False Then
           'MsgBox "Data do caixa geral incorreta.Favor efetuar o Fechamento Geral", vbCritical, "Atenção"
            wPermitirVenda = False
       End If
       rsTEF.Close
'********

 sql = ("Select * from UsuarioCaixa where USU_Nome ='" & txtOperador.text & "' and USU_codigo='99'")
 RsDados.CursorLocation = adUseClient
 RsDados.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
 If Not RsDados.EOF Then

       sql = ("Select * from ControleCaixa where CTR_Supervisor = 99 and " _
                    & "CTR_DataInicial >= '" & Format(Date, "yyyy/mm/dd") & "'")
       
       rsTEF.CursorLocation = adUseClient
       rsTEF.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

       If rsTEF.EOF And wPermitirVenda = True Then
           MsgBox "Por favor, abrir o caixa."
           rsTEF.Close
           RsDados.Close
           Exit Sub
           
       End If
       rsTEF.Close


      sql = "Select * from ControleCaixa Where CTR_supervisor <> '99' and CTR_situacaocaixa = 'A'"
      rsTEF.CursorLocation = adUseClient
      rsTEF.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

       If Not rsTEF.EOF Then
           MsgBox "Por favor, fazer o fechamento dos caixas para fazer o fechamento geral"
           rsTEF.Close
           RsDados.Close
           Exit Sub
       End If
       rsTEF.Close

 End If
RsDados.Close

   If wPermitirVenda = False Then
           MsgBox "Data do caixa geral incorreta.Favor efetuar o Fechamento Geral", vbCritical, "Atenção"
           Exit Sub
   End If

   sql = ("Select * from UsuarioCaixa where USU_Nome ='" & txtOperador.text & "' and USU_TipoUsuario='O'")
   RsDados.CursorLocation = adUseClient
   RsDados.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
   If RsDados.EOF Then
      MsgBox "Operador não Cadastrado", vbCritical, "Aviso"
      RsDados.Close
      Exit Sub
   Else


      If RTrim(RsDados("USU_Senha")) <> txtSenha.text Then
         MsgBox "Senha do Operador não Cadastrado", vbCritical, "Aviso"
         RsDados.Close
         Exit Sub
      Else
         GLB_USU_Nome = Trim(RsDados("USU_Nome"))
         GLB_USU_Codigo = Trim(RsDados("USU_Codigo"))
         CodigoOperador = RsDados("USU_Codigo")
         RsDados.Close
         'SQL = ("Select * from UsuarioCaixa where USU_Nome ='" & txtSupervisor.Text & "' and USU_TipoUsuario='S'")
         'RsDados.CursorLocation = adUseClient
         'RsDados.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
         'If RsDados.EOF Then
            'MsgBox "Supervisor não Cadastrado", vbCritical, "Aviso"
            'RsDados.Close
            'Exit Sub
         'Else
            'If RTrim(RsDados("USU_Senha")) <> txtSenhaSupervisor.Text Then
               'MsgBox "Senha do Supervisor não Cadastrado", vbCritical, "Aviso"
               'RsDados.Close
               'Exit Sub
            'Else
               
                'CodigoSupervisor = RsDados("USU_Codigo")
                'RsDados.Close
                
                sql = ("Select * from ControleCaixa where CTR_Supervisor = 99 and " _
                    & "CTR_DataInicial >= '" & Format(Date, "yyyy/mm/dd") & "'")
                RsDados.CursorLocation = adUseClient
                RsDados.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                If Not RsDados.EOF Then

                    If Trim(RsDados("ctr_situacaocaixa")) = "F" Then
                        MsgBox "Fechamento Geral já foi efetuado não será possivel abrir o caixa."
                        RsDados.Close
                        Exit Sub
                    End If
                Else
                        Call VerificaSaldoCaixa
                        sql = ""
                        sql = "Insert Into ControleCaixa (" _
                            & "CTR_Operador,CTR_Supervisor,CTR_DataInicial," _
                            & "CTR_DataFinal,CTR_SaldoAnterior," _
                            & "CTR_SaldoFinal,CTR_SituacaoCaixa,CTR_NumeroCaixa,CTR_ProtocoloAnterior) " _
                            & "Values (99,99,GetDate(),' ',0,0,'A'," & GLB_Caixa & "," & GuardaSequencia & ")"
                        rdoCNLoja.Execute (sql)

                End If
                RsDados.Close
                
                Call VerificaSaldoCaixa
                rdoCNLoja.BeginTrans
                Screen.MousePointer = vbHourglass
                sql = ""
                sql = "Insert Into ControleCaixa (" _
                    & "CTR_Operador,CTR_Supervisor,CTR_DataInicial," _
                    & "CTR_DataFinal,CTR_SaldoAnterior," _
                    & "CTR_SaldoFinal,CTR_SituacaoCaixa,CTR_NumeroCaixa,CTR_ProtocoloAnterior) " _
                    & "Values (" & CodigoOperador & "," & "0" _
                    & ",getdate(),' '," & ConverteVirgula(Format(saldoAnterior, "00.00")) & ",0,'A'," & GLB_Caixa & "," & GuardaSequencia & ")"
                    rdoCNLoja.Execute sql
                    'Screen.MousePointer = vbNormal
                    rdoCNLoja.CommitTrans
                Call GuardaProtocolo
                Call ComposicaoSaldoAnterior
                
                wPermitirVenda = True
                Call limparArquivosImpressaoTEF
                
''               If VerificaSeEmiteCupom = "S" Then
''                 Retorno = Bematech_FI_LeituraX()
''                 Call VerificaRetornoImpressora("", "", "Leitura X")
''               End If

                Screen.MousePointer = vbNormal
                
                Call AlterarResolucao(resolucaoOriginal.Colunas, resolucaoOriginal.Linhas)

                Unload Me
                Unload frmFundoEscuro
                Call carregaControleCaixa
                frmBandeja.Show vbModal
                
               ' frmControlaCaixa.ZOrder
            End If
         'End If
      'End If
   End If

End Sub

'Private Sub cmdGravar_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    cmdGravar_Click
'End If
'End Sub


Private Sub cmdRetornar_KeyPress(KeyAscii As Integer)
Call AlterarResolucao(resolucaoOriginal.Colunas, resolucaoOriginal.Linhas)
Unload Me
End
End Sub



Private Sub Form_Activate()
    resolucaoOriginal.Colunas = resolucaoTela.Colunas
    resolucaoOriginal.Linhas = resolucaoTela.Linhas
    Call AlterarResolucao(1024, 768)
    
    If GLB_USU_Codigo = Empty Then
        lblMensagem.Caption = ""
    Else
        lblMensagem.Caption = "Login como Modo Administrador"
    End If
    
End Sub

Private Sub Form_Load()
  left = (Screen.Width - Width) / 2
  top = (Screen.Height - Height) / 2
  'frmLoginCaixa.Picture = LoadPicture("C:\sistemas\DMAC Caixa\Imagens\TelaLogin.jpg")
  wFechamentoGeral = False

End Sub
Private Sub VerificaSaldoCaixa()

  sql = "Select max(CTR_Protocolo) as Sequencia from ControleCaixa where CTR_Operador <> '99' AND CTR_numeroCaixa = " & GLB_Caixa

   RsSaldoCaixa.CursorLocation = adUseClient
   RsSaldoCaixa.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
   If RsSaldoCaixa.EOF Then
      MsgBox "Problemas com Saldo do Caixa", vbCritical, "Aviso"
      RsSaldoCaixa.Close
      Call AlterarResolucao(resolucaoOriginal.Colunas, resolucaoOriginal.Linhas)
      Unload Me
      Exit Sub
   Else
      If IsNull(RsSaldoCaixa("Sequencia")) Then
         GuardaSequencia = 0
         saldoAnterior = 0
         
      Else
         GuardaSequencia = RsSaldoCaixa("Sequencia")
         'RsSaldoCaixa.Close

      End If
   End If
   
   RsSaldoCaixa.Close
   
End Sub
Private Sub ComposicaoSaldoAnterior()
    
  sql = "select * from MovimentoCaixa Where MC_Protocolo = " & GuardaSequencia _
      & " and MC_Grupo like '70%'"
  rdoFormaPagamento.CursorLocation = adUseClient
  rdoFormaPagamento.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
       
  If Not rdoFormaPagamento.EOF Then
     Do While Not rdoFormaPagamento.EOF
        If rdoFormaPagamento("MC_Grupo") = "70101" Then
           wGrupo = "11006"
        ElseIf rdoFormaPagamento("MC_Grupo") = "70201" Then
            wGrupo = "11007"
        ElseIf rdoFormaPagamento("MC_Grupo") = "70204" Then
            wGrupo = "11008"
        End If
        sql = "Insert into movimentocaixa (MC_NumeroEcf,MC_NroCaixa,MC_CodigoOperador,MC_Loja, MC_Data, MC_Grupo,MC_Subgrupo, MC_Documento,MC_Serie," _
            & "MC_Valor, MC_banco, MC_Agencia,MC_Contacorrente, MC_bomPara, MC_Parcelas, MC_Remessa,MC_SituacaoEnvio,MC_Protocolo,MC_Pedido,MC_DataProcesso,MC_TipoNota)" _
            & " values(" & GLB_ECF & "," & GLB_Caixa & ",'" & GLB_USU_Codigo & "','" & GLB_Loja & "','" _
            & Format(Date, "yyyy/mm/dd") & "','" & wGrupo & "','',0,'SC'," & ConverteVirgula(rdoFormaPagamento("MC_Valor")) & ",0,0,0,0,0,9,'A'," & GLB_CTR_Protocolo & ",'0','" & Format(Date, "yyyy/mm/dd") & "','V')"
        rdoCNLoja.Execute (sql)
        rdoFormaPagamento.MoveNext
     Loop
  
  
  ''Arrumar
  'Else
 ''      sql = "select * from MovimentoCaixa Where MC_Grupo = '11006' and mc_data >= '" & Format(Date, "yyyy/mm/dd") & "'"
   '     rdoFormaPagamento.CursorLocation = adUseClient
    '    rdoFormaPagamento.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
 '''
 '       If rdoFormaPagamento.EOF Then
 '         sql = "Insert into movimentocaixa (MC_NumeroEcf,MC_NroCaixa,MC_CodigoOperador,MC_Loja, MC_Data, MC_Grupo,MC_Subgrupo, MC_Documento,MC_Serie," _
              & "MC_Valor, MC_banco, MC_Agencia,MC_Contacorrente, MC_bomPara, MC_Parcelas, MC_Remessa,MC_SituacaoEnvio,MC_Protocolo,MC_Pedido,MC_DataProcesso,MC_TipoNota)" _
 '             & " values(" & GLB_ECF & "," & GLB_Caixa & ",'" & GLB_USU_Codigo & "','" & GLB_Loja & "','" _
 '             & Format(Date, "yyyy/mm/dd") & "','" & wGrupo & "','',0,'SC',0,0,0,0,0,0,9,'A'," & GLB_CTR_Protocolo & ",'0','" & Format(Date, "yyyy/mm/dd") & "','V')"
 '         rdoCNLoja.Execute (sql)
 '       End If
  ''''''
  End If
  rdoFormaPagamento.Close
End Sub
Private Sub GuardaProtocolo()
    sql = "Select * from ControleCaixa where CTR_Supervisor <> 99 and CTR_SituacaoCaixa = 'A' and CTR_NumeroCaixa =" & GLB_Caixa
          RsControleCaixa.CursorLocation = adUseClient
          RsControleCaixa.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    GLB_CTR_Protocolo = RsControleCaixa("CTR_Protocolo")
    RsControleCaixa.Close
End Sub


Private Sub Image2_Click()

End Sub

Private Sub txtOperador_GotFocus()
   txtOperador.text = ""
   txtOperador.SelStart = 0
   txtOperador.SelLength = Len(txtOperador.text)
End Sub

Private Sub txtOperador_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    EfetuarLogin
End If

If KeyAscii = 27 Then
   Call AlterarResolucao(resolucaoOriginal.Colunas, resolucaoOriginal.Linhas)
   Unload Me
   Unload frmFundoEscuro
   End
End If

End Sub

Private Sub txtSenha_GotFocus()
   txtSenha.text = ""
   txtSenha.SelStart = 0
   txtSenha.SelLength = Len(txtSenha.text)
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If GLB_USU_Nome = Empty Then
            EfetuarLogin
        Else
            modoAdministrador
        End If
    End If
    
    If KeyAscii = 27 Then
       Call AlterarResolucao(resolucaoOriginal.Colunas, resolucaoOriginal.Linhas)
       Unload Me
       Unload frmFundoEscuro
       End
    End If
End Sub

