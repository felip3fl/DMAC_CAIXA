VERSION 5.00
Begin VB.Form frmOperacoesECF 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Forma de Pagamento"
   ClientHeight    =   7605
   ClientLeft      =   14670
   ClientTop       =   3075
   ClientWidth     =   4905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraPagamento 
      BackColor       =   &H00000000&
      Height          =   3315
      Left            =   135
      TabIndex        =   2
      Top             =   135
      Width           =   3615
      Begin VB.TextBox txtOperacao 
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
         Left            =   90
         TabIndex        =   4
         Top             =   2685
         Width           =   3435
      End
      Begin Balcao2010.chameleonButton chbCancelamenteECF 
         Height          =   720
         Left            =   75
         TabIndex        =   3
         Top             =   1575
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   1270
         BTYPE           =   14
         TX              =   "Cancelamento ECF"
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
         MICON           =   "frmOperacoesECF.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Balcao2010.chameleonButton chbLeituraZ 
         Height          =   720
         Left            =   75
         TabIndex        =   1
         Top             =   870
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   1270
         BTYPE           =   14
         TX              =   "Leitura Z"
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
         MICON           =   "frmOperacoesECF.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Balcao2010.chameleonButton chbLeituraX 
         Height          =   720
         Left            =   75
         TabIndex        =   0
         Top             =   180
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   1270
         BTYPE           =   14
         TX              =   "Leitura X"
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
         MICON           =   "frmOperacoesECF.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblOperacao 
         AutoSize        =   -1  'True
         BackColor       =   &H00AE7411&
         BackStyle       =   0  'Transparent
         Caption         =   "Senha/NroECF"
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
         Left            =   90
         TabIndex        =   5
         Top             =   2370
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmOperacoesECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sql As String
Dim wSenhaOperacaoECFOK As Boolean
Dim wTipoOperacao As String
' TipoOperacao
' C = Cancelamento
' CL = Cancelamento Liberado pelo getente, esperando nro ECF
' Z = Leitura Z
' X = Leitura X

Private Sub chbCancelamenteECF_Click()
    wTipoOperacao = "C"
    wSenhaOperacaoECFOK = False
    txtOperacao.Visible = True
    lblOperacao.Visible = True
    txtOperacao.PasswordChar = "*"
    txtOperacao = ""
    lblOperacao = "Senha do Gerente:"
    txtOperacao.SetFocus
End Sub

Private Sub chbCancelamenteECF_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
      Unload Me
      frmControlaCaixa.txtPedido.SetFocus
  End If
End Sub


Private Sub chbLeituraX_Click()
    Screen.MousePointer = vbHourglass
    wTipoOperacao = "X"
    Call LimpaOperacao
    Retorno = Bematech_FI_LeituraX()
    Call VerificaRetornoImpressora("", "", "Leitura X")
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub chbLeituraX_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
    frmControlaCaixa.txtPedido.SetFocus
  End If
End Sub


Private Sub chbLeituraZ_Click()
    wTipoOperacao = "Z"
    wSenhaOperacaoECFOK = False
    txtOperacao.Visible = True
    lblOperacao.Visible = True
    txtOperacao.PasswordChar = "*"
    txtOperacao = ""
    lblOperacao = "Senha do Gerente:"
    txtOperacao.SetFocus
End Sub

Private Sub chbLeituraZ_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
    frmControlaCaixa.txtPedido.SetFocus
  End If
End Sub

Private Sub Form_Load()
   Call AjustaTela(frmOperacoesECF)
   Call LimpaOperacao
   
End Sub



Private Sub txtOperacao_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Call LimpaOperacao
  End If
  
  If KeyAscii = 13 Then
  
    If Trim(txtOperacao.Text) = "" Or txtOperacao.Text = "'" Then
      MsgBox "Senha inválida!"
    End If
  
    If wTipoOperacao = "Z" Then
        Call VerificaSenhaGerente
        
        If wSenhaOperacaoECFOK = True Then
            If MsgBox("Está operação impossibilitará a emissão de Cupom Fiscal hoje. Deseja Continuar?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
                Screen.MousePointer = vbHourglass
                lblOperacao = "Imprimindo Leitura Z"
'                MsgBox "Leitura Z, antes de compilar, retirar msgbox e incluir cod comentado"
               Retorno = Bematech_FI_ReducaoZ("", "")
               Call VerificaRetornoImpressora("", "", "Redução Z")
                Screen.MousePointer = vbNormal
            End If
        Else
            MsgBox "Senha inválida!"
            txtOperacao.Text = ""
        End If
        Call LimpaOperacao

    ElseIf wTipoOperacao = "C" Then
        Call VerificaSenhaGerente
        If wSenhaOperacaoECFOK = True Then
          txtOperacao.PasswordChar = ""
          txtOperacao = ""
          lblOperacao = "Número ECF:"
          wTipoOperacao = "CL"
        Else
          MsgBox "Senha inválida!"
          txtOperacao.Text = ""
        End If
        
    ElseIf wTipoOperacao = "CL" Then
        Call CancelaECF
        End If
  End If
End Sub

Private Sub LimpaOperacao()
    wSenhaOperacaoECFOK = False
    wTipoOperacao = ""
    txtOperacao = ""
    lblOperacao = ""
    txtOperacao.Visible = False
    lblOperacao.Visible = False
'chbLeituraX.SetFocus
End Sub

Private Sub VerificaSenhaGerente()

    wSenhaOperacaoECFOK = False
    
    sql = "select LTrim(rtrim(usu_Senha)) from usuariocaixa where usu_tipousuario = 'S' " _
        & "and usu_Senha = '" & Trim(txtOperacao) & "'"
    rsOperacoes.CursorLocation = adUseClient
    rsOperacoes.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

    If Not rsOperacoes.EOF = True Then
        wSenhaOperacaoECFOK = True
    End If
    
    rsOperacoes.Close
    

    End Sub
Private Sub CancelaECF()

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
       
        sql = "SELECT TIPONOTA,NumeroPed, SERIE, NF, TOTALNOTA, DATAEMi,Numeroped " _
            & " FROM NFCAPA WHERE " _
            & " SERIE = '" & GLB_SerieCF & "' AND " _
            & " NF = " & txtOperacao.Text
         
        ADOCancela.CursorLocation = adUseClient
        ADOCancela.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            
      If ADOCancela.EOF = True Then
           MsgBox "Cupom Fiscal incorreto"
            ADOSituacao.Close
            ADOCancela.Close
           Exit Sub
      Else
                 
       'Cancelando cupom fiscal
       
        Dim Numerocupom As String
        Dim RetornoStatus As String
    
        If (LocalRetorno = "1") Then 'Grava retorno em arquivo
           Numerocupom = Space(1)
        Else
           Numerocupom = Space(6)
        End If
    
        Retorno = Bematech_FI_NumeroCupom(Numerocupom)
        'Função que analisa o retorno da impressora
'        Call VerificaRetornoImpressora("Número do Último Cupom: ", _
'        NumeroCupom, "Informações da Impressora")
         
        If Val(Numerocupom) = Val(txtOperacao.Text) Then
          Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
          
          Retorno = Bematech_FI_CancelaCupom()
          'Função que analisa o retorno da impressora
          Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
          
          'Cancelando no sistema
          If Retorno = 1 Then
          
''              SQL = "Select * from controlecaixa where " _
''                  & " Ctr_DataInicial between '" & Format(Date, "yyyy/mm/dd") & " 00:00:00' and  '" _
''                  & Format(Date, "yyyy/mm/dd") & " 23:59:59'"
''
''
''              PegaLoja.CursorLocation = adUseClient
''              PegaLoja.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
''
''              SQL = "INSERT INTO MOVIMENTOCAIXA (MC_GRUPO,MC_SubGrupo, MC_DATA, MC_VALOR, " _
''                & "MC_DOCUMENTO, MC_BANCO, MC_AGENCIA, MC_CONTACORRENTE, MC_BOMPARA, " _
''                & "MC_REMESSA,MC_Loja,MC_SituacaoEnvio,MC_Serie, " _
''                & " MC_NumeroECF,MC_CodigoOperador,MC_Pedido,MC_DataProcesso,MC_TipoNota) " _
''                & "VALUES (" & 30105 & ",'', '" & Format(wData, "yyyy/mm/dd") _
''                & "', " & ADOCancela("TotalNota") & ", " _
''                & ADOCancela("NF") & ", " & 0 & ", '" & 0 & "', " & 0 & ", '" _
''                & Format(wData, "yyyy/mm/dd") & "', " & 0 & ",'" _
''                & GLB_Loja & "','A','" & Trim(ADOCancela("Serie")) & "', " & GLB_ECF & ",'" _
''                & PegaLoja("ctr_operador") & "','" & ADOCancela("Numeroped") & "','" _
''                & Format(wData, "yyyy/mm/dd") & "','C')"
''
''                rdoCNLoja.Execute (SQL)
''
''               PegaLoja.Close
             
               sql = "exec SP_Cancela_NotaFiscal " & ADOCancela("NF") & ",'" & Trim(ADOCancela("Serie")) & "'"
               rdoCNLoja.Execute (sql)
               
               MsgBox "Nota cancelada com sucesso", vbInformation, "Aviso"
               Call LimpaOperacao
          Else
             MsgBox "Cupom Fiscal NÃO foi cancelado"
          End If
        Else
          MsgBox "Só é possível o cancelamento da ultima transição ECF"
        End If
        
     End If
         ADOSituacao.Close
         ADOCancela.Close
End Sub
