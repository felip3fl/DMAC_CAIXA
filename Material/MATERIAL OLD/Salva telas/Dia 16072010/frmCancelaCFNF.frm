VERSION 5.00
Begin VB.Form frmCancelaCFNF 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Cancela CF/NF"
   ClientHeight    =   2625
   ClientLeft      =   3030
   ClientTop       =   7650
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCancelaCFNF.frx":0000
   ScaleHeight     =   2625
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtSenha 
      BackColor       =   &H0081E8FA&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2175
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   1725
      Width           =   1035
   End
   Begin VB.TextBox txtValorNF 
      BackColor       =   &H0081E8FA&
      Height          =   315
      Left            =   3285
      TabIndex        =   7
      Top             =   1140
      Width           =   1260
   End
   Begin VB.TextBox txtPedido 
      BackColor       =   &H0081E8FA&
      Height          =   315
      Left            =   2190
      TabIndex        =   6
      Top             =   1140
      Width           =   1035
   End
   Begin VB.TextBox txtSerie 
      BackColor       =   &H0081E8FA&
      Height          =   315
      Left            =   1575
      TabIndex        =   3
      Top             =   1140
      Width           =   555
   End
   Begin VB.TextBox txtNotaFiscal 
      BackColor       =   &H0081E8FA&
      Height          =   315
      Left            =   450
      TabIndex        =   2
      Top             =   1140
      Width           =   1065
   End
   Begin Balcao2010.chameleonButton cmdGrava 
      Height          =   435
      Left            =   4710
      TabIndex        =   9
      Top             =   2040
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   767
      BTYPE           =   13
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCancelaCFNF.frx":355B6
      PICN            =   "frmCancelaCFNF.frx":355D2
      PICH            =   "frmCancelaCFNF.frx":36224
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton cmbSair 
      Height          =   435
      Left            =   4710
      TabIndex        =   10
      Top             =   150
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   767
      BTYPE           =   13
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCancelaCFNF.frx":36E76
      PICN            =   "frmCancelaCFNF.frx":36E92
      PICH            =   "frmCancelaCFNF.frx":376E4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1530
      TabIndex        =   12
      Top             =   1830
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancelar NF/CF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1635
      TabIndex        =   8
      Top             =   240
      Width           =   1665
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3300
      TabIndex        =   5
      Top             =   915
      Width           =   1005
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2205
      TabIndex        =   4
      Top             =   915
      Width           =   960
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1590
      TabIndex        =   1
      Top             =   915
      Width           =   450
   End
   Begin VB.Label lblnroNFCF 
      AutoSize        =   -1  'True
      BackColor       =   &H0081E8FA&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. NF/CF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   465
      TabIndex        =   0
      Top             =   915
      Width           =   990
   End
End
Attribute VB_Name = "frmCancelaCFNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub chameleonButton4_Click()

'Retorno = Bematech_FI_CancelaCupom()

' Unload Me
'End Sub

'Private Sub cmbSair_Click()
' Unload Me
'End Sub

'Private Sub Form_Load()
'frmCancelaCFNF.Left = 105
'frmCancelaCFNF.Top = 7520
'End Sub
Dim SQL As String
'Dim ADOCancela As rdoResultset
Dim ISQL As rdoResultset
'Dim RsPegaGrupoMovCaixa As rdoResultset
Dim WTESTACAMPOS As Boolean
Dim WAUX As Double
Dim wUltimoCupom As Double
Dim wDataDia As Date
Dim WGrupoAtualzado As Double
Dim WnumeroPedido As Double
Dim wWhere As String

Private Sub cmbSair_Click()
 Unload Me
End Sub

Private Sub cmdGrava_Click()

If Trim(UCase(txtSerie.Text = "CF")) Then
   wCancelaVenda = 1
Else
   wCancelaVenda = 2
End If
   
   
WTESTACAMPOS = False
wVerificaTM = False

SQL = ""
SQL = "Select CT_Loja from Controle"

rspegaloja.CursorLocation = adUseClient
rspegaloja.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    

    
    If Not rspegaloja.EOF Then
        wLoja = rspegaloja("CT_Loja")
    End If
    
rspegaloja.Close

    
        If Trim(txtNotaFiscal.Text) = "" Then
            MsgBox "Favor digite o Numero do cupom fiscal", vbInformation, "Aviso"
            txtNotaFiscal.SelStart = 0
            txtNotaFiscal.SelLength = Len(txtNotaFiscal.Text)
            txtNotaFiscal.SetFocus
            Exit Sub
            
        ElseIf IsNumeric(txtNotaFiscal.Text) = False Then
               MsgBox "Numero do Cupom Fiscal Inválido", vbCritical, "Atenção"
               txtNotaFiscal.SelStart = 0
               txtNotaFiscal.SelLength = Len(txtNotaFiscal.Text)
               txtNotaFiscal.SetFocus
               Exit Sub
                                        
        ElseIf Trim(txtSerie.Text) = "" Then
               MsgBox "Favor informe a série", vbInformation, "Atenção"
               txtSerie.SelStart = 0
               txtSerie.SelLength = Len(txtSerie.Text)
               txtSerie.SetFocus
               Exit Sub
        
        ElseIf txtSenha.Text = "" Then
               MsgBox "Favor digite a senha", vbInformation, "Aviso"
               txtSenha.SelStart = 0
               txtSenha.SelLength = Len(txtSenha.Text)
               txtSenha.SetFocus
               Exit Sub
        End If
     

   If txtSerie.Text = "CF" Then
      wWhere = " and EcfNF = " & GLB_ECF
   Else
      wWhere = ""
   End If
   
SQL = "SELECT TIPONOTA,NumeroPed, SERIE, NF, TOTALNOTA, DATAEMI, CT_SENHALIBERACAO " _
        & "FROM NFCAPA, CONTROLE " _
        & "WHERE SERIE = '" & txtSerie.Text & "' AND NF = " & txtNotaFiscal.Text & " " & Where _
        & "AND CT_SENHALIBERACAO = '" & txtSenha.Text & "'"
    
ADOCancela.CursorLocation = adUseClient
ADOCancela.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    


If Not ADOCancela.EOF Then
 
 ' Quando cancelamento de Cupom entra aqui
 
  If wCancelaVenda = 1 Then
         If ADOCancela("TIPONOTA") = "CA" Then
            MsgBox "Este cupom ja foi cancelado", vbCritical, "Atenção"
            txtNotaFiscal.SelStart = 0
            txtNotaFiscal.SelLength = Len(txtNotaFiscal.Text)
            txtNotaFiscal.SetFocus
           
            Exit Sub
                                
        ElseIf ADOCancela("DATAEMI") <> Date Then
            MsgBox "Esta nota não pode ser cancelada data ultrapassada", vbInformation, "Aviso"
           
            Exit Sub
        Else
        SQL = "Select nf from nfcapa where serie = 'CF' and DataEmi = '" & Format(Date, _
                    "mm/dd/yyyy") & "' and ECFNF=" & Val(GLB_ECF) & "  order by nf desc "
         
               ADOCancela.CursorLocation = adUseClient
               ADOCancela.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            
            If Not ADOCancela.EOF Then
                If ADOCancela("nf") <> Val(txtNotaFiscal.Text) Then
                    wUltimoCupom = 0
                    wUltimoCupom = ADOCancela("nf")
                    MsgBox "Ultimo Cupom = " & wUltimoCupom & " " & "  " & "(Você só pode Cancelar o Ultimo)", vbCritical, "Atenção"
                    txtSenha.Text = ""
                    txtSerie.Text = ""
                    txtNotaFiscal.Text = ""
                    txtValorNF.Text = ""
                    txtNotaFiscal.SelStart = 0
                    txtNotaFiscal.SelLength = Len(txtNotaFiscal.Text)
                    txtNotaFiscal.SetFocus
                    Exit Sub
                Else
                    SQL = "UPDATE NFCAPA SET TIPONOTA = 'CA',SituacaoEnvio='A' WHERE NF = " _
                          & txtNotaFiscal.Text & " and Serie = '" & txtSerie.Text & "'"
                    rdoCNLoja.Execute (SQL)
                    WTESTACAMPOS = True
                    
                    'GravaSequenciaLeitura 95, txtNotaFiscal.Text, "0"
                
                    SQL = "Select MC_Grupo,MC_Sequencia from MovimentoCaixa " _
                    & "where MC_Documento =" & txtNotaFiscal.Text & " "
                      '  Set RsPegaGrupoMovCaixa = rdoCNLoja.OpenResultset(SQL)
                          RsPegaGrupoMovCaixa.CursorLocation = adUseClient
                          RsPegaGrupoMovCaixa.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                    
                    If Not RsPegaGrupoMovCaixa.EOF Then
                        Do While Not RsPegaGrupoMovCaixa.EOF
                            If Mid(RsPegaGrupoMovCaixa("MC_Grupo"), 1, 1) = 1 Then
                                wPegaGrupo = RsPegaGrupoMovCaixa("MC_Grupo")
                                WGrupoAtualzado = 0
                                WGrupoAtualzado = 9 & Mid(RsPegaGrupoMovCaixa("MC_Grupo"), 2, Len(RsPegaGrupoMovCaixa("MC_Grupo")))
                                SQL = "Update MovimentoCaixa set MC_Grupo = " & WGrupoAtualzado & ", " _
                                    & "MC_SituacaoEnvio='A' " _
                                    & "where MC_Sequencia = " & RsPegaGrupoMovCaixa("MC_Sequencia") & " " _
                                    & "and MC_Documento = " & txtNotaFiscal.Text & " " _
                                    & "and MC_Grupo = " & wPegaGrupo
                                    rdoCNLoja.Execute (SQL)
                                
                                'GravaSequenciaLeitura 94, RsPegaGrupoMovCaixa("MC_Sequencia"), "0"
                            ElseIf Mid(RsPegaGrupoMovCaixa("MC_Grupo"), 1, 1) = 2 Then
                                wPegaGrupo = RsPegaGrupoMovCaixa("MC_Grupo")
                                WGrupoAtualzado = 0
                                SQL = "Update MovimentoCaixa set MC_Serie = 'CA', " _
                                    & "MC_SituacaoEnvio='A' " _
                                    & "where MC_Sequencia = " & RsPegaGrupoMovCaixa("MC_Sequencia") & " " _
                                    & "and MC_Documento = " & txtNotaFiscal.Text & " " _
                                    & "and MC_Grupo = " & wPegaGrupo
                                    rdoCNLoja.Execute (SQL)
                                
                                'GravaSequenciaLeitura 94, RsPegaGrupoMovCaixa("MC_Sequencia"), "0"
                            
                            End If
                            RsPegaGrupoMovCaixa.MoveNext
                        Loop
                    End If
                End If
            Else
                MsgBox "Informação não encontrada", vbInformation, "Aviso"
                Exit Sub
            End If
        
'        Else: SQL = "UPDATE NFCAPA SET TIPONOTA = 'CA' WHERE NF = " & txtNotaFiscal.Text & "                     "
'            rdoCnLoja.Execute (SQL)
'            WTESTACAMPOS = True
                                        
        If WTESTACAMPOS = True Then
            SQL = "SELECT CT_DATA, CT_SITUACAO FROM CTCAIXA " _
            & "WHERE CT_SITUACAO = 'A'"
            'Set ADOCancela = rdoCNLoja.OpenResultset(SQL)
            ADOCancela.CursorLocation = adUseClient
            ADOCancela.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                                            
            If Not ADOCancela.EOF Then
                wData = ADOCancela("CT_DATA")
                SQL = "INSERT INTO MOVIMENTOCAIXA (MC_GRUPO, MC_DATA, MC_VALOR, " _
                & "MC_DOCUMENTO, MC_BANCO, MC_AGENCIA, MC_CONTACORRENTE, MC_BOMPARA, MC_REMESSA,MC_Loja,MC_SituacaoEnvio,MC_Serie) " _
                & "VALUES (" & 30105 & ", '" & Format(wData, "MM/DD/YYYY") & "', " & ConverteVirgula(txtValorNF.Text) & ", " _
                & "" & txtNotaFiscal.Text & ", " & 0 & ", '" & 0 & "', " & 0 & ", '" & Format(wData, "MM/DD/YYYY") & "', " & 0 & ",'" & wLoja & "','A','A') "
                rdoCNLoja.Execute (SQL)
                
                SQL = ""
                SQL = "Update nfitens set tipomovimentacao = 21, " _
                    & "SituacaoEnvio='A', TipoNota='CA' " _
                    & "Where nf =" & txtNotaFiscal.Text & " " _
                    & "And Serie = '" & txtSerie.Text & "' "
                rdoCNLoja.Execute (SQL)
                
                'AtualizaEstoque txtNotaFiscal.Text, txtSerie.Text, 1
                
                Call CancelaCupomFiscal
                MsgBox "Cupom cancelado com sucesso", vbInformation, "Aviso"
            
            Else
                
                MsgBox "ARQUIVO NÃO ENCONTRADO", vbInformation, "ATENÇÃO"
                Exit Sub
            
            End If
        End If
    End If
    
 
Else
 
 ' Quando cancelamento de Nota entra aqui
 
       If ADOCancela("TIPONOTA") = "CA" Then
            MsgBox "Esta nota ja foi cancelada", vbCritical, "Atenção"
            txtNotaFiscal.SelStart = 0
            txtNotaFiscal.SelLength = Len(txtNotaFiscal.Text)
            txtNotaFiscal.SetFocus
            Call LimpaCampos
            Exit Sub
                                
        ElseIf ADOCancela("DATAEMI") <> Date Then
            MsgBox "Esta nota não pode ser cancelada data ultrapassada", vbInformation, "Aviso"
            txtNotaFiscal.SelStart = 0
            txtNotaFiscal.SelLength = Len(txtNotaFiscal.Text)
            txtNotaFiscal.SetFocus
            Call LimpaCampos
            Exit Sub
                                
        Else: SQL = "UPDATE NFCAPA SET TIPONOTA = 'CA',SituacaoEnvio='A' WHERE NF = " _
                 & txtNotaFiscal.Text & " and Serie = '" & txtSerie.Text & "'"
                rdoCNLoja.Execute (SQL)
                WTESTACAMPOS = True
                
                 'GravaSequenciaLeitura 95, txtNotaFiscal.Text, "0"
                 
                 SQL = "Select MC_Grupo,MC_Sequencia from MovimentoCaixa " _
                    & "where MC_Documento =" & txtNotaFiscal.Text & " "
                   ' Set RsPegaGrupoMovCaixa = rdoCNLoja.OpenResultset(SQL)
                     RsPegaGrupoMovCaixa.CursorLocation = adUseClient
                     RsPegaGrupoMovCaixa.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                    
                    If Not RsPegaGrupoMovCaixa.EOF Then
                        Do While Not RsPegaGrupoMovCaixa.EOF
                            If Mid(RsPegaGrupoMovCaixa("MC_Grupo"), 1, 1) = 1 Then
                                wPegaGrupo = RsPegaGrupoMovCaixa("MC_Grupo")
                                WGrupoAtualzado = 0
                                WGrupoAtualzado = 9 & Mid(RsPegaGrupoMovCaixa("MC_Grupo"), 2, Len(RsPegaGrupoMovCaixa("MC_Grupo")))
                                SQL = "Update MovimentoCaixa set MC_Grupo = " & WGrupoAtualzado & ", " _
                                    & "MC_SituacaoEnvio='A' " _
                                    & "where MC_Sequencia = " & RsPegaGrupoMovCaixa("MC_Sequencia") & " " _
                                    & "and MC_Documento = " & txtNotaFiscal.Text & " " _
                                    & "and MC_Grupo = " & wPegaGrupo
                                    rdoCNLoja.Execute (SQL)
                            
                                'GravaSequenciaLeitura 94, RsPegaGrupoMovCaixa("MC_Sequencia"), "0"
                            ElseIf Mid(RsPegaGrupoMovCaixa("MC_Grupo"), 1, 1) = 2 Then
                                wPegaGrupo = RsPegaGrupoMovCaixa("MC_Grupo")
                                WGrupoAtualzado = 0
                                SQL = "Update MovimentoCaixa set MC_Serie = 'CA', " _
                                    & "MC_SituacaoEnvio='A' " _
                                    & "where MC_Sequencia = " & RsPegaGrupoMovCaixa("MC_Sequencia") & " " _
                                    & "and MC_Documento = " & txtNotaFiscal.Text & " " _
                                    & "and MC_Grupo = " & wPegaGrupo
                                    rdoCNLoja.Execute (SQL)
                                'GravaSequenciaLeitura 94, RsPegaGrupoMovCaixa("MC_Sequencia"), "0"
                            End If
                            RsPegaGrupoMovCaixa.MoveNext
                        Loop
                    End If
                
                
                
        
        If WTESTACAMPOS = True Then
            SQL = "SELECT CT_DATA, CT_SITUACAO FROM CTCAIXA " _
            & "WHERE CT_SITUACAO = 'A'"
            'Set ADOCancela = rdoCNLoja.OpenResultset(SQL)
            ADOCancela.CursorLocation = adUseClient
            ADOCancela.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                                            
            If Not ADOCancela.EOF Then
                wData = ADOCancela("CT_DATA")
                SQL = "INSERT INTO MOVIMENTOCAIXA (MC_GRUPO, MC_DATA, MC_VALOR, " _
                    & "MC_DOCUMENTO, MC_BANCO, MC_AGENCIA, MC_CONTACORRENTE, MC_BOMPARA, MC_REMESSA,MC_Loja,MC_SituacaoEnvio,MC_Serie) " _
                    & "VALUES (" & 30105 & ", '" & Format(wData, "mm/DD/YYYY") & "', " & ConverteVirgula(txtValorNF.Text) & ", " _
                    & "" & txtNotaFiscal.Text & ", " & 0 & ", '" & 0 & "', " & 0 & ", '" & Format(wData, "MM/DD/YYYY") & "', " & 0 & ",'" & wLoja & "','A','" & txtSerie.Text & "') "
                rdoCNLoja.Execute (SQL)
                   
                If WTipoNota = "T" Then
                    SQL = ""
                    SQL = "Update nfitens set tipomovimentacao = 25, " _
                        & "SituacaoEnvio='A',TipoNota='CA' " _
                        & "Where nf =" & txtNotaFiscal.Text & " " _
                        & "And Serie = '" & txtSerie.Text & "' "
                    rdoCNLoja.Execute (SQL)
                ElseIf WTipoNota = "E" Then
                    SQL = ""
                    SQL = "Update nfitens set tipomovimentacao = 14, " _
                        & "SituacaoEnvio='A',TipoNota='CA' " _
                        & "Where nf =" & txtNotaFiscal.Text & " " _
                        & "And Serie = '" & txtSerie.Text & "' "
                    rdoCNLoja.Execute (SQL)
                Else
                    SQL = ""
                    SQL = "Update nfitens set tipomovimentacao = 21, " _
                        & "SituacaoEnvio='A',TipoNota='CA' " _
                        & "Where nf =" & txtNotaFiscal.Text & " " _
                        & "And Serie = '" & txtSerie.Text & "' "
                    rdoCNLoja.Execute (SQL)
                End If
                'AtualizaEstoque txtNotaFiscal.Text, txtSerie.Text, 1
                
                MsgBox "Nota cancelada com sucesso", vbInformation, "Aviso"
            
            Else
                MsgBox "ARQUIVO NÃO ENCONTRADO", vbInformation, "ATENÇÃO"
                Exit Sub
            End If
        
        End If
    End If
End If

Else
        MsgBox "Registro não encontrado", vbInformation, "Aviso"
        txtNotaFiscal.SelStart = 0
        txtNotaFiscal.SelLength = Len(txtNotaFiscal.Text)
        txtNotaFiscal.SetFocus
       ' Call LimpaCampos
        Exit Sub
End If

    Call LimpaCampos
    
    txtNotaFiscal.SetFocus
       
End Sub



Private Sub cmdRetorna_Click()
Unload Me
End Sub

Private Sub cmdGravar_Click()

End Sub

Private Sub Form_Load()


    
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 3
   ' If wCancelaVenda = 1 Then
   '     fraCancelamento.Caption = "Cancelamento do Último Cupom Impresso"
   '     lblNotaFiscal.Caption = "Cupom Fiscal"
   ' Else
   '     fraCancelamento.Caption = "Cancelamento de Nota Fiscal"
   '     lblNotaFiscal.Caption = "Nota Fiscal"
   ' End If
End Sub




Private Sub txtNotaFiscal_LostFocus()
If txtNotaFiscal.Text = "" Then
    Exit Sub
End If

If IsNumeric(txtNotaFiscal.Text) = False Then
     
    txtNotaFiscal.Text = ""
    txtNotaFiscal.SetFocus
    Exit Sub
End If
End Sub

Private Sub txtSenha_LostFocus()
If txtSenha.Text = "" Then
    Exit Sub
End If

SQL = "SELECT CT_SENHALIBERACAO FROM CONTROLE WHERE CT_SENHALIBERACAO = '" & (txtSenha.Text) & "' "

'Set ADOCancela = rdoCNLoja.OpenResultset(SQL)
ADOCancela.CursorLocation = adUseClient
ADOCancela.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic


If Not ADOCancela.EOF Then
    If txtSenha.Text <> ADOCancela("CT_SENHALIBERACAO") Then
        MsgBox "Senha para cancelamento não confere", vbCritical, "Atenção"
        txtSenha.SelStart = 0
        txtSenha.SelLength = Len(txtSenha.Text)
        txtSenha.SetFocus
        Exit Sub
    End If
Else
    MsgBox "Senha não cadastrada", vbCritical, "Atenção"
    txtSenha.SelStart = 0
    txtSenha.SelLength = Len(txtSenha.Text)
    txtSenha.SetFocus
    Exit Sub
End If

ADOCancela.Close

End Sub

Private Sub txtSerie_LostFocus()
txtSerie.Text = UCase(txtSerie.Text)
If txtSerie.Text = "" Then
    Exit Sub
End If


'If txtSenha.Text = "" Then
'    MsgBox "Digite sua Senha", vbCritical, "Atenção"
'    txtSenha.SetFocus
'    txtSerie.Text = ""
'    txtValorNF.Text = ""
'    txtNotaFiscal.Text = ""
'    Exit Sub
'End If

If txtNotaFiscal.Text = "" Then
   MsgBox "Preencha todos os campos", vbCritical, "Atenção"
   txtNotaFiscal.SelStart = 0
   txtNotaFiscal.SelLength = Len(txtNotaFiscal.Text)
   txtNotaFiscal.SetFocus
   Exit Sub
End If

 

SQL = "SELECT TOTALNOTA, NF, SERIE,TipoNota FROM NFCAPA WHERE NF = " & txtNotaFiscal.Text & " " _
    & "AND SERIE = '" & txtSerie.Text & "' "
    
 ADOCancela.CursorLocation = adUseClient
 ADOCancela.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
'Set ADOCancela = rdoCNLoja.OpenResultset(SQL)

If Not ADOCancela.EOF Then
       txtValorNF.Text = Format(ADOCancela("TOTALNOTA"), "0.00")
       WTipoNota = ADOCancela("TipoNota")
    Else
        MsgBox "Registro não encontrado", vbInformation, "Aviso"
        txtNotaFiscal.SelStart = 0
        txtNotaFiscal.SelLength = Len(txtNotaFiscal.Text)
        txtNotaFiscal.SetFocus
End If
ADOCancela.Close

End Sub

Private Sub txtValorNF_GotFocus()



If wCancelaVenda = 1 Then
    If txtSerie.Text = "" Then
        Exit Sub
        ElseIf txtSerie.Text <> "CF" Then
            MsgBox "Série invalida campo CF", vbCritical, "Atenção"
       ' txtSerie.Text = ""
       ' txtValorNF.Text = ""
       ' txtSenha.Text = ""
       ' txtNotaFiscal.Text = ""
        txtSerie.SetFocus
        Exit Sub
        
        
    End If
End If

If wCancelaVenda = 2 Then
    If txtSerie.Text = "" Then
        Exit Sub
        ElseIf txtSerie.Text = "CF" Then
            MsgBox "Série invalida campo NF", vbCritical, "Atenção"
        'txtSerie.Text = ""
        'txtSenha.Text = ""
        'txtValorNF.Text = ""
        'txtNotaFiscal.Text = ""
        txtSerie.SetFocus
        Exit Sub
            
        
    End If
End If

End Sub


Sub LimpaCampos()

        txtSerie.Text = ""
        txtValorNF.Text = ""
        txtSenha.Text = ""
        txtNotaFiscal.Text = ""
        Exit Sub

End Sub

Private Sub CancelaCupomFiscal()
    
    
    Retorno = Bematech_FI_CancelaCupom()
    Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
    If Retorno = 1 Then
        Call AtualizaNumeroCupom
    End If

End Sub





