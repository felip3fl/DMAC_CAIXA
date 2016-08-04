VERSION 5.00
Begin VB.Form frmCancelaCFNF 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Cancela CF/NF"
   ClientHeight    =   2625
   ClientLeft      =   4605
   ClientTop       =   5310
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
      TabIndex        =   4
      Top             =   1725
      Width           =   1035
   End
   Begin VB.TextBox txtValorNF 
      BackColor       =   &H0081E8FA&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3285
      TabIndex        =   3
      Top             =   1140
      Width           =   1260
   End
   Begin VB.TextBox txtPedido 
      BackColor       =   &H0081E8FA&
      Enabled         =   0   'False
      Height          =   315
      Left            =   2190
      TabIndex        =   2
      Top             =   1140
      Width           =   1035
   End
   Begin VB.TextBox txtSerie 
      BackColor       =   &H0081E8FA&
      Height          =   315
      Left            =   1575
      TabIndex        =   1
      Top             =   1140
      Width           =   555
   End
   Begin VB.TextBox txtNotaFiscal 
      BackColor       =   &H0081E8FA&
      Height          =   315
      Left            =   450
      TabIndex        =   0
      Top             =   1140
      Width           =   1065
   End
   Begin Balcao2010.chameleonButton cmdGrava 
      Height          =   435
      Left            =   4710
      TabIndex        =   5
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
      TabIndex        =   6
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   915
      Width           =   990
   End
End
Attribute VB_Name = "frmCancelaCFNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String
Dim wUltimoCupom As Double
Dim WGrupoAtualzado As Double
Dim WnumeroPedido As Double
Dim wWhere As String

Private Sub cmbSair_Click()
 Unload Me
End Sub

Private Sub cmdGrava_Click()


        
        If Trim(txtNotaFiscal.Text) = "" Then
            MsgBox "Favor digite o Numero NF/CF ", vbInformation, "Aviso"
            txtNotaFiscal.SelStart = 0
            txtNotaFiscal.SelLength = Len(txtNotaFiscal.Text)
            txtNotaFiscal.SetFocus
            Exit Sub
            
        ElseIf IsNumeric(txtNotaFiscal.Text) = False Then
               MsgBox "Numero NF/CF Inválido", vbCritical, "Atenção"
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
      '  ElseIf Trim(txtPedido.Text) = "" Then
      '         MsgBox "Favor digite o Numero do Pedido", vbInformation, "Aviso"
      '         txtPedido.SelStart = 0
      '         txtPedido.SelLength = Len(txtPedido.Text)
      '         txtPedido.SetFocus
      '         Exit Sub
      '
      '  ElseIf IsNumeric(txtPedido.Text) = False Then
      '         MsgBox "Numero do Pedido Inválido", vbCritical, "Atenção"
      '         txtPedido.SelStart = 0
      '         txtPedido.SelLength = Len(txtPedido.Text)
      '         txtPedido.SetFocus
      '         Exit Sub
        End If
      
      If Trim(UCase((txtSenha.Text))) <> Trim(UCase(wSenhaLiberacao)) Then
         MsgBox "Senha para cancelamento não confere", vbCritical, "Atenção"
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
     
   
        If Trim(UCase(txtSerie.Text = "CF")) Then
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
                 ADOCancela.Close
                 Exit Sub
              End If
           End If
           ADOCancela.Close
        End If
   
        SQL = "SELECT CTR_DATAINICIAL, CTR_SITUACAOCAIXA FROM  CONTROLECAIXA " _
            & " WHERE CTR_SITUACAOCAIXA = 'A'"
     
        ADOSituacao.CursorLocation = adUseClient
        ADOSituacao.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                                          
        If Not ADOSituacao.EOF Then
           wData = ADOSituacao("CTR_DATAINICIAL")
        Else
           MsgBox "Caixa Fechado", vbInformation, "Aviso"
           ADOSituacao.Close
           Exit Sub
        End If
       
          
        ADOSituacao.Close
     
      
       'SQL = "SELECT TIPONOTA,NumeroPed, SERIE, NF, TOTALNOTA, DATAEMI  " _
            & " FROM NFCAPA WHERE " _
            & " SERIE = '" & txtSerie.Text & "' AND " _
            & " TIPONOTA <> 'CA' and NumeroPed = " & txtPedido.Text And "" _
            & " NF = " & txtNotaFiscal.Text & " " & Where
        
        SQL = "SELECT TIPONOTA,NumeroPed, SERIE, NF, TOTALNOTA, DATAEMI " _
            & " FROM NFCAPA WHERE " _
            & " SERIE = '" & txtSerie.Text & "' AND " _
            & " TIPONOTA <> 'CA' and " _
            & " NF = " & txtNotaFiscal.Text & " " & Where
         
        ADOCancela.CursorLocation = adUseClient
        ADOCancela.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
 

        If Not ADOCancela.EOF Then
           If ADOCancela("DATAEMI") <> Date Then
              MsgBox "Esta NF/CF não pode ser cancelado data ultrapassada", vbInformation, "Aviso"
              ADOCancela.Close
              Exit Sub
           End If
                 
                SQL = "INSERT INTO MOVIMENTOCAIXA (MC_GRUPO, MC_DATA, MC_VALOR, " _
                & "MC_DOCUMENTO, MC_BANCO, MC_AGENCIA, MC_CONTACORRENTE, MC_BOMPARA, MC_REMESSA,MC_Loja,MC_SituacaoEnvio,MC_Serie) " _
                & "VALUES (" & 30105 & ", '" & Format(wData, "MM/DD/YYYY") & "', " & ConverteVirgula(txtValorNF.Text) & ", " _
                & "" & txtNotaFiscal.Text & ", " & 0 & ", '" & 0 & "', " & 0 & ", '" & Format(wData, "MM/DD/YYYY") & "', " & 0 & ",'" _
                & wlblloja & "','A','CA') "
                                              
                rdoCNLoja.Execute (SQL)
               
                If Trim(UCase(txtSerie.Text = "CF")) Then
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
                   MsgBox "Nota cancelada com sucesso", vbInformation, "Aviso"
                   
                End If
         Else
             MsgBox "NF/CF não encontrado", vbInformation, "Aviso"
             ADOCancela.Close
             Exit Sub
         End If
        
         ADOCancela.Close

         SQL = "UPDATE NFCAPA SET TIPONOTA = 'CA',SituacaoEnvio='A' WHERE NF = " _
              & txtNotaFiscal.Text & " and Serie = '" & txtSerie.Text & "'"
                rdoCNLoja.Execute (SQL)
                        
         SQL = "Select MC_Grupo,MC_Sequencia from MovimentoCaixa " _
             & "where MC_Documento =" & txtNotaFiscal.Text & " "
                 
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
                                
                                 
                   ElseIf Mid(RsPegaGrupoMovCaixa("MC_Grupo"), 1, 1) = 2 Then
                          wPegaGrupo = RsPegaGrupoMovCaixa("MC_Grupo")
                          WGrupoAtualzado = 0
                          SQL = "Update MovimentoCaixa set MC_Serie = 'CA', " _
                              & "MC_SituacaoEnvio='A' " _
                              & "where MC_Sequencia = " & RsPegaGrupoMovCaixa("MC_Sequencia") & " " _
                              & "and MC_Documento = " & txtNotaFiscal.Text & " " _
                              & "and MC_Grupo = " & wPegaGrupo
                          rdoCNLoja.Execute (SQL)
                                
                                
                            
                   End If
                   RsPegaGrupoMovCaixa.MoveNext
                   Loop
             End If
             RsPegaGrupoMovCaixa.Close

    Call LimpaCampos
    
    txtNotaFiscal.SetFocus
       
End Sub

Private Sub cmdRetorna_Click()
Unload Me
End Sub

Private Sub Form_Load()
    
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 3

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



Private Sub txtSenha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdGrava_Click
End If

End Sub

Private Sub txtSerie_LostFocus()
txtSerie.Text = UCase(txtSerie.Text)
If txtSerie.Text = "" Then
    Exit Sub
End If

If txtNotaFiscal.Text = "" Then
   MsgBox "Preencha todos os campos", vbCritical, "Atenção"
   txtNotaFiscal.SelStart = 0
   txtNotaFiscal.SelLength = Len(txtNotaFiscal.Text)
   txtNotaFiscal.SetFocus
   Exit Sub
End If

 

'SQL = "SELECT TOTALNOTA, NF, SERIE,TipoNota FROM NFCAPA WHERE NF = " & txtNotaFiscal.Text & " " _
    & "AND SERIE = '" & txtSerie.Text & "' AND numeroped = " & txtPedido.Text & ""
    
SQL = "SELECT TOTALNOTA, NF, SERIE,TipoNota,numeroped FROM NFCAPA WHERE NF = " & txtNotaFiscal.Text & " " _
    & "AND SERIE = '" & UCase(Trim(txtSerie.Text)) & "' and TIPONOTA <> 'CA'"
    
 ADOCancela.CursorLocation = adUseClient
 ADOCancela.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

If Not ADOCancela.EOF Then
       txtValorNF.Text = Format(ADOCancela("TOTALNOTA"), "0.00")
       txtPedido.Text = ADOCancela("numeroped")
       WTipoNota = ADOCancela("TipoNota")
    Else
        MsgBox "NF/Cupom não encontrado ou já cancelado", vbInformation, "Aviso"
        txtSerie.Text = ""
        txtNotaFiscal.SelStart = 0
        txtNotaFiscal.SelLength = Len(txtNotaFiscal.Text)
        txtNotaFiscal.SetFocus
End If
ADOCancela.Close

End Sub



Sub LimpaCampos()

        txtSerie.Text = ""
        txtValorNF.Text = ""
        txtSenha.Text = ""
        txtNotaFiscal.Text = ""
        txtPedido.Text = ""
        Exit Sub

End Sub

'Private Sub CancelaCupomFiscal()
    
    
'    Retorno = Bematech_FI_CancelaCupom()
'    Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
'   If Retorno = 1 Then
'        Call AtualizaNumeroCupom
'    End If

'End Sub





Private Sub txtValorNF_Change()

End Sub
