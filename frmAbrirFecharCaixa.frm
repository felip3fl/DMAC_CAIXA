VERSION 5.00
Begin VB.Form frmAbrirFecharCaixa 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1455
   ClientLeft      =   4080
   ClientTop       =   3675
   ClientWidth     =   4425
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtSenha 
      BackColor       =   &H80000004&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1620
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   765
      Width           =   1380
   End
   Begin VB.TextBox txtUsuario 
      BackColor       =   &H80000004&
      Height          =   315
      Left            =   1620
      TabIndex        =   0
      Top             =   360
      Width           =   1380
   End
   Begin VB.CommandButton cmdCaixaAberto 
      Height          =   765
      Left            =   3180
      Picture         =   "frmAbrirFecharCaixa.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   345
      Width           =   975
   End
   Begin VB.CommandButton cmdCaixaFechado 
      Height          =   765
      Left            =   3180
      Picture         =   "frmAbrirFecharCaixa.frx":10C2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   345
      Width           =   975
   End
   Begin VB.Label lblSenha 
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   360
      TabIndex        =   5
      Top             =   855
      Width           =   1065
   End
   Begin VB.Label lblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   360
      TabIndex        =   4
      Top             =   435
      Width           =   1065
   End
End
Attribute VB_Name = "frmAbrirFecharCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Data As Date
Dim NumeroCaixa As Long
Dim Sequencia As Long
Dim Operador As Long
Dim Operacoes As Long
Dim Controle As Long


Dim PegaLoja As rdoResultset
Dim VerificaUsuario As rdoResultset
Dim PegaDadosAnteriores As rdoResultset
Dim pegadados As rdoResultset
Dim RSLojaControle As rdoResultset

Dim Loja As String
Dim HoraInicial As String
Dim HoraFinal As String
Dim SituacaoCaixa As String
Dim wSituacao As String
Dim WusuarioOk As Boolean


Private Sub cmdCaixaAberto_Click()
   wPegaLojaControle = ""
   SQL = ""
   SQL = "Select CT_Loja from controle"
    Set RSLojaControle = rdoCNLoja.OpenResultset(SQL)
   
   If Not RSLojaControle.EOF Then
        wPegaLojaControle = RSLojaControle("CT_Loja")
   Else
        MsgBox "Loja não encontrada no Controle", vbInformation, "Atenção"
        Exit Sub
   End If
   
   
   wPegaUsuario = ""
   Set VerificaUsuario = rdoCNLoja.OpenResultset("Select * from Usuario where US_Usuario='" & txtusuario.Text & "' and US_Senha='" & txtSenha.Text & "'")
   GLB_USU_Nome = VerificaUsuario("US_usuario")
   If Not VerificaUsuario.EOF Then
      Abre = True
      Fecha = False
      
      Operador = VerificaUsuario("US_Codigo")
      cmdCaixaAberto.Visible = False
      cmdCaixaFechado.Visible = True
      cmdCaixaAberto.Refresh
     ' Esperar 1
      frmAbrirFecharCaixa.Visible = False
      
      VerificaProcesso
      LimpaText
      frmRotinasDiaria.Visible = False
      'frmCaixa.Hide
      AtualizaProcessoFechamento "Controle", "CT_SeqFechamento", "EF"
      frmReforcoSangriaCaixa.Show
   Else
      MsgBox "Usuário ou senha incorretos, verifique.", vbInformation, "Informação"
      txtusuario.SetFocus
      Exit Sub
   End If
   
   VerificaUsuario.Close
   
End Sub

Private Sub cmdCaixaFechado_Click()
   wPegaLojaControle = ""
   SQL = ""
   SQL = "Select CT_Loja from controle"
    Set RSLojaControle = rdoCNLoja.OpenResultset(SQL)
   
   If Not RSLojaControle.EOF Then
        wPegaLojaControle = RSLojaControle("CT_Loja")
   Else
        MsgBox "Loja não encontrada no Controle", vbInformation, "Atenção"
        Exit Sub
   End If
   
   Set VerificaUsuario = rdoCNLoja.OpenResultset("Select * from Usuario where Us_Usuario='" & txtusuario.Text & "' and Us_Senha='" & txtSenha.Text & "'")
   
   If Not VerificaUsuario.EOF Then
      Fecha = True
      Abre = False
      Operador = VerificaUsuario("Us_Codigo")
      cmdCaixaAberto.Visible = True
      cmdCaixaFechado.Visible = False
      cmdCaixaAberto.Refresh
    '  Esperar 1
      frmAbrirFecharCaixa.Visible = False
      VerificaProcesso
      LimpaText
      'frmReforcoSangriaCaixa.Show
   Else
      MsgBox "Usuário ou senha incorretos, verifique.", vbInformation, "Informação"
      txtusuario.SetFocus
      Exit Sub
   End If
   On Error Resume Next
   VerificaUsuario.Close
End Sub

Private Sub Form_Load()

   Left = (Screen.Width - Width) / 2
   Top = (Screen.Height - Height) / 3
   Call ValidaAbertura

   Set PegaLoja = rdoCNLoja.OpenResultset("Select * from Controle")
   
   If Not PegaLoja.EOF Then
      NumeroCaixa = 1
      Loja = PegaLoja("CT_Loja")
      Data = Date
      Sequencia = 0
      HoraInicial = Time
   End If
   
   Abre = False
   Fecha = False
   
   If FecParcial = True Or FecTotal = True Then
      cmdCaixaAberto.Visible = False
      cmdCaixaFechado.Visible = True
   Else
      cmdCaixaAberto.Visible = True
      cmdCaixaFechado.Visible = False
   End If
   
   
   
 End Sub

Sub LimpaText()

    txtusuario.Text = ""
    txtSenha.Text = ""

End Sub
 
Sub VerificaProcesso()

    Set pegadados = rdoCNLoja.OpenResultset("Select Max(CT_Data)as DataMov,Max(Ct_Sequencia) as Seq from CTCaixa where CT_NumeroECF = " & glb_ECF & "")

    If (Not pegadados.EOF) And (pegadados("Seq") > 0) Then
    
       Set PegaDadosAnteriores = rdoCNLoja.OpenResultset("Select * from CtCaixa where ct_Data='" & Format(pegadados("datamov"), "mm/dd/yyyy") & "' and Ct_Sequencia= " & VerificaCaixa("seq") & " order by ct_controle desc")
       wSituacao = PegaDadosAnteriores("CT_Situacao")
       Call VerificaUsuarioEmUso
       
       If WusuarioOk = False Then
          Exit Sub
       End If
          
       If Not PegaDadosAnteriores.EOF Then
          If Abre = True Then

             rdoCNLoja.Execute "Insert into CTCaixa (CT_NumeroECF, CT_Loja, CT_Data, CT_Operador, " _
                          & "CT_HoraInicial, CT_HoraFinal, CT_Situacao, CT_Operacoes, CT_Controle) values (" & glb_ECF & ", " _
                          & "'" & PegaDadosAnteriores("CT_Loja") & "', '" & Format(Date, "mm/dd/yyyy") & "', " _
                          & " " & PegaDadosAnteriores("CT_Operador") & ", " _
                          & "'" & Format(PegaDadosAnteriores("CT_HoraInicial"), "hh:mm") & "', " _
                          & "null,'A', " _
                          & " 0," & PegaDadosAnteriores("CT_Controle") & " )"

             If PegaDadosAnteriores("CT_Situacao") = "T" Then
                'rdoCnLoja.Execute "Update CTCaixa set CT_Data= '" & Format(Date, "mm/dd/yyyy") & "', CT_Operador=" & Operador & ", CT_HoraInicial='" & Format(Time, "hh:mm") & "', " _
                             & "CT_HoraFinal=(null), CT_Situacao='A', CT_Operacoes=0, CT_Controle=1 where CT_sequencia=" & VerificaCaixa("seq") + 1 & " and CT_situacao='T' "
                
                WAbrirCaixa = True
                
                SQL = "Update controle set CT_TipoArquivo= 1"
                rdoCNLoja.Execute (SQL)
                
                If WAbrirCaixa = True Then
                    frmRotinasDiaria.cmdProcessar.Enabled = False
                    Call ProcessaRotinasDiarias
                Else
                    frmRotinasDiaria.cmdProcessar.Enabled = True
                End If
                
                WAbrirCaixa = False
                
             ElseIf PegaDadosAnteriores("CT_Situacao") = "P" Then
                Controle = PegaDadosAnteriores("Ct_Controle") + 1

                rdoCNLoja.Execute "Update CTCaixa set Ct_Data= '" & Format(Date, "mm/dd/yyyy") & "', Ct_Operador=" & Operador & ",Ct_HoraInicial='" & Format(Time, "hh:mm") & "', " _
                             & "Ct_HoraFinal=(null), Ct_Situacao='A', Ct_Operacoes=0, Ct_Controle=" & Controle & " where Ct_sequencia=" & VerificaCaixa("seq") + 1 & " and Ct_situacao='P' "
             End If
          
          
          ElseIf Fecha = True Then
             
'             If SituacaoFechaRetaguarda = True Then
                SQL = ""
                SQL = "Update FechamentoRetaguarda Set FR_DataFechamento = '" & Format(pegadados("DataMov"), "mm/dd/yyyy") & "', " _
                    & "FR_SituacaoFechamento = 'A'"
                rdoCnLojaBach.Execute (SQL)
'             Else
'                MsgBox "Não é possível fazer o fechamento do caixa," & Chr(10) & Chr(13) & "pois você não fez o fechamento da Retaguarda", vbCritical
'                frmAbrirFecharCaixa.Hide
'                SitFechaReta = False
'                frmFechamentoRetaguarda.Show 1
'             End If
             If FecParcial = True Then
                rdoCNLoja.Execute "Update CTCaixa set Ct_HoraFinal='" & Format(Time, "hh:mm") & "', Ct_Situacao='P' where Ct_sequencia=" & VerificaCaixa("seq") & " "
                frmAbrirFecharCaixa.Hide
                'frmCaixa.Hide
             ElseIf FecTotal = True Then
                rdoCNLoja.Execute "Update CTCaixa set Ct_HoraFinal='" & Format(Time, "hh:mm") & "', Ct_Situacao='T' where Ct_sequencia=" & VerificaCaixa("seq") & " "
                
                frmAbrirFecharCaixa.Hide
                'frmCaixa.Hide
                WAbrirCaixa = True
                
                SQL = "Update controle set CT_TipoArquivo= 2"
                rdoCNLoja.Execute (SQL)
                
                If WAbrirCaixa = True Then
                    frmRotinasDiaria.cmdProcessar.Enabled = False
                    Call ProcessaRotinasDiarias
                Else
                    frmRotinasDiaria.cmdProcessar.Enabled = True
                End If
                
                WAbrirCaixa = False
             End If
          
          End If
       End If
    Else
        rdoCNLoja.Execute "Insert into CTCaixa (CT_NumeroECF, CT_Loja, CT_Data, CT_Operador, " _
                     & "CT_HoraInicial, CT_HoraFinal, CT_Situacao, CT_Operacoes, CT_Controle) values (" & glb_ECF & ", " _
                     & "'" & PegaDadosAnteriores("CT_Loja") & "', '" & Format(Date, "mm/dd/yyyy") & "', " _
                     & " " & PegaDadosAnteriores("CT_Operador") & ", " _
                     & "'" & Format(PegaDadosAnteriores("CT_HoraInicial"), "hh:mm") & "', " _
                     & "null,'A', " _
                     & " 0," & PegaDadosAnteriores("CT_Controle") & " )"
        
        WAbrirCaixa = True
        
        SQL = "Update controle set CT_TipoArquivo= 1"
        rdoCNLoja.Execute (SQL)
        
        If WAbrirCaixa = True Then
            frmRotinasDiaria.cmdProcessar.Enabled = False
            Call ProcessaRotinasDiarias
        Else
            frmRotinasDiaria.cmdProcessar.Enabled = True
        End If
        
        WAbrirCaixa = False
    End If
End Sub




Private Sub VerificaUsuarioEmUso()
      
       
       
       WusuarioOk = True
       SQL = ""
       SQL = "Select Us_Codigo from Usuario where Us_Usuario= '" & Trim(txtusuario.Text) & "'"
       Set RsUsuario = rdoCNLoja.OpenResultset(SQL)
       
       If Not RsUsuario.EOF Then
          If RsUsuario("Us_Codigo") <> PegaDadosAnteriores("Ct_Operador") Then
             If wSituacao = "A" Then
                MsgBox "Somente o usuario que abriu pode fechar o caixa", vbCritical, "Atenção"
                WusuarioOk = False
                Exit Sub
            End If
          End If
       Else
          MsgBox "Usuario não cadastrado", vbCritical, "Atenção"
          WusuarioOk = False
          Exit Sub
       End If

End Sub

Private Sub txtusuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtusuario.Text = "" Then
            txtusuario.SetFocus
            Exit Sub
        End If
        txtSenha.SetFocus
    End If
End Sub
