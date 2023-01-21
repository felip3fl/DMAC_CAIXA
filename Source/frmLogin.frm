VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Senha"
   ClientHeight    =   1350
   ClientLeft      =   4860
   ClientTop       =   4080
   ClientWidth     =   2505
   LinkTopic       =   "Form1"
   ScaleHeight     =   1350
   ScaleWidth      =   2505
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   795
      TabIndex        =   1
      Top             =   825
      Width           =   795
   End
   Begin VB.CommandButton cmdRetornar 
      Caption         =   "&Retornar"
      Height          =   330
      Left            =   1590
      TabIndex        =   2
      Top             =   825
      Width           =   765
   End
   Begin VB.TextBox txtSenha 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   810
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   1515
   End
   Begin VB.Label lblSenha 
      AutoSize        =   -1  'True
      Caption         =   "Senha :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   165
      TabIndex        =   3
      Top             =   390
      Width           =   555
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conexaosql1 As New rdoConnection
Dim R As rdoResultset
Dim ws As String
Dim wu As String

Function MontaSenhaDia() As String
    'A montagem da senha é = ((HH+Loja)*100/Loja)*36
    Dim rsLoja As rdoResultset
    Dim wLoja As Double
    Dim wSenhaMontada As Double
    Dim wDataSenha As String
    Dim wHora As Double
    Dim wDia As Double
    
    SQL = ""
    SQL = "Select CT_Loja from Controle"
        Set rsLoja = rdoCnLoja.OpenResultset(SQL)
    If Not rsLoja.EOF Then
        wDia = Val(Format(Date, "DD"))
        wHora = Val(Format(Time, "hh"))
        wLoja = Val(rsLoja("CT_Loja"))
        If wLoja = 0 Then
            wLoja = 1
        End If
        wSenhaMontada = ((wDia + wHora + wLoja) * 100 / wLoja) * 36
        MontaSenhaDia = Val(wSenhaMontada)
    End If
        

End Function



        
Private Sub cmdOK_Click()
    
'    If Trim(txtSenha.Text) = Trim(MontaSenhaDia) Then
    '
    '---------------------------Destrava o sistema se a senha estiver correta------------
    '
        If frmLogin.Caption = "Senha" Then
            SQL = ""
            SQL = "Update ControleECF set CT_SituacaoCaixa='F' where CT_ECF=" & Val(glb_ECF)
                rdoCnLoja.Execute (SQL)
            
            MsgBox "O sistema foi destravado com sucesso", vbInformation, "Sucesso"
        ElseIf frmLogin.Caption = "OFF" Then
            SQL = ""
            SQL = "Update Controle set CT_BancosOnLine = 'N' "
            rdoCnLoja.Execute (SQL)
            GLB_BancosOnline = "N"
            mdiBalcao.MnuLojaOffLine.Checked = True
            mdiBalcao.MnuLojaOnLine.Checked = False
            MsgBox "O sistema foi alterado para o modo Off-Line", vbInformation, "Atenção"
        End If
        Unload Me
'    Else
'        MsgBox "Senha incorreta", vbCritical, "Aviso"
'        txtSenha.Text = ""
'        txtSenha.SetFocus
'    End If
    
End Sub

Private Sub cmdRetornar_Click()
    
    Unload Me
    
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If frmLogin.Caption = "Senha" Then
        DescarregaForms
    End If
    
End Sub
