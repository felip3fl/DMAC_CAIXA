VERSION 5.00
Begin VB.Form frmInicio 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conectar"
   ClientHeight    =   1980
   ClientLeft      =   3360
   ClientTop       =   4620
   ClientWidth     =   3030
   Icon            =   "frmInicio.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmb_Caixa 
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      Top             =   765
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sair"
      Height          =   375
      Left            =   1530
      TabIndex        =   2
      Top             =   1305
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   285
      TabIndex        =   1
      Top             =   1305
      Width           =   1215
   End
   Begin VB.ComboBox cmb_loja 
      Height          =   315
      Left            =   1515
      TabIndex        =   0
      Top             =   330
      Width           =   1035
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Caixa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   285
      TabIndex        =   5
      Top             =   765
      Width           =   1080
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Loja"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   285
      TabIndex        =   4
      Top             =   330
      Width           =   1080
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmb_loja_Click()

    Dim i As Integer

    Sql = "Select CXA_NumeroCaixa from ParametroSistema order by CXA_NumeroCaixa"
    
    rdoParametroINI.CursorLocation = adUseClient
    rdoParametroINI.Open Sql, adoCNAccess, adOpenForwardOnly, adLockPessimistic
    
        If Not rdoParametroINI.EOF Then
            cmb_Caixa.Clear
            
            For i = 1 To Trim(rdoParametroINI("CXA_NumeroCaixa"))
                cmb_Caixa.AddItem i
            Next i

            Screen.MousePointer = 0
            cmb_Caixa.ListIndex = 0
        End If
    
    rdoParametroINI.Close

End Sub

Private Sub cmb_loja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Command1_Click
    End If

'txtCondPagamento.Text = Mid(cmbCondPag.Text, 1, 1)
End Sub

Private Sub Command1_Click()
 
    If adoCNAccess.State = 1 Then
         adoCNAccess.Close
    End If
     
    lsDSN = "Driver={Microsoft Access Driver (*.mdb)};" & _
              "Dbq=c:\sistemas\DMACini.mdb;" & _
              "Uid=Admin; Pwd=astap36"
      adoCNAccess.Open lsDSN
     
    Sql = "Select * from ConexaoSistema where GLB_Loja = '" & Trim(cmb_loja.text) & "'"
     
    rdoConexaoINI.CursorLocation = adUseClient
    rdoConexaoINI.Open Sql, adoCNAccess, adOpenForwardOnly, adLockPessimistic
     
      If Not rdoConexaoINI.EOF Then
           GLB_Servidor = Trim(rdoConexaoINI("GLB_ServidorRetaguarda"))
           GLB_Loja = Trim(rdoConexaoINI("GLB_Loja"))
           GLB_Banco = Trim(rdoConexaoINI("GLB_BancoRetaguarda"))
           GLB_Servidorlocal = Trim(rdoConexaoINI("GLB_ServidorLocal"))
           Glb_BancoLocal = Trim(rdoConexaoINI("GLB_BancoLocal"))
           'GLB_EnderecoPortal = Trim(rdoConexaoINI("GLB_Portal"))
           GLB_EnderecoPastaRESP = Trim(rdoConexaoINI("GLB_EnderecoResp"))
           GLB_EnderecoPastaFIL = Trim(rdoConexaoINI("GLB_EnderecoFil"))
      End If
      
      rdoConexaoINI.Close
      ConectaODBC
      adoCNAccess.Close
      
      GLB_Caixa = Val(cmb_Caixa.text)
      
Continua:
    
        If GLB_ConectouOK = True Then
           Call DadosLoja
           
           Sql = "Select * from ControleCaixa Where CTR_Supervisor = 99 and CTR_SituacaoCaixa='F' " _
               & "and CTR_datainicial >= '" & Format(Date, "yyyy/mm/dd") & "'"
                RsDados.CursorLocation = adUseClient
                
                RsDados.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                If Not RsDados.EOF Then
                   MsgBox "Fechamento Geral de hoje já foi efetuado. Não é possivel abrir o Caixa."
                   wPermitirVenda = False
                   RsDados.Close
                   Unload Me
                   Exit Sub
                End If
                RsDados.Close
           
                If carregaControleCaixa = False Then
                   frmFundoEscuro.Show
                   frmLoginCaixa.Show
                Else
                   frmBandeja.Show
    
                End If
        Else
            MsgBox "Erro ao conectar-se ao Banco de Dados", vbCritical, "Atenção"
            Exit Sub
        End If
        
        Unload Me

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
'Skin1.LoadSkin "c:\WINDOWS\system\skin.skn"
'Skin1.ApplySkin Me.Hwnd

Sql = "Select GLB_LOJA from ConexaoSistema GROUP BY GLB_LOJA"
 
rdoConexaoINI.CursorLocation = adUseClient
rdoConexaoINI.Open Sql, adoCNAccess, adOpenForwardOnly, adLockPessimistic
 
        If Not rdoConexaoINI.EOF Then
            cmb_loja.Clear
            Do While Not rdoConexaoINI.EOF
                cmb_loja.AddItem Trim(rdoConexaoINI("GLB_LOJA"))
                rdoConexaoINI.MoveNext
            Loop
            Screen.MousePointer = 0
        End If

rdoConexaoINI.Close
cmb_loja.ListIndex = 0

Exit Sub
ConexaoErro:
MsgBox "Erro ao abrir banco de Dados da Loja! "
End

End Sub


'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'adoBDIBC2000.Close
'adoCNLoja.Close
'End Sub
