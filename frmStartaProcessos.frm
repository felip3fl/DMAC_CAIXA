VERSION 5.00
Begin VB.Form frmStartaProcessos 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Starta Processos"
   ClientHeight    =   5925
   ClientLeft      =   255
   ClientTop       =   3630
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5925
   ScaleMode       =   0  'User
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPedido 
      Height          =   285
      Left            =   255
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   300
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Timer Timer1 
      Left            =   420
      Top             =   3780
   End
End
Attribute VB_Name = "frmStartaProcessos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
 
 Call AjustaTela(frmStartaProcessos)
 
      Screen.MousePointer = 11
      
  frmControlaCaixa.webInternet1.Picture = LoadPicture(endIMG("topo1024768hd"))
  frmStartaProcessos.Picture = LoadPicture(endIMG("FundoProcessa"))
 Call StatusAtualizacao
 
    wPedido = pedido
    
    sql = "Select top 1 serie from nfcapa where numeroped = " & pedido
    rsComplementoVenda.CursorLocation = adUseClient
    rsComplementoVenda.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If rsComplementoVenda.EOF Then
        Esperar 1
    ElseIf rsComplementoVenda("serie") = "NE" Then
        Call CriaNFE(NroNotaFiscal, pedido)
    ElseIf rsComplementoVenda("serie") = "CE" Then
        Call CriaSAT(NroNotaFiscal, pedido)
    Else
        Esperar 1
    End If
    
    rsComplementoVenda.Close

 Screen.MousePointer = 0
 Unload Me
End Sub

Private Sub StatusAtualizacao()

   sql = "exec sp_totaliza_capa_nota_fiscal_Loja " & pedido
         RsDadosTef.CursorLocation = adUseClient
         RsDadosTef.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
  
   sql = "exec SP_Atualiza_Processos_Venda " & pedido & ", " & NroNotaFiscal & ", " & GLB_CTR_Protocolo & ", " & GLB_Caixa
          rdoCNLoja.Execute sql

   sql = "exec SP_Atualiza_Processos_Venda_Central"
          rdoCNLoja.Execute sql

   'ConectaODBCMatriz
    
   'If wConectouRetaguarda = True Then
        'sql = "Exec SP_Est_Transferencia_destino '" & RTrim(LTrim(GLB_Loja)) & "'"
        'rdoCNRetaguarda.Execute sql
   'End If

End Sub


Sub Esperar(ByVal tempo As Integer)
    Dim StartTime As Long
    StartTime = Timer
    Do While Timer < StartTime + tempo
       DoEvents
    Loop
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Imagemfundo_Click()

End Sub

