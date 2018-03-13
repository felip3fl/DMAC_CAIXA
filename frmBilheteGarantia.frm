VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmBilheteGarantia 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Bilhete Garantia Estendida"
   ClientHeight    =   7605
   ClientLeft      =   240
   ClientTop       =   1905
   ClientWidth     =   19830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   19830
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser WebNavegadorTermo 
      Height          =   1620
      Left            =   16245
      TabIndex        =   6
      Top             =   3585
      Width           =   3195
      ExtentX         =   5636
      ExtentY         =   2857
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer timerImprimir 
      Left            =   15
      Top             =   0
   End
   Begin VB.PictureBox imgOcultaNavegador 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   1215
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraPedido 
      BackColor       =   &H00000000&
      Caption         =   "Pedido G.E."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   885
      Left            =   285
      TabIndex        =   1
      Top             =   6510
      Width           =   1905
      Begin VB.TextBox txtNumeroPedido 
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
         Left            =   120
         MaxLength       =   6
         TabIndex        =   0
         Top             =   285
         Width           =   1680
      End
   End
   Begin Balcao2010.chameleonButton cmdImprimirTodos 
      Height          =   600
      Left            =   11025
      TabIndex        =   2
      Top             =   6630
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1058
      BTYPE           =   11
      TX              =   "Imprimir"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   14737632
      FCOLO           =   14737632
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBilheteGarantia.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SHDocVwCtl.WebBrowser webNavegador 
      Height          =   5985
      Left            =   300
      TabIndex        =   5
      Top             =   300
      Width           =   12675
      ExtentX         =   22357
      ExtentY         =   10557
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label lblStatusImpressao 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Status Impressao"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   0
      TabIndex        =   3
      Top             =   6645
      Visible         =   0   'False
      Width           =   15210
   End
End
Attribute VB_Name = "frmBilheteGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pedido As String

Dim rsCapaBilhete As New ADODB.Recordset
Dim rsItemBilhete As New ADODB.Recordset

Dim rsItensGarantia As New ADODB.Connection

Dim codigoBilheteHTML As String
Dim codigoBilheteDMACHTML As String
Dim codigoTermoHTML As String

Public Sub statusFuncionamento()
    Dim mensagem As String
    
    mensagem = "Imprimindo Garantia Estendida" & " "
    If lblStatusImpressao.Caption = mensagem & "  . . . ." Then
        lblStatusImpressao.Caption = mensagem & ".   . . ."
    ElseIf lblStatusImpressao.Caption = mensagem & ".   . . ." Then
        lblStatusImpressao.Caption = mensagem & ". .   . ."
    ElseIf lblStatusImpressao.Caption = mensagem & ". .   . ." Then
        lblStatusImpressao.Caption = mensagem & ". . .   ."
    ElseIf lblStatusImpressao.Caption = mensagem & ". . .   ." Then
        lblStatusImpressao.Caption = mensagem & ". . . .  "
    ElseIf lblStatusImpressao.Caption = mensagem & ". . . .  " Then
        lblStatusImpressao.Caption = mensagem & "  . . . ."
    Else
        lblStatusImpressao.Caption = mensagem & "  . . . ."
    End If
End Sub

Private Sub Form_Activate()
    If frmFormaPagamento.txtPedido.text <> "" Then
        fraPedido.Visible = False
    End If
End Sub

Private Sub Form_GotFocus()
    txtNumeroPedido.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        encerrar
    End If
End Sub

Private Sub ocultaNavegador(ativa As Boolean)
    imgOcultaNavegador.left = webNavegador.left
    imgOcultaNavegador.top = webNavegador.top
    imgOcultaNavegador.Width = webNavegador.Width
    imgOcultaNavegador.Height = webNavegador.Height
    imgOcultaNavegador.Visible = ativa
End Sub

Private Sub cmdEncerrar_Click()
    rsItensGarantia.Close
    
    txtNumeroPedido.text = Empty
    lblStatusImpressao.Visible = False
    cmdImprimirTodos.Visible = True
    fraPedido.Visible = True
    
    Unload Me
End Sub

Private Sub cmdImprimirTodos_Click()
    Screen.MousePointer = 11

    'lblStatusImpressao.Caption = mensagemImpressao & " " & statusImpressao
    statusFuncionamento
    lblStatusImpressao.Visible = True
    cmdImprimirTodos.Visible = False
    fraPedido.Visible = False
    
    ocultaNavegador True

    criaBilhete True
    
    timerImprimir.Enabled = True
    timerImprimir.Interval = 800
    
End Sub

Public Sub verificaNumeroPedido()

    If frmFormaPagamento.txtPedido.text <> "" Then
    
        frmBilheteGarantia.Visible = False
        
        txtNumeroPedido.text = frmFormaPagamento.txtPedido.text
        txtNumeroPedido_KeyPress (13)
        
        fraPedido.Visible = False
        cmdImprimirTodos_Click
        
    Else
        'frmBilheteGarantia.Visible = True
        'txtNumeroPedido.Enabled = True
        'txtNumeroPedido.SetFocus
        
        'frmBilheteGarantia.Show 1'
        
    End If
    
End Sub

Private Function carregaListaItem(numeroPedido As String) As Boolean

    Dim Sql As String
    
    carregaListaItem = True
    
    Sql = "SP_FIN_GE_Monta_Campos_capa " & numeroPedido
    
    rsCapaBilhete.CursorLocation = adUseClient
    rsCapaBilhete.Open Sql, rdoCNLoja, adOpenDynamic, adLockPessimistic
    
    'Set rsItensGarantia = rdoCNLoja.OpenResultset(SQL, rdOpenKeyset)
    
    If rsCapaBilhete.EOF Then
        carregaListaItem = False
        rsCapaBilhete.Close
    End If
    
End Function

Private Sub atualizaNumeroCertificado(numeroPedido As String, ByRef rsItens)
    Dim Sql As String
    
    Do While Not rsItens.EOF
        Sql = Sql & "update nfitens set " & vbNewLine & _
              "certificadoInicio = (select cts_certificado from ControleSistema)+1, " & vbNewLine & _
              "certificadoFim = ((select cts_certificado from ControleSistema)+qtdeGarantia) " & vbNewLine & _
              "where numeroPed = " & numeroPedido & " and garantiaEstendida = 'S' and " & vbNewLine & _
              "referencia = " & rsItens("Referencia") & vbNewLine & vbNewLine
              
        Sql = Sql & "update ControleSistema set " & vbNewLine & _
              "cts_certificado = (select certificadoFIM from nfitens " & vbNewLine & _
              "where numeroPed = " & numeroPedido & " and garantiaEstendida = 'S' and " & vbNewLine & _
              "referencia = " & rsItens("Referencia") & ")" & vbNewLine & vbNewLine
        rsItens.MoveNext
    Loop

    rdoCNLoja.Execute (Sql)
    'rsItensGarantia.Close
    ''carregaListaItem numeroPedido '''''''''''''''''''''''''''''''''''''''''''''''

End Sub

Private Function possuiNumeroCertificado(numeroPedido As String) As Boolean
    Dim Sql As String
    Dim rsProdutoGarantia As New ADODB.Recordset
    
    possuiNumeroCertificado = True

    Sql = "select certificadoInicio from nfitens " & _
          "where numeroPed = " & numeroPedido & " and " & _
          "referencia = " & rsItemBilhete("Referencia")
          
    Set rsProdutoGarantia = rdoCNLoja.OpenResultset(Sql)
        If IsNull(rsProdutoGarantia("certificadoInicio")) Then
            possuiNumeroCertificado = False
        End If
    rsProdutoGarantia.Close
End Function

Private Sub timerImprimir_Timer()
    
    statusFuncionamento
    If Not rsItemBilhete.EOF Then
        webNavegador.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
        webNavegador.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
        WebNavegadorTermo.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
        WebNavegadorTermo.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
        
        rsItemBilhete.MoveNext
        If Not rsItemBilhete.EOF Then
            criaBilhete True
            criaTermo
        End If
    Else
        Screen.MousePointer = 0
        encerrar
    End If

End Sub

Private Sub txtNumeroPedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        encerrar
    End If
        
    If KeyAscii = 13 Then
        Screen.MousePointer = 11
        
        If IsNumeric(txtNumeroPedido.text) Then
           
            pedido = txtNumeroPedido.text
           
            possuiNumCertificado (pedido)
          
            
            encerraConexao
           
            If carregaListaItem(pedido) Then
                
                carregaItem pedido
              
                
                criaBilhete False
                
                criaTermo
                
      
                ocultaNavegador False
                cmdImprimirTodos.Visible = True
                

                
            Else
                MsgBox "Não há contrato de Garantia Estendida para o pedido " & pedido, vbExclamation, Me.Caption
                txtNumeroPedido.text = Empty
            End If
        Else
            txtNumeroPedido.text = Empty
        End If
        
        Screen.MousePointer = 0

        
    End If
End Sub

Private Sub possuiNumCertificado(numeroPedido)

    Dim Sql As String
    Dim rsItens As New ADODB.Recordset
    
    Sql = "select itens.referencia as referencia " & vbNewLine & _
          "from nfitens itens, nfcapa capa " & vbNewLine & _
          "where itens.garantiaEstendida = 'S' " & vbNewLine & _
          "and itens.certificadoInicio = '' " & vbNewLine & _
          "and capa.numeroPed = itens.numeroPed " & vbNewLine & _
          "and capa.tipoNota = 'V'" & vbNewLine & _
          "and capa.garantiaEstendida = 'S' " & vbNewLine & _
          "and capa.numeroPed = " & numeroPedido & "" & vbNewLine & _
          "order by itens.item"

    rsItens.CursorLocation = adUseClient
    rsItens.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

        If Not rsItens.EOF Then
            atualizaNumeroCertificado pedido, rsItens
        End If
        
    rsItens.Close
    
End Sub

''''' CAMPOS ITENS ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub carregaItem(numeroPedido As String)

    Dim Sql As String
    
    rdoCNLoja.Execute ("exec SP_FIN_GE_Monta_Campos_item " & numeroPedido)

    Sql = "SELECT * FROM temp_GE_Itens order by certificadoInicio"
    
    rsItemBilhete.CursorLocation = adUseClient
    rsItemBilhete.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If rsItemBilhete.EOF Then
        MsgBox "Erro ao monta item GE", vbCritical, "GE"
    End If
    
End Sub

Private Sub criaBilhete(corParaImpressao As Boolean)

    Dim BilheteHTML As String
    Dim campoCapa As String
    Dim campoItem As String
    Dim campo As String
    
    If corParaImpressao = True Then
        BilheteHTML = codigoBilheteHTML
    Else
        BilheteHTML = codigoBilheteDMACHTML
    End If
    
    campoCapa = "DMACCapa"
    campoItem = "DMACItem"
    
    campo = campoBilhete(BilheteHTML, campoCapa)
    Do While campo <> ""
        BilheteHTML = Replace(BilheteHTML, campoCapa & campo, rsCapaBilhete(campo))
        campo = campoBilhete(BilheteHTML, campoCapa)
    Loop
    
    campo = campoBilhete(BilheteHTML, campoItem)
    Do While campo <> ""
        BilheteHTML = Replace(BilheteHTML, campoItem & campo, rsItemBilhete(campo))
        campo = campoBilhete(BilheteHTML, campoItem)
    Loop

    Open "C:\Sistemas\DMAC Caixa\GarantiaEstendida\Bilhete.htm" For Output As #1
         Print #1, BilheteHTML
    Close #1

    webNavegador.Navigate "C:\Sistemas\DMAC Caixa\GarantiaEstendida\Bilhete.htm"

End Sub

Private Function Replace(Source As String, Find As String, ReplaceStr As String, _
    Optional ByVal Start As Long = 1, Optional Count As Long = -1, _
    Optional Compare As VbCompareMethod = vbBinaryCompare) As String

    Dim findLen As Long
    Dim replaceLen As Long
    Dim Index As Long
    Dim counter As Long
    
    findLen = Len(Find)
    replaceLen = Len(ReplaceStr)
    If findLen = 0 Then Err.Raise 5
    
    If Start < 1 Then Start = 1
    Index = Start
    
    Replace = Source
    
    Do
        Index = InStr(Index, Replace, Find, Compare)
        If Index = 0 Then Exit Do
        If findLen = replaceLen Then
            Mid$(Replace, Index, findLen) = ReplaceStr
        Else
            Replace = left$(Replace, Index - 1) & ReplaceStr & Mid$(Replace, _
                Index + findLen)
        End If
        Index = Index + replaceLen
        counter = counter + 1
    Loop Until counter = Count
    
    If Start > 1 Then Replace = Mid$(Replace, Start)

End Function


Private Function campoBilhete(codigoBilhete As String, campo As String) As String
    
    If codigoBilhete Like "*" & campo & "*" Then
        Dim inicioCampo, fimCampo As Integer
    
        inicioCampo = (InStr(codigoBilhete, campo)) + (Len(campo))
        fimCampo = (InStr(inicioCampo, codigoBilhete, "<")) - inicioCampo
    
        If inicioCampo + fimCampo <> 0 Then
            campoBilhete = Mid$(codigoBilhete, inicioCampo, fimCampo)
        End If
        
    Else
        campoBilhete = ""
    End If
    
End Function

Private Sub webNavegador_GotFocus()
    frmBilheteGarantia.SetFocus
End Sub

Private Sub Form_Load()

    'mensagemImpressao = "         Imprimindo Garantia Estendida"
    'statusImpressao = ". . . . ."
    
   ' txtNumeroPedido.SetFocus
    
    limpaVariavel
    Call AjustaTela(Me)
    lblStatusImpressao.Height = Me.Height
    ocultaNavegador True
    obterCertificado
    verificaNumeroPedido
    
End Sub

Private Sub obterCertificado()
    Dim fso As New FileSystemObject
    Dim mensagemArquivoTXT As TextStream

    Set mensagemArquivoTXT = fso.OpenTextFile _
    ("C:\Sistemas\DMAC Caixa\GarantiaEstendida\configBilhete")
    codigoBilheteHTML = mensagemArquivoTXT.ReadAll
    mensagemArquivoTXT.Close
    
    Set mensagemArquivoTXT = fso.OpenTextFile _
    ("C:\Sistemas\DMAC Caixa\GarantiaEstendida\configBilheteDMAC")
    codigoBilheteDMACHTML = mensagemArquivoTXT.ReadAll
    mensagemArquivoTXT.Close
    
    Set mensagemArquivoTXT = fso.OpenTextFile _
    ("C:\Sistemas\DMAC Caixa\GarantiaEstendida\configTermo")
    codigoTermoHTML = mensagemArquivoTXT.ReadAll
    mensagemArquivoTXT.Close
End Sub

Private Sub encerrar()
    Dim Sql As String
    
    Sql = "IF object_id('temp_GE_Itens') IS NOT NULL" & vbNewLine & _
          "BEGIN" & vbNewLine & _
          "   drop table temp_GE_Itens" & vbNewLine & _
          "END"
    
    rdoCNLoja.Execute (Sql)
    encerraConexao
    Screen.MousePointer = 0
    
    Unload Me
    
End Sub



Private Sub encerraConexao()
    If rsCapaBilhete.State = 1 Then rsCapaBilhete.Close
    If rsItemBilhete.State = 1 Then rsItemBilhete.Close
End Sub

Private Sub limpaVariavel()
    pedido = Empty
    codigoBilheteHTML = Empty
End Sub

Private Sub criaTermo()
    
    Dim TermoHTML As String
    Dim campoCapa As String
    Dim campoItem As String
    Dim campo As String
    
    TermoHTML = codigoTermoHTML
    
    campoCapa = "DMACCapa"
    campoItem = "DMACItem"
    
    campo = campoBilhete(TermoHTML, campoCapa)
    Do While campo <> ""
        TermoHTML = Replace(TermoHTML, campoCapa & campo, rsCapaBilhete(campo))
        campo = campoBilhete(TermoHTML, campoCapa)
    Loop
    
    campo = campoBilhete(TermoHTML, campoItem)
    Do While campo <> ""
        TermoHTML = Replace(TermoHTML, campoItem & campo, rsItemBilhete(campo))
        campo = campoBilhete(TermoHTML, campoItem)
    Loop
    
    Open "C:\Sistemas\DMAC Caixa\GarantiaEstendida\Termo.htm" For Output As #1
    Print #1, TermoHTML
    Close #1
    
    WebNavegadorTermo.Navigate "C:\Sistemas\DMAC Caixa\GarantiaEstendida\Termo.htm"
    
End Sub

