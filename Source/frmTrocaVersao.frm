VERSION 5.00
Begin VB.Form frmTrocaVersao 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Atualização do sistema"
   ClientHeight    =   1425
   ClientLeft      =   7530
   ClientTop       =   3900
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1425
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer timerVerificacao 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8280
      Top             =   420
   End
   Begin VB.Timer timerContagemRegresiva 
      Interval        =   1000
      Left            =   6510
      Top             =   600
   End
   Begin VB.CommandButton cmdCancela 
      Caption         =   "Cancelar e continuar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3945
      TabIndex        =   1
      Top             =   2415
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Image imgLogoBlack 
      Height          =   1440
      Left            =   0
      Picture         =   "frmTrocaVersao.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Image imgLogoWhite 
      Height          =   1440
      Left            =   15
      Picture         =   "frmTrocaVersao.frx":1A20
      Top             =   45
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblErro 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   72
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1680
      Left            =   585
      TabIndex        =   3
      Top             =   -360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblAtualiza 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Atualizando em "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   1500
      TabIndex        =   2
      Top             =   780
      Width           =   5340
   End
   Begin VB.Label lblMensagem 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Verificando se há atualização do sitema . . ."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1500
      TabIndex        =   0
      Top             =   285
      Width           =   8640
   End
End
Attribute VB_Name = "frmTrocaVersao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim msgBotaoAtualiza As String
Dim tempoRestante As Byte

Dim endPastaRemoto As String
Dim endPastaLocal As String

Dim nomeExecutavel As String
Dim nomeArquivoBAT As String

Dim I, k, tam As Byte
Dim nomeAplicativo As String
Dim lojaAtualizacao As String

Dim arquivoini, Data As String


Private Function lerArquivoNomeSistema(enderecoPastaSistema) As String

    Dim mensagemArquivoTXT As TextStream
    Dim fso As New FileSystemObject
    
    Set mensagemArquivoTXT = fso.OpenTextFile(enderecoPastaSistema & "trocaversao")
    lerArquivoNomeSistema = mensagemArquivoTXT.ReadLine
    mensagemArquivoTXT.Close
    
End Function

Private Sub imgLogoWhite_Click()
    tempoRestante = 0
    lblAtualiza_Click
    timerContagemRegresiva.Enabled = False
End Sub

Private Sub imgLogoBlack_Click()
    tempoRestante = 0
    lblAtualiza_Click
    timerContagemRegresiva.Enabled = False
End Sub

Private Sub Label1_Click()

End Sub

Private Sub lblAtualiza_Click()

    lblAtualiza.Enabled = False
    lblAtualiza.Caption = "Atualizando sistema . . . "
    cmdCancela.Enabled = False
    timerContagemRegresiva.Enabled = False

    lblMensagem.Caption = "Atualizando versão "
    lblMensagem.Refresh
    
    criaTXTComando montaComandoCMD
    Shell endPastaLocal & nomeArquivoBAT, vbHide
    
    Esperar 3
    
    Unload Me
    
End Sub


Private Function obterServidorPasta()
    Dim Sql As String
    Dim rsControle As New ADODB.Recordset
    
   ' If (arquivoini = "DMACini.mdb") Then
        Sql = "select rtrim(cts_ServidorAtualizacao) as ServidorAtualizacao from controlesistema"
   ' Else
        'sql = "select rtrim(ct_ServidorAtualizacao) as ServidorAtualizacao from controle"
   ' End If
    
    rsControle.CursorLocation = adUseClient
    rsControle.Open Sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
        obterServidorPasta = rsControle("ServidorAtualizacao")
    
    rsControle.Close
End Function

Private Function obterLoja() As String

On Error GoTo TrataErro
  
    Dim Sql As String
    Dim rsControle As New ADODB.Recordset
    
    Sql = "select rtrim(CTS_Loja) as loja from controlesistema"

    rsControle.CursorLocation = adUseClient
    rsControle.Open Sql, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
        obterLoja = rsControle("loja")
    
    rsControle.Close
    
TrataErro:
    obterLoja = ""
    
End Function

Private Sub verificaAppExecucao()
    If App.PrevInstance Then
        End
    End If
End Sub

Private Sub Form_Load()
    
    
    On Error GoTo TrataErro
    
    nomeExecutavel = "dmac_caixa.exe"
    Mensagem nomeExecutavel
    lblMensagem.Caption = "Verificando se há atualização do sitema . . ."
    
    AlwaysOnTop Me, True
    
    lblAtualiza.Caption = ""
    I = 1
    k = 0
    
    nomeArquivoBAT = "update.bat"
    
    tempoRestante = 2
    
    timerContagemRegresiva.Interval = 1000
    timerContagemRegresiva.Enabled = True
    
    
    imgLogoBlack.left = 100
    imgLogoWhite.left = 100
    lblErro.left = 100
   ' Me.Visible = False
    
    CenterForm Me
  
    'endPastaLocal = "c:\sistemas\" & pastaAtual
    endPastaLocal = "c:\sistemas\dmac caixa\"
    
    endPastaRemoto = obterServidorPasta
    'endPastaRemoto = "\\128.0.0.30\sistemas\"
    
    tam = Len(pastaAtual)
    
    endPastaRemoto = obterServidorPasta & pastaAtual
    'endPastaRemoto = "\\128.0.0.30\sistemas\dmac venda\"

    
    lojaAtualizacao = obterLoja
    
    Call atualizaVersao
    
    If k <= 0 Then
      
        Me.Caption = "Atualização do sistema (Versão " & App.Major & "." & App.Minor & "." & App.Revision & ")"
        msgBotaoAtualiza = "Atualizando em "
        
        timerVerificacao.Enabled = True
 
    End If


    Exit Sub
    
TrataErro:
    exibirErro Err.Number
    
End Sub

Private Sub validaUnidade()
    If Mid(LCase(App.Path), 1, 1) <> "c" Then

        End
    End If
End Sub

Private Sub atualizaVersao()
    
    Dim fso As New FileSystemObject
    Dim arqRemoto As File
    Dim arqLocal As File
    Dim id As String
    Dim comandoCDM As String
    
    On Error GoTo TrataErro
    

    'If Dir(endPastaLocal & nomeExecutavel) = "" Then

        'montaComandoCMD

        
    'Else
        Set arqRemoto = fso.GetFile(endPastaRemoto & nomeExecutavel)
        Set arqLocal = fso.GetFile(endPastaLocal & nomeExecutavel)
            Data = arqLocal.DateCreated
        If arqLocal.DateLastModified <> arqRemoto.DateLastModified Or arqLocal.Size <> arqRemoto.Size Then
            'If Right(listaArquivoAtualizacao(i), 3) = "exe" Then
                k = k + 1
                Mensagem nomeExecutavel
            'Else
              '  Kill endPastaLocal & listaArquivoAtualizacao(i)
                'FileCopy endPastaRemoto & listaArquivoAtualizacao(i), endPastaLocal & listaArquivoAtualizacao(i)
            'End If
        End If
    'End If
     
    'Loop
    
    Exit Sub
TrataErro:

    exibirErro Err.Number
    
End Sub



Private Function montaComandoCMD()
    
    montaComandoCMD = "taskkill  /im " & nomeExecutavel & " /f"
    
    montaComandoCMD = montaComandoCMD & vbNewLine & _
                 "xcopy " & Chr(34) & left(endPastaRemoto, Len(endPastaRemoto) - 1) & "" & Chr(34) & _
                 " " & Chr(34) & endPastaLocal & Chr(34) & " /y /c /e /i"
    
    montaComandoCMD = montaComandoCMD & vbNewLine & _
                 "del " & Chr(34) & endPastaLocal & nomeArquivoBAT & Chr(34)
    
    If lojaAtualizacao <> "" Then

        montaComandoCMD = montaComandoCMD & vbNewLine & _
                     "xcopy " & Chr(34) & left(Replace(endPastaRemoto, "Sistemas", "Sistemas por loja\" & _
                     lojaAtualizacao & ""), Len(endPastaRemoto) - 1) & _
                     "" & Chr(34) & " " & Chr(34) & endPastaLocal & Chr(34) & " /y /c /e /i"
                     
    End If
                     
End Function

Private Sub criaTXTComando(comando As String)

On Error GoTo TrataErro
    
 
    Open endPastaLocal & nomeArquivoBAT For Output As #1
         Print #1, comando
    Close #1

    Exit Sub
    
TrataErro:
    Select Case Err.Number
    Case Else
        'mensagemErroDesconhecido Err, "Erro na criação do arquivo"
    End Select
End Sub

Private Sub Mensagem(ByVal nomePrograma As String)
    
    nomePrograma = UCase(nomePrograma)
    If nomePrograma = "SUP.EXE" Then
        nomePrograma = "Suprimentos"
    ElseIf nomePrograma = "GER.EXE" Or nomePrograma = "GERLOJA.EXE" Then
        nomePrograma = "Gerencial"
    Else
        nomePrograma = Replace(nomePrograma, "_", " ")
        nomePrograma = Replace(nomePrograma, ".EXE", "")
        nomePrograma = Replace(nomePrograma, "-", " ")
        nomePrograma = Replace(nomePrograma, ".", " ")
    End If

    If nomePrograma Like "DMAC*" Then
        imgLogoBlack.Visible = True
        imgLogoWhite.Visible = False
        lblAtualiza.BackColor = vbBlack
        lblMensagem.BackColor = vbBlack
        lblAtualiza.ForeColor = vbWhite
        lblMensagem.ForeColor = vbWhite
        Me.BackColor = vbBlack
    Else
        imgLogoWhite.Visible = True
        imgLogoBlack.Visible = False
        lblAtualiza.BackColor = vbWhite
        lblMensagem.BackColor = vbWhite
        lblAtualiza.ForeColor = vbBlack
        lblMensagem.ForeColor = vbBlack
        Me.BackColor = vbWhite
    End If

    lblMensagem.Caption = "Uma nova versão do sistema " & nomePrograma & " está disponível"
    If lojaAtualizacao <> "" Then
        lblMensagem.Caption = lblMensagem.Caption & " para a loja " & lojaAtualizacao & vbNewLine
    End If
    'lblMensagem.Caption = lblMensagem.Caption & "Deseja atualizar agora?"
End Sub



Private Sub exibirErro(mensagemErro As String)

    imgLogoWhite.Visible = False
    'timerAtualiza.Enabled = False
    timerContagemRegresiva.Enabled = False
    lblErro.Visible = True
    lblMensagem.Caption = "Ocorreu um erro ao atualizar a versão"
    lblAtualiza.Caption = "Erro " & mensagemErro & ""
    lblMensagem.ForeColor = vbRed
    lblAtualiza.ForeColor = vbRed
    timerVerificacao.Enabled = True
    
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub timerContagemRegresiva_Timer()
    If tempoRestante = 0 Then
        lblAtualiza_Click
        timerContagemRegresiva.Enabled = False
    Else
        lblAtualiza.Caption = msgBotaoAtualiza & "" & tempoRestante & " segundos"
        tempoRestante = tempoRestante - 1
        'AlwaysOnTop Me, True
        'Me.SetFocus
        If k = 0 Then Unload Me
    End If
End Sub

Sub Esperar(ByVal Tempo As Integer)

    Dim StartTime As Long
    StartTime = Timer
    Do While Timer < StartTime + Tempo
       DoEvents
    Loop
    
End Sub

'Private Function verificaNovoArquivo(endSitema As String, endRemoto As String, Optional tipoArquivo As String) As Boolean
'
'    Dim nomeExecutavel As String
'    On Error GoTo TrataErro
'
'    If tipoArquivo = Empty Then tipoArquivo = "*.exe"
'
'    listaArquivoAtualizacao(i) = Dir(endRemoto, vbDirectory)
'

'    i = i + 1
'    listaArquivoAtualizacao(i) = Dir
'

'
'    Do While listaArquivoAtualizacao(i) <> ""
'        'If Len(listaArquivoAtualizacao(i)) >= 3 Then i = i + 1
'        i = i + 1
'        listaArquivoAtualizacao(i) = Dir
'        'end if
'    Loop
'
'    If listaArquivoAtualizacao(i) = "" Then i = i - 1
'
'    Exit Function
'
'TrataErro:
'    Select Case Err.Number
'        Case 5
'
'            verificaNovoArquivo = False
'        Case Else
'                exibirErro Err.Number
'    End Select
'End Function

Private Function pastaAtual()
    Dim pasta As String
    pastaAtual = Replace((LCase(App.Path)), "c:\sistemas\", "") & "\"
End Function



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



Private Sub timerVerificacao_Timer()
    tempoRestante = tempoRestante + 1
    If tempoRestante > 2 Then
        tempoRestante = 0
        Unload Me
    End If
End Sub


'public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'
Public Function AlwaysOnTop(FrmID As Form, ByVal OnTop As Boolean) As Boolean
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    If OnTop = True Then
        AlwaysOnTop = SetWindowPos(FrmID.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        AlwaysOnTop = SetWindowPos(FrmID.Hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Function

Public Sub CenterForm(X As Form)

    X.Move ((Screen.Width - X.Width) / 2), ((Screen.Height - X.Height) / 2) - 700

End Sub


