Attribute VB_Name = "modTef_Caixa_Dmac"
Option Explicit


Public Function sequencial_Tef_Vbi() As Integer
 tef_sql = "select CTS_Sequencial from  ControleSistema"
 ADOTef.CursorLocation = adUseClient
 ADOTef.Open tef_sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    If Not ADOTef.EOF Then
        tef_sql = "update ControleSistema set CTS_Sequencial=" & (ADOTef("CTS_Sequencial") + 3)
        rdoCNLoja.BeginTrans
        rdoCNLoja.Execute tef_sql
        rdoCNLoja.CommitTrans
        sequencial_Tef_Vbi = ADOTef("CTS_Sequencial")
    End If
ADOTef.Close

End Function
Public Function ususrio_senha_Tef_Vbi() As Integer
 tef_sql = "SELECT  * FROM  Usuariocaixa WHERE USU_Codigo='" & GLB_Loja & "'"
 ADOTef.CursorLocation = adUseClient
 ADOTef.Open tef_sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    If Not ADOTef.EOF Then
        tef_usuario = Trim(ADOTef("USU_Nome"))
        tef_senha = Trim(ADOTef("USU_Senha"))
    End If
ADOTef.Close

End Function

Public Function Grava_Cupom(ByVal resposta As String)
Dim fso As New FileSystemObject
Dim mensagemArquivoTXT As TextStream
VereficaArquivos ("C:\Sistemas\DMAC Caixa\Tef_Cupom")
   Set mensagemArquivoTXT = fso.OpenTextFile _
                    ("C:\Sistemas\DMAC Caixa\Tef_Cupom")
                    tef_cupom = mensagemArquivoTXT.ReadAll
                    mensagemArquivoTXT.Close
                        Open "C:\Sistemas\DMAC Caixa\Tef_Cupom" For Output As #1
                        Print #1, tef_cupom
                        Print #1, resposta
                        Close #1
End Function
Public Function Imprimir_Tef()

    Dim fso As FileSystemObject
    Dim FileName As File
    Dim TextStream As TextStream
    Dim strText As String
    Dim Texto   As String
    Dim texto1   As Variant

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists("C:\Sistemas\DMAC Caixa\Tef_Cupom") Then
        Set FileName = fso.GetFile("C:\Sistemas\DMAC Caixa\Tef_Cupom")
    Else
        'Do something else
    End If
      'strText = TextStream.ReadLine
    
    Debug.Print Mid(strText, 1, 1)
        
    Set TextStream = FileName.OpenAsTextStream(ForReading, TristateUseDefault)
        
     
    Do While Not TextStream.AtEndOfStream
        strText = TextStream.ReadLine
        If Trim(strText) <> "" Then
                         GLB_Impressora00 = "SAT"
                         impressoraRelatorio ("[INICIO]")
                                                
                       texto1 = Split(strText, ";")
                        For I = LBound(texto1) To UBound(texto1)
                              impressoraRelatorio (texto1(I))
                        Next
                        impressoraRelatorio ("[FIM]")
       End If
    Loop
    
   
    TextStream.Close
    Set TextStream = Nothing
    Set fso = Nothing
 Open "C:\Sistemas\DMAC Caixa\Tef_Cupom " For Output As #1
 Print #1, ""
 Close #1
End Function


Public Function Fecha_Log()
Dim fso As FileSystemObject
    Dim FileName As File
    Dim ts As TextStream
    Dim mensagemArquivoTXT As TextStream
    
    
Set fso = New Scripting.FileSystemObject
Set ts = fso.CreateTextFile("C:\Sistemas\DMAC Caixa\Tef_Log\Tef_Log_" & Format(Date, "DDMMYY"), ForWriting, True)
ts.Close

        
                Set mensagemArquivoTXT = fso.OpenTextFile _
                    ("C:\Sistemas\DMAC Caixa\Tef_Diario")
                    tef_cupom = mensagemArquivoTXT.ReadAll
                    mensagemArquivoTXT.Close
                    



 Open "C:\Sistemas\DMAC Caixa\Tef_Diario " For Output As #1
 Print #1, ""
 Close #1

 Open "C:\Sistemas\DMAC Caixa\Tef_Log\Tef_Log_" & Format(Date, "DDMMYY") For Output As #1
 Print #1, tef_cupom
 Close #1
 
 tef_cupom = ""

End Function

Public Function Grava_Log_Diario(ByVal resposta As String)
Dim fso As New FileSystemObject
Dim mensagemArquivoTXT As TextStream
VereficaArquivos ("C:\Sistemas\DMAC Caixa\Tef_Diario")
   Set mensagemArquivoTXT = fso.OpenTextFile _
                    ("C:\Sistemas\DMAC Caixa\Tef_Diario")
                    tef_cupom = mensagemArquivoTXT.ReadAll
                    mensagemArquivoTXT.Close
                        Open "C:\Sistemas\DMAC Caixa\Tef_Diario " For Output As #1
                        Print #1, tef_cupom
                        Print #1, resposta
                        Close #1
End Function




Public Function Verifica_Tef_Pos()
Dim sql As String

verifica_pos = False
 verifica_tef = False

sql = "select cts_Tef,CTS_LiberaPOS from  ControleSistema"
 verefica.CursorLocation = adUseClient
 verefica.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
 If Not verefica.EOF Then
    If Trim(verefica("cts_Tef")) = "S" Then
        verifica_tef = True
    End If
     If Trim(verefica("CTS_LiberaPOS")) = "S" Then
        verifica_pos = True
    End If
 Else
 
 End If
verefica.Close
End Function
Public Function Fecha_Log_Diario()
Dim fso As New FileSystemObject
Dim mensagemArquivoTXT As TextStream
   Set mensagemArquivoTXT = fso.OpenTextFile _
                    ("C:\Sistemas\DMAC Caixa\Tef_Diario")
                    tef_cupom = mensagemArquivoTXT.ReadAll
                    mensagemArquivoTXT.Close
                        Open "C:\Sistemas\DMAC Caixa\Tef_Diario " For Output As #1
                        Print #1, tef_cupom
                        Print #1, "============================ Fim ==========================="
                        Close #1
End Function
Public Function VereficaArquivos(ByVal url As String)
Dim FileName As File
Dim fso As Object
Dim ts As TextStream
Set fso = CreateObject("Scripting.FileSystemObject")
 
 If fso.FileExists(url) Then
        Set FileName = fso.GetFile(url)
        
    Else
        Set ts = fso.CreateTextFile(url, ForWriting, True)
        ts.WriteLine " "
        ts.Close
    End If
End Function

Public Function ImprimeTef_1() As String
Dim fso As New FileSystemObject
Dim mensagemArquivoTXT As TextStream
VereficaArquivos ("C:\Sistemas\DMAC Caixa\Tef_Cupom")
   Set mensagemArquivoTXT = fso.OpenTextFile _
                    ("C:\Sistemas\DMAC Caixa\Tef_Cupom")
                    tef_cupom = mensagemArquivoTXT.ReadAll
                    mensagemArquivoTXT.Close
        ImprimeTef_1 = tef_cupom

End Function


