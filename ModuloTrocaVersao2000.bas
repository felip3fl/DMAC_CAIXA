Attribute VB_Name = "ModuloTrocaVersao2000"
Global STRVersao As String
Global nomeArquivo As String, arquivoSaida As String
Global F As String
Global Versao_Atual As String
Global wVerificaSeAquivoExiste

Public Function AtualizaVersao()

nomeArquivo = "Versaobc2000.txt"
arquivoSaida = App.Path & "\Versaobc2000.txt"
    
    STRVersao = "recv " & nomeArquivo & " " & arquivoSaida
    
F = FreeFile

        If mdiBalcao.Inet1.StillExecuting Then
            mdiBalcao.Inet1.Cancel
        
        End If

    With mdiBalcao.Inet1
        .URL = "ftp://ftp2@ftp8.zeronet.com.br"
        .UserName = "ftpdemeo2"
        .Password = "q2w3e4"
        .Execute , STRVersao
         TerminaComando
         
    End With
    
     Open arquivoSaida For Input As #F

        Line Input #F, STRVersao

     Close #F

    Versao_Atual = App.Major & App.Minor & App.Revision

    If STRVersao <> "" And Versao_Atual < STRVersao Then
        
        mdiBalcao.SSPanel1.Visible = True
        
       ' frmTrocaVersao.Show
        Call RenomeiaExecutavelAntigo
                
        nomeArquivo = "bc2000.exe"
        arquivoSaida = App.Path & "\bc2000.exe"
        
        STRVersao = "recv " & nomeArquivo & " " & arquivoSaida
    
    
        With mdiBalcao.Inet1
            .Execute , STRVersao
            TerminaComando
            .Execute , "quit"
            
        End With
        
        

        wVerificaSeAquivoExiste = Dir(App.Path & "\Versaobc2000.txt")
        
        If wVerificaSeAquivoExiste <> "" Then
           Kill App.Path & "\Versaobc2000.txt"
        End If
        
        mdiBalcao.SSPanel1.Visible = False
        MsgBox "A troca de versão foi concluída com sucesso.", vbInformation, ""
        
        End
        Exit Function
    Else
        
        mdiBalcao.Inet1.Execute , "quit"
        
        wVerificaSeAquivoExiste = Dir(App.Path & "\Versaobc2000.txt")
        
        If wVerificaSeAquivoExiste <> "" Then
           Kill App.Path & "\Versaobc2000.txt"
        End If
        
        TerminaComando
    
    End If
    
End Function
Private Sub RenomeiaExecutavelAntigo()
 Dim SourceFile As String
    
      
    Screen.MousePointer = 11
    
    On Error Resume Next
    
    DirApp = App.Path & "\"
    
                wVerificaSeAquivoExiste = Dir(DirApp & "bc2000OLD2.EXE")
                If wVerificaSeAquivoExiste <> "" Then
                   Kill DirApp & "bc2000OLD2.EXE"
                   Err.Clear
                End If
                
                
                wVerificaSeAquivoExiste = Dir(DirApp & "bc2000OLD.EXE")
                If wVerificaSeAquivoExiste <> "" Then
                   Name DirApp & "bc2000OLD.EXE" As DirApp & "bc2000OLD2.EXE"
                   Err.Clear
                End If
                
                wVerificaSeAquivoExiste = Dir(DirApp & "bc2000.EXE")
                  
                If wVerificaSeAquivoExiste <> "" Then
                   Name DirApp & "bc2000.EXE" As DirApp & "bc2000OLD.EXE"
                   Err.Clear
                End If
                
                
                If Err.Number <> 0 Then
                    MsgBox "Erro durante a troca de versão. Tente novamente.", vbCritical, "Troca de Versão"
                    MsgBox Err.Number & ": " & Err.Description, vbCritical, "Troca de Versão"
                    
                    Err.Clear
                    If Dir(DirApp & "bc2000.exe") = "" Then
                        Name DirApp & "bc2000OLD.EXE" As DirApp & "bc2000.EXE"
                        If Err.Number <> 0 Then
                             Name DirApp & "bc2000OLD2.EXE" As DirApp & "bc2000.EXE"
                        End If
                        MsgBox "Ocorreram erros durante a troca de versão. Tente novamente.", vbInformation, "Troca de Versão"
                    End If
                Else
                 '   MsgBox "A troca de versão foi concluída com sucesso.", vbInformation, "Troca de Versão"
                End If
       
    Screen.MousePointer = 0
    Exit Sub

End Sub



Private Sub TerminaComando()
    
    Do While mdiBalcao.Inet1.StillExecuting
        DoEvents
    Loop
    
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)


    Select Case State
        Case icResolvingHost
            mdiBalcao.lblStatus.Caption = "Resolvendo Host"
        Case icHostResolved
            mdiBalcao.lblStatus.Caption = "Host Resolvido"
        Case icConnecting
            mdiBalcao.lblStatus.Caption = "Conectando ..."
        Case icConnected
            mdiBalcao.lblStatus.Caption = "Conectado"
        Case icRequesting
            mdiBalcao.lblStatus.Caption = "Requesitando ..."
        Case icRequestSent
            mdiBalcao.lblStatus.Caption = "Requesição enviada"
        Case icReceivingResponse
            mdiBalcao.lblStatus.Caption = "Recebendo ..."
        Case icResponseReceived
            mdiBalcao.lblStatus.Caption = "Resposta recebida"
        Case icDisconnecting
            mdiBalcao.lblStatus.Caption = "Desconectando ..."
        Case icDisconnected
            mdiBalcao.lblStatus.Caption = "Desconectado"
        Case icError
            mdiBalcao.lblStatus.Caption = itcFTP.ResponseInfo
            mdiBalcao.txtlog.Text = mdiBalcao.txtlog.Text & itcFTP.ResponseInfo & vbCrLf
        Case icResponseCompleted
            mdiBalcao.lblStatus.Caption = "operacao Completa"

    End Select
           
End Sub
