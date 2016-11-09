Attribute VB_Name = "ModuloInicial"
Sub main()

    ConectaODBC
    If GLB_ConectouOK = True Then
       SQL = "Select ControleCaixa.*,USU_Codigo,USU_Nome from ControleCaixa,UsuarioCaixa" _
            & " Where CTR_Operador = USU_Codigo and CTR_SituacaoCaixa='A' and CTR_NumeroCaixa = " & wNumeroCaixa
           ' Set RsDados = rdoCNLoja.OpenResultset(SQL)
            RsDados.CursorLocation = adUseClient
            RsDados.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
            If RsDados.EOF = False Then
               GLB_USU_Nome = RsDados("USU_Nome")
               GLB_USU_Codigo = RsDados("USU_Codigo")
               GLB_CTR_Protocolo = RsDados("CTR_Protocolo")
          ' GLB_CTR_Protocolo = 0
            End If
            If RsDados.EOF Then
               RsDados.Close
               frmFundoEscuro.Show
               
               frmLoginCaixa.Show vbModal
               frmLoginCaixa.ZOrder
            Else
               RsDados.Close
               frmControlaCaixa.Show
               frmControlaCaixa.ZOrder
            End If
    Else
        MsgBox "Erro ao conectar-se ao Banco de Dados", vbCritical, "Atenção"
        Exit Sub
    End If

End Sub
