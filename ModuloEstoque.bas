Attribute VB_Name = "ModEstoque"


Dim wQuant As Double
Dim wGuardaEstoque As Double
Dim rsMovEstoque As rdoResultset
Dim EstoqueLoja As rdoResultset
Dim EstoqueCentral As rdoResultset
Dim Divergencia As rdoResultset
Dim Controle As rdoResultset
Dim RsPegaDataCaixa As rdoResultset
Dim rsPegaTMnfCapa As rdoResultset
Dim Wcheksaida As Double
Dim WsinalQuant As String
Dim Wsinalvenda As String
Dim wDataVenda As String

Dim wPegaSequenciaEstoque As String
Dim Wchekentrada As Double
Dim wAjusteEntrada As Double
Dim wAjusteSaida As Double
Dim wTransfEntrada As Double
Dim wTransfSaida As Double
Dim wCalculaSaida As Double
Dim wCalculaEntrada As Double
Dim wVendaDevolucao As Double
Dim wVenda As Double
Dim QuantDevolucao As Double
Dim wPegaSequencia As Double
Dim wQuantCancelamento As Double
Dim wCancelamentoVenda As Double
Dim wQuantCancelamentoTransf As Double
Dim wCancelamentoTransf  As Double
Dim Wentrada As Long
Dim wSaida As Long
Dim WSRS As Double

Function AtualizaEstoque(ByVal NumeroDocumento As Double, ByVal SerieDocumento As String, ByVal TipoAtualizacaoEstoque As Double)
          
          If NumeroDocumento > 0 Then
              SQL = "Select el_estoque,EL_UltimaVenda,nfitens.* from EstoqueLoja, nfitens " _
                  & "where nfitens.nf = " & NumeroDocumento & " and nfitens.serie='" & SerieDocumento & "' " _
                  & "and TipoNota not in ('PA') and el_referencia = nfitens.referencia "
              Set RsdadosItens = rdoCnLoja.OpenResultset(SQL)
              
              If Not RsdadosItens.EOF Then
                  Do While Not RsdadosItens.EOF
                      Wchekentrada = 0
                      wAjusteEntrada = 0
                      wAjusteSaida = 0
                      wTransfEntrada = 0
                      wTransfSaida = 0
                      wCalculaSaida = 0
                      wCalculaEntrada = 0
                      Wentrada = 0
                      wSaida = 0
                      wVendaDevolucao = 0
                      wVenda = 0
                      QuantDevolucao = 0
                      wQuantCancelamento = 0
                      wCancelamentoVenda = 0
                      wDataVenda = 0
                      wQuantCancelamentoTransf = 0
                      wCancelamentoTransf = 0
                      TipoAtualizacaoEstoque = Mid(RsdadosItens("TIPOMOVIMENTACAO"), 1, 1)
                        
                            If TipoAtualizacaoEstoque = 1 Then
                                'Wsaida = RsdadosItens("QTDE")
                                'Wentrada = 0
                                Wcheksaida = (RsdadosItens("NF"))
                                Wchekentrada = 0
                                If RsdadosItens("TIPOMOVIMENTACAO") = 11 Then
                                   wDataVenda = Format(Date, "DD/MM/YYYY")
                                   wVenda = RsdadosItens("QTDE")
                                'Else
                                   'wVenda = 0
                                End If
                                If RsdadosItens("TipoMovimentacao") = 12 Then
                                    wTransfSaida = RsdadosItens("QTDE")
                                    wDataVenda = Format(RsdadosItens("EL_UltimaVenda"))
                                Else
                                    wTransfSaida = 0
                                End If
                                If RsdadosItens("TipoMovimentacao") = 13 Then
                                    WSRS = RsdadosItens("QTDE")
                                    wQuantSRS = (WSRS * -1)
                                End If
                                
                            ElseIf TipoAtualizacaoEstoque = 2 Then
                                'Wentrada = RsdadosItens("QTDE")
                                'Wsaida = 0
                                Wchekentrada = RsdadosItens("NF")
                                Wcheksaida = 0
                                If RsdadosItens("TIPOMOVIMENTACAO") = 21 Then
                                   wVendaDevolucao = RsdadosItens("QTDE")
                                   wVendaDevolucao = (wVendaDevolucao * -1)
                                   wCancelamentoVenda = wVendaDevolucao
                                   wQuantCancelamento = RsdadosItens("QTDE")
                                   Wchekentrada = 0
                                   Wcheksaida = (Wchekentrada * -1)
                                Else
                                   wVendaDevolucao = 0
                                End If
                                If RsdadosItens("TIPOMOVIMENTACAO") = 22 Then
                                   wTransfEntrada = RsdadosItens("QTDE")
                                Else
                                   wTransfEntrada = 0
                                End If
                                If RsdadosItens("TIPOMOVIMENTACAO") = 23 Then
                                   'Wvenda = RsDadosItens("QTDE")
                                   'Wvenda = (Wvenda * -1)
                                   QuantDevolucao = RsdadosItens("QTDE")
                                'Else
                                   'wVenda = 0
                                End If
                                If RsdadosItens("TipoMovimentacao") = 24 Then
                                    WSRE = RsdadosItens("QTDE")
                                End If
                                If RsdadosItens("TipoMovimentacao") = 25 Then
                                    wCancelamentoTransf = RsdadosItens("QTDE")
                                    wCancelamentoTransf = (wCancelamentoTransf * -1)
                                    wQuantCancelamentoTransf = RsdadosItens("QTDE")
                                    Wchekentrada = 0
                                End If
                            
                            
                            ElseIf TipoAtualizacaoEstoque = 3 Then
                                If RsdadosItens("TIPOMOVIMENTACAO") = 32 Then
                                    wAjusteSaida = RsdadosItens("QTDE")
                                    wAjusteEntrada = 0
                                    Wcheksaida = RsdadosItens("NF")
                                    Wchekentrada = 0
                                ElseIf RsdadosItens("TIPOMOVIMENTACAO") = 31 Then
                                    wAjusteEntrada = RsdadosItens("QTDE")
                                    wAjusteSaida = 0
                                    Wchekentrada = RsdadosItens("NF")
                                    Wcheksaida = 0
                                End If
                            End If
                                           
                            wCalculaSaida = Val(wVenda + wTransfSaida + wAjusteSaida + WSRS)
                            wCalculaEntrada = Val(wTransfEntrada + wAjusteEntrada + QuantDevolucao + wQuantCancelamento + WSRE + wQuantCancelamentoTransf)
                
                'If Not RsdadosItens.EOF Then
                     
                         wQuant = RsdadosItens("QTDE")
                         wGuardaEstoque = RsdadosItens("EL_Estoque")
                        
                        If TipoAtualizacaoEstoque = 1 Then
                            wQuant = (wQuant * -1)
                        End If
                         SQL = ""
                         'WVENDA = 1
                         'SQL = "Update estoqueloja set el_estoque= el_estoque  + " & wQuant _
                         & ", el_vendames = el_vendames + " & wVenda + wCancelamentoVenda & " " _
                         & ", EL_UltimaVenda='" & Format(Date, "MM/DD/YYYY") & "' " _
                         & "  where el_referencia = '" & RsdadosItens("Referencia") & "'"
                         'rdoCnLoja.Execute (SQL)
                      
                         Call AtualizaMovimentoEstoque
                         'GravaSequenciaLeitura 2, wPegaSequencia, 0
                         RsdadosItens.MoveNext
            Loop
        End If
    End If
    
End Function


Sub AtualizaMovimentoEstoque()
    
    SQL = ""
    SQL = "Select CT_Data,CT_loja from CTcaixa order by CT_data desc"
    Set RsPegaDataCaixa = rdoCnLoja.OpenResultset(SQL)
    If Not RsPegaDataCaixa.EOF Then
       wLoja = RsPegaDataCaixa("CT_loja")
       If IsDate(Wdata) = False Or Wdata = "" Or IsNull(Wdata) = True Then
        
          If Not RsPegaDataCaixa.EOF Then
             Wdata = Format(RsPegaDataCaixa("CT_data"), "dd/mm/yyyy")
          Else
             Wdata = Format(Date, "dd/mm/yyyy")
          End If
       End If
    End If
    
    
    SQL = "Select * from MovimentacaoEstoque where ME_DataMovimento= '" & Format(Wdata, "mm/dd/yyyy") & "' and ME_Referencia = '" & RsdadosItens("Referencia") & "'"
    Set rsMovEstoque = rdoCnLoja.OpenResultset(SQL)
    
        If rsMovEstoque.EOF Then
      
                SQL = "INSERT INTO MovimentacaoEstoque (ME_DataMovimento, ME_Loja, ME_Referencia, ME_Venda, ME_TransferenciaSaida, " _
                & "ME_AjusteSaida,ME_SRS,ME_DevolucaoCompras,ME_SRE,ME_TranferenciaEntrada,ME_AjusteEntrada,ME_EstoqueFinal, " _
                & "ME_CheklistEntrada,ME_ChekListSaida,ME_EstoqueInicial,ME_MovimentoOK,ME_Situacao,ME_DevolucaoVenda)" _
                & "VALUES ('" & Format(Wdata, "mm/dd/yyyy") & "', '" & wLoja & "','" & RsdadosItens("Referencia") & "', " _
                & "" & (wVenda + wVendaDevolucao) & ", " & wTransfSaida & " , " & wAjusteSaida & ", " _
                & "" & WSRS & ", 0, " & 0 & ", " & wTransfEntrada & " ," & wAjusteEntrada & " ," & (wGuardaEstoque + wCalculaEntrada - wCalculaSaida) & "," & Wchekentrada & "," & Wcheksaida & "," & wGuardaEstoque & ",'S',9, " & QuantDevolucao & ")"
                rdoCnLoja.Execute (SQL)
                
                SQL = "Select ME_Sequencia from MovimentacaoEstoque where ME_DataMovimento= '" & Format(Wdata, "mm/DD/yyyy") & "' and ME_Referencia = '" & RsdadosItens("Referencia") & "'"
                    Set rsMovEstoque = rdoCnLoja.OpenResultset(SQL)
                    
                If Not rsMovEstoque.EOF Then
                    wPegaSequencia = rsMovEstoque("ME_Sequencia")
                End If
                
        Else
            
                
                SQL = "UPDATE MovimentacaoEstoque set " _
                & "ME_Venda = ME_venda + " & wVenda & " + " & wVendaDevolucao & ", " _
                & "ME_TransferenciaSaida =ME_TransferenciaSaida + " & wTransfSaida + wCancelamentoTransf & " , " _
                & "ME_AjusteSaida =ME_AjusteSaida + " & wAjusteSaida & " , " _
                & "ME_SRS = ME_SRS + " & WSRS & " , " _
                & "ME_DevolucaoVenda =ME_DevolucaoVenda + " & QuantDevolucao & " , " _
                & "ME_DevolucaoCompras = 0 , " _
                & "ME_SRE = ME_SRE + " & 0 & " , " _
                & "ME_TranferenciaEntrada =ME_TranferenciaEntrada + " & wTransfEntrada & " , " _
                & "ME_AjusteEntrada =ME_AjusteEntrada + " & wAjusteEntrada & " , " _
                & "ME_EstoqueFinal = ME_EstoqueFinal + (" & wCalculaEntrada & " - " & wCalculaSaida & "), " _
                & "ME_ChekListEntrada = ME_ChekListEntrada + " & Wchekentrada & " , " _
                & "ME_ChekListSaida = ME_ChekListSaida + " & Wcheksaida & ", " _
                & "ME_MovimentoOK = 'S', " _
                & "ME_Situacao = '9' " _
                & "where ME_Sequencia = " & rsMovEstoque("ME_Sequencia") & " "
                
                rdoCnLoja.Execute (SQL)
                wPegaSequencia = rsMovEstoque("ME_Sequencia")
        End If
        
        
        If Wchekentrada <> 0 Then
           SQL = ""
           SQL = "Select TM from nfcapa " _
                & "Where nf = " & Wchekentrada & " "
                Set rsPegaTMnfCapa = rdoCnLoja.OpenResultset(SQL)
        ElseIf Wcheksaida <> 0 Then
            SQL = ""
            SQL = "Select TM from nfcapa " _
                & "Where nf = " & Wcheksaida & " "
                Set rsPegaTMnfCapa = rdoCnLoja.OpenResultset(SQL)
        End If
        TipoAtualizacaoEstoque = Mid(RsdadosItens("tipomovimentacao"), 1, 1)
        If TipoAtualizacaoEstoque = 2 Then
            wCalculaEntrada = (wCalculaEntrada * -1)
        End If
        If wVerificaTM = False Then
            'MsgBox "Colocar a atualizacao no dbf aqui", vbExclamation
            SQL = ""
            'SQL = "insert into EstqLoja (Referencia,Quantidade,Situacao) " _
                & "values('" & RsdadosItens("Referencia") & "'," & wCalculaSaida & " + " & wCalculaEntrada & ",'A')"
            '    DBFBanco.Execute (SQL)
        End If

End Sub


 Sub AtualizaEstoqueAnterior()
 
    frmRotinasDiaria.lblProcessos.Caption = "Atualizando Estoque Anterior"
    frmAguarde.lblMensagem.Caption = "Aguarde, Girando o Estoque"
    frmAguarde.PrbContador.Visible = False
    AtualizaProcessoFechamento "Controle", "CT_SeqFechamento", "E"
  
    rdoCnLoja.Execute "update estoqueloja set el_estoqueanterior = 999999 "
    rdoCnLoja.Execute "update estoqueloja set el_estoqueanterior = el_estoque"
    
    AtualizaProcessoFechamento "Controle", "CT_SeqFechamento", "M"
    

End Sub


Function AuditorEstoque(ByVal Data As String)
    
    frmAguarde.lblMensagem.Caption = "Aguarde, Auditando o Estoque"
    ConfereEstoque Data, 1 'NfItens
    ConfereEstoque Data, 2 'ControleEstoque
    AtualizaProcessoFechamento "Controle", "CT_SeqFechamento", "D"
    
End Function

Function RecalculaEstoque(ByVal Referencia As String, ByVal Saidas As Integer, ByVal Entradas As Integer, ByVal EstoqueInicial As Integer)
    
    SQL = ""
    'SQL = "Update EstoqueLoja set EL_Estoque=EL_EstoqueAnterior + " & (Entradas - Saidas) & " " _
        & "where EL_Referencia='" & Referencia & "' "
    '   rdoCnLoja.Execute (SQL)
        
End Function

Function GravaControleEstoque(ByVal Referencia As String, ByVal qtde As Integer, ByVal Tipo As String, ByVal Documento As Double)

    SQL = ""
    SQL = "Insert into ControleEstoque (CE_Referencia,CE_Quantidade,CE_NumeroDocumento,CE_TipoDocumento,CE_Data) " _
        & "Values ('" & Referencia & "'," & qtde & "," & Documento & ",'" & Tipo & "','" & Format(Date, "mm/dd/yyyy") & "') "
        rdoCnLoja.Execute (SQL)

End Function


Function ConfereEstoque(ByVal Data As String, ByVal Tipo As Integer)
    Dim rsPegaReferencias As rdoResultset
    Dim rsPegaRefControle As rdoResultset
    Dim wSaidaItens As Integer
    Dim wEntradaItens As Integer
    Dim wReferencia As String
    
    
    If Tipo = 1 Then
        SQL = ""
        SQL = "Select Count(Referencia) as Total from NfItens " _
            & "where DataEmi='" & Format(Data, "mm/dd/yyyy") & "' and TipoNota in ('T','V','E') "
            Set rsPegaReferencias = rdoCnLoja.OpenResultset(SQL)
        If Not rsPegaReferencias.EOF Then
            If rsPegaReferencias("Total") > 0 Then
                frmAguarde.PrbContador.Visible = True
                frmAguarde.PrbContador.Max = rsPegaReferencias("Total")
                frmAguarde.PrbContador.Value = 0
                rsPegaReferencias.Close
                i = 0
            End If
        End If
        SQL = ""
        SQL = "Select Referencia,TipoNota,sum(Qtde) as QtdeItens from NfItens " _
            & "where DataEmi = '" & Format(Data, "mm/dd/yyyy") & "' " _
            & "and TipoNota in ('T','V','E') " _
            & "group by Referencia,TipoNota order by Referencia"
            Set rsPegaReferencias = rdoCnLoja.OpenResultset(SQL)
        If Not rsPegaReferencias.EOF Then
            i = 0
            wReferencia = rsPegaReferencias("Referencia")
            
            wEntradaItens = 0
            wSaidaItens = 0
            Do While Not rsPegaReferencias.EOF
                i = i + 1
                frmAguarde.PrbContador.Value = i
                If wReferencia = rsPegaReferencias("Referencia") Then
                    If rsPegaReferencias("TipoNota") = "T" Or rsPegaReferencias("TipoNota") = "V" Then
                        wSaidaItens = wSaidaItens + rsPegaReferencias("QtdeItens")
                    Else
                        wEntradaItens = wEntradaItens + rsPegaReferencias("QtdeItens")
                    End If
                Else
                    SQL = ""
                    SQL = "Select CE_Quantidade from ControleEstoque " _
                        & "where CE_Referencia='" & wReferencia & "' " _
                        & "and CE_Data='" & Format(Data, "mm/dd/yyyy") & "' "
                    Set rsPegaRefControle = rdoCnLoja.OpenResultset(SQL)
                    Do While Not rsPegaRefControle.EOF
                        If rsPegaRefControle("CE_Quantidade") > 0 Then
                            wEntradaItens = wEntradaItens + rsPegaRefControle("CE_Quantidade")
                        Else
                            wSaidaItens = wSaidaItens + (rsPegaRefControle("CE_Quantidade") * -1)
                        End If
                        rsPegaRefControle.MoveNext
                    Loop
                    ConfereNfItensXMovimentacao wReferencia, Data, wSaidaItens, wEntradaItens
                    wEntradaItens = 0
                    wSaidaItens = 0
                    If rsPegaReferencias("TipoNota") = "T" Or rsPegaReferencias("TipoNota") = "V" Then
                        wSaidaItens = wSaidaItens + rsPegaReferencias("QtdeItens")
                    Else
                        wEntradaItens = wEntradaItens + rsPegaReferencias("QtdeItens")
                    End If
                End If
                wReferencia = rsPegaReferencias("Referencia")
                rsPegaReferencias.MoveNext
            Loop
            SQL = ""
            SQL = "Select CE_Quantidade from ControleEstoque " _
                & "where CE_Referencia='" & wReferencia & "' " _
                & "and CE_Data='" & Format(Data, "mm/dd/yyyy") & "' "
            Set rsPegaRefControle = rdoCnLoja.OpenResultset(SQL)
            Do While Not rsPegaRefControle.EOF
                i = i + 1
                frmAguarde.PrbContador.Value = i
                If rsPegaRefControle("CE_Quantidade") > 0 Then
                    wEntradaItens = wEntradaItens + rsPegaRefControle("CE_Quantidade")
                Else
                    wSaidaItens = wSaidaItens + (rsPegaRefControle("CE_Quantidade") * -1)
                End If
                rsPegaRefControle.MoveNext
            Loop
            ConfereNfItensXMovimentacao wReferencia, Data, wSaidaItens, wEntradaItens
            wEntradaItens = 0
            wSaidaItens = 0
        End If
    ElseIf Tipo = 2 Then
        SQL = ""
        SQL = "Select Count(CE_Referencia) as Total from ControleEstoque " _
            & "where CE_Data='" & Format(Data, "mm/dd/yyyy") & "' "
            Set rsPegaReferencias = rdoCnLoja.OpenResultset(SQL)
        If Not rsPegaReferencias.EOF Then
            If rsPegaReferencias("Total") > 0 Then
                frmAguarde.PrbContador.Visible = True
                frmAguarde.PrbContador.Max = rsPegaReferencias("Total")
                frmAguarde.PrbContador.Value = 0
                rsPegaReferencias.Close
                i = 0
            End If
        End If
        
        SQL = ""
        SQL = "Select * from ControleEstoque " _
            & "where CE_Data='" & Format(Data, "mm/dd/yyyy") & "' order by CE_referencia"
        Set rsPegaRefControle = rdoCnLoja.OpenResultset(SQL)
        If Not rsPegaRefControle.EOF Then
            i = 0
            wReferencia = rsPegaRefControle("CE_Referencia")
            wEntradaItens = 0
            wSaidaItens = 0
            Do While Not rsPegaRefControle.EOF
                i = i + 1
                frmAguarde.PrbContador.Value = i
                If wReferencia = rsPegaRefControle("CE_Referencia") Then
                    If rsPegaRefControle("CE_Quantidade") > 0 Then
                        wEntradaItens = wEntradaItens + rsPegaRefControle("CE_Quantidade")
                    Else
                        wSaidaItens = wSaidaItens + (rsPegaRefControle("CE_Quantidade") * -1)
                    End If
                Else
                    SQL = ""
                    SQL = "Select Referencia,TipoNota,sum(Qtde) as QtdeItens from NfItens " _
                        & "where DataEmi = '" & Format(Data, "mm/dd/yyyy") & "' " _
                        & "and TipoNota in ('T','V','E') and Referencia='" & wReferencia & "'" _
                        & "group by Referencia,TipoNota order by Referencia"
                        Set rsPegaReferencias = rdoCnLoja.OpenResultset(SQL)
                    Do While Not rsPegaReferencias.EOF
                        If rsPegaReferencias("TipoNota") = "T" Or rsPegaReferencias("TipoNota") = "V" Then
                            wSaidaItens = wSaidaItens + rsPegaReferencias("QtdeItens")
                        Else
                            wEntradaItens = wEntradaItens + rsPegaReferencias("QtdeItens")
                        End If
                        rsPegaReferencias.MoveNext
                    Loop
                    ConfereNfItensXMovimentacao wReferencia, Data, wSaidaItens, wEntradaItens
                    wEntradaItens = 0
                    wSaidaItens = 0
                    If rsPegaRefControle("CE_Quantidade") > 0 Then
                        wEntradaItens = wEntradaItens + rsPegaRefControle("CE_Quantidade")
                    Else
                        wSaidaItens = wSaidaItens + (rsPegaRefControle("CE_Quantidade") * -1)
                    End If
                End If
                wReferencia = rsPegaRefControle("CE_Referencia")
                rsPegaRefControle.MoveNext
            Loop
            SQL = ""
            SQL = "Select Referencia,TipoNota,sum(Qtde) as QtdeItens from NfItens " _
                & "where DataEmi = '" & Format(Data, "mm/dd/yyyy") & "' " _
                & "and TipoNota in ('T','V','E') and Referencia='" & wReferencia & "'" _
                & "group by Referencia,TipoNota order by Referencia"
                Set rsPegaReferencias = rdoCnLoja.OpenResultset(SQL)
            Do While Not rsPegaReferencias.EOF
                i = i + 1
                frmAguarde.PrbContador.Value = 1
                If rsPegaReferencias("TipoNota") = "T" Or rsPegaReferencias("TipoNota") = "V" Then
                    wSaidaItens = wSaidaItens + rsPegaReferencias("QtdeItens")
                Else
                    wEntradaItens = wEntradaItens + rsPegaReferencias("QtdeItens")
                End If
                rsPegaReferencias.MoveNext
            Loop
            ConfereNfItensXMovimentacao wReferencia, Data, wSaidaItens, wEntradaItens
            wEntradaItens = 0
            wSaidaItens = 0
        End If
                
    End If
End Function

Function ConfereMovimentacaoXestoque(ByVal Referencia As String, ByVal Data As String) As Boolean
    Dim rsMvEstoque As rdoResultset
    Dim rsMovXEstoque As rdoResultset
    
    SQL = ""
    SQL = "Select ME_EstoqueFinal,ME_EstoqueInicial,EL_Estoque " _
        & "from MovimentacaoEstoque,EstoqueLoja " _
        & "where ME_Referencia='" & Referencia & "' " _
        & "and ME_DataMovimento='" & Format(Data, "mm/dd/yyyy") & "' " _
        & "and EL_referencia=ME_Referencia " _
        & "and ME_estoqueFinal = EL_estoque "
    Set rsMovXEstoque = rdoCnLoja.OpenResultset(SQL)
    If Not rsMovXEstoque.EOF Then
        ConfereMovimentacaoXestoque = True
    Else
        ConfereMovimentacaoXestoque = False
    End If
    rsMovXEstoque.Close
    
End Function

Function ConfereNfItensXMovimentacao(ByVal Referencia As String, ByVal Data As String, ByVal Saidas As Integer, ByVal Entradas As Integer) As Boolean

    Dim rsMvEstoque As rdoResultset
    
    SQL = "Select ME_Referencia,ME_EstoqueInicial,ME_EStoqueFinal,(ME_Venda+ME_TransferenciaSaida+ME_AjusteSaida+ME_SRS) as ME_Saidas, " _
        & "(ME_DevolucaoCompras+ME_SRE+ME_TranferenciaEntrada+ME_AjusteEntrada+ME_DevolucaoVenda) as ME_Entradas " _
        & "from MovimentacaoEstoque " _
        & "where ME_Referencia='" & Referencia & "' " _
        & "and ME_DataMovimento='" & Format(Data, "mm/dd/yyyy") & "' " _
        & "group by ME_Referencia,ME_Venda,ME_TransferenciaSaida,ME_AjusteSaida,ME_SRS,ME_EstoqueInicial,ME_EStoqueFinal, " _
        & "ME_DevolucaoCompras , ME_SRE, ME_TranferenciaEntrada, ME_AjusteEntrada, ME_DevolucaoVenda "
    Set rsMvEstoque = rdoCnLoja.OpenResultset(SQL)
    If Not rsMvEstoque.EOF Then
        If rsMvEstoque("ME_Saidas") <> Saidas Or rsMvEstoque("ME_Entradas") <> Entradas Or rsMvEstoque("ME_EstoqueFinal") <> (rsMvEstoque("ME_EstoqueInicial") - Saidas + Entradas) Then
            AcertaMovimentacaoEstoque Data, Referencia, "S"
            RecalculaEstoque Referencia, Saidas, Entradas, rsMvEstoque("ME_EstoqueInicial")
        Else
            If ConfereMovimentacaoXestoque(Referencia, Data) = False Then
                RecalculaEstoque Referencia, Saidas, Entradas, rsMvEstoque("ME_EstoqueFinal")
            End If
        End If
    Else
        AcertaMovimentacaoEstoque Data, Referencia, "N"
    End If
    
End Function


Function AcertaMovimentacaoEstoque(ByVal Data As String, ByVal Referencia As String, ByVal Existe As String)
    
    Dim rsAcertaMovimento As rdoResultset
    Dim rsControleEstoque As rdoResultset
    Dim wVenda As Integer
    Dim wTransfEntrada As Integer
    Dim wTransfSaida As Integer
    Dim wAjusteSaida As Integer
    Dim wAjusteEntrada As Integer
    Dim wDevolucaoVenda As Integer
    Dim wTotalEstoqueEntrada As Integer
    Dim wTemMovimento As Boolean
    Dim wTotalEstoqueSaida As Integer
    Dim wCompras As Integer
    Dim rsEstoqueInicial As rdoResultset

    wVenda = 0
    wTransfEntrada = 0
    wTransfSaida = 0
    wAjusteSaida = 0
    wAjusteEntrada = 0
    wDevolucaoVenda = 0
    wTotalEstoqueEntrada = 0
    wTotalEstoqueSaida = 0
    wCompras = 0
    wTemMovimento = False
    
    SQL = ""
    SQL = "Select Referencia,TipoNota,sum(Qtde) as QtdeItens from NfItens " _
        & "where DataEmi = '" & Format(Data, "mm/dd/yyyy") & "' " _
        & "and TipoNota in ('T','V','E') and Referencia='" & Referencia & "'" _
        & "group by Referencia,TipoNota order by Referencia"
        Set rsAcertaMovimento = rdoCnLoja.OpenResultset(SQL)
    Do While Not rsAcertaMovimento.EOF
        wTemMovimento = True
        If rsAcertaMovimento("TipoNota") = "V" Then
            wVenda = wVenda + rsAcertaMovimento("QtdeItens")
        ElseIf rsAcertaMovimento("TipoNota") = "T" Then
            wTransfSaida = wTransfSaida + rsAcertaMovimento("QtdeItens")
        ElseIf rsAcertaMovimento("TipoNota") = "E" Then
            wDevolucaoVenda = wDevolucaoVenda + rsAcertaMovimento("QtdeItens")
        End If
        rsAcertaMovimento.MoveNext
    Loop
    
    SQL = ""
    SQL = "Select CE_Quantidade,CE_TipoDocumento from ControleEstoque " _
        & "where CE_referencia='" & Referencia & "' " _
        & "and CE_Data='" & Format(Data, "mm/dd/yyyy") & "'"
        Set rsControleEstoque = rdoCnLoja.OpenResultset(SQL)
    Do While Not rsControleEstoque.EOF
        wTemMovimento = True
        If rsControleEstoque("CE_TipoDocumento") = "A" Then
            If rsControleEstoque("CE_Quantidade") > 0 Then
                wAjusteEntrada = wAjusteEntrada + rsControleEstoque("CE_Quantidade")
            Else
                wAjusteSaida = wAjusteSaida + (rsControleEstoque("CE_Quantidade") * -1)
            End If
        ElseIf rsControleEstoque("CE_tipoDocumento") = "T" Then
            wTransfEntrada = wTransfEntrada + rsControleEstoque("CE_Quantidade")
        ElseIf rsControleEstoque("CE_tipoDocumento") = "C" Then
            wCompras = wCompras + rsControleEstoque("CE_Quantidade")
        End If
        rsControleEstoque.MoveNext
    Loop
    
    If wTemMovimento = True Then
        If Existe = "S" Then
            wTotalEstoqueEntrada = (wTransfEntrada + wDevolucaoVenda + wAjusteEntrada + wCompras)
            wTotalEstoqueSaida = (wVenda + wTransfSaida + wAjusteSaida)
            
            SQL = ""
            SQL = "Update MovimentacaoEstoque set ME_Venda = " & wVenda & ", " _
                & "ME_TransferenciaSaida = " & wTransfSaida & ", " _
                & "ME_AjusteSaida=" & wAjusteSaida & ", " _
                & "ME_TranferenciaEntrada=" & wTransfEntrada & ", " _
                & "ME_AjusteEntrada = " & wAjusteEntrada & ", " _
                & "ME_DevolucaoVenda=" & wDevolucaoVenda & ", " _
                & "ME_SRE=" & wCompras & ", " _
                & "ME_EstoqueFinal=(ME_EstoqueInicial - " & wTotalEstoqueSaida & ") + " & wTotalEstoqueEntrada & ", " _
                & "ME_MovimentoOK='N' " _
                & "where ME_Datamovimento='" & Format(Data, "mm/dd/yyyy") & "' " _
                & "and ME_Referencia = '" & Referencia & "'"
            rdoCnLoja.Execute (SQL)
        ElseIf Existe = "N" Then
            wTotalEstoqueEntrada = (wTransfEntrada + wDevolucaoVenda + wAjusteEntrada)
            wTotalEstoqueSaida = (wVenda + wTransfSaida + wAjusteSaida)
            SQL = ""
            SQL = "Select EL_EstoqueAnterior from EstoqueLoja " _
                & "where EL_Referencia='" & Referencia & "'"
                Set rsEstoqueInicial = rdoCnLoja.OpenResultset(SQL)
            If Not rsEstoqueInicial.EOF Then
                SQL = ""
                SQL = "Insert into MovimentacaoEstoque (ME_Loja,ME_Referencia,ME_EstoqueInicial,ME_DataMovimento,ME_Venda, " _
                    & "ME_TransferenciaSaida,ME_AjusteSaida,ME_TranferenciaEntrada,ME_AjusteEntrada, " _
                    & "ME_DevolucaoVenda,ME_estoqueFinal,ME_SRE,ME_MovimentoOK,ME_Situacao) " _
                    & "Values ('" & AchaLojaControle & "','" & Referencia & "'," & rsEstoqueInicial("EL_EstoqueAnterior") & ",'" & Format(Data, "mm/dd/yyyy") & "'," & wVenda & ", " _
                    & "" & wTransfSaida & ", " & wAjusteSaida & ", " & wTransfEntrada & ", " & wAjusteEntrada & ", " _
                    & "" & wDevolucaoVenda & ", " & (rsEstoqueInicial("EL_EstoqueAnterior") - wTotalEstoqueSaida) + wTotalEstoqueEntrada & "," & wCompras & ",'N',9)"
                    rdoCnLoja.Execute (SQL)
                
                RecalculaEstoque Referencia, wTotalEstoqueSaida, wTotalEstoqueEntrada, rsEstoqueInicial("EL_EstoqueAnterior")
            End If
        End If
    End If
    
End Function


Function AcertaEstoque(ByVal Data As String, ByVal Referencia As String)
    
    Dim rsAcertaMovimento As rdoResultset
    Dim rsControleEstoque As rdoResultset
    Dim wVenda As Integer
    Dim wTransfEntrada As Integer
    Dim wTransfSaida As Integer
    Dim wAjusteSaida As Integer
    Dim wAjusteEntrada As Integer
    Dim wDevolucaoVenda As Integer
    Dim wTotalEstoqueEntrada As Integer
    Dim wTemMovimento As Boolean
    Dim wTotalEstoqueSaida As Integer
    Dim wCompras As Integer
    Dim rsEstoqueAnterior As rdoResultset

    wVenda = 0
    wTransfEntrada = 0
    wTransfSaida = 0
    wAjusteSaida = 0
    wAjusteEntrada = 0
    wDevolucaoVenda = 0
    wTotalEstoqueEntrada = 0
    wTotalEstoqueSaida = 0
    wCompras = 0
    wTemMovimento = False
    
    SQL = ""
    SQL = "Select Referencia,TipoNota,sum(Qtde) as QtdeItens from NfItens " _
        & "where DataEmi = '" & Format(Data, "mm/dd/yyyy") & "' " _
        & "and TipoNota in ('T','V','E') and Referencia='" & Referencia & "'" _
        & "group by Referencia,TipoNota order by Referencia"
        Set rsAcertaMovimento = rdoCnLoja.OpenResultset(SQL)
    Do While Not rsAcertaMovimento.EOF
        wTemMovimento = True
        If rsAcertaMovimento("TipoNota") = "V" Then
            wVenda = wVenda + rsAcertaMovimento("QtdeItens")
        ElseIf rsAcertaMovimento("TipoNota") = "T" Then
            wTransfSaida = wTransfSaida + rsAcertaMovimento("QtdeItens")
        ElseIf rsAcertaMovimento("TipoNota") = "E" Then
            wDevolucaoVenda = wDevolucaoVenda + rsAcertaMovimento("QtdeItens")
        End If
        rsAcertaMovimento.MoveNext
    Loop
    
    SQL = ""
    SQL = "Select CE_Quantidade,CE_TipoDocumento from ControleEstoque " _
        & "where CE_referencia='" & Referencia & "' " _
        & "and CE_Data='" & Format(Data, "mm/dd/yyyy") & "'"
        Set rsControleEstoque = rdoCnLoja.OpenResultset(SQL)
    Do While Not rsControleEstoque.EOF
        wTemMovimento = True
        If rsControleEstoque("CE_TipoDocumento") = "A" Then
            If rsControleEstoque("CE_Quantidade") > 0 Then
                wAjusteEntrada = wAjusteEntrada + rsControleEstoque("CE_Quantidade")
            Else
                wAjusteSaida = wAjusteSaida + (rsControleEstoque("CE_Quantidade") * -1)
            End If
        ElseIf rsControleEstoque("CE_tipoDocumento") = "T" Then
            wTransfEntrada = wTransfEntrada + rsControleEstoque("CE_Quantidade")
        ElseIf rsControleEstoque("CE_tipoDocumento") = "C" Then
            wCompras = wCompras + rsControleEstoque("CE_Quantidade")
        End If
        rsControleEstoque.MoveNext
    Loop
    
    wTotalEstoqueEntrada = (wTransfEntrada + wDevolucaoVenda + wAjusteEntrada + wCompras)
    wTotalEstoqueSaida = (wVenda + wTransfSaida + wAjusteSaida)
                
    SQL = ""
    SQL = "Select EL_EstoqueAnterior from EstoqueLoja " _
        & "where EL_Referencia='" & Referencia & "'"
        Set rsEstoqueAnterior = rdoCnLoja.OpenResultset(SQL)
    If Not rsEstoqueAnterior.EOF Then
        RecalculaEstoque Referencia, wTotalEstoqueSaida, wTotalEstoqueEntrada, rsEstoqueAnterior("EL_EstoqueAnterior")
    End If
    
End Function

Function AcertaEstoqueDBF(ByVal Data As String)
    Dim RsPegaEstoqueLoja As rdoResultset
    
    FileCopy Mid(WbancoAccess, 1, Len(WbancoAccess) - 8) & "vazios\estoqlj.dbf", Mid(WbancoAccess, 1, Len(WbancoAccess) - 8) & "estoqlj.dbf"
    
    SQL = ""
    SQL = "Delete * from EstqLojaDBF"
        rdoCnLoja.Execute (SQL)
    
    'SQL = ""
    'SQL = "Delete * from EstoqLjDBF"
        'rdoCnLoja.Execute (SQL)
    
    SQL = ""
    SQL = "Select count(*) as Total from EstoqueLoja"
        Set RsPegaEstoqueLoja = rdoCnLoja.OpenResultset(SQL)
    If Not RsPegaEstoqueLoja.EOF Then
        frmAguarde.PrbContador.Visible = True
        frmAguarde.PrbContador.Max = RsPegaEstoqueLoja("Total")
        frmAguarde.PrbContador.Value = 0
        frmAguarde.lblMensagem.Caption = "Aguarde, Estoque DBF"
        frmAguarde.Refresh
        RsPegaEstoqueLoja.Close
    End If
    
    SQL = ""
    SQL = "Select EL_Referencia,EL_Estoque from EstoqueLoja order by EL_Referencia"
        Set RsPegaEstoqueLoja = rdoCnLoja.OpenResultset(SQL)
    If Not RsPegaEstoqueLoja.EOF Then
        i = 0
        frmAguarde.lblMensagem.Caption = "Aguarde, Estoque DBF"
        Do While Not RsPegaEstoqueLoja.EOF
            i = i + 1
            frmAguarde.PrbContador.Value = i
            SQL = ""
            SQL = "Insert into EstoqLjDBF (Referencia,Estoque,DataEstq,Situacao) " _
                & "Values ('" & RsPegaEstoqueLoja("EL_Referencia") & "'," & RsPegaEstoqueLoja("EL_Estoque") & ", " _
                & "'" & Format(Data, "mm/dd/yyyy") & "','A')"
            rdoCnLoja.Execute (SQL)
            RsPegaEstoqueLoja.MoveNext
        Loop
        SQL = ""
        SQL = "Update ControleDBF set Backup='A'"
            rdoCnLoja.Execute (SQL)
    End If
    
    FileCopy Mid(WbancoAccess, 1, Len(WbancoAccess) - 8) & "estoqlj.dbf", WbancoDbf & "estoqlj.dbf"
    AtualizaProcessoFechamento "Controle", "CT_SeqFechamento", "P"
        
End Function

Function GravaProduto(ByVal Referencia As String)
    Dim rdoDadosProdu As rdoResultset
    
    Screen.MousePointer = 11
    On Error Resume Next
    If ConectaODBC(ConexaoBach, "sa", "jeda36") = False Then
        MsgBox "Erro conectando ao banco de dados.", vbCritical, "Atenção"
        Exit Function
    End If
    SQL = ""
    SQL = "Select * from Produto where PR_Referencia='" & Referencia & "'"
        Set rdoDadosProdu = ConexaoBach.OpenResultset(SQL)
    If Not rdoDadosProdu.EOF Then
        On Error Resume Next
        BeginTrans
        
        SQL = ""
        SQL = "Delete Produto where Pr_referencia='" & Referencia & "' "
            rdoCnLoja.Execute (SQL)
        
        'SQL = "Insert into Produto (PR_Referencia,PR_CodigoFornecedor,PR_CodigoBarra,PR_Descricao,PR_DataCadastro,PR_Linha,PR_Secao,PR_Classe,PR_Bloqueio,PR_ClasseABC,PR_ClasseFiscal,PR_Unidade,PR_UnidadeDistribuicao,PR_PercentualComissao,PR_ICMSEntrada,PR_ICMSSaida,PR_AliquotaIPI, " _
            & "PR_CodigoIPI,PR_CodigoReducaoICMS,PR_PrecoFornecedor,PR_DescontoFornecedor,PR_ItemCondicoesGerais, PR_PercentualFrete,PR_PercentualEmbalagem,PR_CustoMedio1,PR_CustoMedio2,PR_CustoMedio3,PR_CustoMedioLiquido1,PR_CustoMedioLiquido2,PR_CustoMedioLiquido3,PR_PrecoCusto1,PR_PrecoCusto2," _
            & "PR_PrecoCusto3,PR_CustoLiquido1,PR_CustoLiquido2,PR_CustoLiquido3,PR_PrecoEntrada1,PR_PrecoEntrada2, PR_PrecoEntrada3,PR_DataPrecoCusto1,PR_DataPrecoCusto2,PR_DataPrecoCusto3,PR_PrecoVenda1,PR_PrecoVenda2, PR_PrecoVenda3,PR_DataPrecoVenda1,PR_DataPrecoVenda2,PR_DataPrecoVenda3,PR_PrecoVendaObjetivo,PR_PaginaListaPreco, " _
            & "PR_Peso,PR_MenorUnidadeCompra,PR_MetodoCompra,PR_TipoCalculoReposicao,PR_MetodoDistribuicao,PR_Residencia,PR_MargemObjetiva,PR_MargemPrevista,PR_Markup,PR_Comprador,PR_EmiteEtiqueta,PR_Situacao,PR_SubstituicaoTributaria, PR_DeducoesVenda,PR_IcmPdv,PR_DescricaoPDV,PR_Grupo,PR_Complemento,PR_Sazonal,PR_CodigoProdutoNoFornecedor, " _
            & "PR_PrecoVendaLiquido1,PR_PrecoVendaLiquido2,PR_PrecoVendaLiquido3,PR_HoraManutencao) Values ('" & rdoDadosProdu("PR_Referencia") & "'," _
            & rdoDadosProdu("PR_CodigoFornecedor") & ", '" & rdoDadosProdu("PR_CodigoBarra") & "', '" & rdoDadosProdu("PR_Descricao") & "', '" & Format(rdoDadosProdu("PR_DataCadastro"), "mm/dd/yyyy") & "', " & rdoDadosProdu("PR_Linha") & ", " & rdoDadosProdu("PR_Secao") & ",'" & rdoDadosProdu("PR_Classe") & "' , " _
            & "'" & rdoDadosProdu("PR_Bloqueio") & "', '" & rdoDadosProdu("PR_ClasseABC") & "', '" & rdoDadosProdu("PR_ClasseFiscal") & "', '" & rdoDadosProdu("PR_Unidade") & "', " & rdoDadosProdu("PR_UnidadeDistribuicao") & ", 2.00, " & ConverteVirgula(rdoDadosProdu("PR_ICMSEntrada")) & ", " & ConverteVirgula(rdoDadosProdu("PR_ICMSSaida")) & ", " _
            & ConverteVirgula(rdoDadosProdu("PR_AliquotaIPI")) & ", " & rdoDadosProdu("PR_CodigoIPI") & ", " & rdoDadosProdu("PR_CodigoReducaoICMS") & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoFornecedor")) & ", " & ConverteVirgula(rdoDadosProdu("PR_DescontoFornecedor")) & ", " & rdoDadosProdu("PR_ItemCondicoesGerais") & ", " & ConverteVirgula(rdoDadosProdu("PR_PercentualFrete")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PercentualEmbalagem")) & ", " _
            & ConverteVirgula(rdoDadosProdu("PR_CustoMedio1")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoMedio2")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoMedio3")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoMedioLiquido1")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoMedioLiquido2")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoMedioLiquido3")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoCusto1")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoCusto2")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoCusto3")) & ", " _
            & ConverteVirgula(rdoDadosProdu("PR_CustoLiquido1")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoLiquido2")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoLiquido3")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoEntrada1")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoEntrada2")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoEntrada3")) & ", '" & Format(rdoDadosProdu("PR_DataPrecoCusto1"), "mm/dd/yyyy") & "','" & Format(rdoDadosProdu("PR_DataPrecoCusto2"), "mm/dd/yyyy") & "', '" & Format(rdoDadosProdu("PR_DataPrecoCusto3"), "mm/dd/yyyy") & "', " & ConverteVirgula(rdoDadosProdu("PR_PrecoVenda1")) & ", " _
            & ConverteVirgula(rdoDadosProdu("PR_PrecoVenda2")) & "," & ConverteVirgula(rdoDadosProdu("PR_PrecoVenda3")) & ", '" & Format(rdoDadosProdu("PR_DataPrecoVenda1"), "mm/dd/yyyy") & "',' " & Format(rdoDadosProdu("PR_DataPrecoVenda2"), "mm/dd/yyyy") & "' ,'" & Format(rdoDadosProdu("PR_DataPrecoVenda3"), "mm/dd/yyyy") & "', " & ConverteVirgula(rdoDadosProdu("PR_PrecoVendaObjetivo")) & ", " & rdoDadosProdu("PR_PaginaListaPreco") & ", " & ConverteVirgula(rdoDadosProdu("PR_Peso")) & ", " & rdoDadosProdu("PR_MenorUnidadeCompra") & ", " & rdoDadosProdu("PR_MetodoCompra") & ", " _
            & rdoDadosProdu("PR_TipoCalculoReposicao") & " , " & rdoDadosProdu("PR_MetodoDistribuicao") & ", " & rdoDadosProdu("PR_Residencia") & ", " & ConverteVirgula(rdoDadosProdu("PR_MargemObjetiva")) & ", " & ConverteVirgula(rdoDadosProdu("PR_MargemPrevista")) & ", " & ConverteVirgula(rdoDadosProdu("PR_Markup")) & ", " & rdoDadosProdu("PR_Comprador") & ", '" & rdoDadosProdu("PR_EmiteEtiqueta") & "', '" & rdoDadosProdu("PR_Situacao") & "', '" & rdoDadosProdu("PR_SubstituicaoTributaria") & "', " _
            & ConverteVirgula(rdoDadosProdu("PR_DeducoesVenda")) & ", " & ConverteVirgula(rdoDadosProdu("PR_IcmPdv")) & ", '" & rdoDadosProdu("PR_DescricaoPDV") & "', " & rdoDadosProdu("PR_Grupo") & ", '" & rdoDadosProdu("PR_Complemento") & "', '" & rdoDadosProdu("PR_Sazonal") & "', '" & rdoDadosProdu("PR_CodigoProdutoNoFornecedor") & "', " & ConverteVirgula(IIf(IsNull(rdoDadosProdu("PR_PrecoVendaLiquido1")), 0, rdoDadosProdu("PR_PrecoVendaLiquido1"))) & ", " & ConverteVirgula(IIf(IsNull(rdoDadosProdu("PR_PrecoVendaLiquido2")), 0, rdoDadosProdu("PR_PrecoVendaLiquido2"))) & ", " & ConverteVirgula(IIf(IsNull(rdoDadosProdu("PR_PrecoVendaLiquido3")), 0, rdoDadosProdu("PR_PrecoVendaLiquido3"))) & ", '' )"
       
        
        SQL = "Insert into Produto (PR_Referencia,PR_CodigoFornecedor,PR_CodigoBarra,PR_Descricao,PR_DataCadastro,PR_Linha,PR_Secao,PR_Classe,PR_Bloqueio,PR_ClasseABC,PR_ClasseFiscal,PR_Unidade,PR_UnidadeDistribuicao,PR_PercentualComissao,PR_ICMSEntrada,PR_ICMSSaida,PR_AliquotaIPI, " _
            & "PR_CodigoIPI,PR_CodigoReducaoICMS,PR_PrecoFornecedor,PR_DescontoFornecedor,PR_ItemCondicoesGerais, PR_PercentualFrete,PR_PercentualEmbalagem,PR_CustoMedio1,PR_CustoMedio2,PR_CustoMedio3,PR_CustoMedioLiquido1,PR_CustoMedioLiquido2,PR_CustoMedioLiquido3,PR_PrecoCusto1,PR_PrecoCusto2," _
            & "PR_PrecoCusto3,PR_CustoLiquido1,PR_CustoLiquido2,PR_CustoLiquido3,PR_PrecoEntrada1,PR_PrecoEntrada2, PR_PrecoEntrada3,PR_DataPrecoCusto1,PR_DataPrecoCusto2,PR_DataPrecoCusto3,PR_PrecoVenda1,PR_PrecoVenda2, PR_PrecoVenda3,PR_DataPrecoVenda1,PR_DataPrecoVenda2,PR_DataPrecoVenda3,PR_PrecoVendaObjetivo,PR_PaginaListaPreco, " _
            & "PR_Peso,PR_MenorUnidadeCompra,PR_MetodoCompra,PR_TipoCalculoReposicao,PR_MetodoDistribuicao,PR_Residencia,PR_MargemObjetiva,PR_MargemPrevista,PR_Markup,PR_Comprador,PR_EmiteEtiqueta,PR_Situacao,PR_SubstituicaoTributaria, PR_DeducoesVenda,PR_IcmPdv,PR_DescricaoPDV,PR_Grupo,PR_Complemento,PR_Sazonal,PR_CodigoProdutoNoFornecedor, " _
            & "PR_PrecoVendaLiquido1,PR_PrecoVendaLiquido2,PR_PrecoVendaLiquido3,PR_HoraManutencao) Values ('" & rdoDadosProdu("PR_Referencia") & "'," _
            & rdoDadosProdu("PR_CodigoFornecedor") & ", '" & rdoDadosProdu("PR_CodigoBarra") & "', '" & rdoDadosProdu("PR_Descricao") & "', '" & Format(rdoDadosProdu("PR_DataCadastro"), "mm/dd/yyyy") & "', " & rdoDadosProdu("PR_Linha") & ", " & rdoDadosProdu("PR_Secao") & ",'" & rdoDadosProdu("PR_Classe") & "' , " _
            & "'" & rdoDadosProdu("PR_Bloqueio") & "', '" & rdoDadosProdu("PR_ClasseABC") & "', '" & rdoDadosProdu("PR_ClasseFiscal") & "', '" & rdoDadosProdu("PR_Unidade") & "', " & rdoDadosProdu("PR_UnidadeDistribuicao") & ", 2.00, " & ConverteVirgula(rdoDadosProdu("PR_ICMSEntrada")) & ", " & ConverteVirgula(rdoDadosProdu("PR_ICMSSaida")) & ", " _
            & ConverteVirgula(rdoDadosProdu("PR_AliquotaIPI")) & ", " & rdoDadosProdu("PR_CodigoIPI") & ", " & rdoDadosProdu("PR_CodigoReducaoICMS") & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoFornecedor")) & ", " & ConverteVirgula(rdoDadosProdu("PR_DescontoFornecedor")) & ", " & rdoDadosProdu("PR_ItemCondicoesGerais") & ", " & ConverteVirgula(rdoDadosProdu("PR_PercentualFrete")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PercentualEmbalagem")) & ", " _
            & ConverteVirgula(rdoDadosProdu("PR_CustoMedio1")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoMedio2")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoMedio3")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoMedioLiquido1")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoMedioLiquido2")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoMedioLiquido3")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoCusto1")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoCusto2")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoCusto3")) & ", " _
            & ConverteVirgula(rdoDadosProdu("PR_CustoLiquido1")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoLiquido2")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoLiquido3")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoEntrada1")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoEntrada2")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoEntrada3")) & ", '" & Format(rdoDadosProdu("PR_DataPrecoCusto1"), "mm/dd/yyyy") & "','" & Format(rdoDadosProdu("PR_DataPrecoCusto2"), "mm/dd/yyyy") & "', '" & Format(rdoDadosProdu("PR_DataPrecoCusto3"), "mm/dd/yyyy") & "', " & ConverteVirgula(rdoDadosProdu("PR_PrecoVenda1")) & ", " _
            & ConverteVirgula(rdoDadosProdu("PR_PrecoVenda2")) & "," & ConverteVirgula(rdoDadosProdu("PR_PrecoVenda3")) & ", '" & Format(rdoDadosProdu("PR_DataPrecoVenda1"), "mm/dd/yyyy") & "',' " & Format(rdoDadosProdu("PR_DataPrecoVenda2"), "mm/dd/yyyy") & "' ,'" & Format(rdoDadosProdu("PR_DataPrecoVenda3"), "mm/dd/yyyy") & "', " & ConverteVirgula(rdoDadosProdu("PR_PrecoVendaObjetivo")) & ", " & rdoDadosProdu("PR_PaginaListaPreco") & ", " & ConverteVirgula(rdoDadosProdu("PR_Peso")) & ", " & rdoDadosProdu("PR_MenorUnidadeCompra") & ", " & rdoDadosProdu("PR_MetodoCompra") & ", " _
            & rdoDadosProdu("PR_TipoCalculoReposicao") & " , " & rdoDadosProdu("PR_MetodoDistribuicao") & ", " & rdoDadosProdu("PR_Residencia") & ", " & ConverteVirgula(rdoDadosProdu("PR_MargemObjetiva")) & ", " & ConverteVirgula(rdoDadosProdu("PR_MargemPrevista")) & ", " & ConverteVirgula(rdoDadosProdu("PR_Markup")) & ", " & rdoDadosProdu("PR_Comprador") & ", '" & rdoDadosProdu("PR_EmiteEtiqueta") & "', '" & rdoDadosProdu("PR_Situacao") & "', '" & rdoDadosProdu("PR_SubstituicaoTributaria") & "', " _
            & ConverteVirgula(rdoDadosProdu("PR_DeducoesVenda")) & ", " & ConverteVirgula(rdoDadosProdu("PR_IcmPdv")) & ", '" & rdoDadosProdu("PR_DescricaoPDV") & "', " & rdoDadosProdu("PR_Grupo") & ", '" & rdoDadosProdu("PR_Complemento") & "', '" & rdoDadosProdu("PR_Sazonal") & "', '" & rdoDadosProdu("PR_CodigoProdutoNoFornecedor") & "', " & ConverteVirgula(IIf(IsNull(rdoDadosProdu("PR_PrecoVendaLiquido1")), 0, rdoDadosProdu("PR_PrecoVendaLiquido1"))) & ", " & ConverteVirgula(IIf(IsNull(rdoDadosProdu("PR_PrecoVendaLiquido2")), 0, rdoDadosProdu("PR_PrecoVendaLiquido2"))) & ", " & ConverteVirgula(IIf(IsNull(rdoDadosProdu("PR_PrecoVendaLiquido3")), 0, rdoDadosProdu("PR_PrecoVendaLiquido3"))) & ", '' )"
        rdoCnLoja.Execute (SQL)
        If Err.Number = 0 Then
            CommitTrans
            GravaEstoqueLoja (rdoDadosProdu("PR_Referencia"))
'            On Error Resume Next
'            BeginTrans
'            SQL = "insert into ProduLjDBF (DESCRICAO,ORCODPRO,ALIQIPI,CODIPI,CONTROLE,TRIBUTO,VENVAR1,CLASSFISC,UNIDADE, " _
'                & "PRECUS1,PROMO,BCOMIS,CSPROD,LINHA,SECAO,FORNECEDOR,TIPO,PESO,PAG,SUBTRIBUT,ICMPDV,CODBARRA,SITUACAO) " _
'                & "Values('" & rdoDadosProdu("PR_Descricao") & "', '" & rdoDadosProdu("PR_Referencia") & "', " & rdoDadosProdu("PR_AliquotaIPI") & ", " & rdoDadosProdu("PR_CodigoIPI") & ", " & rdoDadosProdu("PR_CodigoReducaoICMS") & ", " _
'                & "" & rdoDadosProdu("PR_ICMSSaida") & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoVenda1")) & ", '" & rdoDadosProdu("PR_ClasseFiscal") & "', '" & rdoDadosProdu("PR_Unidade") & "', " & ConverteVirgula(Format(rdoDadosProdu("PR_PrecoCusto1"), "0.00")) & ", " _
'                & "1, 0, " & rdoDadosProdu("PR_Bloqueio") & ", " & rdoDadosProdu("PR_Linha") & ", " & rdoDadosProdu("PR_Secao") & ", " & rdoDadosProdu("PR_CodigoFornecedor") & ", '" & rdoDadosProdu("PR_Classe") & "', " & ConverteVirgula(Format(rdoDadosProdu("PR_Peso"), "0.000")) & ", " & rdoDadosProdu("PR_PaginaListaPreco") & ", '" & IIf(IsNull(rdoDadosProdu("PR_SubstituicaoTributaria")), "N", rdoDadosProdu("PR_SubstituicaoTributaria")) & "', " _
'                & "" & ConverteVirgula(Format(rdoDadosProdu("PR_IcmPdv"), "0.00")) & ", '" & rdoDadosProdu("PR_CodigoBarra") & "', '" & rdoDadosProdu("PR_Situacao") & "')"
'                rdoCnLoja.Execute (SQL)
'            If Err.Number = 0 Then
'                CommitTrans
'            Else
'                Rollback
'            End If
            MsgBox "Referencia Gravada com sucesso", vbInformation, "Sucesso"
        Else
            Rollback
            MsgBox "Não foi possivel gravar a referencia, Verifique se você esta conectado a internet", vbCritical, "Atenção"
        End If
    Else
        Screen.MousePointer = 0
        MsgBox "Referencia não encontrada", vbCritical, "Aviso"
    End If
    ConexaoBach.Close
    Screen.MousePointer = 0
End Function

Function GravaCodigoBarras(ByVal Referencia As String)
    Dim rdoDadosProdu As rdoResultset
    
    Screen.MousePointer = 11
    On Error Resume Next
    SQL = ""
    SQL = "Select * from ProdutoBarras where PRB_CodigoBarras='" & UCase(Referencia) & "'"
        Set rdoDadosProdu = Conexao.OpenResultset(SQL)
    If Not rdoDadosProdu.EOF Then
        On Error Resume Next
        
        SQL = ""
        SQL = "Delete ProdutoBarras where PRB_CodigoBarras='" & UCase(Referencia) & "' "
            rdoCnLoja.Execute (SQL)
        
        BeginTrans
        
        'SQL = "Insert into ProdutoBarras (PRB_Referencia,PRB_CodigoBarras,PRB_CodigoFornecedor,PRB_Embalagem,PRB_HoraManutencao,PRB_TipoCodigo) Values ('" & rdoDadosProdu("PRB_Referencia") & "'," _
            & rdoDadosProdu("PRB_CodigoBarras") & ", '" & rdoDadosProdu("PRB_CodigoFornecedor") & "', " & rdoDadosProdu("PRB_Embalagem") & ", '" & Format(rdoDadosProdu("PRB_HoraManutencao"), "mm/dd/yyyy hh:mm:ss") & "', '" & rdoDadosProdu("PRB_TipoCodigo") & "')"
       
        
        SQL = "Insert into ProdutoBarras (PRB_Referencia,PRB_CodigoBarras,PRB_CodigoFornecedor,PRB_Embalagem,PRB_HoraManutencao,PRB_TipoCodigo) Values ('" & rdoDadosProdu("PRB_Referencia") & "','" _
            & rdoDadosProdu("PRB_CodigoBarras") & "', '" & rdoDadosProdu("PRB_CodigoFornecedor") & "', " & rdoDadosProdu("PRB_Embalagem") & ", '" & Format(rdoDadosProdu("PRB_HoraManutencao"), "mm/dd/yyyy hh:mm:ss") & "', '" & rdoDadosProdu("PRB_TipoCodigo") & "')"
        rdoCnLoja.Execute (SQL)
        If Err.Number = 0 Then
            CommitTrans
            MsgBox "Código de Barras Gravado com sucesso", vbInformation, "Sucesso"
        Else
            Rollback
            MsgBox "Não foi possivel gravar a referencia, Verifique se você esta conectado a internet", vbCritical, "Atenção"
        End If
    Else
        Screen.MousePointer = 0
        MsgBox "Código de Barras não encontrado", vbCritical, "Aviso"
    End If
    Screen.MousePointer = 0
End Function

Function GravaFornecedor(ByVal Forne As Integer)
    Dim rdoProduBarra As rdoResultset
    Dim rdoDadosProdu As rdoResultset
    
    On Error Resume Next
    If ConectaODBC(ConexaoBach, "sa", "jeda36") = False Then
        MsgBox "Erro conectando ao banco de dados.", vbCritical, "Atenção"
        Exit Function
    End If
    
    SQL = ""
    SQL = "Select * From Produto Where PR_CodigoFornecedor = " & Forne & ""
    Set rdoDadosProdu = ConexaoBach.OpenResultset(SQL)
    
    If Not rdoDadosProdu.EOF Then
        Do While Not rdoDadosProdu.EOF
            SQL = ""
            SQL = "Delete Produto where Pr_referencia='" & rdoDadosProdu("PR_Referencia") & "' "
                rdoCnLoja.Execute (SQL)
            
            Err.Number = 0
            
            rdoCnLoja.BeginTrans
            'SQL = "Insert into Produto (PR_Referencia,PR_CodigoFornecedor,PR_CodigoBarra,PR_Descricao,PR_DataCadastro,PR_Linha,PR_Secao,PR_Classe,PR_Bloqueio,PR_ClasseABC,PR_ClasseFiscal,PR_Unidade,PR_UnidadeDistribuicao,PR_PercentualComissao,PR_ICMSEntrada,PR_ICMSSaida,PR_AliquotaIPI, " _
                & "PR_CodigoIPI,PR_CodigoReducaoICMS,PR_PrecoFornecedor,PR_DescontoFornecedor,PR_ItemCondicoesGerais, PR_PercentualFrete,PR_PercentualEmbalagem,PR_CustoMedio1,PR_CustoMedio2,PR_CustoMedio3,PR_CustoMedioLiquido1,PR_CustoMedioLiquido2,PR_CustoMedioLiquido3,PR_PrecoCusto1,PR_PrecoCusto2," _
                & "PR_PrecoCusto3,PR_CustoLiquido1,PR_CustoLiquido2,PR_CustoLiquido3,PR_PrecoEntrada1,PR_PrecoEntrada2, PR_PrecoEntrada3,PR_DataPrecoCusto1,PR_DataPrecoCusto2,PR_DataPrecoCusto3,PR_PrecoVenda1,PR_PrecoVenda2, PR_PrecoVenda3,PR_DataPrecoVenda1,PR_DataPrecoVenda2,PR_DataPrecoVenda3,PR_PrecoVendaObjetivo,PR_PaginaListaPreco, " _
                & "PR_Peso,PR_MenorUnidadeCompra,PR_MetodoCompra,PR_TipoCalculoReposicao,PR_MetodoDistribuicao,PR_Residencia,PR_MargemObjetiva,PR_MargemPrevista,PR_Markup,PR_Comprador,PR_EmiteEtiqueta,PR_Situacao,PR_SubstituicaoTributaria, PR_DeducoesVenda,PR_IcmPdv,PR_DescricaoPDV,PR_Grupo,PR_Complemento,PR_Sazonal,PR_CodigoProdutoNoFornecedor, " _
                & "PR_PrecoVendaLiquido1,PR_PrecoVendaLiquido2,PR_PrecoVendaLiquido3,PR_HoraManutencao) Values ('" & rdoDadosProdu("PR_Referencia") & "'," _
                & rdoDadosProdu("PR_CodigoFornecedor") & ", '" & rdoDadosProdu("PR_CodigoBarra") & "', '" & rdoDadosProdu("PR_Descricao") & "', '" & Format(rdoDadosProdu("PR_DataCadastro"), "mm/dd/yyyy") & "', " & rdoDadosProdu("PR_Linha") & ", " & rdoDadosProdu("PR_Secao") & ",'" & rdoDadosProdu("PR_Classe") & "' , " _
                & "'" & rdoDadosProdu("PR_Bloqueio") & "', '" & rdoDadosProdu("PR_ClasseABC") & "', '" & rdoDadosProdu("PR_ClasseFiscal") & "', '" & rdoDadosProdu("PR_Unidade") & "', " & rdoDadosProdu("PR_UnidadeDistribuicao") & ", 2.00, " & ConverteVirgula(rdoDadosProdu("PR_ICMSEntrada")) & ", " & ConverteVirgula(rdoDadosProdu("PR_ICMSSaida")) & ", " _
                & ConverteVirgula(rdoDadosProdu("PR_AliquotaIPI")) & ", " & rdoDadosProdu("PR_CodigoIPI") & ", " & rdoDadosProdu("PR_CodigoReducaoICMS") & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoFornecedor")) & ", " & ConverteVirgula(rdoDadosProdu("PR_DescontoFornecedor")) & ", " & rdoDadosProdu("PR_ItemCondicoesGerais") & ", " & ConverteVirgula(rdoDadosProdu("PR_PercentualFrete")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PercentualEmbalagem")) & ", " _
                & ConverteVirgula(rdoDadosProdu("PR_CustoMedio1")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoMedio2")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoMedio3")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoMedioLiquido1")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoMedioLiquido2")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoMedioLiquido3")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoCusto1")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoCusto2")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoCusto3")) & ", " _
                & ConverteVirgula(rdoDadosProdu("PR_CustoLiquido1")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoLiquido2")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoLiquido3")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoEntrada1")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoEntrada2")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoEntrada3")) & ", '" & Format(rdoDadosProdu("PR_DataPrecoCusto1"), "mm/dd/yyyy") & "','" & Format(rdoDadosProdu("PR_DataPrecoCusto2"), "mm/dd/yyyy") & "', '" & Format(rdoDadosProdu("PR_DataPrecoCusto3"), "mm/dd/yyyy") & "', " & ConverteVirgula(rdoDadosProdu("PR_PrecoVenda1")) & ", " _
                & ConverteVirgula(rdoDadosProdu("PR_PrecoVenda2")) & "," & ConverteVirgula(rdoDadosProdu("PR_PrecoVenda3")) & ", '" & Format(rdoDadosProdu("PR_DataPrecoVenda1"), "mm/dd/yyyy") & "',' " & Format(rdoDadosProdu("PR_DataPrecoVenda2"), "mm/dd/yyyy") & "' ,'" & Format(rdoDadosProdu("PR_DataPrecoVenda3"), "mm/dd/yyyy") & "', " & ConverteVirgula(rdoDadosProdu("PR_PrecoVendaObjetivo")) & ", " & rdoDadosProdu("PR_PaginaListaPreco") & ", " & ConverteVirgula(rdoDadosProdu("PR_Peso")) & ", " & rdoDadosProdu("PR_MenorUnidadeCompra") & ", " & rdoDadosProdu("PR_MetodoCompra") & ", " _
                & rdoDadosProdu("PR_TipoCalculoReposicao") & " , " & rdoDadosProdu("PR_MetodoDistribuicao") & ", " & rdoDadosProdu("PR_Residencia") & ", " & ConverteVirgula(rdoDadosProdu("PR_MargemObjetiva")) & ", " & ConverteVirgula(rdoDadosProdu("PR_MargemPrevista")) & ", " & ConverteVirgula(rdoDadosProdu("PR_Markup")) & ", " & rdoDadosProdu("PR_Comprador") & ", '" & rdoDadosProdu("PR_EmiteEtiqueta") & "', '" & rdoDadosProdu("PR_Situacao") & "', '" & rdoDadosProdu("PR_SubstituicaoTributaria") & "', " _
                & ConverteVirgula(rdoDadosProdu("PR_DeducoesVenda")) & ", " & ConverteVirgula(rdoDadosProdu("PR_IcmPdv")) & ", '" & rdoDadosProdu("PR_DescricaoPDV") & "', " & rdoDadosProdu("PR_Grupo") & ", '" & rdoDadosProdu("PR_Complemento") & "', '" & rdoDadosProdu("PR_Sazonal") & "', '" & rdoDadosProdu("PR_CodigoProdutoNoFornecedor") & "', " & ConverteVirgula(IIf(IsNull(rdoDadosProdu("PR_PrecoVendaLiquido1")), 0, rdoDadosProdu("PR_PrecoVendaLiquido1"))) & ", " & ConverteVirgula(IIf(IsNull(rdoDadosProdu("PR_PrecoVendaLiquido2")), 0, rdoDadosProdu("PR_PrecoVendaLiquido2"))) & ", " & ConverteVirgula(IIf(IsNull(rdoDadosProdu("PR_PrecoVendaLiquido3")), 0, rdoDadosProdu("PR_PrecoVendaLiquido3"))) & ", '' )"
           
            
            SQL = "Insert into Produto (PR_Referencia,PR_CodigoFornecedor,PR_CodigoBarra,PR_Descricao,PR_DataCadastro,PR_Linha,PR_Secao,PR_Classe,PR_Bloqueio,PR_ClasseABC,PR_ClasseFiscal,PR_Unidade,PR_UnidadeDistribuicao,PR_PercentualComissao,PR_ICMSEntrada,PR_ICMSSaida,PR_AliquotaIPI, " _
                & "PR_CodigoIPI,PR_CodigoReducaoICMS,PR_PrecoFornecedor,PR_DescontoFornecedor,PR_ItemCondicoesGerais, PR_PercentualFrete,PR_PercentualEmbalagem,PR_CustoMedio1,PR_CustoMedio2,PR_CustoMedio3,PR_CustoMedioLiquido1,PR_CustoMedioLiquido2,PR_CustoMedioLiquido3,PR_PrecoCusto1,PR_PrecoCusto2," _
                & "PR_PrecoCusto3,PR_CustoLiquido1,PR_CustoLiquido2,PR_CustoLiquido3,PR_PrecoEntrada1,PR_PrecoEntrada2, PR_PrecoEntrada3,PR_DataPrecoCusto1,PR_DataPrecoCusto2,PR_DataPrecoCusto3,PR_PrecoVenda1,PR_PrecoVenda2, PR_PrecoVenda3,PR_DataPrecoVenda1,PR_DataPrecoVenda2,PR_DataPrecoVenda3,PR_PrecoVendaObjetivo,PR_PaginaListaPreco, " _
                & "PR_Peso,PR_MenorUnidadeCompra,PR_MetodoCompra,PR_TipoCalculoReposicao,PR_MetodoDistribuicao,PR_Residencia,PR_MargemObjetiva,PR_MargemPrevista,PR_Markup,PR_Comprador,PR_EmiteEtiqueta,PR_Situacao,PR_SubstituicaoTributaria, PR_DeducoesVenda,PR_IcmPdv,PR_DescricaoPDV,PR_Grupo,PR_Complemento,PR_Sazonal,PR_CodigoProdutoNoFornecedor, " _
                & "PR_PrecoVendaLiquido1,PR_PrecoVendaLiquido2,PR_PrecoVendaLiquido3,PR_HoraManutencao) Values ('" & rdoDadosProdu("PR_Referencia") & "'," _
                & rdoDadosProdu("PR_CodigoFornecedor") & ", '" & rdoDadosProdu("PR_CodigoBarra") & "', '" & rdoDadosProdu("PR_Descricao") & "', '" & Format(rdoDadosProdu("PR_DataCadastro"), "mm/dd/yyyy") & "', " & rdoDadosProdu("PR_Linha") & ", " & rdoDadosProdu("PR_Secao") & ",'" & rdoDadosProdu("PR_Classe") & "' , " _
                & "'" & rdoDadosProdu("PR_Bloqueio") & "', '" & rdoDadosProdu("PR_ClasseABC") & "', '" & rdoDadosProdu("PR_ClasseFiscal") & "', '" & rdoDadosProdu("PR_Unidade") & "', " & rdoDadosProdu("PR_UnidadeDistribuicao") & ", 2.00, " & ConverteVirgula(rdoDadosProdu("PR_ICMSEntrada")) & ", " & ConverteVirgula(rdoDadosProdu("PR_ICMSSaida")) & ", " _
                & ConverteVirgula(rdoDadosProdu("PR_AliquotaIPI")) & ", " & rdoDadosProdu("PR_CodigoIPI") & ", " & rdoDadosProdu("PR_CodigoReducaoICMS") & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoFornecedor")) & ", " & ConverteVirgula(rdoDadosProdu("PR_DescontoFornecedor")) & ", " & rdoDadosProdu("PR_ItemCondicoesGerais") & ", " & ConverteVirgula(rdoDadosProdu("PR_PercentualFrete")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PercentualEmbalagem")) & ", " _
                & ConverteVirgula(rdoDadosProdu("PR_CustoMedio1")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoMedio2")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoMedio3")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoMedioLiquido1")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoMedioLiquido2")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoMedioLiquido3")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoCusto1")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoCusto2")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoCusto3")) & ", " _
                & ConverteVirgula(rdoDadosProdu("PR_CustoLiquido1")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoLiquido2")) & ", " & ConverteVirgula(rdoDadosProdu("PR_CustoLiquido3")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoEntrada1")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoEntrada2")) & ", " & ConverteVirgula(rdoDadosProdu("PR_PrecoEntrada3")) & ", '" & Format(rdoDadosProdu("PR_DataPrecoCusto1"), "mm/dd/yyyy") & "','" & Format(rdoDadosProdu("PR_DataPrecoCusto2"), "mm/dd/yyyy") & "', '" & Format(rdoDadosProdu("PR_DataPrecoCusto3"), "mm/dd/yyyy") & "', " & ConverteVirgula(rdoDadosProdu("PR_PrecoVenda1")) & ", " _
                & ConverteVirgula(rdoDadosProdu("PR_PrecoVenda2")) & "," & ConverteVirgula(rdoDadosProdu("PR_PrecoVenda3")) & ", '" & Format(rdoDadosProdu("PR_DataPrecoVenda1"), "mm/dd/yyyy") & "',' " & Format(rdoDadosProdu("PR_DataPrecoVenda2"), "mm/dd/yyyy") & "' ,'" & Format(rdoDadosProdu("PR_DataPrecoVenda3"), "mm/dd/yyyy") & "', " & ConverteVirgula(rdoDadosProdu("PR_PrecoVendaObjetivo")) & ", " & rdoDadosProdu("PR_PaginaListaPreco") & ", " & ConverteVirgula(rdoDadosProdu("PR_Peso")) & ", " & rdoDadosProdu("PR_MenorUnidadeCompra") & ", " & rdoDadosProdu("PR_MetodoCompra") & ", " _
                & rdoDadosProdu("PR_TipoCalculoReposicao") & " , " & rdoDadosProdu("PR_MetodoDistribuicao") & ", " & rdoDadosProdu("PR_Residencia") & ", " & ConverteVirgula(rdoDadosProdu("PR_MargemObjetiva")) & ", " & ConverteVirgula(rdoDadosProdu("PR_MargemPrevista")) & ", " & ConverteVirgula(rdoDadosProdu("PR_Markup")) & ", " & rdoDadosProdu("PR_Comprador") & ", '" & rdoDadosProdu("PR_EmiteEtiqueta") & "', '" & rdoDadosProdu("PR_Situacao") & "', '" & rdoDadosProdu("PR_SubstituicaoTributaria") & "', " _
                & ConverteVirgula(rdoDadosProdu("PR_DeducoesVenda")) & ", " & ConverteVirgula(rdoDadosProdu("PR_IcmPdv")) & ", '" & rdoDadosProdu("PR_DescricaoPDV") & "', " & rdoDadosProdu("PR_Grupo") & ", '" & rdoDadosProdu("PR_Complemento") & "', '" & rdoDadosProdu("PR_Sazonal") & "', '" & rdoDadosProdu("PR_CodigoProdutoNoFornecedor") & "', " & ConverteVirgula(IIf(IsNull(rdoDadosProdu("PR_PrecoVendaLiquido1")), 0, rdoDadosProdu("PR_PrecoVendaLiquido1"))) & ", " & ConverteVirgula(IIf(IsNull(rdoDadosProdu("PR_PrecoVendaLiquido2")), 0, rdoDadosProdu("PR_PrecoVendaLiquido2"))) & ", " & ConverteVirgula(IIf(IsNull(rdoDadosProdu("PR_PrecoVendaLiquido3")), 0, rdoDadosProdu("PR_PrecoVendaLiquido3"))) & ", '' )"
            rdoCnLoja.Execute (SQL)
            
            If Err.Number = 0 Then
                rdoCnLoja.CommitTrans
                GravaEstoqueLoja (rdoDadosProdu("PR_Referencia"))
            Else
                rdoCnLoja.RollbackTrans
            End If
            
            SQL = ""
            SQL = "Delete ProdutoBarras where PRB_CodigoBarras='" & rdoDadosProdu("PR_Referencia") & "' "
                rdoCnLoja.Execute (SQL)
            
            SQL = ""
            SQL = "Select * From ProdutoBarras Where PRB_Referencia = '" & rdoDadosProdu("PR_Referencia") & "' and PRB_TipoCodigo Not In ('F')"
            Set rdoProduBarra = ConexaoBach.OpenResultset(SQL)
            
            If Not rdoProduBarra.EOF Then
                Do While Not rdoProduBarra.EOF
                    rdoCnLoja.BeginTrans
                     
                    'SQL = "Insert into ProdutoBarras (PRB_Referencia,PRB_CodigoBarras,PRB_CodigoFornecedor,PRB_Embalagem,PRB_HoraManutencao,PRB_TipoCodigo) Values ('" & rdoProduBarra("PRB_Referencia") & "'," _
                         & rdoProduBarra("PRB_CodigoBarras") & "', '" & rdoProduBarra("PRB_CodigoFornecedor") & "', " & rdoProduBarra("PRB_Embalagem") & ", '" & Format(rdoProduBarra("PRB_HoraManutencao"), "mm/dd/yyyy hh:mm:ss") & "', '" & rdoProduBarra("PRB_TipoCodigo") & "')"
                     
                    SQL = "Insert into ProdutoBarras (PRB_Referencia,PRB_CodigoBarras,PRB_CodigoFornecedor,PRB_Embalagem,PRB_HoraManutencao,PRB_TipoCodigo) Values ('" & rdoProduBarra("PRB_Referencia") & "','" _
                        & rdoProduBarra("PRB_CodigoBarras") & "', '" & rdoProduBarra("PRB_CodigoFornecedor") & "', " & rdoProduBarra("PRB_Embalagem") & ", '" & Format(rdoProduBarra("PRB_HoraManutencao"), "mm/dd/yyyy hh:mm:ss") & "', '" & rdoProduBarra("PRB_TipoCodigo") & "')"
                    rdoCnLoja.Execute (SQL)
                     
                    If Err.Number = 0 Then
                        rdoCnLoja.CommitTrans
                    Else
                        rdoCnLoja.RollbackTrans
                    End If
                
                    rdoProduBarra.MoveNext
                Loop
            End If
            rdoDadosProdu.MoveNext
        Loop
    Else
        MsgBox "Fornecedor não cadastrado.", vbCritical, "Atenção"
        Exit Function
    End If
    
    
    ConexaoBach.Close
    
End Function

Function GravaEstoqueLoja(ByVal Referencia As String)
    Dim rdoEstoque As rdoResultset
    
    On Error Resume Next
    
    SQL = ""
    SQL = "Select * From EstoqueLoja Where EL_Referencia = '" & Referencia & "'"
    Set rdoEstoque = rdoCnLoja.OpenResultset(SQL)
    
    If rdoEstoque.EOF Then
        rdoCnLoja.BeginTrans
        SQL = ""
        SQL = "Insert into EstoqueLoja (EL_Referencia,EL_Loja,EL_Estoque,EL_EstoqueAnterior) " _
            & "Values ('" & Referencia & "','" & AchaLojaControle & "',0,0)"
            rdoCnLoja.Execute (SQL)
        If Err.Number = 0 Then
            rdoCnLoja.CommitTrans
        Else
            rdoCnLoja.RollbackTrans
        End If
    End If

End Function


Function ComparaEstoqueAnterior() As Boolean
    Dim rsCompEstAnt As rdoResultset

    SQL = ""
    SQL = "Select EL_Referencia from EstoqueLoja " _
        & "where EL_Estoque<>EL_EstoqueAnterior"
    Set rsCompEstAnt = rdoCnLoja.OpenResultset(SQL)
    If Not rsCompEstAnt.EOF Then
        ComparaEstoqueAnterior = False
    Else
        ComparaEstoqueAnterior = True
    End If
    
    
End Function






