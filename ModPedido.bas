Attribute VB_Name = "ModPedido"
Option Explicit
Global wDescontoFechamento As Boolean
Global wPreencherCliente As Boolean
Global wNumeroClientePedido As Double


Function CalculaDesconto(ByVal TipoDesconto As Integer, ByVal ValorDesconto As Double, ByVal ValorItem As Double) As Double

    Dim wDesconto As Double

'
'--------------------------------------Calcula Desconto Por %(1)---------------------------
'
    
    If TipoDesconto = 1 Then
        wDesconto = (ValorItem * ValorDesconto) / 100
        ValorItem = ValorItem - wDesconto
        CalculaDesconto = ValorItem
    End If
    
'
'---------------------------------------Calcula Desconto Por - (2) -------------------------
'
    If TipoDesconto = 2 Then
        ValorItem = ValorItem - ValorDesconto
        CalculaDesconto = Format(ValorItem, "0.00")
    End If
        

End Function

Function DeletaItensPedido(ByVal TipoDelecao As Integer, ByVal Referencia As String, ByVal NumeroPedido As Double, ByVal Item As Integer)

'
'----------------------------------Deleta Itens Pedido (1)--------------------------------
'
    If TipoDelecao = 1 Then
        SQL = ""
        SQL = "Delete NfItens " _
            & "where NumeroPed=" & NumeroPedido & " " _
            & "and Referencia='" & Referencia & "' " _
            & "and Item = " & Item & ""
            rdoCnLoja.Execute (SQL)
        
'
'-----------------------------------Deleta Pedido Inteiro (2) ------------------------------
'

    ElseIf TipoDelecao = 2 Then
        SQL = ""
        SQL = "Delete NfItens " _
        & "where NumeroPed=" & NumeroPedido & " " _
        & " and tiponota = 'PD'"
            rdoCnLoja.Execute (SQL)
    
        SQL = ""
        SQL = "Delete NfCapa " _
            & "where NumeroPed=" & NumeroPedido & " " _
            & " and tiponota = 'PD'"
            rdoCnLoja.Execute (SQL)
            
    End If

End Function

Function PesquisaCliente(ByVal TipoPesquisa As Integer, ByVal Cliente As String, ByRef NomerdoResultset) As Boolean

'
'--------------------------------Pesquisa Pelo Codigo do Cliente (1)-------------------------
'
    DescricaoOperacao "Pesquisando Cliente"
    If TipoPesquisa = 1 Then
        SQL = ""
        SQL = "Select CE_Razao ,CE_CodigoCliente from Cliente " _
            & "where CE_CodigoCliente = " & Cliente & " "
            Set NomerdoResultset = rdoCnLoja.OpenResultset(SQL)

'
'-------------------------------Pesquisa por cgc ou cpf (2) ---------------------------------
'
    ElseIf TipoPesquisa = 2 Then
        SQL = ""
        SQL = ""
        SQL = "Select CE_Razao ,CE_CodigoCliente from Cliente " _
            & "where CE_Cgc = '" & Cliente & "' "
            Set NomerdoResultset = rdoCnLoja.OpenResultset(SQL)
    
'
'-------------------------------Pesquisa Pelo Nome Cliente (3) ---------------------------------
'
    ElseIf TipoPesquisa = 3 Then
        SQL = ""
        SQL = ""
        SQL = "Select CE_razao,CE_CodigoCliente from Cliente " _
            & "where CE_Razao like '" & UCase(Cliente) & "%' order by CE_Razao"
            Set NomerdoResultset = rdoCnLoja.OpenResultset(SQL)
    
'
'-------------------------------Pesquisa Cliente Tela frmCadCliente(4) --------------------------
'
    ElseIf TipoPesquisa = 4 Then
        SQL = ""
        SQL = ""
        SQL = "Select * from Cliente " _
            & "where CE_CodigoCliente = " & Cliente & " order by CE_CodigoCliente"
            Set NomerdoResultset = rdoCnLoja.OpenResultset(SQL)
    
    Else
        Exit Function
    End If
    If Not NomerdoResultset.EOF Then
        PesquisaCliente = True
    Else
        PesquisaCliente = False
    End If
    DescricaoOperacao "Pronto"
    
End Function

Sub CriaCotacaoHtml(ByVal Pedido As Double)
    Dim IntFile1
    Dim NomeArquivo As String
    Dim rsInfLoja As rdoResultset
    Dim InfLoja As String
    Dim rsPedido As rdoResultset
    Dim Logo  As String
    Dim SomaPedido As Double
    Dim rdoCliente As rdoResultset
    Dim rdoValidadeCotacao As rdoResultset
    Dim Vendedor As String
    Dim CondPag As String
    Dim Razao As String
    Dim CepLoja As String
    Dim UfLoja As String
    Dim TelefoneLoja As String
    Dim EndLoja As String
    Dim BairroLoja As String
    
    CepLoja = ""
    UfLoja = ""
    TelefoneLoja = ""
    EndLoja = ""
    BairroLoja = ""
    Razao = ""
    
    SQL = ""
    SQL = "Select CT_ValidadeCotacao From Controle"
    Set rdoValidadeCotacao = rdoCnLojaBach.OpenResultset(SQL)
    
    SQL = ""
    SQL = "Select * from Lojas where LO_Loja='" & AchaLojaControle & "'"
    Set rsInfLoja = rdoCnLoja.OpenResultset(SQL)
    If Not rsInfLoja.EOF Then
        CepLoja = rsInfLoja("LO_Cep")
        UfLoja = rsInfLoja("LO_UF")
        TelefoneLoja = rsInfLoja("LO_Telefone")
        EndLoja = rsInfLoja("LO_Endereco")
        BairroLoja = rsInfLoja("LO_Bairro")
        Razao = rsInfLoja("LO_Razao")
    End If
    NomeArquivo = "Cot" & Pedido
    

    IntFile1 = FreeFile()
    Open GLB_Cotacao & NomeArquivo & ".html" For Output Access Write As #IntFile1
     
    Print #IntFile1, "<html>"
    Print #IntFile1, "<head>"
    Print #IntFile1, "<title>" & Razao & "</title>"
    Print #IntFile1, "</head>"
    
    SQL = ""
    SQL = "Select CE_Razao,CE_Endereco,CE_Cep,CE_Bairro,CE_Municipio,CE_Estado,CE_Telefone from Cliente,NfCapa " _
        & "where CE_CodigoCliente=Cliente and NumeroPed=" & Pedido & ""
    Set rdoCliente = rdoCnLojaBach.OpenResultset(SQL)
    If Not rdoCliente.EOF Then
        Print #IntFile1, "<table border=0 cellpadding=0 cellspacing=0 width=750 height=39>"
        Print #IntFile1, "<tr>"
        If AchaLojaControle = "85" Then
            Print #IntFile1, "<td width=8 rowspan=3 height=1><img border=0 src=http://www.afgferramentas.com.br/images/logoafg.jpg width=245 height=60 align=left>"
        ElseIf AchaLojaControle = "314" Then
            Print #IntFile1, "<td width=8 rowspan=3 height=1><img border=0 src=http://www.dmmotores.com.br/images/logodm.jpg width=245 height=60 align=left>"
        Else
            Print #IntFile1, "<td width=8 rowspan=3 height=1><img border=0 src=http://www.demeo.com.br/images/logodemeo.jpg width=245 height=60 align=left>"
        End If
        Print #IntFile1, "</td>"
        Print #IntFile1, "<td width=333 height=1>"
        Print #IntFile1, "<font face=System color=#000080>" & UCase(rdoCliente("CE_Razao")) & " </font>"
        Print #IntFile1, "</td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=333 height=0>"
        Print #IntFile1, "<font face=System color=#000080>" & UCase(rdoCliente("CE_Endereco")) & " - " & UCase(rdoCliente("CE_Bairro")) & "</font>"
        Print #IntFile1, "</td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=333 height=0>"
        Print #IntFile1, "<font face=System color=#000080>" & UCase(Right(String(8, "0") & rdoCliente("CE_CEP"), 8)) & " - " & UCase(rdoCliente("CE_Municipio")) & " - " & UCase(rdoCliente("CE_estado")) & " - Fone: " & UCase(rdoCliente("CE_Telefone")) & " </font>"
        Print #IntFile1, "</td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "</table>"
        Print #IntFile1, "<div align=justify style=width: 750; height: 56>"
        Print #IntFile1, "<div align=justify>"
        Print #IntFile1, "<table border=0 cellpadding=0 cellspacing=0 width=336 height=41>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=334 height=21><font size=1 color=#000080 face=Arial>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & EndLoja & " - " & BairroLoja & "</font></td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=334 height=20><font face=Arial size=1 color=#000080>SAO PAULO - " & UfLoja & " - CEP : " & UCase(Right(String(8, "0") & CepLoja, 8)) & " - FONE : " & TelefoneLoja & "</font></td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "</table>"
        Print #IntFile1, "</div>"
        Print #IntFile1, "<table border=0 cellpadding=0 cellspacing=0 width=750 height=48>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=700 rowspan=3><img border=0 src=http://www.demeo.com.br/images/truck.jpg width=750 height=40 align=right></td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "</table>"
        Print #IntFile1, "</div>"
        Print #IntFile1, "<div align=left>"
        Print #IntFile1, "<table border=1 cellpadding=1 cellspacing=0 width=750 height=36 solid; border-width: 0 bordercolor=#000080>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=640 height=20><font face=Arial color=#000080><b>Referencia / Descrição</b></font></td>"
        Print #IntFile1, "<td width=41 height=20><font face=Arial color=#000080><b>Qtde</b></font></td>"
        Print #IntFile1, "<td width=76 height=20><font face=Arial color=#000080><b>Valor</b></font></td>"
        Print #IntFile1, "<td width=66 height=20><font face=Arial color=#000080><b>Desconto</b></font></td>"
        Print #IntFile1, "<td width=79 height=20><font face=Arial color=#000080><b>Total</b></font></td>"
        Print #IntFile1, "</tr>"
    Else
        Print #IntFile1, "<table border=0 cellpadding=0 cellspacing=0 width=750 height=39>"
        Print #IntFile1, "<tr>"
        If AchaLojaControle = "85" Then
            Print #IntFile1, "<td width=8 rowspan=3 height=1><img border=0 src=http://www.afgferramentas.com.br/images/logoafg.jpg width=245 height=60 align=left>"
        ElseIf AchaLojaControle = "314" Then
            Print #IntFile1, "<td width=8 rowspan=3 height=1><img border=0 src=http://www.dmmotores.com.br/images/logodm.jpg width=245 height=60 align=left>"
        Else
            Print #IntFile1, "<td width=8 rowspan=3 height=1><img border=0 src=http://www.demeo.com.br/images/logodemeo.jpg width=245 height=60 align=left>"
        End If
        Print #IntFile1, "</td>"
        Print #IntFile1, "<td width=333 height=1>"
        Print #IntFile1, "<font face=System color=#000080>CONSUMIDOR</font>"
        Print #IntFile1, "</td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=333 height=0>"
        Print #IntFile1, "<font face=System color=#000080>CONSUMIDOR</font>"
        Print #IntFile1, "</td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=333 height=0>"
        Print #IntFile1, "<font face=System color=#000080>00000000 - CONSUMIDOR </font>"
        Print #IntFile1, "</td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "</table>"
        Print #IntFile1, "<div align=justify style=width: 750; height: 56>"
        Print #IntFile1, "<div align=justify>"
        Print #IntFile1, "<table border=0 cellpadding=0 cellspacing=0 width=336 height=41>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=334 height=21><font size=1 color=#000080 face=Arial>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & EndLoja & " - " & BairroLoja & "</font></td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=334 height=20><font face=Arial size=1 color=#000080>SAO PAULO - " & UfLoja & " - CEP : " & UCase(Right(String(8, "0") & CepLoja, 8)) & " - FONE : " & TelefoneLoja & "</font></td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "</table>"
        Print #IntFile1, "</div>"
        Print #IntFile1, "<table border=0 cellpadding=0 cellspacing=0 width=750 height=48>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=700 rowspan=3><img border=0 src=http://www.demeo.com.br/images/truck.jpg width=750 height=40 align=right></td>"
        Print #IntFile1, "</tr>"
        Print #IntFile1, "</table>"
        Print #IntFile1, "</div>"
        Print #IntFile1, "<div align=left>"
        Print #IntFile1, "<table border=1 cellpadding=1 cellspacing=0 width=759 height=36 solid; border-width: 0 bordercolor=#000080>"
        Print #IntFile1, "<tr>"
        Print #IntFile1, "<td width=640 height=20><font face=Arial color=#000080><b>Referencia / Descrição</b></font></td>"
        Print #IntFile1, "<td width=41 height=20><font face=Arial color=#000080><b>Qtde</b></font></td>"
        Print #IntFile1, "<td width=76 height=20><font face=Arial color=#000080><b>Valor</b></font></td>"
        Print #IntFile1, "<td width=66 height=20><font face=Arial color=#000080><b>Desconto</b></font></td>"
        Print #IntFile1, "<td width=79 height=20><font face=Arial color=#000080><b>Total</b></font></td>"
        Print #IntFile1, "</tr>"
    End If
    
    SQL = ""
    
    SQL = "Select PR_Descricao,Qtde,VlUnit,(VLTotItem - VlUnit2) as Desconto,VlUnit2,Referencia,VE_Nome,CP_Condicao " _
        & "From Produto, NfItens, Vende, CondicaoPagamento, NFCapa " _
        & "Where PR_Referencia = Referencia " _
        & "and NfCapa.NumeroPed=NfItens.NumeroPed " _
        & "and NfItens.NumeroPed=" & Pedido & " " _
        & "and VE_Codigo=NfCapa.Vendedor " _
        & "and CP_Codigo=*NfCapa.CondPag "
    Set rsPedido = rdoCnLoja.OpenResultset(SQL)
    SomaPedido = 0
    Vendedor = ""
    CondPag = ""
    EndLoja = ""
    CepLoja = ""
    TelefoneLoja = ""
    BairroLoja = ""
    UfLoja = ""
    If Not rsPedido.EOF Then
        Vendedor = rsPedido("VE_Nome")
        CondPag = IIf(IsNull(rsPedido("CP_Condicao")), "A VISTA", rsPedido("CP_Condicao"))
        Do While Not rsPedido.EOF
            Print #IntFile1, "<tr>"
            Print #IntFile1, "<td width=640 height=16><font face=Verdana size=2 color=#000080>" & rsPedido("Referencia") & " " & rsPedido("PR_Descricao") & "</font></td>"
            Print #IntFile1, "<td width=41 height=16 align=right><font face=Verdana size=2 color=#000080>" & rsPedido("Qtde") & "</font></td>"
            Print #IntFile1, "<td width=76 height=16 align=right><font face=Verdana size=2 color=#000080>" & Format(rsPedido("VlUnit"), "##,###,###0.00") & "</font></td>"
            Print #IntFile1, "<td width=66 height=16 align=right><font face=Verdana size=2 color=#000080>" & Format(rsPedido("Desconto"), "##,###,###0.00") & "</font></td>"
            Print #IntFile1, "<td width=79 height=16 align=right><font face=Verdana size=2 color=#000080>" & Format(rsPedido("VlUnit2"), "##,###,###0.00") & "</font></td>"
            
            Print #IntFile1, "</tr>"
            SomaPedido = SomaPedido + Format(rsPedido("VlUnit2"), "##,###,###0.00")
            rsPedido.MoveNext
        Loop
    End If
    Print #IntFile1, "</table>"
    Print #IntFile1, "</div>"
    Print #IntFile1, "</p>"
      
    Print #IntFile1, "</table>"
    
    Print #IntFile1, "</div>"
    Print #IntFile1, "<table border=1 cellpadding=1 cellspacing=1 width=223 align=right bordercolor=#000080>"
    Print #IntFile1, "<tr>"
    Print #IntFile1, "<td width=120 align=left nowrap bgcolor=#FFFFFF bordercolor=#FFFFFF bordercolorlight=#FFFFFF bordercolordark=#FFFFFF><font face=Arial color=#000080>Total Pedido</font></td>"
    Print #IntFile1, "<td width=120 align=Right><font color=#000080>" & Format(SomaPedido, "###,###,###,##0.00") & "</font></td>"
    Print #IntFile1, "</tr>"
    Print #IntFile1, "</table>"
    Print #IntFile1, "<p>&nbsp;</p>"
    Print #IntFile1, "<p>&nbsp;</p>"
    Print #IntFile1, "<div align=justify>"
    Print #IntFile1, "<table border=0 cellpadding=0 cellspacing=0 width=750 height=62>"
    Print #IntFile1, "<tr>"
    Print #IntFile1, "<td width=259 height=21><font color=#000080 face=System>COND PAGTO&nbsp;&nbsp; : " & UCase(CondPag) & "</font></td>"
    Print #IntFile1, "</tr>"
    Print #IntFile1, "<tr>"
    Print #IntFile1, "<td width=259 height=21><font color=#000080 face=System>VALIDADE&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : " & Format(DateAdd("D", rdoValidadeCotacao("CT_ValidadeCotacao"), Date), "dd/mm/yyyy") & "</font></td>"
    Print #IntFile1, "</tr>"
    Print #IntFile1, "<tr>"
    Print #IntFile1, "<td width=259 height=20><font color=#000080 face=System>VENDEDOR&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: " & UCase(Vendedor) & "</font></td>"
    Print #IntFile1, "</tr>"
    Print #IntFile1, "</table>"
    Print #IntFile1, "</div>"
    
    Print #IntFile1, "</body>"

    Print #IntFile1, "</html>"

    Close #IntFile1
End Sub


