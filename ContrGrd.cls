VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlaGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarCabecalho As String
Private mvarFormato As String
Private mvarTamanho As String
Private mvarValores As String
Private mvarCampos As String
Private mvarAlinhamento As String
Private mvarNomeFormulario As String

Dim matCabecalho() As String
Dim matFormato() As String
Dim matTamanho() As String
Dim matValores() As String
Dim matCampos() As String
Dim matAlinhamento() As String

Dim Formulario As Form

'constantes para formato
Private Const moeda = "R$ ###,###,###,###0.00"
Private Const NumDecimal = "##,###,###,###0.00"
Private Const Porcentagem = "##,###,###,###0.00"
Private Const Data = "dd/mm/yyyy"
Private Const Maiusculo = ">"

'constantes para alinhamento
Private Const Esquerda = 1
Private Const Centro = 4
Private Const Direita = 7

Public ConexaoGrid As Variant
Public SQL As String
Public UseValores As Boolean
Public Colunas As Long
Public NomeGrid As String
Public LinhasVisiveis As Long

Public Event RegistroNaoExiste()
Public Event BeforeFormat(ByVal Coluna As Long, ByRef Valor As Variant)
Public Event BeforeAdd(ByRef rdoInterno As Variant)
Public Event Acabou()

Public Sub SelecionaLinha(ByVal KeyCode As Integer)
    Select Case KeyCode
        Case vbKeyUp:
            If Formulario.Controls(NomeGrid).Row > Formulario.Controls(NomeGrid).FixedRows Then
                Formulario.Controls(NomeGrid).Row = Formulario.Controls(NomeGrid).Row - 1
                Formulario.Controls(NomeGrid).ColSel = Formulario.Controls(NomeGrid).Cols - 1
            End If
        Case vbKeyDown:
            If Formulario.Controls(NomeGrid).Row < Formulario.Controls(NomeGrid).Rows - 1 Then
                Formulario.Controls(NomeGrid).Row = Formulario.Controls(NomeGrid).Row + 1
                Formulario.Controls(NomeGrid).ColSel = Formulario.Controls(NomeGrid).Cols - 1
            End If
    End Select
    
    TornaLinhaVisivel Formulario.Controls(NomeGrid).Row

End Sub

Public Sub TornaLinhaVisivel(ByVal NumeroLinha As Long)

    If NumeroLinha < Formulario.Controls(NomeGrid).TopRow Or NumeroLinha > (Formulario.Controls(NomeGrid).TopRow + (LinhasVisiveis - 1)) Then
        If (Formulario.Controls(NomeGrid).Rows - Formulario.Controls(NomeGrid).TopRow) >= LinhasVisiveis Then
            Formulario.Controls(NomeGrid).TopRow = NumeroLinha
        Else
            Formulario.Controls(NomeGrid).TopRow = Formulario.Controls(NomeGrid).Rows - LinhasVisiveis
        End If
    End If

End Sub

'Public Sub Preencher()
'
'    Dim Indice As Long
'    Dim Maximo As Long
'    Dim StringAdicionar As String
'    Dim ChrTAB As String * 1
'    Dim Dado As Variant
'    'Dim RdoDados As rdoResultset
'
'    On Error Resume Next
'
'    ChrTAB = Chr(vbKeyTab)
'
'    Formulario.Controls(NomeGrid).Redraw = False
'    If UseValores Then
'        Maximo = Colunas - 1
'        StringAdicionar = ""
'        For Indice = 0 To Maximo Step 1
'            Dado = Trim(matValores(Indice))
'
'            RaiseEvent BeforeFormat(Indice, Dado)
'
'            Select Case UCase(Trim(matFormato(Indice)))
'                Case "MOEDA"
'                    Dado = Format(Dado, Moeda)
'                Case "DECIMAL"
'                    Dado = Format(Dado, NumDecimal)
'                Case "PORCENTAGEM"
'                    Dado = Format(Dado, Porcentagem) & "%"
'                Case "DATA"
'                    Dado = Format(Dado, Data)
'                Case "MAIUSCULO"
'                    Dado = Format(Dado, Maiusculo)
'                Case "CARACTERE", "NUMERO"
'
'                Case Else
'                    If Trim(matFormato(Indice)) <> "" Then
'                        Dado = Format(Dado, Trim(matFormato(Indice)))
'                    End If
'            End Select
'
'            StringAdicionar = StringAdicionar & Dado & ChrTAB
'        Next Indice
'        Formulario.Controls(NomeGrid).AddItem StringAdicionar
'        RaiseEvent Acabou
'    Else
'        'Set RdoDados = ConexaoGrid.OpenResultset(SQL, Options:=rdExecDirect)
'
''        If RdoDados.EOF Then
''            RdoDados.Close
''            Formulario.Controls(NomeGrid).Redraw = True
''            RaiseEvent RegistroNaoExiste
''            Exit Sub
''        End If
'
'        Do While Not RdoDados.EOF
'            StringAdicionar = ""
'            Maximo = UBound(matCampos)
'            For Indice = 0 To Maximo Step 1
'                Dado = RdoDados(matCampos(Indice))
'
'                RaiseEvent BeforeFormat(Indice, Dado)
'
'                Select Case UCase(Trim(matFormato(Indice)))
'                    Case "MOEDA"
'                        Dado = Format(Dado, Moeda)
'                    Case "DECIMAL"
'                        Dado = Format(Dado, NumDecimal)
'                    Case "PORCENTAGEM"
'                        Dado = Format(Dado, Porcentagem) & "%"
'                    Case "DATA"
'                        Dado = Format(Dado, Data)
'                    Case "MAIUSCULO"
'                        Dado = Format(Dado, Maiusculo)
'                    Case "CARACTERE", "NUMERO"
'
'                    Case Else
'                       If Trim(matFormato(Indice)) <> "" Then
'                           Dado = Format(Dado, Trim(matFormato(Indice)))
'                       End If
'                End Select
'
'                StringAdicionar = StringAdicionar & Dado & ChrTAB
'            Next Indice
'
'            RaiseEvent BeforeAdd(RdoDados)
'
'            Formulario.Controls(NomeGrid).AddItem StringAdicionar
'
'            RdoDados.MoveNext
'        Loop
'        RdoDados.Close
'        RaiseEvent Acabou
'    End If
'    Formulario.Controls(NomeGrid).Redraw = True
'
'End Sub

Public Sub MontaCabecalho()

    Dim Indice As Long
    Dim Maximo As Long
    Dim TipoAlinhamento As Long
    Dim AlinCabecalho As Long
    Dim Inicio As Long
    Dim Linha As Long
    
    On Error Resume Next
    
    Formulario.Controls(NomeGrid).Cols = Colunas
    
    Maximo = Colunas - 1
    
    For Linha = 0 To Formulario.Controls(NomeGrid).FixedRows - 1 Step 1
        For Indice = 0 To Maximo Step 1
            Select Case Left(matCabecalho(Indice + (Linha * (Maximo + 1))), 1)
                Case "<"
                    AlinCabecalho = Esquerda
                    Inicio = 2
                Case "^"
                    AlinCabecalho = Centro
                    Inicio = 2
                Case ">"
                    AlinCabecalho = Direita
                    Inicio = 2
                Case Else
                    AlinCabecalho = 0
                    Inicio = 1
            End Select
            
            Formulario.Controls(NomeGrid).TextMatrix(Linha, Indice) = Mid(matCabecalho(Indice + (Linha * (Maximo + 1))), Inicio)
            If Linha = 0 Then
                Formulario.Controls(NomeGrid).ColWidth(Indice) = Val(matTamanho(Indice))
                
                Select Case UCase(matAlinhamento(Indice))
                    Case "ESQUERDA"
                        TipoAlinhamento = Esquerda
                    Case "DIREITA"
                        TipoAlinhamento = Direita
                    Case "CENTRO"
                        TipoAlinhamento = Centro
                End Select
                Formulario.Controls(NomeGrid).ColAlignment(Indice) = TipoAlinhamento
                
                Formulario.Controls(NomeGrid).FixedAlignment(Indice) = AlinCabecalho
            End If
        Next Indice
    Next Linha

End Sub

Public Sub Clear()

    Formulario.Controls(NomeGrid).Redraw = False
    Formulario.Controls(NomeGrid).Rows = Formulario.Controls(NomeGrid).FixedRows + 1
    Formulario.Controls(NomeGrid).AddItem ""
    Formulario.Controls(NomeGrid).RemoveItem Formulario.Controls(NomeGrid).FixedRows
    Formulario.Controls(NomeGrid).Redraw = True

End Sub

Public Property Let NomeFormulario(ByVal vData As String)
    
    Dim Indice As Long
    Dim Maximo As Long
    
    Maximo = Forms.Count - 1
    
    mvarNomeFormulario = ""
    
    For Indice = 0 To Maximo Step 1
        If UCase(Forms(Indice).Name) = UCase(vData) Then
            mvarNomeFormulario = vData
            Set Formulario = Forms(Indice)
            Exit For
        End If
    Next Indice
    
    If mvarNomeFormulario = "" Then
        Err.Raise -10, "Classe Cadastro", "Nome de formulario não encontrado."
    End If

End Property

Public Property Get NomeFormulario() As String
    
    NomeFormulario = mvarNomeFormulario

End Property

Public Property Let Alinhamento(ByVal vData As String)
    
    mvarAlinhamento = vData
    MontarMatriz vData, matAlinhamento

End Property

Public Property Get Alinhamento() As String
    
    Alinhamento = mvarAlinhamento

End Property

Public Property Let Campos(ByVal vData As String)
    
    mvarCampos = vData
    MontarMatriz vData, matCampos

End Property

Public Property Get Campos() As String
    
    Campos = mvarCampos

End Property

Public Property Let Valores(ByVal vData As String)
    
    mvarValores = vData
    MontarMatriz vData, matValores

End Property

Public Property Get Valores() As String
    
    Valores = mvarValores

End Property

Public Property Let Tamanho(ByVal vData As String)
    
    mvarTamanho = vData
    MontarMatriz vData, matTamanho

End Property

Public Property Get Tamanho() As String
    
    Tamanho = mvarTamanho

End Property

Public Property Let Formato(ByVal vData As String)
    
    mvarFormato = vData
    MontarMatriz vData, matFormato

End Property

Public Property Get Formato() As String
    
    Formato = mvarFormato

End Property

Public Property Let Cabecalho(ByVal vData As String)
    
    mvarCabecalho = vData
    MontarMatriz vData, matCabecalho

End Property

Public Property Get Cabecalho() As String
    
    Cabecalho = mvarCabecalho

End Property

Private Sub Class_Initialize()
    
    mvarCabecalho = ""
    mvarFormato = ""
    mvarTamanho = ""
    mvarValores = ""
    mvarCampos = ""
    mvarAlinhamento = ""
    mvarNomeFormulario = ""
    LinhasVisiveis = 2
    
    UseValores = False

End Sub

Private Sub MontarMatriz(ByVal vData As String, ByRef Matriz As Variant)

    Dim Tamanho As Long
    Dim CharPos As Long
    Dim NomeLido As String
    Dim Indice As Long
    Dim PosAnt As Long
    
    Tamanho = Len(vData)
    
    ReDim Matriz(0) As String
    
    If Tamanho > 0 Then
        CharPos = InStr(vData, ";")
        If CharPos = 0 Then
            Matriz(0) = vData
        Else
            PosAnt = 1
            Indice = 0
            Do While CharPos <> 0
                NomeLido = Mid(vData, PosAnt, CharPos - PosAnt)
                PosAnt = CharPos + 2
                CharPos = InStr(CharPos + 1, vData, ";")
                Matriz(Indice) = NomeLido
                Indice = Indice + 1
                ReDim Preserve Matriz(Indice) As String
            Loop
            NomeLido = Mid(vData, PosAnt)
            Matriz(Indice) = NomeLido
        End If
    End If

End Sub


