VERSION 5.00
Begin VB.Form frmReemissaoNotaFiscal 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Reemissao de notas fiscais"
   ClientHeight    =   2475
   ClientLeft      =   12435
   ClientTop       =   4170
   ClientWidth     =   5070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   2430
      Left            =   75
      TabIndex        =   4
      Top             =   0
      Width           =   4890
      Begin VB.TextBox txtPedido 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2055
         TabIndex        =   2
         Top             =   1305
         Width           =   1035
      End
      Begin VB.TextBox txtNotaFiscal 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   315
         TabIndex        =   0
         Top             =   1305
         Width           =   1065
      End
      Begin VB.TextBox txtSerie 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1425
         TabIndex        =   1
         Top             =   1305
         Width           =   555
      End
      Begin VB.TextBox txtValorNF 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3150
         TabIndex        =   3
         Top             =   1290
         Width           =   1260
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nota Fiscal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1695
         TabIndex        =   9
         Top             =   570
         Width           =   1200
      End
      Begin VB.Label lblPedido 
         AutoSize        =   -1  'True
         BackColor       =   &H0081E8FA&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro.Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2070
         TabIndex        =   8
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label lblValorTotal 
         AutoSize        =   -1  'True
         BackColor       =   &H0081E8FA&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3165
         TabIndex        =   7
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label lblNotaFiscal 
         BackColor       =   &H0081E8FA&
         BackStyle       =   0  'Transparent
         Caption         =   "Nota Fiscal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   315
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblSerie 
         AutoSize        =   -1  'True
         BackColor       =   &H0081E8FA&
         BackStyle       =   0  'Transparent
         Caption         =   "Série"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1440
         TabIndex        =   5
         Top             =   1065
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmReemissaoNotaFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsPegaLojaControle As rdoResultset
Dim FIN As rdoResultset
Dim WARQUIVO As String
Dim sql As String
Dim wUltimoCupom As Double
Dim WGrupoAtualzado As Double
Dim wNumeroPedido As Double
Dim wWhere As String

Private Sub FinalizarReimpressao()
    
    If txtSerie.text <> PegaSerieNota Then
      Exit Sub
    End If
       
          If txtNotaFiscal.text = "" Then
            MsgBox "Preencha todos os campos", vbCritical, "Atenção"
            txtNotaFiscal.SetFocus
            txtSerie.text = ""
            txtValorNF.text = ""
            txtPedido.text = ""
            Exit Sub
        End If
        If txtSerie.text = "" Then
            MsgBox "Preencha todos os campos", vbCritical, "Atenção"
            txtNotaFiscal.SetFocus
            txtSerie.text = ""
            txtValorNF.text = ""
            txtPedido.text = ""

            Exit Sub
        End If
  
         If txtSerie.text <> PegaSerieNota Then
            MsgBox "Tela somente para reimprimir nota Serie " & PegaSerieNota & ".", vbCritical, "Atenção"
            txtNotaFiscal.SetFocus
            txtSerie.text = ""
            txtValorNF.text = ""
            txtPedido.text = ""
            Exit Sub
        End If
    
             
     sql = "Select * From Nfcapa " _
         & "Where nf = " & txtNotaFiscal.text & " and serie = '" & txtSerie.text & "'"

            RsDados.CursorLocation = adUseClient
            RsDados.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

             If RsDados.EOF = False Then
                 wPegaDesconto = IIf(IsNull(RsDados("Desconto")) = True, 0, RsDados("Desconto"))
                 wPegaFrete = IIf(IsNull(RsDados("FreteCobr")) = True, 0, RsDados("FreteCobr"))
                 wNotaFiscalReemissao = RsDados("NF")
                 wSerieReemissao = Trim(RsDados("Serie"))
                 If MsgBox("Deseja Reemitir a Nota Fiscal Nº " & wNotaFiscalReemissao, vbYesNo + vbQuestion, "Atenção") = vbYes Then
                    RsDados.Close
                    EmiteNotafiscal wNotaFiscalReemissao, wSerieReemissao
                    Limpar
                    txtNotaFiscal.SetFocus
                    Exit Sub
                 End If
             Else
                 MsgBox "Nota fiscal não encontrada Nº " & wNotaFiscalReemissao, vbCritical, "Atenção"
                 RsDados.Close
                 Exit Sub
             End If

    RsDados.Close
    Limpar
    txtNotaFiscal.SetFocus
    
End Sub



Private Sub cmdRetorna1_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   txtNotaFiscal.SetFocus
End Sub

Private Sub Form_Load()
    
    left = 0
    top = (Screen.Height - Height) / 3
    
    txtNotaFiscal.text = ""
    txtSerie.text = ""
    txtValorNF.text = ""
    txtPedido.text = ""
Call AjustaTela(frmReemissaoNotaFiscal)

End Sub

Private Sub txtAux_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
End If
If KeyAscii = 13 Then
      txtSerie.SetFocus
      Call FinalizarReimpressao
End If

End Sub

Private Sub txtNotaFiscal_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
   Unload Me
End If

If KeyAscii = 13 Then
   If txtNotaFiscal <> "" Then
      txtSerie.SetFocus
   End If
End If


End Sub

Private Sub txtNotaFiscal_LostFocus()
If txtNotaFiscal.text = "" Then
    Exit Sub
End If

If IsNumeric(txtNotaFiscal.text) = False Then
    
    txtNotaFiscal.text = ""
    txtNotaFiscal.SetFocus
    Exit Sub
End If
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If txtSerie.text <> "" Then
      If txtValorNF.text = "" Then
         Call txtSerie_LostFocus
         txtSerie.SetFocus
      Else
         Call FinalizarReimpressao
      End If
   End If
End If

If KeyAscii = 27 Then
   txtNotaFiscal.SelStart = 0
   txtNotaFiscal.SelLength = Len(txtNotaFiscal.text)
   txtNotaFiscal.SetFocus

End If



End Sub

Private Sub txtSerie_LostFocus()
 txtSerie.text = UCase(txtSerie.text)
If txtSerie.text = "" Then
    Exit Sub
End If

 If txtSerie.text <> PegaSerieNota Then
    Exit Sub
 End If

If txtNotaFiscal.text = "" Then
   MsgBox "Preencha todos os campos", vbCritical, "Atenção"
   txtNotaFiscal.SelStart = 0
   txtNotaFiscal.SelLength = Len(txtNotaFiscal.text)
   txtNotaFiscal.SetFocus
   Exit Sub
End If


sql = "SELECT TOTALNOTA, NF, SERIE,TipoNota,numeroped FROM NFCAPA WHERE NF = " & txtNotaFiscal.text & " " _
    & "AND SERIE = '" & UCase(Trim(txtSerie.text)) & "' and TIPONOTA <> 'C'"
    
 ADOCancela.CursorLocation = adUseClient
 ADOCancela.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

    If Not ADOCancela.EOF Then
       txtValorNF.text = Format(ADOCancela("TOTALNOTA"), "0.00")
       txtPedido.text = ADOCancela("numeroped")

    Else
        MsgBox "NF não encontrada", vbInformation, "Aviso"
        txtNotaFiscal.SelStart = 0
        txtNotaFiscal.SelLength = Len(txtNotaFiscal.text)
        txtNotaFiscal.SetFocus
    End If
ADOCancela.Close
End Sub


Private Sub Limpar()
txtPedido.text = ""
txtSerie.text = ""
txtValorNF.text = ""
txtNotaFiscal.text = ""
txtNotaFiscal.SetFocus
End Sub
