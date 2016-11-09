VERSION 5.00
Begin VB.Form frmEmissaoRomaneio 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Romaneio"
   ClientHeight    =   2640
   ClientLeft      =   5640
   ClientTop       =   4215
   ClientWidth     =   5040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2640
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   2430
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   4890
      Begin VB.TextBox txtValorNF 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3150
         TabIndex        =   4
         Top             =   1305
         Width           =   1260
      End
      Begin VB.TextBox txtSerie 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   1305
         Width           =   555
      End
      Begin VB.TextBox txtNotaFiscal 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   315
         TabIndex        =   2
         Top             =   1305
         Width           =   1065
      End
      Begin VB.TextBox txtPedido 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2055
         TabIndex        =   1
         Top             =   1305
         Width           =   1035
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
         TabIndex        =   9
         Top             =   1065
         Width           =   450
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
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
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
         TabIndex        =   6
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Romaneio"
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
         TabIndex        =   5
         Top             =   570
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmEmissaoRomaneio"
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
    
    If txtSerie.Text <> "00" Then
      Exit Sub
    End If
       
          If txtNotaFiscal.Text = "" Then
            MsgBox "Preencha todos os campos", vbCritical, "Atenção"
            txtNotaFiscal.SetFocus
            txtSerie.Text = ""
            txtValorNF.Text = ""
            Exit Sub
        End If
        If txtSerie.Text = "" Then
            MsgBox "Preencha todos os campos", vbCritical, "Atenção"
            txtNotaFiscal.SetFocus
            txtSerie.Text = ""
            txtValorNF.Text = ""
            Exit Sub
        End If
   
         If txtSerie.Text <> "00" Then
            MsgBox "Tela somente para imprimir Romaneio.", vbCritical, "Atenção"
            txtNotaFiscal.SetFocus
            txtSerie.Text = ""
            txtValorNF.Text = ""
            Exit Sub
        End If
    
             
     sql = "Select * From Nfcapa " _
         & "Where nf = " & txtNotaFiscal.Text & " and serie = '" & txtSerie.Text & "'"

            RsDados.CursorLocation = adUseClient
            RsDados.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

             If RsDados.EOF = False Then
                 wPegaDesconto = RsDados("Desconto")
                 wPegaFrete = RsDados("FreteCobr")
                 wNotaFiscalReemissao = RsDados("NF")
                 wSerieReemissao = Trim(RsDados("Serie"))
                    RsDados.Close
                    'EmiteNotafiscal wNotaFiscalReemissao, wSerieReemissao
                    wlblloja = Trim(GLB_Loja)
                    NroNotaFiscal = txtNotaFiscal.Text
                    
                    wQdteViasImpressao = 1
                    Call BuscaQtdeViaImpressaoMovimento
  
                    For i = 1 To wQdteViasImpressao
                       Call ImprimeRomaneio
                    Next i
                    

 '                  cmdImprimir1.Visible = False
                    Limpar
                    txtNotaFiscal.SetFocus

             Else
                 MsgBox "Romaneio não encontrado Nº " & wNotaFiscalReemissao, vbCritical, "Atenção"
                 RsDados.Close
                 Exit Sub
             End If

  '  RsDados.Close
  '  cmdImprimir1.Visible = False
    Limpar
    txtNotaFiscal.SetFocus
    
End Sub


Private Sub cmdImprimir1_Click()

End Sub

Private Sub Form_Activate()
 txtNotaFiscal.SetFocus
End Sub

Private Sub Form_Load()

 '   Left = 0
 '   Top = (Screen.Height - Height) / 3
    
    txtNotaFiscal.Text = ""
    txtSerie.Text = ""
    txtValorNF.Text = ""

Call AjustaTela(frmEmissaoRomaneio)
    
End Sub

Private Sub Label4_Click()
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
If txtNotaFiscal.Text = "" Then
    Exit Sub
End If

If IsNumeric(txtNotaFiscal.Text) = False Then
    
    txtNotaFiscal.Text = ""
    txtNotaFiscal.SetFocus
    Exit Sub
End If
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   If txtSerie.Text <> "" Then
      If txtValorNF.Text = "" Then
         Call txtSerie_LostFocus
         txtSerie.SetFocus
      Else
         Call FinalizarReimpressao
      End If
   End If
End If

If KeyAscii = 27 Then
   txtNotaFiscal.SelStart = 0
   txtNotaFiscal.SelLength = Len(txtNotaFiscal.Text)
   txtNotaFiscal.SetFocus

End If
End Sub

Private Sub txtSerie_LostFocus()
 txtSerie.Text = UCase(txtSerie.Text)
If txtSerie.Text = "" Then
    Exit Sub
End If

 If txtSerie.Text <> "00" Then
    Exit Sub
 End If

If txtNotaFiscal.Text = "" Then
   MsgBox "Preencha todos os campos", vbCritical, "Atenção"
   txtNotaFiscal.SelStart = 0
   txtNotaFiscal.SelLength = Len(txtNotaFiscal.Text)
   txtNotaFiscal.SetFocus
   Exit Sub
End If

sql = "SELECT TOTALNOTA, NF, SERIE,TipoNota,numeroped FROM NFCAPA WHERE NF = " & txtNotaFiscal.Text & " " _
    & "AND SERIE = '" & UCase(Trim(txtSerie.Text)) & "' and TIPONOTA <> 'C'"
    
 ADOCancela.CursorLocation = adUseClient
 ADOCancela.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

    If Not ADOCancela.EOF Then
       txtValorNF.Text = Format(ADOCancela("TOTALNOTA"), "0.00")
       txtPedido.Text = ADOCancela("numeroped")
        'txtAux.Enabled = True
        'txtAux.SetFocus
    Else
        MsgBox "Romaneio não encontrado", vbInformation, "Aviso"
'        txtAux.Enabled = False
        txtSerie.Text = ""
        txtNotaFiscal.SelStart = 0
        txtNotaFiscal.SelLength = Len(txtNotaFiscal.Text)
        txtNotaFiscal.SetFocus
    End If
ADOCancela.Close
'.SetFocus
End Sub


Private Sub Limpar()
txtPedido.Text = ""
txtSerie.Text = ""
txtValorNF.Text = ""
txtNotaFiscal.Text = ""
txtNotaFiscal.SetFocus
End Sub


