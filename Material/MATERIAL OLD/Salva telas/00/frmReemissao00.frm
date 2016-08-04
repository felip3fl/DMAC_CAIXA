VERSION 5.00
Begin VB.Form frmReemissao00 
   Caption         =   "Romaneio"
   ClientHeight    =   2640
   ClientLeft      =   4245
   ClientTop       =   1755
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmReemissao00.frx":0000
   ScaleHeight     =   2640
   ScaleWidth      =   5265
   Begin VB.TextBox txtValorNF 
      BackColor       =   &H0081E8FA&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3045
      TabIndex        =   4
      Top             =   1095
      Width           =   1260
   End
   Begin VB.TextBox txtSerie 
      BackColor       =   &H0081E8FA&
      Height          =   315
      Left            =   1335
      TabIndex        =   3
      Top             =   1095
      Width           =   555
   End
   Begin VB.TextBox txtNotaFiscal 
      BackColor       =   &H0081E8FA&
      Height          =   315
      Left            =   210
      TabIndex        =   2
      Top             =   1095
      Width           =   1065
   End
   Begin VB.TextBox txtPedido 
      BackColor       =   &H0081E8FA&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1950
      TabIndex        =   0
      Top             =   1095
      Width           =   1035
   End
   Begin Balcao2010.chameleonButton cmdRetorna1 
      Height          =   435
      Left            =   4470
      TabIndex        =   1
      Top             =   105
      Width           =   405
      _extentx        =   714
      _extenty        =   767
      btype           =   13
      tx              =   ""
      enab            =   -1  'True
      font            =   "frmReemissao00.frx":355B6
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   14215660
      bcolo           =   14215660
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmReemissao00.frx":355E2
      picn            =   "frmReemissao00.frx":35600
      pich            =   "frmReemissao00.frx":35E54
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin Balcao2010.chameleonButton cmdImprimir1 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   4470
      TabIndex        =   5
      Top             =   2010
      Width           =   405
      _extentx        =   714
      _extenty        =   767
      btype           =   13
      tx              =   ""
      enab            =   -1  'True
      font            =   "frmReemissao00.frx":366A8
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   14215660
      bcolo           =   14215660
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmReemissao00.frx":366D4
      picn            =   "frmReemissao00.frx":366F2
      pich            =   "frmReemissao00.frx":37346
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
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
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   210
      TabIndex        =   10
      Top             =   870
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3060
      TabIndex        =   9
      Top             =   870
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1965
      TabIndex        =   8
      Top             =   870
      Width           =   960
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1350
      TabIndex        =   7
      Top             =   885
      Width           =   450
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
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   990
      TabIndex        =   6
      Top             =   255
      Width           =   1080
   End
End
Attribute VB_Name = "frmReemissao00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsPegaLojaControle As rdoResultset
Dim FIN As rdoResultset
Dim WARQUIVO As String
Dim SQL As String
Dim wUltimoCupom As Double
Dim WGrupoAtualzado As Double
Dim WnumeroPedido As Double
Dim wWhere As String

Private Sub cmdImprimir1_Click()
    
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
        
      '  If txtSerie.Text = "CF" Then
      '      MsgBox "Cupom Fiscal não pode ser REIMPRESSO.", vbCritical, "Atenção"
      '      txtNotaFiscal.SetFocus
      '      txtSerie.Text = ""
      '      txtValorNF.Text = ""
      '      Exit Sub
      '  End If
      
         If txtSerie.Text <> "00" Then
            MsgBox "Tela somente para imprimir Romaneio.", vbCritical, "Atenção"
            txtNotaFiscal.SetFocus
            txtSerie.Text = ""
            txtValorNF.Text = ""
            Exit Sub
        End If
    
             
     SQL = "Select * From Nfcapa " _
         & "Where nf = " & txtNotaFiscal.Text & " and serie = '" & txtSerie.Text & "'"

            RsDados.CursorLocation = adUseClient
            RsDados.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

             If RsDados.EOF = False Then
                 wNotaFiscalReemissao = RsDados("NF")
                 wSerieReemissao = Trim(RsDados("Serie"))
             '    If MsgBox("Deseja Reemitir a Nota Fiscal Nº " & wNotaFiscalReemissao, vbYesNo + vbQuestion, "Atenção") = vbYes Then
             '       RsDados.Close
                    EmiteNotafiscal wNotaFiscalReemissao, wSerieReemissao
                    cmdImprimir1.Visible = False
                    Limpar
                    txtNotaFiscal.SetFocus
             '       Exit Sub
             '    End If
             Else
                 MsgBox "Romaneio não encontrado Nº " & wNotaFiscalReemissao, vbCritical, "Atenção"
                 RsDados.Close
                 Exit Sub
             End If

    RsDados.Close
    cmdImprimir1.Visible = False
    Limpar
    txtNotaFiscal.SetFocus
    
End Sub



Private Sub cmdRetorna1_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    'Left = (Screen.Width - Width) / 2
    'Top = (Screen.Height - Height) / 3
    txtNotaFiscal.Text = ""
    txtSerie.Text = ""
    txtValorNF.Text = ""
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
      cmdRetorna1.SetFocus
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

 

'SQL = "SELECT TOTALNOTA, NF, SERIE,TipoNota FROM NFCAPA WHERE NF = " & txtNotaFiscal.Text & " " _
    & "AND SERIE = '" & txtSerie.Text & "' AND numeroped = " & txtPedido.Text & ""
    
SQL = "SELECT TOTALNOTA, NF, SERIE,TipoNota,numeroped FROM NFCAPA WHERE NF = " & txtNotaFiscal.Text & " " _
    & "AND SERIE = '" & UCase(Trim(txtSerie.Text)) & "' and TIPONOTA <> 'CA'"
    
 ADOCancela.CursorLocation = adUseClient
 ADOCancela.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

    If Not ADOCancela.EOF Then
       txtValorNF.Text = Format(ADOCancela("TOTALNOTA"), "0.00")
       txtPedido.Text = ADOCancela("numeroped")
      ' WTipoNota = ADOCancela("TipoNota")
        cmdImprimir1.Visible = True
        cmdImprimir1.SetFocus
    Else
        MsgBox "Romaneio não encontrado", vbInformation, "Aviso"
        cmdImprimir1.Visible = False
        txtSerie.Text = ""
        txtNotaFiscal.SelStart = 0
        txtNotaFiscal.SelLength = Len(txtNotaFiscal.Text)
        txtNotaFiscal.SetFocus
End If
ADOCancela.Close

End Sub


Private Sub Limpar()
txtPedido.Text = ""
txtSerie.Text = ""
txtValorNF.Text = ""
txtNotaFiscal.Text = ""
txtNotaFiscal.SetFocus
End Sub


