VERSION 5.00
Begin VB.Form frmReemissaoNotaFiscal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reemissao de notas fiscais"
   ClientHeight    =   2640
   ClientLeft      =   5910
   ClientTop       =   4020
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmReemissaoNotaFiscal.frx":0000
   ScaleHeight     =   2640
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   Begin Balcao2010.chameleonButton cmdImprimir1 
      Height          =   435
      Left            =   4710
      TabIndex        =   4
      Top             =   2040
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   767
      BTYPE           =   13
      TX              =   ""
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmReemissaoNotaFiscal.frx":355B6
      PICN            =   "frmReemissaoNotaFiscal.frx":355D2
      PICH            =   "frmReemissaoNotaFiscal.frx":36224
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtPedido 
      BackColor       =   &H0081E8FA&
      Enabled         =   0   'False
      Height          =   315
      Left            =   2190
      TabIndex        =   2
      Top             =   1140
      Width           =   1035
   End
   Begin Balcao2010.chameleonButton cmdRetorna1 
      Height          =   435
      Left            =   4710
      TabIndex        =   5
      Top             =   150
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   767
      BTYPE           =   13
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmReemissaoNotaFiscal.frx":36E76
      PICN            =   "frmReemissaoNotaFiscal.frx":36E92
      PICH            =   "frmReemissaoNotaFiscal.frx":376E4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtNotaFiscal 
      BackColor       =   &H0081E8FA&
      Height          =   315
      Left            =   450
      TabIndex        =   0
      Top             =   1140
      Width           =   1065
   End
   Begin VB.TextBox txtSerie 
      BackColor       =   &H0081E8FA&
      Height          =   315
      Left            =   1575
      TabIndex        =   1
      Top             =   1140
      Width           =   555
   End
   Begin VB.TextBox txtValorNF 
      BackColor       =   &H0081E8FA&
      Enabled         =   0   'False
      Height          =   315
      Left            =   3285
      TabIndex        =   3
      Top             =   1140
      Width           =   1260
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
      Left            =   1590
      TabIndex        =   9
      Top             =   930
      Width           =   450
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
      Left            =   2205
      TabIndex        =   8
      Top             =   915
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3300
      TabIndex        =   7
      Top             =   915
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
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   450
      TabIndex        =   6
      Top             =   915
      Width           =   1215
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
Dim SQL As String
Dim wUltimoCupom As Double
Dim WGrupoAtualzado As Double
Dim WnumeroPedido As Double
Dim wWhere As String

Private Sub cmdImprimir1_Click()
    
    If txtSerie.Text <> "S2" Then
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
        
        If txtSerie.Text = "CF" Then
            MsgBox "Cupom Fiscal não pode ser REIMPRESSO.", vbCritical, "Atenção"
            txtNotaFiscal.SetFocus
            txtSerie.Text = ""
            txtValorNF.Text = ""
            Exit Sub
        End If
    
    
    'SQL = "Select * From Nfitens " _
             & "Where numeroped = " & txtPedido.Text & "" _
             & " order by Item"
             
     SQL = "Select * From Nfcapa " _
             & "Where nf = " & txtNotaFiscal.Text & " and serie = " & txtSerie.Text & "" _
             & " order by Item"

             RsDados.CursorLocation = adUseClient
             RsDados.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

             If RsDados.EOF = False Then
                 wNotaFiscalReemissao = RsDados("NF")
                 wSerieReemissao = Trim(RsDados("Serie"))

                 If wNotaFiscalReemissao <> 0 Then
                    If MsgBox("Deseja Reemitir a Nota Fiscal Nº " & wNotaFiscalReemissao, vbYesNo + vbQuestion, "Atenção") = vbYes Then
                       RsDados.Close
                       EmiteNotafiscal wNotaFiscalReemissao, wSerieReemissao
                       cmdImprimir1.Enabled = False
                    Exit Sub
                    End If
                 Else
                    MsgBox "Nota fiscal não encontrada Nº " & wNotaFiscalReemissao, vbCritical, "Atenção"
                    RsDados.Close
                    Exit Sub
                 End If
             Else
             MsgBox "Nota fiscal não encontrada Nº " & wNotaFiscalReemissao, vbCritical, "Atenção"
              
             RsDados.Close
             Exit Sub
             End If

             RsDados.Close
 
    
    txtNumero.Text = ""
    txtSerie.Text = ""
    txtValorNF.Text = ""
    txtNotaFiscal.Text = ""
    txtNotaFiscal.SetFocus
End Sub


Private Sub cmdRetorna_Click()
    Unload Me
End Sub

Private Sub cmdRetorna1_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 3
    txtNotaFiscal.Text = ""
    txtSerie.Text = ""
    txtValorNF.Text = ""
End Sub



Private Sub txtNotaFiscal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtSerie.SetFocus
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




Private Sub txtSerie_LostFocus()
 txtSerie.Text = UCase(txtSerie.Text)
If txtSerie.Text = "" Then
    Exit Sub
End If

' If txtSerie.Text <> "S2" Then
'    Exit Sub
' End If

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
        cmdImprimir1.Enabled = True
    Else
        MsgBox "NF não encontrada", vbInformation, "Aviso"
        cmdImprimir1.Enabled = False
        txtSerie.Text = ""
        txtNotaFiscal.SelStart = 0
        txtNotaFiscal.SelLength = Len(txtNotaFiscal.Text)
        txtNotaFiscal.SetFocus
End If
ADOCancela.Close

End Sub

