VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "ACTSKN43.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmParametroCaixa 
   BackColor       =   &H00FBD5D2&
   Caption         =   "Parametros Gerais"
   ClientHeight    =   5220
   ClientLeft      =   2670
   ClientTop       =   2880
   ClientWidth     =   5790
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   5220
   ScaleWidth      =   5790
   Begin TabDlg.SSTab SSTab1 
      Height          =   4845
      Left            =   105
      TabIndex        =   0
      Top             =   60
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   8546
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   16504274
      TabCaption(0)   =   "Caixa"
      TabPicture(0)   =   "frmParametroCaixa.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdParametroCaixa"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdRetornar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdGravar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraLogin"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmParametroCaixa.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmParametroCaixa.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "frmParametroCaixa.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "frmParametroCaixa.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Tab 5"
      TabPicture(5)   =   "frmParametroCaixa.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      Begin VB.Frame fraLogin 
         ForeColor       =   &H00FF0000&
         Height          =   975
         Left            =   165
         TabIndex        =   3
         Top             =   3150
         Width           =   5250
         Begin VB.TextBox txtNroCaixa 
            Height          =   315
            Left            =   2625
            TabIndex        =   10
            Top             =   210
            Width           =   585
         End
         Begin VB.TextBox txtNroECF 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   4530
            TabIndex        =   9
            Top             =   210
            Width           =   585
         End
         Begin VB.CheckBox chkCaixaECF 
            Caption         =   "Caixa ECF"
            Height          =   195
            Left            =   150
            TabIndex        =   8
            Top             =   645
            Width           =   1035
         End
         Begin VB.CheckBox chkCaixaSN 
            Caption         =   "Caixa SN"
            Height          =   195
            Left            =   1605
            TabIndex        =   7
            Top             =   645
            Width           =   960
         End
         Begin VB.CheckBox chkCaixa00 
            Caption         =   "Caixa 00"
            Height          =   195
            Left            =   2835
            TabIndex        =   6
            Top             =   645
            Width           =   915
         End
         Begin VB.CheckBox chkCaixaSM 
            Caption         =   "Caixa SM"
            Height          =   195
            Left            =   4155
            TabIndex        =   5
            Top             =   645
            Width           =   975
         End
         Begin VB.TextBox txtLoja 
            Height          =   315
            Left            =   540
            TabIndex        =   4
            Top             =   225
            Width           =   630
         End
         Begin ACTIVESKINLibCtl.SkinLabel sklNumerodoECF 
            Height          =   270
            Left            =   3630
            OleObjectBlob   =   "frmParametroCaixa.frx":00A8
            TabIndex        =   11
            Top             =   315
            Width           =   825
         End
         Begin ACTIVESKINLibCtl.SkinLabel sklNroCaixa 
            Height          =   195
            Left            =   1620
            OleObjectBlob   =   "frmParametroCaixa.frx":011A
            TabIndex        =   12
            Top             =   315
            Width           =   915
         End
         Begin ACTIVESKINLibCtl.SkinLabel sklLoja 
            Height          =   195
            Left            =   150
            OleObjectBlob   =   "frmParametroCaixa.frx":0190
            TabIndex        =   13
            Top             =   330
            Width           =   330
         End
      End
      Begin VB.CommandButton cmdGravar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Gravar"
         Height          =   435
         Left            =   3090
         TabIndex        =   2
         Top             =   4260
         Width           =   1155
      End
      Begin VB.CommandButton cmdRetornar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Retornar"
         Height          =   435
         Left            =   4245
         TabIndex        =   1
         Top             =   4260
         Width           =   1155
      End
      Begin MSFlexGridLib.MSFlexGrid grdParametroCaixa 
         Height          =   2535
         Left            =   150
         TabIndex        =   14
         Top             =   600
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4471
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   16504274
         BackColorBkg    =   16504274
         GridColor       =   16761024
         GridColorFixed  =   16761024
         FillStyle       =   1
         FormatString    =   "<Loja |<Nro Caixa|<Nro ECF|<Caixa ECF|<Caixa SN|<Caixa 00|<Caixa SM"
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   210
      OleObjectBlob   =   "frmParametroCaixa.frx":01F6
      Top             =   4485
   End
End
Attribute VB_Name = "frmParametroCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim RsDados As New ADODB.Recordset
Dim SQL As String
Dim wnomecampo As String
Dim wtipocampo As String

Private Sub chkCaixa00_Click()
    If chkCaixa00.Value = 1 Then
       chkCaixa00.Tag = "S"
    Else
       chkCaixa00.Tag = "N"
    End If


End Sub

Private Sub chkCaixaECF_Click()
    If chkCaixaECF.Value = 1 Then
       chkCaixaECF.Tag = "S"
    Else
       chkCaixaECF.Tag = "N"
    End If
End Sub

Private Sub chkCaixaSM_Click()
    If chkCaixaSM.Value = 1 Then
       chkCaixaSM.Tag = "S"
    Else
       chkCaixaSM.Tag = "N"
    End If
End Sub

Private Sub chkCaixaSN_Click()
    If chkCaixaSN.Value = 1 Then
       chkCaixaSN.Tag = "S"
    Else
       chkCaixaSN.Tag = "N"
    End If
End Sub

Private Sub cmdGravar_Click()
rdoCNLoja.BeginTrans
Screen.MousePointer = vbHourglass

SQL = "Insert Into ParametroCaixa (PAR_Loja," _
    & "PAR_NroCaixa,PAR_NroECF,PAR_CaixaECF,PAR_CaixaSN,PAR_Caixa00,PAR_CaixaSM) " _
    & "Values ('" & txtLoja.Text & "'," & txtNroCaixa.Text & "," & txtNroECF.Text & ",'" _
    & chkCaixaECF.Tag & "','" & chkCaixaSN.Tag & "','" & chkCaixa00.Tag & "','" & chkCaixaSM.Tag & "')"
    rdoCNLoja.Execute SQL
    Screen.MousePointer = vbNormal
    rdoCNLoja.CommitTrans

    Call CarregaGrid
    Call LimpaTela

End Sub

Private Sub cmdRetornar_Click()
' Unload Me
wnomecampo = "PAR_CampoTeste      "
wtipocampo = "char (10) Default 'A1' "
rdoCNLoja.BeginTrans
Screen.MousePointer = vbHourglass

SQL = "Alter Table ParametroCaixa Add " _
      & wnomecampo & wtipocampo
rdoCNLoja.Execute SQL
Screen.MousePointer = vbNormal
rdoCNLoja.CommitTrans
End Sub

Private Sub Form_Load()
  Skin1.LoadSkin "c:\WINDOWS\system\skin.skn"
  Skin1.ApplySkin Me.hwnd
  Left = (Screen.Width - Width) / 2
  Top = (Screen.Height - Height) / 2
  Call CarregaGrid
  Call LimpaTela
End Sub
Private Sub CarregaGrid()
  grdParametroCaixa.Rows = 1
  SQL = ("Select * from ParametroCaixa")
  RsDados.CursorLocation = adUseClient
  RsDados.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

  
  If Not RsDados.EOF Then
     Do While Not RsDados.EOF
        grdParametroCaixa.AddItem RsDados("PAR_Loja") & Chr(9) _
        & RsDados("PAR_NroCaixa") & Chr(9) & RsDados("PAR_NroECF") & Chr(9) _
        & RsDados("PAR_CaixaECF") & Chr(9) & RsDados("PAR_CaixaSN") & Chr(9) _
        & RsDados("PAR_Caixa00") & Chr(9) & RsDados("PAR_CaixaSM") & Chr(9)
        RsDados.MoveNext
     Loop
  End If
  RsDados.Close
End Sub
Private Sub LimpaTela()
  chkCaixaECF.Value = False
  chkCaixaECF.Value = False
  chkCaixaECF.Value = False
  chkCaixaECF.Value = False
  chkCaixaSN.Tag = "N"
  chkCaixaECF.Tag = "N"
  chkCaixa00.Tag = "N"
  chkCaixaSM.Tag = "N"
  txtLoja.Text = " "
  txtNroCaixa.Text = 0
  txtNroECF.Text = 0
End Sub


Private Sub grdParametroCaixa_DblClick()
  
  rdoCNLoja.BeginTrans
  Screen.MousePointer = vbHourglass

  
  SQL = "Delete Parametrocaixa where PAR_NroCaixa=" _
       & grdParametroCaixa.TextMatrix(grdParametroCaixa.Row, 1)
       
  rdoCNLoja.Execute SQL
  Screen.MousePointer = vbNormal
  rdoCNLoja.CommitTrans

  
  Call CarregaGrid
End Sub

