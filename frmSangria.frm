VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7u.ocx"
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmSangria 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "T. Numerário"
   ClientHeight    =   9090
   ClientLeft      =   885
   ClientTop       =   1650
   ClientWidth     =   18900
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   18900
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameDividiModalidade 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   15195
      TabIndex        =   32
      Top             =   585
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtValorNovoModalidade 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   990
         TabIndex        =   36
         Text            =   "0,00"
         Top             =   585
         Width           =   1455
      End
      Begin VB.CheckBox chkDividirModalidade 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Dividir modalidade (Administrador)"
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
         Height          =   300
         Left            =   300
         TabIndex        =   35
         Top             =   195
         Width           =   3495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   285
         TabIndex        =   33
         Top             =   630
         Width           =   510
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Left            =   0
         Top             =   0
         Width           =   4170
      End
   End
   Begin VB.Frame frmSaldo 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   12390
      TabIndex        =   21
      Top             =   6645
      Visible         =   0   'False
      Width           =   9075
      Begin VB.CheckBox chkAtualizaSaldoFuturo 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Atualizar saldos futuro"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3285
         TabIndex        =   29
         Top             =   90
         Width           =   2235
      End
      Begin VB.TextBox txtSaldoDiferenca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1550
         TabIndex        =   27
         Text            =   "0,00"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtSaldoTela 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1550
         TabIndex        =   26
         Text            =   "0,00"
         Top             =   150
         Width           =   1455
      End
      Begin VB.TextBox txtSaldoNovo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1550
         TabIndex        =   22
         Text            =   "0,00"
         Top             =   495
         Width           =   1455
      End
      Begin VSFlex7UCtl.VSFlexGrid grdAuditorMovimento 
         Height          =   4380
         Left            =   150
         TabIndex        =   24
         Top             =   2250
         Visible         =   0   'False
         Width           =   7245
         _cx             =   12779
         _cy             =   7726
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   14737632
         ForeColor       =   4210752
         BackColorFixed  =   0
         ForeColorFixed  =   16777215
         BackColorSel    =   3421236
         ForeColorSel    =   16777215
         BackColorBkg    =   0
         BackColorAlternate=   12632256
         GridColor       =   14737632
         GridColorFixed  =   8421504
         TreeColor       =   8421504
         FloodColor      =   16777215
         SheetBorder     =   8421504
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSangria.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   -2147483633
         ForeColorFrozen =   4210752
         WallPaperAlignment=   4
      End
      Begin Balcao2010.chameleonButton cmdAtualizarSaldo 
         Height          =   585
         Left            =   3270
         TabIndex        =   31
         Top             =   540
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1032
         BTYPE           =   14
         TX              =   "Atualizar Saldo"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   2500134
         BCOLO           =   4210752
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   5263440
         MPTR            =   1
         MICON           =   "frmSangria.frx":00AC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   6810
         Left            =   0
         Top             =   0
         Width           =   7575
      End
      Begin VB.Label lblPrevia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Atual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   150
         TabIndex        =   30
         Top             =   1890
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diferença:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   150
         TabIndex        =   28
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblSaldo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   150
         TabIndex        =   25
         Top             =   150
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Novo Saldo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   150
         TabIndex        =   23
         Top             =   495
         Width           =   1290
      End
   End
   Begin VSFlex7UCtl.VSFlexGrid grdAnaliticoSangria 
      Height          =   2655
      Left            =   10515
      TabIndex        =   20
      Top             =   4095
      Visible         =   0   'False
      Width           =   3555
      _cx             =   6271
      _cy             =   4683
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   14737632
      ForeColor       =   4210752
      BackColorFixed  =   0
      ForeColorFixed  =   16777215
      BackColorSel    =   3421236
      ForeColorSel    =   16777215
      BackColorBkg    =   0
      BackColorAlternate=   12632256
      GridColor       =   14737632
      GridColorFixed  =   8421504
      TreeColor       =   8421504
      FloodColor      =   16777215
      SheetBorder     =   8421504
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSangria.frx":00C8
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   -2147483633
      ForeColorFrozen =   4210752
      WallPaperAlignment=   4
   End
   Begin VSFlex7UCtl.VSFlexGrid grdAnaliticoVenda 
      Height          =   5040
      Left            =   5610
      TabIndex        =   11
      Top             =   1995
      Visible         =   0   'False
      Width           =   4995
      _cx             =   8811
      _cy             =   8890
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   14737632
      ForeColor       =   4210752
      BackColorFixed  =   0
      ForeColorFixed  =   16777215
      BackColorSel    =   3421236
      ForeColorSel    =   16777215
      BackColorBkg    =   0
      BackColorAlternate=   12632256
      GridColor       =   14737632
      GridColorFixed  =   8421504
      TreeColor       =   8421504
      FloodColor      =   16777215
      SheetBorder     =   8421504
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSangria.frx":0129
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   -2147483633
      ForeColorFrozen =   4210752
      WallPaperAlignment=   4
   End
   Begin VB.Frame frameDataAdministrador 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1320
      Left            =   5610
      TabIndex        =   14
      Top             =   825
      Width           =   7350
      Begin VB.PictureBox picAvancar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DrawStyle       =   5  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2715
         MouseIcon       =   "frmSangria.frx":01C5
         Picture         =   "frmSangria.frx":04CF
         ScaleHeight     =   375
         ScaleWidth      =   240
         TabIndex        =   16
         ToolTipText     =   "Avança"
         Top             =   555
         Width           =   240
      End
      Begin VB.PictureBox picVoltar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DrawStyle       =   5  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   915
         MouseIcon       =   "frmSangria.frx":075E
         Picture         =   "frmSangria.frx":0A68
         ScaleHeight     =   375
         ScaleWidth      =   240
         TabIndex        =   15
         ToolTipText     =   "Retorna"
         Top             =   555
         Width           =   240
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   315
         Left            =   1305
         TabIndex        =   17
         Top             =   585
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alteração sangria de outros dias (Administrador)"
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
         Left            =   285
         TabIndex        =   19
         Top             =   195
         Width           =   4065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   300
         TabIndex        =   18
         Top             =   630
         Width           =   435
      End
      Begin VB.Image fraFechamentoAnterior 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Left            =   0
         Top             =   0
         Width           =   4770
      End
   End
   Begin VB.ComboBox cmbGrupoAuxiliar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   525
      TabIndex        =   10
      Top             =   6075
      Width           =   4620
   End
   Begin VB.TextBox txtReforco 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   4080
      TabIndex        =   5
      Text            =   "0"
      Top             =   7125
      Width           =   1050
   End
   Begin VB.TextBox txtTNnumero 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1635
      TabIndex        =   0
      Text            =   "0"
      Top             =   6720
      Width           =   555
   End
   Begin VB.TextBox txtRetirada 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   300
      Left            =   4080
      TabIndex        =   1
      Text            =   "0"
      Top             =   6720
      Width           =   1050
   End
   Begin VSFlex7UCtl.VSFlexGrid grdMovimentoCaixa 
      Height          =   5040
      Left            =   300
      TabIndex        =   9
      Top             =   825
      Width           =   5055
      _cx             =   8916
      _cy             =   8890
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   14737632
      ForeColor       =   4210752
      BackColorFixed  =   0
      ForeColorFixed  =   16777215
      BackColorSel    =   3421236
      ForeColorSel    =   16777215
      BackColorBkg    =   0
      BackColorAlternate=   12632256
      GridColor       =   14737632
      GridColorFixed  =   8421504
      TreeColor       =   8421504
      FloodColor      =   16777215
      SheetBorder     =   8421504
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSangria.frx":0CF6
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   -2147483633
      ForeColorFrozen =   4210752
      WallPaperAlignment=   4
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdModalidadeVenda 
      Height          =   1080
      Left            =   10665
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   4170
      _cx             =   7355
      _cy             =   1905
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   14737632
      ForeColor       =   4210752
      BackColorFixed  =   0
      ForeColorFixed  =   16777215
      BackColorSel    =   3421236
      ForeColorSel    =   16777215
      BackColorBkg    =   0
      BackColorAlternate=   12632256
      GridColor       =   14737632
      GridColorFixed  =   8421504
      TreeColor       =   8421504
      FloodColor      =   16777215
      SheetBorder     =   8421504
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSangria.frx":0D95
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   5
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   -2147483633
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin VB.Image imgLogo 
      Height          =   1005
      Left            =   15975
      Top             =   3825
      Width           =   3000
   End
   Begin VB.Label lblAnalitico 
      BackColor       =   &H00000000&
      Caption         =   "Analítico"
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
      Height          =   255
      Left            =   5610
      TabIndex        =   34
      Top             =   585
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label lblModalidade 
      BackColor       =   &H00000000&
      Caption         =   "Troca de Modalidade"
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
      Height          =   255
      Left            =   10635
      TabIndex        =   13
      Top             =   585
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   885
      Left            =   300
      Top             =   6645
      Width           =   5055
   End
   Begin VB.Image imgSelecao 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   300
      Top             =   5985
      Width           =   5055
   End
   Begin VB.Label lblRetirada 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Retirada"
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
      Left            =   300
      TabIndex        =   8
      Top             =   585
      Width           =   735
   End
   Begin VB.Label lblReforco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reforço de Caixa"
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
      Height          =   405
      Left            =   525
      TabIndex        =   7
      Top             =   7185
      Width           =   1485
   End
   Begin VB.Label lblTipoReforco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dinheiro "
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
      Height          =   405
      Left            =   2070
      TabIndex        =   6
      Top             =   7185
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T.Numerário "
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
      Height          =   405
      Left            =   525
      TabIndex        =   4
      Top             =   6765
      Width           =   1110
   End
   Begin VB.Label lblCabec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transfencia de Numerário / Reforço de Caixa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   300
      TabIndex        =   3
      Top             =   200
      Width           =   4770
   End
   Begin VB.Label lbEspecie 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dinheiro"
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
      Height          =   615
      Left            =   2265
      TabIndex        =   2
      Top             =   6765
      Width           =   1770
   End
End
Attribute VB_Name = "frmSangria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wSubTotal As Double
Dim wSubTotal_S As Double
Dim wSubTotal2 As Double
Dim wTotalEntrada As Double
Dim wSubTotalEntfin As Double
Dim wSubTotalEntFat As Double
Dim wTotalSaldo As Double
Dim wTotalSaldo_S As Double
Dim wControlacor As Long
Dim wConfigCor As Long
Dim sql As String
Dim Cor As String
Dim Cor1 As String
Dim Cor2 As String
Dim Cor3 As String
Dim wData As Date
Dim wGuardaRow As Integer
Dim wTotalTipoNota As Double
Dim wVenda As Double
Dim wCancelamento As Double
Dim wDevolucao As Double
Dim wTR As Double
Dim wValorGrid As Double
Dim wGrupoAux As Long
Dim wWhere As String
Dim Idx As Integer

Dim rdoDataFechamentoRetaguarda As New ADODB.Recordset

Private Sub ChkModoAdministrador_Click()
'  Call CarregaMovimentocaixa
'  Call BuscaTransNumerico
'  Call CarregaModalidadeVenda
End Sub

Private Sub chkAtualizaSaldoFuturo_Click()
    Call carregaGridAuditor
End Sub

Private Sub chkDividirModalidade_Click()
    txtValorNovoModalidade.text = "0,00"
    txtValorNovoModalidade.Enabled = True
End Sub

Private Sub cmbGrupoAuxiliar_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
  End If
End Sub

Private Sub cmbSair_Click()
 Unload Me
End Sub

Private Sub cmdRetornar_Click()
  Unload Me
End Sub

Private Sub cmdRetornar_KeyPress(KeyAscii As Integer)
Unload Me
End Sub


Private Sub cmdAtualizarSaldo_Click()

    If chkAtualizaSaldoFuturo.Value = 0 Then
        sql = "update movimentocaixa set mc_valor = " & ConverteVirgula(Format(txtSaldoNovo.text, "##,###0.00")) & vbNewLine & _
              "where mc_grupo = '11006' and mc_protocolo = " & rdoDataFechamentoRetaguarda("protocolo") & " and mc_data = '" & Format(mskData.text, "YYYY/MM/DD") & "'"
        notificacaoEmail "Modificação de saldo realizado! Data: " & Format(mskData.text, "DD/MM/YYYY") & " - Valor: " & Format(txtSaldoNovo.text, "##,###0.00")
    Else
        sql = "update movimentocaixa set mc_valor = mc_valor + (" & Replace(txtSaldoDiferenca.text, ",", ".") & ")" & vbNewLine & _
              "where mc_grupo = '11006' and mc_nroCaixa = " & GLB_Caixa & " and mc_data >= '" & Format(mskData.text, "YYYY/MM/DD") & "'"
        notificacaoEmail "Modificação de saldo realizado! Data: " & Format(mskData.text, "DD/MM/YYYY") & " até " & Format(GLB_DataInicial, "DD/MM/YYYY") & " - Valor: " & Format(txtSaldoDiferenca.text, "##,###0.00")
    End If
    'MsgBox sql
    rdoCNLoja.Execute sql
    
    notificacaoEmail sql
    
    frmSaldo.Visible = False
        Call CarregaMovimentocaixa
        Call BuscaTransNumerico
        Call CarregaModalidadeVenda
    
End Sub



Private Sub Form_Load()
  
    defineImpressora
  
    frmSaldo.BackColor = vbBlack
  
  If GLB_Administrador = True Then
    mskData.text = Date
    frameDataAdministrador.Visible = True
    frameDividiModalidade.top = grdModalidadeVenda.top
    grdModalidadeVenda.top = (grdModalidadeVenda.top + (frameDividiModalidade.Height))
  Else
    frameDataAdministrador.Visible = False
    frameDividiModalidade.Visible = False
  End If
  
  
  grdAnaliticoVenda.top = 825

  Call AjustaTela(frmSangria)
  
  grdMovimentoCaixa.Row = 14
  For i = 0 To grdMovimentoCaixa.Cols - 1
    grdMovimentoCaixa.Col = i
    grdMovimentoCaixa.CellBackColor = &HC0C0FF
  Next i

  Call CarregaMovimentocaixa
  Call BuscaTransNumerico
  Call CarregaModalidadeVenda
  
  'frameDividiModalidade.top = (grdModalidadeVenda.top + grdModalidadeVenda.Width)
  frameDividiModalidade.left = (grdModalidadeVenda.left)
  
  grdAnaliticoVenda.Height = grdMovimentoCaixa.Height
  grdModalidadeVenda.Height = grdMovimentoCaixa.Height
  
  grdAnaliticoSangria.Height = grdMovimentoCaixa.Height
  grdAnaliticoSangria.left = grdAnaliticoVenda.left
  grdAnaliticoSangria.top = grdMovimentoCaixa.top
  
  grdAnaliticoVenda.Height = grdMovimentoCaixa.Height
  grdAnaliticoVenda.Height = grdMovimentoCaixa.Height
  
  frmSaldo.left = frameDataAdministrador.left
  frmSaldo.top = frameDataAdministrador.top
  
  imgLogo.Picture = LoadPicture(endIMG("logo"))
  
End Sub

Private Sub CarregaMovimentocaixa()

  grdMovimentoCaixa.Rows = 1
  grdMovimentoCaixa.AddItem "Dinheiro" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50101"             '1   1   1
  grdMovimentoCaixa.AddItem "Cheque" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50201"               '2   2   2
  grdMovimentoCaixa.AddItem "Visa" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50301"                 '3   3   3
  grdMovimentoCaixa.AddItem "MasterCard" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50302"           '4   4   4
  grdMovimentoCaixa.AddItem "Amex" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50303"                 '5   5   5
  grdMovimentoCaixa.AddItem "BNDES" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50304"                '6   6   6
  grdMovimentoCaixa.AddItem "Rede Shop" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50203"            '7   7   7
  grdMovimentoCaixa.AddItem "Visa Elec." & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50206"           '7   8   8
  grdMovimentoCaixa.AddItem "Nota de Credito" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50701"      '8   9   9
  grdMovimentoCaixa.AddItem "Hypercard" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50205"            '9   10  10
  grdMovimentoCaixa.AddItem "Entrada Faturada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50502"     '10  11  11
  grdMovimentoCaixa.AddItem "Entrada Financiada" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50602"   '11  12  12
  grdMovimentoCaixa.AddItem "Reforço de Caixa" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50801"     '12  13  13
  grdMovimentoCaixa.AddItem "* Garantia Estendida" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50009"           '14
  grdMovimentoCaixa.AddItem ""                                                                               '13  14  15
  grdMovimentoCaixa.AddItem "** Saldo Anterior**"                                                            '14  15  16
  grdMovimentoCaixa.AddItem "Dinheiro" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50806"             '16  16  17
  grdMovimentoCaixa.AddItem "Cheque" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50807"               '17  17  18
  grdMovimentoCaixa.AddItem "Total Saldo Anterior" & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & 0 & Chr(9) & "50803" '18  18  19
  grdMovimentoCaixa.AddItem ""                                                                               '19  19  20
  grdMovimentoCaixa.AddItem "*** TOTAL CAIXA"                                                                '20  20  21

  wTotalSaldo = 0
  wTotalSaldo_S = 0
  wSubTotal = 0
  wSubTotal_S = 0
  grdMovimentoCaixa.Row = 1

    grdAnaliticoVenda.Visible = False
    lblAnalitico.Visible = False
    grdAnaliticoSangria.Visible = False

    If rdoDataFechamentoRetaguarda.State = 1 Then
        rdoDataFechamentoRetaguarda.Close
    End If

    If GLB_Administrador = False Then
        sql = "Select Max(CTr_DataInicial)as DataMov,Max(Ctr_Protocolo) as protocolo " _
        & "from ControleCaixa where CTR_Supervisor <> 99 and CTr_NumeroCaixa = " & GLB_Caixa
    Else
        sql = "Select Max(CTr_DataInicial)as DataMov,Max(Ctr_Protocolo) as protocolo " _
        & "from ControleCaixa where CTR_Supervisor <> 99 and CTr_NumeroCaixa = " & GLB_Caixa _
        & "and substring(convert(char(10),CTr_DataInicial,111),1,10) = '" & Format(mskData.text, "YYYY/MM/DD") & "'"
    End If
        
        rdoDataFechamentoRetaguarda.CursorLocation = adUseClient
        rdoDataFechamentoRetaguarda.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
 
        If IsNull(rdoDataFechamentoRetaguarda("DataMov")) = True Then
            Exit Sub
        End If
 
 sql = ("select mc_Grupo,sum(MC_Valor) as TotalModalidade,Count(*) as Quantidade from movimentocaixa" _
       & " Where MC_NumeroEcf = " & GLB_ECF & " and MC_NroCaixa=" & GLB_Caixa & " and MC_Protocolo = " & rdoDataFechamentoRetaguarda("protocolo") _
       & " and MC_Data ='" & Format(rdoDataFechamentoRetaguarda("DataMov"), "yyyy/mm/dd") & "' and  MC_Serie <> '00' and (MC_Grupo like '10%' or MC_Grupo like '11%'" _
       & " or MC_Grupo like '50%' or MC_Grupo like '20%') and MC_TipoNota in ('V','T','E','S') group by mc_grupo")
       rdoFormaPagamento.CursorLocation = adUseClient
       rdoFormaPagamento.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
       
  wData = Format(rdoDataFechamentoRetaguarda("datamov"), "yyyy/mm/dd")
 'rdoDataFechamentoRetaguarda.Close
  
  If Not rdoFormaPagamento.EOF Then
     Do While Not rdoFormaPagamento.EOF
        If rdoFormaPagamento("MC_Grupo") = "10101" Then
           grdMovimentoCaixa.TextMatrix(1, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10201" Then
           grdMovimentoCaixa.TextMatrix(2, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("mc_grupo") = "10301" Then
           grdMovimentoCaixa.TextMatrix(3, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10302" Then
           grdMovimentoCaixa.TextMatrix(4, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10303" Then
           grdMovimentoCaixa.TextMatrix(5, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10304" Then
           grdMovimentoCaixa.TextMatrix(6, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10203" Then
           grdMovimentoCaixa.TextMatrix(7, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10206" Then
           grdMovimentoCaixa.TextMatrix(8, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10701" Then
           grdMovimentoCaixa.TextMatrix(9, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10205" Then
           grdMovimentoCaixa.TextMatrix(10, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "11004" Then
           grdMovimentoCaixa.TextMatrix(11, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
           wTotalEntrada = (wTotalEntrada + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "11005" Then
           grdMovimentoCaixa.TextMatrix(12, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
           wTotalEntrada = (wTotalEntrada + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "10801" Then
           grdMovimentoCaixa.TextMatrix(13, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal = (wSubTotal + rdoFormaPagamento("TotalModalidade"))
           
        ElseIf rdoFormaPagamento("MC_Grupo") = "11009" Then
           grdMovimentoCaixa.TextMatrix(14, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           'wtotalGarantia = (wtotalGarantia + rdoFormaPagamento("TotalModalidade"))
           wTotalSaldo = (wSubTotal - rdoFormaPagamento("TotalModalidade"))
           
        ElseIf rdoFormaPagamento("MC_Grupo") = "11006" Then
           grdMovimentoCaixa.TextMatrix(17, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wTotalSaldo = (wTotalSaldo + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "11007" Then
           grdMovimentoCaixa.TextMatrix(18, 1) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wTotalSaldo = (wTotalSaldo + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50101" Then
           grdMovimentoCaixa.TextMatrix(1, 2) = CDbl(grdMovimentoCaixa.TextMatrix(1, 2)) + CDbl(Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00"))
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50201" Then
           grdMovimentoCaixa.TextMatrix(2, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("mc_grupo") = "50301" Then
           grdMovimentoCaixa.TextMatrix(3, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50302" Then
           grdMovimentoCaixa.TextMatrix(4, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50303" Then
           grdMovimentoCaixa.TextMatrix(5, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50304" Then
           grdMovimentoCaixa.TextMatrix(6, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50203" Then
           grdMovimentoCaixa.TextMatrix(7, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50206" Then
           grdMovimentoCaixa.TextMatrix(8, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50701" Then
           grdMovimentoCaixa.TextMatrix(9, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50205" Then
           grdMovimentoCaixa.TextMatrix(10, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50502" Then
           grdMovimentoCaixa.TextMatrix(11, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50602" Then
           grdMovimentoCaixa.TextMatrix(12, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
         ElseIf rdoFormaPagamento("MC_Grupo") = "50801" Then
           grdMovimentoCaixa.TextMatrix(13, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
           
        ElseIf rdoFormaPagamento("MC_Grupo") = "50009" Then
        
           grdMovimentoCaixa.TextMatrix(1, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           grdMovimentoCaixa.TextMatrix(14, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
           
           'wtotalGarantia = (wtotalGarantia + rdoFormaPagamento("TotalModalidade"))
           
         'ElseIf rdoFormaPagamento("MC_Grupo") = "50804" Then
           'grdMovimentoCaixa.TextMatrix(15, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           'wTotalSaldo_S = (wTotalSaldo_S + rdoFormaPagamento("TotalModalidade"))
         ElseIf rdoFormaPagamento("MC_Grupo") = "50806" Then
           grdMovimentoCaixa.TextMatrix(17, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wTotalSaldo_S = (wTotalSaldo_S + rdoFormaPagamento("TotalModalidade"))
         ElseIf rdoFormaPagamento("MC_Grupo") = "50807" Then
           grdMovimentoCaixa.TextMatrix(18, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wTotalSaldo_S = (wTotalSaldo_S + rdoFormaPagamento("TotalModalidade"))
        ElseIf rdoFormaPagamento("MC_Grupo") = "50803" Then
           grdMovimentoCaixa.TextMatrix(19, 2) = Format(rdoFormaPagamento("TotalModalidade"), "###,###,###,##0.00")
           wSubTotal_S = (wSubTotal_S + rdoFormaPagamento("TotalModalidade"))
        
        End If
       rdoFormaPagamento.MoveNext
     Loop
     
     grdMovimentoCaixa.TextMatrix(1, 3) = Format((grdMovimentoCaixa.TextMatrix(1, 1) - grdMovimentoCaixa.TextMatrix(1, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(2, 3) = Format((grdMovimentoCaixa.TextMatrix(2, 1) - grdMovimentoCaixa.TextMatrix(2, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(3, 3) = Format((grdMovimentoCaixa.TextMatrix(3, 1) - grdMovimentoCaixa.TextMatrix(3, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(4, 3) = Format((grdMovimentoCaixa.TextMatrix(4, 1) - grdMovimentoCaixa.TextMatrix(4, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(5, 3) = Format((grdMovimentoCaixa.TextMatrix(5, 1) - grdMovimentoCaixa.TextMatrix(5, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(6, 3) = Format((grdMovimentoCaixa.TextMatrix(6, 1) - grdMovimentoCaixa.TextMatrix(6, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(7, 3) = Format((grdMovimentoCaixa.TextMatrix(7, 1) - grdMovimentoCaixa.TextMatrix(7, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(8, 3) = Format((grdMovimentoCaixa.TextMatrix(8, 1) - grdMovimentoCaixa.TextMatrix(8, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(9, 3) = Format((grdMovimentoCaixa.TextMatrix(9, 1) - grdMovimentoCaixa.TextMatrix(9, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(10, 3) = Format((grdMovimentoCaixa.TextMatrix(10, 1) - grdMovimentoCaixa.TextMatrix(10, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(11, 3) = Format((grdMovimentoCaixa.TextMatrix(11, 1) - grdMovimentoCaixa.TextMatrix(11, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(12, 3) = Format((grdMovimentoCaixa.TextMatrix(12, 1) - grdMovimentoCaixa.TextMatrix(12, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(13, 3) = Format((grdMovimentoCaixa.TextMatrix(13, 1) - grdMovimentoCaixa.TextMatrix(13, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(14, 3) = Format((grdMovimentoCaixa.TextMatrix(14, 1) - grdMovimentoCaixa.TextMatrix(14, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(17, 3) = Format((grdMovimentoCaixa.TextMatrix(17, 1) - grdMovimentoCaixa.TextMatrix(17, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(18, 3) = Format((grdMovimentoCaixa.TextMatrix(18, 1) - grdMovimentoCaixa.TextMatrix(18, 2)), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(19, 3) = Format((grdMovimentoCaixa.TextMatrix(19, 1) - grdMovimentoCaixa.TextMatrix(19, 2)), "###,###,###,##0.00")
     
     grdMovimentoCaixa.TextMatrix(19, 1) = Format(wTotalSaldo, "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(19, 2) = Format(wTotalSaldo_S, "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(19, 3) = Format((wTotalSaldo - wTotalSaldo_S), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(21, 1) = Format((wSubTotal + wTotalSaldo), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(21, 2) = Format((wSubTotal_S + wTotalSaldo_S), "###,###,###,##0.00")
     grdMovimentoCaixa.TextMatrix(21, 3) = Format(((wSubTotal + wTotalSaldo) - (wSubTotal_S + wTotalSaldo_S)), "###,###,###,##0.00")

  End If
  
  rdoFormaPagamento.Close
  wSubTotal = 0
  wSubTotal_S = 0
  
End Sub

Private Sub grdAnaliticoSangria_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
     grdAnaliticoSangria.Visible = False
     grdModalidadeVenda.Visible = False
     lblModalidade.Visible = False
  End If
End Sub

Private Sub grdAnaliticoSangria_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 46 And GLB_Administrador = True Then
        If MsgBox("Deseja deleta a Transferencia de valor R$ " & grdAnaliticoSangria.TextMatrix(grdAnaliticoSangria.Row, 0) & " do movimento?", vbYesNo + vbQuestion, "Atenção") = vbYes Then
            sql = "delete movimentocaixa " & vbNewLine & _
                  "where mc_grupo = '" & grdMovimentoCaixa.TextMatrix(grdMovimentoCaixa.Row, 4) & "'" & vbNewLine & _
                  "and mc_nrocaixa = '" & GLB_Caixa & "'" & vbNewLine & _
                  "and MC_Protocolo = '" & rdoDataFechamentoRetaguarda("protocolo") & "'" & vbNewLine & _
                  "and MC_Sequencia = '" & grdAnaliticoSangria.TextMatrix(grdAnaliticoSangria.Row, 2) & "'" & vbNewLine & _
                  "and MC_valor = " & ConverteVirgula(grdAnaliticoSangria.TextMatrix(grdAnaliticoSangria.Row, 0)) & ""
            rdoCNLoja.Execute (sql)
            grdMovimentoCaixa_DblClick
        End If
  End If
End Sub

Private Sub grdAnaliticoVenda_DblClick()
  If grdAnaliticoVenda.Rows > 1 Then
        If MsgBox("Deseja Imprimir Analítico?", vbYesNo + vbQuestion, "Atenção") = vbYes Then
           Call ImprimeAnaliticoVenda
        Else
           If MsgBox("Deseja Trocar Modalidade ", vbYesNo + vbQuestion, "Atenção") = vbYes Then
              lblModalidade.Visible = True
              If GLB_Administrador = True Then frameDividiModalidade.Visible = True
              grdModalidadeVenda.Visible = True
              grdModalidadeVenda.SetFocus
              Exit Sub
           End If
        End If
  End If
End Sub

Private Sub grdAnaliticoVenda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
     grdAnaliticoVenda.Visible = False
     lblAnalitico.Visible = False
      frameDividiModalidade.Visible = False
     grdModalidadeVenda.Visible = False
     lblModalidade.Visible = False
  End If
End Sub

Private Sub grdAuditorMovimento_DblClick()
    mskData.text = Format(grdAuditorMovimento.TextMatrix(grdAuditorMovimento.Row, 0), "DD/MM/YYYY")
    mskData_KeyPress 13
    frmSaldo.Visible = False
End Sub

Private Sub grdAuditorMovimento_KeyPress(KeyAscii As Integer)
        
    If KeyAscii = 27 Then
        frmSaldo.Visible = False
    End If
End Sub

Private Sub grdModalidadeVenda_DblClick()
    If chkDividirModalidade.Value = 1 Then
        If txtValorNovoModalidade.text > 0 Then
            Call DividiModalidade
        Else
            MsgBox "Valor inválido.", vbExclamation, "Divisão de Modalidade"
        End If
    Else
        Call TrocaModalidade
    End If
End Sub

Private Sub grdModalidadeVenda_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
      frameDividiModalidade.Visible = False
     grdModalidadeVenda.Visible = False
     lblModalidade.Visible = False
     grdAnaliticoVenda.SetFocus
  End If
End Sub

Private Sub grdMovimentoCaixa_DblClick()

If GLB_Administrador = True Then

    'wCampoAdminstrador = grdMovimentoCaixa.TextMatrix(grdMovimentoCaixa.Row, 0)

    If grdMovimentoCaixa.TextMatrix(grdMovimentoCaixa.Row, 4) = "50806" And grdMovimentoCaixa.Col = 1 Then

        txtSaldoTela.text = grdMovimentoCaixa.TextMatrix(grdMovimentoCaixa.Row, 1)

        'txtSaldoNovo.Text = Empty
        txtSaldoNovo.text = txtSaldoTela.text
        frmSaldo.Visible = True
        txtSaldoNovo.SetFocus
        carregaGridAuditor
        
        
        'GravaSaldoCaixa

    End If
    
End If

    grdAnaliticoSangria.Visible = False
    grdAnaliticoVenda.Visible = False
    lblAnalitico.Visible = False
    If grdMovimentoCaixa.Col = 1 Then
        Call CarregaAnaliticoVenda
    ElseIf grdMovimentoCaixa.Col = 2 Then
        Call CarregaAnaliticoSangria
    End If
    
End Sub

Private Sub grdMovimentoCaixa_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
  End If
End Sub

Private Sub sklDataMovimento_Click()

End Sub

Private Sub mskData_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CarregaMovimentocaixa
        Call BuscaTransNumerico
        Call CarregaModalidadeVenda
    End If
End Sub

Private Sub picAvancar_Click()
'    If KeyAscii = 13 Then

        mskData.text = Format(CDate(mskData.text) + 1, "DD/MM/YYYY")

        Call CarregaMovimentocaixa
        Call BuscaTransNumerico
        Call CarregaModalidadeVenda
        'rdoDataFechamentoRetaguarda.Close
        
    'End If
End Sub

Private Sub picVoltar_Click()
    'If KeyAscii = 13 Then
        
        mskData.text = Format(CDate(mskData.text) - 1, "DD/MM/YYYY")
    
        Call CarregaMovimentocaixa
        Call BuscaTransNumerico
        Call CarregaModalidadeVenda
        'rdoDataFechamentoRetaguarda.Close
        
    'End If
End Sub

Private Sub txtReforco_GotFocus()
txtReforco.SelStart = 0
txtReforco.SelLength = Len(txtReforco.text)
End Sub

Private Sub txtReforco_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If txtReforco.text > 0 Then
        Call ImprimeReforcoSangria
        Call GravaReforcoCaixa
        Call CarregaMovimentocaixa
        txtReforco.text = 0
        txtReforco.SelStart = 0
        txtReforco.SelLength = Len(txtReforco.text)
        txtReforco.SetFocus

     End If
  End If
  
  If KeyAscii = 27 Then
    Unload Me
  End If
End Sub

Private Sub txtRetirada_GotFocus()
    txtRetirada.SelStart = 0
    txtRetirada.SelLength = Len(txtRetirada.text)
End Sub

Private Sub verificarValorNegativo()
'ricardo versao 01
     Dim rdoSelectCapa As New ADODB.Recordset

    sql = ("Select * from Movimentocaixa Where MC_NumeroECF = " & GLB_ECF & "" _
     & " and  mc_tipoNota = 'V' and mc_data = '" & Format(Date, "yyyy/mm/dd") _
     & "' " & " and MC_Protocolo = " & GLB_CTR_Protocolo & " ")


     rdoSelectCapa.CursorLocation = adUseClient
     rdoSelectCapa.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
     
   If Not rdoSelectCapa.EOF Then
        Do While Not rdoSelectCapa.EOF
            If Mid(rdoSelectCapa("mc_valor"), 1, 1) = "-" Then
               MsgBox "Numero Negativo", vbCritical, "Aviso"
               
               
               Exit Sub
               
            End If
            rdoSelectCapa.MoveNext
       Loop
   End If
     
End Sub

Private Sub txtRetirada_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
  
     If txtRetirada.text = "" Then
        Exit Sub
     End If
     
     verificarValorNegativo
     
     If Mid(txtRetirada.text, 1, 1) = "-" Then
        MsgBox "Valor não pode ser negativo", vbCritical, "Aviso"
        Exit Sub
     End If
     
     If txtRetirada.text > 0 Or GLB_Administrador = True Then
        wValorGrid = grdMovimentoCaixa.TextMatrix(grdMovimentoCaixa.Row, 2)
        If grdMovimentoCaixa.Row = 1 Then wValorGrid = wValorGrid + grdMovimentoCaixa.TextMatrix(14, 3)
        If CDbl(Format((wValorGrid + txtRetirada.text), "0.00")) > _
            CDbl(Format(grdMovimentoCaixa.TextMatrix(grdMovimentoCaixa.Row, 1), "0.00")) Then
            MsgBox "Saldo Insuficiente", vbCritical, "Aviso"
            txtRetirada.text = 0
            txtRetirada.SelStart = 0
            txtRetirada.SelLength = Len(txtRetirada.text)
            txtRetirada.SetFocus
            Exit Sub
        End If

        wGuardaRow = grdMovimentoCaixa.Row
        If txtRetirada.text > 0 Then
            If Mid(cmbGrupoAuxiliar.text, 1, 5) = "30106" Or Mid(cmbGrupoAuxiliar.text, 1, 5) = "30107" Then
                 Call ImprimeReforcoSangria
            End If
            
            Call ImprimeReforcoSangria
        End If
        Call GravaMovimentoCaixa
        Call CarregaMovimentocaixa
        Call BuscaTransNumerico
        Call CarregaModalidadeVenda
        
        grdMovimentoCaixa.Row = wGuardaRow
        txtRetirada.text = 0
        txtRetirada.SelStart = 0
        txtRetirada.SelLength = Len(txtRetirada.text)
        txtRetirada.SetFocus
     End If
     
     If grdMovimentoCaixa.Rows - 1 > grdMovimentoCaixa.Row Then
        grdMovimentoCaixa.Row = grdMovimentoCaixa.Row + 1
        If grdMovimentoCaixa.Row > 17 Then
           grdMovimentoCaixa.Row = 1
        End If
        If (grdMovimentoCaixa.Row = 13) Then
           grdMovimentoCaixa.Row = grdMovimentoCaixa.Row + 2
        End If
        
        If Not grdMovimentoCaixa.RowIsVisible(grdMovimentoCaixa.Row) Then
             grdMovimentoCaixa.TopRow = grdMovimentoCaixa.Row
        End If
    End If
  End If
  
  If KeyAscii = 27 Then
    Unload Me
  End If

End Sub


Private Sub txtSaldoNovo_LostFocus()
    If txtSaldoNovo.text = Empty Then
        txtSaldoNovo.text = "0,00"
    Else
        txtSaldoNovo.text = Format(txtSaldoNovo.text, "#0.00")
        
    End If
    txtSaldoDiferenca.text = Format(CDbl(txtSaldoNovo.text) - CDbl(txtSaldoTela.text), "#0.00")
End Sub

Private Sub txtTNnumero_GotFocus()
  txtTNnumero.text = 0
  txtTNnumero.SelStart = 0
  txtTNnumero.SelLength = Len(txtTNnumero.text)
End Sub

Private Sub txtTNnumero_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    If txtTNnumero.text < 1 Then
       MsgBox "Nro TN Incorreto", vbCritical, "Aviso"
       txtTNnumero.SelStart = 0
       txtTNnumero.SelLength = Len(txtTNnumero.text)
       txtTNnumero.SetFocus
    Else
       txtTNnumero.Enabled = False
       txtReforco.Enabled = False
       txtRetirada.Enabled = True
       grdMovimentoCaixa.Row = 1
       txtRetirada.SelStart = 0
       txtRetirada.SelLength = Len(txtRetirada.text)
       
       txtRetirada.SetFocus
       
    End If
 End If
 
   If KeyAscii = 27 Then
    Unload Me
  End If
    
End Sub

Private Sub ImprimeReforcoSangria()
  Dim wSangrialinha1 As String
  Dim wSangrialinha2 As String
    
    Screen.MousePointer = 11
    
    impressoraRelatorio "[INICIO]"
    
    impressoraRelatorio "________________________________________________"
    impressoraRelatorio "         RETIRADA / REFORÇO DE CAIXA            "
    impressoraRelatorio "                                                "
    impressoraRelatorio left("Loja         " & GLB_Loja & Space(48), 48)
    impressoraRelatorio left("Caixa Nro.   " & GLB_Caixa & Space(48), 48)
    impressoraRelatorio left("Operador     " & GLB_USU_Nome & Space(48), 48)
    impressoraRelatorio left("Protocolo    " & GLB_CTR_Protocolo & Space(48), 48)
    impressoraRelatorio left("Data/Hora    " & Format(Date, "DD/MM/YYYY") & " - " & Format(Time, "HH:MM") & Space(48), 48)
    impressoraRelatorio "                                                "
    impressoraRelatorio "________________________________________________"
 
   
   If txtReforco.text > 0 Then
     
        impressoraRelatorio "                                                "
        impressoraRelatorio "::::::::::::::  REFORÇO DE CAIXA  ::::::::::::::"
        impressoraRelatorio "                                                "
        impressoraRelatorio "                                                "
        impressoraRelatorio left(Space(10) & "R$ " & Format(txtReforco.text, "###,###,##0.00") & Space(38), 48)
        impressoraRelatorio "                                                "

        imprimeCampoGerenteOperador
 
    ElseIf txtRetirada.text > 0 Then
    
        impressoraRelatorio "                                                "
        impressoraRelatorio "::::::::::::  " & UCase(Mid(cmbGrupoAuxiliar.text, 7, 20)) & "  ::::::::::::"
        impressoraRelatorio "                                                "
        impressoraRelatorio "                                                "
        impressoraRelatorio left(Space(10) & "R$ " & Format(txtRetirada.text, "###,###,##0.00") & Space(38), 48)
        impressoraRelatorio "                                                "

        imprimeCampoGerenteOperador
        
    End If
   
   
   
   
   
   
   'TxtFormatado = “Teste de formatação relatório gerencial !!!” + Chr(10) + cItalico + cNegrito + cCondensado + cSublinhado + cExpandido
   'Abre relatório gerencial
        'Retorno = Bematech_FI_AbreRelatorioGerencialMFD("01")
   'impressão texto formatado

   'Encerra relatório gerencial
    impressoraRelatorio "[FIM]"
 
    Screen.MousePointer = 0
End Sub


Private Sub grdMovimentoCaixa_EnterCell()

 If grdMovimentoCaixa.Row > 0 Then
     
    ' If grdMovimentoCaixa.TextMatrix(grdMovimentoCaixa.Row, 0) = "" Then
    '    lbEspecie.Caption = "**Saldo anterior"
    ' Else
        lbEspecie.Caption = grdMovimentoCaixa.TextMatrix(grdMovimentoCaixa.Row, 0)
    ' End If
      
     If grdMovimentoCaixa.TextMatrix(grdMovimentoCaixa.Row, 4) <> "" And _
        grdMovimentoCaixa.TextMatrix(grdMovimentoCaixa.Row, 4) <> "50803" Then
        wGrupoAux = grdMovimentoCaixa.TextMatrix(grdMovimentoCaixa.Row, 4)
        txtRetirada.text = 0
        If txtTNnumero.text > 0 Then
           txtRetirada.Enabled = True
           txtRetirada.SetFocus
        End If
     Else
        wGrupoAux = 0
        txtRetirada.text = ""
        txtRetirada.Enabled = False
        lbEspecie.Caption = ""
        '--
        cmbGrupoAuxiliar.AddItem ""
        cmbGrupoAuxiliar.Clear
        '--
        
        Exit Sub
     End If
     Call CarregaGrupoAuxiliar
 End If
End Sub
Private Sub GravaMovimentoCaixa()

    If GLB_Administrador = True And CDbl(txtRetirada.text) = 0 Then
    
        If MsgBox("Deseja deleta todas as Transferencia desse grupo?", vbYesNo + vbQuestion, "Atenção") = vbYes Then
            sql = "delete movimentocaixa " & vbNewLine & _
                  "where mc_grupo = '" & grdMovimentoCaixa.TextMatrix(grdMovimentoCaixa.Row, 4) & "'" & vbNewLine & _
                  "and mc_nrocaixa = '" & GLB_Caixa & "'" & vbNewLine & _
                  "and MC_Protocolo = '" & rdoDataFechamentoRetaguarda("protocolo") & "'"
                  
                  rdoCNLoja.Execute (sql)
        End If
    
    Else
    
        sql = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja, MC_Data, MC_Grupo,MC_SubGrupo, MC_Documento,MC_Serie," _
                     & "MC_Valor, MC_banco, MC_Agencia,MC_Contacorrente, MC_bomPara, MC_Parcelas, MC_Remessa,MC_SituacaoEnvio," _
                     & "MC_NroCaixa,MC_Protocolo,MC_GrupoAuxiliar,MC_Pedido,MC_DataProcesso,MC_TipoNota)" _
                     & " values(" & GLB_ECF & ",'" & GLB_USU_Codigo & "','" & GLB_Loja & "','" _
                     & Format(wData, "yyyy/mm/dd") & "'," & grdMovimentoCaixa.TextMatrix(grdMovimentoCaixa.Row, 4) _
                     & ",''," & txtTNnumero.text & ",'TN'," _
                     & ConverteVirgula(Format(txtRetirada.text, "##,###0.00")) & ", " _
                     & "0,0,0,0,0,9,'A','" & GLB_Caixa & "'," & rdoDataFechamentoRetaguarda("protocolo") & "," & Mid(cmbGrupoAuxiliar.text, 1, 5) & "," _
                     & "'" & pedido & "','" & Format(wData, "yyyy/mm/dd") & "','V')"
        rdoCNLoja.Execute (sql)
        
    End If
      
End Sub
Private Sub GravaReforcoCaixa()
  sql = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja, MC_Data, MC_Grupo,MC_SubGrupo, MC_Documento,MC_Serie," _
               & "MC_Valor, MC_banco, MC_Agencia,MC_Contacorrente, MC_bomPara, MC_Parcelas, MC_Remessa,MC_SituacaoEnvio," _
               & "MC_NroCaixa,MC_Protocolo,MC_Pedido, MC_DataProcesso,MC_TipoNota)" _
               & " values(" & GLB_ECF & ",'" & GLB_USU_Codigo & "','" & GLB_Loja & "','" _
               & Format(wData, "yyyy/mm/dd") & "',10801,'',0,'RC'," _
               & ConverteVirgula(Format(txtReforco.text, "##,###0.00")) & ", " _
               & "0,0,0,0,0,9,'A','" & GLB_Caixa & "'," & GLB_CTR_Protocolo & ",'" & pedido & "','" & Format(wData, "yyyy/mm/dd") & "','V')"
      rdoCNLoja.Execute (sql)
End Sub
Private Sub CarregaGrupoAuxiliar()
 '"50701"
   
  cmbGrupoAuxiliar.Clear
  
  If wGrupoAux = 0 Then
     cmbGrupoAuxiliar.AddItem ""
     cmbGrupoAuxiliar.ListIndex = 0
     Exit Sub
  End If
  
  If wGrupoAux = 50101 Or wGrupoAux = 50201 Or wGrupoAux = 50203 Or wGrupoAux = 50205 Or wGrupoAux = 50206 Or wGrupoAux = 50301 Or wGrupoAux = 50302 _
                       Or wGrupoAux = 50303 Or wGrupoAux = 50304 Or wGrupoAux = 50701 Then
      wWhere = " and MO_OrdemApresentacao = '" & wGrupoAux & "'"
  ElseIf wGrupoAux = 50009 Then
      wWhere = " and MO_OrdemApresentacao = '" & wGrupoAux & "'"
      txtRetirada.text = grdMovimentoCaixa.TextMatrix(grdMovimentoCaixa.Row, 1) - grdMovimentoCaixa.TextMatrix(grdMovimentoCaixa.Row, 2)
  ElseIf wGrupoAux = 50204 Or wGrupoAux = 50502 Or wGrupoAux = 50602 Or wGrupoAux = 50804 Then
      'wWhere = " and MO_Grupo in(30101,30201,30107,30106)"
      wWhere = " and MO_OrdemApresentacao = '" & wGrupoAux & "'"
      txtRetirada.text = grdMovimentoCaixa.TextMatrix(grdMovimentoCaixa.Row, 1) - grdMovimentoCaixa.TextMatrix(grdMovimentoCaixa.Row, 2)
  ElseIf wGrupoAux = 50801 Or wGrupoAux = 50806 Then
      wWhere = " and MO_Grupo in(30101,30106)"
  ElseIf wGrupoAux = 50807 Then
      wWhere = " and MO_Grupo in(30201,30107)"
  ElseIf wGrupoAux = 50205 Then
      wWhere = " and MO_Grupo in(50205)"
  End If
  'txtRetirada
  sql = "select * From Modalidade WHERE MO_Grupo LIKE '30%'" & wWhere & " order by Mo_Grupo"
       rdoModalidade.CursorLocation = adUseClient
       rdoModalidade.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
  If Not rdoModalidade.EOF Then
     Do While Not rdoModalidade.EOF
           cmbGrupoAuxiliar.AddItem rdoModalidade("MO_Grupo") & "-" & rdoModalidade("MO_Descricao")
          
        rdoModalidade.MoveNext
     Loop
     cmbGrupoAuxiliar.ListIndex = 0
     Else
     cmbGrupoAuxiliar.AddItem 99999 & " - Cadastrar Ordem de apresentação"
     cmbGrupoAuxiliar.ListIndex = 0
  End If
  wWhere = ""
  rdoModalidade.Close
End Sub

Private Sub BuscaTransNumerico()

    If IsNull(rdoDataFechamentoRetaguarda("protocolo")) = True Then Exit Sub

   sql = "select (max(mc_documento) + 1) as TNum from movimentocaixa where mc_serie = 'TN' and mc_protocolo = " & rdoDataFechamentoRetaguarda("protocolo")
   rdoModalidade.CursorLocation = adUseClient
   rdoModalidade.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
   If IsNull(rdoModalidade("TNum")) = True Then
       txtTNnumero.text = 1
   Else
       txtTNnumero.text = rdoModalidade("TNum")
   End If
   txtTNnumero.Enabled = False
   txtRetirada.Enabled = True
   rdoModalidade.Close
End Sub
'------------------------------------------------------------
'  novo
'-------------------------------------------------------------
Sub CarregaAnaliticoVenda()
    grdAnaliticoVenda.Rows = 1
    grdAnaliticoVenda.Visible = False
    lblAnalitico.Visible = False
    
    sql = "select mc_documento,mc_serie,mo_descricao,mc_valor,mc_Sequencia from movimentocaixa,Modalidade" _
        & " where MC_Grupo=mo_grupo and MC_grupo='" & "1" + Mid(grdMovimentoCaixa.TextMatrix(grdMovimentoCaixa.Row, 4), 2, 5) _
        & "' and MC_Data ='" & Format(wData, "yyyy/mm/dd") & "' and MC_Protocolo = " & rdoDataFechamentoRetaguarda("protocolo") & " and MC_TipoNota in ('V','T','E','S')"
    RsMovimentoCaixa.CursorLocation = adUseClient
    RsMovimentoCaixa.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
   
    If Not RsMovimentoCaixa.EOF Then
        Do While Not RsMovimentoCaixa.EOF
           grdAnaliticoVenda.AddItem RsMovimentoCaixa("mc_documento") & Chr(9) & RsMovimentoCaixa("mc_serie") _
           & Chr(9) & RsMovimentoCaixa("mo_descricao") & Chr(9) & Format(RsMovimentoCaixa("mc_Valor"), "###,###,##0.00") _
           & Chr(9) & RsMovimentoCaixa("mc_sequencia")
           RsMovimentoCaixa.MoveNext
        Loop
        grdAnaliticoVenda.Visible = True
        lblAnalitico.Visible = True
        grdAnaliticoVenda.SetFocus
    End If

    RsMovimentoCaixa.Close
End Sub

Sub CarregaAnaliticoSangria()

    grdAnaliticoSangria.Rows = 1
    grdAnaliticoSangria.Visible = False
    
    Dim Grupo As String
    
    'grdMovimentoCaixa.TextMatrix(grdMovimentoCaixa.Row, 4)
    
    Grupo = grdMovimentoCaixa.TextMatrix(grdMovimentoCaixa.Row, 4)
    If Grupo = "50101" Then
        Grupo = Grupo + ",50009"
    End If
    
    
    sql = "select MC_Valor,MO_Descricao,MC_Sequencia from movimentocaixa, Modalidade" _
        & " where MC_GrupoAuxiliar=mo_grupo and MC_grupo in (" & Grupo _
        & ") and MC_Data ='" & Format(wData, "yyyy/mm/dd") & "' and MC_Protocolo = " & rdoDataFechamentoRetaguarda("protocolo") & " and MC_tipoNota not in ('C')"
    RsMovimentoCaixa.CursorLocation = adUseClient
    RsMovimentoCaixa.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
   
    If Not RsMovimentoCaixa.EOF Then
        Do While Not RsMovimentoCaixa.EOF
           grdAnaliticoSangria.AddItem Format(RsMovimentoCaixa("mc_Valor"), "###,###,##0.00") _
           & Chr(9) & RsMovimentoCaixa("MO_Descricao") _
           & Chr(9) & RsMovimentoCaixa("MC_Sequencia")
           RsMovimentoCaixa.MoveNext
        Loop
        grdAnaliticoSangria.Visible = True
        grdAnaliticoSangria.SetFocus
    End If

    RsMovimentoCaixa.Close
    
End Sub

Sub ImprimeAnaliticoVenda()
 
 
Screen.MousePointer = 11
    Retorno = Bematech_FI_AbreRelatorioGerencialMFD("01")
 
    Retorno = Bematech_FI_UsaRelatorioGerencialMFD("________________________________________________" & _
                   "          RELATORIO ANALITICO DE VENDA          " & _
                   left("Loja " & Format(GLB_Loja, "000") & Space(10), 10) & _
                   right(Space(38) & (Format(Trim(wData), "dd/mm/yyyy")), 38) & _
                   "________________________________________________")
    
    Retorno = Bematech_FI_UsaRelatorioGerencialMFD("                                                " & _
                   left("NF " & Space(10), 10) & left("SERIE" & Space(8), 8) & _
                   left("FORMA PAGAMENTO" & Space(20), 20) & left("VALOR " & Space(10), 10) & _
                   "                                                ")
 
     For Idx = 1 To grdAnaliticoVenda.Rows - 1 Step 1
     
     Retorno = Bematech_FI_UsaRelatorioGerencialMFD(left(grdAnaliticoVenda.TextMatrix(Idx, 0) & Space(10), 10) & _
                   left(grdAnaliticoVenda.TextMatrix(Idx, 1) & Space(8), 8) & _
                   left(grdAnaliticoVenda.TextMatrix(Idx, 2) & Space(20), 20) & _
                   right(Space(10) & Format(grdAnaliticoVenda.TextMatrix(Idx, 3), "###,###,##0.00"), 10))
     Next Idx
    
     Retorno = Bematech_FI_FechaRelatorioGerencial()
 
     Screen.MousePointer = 0
     
End Sub
Private Sub CarregaModalidadeVenda()
    sql = "SELECT MO_Grupo,MO_Descricao FROM Modalidade WHERE MO_Grupo between '10101' and '10701'"
    adoConsulta.CursorLocation = adUseClient
    adoConsulta.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
             
    If Not adoConsulta.EOF Then
        grdModalidadeVenda.Rows = 1
        Do While Not adoConsulta.EOF
            grdModalidadeVenda.AddItem adoConsulta("MO_Grupo") & Chr(9) & adoConsulta("MO_Descricao") & Chr(9)
            
        adoConsulta.MoveNext
        Loop
    End If
    adoConsulta.Close
End Sub
Private Sub TrocaModalidade()
        sql = "update MovimentoCaixa set mc_grupo = " & grdModalidadeVenda.TextMatrix(grdModalidadeVenda.Row, 0) _
            & " where MC_Sequencia = " & grdAnaliticoVenda.TextMatrix(grdAnaliticoVenda.Row, 4)
        rdoCNLoja.Execute sql
        Unload Me
        frmSangria.Show vbModal
End Sub

Private Sub DividiModalidade()
        sql = "exec SP_alterar_dividir_modalidade " & grdAnaliticoVenda.TextMatrix(grdAnaliticoVenda.Row, 4) & "," _
              & grdModalidadeVenda.TextMatrix(grdModalidadeVenda.Row, 0) & "," _
              & ConverteVirgula(txtValorNovoModalidade.text) & ""
              
        rdoCNLoja.Execute sql
        
        notificacaoEmail "Modificação de modalidade realizada! NF " & grdAnaliticoVenda.TextMatrix(grdAnaliticoVenda.Row, 0) _
                         & " - Serie: " & grdAnaliticoVenda.TextMatrix(grdAnaliticoVenda.Row, 1)
        notificacaoEmail "Forma Pagamento original: " & grdAnaliticoVenda.TextMatrix(grdAnaliticoVenda.Row, 2) _
                         & "Valor Original: " & grdAnaliticoVenda.TextMatrix(grdAnaliticoVenda.Row, 3)
        notificacaoEmail "Nova Modalidade: " & grdModalidadeVenda.TextMatrix(grdModalidadeVenda.Row, 1) _
                         & "Novo Valor: " & txtValorNovoModalidade.text
        notificacaoEmail sql
        
        Unload Me
        frmSangria.Show vbModal
        
End Sub

Private Sub carregaGridAuditor()
   Dim rdoDataFechamentoRetaguarda2 As New ADODB.Recordset
        Dim saldo As Double
        Dim saldoAnterior As Double
        Dim saldo_s As Double
        Dim saldoAnterior_s As Double
        Dim saldoProximoDia As Double
        Dim msgStatus As String
        
        grdAuditorMovimento.Rows = 1

        sql = "select CTR_DataInicial, CTR_DataFinal, CTR_NumeroCaixa, CTR_Protocolo, USU_Nome from controlecaixa,UsuarioCaixa " _
        & "where CTR_Operador = USU_Codigo and CTR_Supervisor <> 99 and  USU_TipoUsuario = 'O' and CTR_NumeroCaixa = '" & GLB_Caixa & "' " _
        & "and CTR_DataInicial >= '" & Format(mskData.text, "YYYY/MM/DD") & "' " _
        & "order by CTR_Protocolo"
        
        rdoDataFechamentoRetaguarda2.CursorLocation = adUseClient
        rdoDataFechamentoRetaguarda2.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

        
        Do While Not rdoDataFechamentoRetaguarda2.EOF
        
                sql = ("select mc_Grupo,mc_GrupoAuxiliar,sum(MC_Valor) as TotalModalidade,Count(*) as Quantidade from movimentocaixa" _
                      & " Where MC_Protocolo in (" & rdoDataFechamentoRetaguarda2("CTR_Protocolo") _
                      & ") and  MC_Serie <> '00' and MC_tiponota <> 'C' and mc_grupo in ('10101','50101','11006','50806','50009','10801','50801') group by mc_grupo,mc_GrupoAuxiliar")
                
                rdoFechamentoGeral.CursorLocation = adUseClient
                rdoFechamentoGeral.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                
                saldo = 0
                saldoAnterior = 0
                saldo_s = 0
                saldoAnterior_s = 0
                
                If Not rdoFechamentoGeral.EOF Then
                    grdAuditorMovimento.Visible = True
                    Do While Not rdoFechamentoGeral.EOF
                        If rdoFechamentoGeral("MC_Grupo") = "10101" Then
                            saldo = saldo + rdoFechamentoGeral("TotalModalidade")
                        ElseIf rdoFechamentoGeral("MC_Grupo") = "50101" Or rdoFechamentoGeral("MC_Grupo") = "50009" Then
                            saldo_s = saldo_s + rdoFechamentoGeral("TotalModalidade")
                            
                        'REFORCO DE CAIXA - - - - - - - - - - - - - - - - - - - - - - - - - -
                        ElseIf rdoFechamentoGeral("MC_Grupo") = "10801" Then
                            saldo = saldo + rdoFechamentoGeral("TotalModalidade")
                        ElseIf rdoFechamentoGeral("MC_Grupo") = "50801" And rdoFechamentoGeral("mc_GrupoAuxiliar") = "50801" Then
                            saldo_s = saldo_s + rdoFechamentoGeral("TotalModalidade")
                        ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                        
                        ElseIf rdoFechamentoGeral("MC_Grupo") = "11006" Then
                            saldoAnterior = saldoAnterior + rdoFechamentoGeral("TotalModalidade")
                            
                            If mskData.text = left(rdoDataFechamentoRetaguarda2("CTR_DataInicial"), 10) Then
                                saldoAnterior = CDbl(txtSaldoNovo.text)
                            ElseIf chkAtualizaSaldoFuturo.Value = 1 Then
                                saldoAnterior = saldoAnterior + CDbl(txtSaldoDiferenca.text)
                            End If
                            
                            If CStr(saldoAnterior) = CStr(saldoProximoDia) Then
                                msgStatus = "OK"
                            ElseIf saldoProximoDia = 0 Then
                                msgStatus = "OK"
                            Else
                                msgStatus = "Saldo inicial com Divergência! "
                            End If
                            
                        ElseIf rdoFechamentoGeral("MC_Grupo") = "50806" Then
                            saldoAnterior_s = saldoAnterior_s + rdoFechamentoGeral("TotalModalidade")
                        End If
                        rdoFechamentoGeral.MoveNext
                    Loop
                Else
                    grdAuditorMovimento.Visible = False
                End If
                rdoFechamentoGeral.Close
                saldoProximoDia = (saldo - saldo_s) + (saldoAnterior - saldoAnterior_s)
    
        
            grdAuditorMovimento.AddItem Format(rdoDataFechamentoRetaguarda2("CTR_DataInicial"), "DD/MM/YY HH:MM") & Chr(9) _
                                            & Format(rdoDataFechamentoRetaguarda2("CTR_Protocolo"), "###00") & Chr(9) _
                                            & Format(saldoAnterior, "#0.00") & Chr(9) _
                                            & Format(saldoProximoDia, "#0.00") & Chr(9) _
                                            & msgStatus
                                            
            If msgStatus <> "OK" Then
                grdAuditorMovimento.Row = grdAuditorMovimento.Rows - 1
                For i = 0 To grdAuditorMovimento.Cols - 1
                    grdAuditorMovimento.Col = i
                    grdAuditorMovimento.CellForeColor = vbRed
                Next i
            End If
                                            
            rdoDataFechamentoRetaguarda2.MoveNext
            
        Loop
        
        grdAuditorMovimento.Row = 0
        rdoDataFechamentoRetaguarda2.Close
        cmdAtualizarSaldo.SetFocus
End Sub

Private Sub txtSaldoNovo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        Call carregaGridAuditor
        
    ElseIf KeyAscii = 27 Then
        frmSaldo.Visible = False
    End If


End Sub

Private Sub txtValorNovoModalidade_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNumerico(KeyAscii)
End Sub

Private Sub txtValorNovoModalidade_LostFocus()
    If txtValorNovoModalidade.text = Empty Then
        txtValorNovoModalidade.text = "0,00"
    Else
        txtValorNovoModalidade.text = Format(txtValorNovoModalidade.text, "##,###0.00")
    End If
End Sub

