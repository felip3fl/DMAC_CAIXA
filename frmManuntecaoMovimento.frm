VERSION 5.00
Object = "{D76D7120-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7u.ocx"
Begin VB.Form frmManuntecaoMovimento 
   BackColor       =   &H00000000&
   Caption         =   "Manunteção Movimento"
   ClientHeight    =   7860
   ClientLeft      =   1800
   ClientTop       =   2280
   ClientWidth     =   17955
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   17955
   Begin VSFlex7UCtl.VSFlexGrid grdMovimentosDisponiveis 
      Height          =   6870
      Left            =   300
      TabIndex        =   0
      Top             =   600
      Width           =   6345
      _cx             =   11192
      _cy             =   12118
      _ConvInfo       =   1
      Appearance      =   2
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmManuntecaoMovimento.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   4
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
   Begin VB.Label lblCabec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Controle Movimento"
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
      TabIndex        =   1
      Top             =   200
      Width           =   2175
   End
End
Attribute VB_Name = "frmManuntecaoMovimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
