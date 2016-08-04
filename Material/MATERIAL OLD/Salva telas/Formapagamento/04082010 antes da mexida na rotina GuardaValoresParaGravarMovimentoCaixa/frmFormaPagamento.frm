VERSION 5.00
Begin VB.Form frmFormaPagamento 
   BorderStyle     =   0  'None
   Caption         =   "Forma de Pagamento"
   ClientHeight    =   9000
   ClientLeft      =   4425
   ClientTop       =   1170
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFormaPagamento.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmEstorno 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1470
      Left            =   60
      TabIndex        =   50
      Top             =   735
      Visible         =   0   'False
      Width           =   1755
      Begin Balcao2010.chameleonButton chbEstorno 
         Height          =   450
         Left            =   3045
         TabIndex        =   67
         Top             =   2880
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   794
         BTYPE           =   13
         TX              =   "chameleonButton1"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
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
         MICON           =   "frmFormaPagamento.frx":AFCC2
         PICN            =   "frmFormaPagamento.frx":AFCDE
         PICH            =   "frmFormaPagamento.frx":B0930
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtEstorno 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   1470
         TabIndex        =   61
         Text            =   "Text1"
         Top             =   2940
         Width           =   1185
      End
      Begin VB.CommandButton chbDinheiroEst 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   975
         Picture         =   "frmFormaPagamento.frx":B1582
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   990
         Width           =   750
      End
      Begin VB.CommandButton chbVisaElectronEst 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2505
         Picture         =   "frmFormaPagamento.frx":B44A4
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   990
         Width           =   750
      End
      Begin VB.CommandButton chbRedeShopEst 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3270
         Picture         =   "frmFormaPagamento.frx":B7912
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   990
         Width           =   750
      End
      Begin VB.CommandButton chbVisaEst 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   975
         Picture         =   "frmFormaPagamento.frx":BC688
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   1545
         Width           =   750
      End
      Begin VB.CommandButton chbHiperCardEst 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3285
         Picture         =   "frmFormaPagamento.frx":BF78A
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   1560
         Width           =   750
      End
      Begin VB.CommandButton chbChequeEst 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1740
         Picture         =   "frmFormaPagamento.frx":C2604
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   990
         Width           =   750
      End
      Begin VB.CommandButton chbBNDESESt 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4035
         Picture         =   "frmFormaPagamento.frx":C5106
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   990
         Width           =   750
      End
      Begin VB.CommandButton chbMasterCardEst 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1740
         Picture         =   "frmFormaPagamento.frx":C7A88
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   1545
         Width           =   750
      End
      Begin VB.CommandButton chbNotaCreditoEst 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4050
         Picture         =   "frmFormaPagamento.frx":CA40A
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   1560
         Width           =   750
      End
      Begin VB.CommandButton chbAmexEst 
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2505
         Picture         =   "frmFormaPagamento.frx":CD08C
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1560
         Width           =   750
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Para"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   3030
         TabIndex        =   66
         Top             =   2115
         Width           =   600
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "De"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1485
         TabIndex        =   65
         Top             =   2130
         Width           =   330
      End
      Begin VB.Label lblPara 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Label3"
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   3015
         TabIndex        =   64
         Top             =   2385
         Width           =   1350
      End
      Begin VB.Label lblDe 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Label2"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1470
         TabIndex        =   63
         Top             =   2370
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Valor estorno"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1485
         TabIndex        =   62
         Top             =   2730
         Width           =   1245
      End
   End
   Begin Balcao2010.chameleonButton chbteste 
      Height          =   525
      Left            =   3030
      TabIndex        =   45
      Top             =   3660
      Visible         =   0   'False
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   926
      BTYPE           =   3
      TX              =   "567,00"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16558443
      BCOLO           =   16558443
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFormaPagamento.frx":CF8E2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraNModalidades 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2370
      Left            =   600
      TabIndex        =   39
      Top             =   5355
      Width           =   4725
      Begin VB.CommandButton chbEstornaFormaPagto 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Estorna "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   2385
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1605
         Width           =   1155
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   3570
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1605
         Width           =   1155
      End
      Begin VB.CommandButton chbAmex 
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   2385
         Picture         =   "frmFormaPagamento.frx":CF8FE
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   825
         Width           =   1155
      End
      Begin VB.CommandButton chbNotaCredito 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   1200
         Picture         =   "frmFormaPagamento.frx":D2154
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1605
         Width           =   1155
      End
      Begin VB.CommandButton chbMasterCard 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   1200
         Picture         =   "frmFormaPagamento.frx":D4DD6
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   825
         Width           =   1155
      End
      Begin VB.CommandButton chbBNDES 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   15
         Picture         =   "frmFormaPagamento.frx":D7758
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1605
         Width           =   1155
      End
      Begin VB.CommandButton chbCheque 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   1200
         Picture         =   "frmFormaPagamento.frx":DA0DA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   45
         Width           =   1155
      End
      Begin VB.CommandButton chbHiperCard 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   3570
         Picture         =   "frmFormaPagamento.frx":DCBDC
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   825
         Width           =   1155
      End
      Begin VB.CommandButton chbVisa 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   15
         Picture         =   "frmFormaPagamento.frx":DFA56
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   825
         Width           =   1155
      End
      Begin VB.CommandButton chbRedeShop 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   3570
         Picture         =   "frmFormaPagamento.frx":E2B58
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   45
         Width           =   1155
      End
      Begin VB.CommandButton chbVisaElectron 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   2385
         Picture         =   "frmFormaPagamento.frx":E78CE
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   45
         Width           =   1155
      End
      Begin VB.CommandButton chbDinheiro 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   15
         Picture         =   "frmFormaPagamento.frx":EAD3C
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   45
         Width           =   1155
      End
   End
   Begin VB.TextBox lblParc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1470
      TabIndex        =   38
      Text            =   "Parcelas"
      Top             =   3420
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.CommandButton cmdCondicao 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Condição "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   540
      MaskColor       =   &H00C0E0FF&
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6255
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame frmcond 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2325
      Left            =   -1815
      TabIndex        =   27
      Top             =   -480
      Visible         =   0   'False
      Width           =   4980
      Begin VB.TextBox lblTootip1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1275
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   600
         Width           =   2610
      End
      Begin VB.TextBox lblTootip 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   225
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   345
         Width           =   3720
      End
      Begin VB.TextBox lblValorTotalPedido 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1995
         TabIndex        =   34
         Text            =   "lblValorTotalPedido"
         Top             =   60
         Width           =   2265
      End
      Begin VB.TextBox lblTotalPedido 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   255
         TabIndex        =   33
         Text            =   "Total do Pedido R$"
         Top             =   60
         Width           =   1695
      End
      Begin VB.Frame fraRecebimento 
         BackColor       =   &H00C00000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   180
         TabIndex        =   28
         Top             =   885
         Visible         =   0   'False
         Width           =   4140
         Begin VB.TextBox lblFatFin 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            BorderStyle     =   0  'None
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
            Height          =   420
            Left            =   135
            TabIndex        =   32
            Top             =   210
            Width           =   1965
         End
         Begin VB.TextBox lblEntrada 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Left            =   135
            TabIndex        =   31
            Top             =   720
            Width           =   1620
         End
         Begin VB.TextBox lblValorFatFin 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            BorderStyle     =   0  'None
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
            Height          =   405
            Left            =   2775
            TabIndex        =   30
            Top             =   225
            Width           =   1260
         End
         Begin VB.TextBox lblApagar 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   480
            Left            =   2685
            TabIndex        =   29
            Top             =   705
            Width           =   1350
         End
      End
   End
   Begin VB.Frame FraParcelas 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Parcelas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1470
      TabIndex        =   25
      Top             =   4170
      Visible         =   0   'False
      Width           =   735
      Begin VB.TextBox txtParcelas 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   105
         MaxLength       =   2
         TabIndex        =   26
         Top             =   105
         Width           =   480
      End
   End
   Begin VB.TextBox txtNaotemtroco 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   525
      TabIndex        =   24
      Top             =   3405
      Width           =   1965
   End
   Begin VB.TextBox txtIdentificadequeTelaqueveio 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   23
      Top             =   180
      Visible         =   0   'False
      Width           =   660
   End
   Begin Balcao2010.chameleonButton chbOK 
      Height          =   495
      Left            =   5115
      TabIndex        =   12
      Top             =   8250
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
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
      BCOL            =   14933984
      BCOLO           =   14933984
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFormaPagamento.frx":EDC5E
      PICN            =   "frmFormaPagamento.frx":EDC7A
      PICH            =   "frmFormaPagamento.frx":EE8CC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtValorModalidade 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3015
      TabIndex        =   21
      Top             =   4470
      Width           =   2340
   End
   Begin VB.TextBox txtValoraPagar 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3015
      TabIndex        =   20
      Top             =   1350
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.TextBox txtValorPago 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3015
      TabIndex        =   19
      Top             =   2085
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.TextBox txtValorFalta 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3015
      TabIndex        =   18
      Top             =   2820
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.TextBox txtValorTroco 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3015
      TabIndex        =   17
      Top             =   3630
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.TextBox txtTipoNota 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   16
      Top             =   195
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.TextBox txtPedido 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   15
      Top             =   195
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.TextBox txtSerie 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1335
      TabIndex        =   14
      Top             =   195
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.TextBox txtNroNotaFiscal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   435
      TabIndex        =   13
      Top             =   210
      Visible         =   0   'False
      Width           =   855
   End
   Begin Balcao2010.chameleonButton chbTroco 
      Height          =   810
      Left            =   630
      TabIndex        =   44
      Top             =   3600
      Visible         =   0   'False
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1429
      BTYPE           =   13
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFormaPagamento.frx":EF51E
      PICN            =   "frmFormaPagamento.frx":EF53A
      PICH            =   "frmFormaPagamento.frx":F196C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton chbValorFalta 
      Height          =   525
      Left            =   3000
      TabIndex        =   46
      Top             =   2805
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   926
      BTYPE           =   3
      TX              =   "0"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16558443
      BCOLO           =   16558443
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFormaPagamento.frx":F3D9E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton chbValorPago 
      Height          =   525
      Left            =   3030
      TabIndex        =   47
      Top             =   2055
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   926
      BTYPE           =   3
      TX              =   "0"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16558443
      BCOLO           =   15002065
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFormaPagamento.frx":F3DBA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton chbValoraPagar 
      Height          =   525
      Left            =   3015
      TabIndex        =   48
      Top             =   1335
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   926
      BTYPE           =   3
      TX              =   "0"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16558443
      BCOLO           =   16558443
      FCOL            =   4210752
      FCOLO           =   4210752
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFormaPagamento.frx":F3DD6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton chbSair 
      Height          =   495
      Left            =   5115
      TabIndex        =   49
      Top             =   135
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
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
      BCOL            =   14933984
      BCOLO           =   14933984
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFormaPagamento.frx":F3DF2
      PICN            =   "frmFormaPagamento.frx":F3E0E
      PICH            =   "frmFormaPagamento.frx":F4660
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblTroco 
      AutoSize        =   -1  'True
      BackColor       =   &H00AE7411&
      BackStyle       =   0  'Transparent
      Caption         =   "Troco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0002B9EE&
      Height          =   360
      Left            =   1440
      TabIndex        =   43
      Top             =   3885
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label lblFaltaPagar 
      AutoSize        =   -1  'True
      BackColor       =   &H00AE7411&
      BackStyle       =   0  'Transparent
      Caption         =   "A Pagar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0002B9EE&
      Height          =   360
      Left            =   675
      TabIndex        =   42
      Top             =   3015
      Width           =   1125
   End
   Begin VB.Label lblValorPago 
      AutoSize        =   -1  'True
      BackColor       =   &H00AE7411&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0002B9EE&
      Height          =   360
      Left            =   675
      TabIndex        =   41
      Top             =   2250
      Width           =   1560
   End
   Begin VB.Label lblTotalaPagar 
      AutoSize        =   -1  'True
      BackColor       =   &H00AE7411&
      BackStyle       =   0  'Transparent
      Caption         =   "Total a Pagar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0002B9EE&
      Height          =   360
      Left            =   675
      TabIndex        =   40
      Top             =   1560
      Width           =   1860
   End
   Begin VB.Label lblModalidade 
      AutoSize        =   -1  'True
      BackColor       =   &H00AE7411&
      BackStyle       =   0  'Transparent
      Caption         =   "Modalidade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0002B9EE&
      Height          =   360
      Left            =   600
      TabIndex        =   22
      Top             =   4635
      Width           =   1650
   End
End
Attribute VB_Name = "frmFormaPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Asteristico As Boolean
Dim AbreTela As Long
Dim NroPedido As Long
Dim SQL As String
'Dim wtotalitens As Long
'Dim wTotalVenda As Long
'Dim NroItens As Long
Dim ContadorItens As Long
Dim CodigoModalidade As String
Dim ValoraPagar As Double
Dim ValorPago As Double
Dim ValorFalta As Double
Dim ValorTroco As Double
Dim ValorDinheiro As Double
Dim ValorModalidade As Double
Dim wValorCodigoZero As Double

Dim ValorFaturada As Double

Dim wValorDados As String
Dim wSequencia As Integer
Dim wCodigo As Integer
Dim wPegaPessoaCli As String
Dim EntFaturada As Double

Dim EntFinanciada As Double
Dim wTotalNota As Double
Dim wTotalNotaFatFin As Double

Dim NomeCartao As String

Dim ValCartaoVisa As Double
Dim ValCartaoMastercard As Double
Dim ValCartaoAmex As Double
Dim ValCartaoDiners As Double
Dim ValorPagamentoCartao As Double

Dim ValTEFVisaElectron As Double
Dim ValTEFRedeShop As Double
Dim ValTEFHiperCard As Double

Dim ValNotaCredito As Double
Dim AvistaReceber As Double
Dim Modalidade As Double
Dim WParcelas As Integer
Dim wIndicePreco As Double
Dim wVerificaAVR As Boolean
Dim wGrupo As Long
Dim ValorFinanciada As Double

Dim wCodigoModalidadeDINHEIRO As String
Dim WCodigoModalidadeAMEX As String
Dim WCodigoModalidadeCHEQUE As String
Dim wCodigoModalidadeDINNERS As String
Dim wCodigoModalidadeMASTERCARD As String
Dim wCodigoModalidadeNOTACREDITO As String
Dim wTEFVisaElectron As String
Dim wTEFRedeShop As String
Dim wTEFHiperCard As String
Dim WCodigoModalidadeVISA As String
Dim wtempo As Long
Dim wMostraGrideCondicao As Boolean

Dim wControle As String
Dim wEstorno As Boolean
Dim wEstornoModalidade As String
Dim wControlaEstorno As Boolean

Dim wValorModalidadeIncorreto As Boolean



Private Sub GuardaValoresParaGravarMovimentoCaixa()

wValorModalidadeIncorreto = False

      If wEstorno = True Then
         txtValorModalidade.Text = txtEstorno.Text
      End If
      
      If txtValorModalidade.Text <> "" Then
         If txtValorModalidade.Text <> "," Then
            Modalidade = Format(txtValorModalidade.Text, "0.00")
            
            txtValorPago.Text = CDbl(txtValorPago.Text) + Format(CDbl(Modalidade), "0.00")
            chbValorPago.Caption = Format(txtValorPago.Text, "##,###0.00")
            txtValorFalta.Text = CDbl(txtValoraPagar.Text) - CDbl(txtValorPago.Text)
            chbValorFalta.Caption = Format(txtValorFalta.Text, "##,###0.00")
            
            If txtValorFalta.Text > 0 Then
                txtValorFalta.Text = Format(txtValorFalta.Text, "0.00")
                chbValorFalta.Caption = Format(txtValorFalta.Text, "##,###0.00")
            Else
                txtValorFalta.Text = "0,00"
                chbValorFalta.Caption = Format(txtValorFalta.Text, "##,###0.00")
            End If
            If lblModalidade.Caption = "DINHEIRO" Then
               TotPagar = TotPagar + Modalidade
               ValDinheiro = ValDinheiro + Format(CDbl(Modalidade), "##,###0.00")
            End If
            
            If lblModalidade.Caption = "CHEQUE" Then
               
                  If Val(ConverteVirgula(Format(Modalidade, "0.00"))) < Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) Or _
                     Val(ConverteVirgula(Format(Modalidade, "0.00"))) = Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) Then
                     TotPagar = TotPagar + Modalidade
                     ValCheque = ValCheque + Val(ConverteVirgula(Format(txtValorModalidade.Text, "0.00")))
                  Else
                     MsgBox "O valor no cheque não poderá ser maior que o valor do pedido"
                     wValorModalidadeIncorreto = True
                     txtValorPago.Text = txtValorPago.Text - txtValorModalidade.Text
                     chbValorPago.Caption = Format(txtValorPago.Text, "##,###0.00")
                     txtValorModalidade.Text = ""
                     ZeraVariaveis
                     txtValorModalidade.SetFocus
                     Exit Sub
                  End If
              
            End If
           
            If wTEFVisaElectron = "0401" Then
               TotPagar = TotPagar + Modalidade
               ValTEFVisaElectron = ValTEFVisaElectron + Val(ConverteVirgula(Format(Modalidade, "0.00")))
               wTEFVisaElectron = ""
            End If
            
            If wTEFRedeShop = "0402" Then
               TotPagar = TotPagar + Modalidade
               ValTEFRedeShop = ValTEFRedeShop + Val(ConverteVirgula(Format(Modalidade, "0.00")))
               wTEFRedeShop = ""
            End If
            
            If wTEFHiperCard = "0403" Then
               TotPagar = TotPagar + Modalidade
               ValTEFHiperCard = ValTEFHiperCard + Val(ConverteVirgula(Format(Modalidade, "0.00")))
               wTEFHiperCard = ""
            End If
            
            If lblModalidade.Caption = "NOTA DE CRÉD." Then
               TotPagar = TotPagar + Modalidade
               ValNotaCredito = ValNotaCredito + Val(ConverteVirgula(Format(Modalidade, "0.00")))
            End If
            
            If lblModalidade.Caption = "VISA" Or lblModalidade.Caption = "MASTERCARD" Or _
               lblModalidade.Caption = "AMEX" Or lblModalidade.Caption = "DINNERS" Then
               NomeCartao = lblModalidade.Caption
               
               If NomeCartao = "VISA" Then
                  ValCartaoVisa = ValCartaoVisa + Val(ConverteVirgula(Format(txtValorModalidade.Text, "0.00")))
                  ValorPagamentoCartao = ValorPagamentoCartao + ValCartaoVisa
                 
                  If ValorPagamentoCartao < Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) Or _
                     ValorPagamentoCartao = Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) Then
                     TotPagar = TotPagar + Modalidade
                  Else
                     MsgBox "O valor no cartão não poderá ser maior que o valor do pedido"
                     wValorModalidadeIncorreto = True
                     txtValorPago.Text = txtValorPago.Text - txtValorModalidade.Text
                     chbValorPago.Caption = Format(txtValorPago.Text, "##,###0.00")
                     txtValorModalidade.Text = ""
                     ZeraVariaveis
                     txtValorModalidade.SetFocus
                     Exit Sub
                 End If
               End If
               
               If NomeCartao = "MASTERCARD" Then
                  ValCartaoMastercard = ValCartaoMastercard + Val(ConverteVirgula(Format(txtValorModalidade.Text, "0.00")))
                  ValorPagamentoCartao = ValorPagamentoCartao + ValCartaoMastercard
                  
                  If ValorPagamentoCartao < Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) Or _
                     ValorPagamentoCartao = Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) Then
                     TotPagar = TotPagar + Modalidade
                  Else
                     MsgBox "O valor no cartão não poderá ser maior que o valor do pedido"
                     wValorModalidadeIncorreto = True
                     txtValorPago.Text = txtValorPago.Text - txtValorModalidade.Text
                     chbValorPago.Caption = Format(txtValorPago.Text, "##,###0.00")
                     txtValorModalidade.Text = ""
                     ZeraVariaveis
                     txtValorModalidade.SetFocus
                     Exit Sub
                  End If
               End If

               
               If NomeCartao = "AMEX" Then
                  ValCartaoAmex = ValCartaoAmex + Val(ConverteVirgula(Format(txtValorModalidade.Text, "0.00")))
                  ValorPagamentoCartao = ValorPagamentoCartao + ValCartaoAmex
                  If ValorPagamentoCartao < Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) Or ValorPagamentoCartao = Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) Then
                     TotPagar = TotPagar + Modalidade
                  Else
                     MsgBox "O valor no cartão não poderá ser maior que o valor do pedido"
                     wValorModalidadeIncorreto = True
                     txtValorPago.Text = txtValorPago.Text - txtValorModalidade.Text
                     chbValorPago.Caption = Format(txtValorPago.Text, "##,###0.00")
                     txtValorModalidade.Text = ""
                     ZeraVariaveis
                     txtValorModalidade.SetFocus
                     Exit Sub
                  End If
               End If
               
               If NomeCartao = "DINNERS" Then
                  ValCartaoDiners = ValCartaoDiners + Val(ConverteVirgula(Format(txtValorModalidade.Text, "0.00")))
                  ValorPagamentoCartao = ValorPagamentoCartao + ValCartaoDiners
                  If ValorPagamentoCartao < Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) Or ValorPagamentoCartao = Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) Then
                     TotPagar = TotPagar + Modalidade
                  Else
                     MsgBox "O valor no cartão não poderá ser maior que o valor do pedido"
                     wValorModalidadeIncorreto = True
                     txtValorPago.Text = txtValorPago.Text - txtValorModalidade.Text
                     chbValorPago.Caption = Format(txtValorPago.Text, "##,###0.00")
                     txtValorModalidade.Text = ""
                     ZeraVariaveis
                     txtValorModalidade.SetFocus
                     Exit Sub
                  End If
               End If
               
            End If
            
            txtValorModalidade.Text = Val(ConverteVirgula(Format(Modalidade, "0.00")))
               
            If txtValorFalta.Text <> "0,00" Then 'And EstornaCheque = True Then
              ' TotPagar = TotPagar + ValCheque
               'TotPagar = TotPagar - txtValorfalta.text
            Else
               txtValorPago.Text = Format(TotPagar, "0.00")
               chbValorPago.Caption = Format(txtValorPago.Text, "##,###0.00")
               txtValorFalta.Text = CDbl(txtValoraPagar.Text) - CDbl(TotPagar)
               txtValorFalta.Text = Format(txtValorFalta.Text, "##,###0.00")
               chbValorFalta.Caption = Format(txtValorFalta.Text, "##,###0.00")
            End If
            
            If txtValorFalta.Text < 0 Then
               txtValorFalta.Text = "0,00"
               chbValorFalta.Caption = Format(txtValorFalta.Text, "##,###0.00")
            End If
'-- aonde esta o troco
            txtValorTroco.Text = Val(ConverteVirgula(Format(TotPagar, "0.00"))) - Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00")))
            txtValorTroco.Text = Format(txtValorTroco.Text, "##,###0.00")
            
            If txtValorTroco.Text > ValDinheiro Then
               MsgBox "O valor do troco não poderá ser maior que o valor pago em dinheiro"
               wValorModalidadeIncorreto = True
               txtValorPago.Text = txtValorPago.Text - txtValorModalidade.Text
               chbValorPago.Caption = Format(txtValorPago.Text, "##,###0.00")
               chbTroco.Visible = False
               ZeraVariaveis
               txtValorModalidade.SetFocus
               Exit Sub
            End If
            
            If txtValorTroco.Text > 0 Then
               txtNaotemtroco.Visible = False
               lblTroco.Visible = True
               chbTroco.Visible = True
               chbteste.Visible = True
               chbteste.Caption = txtValorTroco.Text
            End If
            
            'If txtValorTroco.Text > 0 Then
               
            'End If
            
            Call GravarModalidade
            
            If Faturada = True Then
               If EntFaturada <> "0.00" Then
                  If Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) > Format(TotPagar, "0.00") Then
                     txtValorModalidade.Text = ""
                     lblModalidade.Caption = "Modalidade"
                    
                  Else
                     lblFatFin.Visible = True
                     lblFatFin.Text = "FATURADA"
                     lblValorFatFin.Visible = True
                     lblValorFatFin.Text = Format(ValorFaturada, "0.00")
                     txtValorModalidade.Text = ""
                     lblModalidade.Caption = "Modalidade"
                  End If
               End If
            ElseIf EntFinanciada <> "0.00" Then
               If Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) > Format(TotPagar, "0.00") Then
                  txtValorModalidade.Text = ""
                  lblModalidade.Caption = "Modalidade"
                
               Else
                 
                  lblFatFin.Visible = True
                  lblFatFin.Text = "FINANCIADA"
                  lblValorFatFin.Visible = True
                  lblValorFatFin.Text = Format(ValorFinanciada, "0.00")
                  txtValorModalidade.Text = ""
                  lblModalidade.Caption = "Modalidade"
               End If
            Else
               If Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) > Format(TotPagar, "0.00") Then
                  txtValorModalidade.Text = ""
                  lblModalidade.Caption = "Modalidade"
               Else
                  txtValorModalidade.Text = ""
                  lblModalidade.Caption = "Modalidade"
                 
               End If
            End If
          End If
      End If
       txtValorPago.Text = Format(txtValorPago.Text, "0.00")
       chbValorPago.Caption = Format(txtValorPago.Text, "##,###0.00")
     
    ' Call GravarModalidade
     

End Sub

'----------------
Private Sub EstornaValoresParaGravarMovimentoCaixa()
 
      If wEstorno = True Then
         txtValorModalidade.Text = txtEstorno.Text
      End If
 
      If txtEstorno.Text <> "" Then
         If txtEstorno.Text <> "," Then
            Modalidade = Format(txtEstorno.Text, "0.00")
            
            txtValorPago.Text = CDbl(txtValorPago.Text) - Format(CDbl(Modalidade), "0.00")
            chbValorPago.Caption = Format(txtValorPago.Text, "##,###0.00")
            txtValorFalta.Text = CDbl(txtValoraPagar.Text) + CDbl(txtValorPago.Text)
            chbValorFalta.Caption = Format(txtValorFalta.Text, "##,###0.00")
            
            If txtValorFalta.Text > 0 Then
                txtValorFalta.Text = Format(txtValorFalta.Text, "0.00")
                chbValorFalta.Caption = Format(txtValorFalta.Text, "##,###0.00")
            Else
                txtValorFalta.Text = "0,00"
                chbValorFalta.Caption = Format(txtValorFalta.Text, "##,###0.00")
            End If
            If lblModalidade.Caption = "DINHEIRO" Then
               TotPagar = TotPagar - Modalidade
               ValDinheiro = ValDinheiro - Format(CDbl(Modalidade), "##,###0.00")
            End If
            
            If lblModalidade.Caption = "CHEQUE" Then
               
                  If Val(ConverteVirgula(Format(Modalidade, "0.00"))) < Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) Or Val(ConverteVirgula(Format(Modalidade, "0.00"))) = Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) Then
                     TotPagar = TotPagar - Modalidade
                     ValCheque = ValCheque - Val(ConverteVirgula(Format(txtEstorno.Text, "0.00")))
                  Else
                     MsgBox "O valor no cheque não poderá ser maior que o valor do pedido"
                     wValorModalidadeIncorreto = True
                     txtValorPago.Text = txtValorPago.Text + txtEstorno.Text
                     chbValorPago.Caption = Format(txtValorPago.Text, "##,###0.00")
                     txtEstorno.Text = ""
                     ZeraVariaveis
                     txtEstorno.SetFocus
                     Exit Sub
                  End If
              
            End If
           
            If wTEFVisaElectron = "0401" Then
               TotPagar = TotPagar - Modalidade
               ValTEFVisaElectron = ValTEFVisaElectron - Val(ConverteVirgula(Format(Modalidade, "0.00")))
            End If
            
            If wTEFRedeShop = "0402" Then
               TotPagar = TotPagar - Modalidade
               ValTEFRedeShop = ValTEFRedeShop - Val(ConverteVirgula(Format(Modalidade, "0.00")))
            End If
            
            If wTEFHiperCard = "0403" Then
               TotPagar = TotPagar - Modalidade
               ValTEFHiperCard = ValTEFHiperCard - Val(ConverteVirgula(Format(Modalidade, "0.00")))
            End If
            
            
            If lblModalidade.Caption = "NOTA DE CRÉD." Then
               TotPagar = TotPagar - Modalidade
               ValNotaCredito = ValNotaCredito - Val(ConverteVirgula(Format(Modalidade, "0.00")))
            End If
            
            If lblModalidade.Caption = "VISA" Or lblModalidade.Caption = "MASTERCARD" Or lblModalidade.Caption = "AMEX" Or lblModalidade.Caption = "DINNERS" Then
               NomeCartao = lblModalidade.Caption
               
               If NomeCartao = "VISA" Then
                  ValCartaoVisa = ValCartaoVisa - Val(ConverteVirgula(Format(txtEstorno.Text, "0.00")))
                  ValorPagamentoCartao = ValorPagamentoCartao - ValCartaoVisa
                 
                 If ValorPagamentoCartao < Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) Or ValorPagamentoCartao = Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) Then
                     TotPagar = TotPagar - Modalidade
                  Else
                     MsgBox "O valor no cartão não poderá ser maior que o valor do pedido"
                     wValorModalidadeIncorreto = True
                     ZeraVariaveis
                     txtEstorno.SetFocus
                     Exit Sub
                  End If
               End If
               
               If NomeCartao = "MASTERCARD" Then
                  ValCartaoMastercard = ValCartaoMastercard - Val(ConverteVirgula(Format(txtEstorno.Text, "0.00")))
                  ValorPagamentoCartao = ValorPagamentoCartao - ValCartaoMastercard
                  If ValorPagamentoCartao < Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) Or ValorPagamentoCartao = Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) Then
                     TotPagar = TotPagar - Modalidade
                  Else
                     MsgBox "O valor no cartão não poderá ser maior que o valor do pedido"
                     ZeraVariaveis
                     txtEstorno.SetFocus
                     Exit Sub
                  End If
               End If

               
               If NomeCartao = "AMEX" Then
                  ValCartaoAmex = ValCartaoAmex - Val(ConverteVirgula(Format(txtEstorno.Text, "0.00")))
                  ValorPagamentoCartao = ValorPagamentoCartao - ValCartaoAmex
                  If ValorPagamentoCartao < Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) Or ValorPagamentoCartao = Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) Then
                     TotPagar = TotPagar - Modalidade
                  Else
                     MsgBox "O valor no cartão não poderá ser maior que o valor do pedido"
                     ZeraVariaveis
                     txtEstorno.SetFocus
                     Exit Sub
                  End If
               End If
               
               If NomeCartao = "DINNERS" Then
                  ValCartaoDiners = ValCartaoDiners - Val(ConverteVirgula(Format(txtEstorno.Text, "0.00")))
                  ValorPagamentoCartao = ValorPagamentoCartao - ValCartaoDiners
                  If ValorPagamentoCartao < Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) Or ValorPagamentoCartao = Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) Then
                     TotPagar = TotPagar - Modalidade
                  Else
                     MsgBox "O valor no cartão não poderá ser maior que o valor do pedido"
                     ZeraVariaveis
                     txtEstorno.SetFocus
                     Exit Sub
                  End If
               End If
               
            End If
            
            txtEstorno.Text = Val(ConverteVirgula(Format(Modalidade, "0.00")))
               
            If txtValorFalta.Text <> "0,00" Then 'And EstornaCheque = True Then
              ' TotPagar = TotPagar + ValCheque
               'TotPagar = TotPagar - txtValorfalta.text
            Else
               txtValorPago.Text = Format(TotPagar, "0.00")
               chbValorPago.Caption = Format(txtValorPago.Text, "##,###0.00")
               txtValorFalta.Text = CDbl(txtValoraPagar.Text) - CDbl(TotPagar)
               txtValorFalta.Text = Format(txtValorFalta.Text, "##,###0.00")
               chbValorFalta.Caption = Format(txtValorFalta.Text, "##,###0.00")
            End If
            
            If txtValorFalta.Text < 0 Then
               txtValorFalta.Text = "0,00"
               chbValorFalta.Caption = Format(txtValorFalta.Text, "##,###0.00")
            End If
'-- aonde esta o troco
            txtValorTroco.Text = Val(ConverteVirgula(Format(TotPagar, "0.00"))) - Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00")))
            txtValorTroco.Text = Format(txtValorTroco.Text, "##,###0.00")
            If txtValorTroco.Text > 0 Then
               lblTroco.Visible = True
               chbTroco.Visible = True
               chbteste.Visible = True
               chbteste.Caption = txtValorTroco.Text
            End If
            If txtValorTroco.Text > ValDinheiro Then
               MsgBox "O valor do troco não poderá ser maior que o valor pago em dinheiro"
               chbTroco.Visible = False
               ZeraVariaveis
               txtEstorno.SetFocus
               Exit Sub
            End If
            
            If txtValorTroco.Text > 0 Then
               txtNaotemtroco.Visible = False
            End If
            
            If Faturada = True Then
               If EntFaturada <> "0.00" Then
                  If Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) > Format(TotPagar, "0.00") Then
                     txtEstorno.Text = ""
                     lblModalidade.Caption = "Modalidade"
                    
                  Else
                     lblFatFin.Visible = True
                     lblFatFin.Text = "FATURADA"
                     lblValorFatFin.Visible = True
                     lblValorFatFin.Text = Format(ValorFaturada, "0.00")
                     txtEstorno.Text = ""
                     lblModalidade.Caption = "Modalidade"
                  End If
               End If
            ElseIf EntFinanciada <> "0.00" Then
               If Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) > Format(TotPagar, "0.00") Then
                  txtEstorno.Text = ""
                  lblModalidade.Caption = "Modalidade"
                
               Else
                 
                  lblFatFin.Visible = True
                  lblFatFin.Text = "FINANCIADA"
                  lblValorFatFin.Visible = True
                  lblValorFatFin.Text = Format(ValorFinanciada, "0.00")
                  txtEstorno.Text = ""
                  lblModalidade.Caption = "Modalidade"
               End If
            Else
               If Val(ConverteVirgula(Format(txtValoraPagar.Text, "0.00"))) > Format(TotPagar, "0.00") Then
                  txtEstorno.Text = ""
                  lblModalidade.Caption = "Modalidade"
               Else
                  txtEstorno.Text = ""
                  lblModalidade.Caption = "Modalidade"
                 
               End If
            End If
          End If
      End If
       txtValorPago.Text = Format(txtValorPago.Text, "0.00")
       chbValorPago.Caption = Format(txtValorPago.Text, "##,###0.00")
  

End Sub

'----------------

Private Sub ZeraVariaveis()
ValorPagamentoCartao = 0
ValDinheiro = 0
ValCheque = 0
ValCartaoAmex = 0
ValCartaoDiners = 0
ValCartaoMastercard = 0
ValCartaoVisa = 0
ValTEFVisaElectron = 0
ValTEFRedeShop = 0
ValTEFHiperCard = 0
TotPagar = 0


ValNotaCredito = 0
txtValorPago.Text = 0
chbValorPago.Caption = Format(txtValorPago.Text, "##,###0.00")
txtValorFalta.Text = ""
chbValorFalta.Caption = Format(txtValorFalta.Text, "##,###0.00")
txtValorModalidade.Text = ""
End Sub


Private Sub chameleonButton1_Click()

End Sub

Private Sub chbAmexEst_Click()
If wControle = "De" Then
   wControle = ""
Else
   wControle = "De"
End If

If wControle = "De" Then
   lblDe = "Amex"
Else
   lblPara = "Amex"
End If
End Sub

Private Sub chbBNDESESt_Click()
If wControle = "De" Then
   wControle = ""
Else
   wControle = "De"
End If

If wControle = "De" Then
   lblDe = "BNDES"
Else
   lblPara = "BNDES"
End If
End Sub

Private Sub chbChequeEst_Click()
If wControle = "De" Then
   wControle = ""
Else
   wControle = "De"
End If

If wControle = "De" Then
   lblDe = "CHEQUE"
Else
   lblPara = "CHEQUE"
End If
End Sub
'End Sub

Private Sub chbDinheiroEst_Click()

'wEstornoModalidade = "Dinheiro"

If wControle = "De" Then
   wControle = ""
Else
   wControle = "De"
End If

If wControle = "De" Then
   lblDe = "Dinheiro"
Else
   lblPara = "Dinheiro"
End If
End Sub

Private Sub chbEstornaFormaPagto_Click()
 

'Permite estornar valores de uma forma de pagamento e inserir em outra.
'Parâmetros:
'FormaOrigem: STRING com a forma de pagamento de onde o valor será estornado, com até 16 caracteres.
'FormaDestino: STRING com a forma de pagamento onde o valor será inserido, com até 16 caracteres.
'Valor: STRING com o valor a ser estornado com até 14 dígitos. Não pode ser maior que o total da forma de pagamento de origem.
'Possíveis retornos da Função (INTEIRO):
'0: Erro de comunicação.
'1: OK.
'-2: Parâmetro inválido na função.
'-4: O arquivo de inicialização BemaFI32.ini não foi encontrado no diretório de sistema do Windows.
'-5: Erro ao abrir a porta de comunicação.
'-27: Status da impressora diferente de 6,0,0 (ACK, ST1 e ST2).
'-30: Função não compatível com a impressora YANCO.

frmEstorno.Visible = True

'iRetorno = Bematech_FI_EstornoFormasPagamento("Ticket", "Dinheiro", "50,00")


End Sub

Private Sub chbEstorno_Click()

wControlaEstorno = False
 
If lblDe = "Dinheiro" Then
   If txtEstorno.Text > ValDinheiro Then
      wControlaEstorno = True
   End If
End If

If lblDe = "CHEQUE" Then
   If txtEstorno.Text > ValCheque Then
      wControlaEstorno = True
   End If
End If

If lblDe = "VisaElectronTEF" Then
   If txtEstorno.Text > ValCheque Then
      wControlaEstorno = True
   End If
End If

If lblDe = "RedeShopTEF" Then
   If txtEstorno.Text > ValTEFVisaElectron Then
      wControlaEstorno = True
   End If
End If


If wControlaEstorno = True Then
   MsgBox "Valor a estornar não pode ser maior que o valor da Modalidade pago", vbCritical, "Atenção"
   Exit Sub
End If

End Sub

Private Sub chbHiperCard_Click()
  frmcond.Visible = False
  'fraCartao.Visible = False
  FraParcelas.Visible = False
  lblParc.Visible = False
  'lblModalidade.Caption = "TEFHiperCard"
  lblModalidade.Caption = "TEF"
  txtValorModalidade.Enabled = True
  txtValorModalidade.SetFocus
  CodigoModalidade = "0403"
  wTEFHiperCard = "0403"
End Sub

Private Sub chbHiperCardEst_Click()
If wControle = "De" Then
   wControle = ""
Else
   wControle = "De"
End If

If wControle = "De" Then
   lblDe = "HiperCard"
Else
   lblPara = "HiperCard"
End If
End Sub

Private Sub chbMasterCardEst_Click()
If wControle = "De" Then
   wControle = ""
Else
   wControle = "De"
End If

If wControle = "De" Then
   lblDe = "MasterCard"
Else
   lblPara = "MasterCard"
End If
End Sub

Private Sub chbNotaCreditoEst_Click()
If wControle = "De" Then
   wControle = ""
Else
   wControle = "De"
End If

If wControle = "De" Then
   lblDe = "NotaCredito"
Else
   lblPara = "NotaCredito"
End If
End Sub

Private Sub chbRedeShop_Click()
  frmcond.Visible = False
  'fraCartao.Visible = False
  FraParcelas.Visible = False
  lblParc.Visible = False
  lblModalidade.Caption = "TEF"
  'lblModalidade.Caption = "TEFRedeShop"
  txtValorModalidade.Enabled = True
  txtValorModalidade.SetFocus
  CodigoModalidade = "0402"
  wTEFRedeShop = "0402"
  
End Sub

Private Sub chbRedeShopEst_Click()
If wControle = "De" Then
   wControle = ""
Else
   wControle = "De"
End If

If wControle = "De" Then
   lblDe = "RedeShopTEF"
Else
   lblPara = "RedeShopTEF"
End If
End Sub

Private Sub chbSair_Click()

'If txtTipoNota.Text = "CUPOM" Then
'      CancelaCupomFiscal
'End If

'       rdoCNLoja.BeginTrans
'       Screen.MousePointer = vbHourglass
'       SQL = "delete FormaPagamento where FOP_numeroPedido = " & Pedido & "  "
'       rdoCNLoja.Execute SQL
'       Screen.MousePointer = vbNormal
'       rdoCNLoja.CommitTrans

Unload Me
End Sub


Private Sub chbAmex_Click()
  lblModalidade.Caption = "AMEX"
  FraParcelas.Visible = True
  lblParc.Visible = True
  txtValorModalidade.Enabled = True
  txtValorModalidade.SetFocus
  CodigoModalidade = "0303"
  WCodigoModalidadeAMEX = "0303"
  'fraFormaPagamento.Visible = False
End Sub

Private Sub chbAmex_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
     'fraFormaPagamento.Visible = True
    lblModalidade.Caption = ""
    'fraCartao.Visible = False
    FraParcelas.Visible = False
    lblParc.Visible = False
End If
End Sub

Private Sub chbCartao_Click()
  frmcond.Visible = False
  lblModalidade.Caption = "CARTÃO"
  txtValorModalidade.Enabled = False
  'fraFormaPagamento.Visible = False
  'fraCartao.Visible = True
  
  FraParcelas.Visible = True
  lblParc.Visible = True
  txtParcelas.SelStart = 0
  txtParcelas.SelLength = Len(txtParcelas.Text)
  
End Sub

Private Sub chbCheque_Click()
  frmcond.Visible = False
  'fraCartao.Visible = False
  FraParcelas.Visible = False
  lblParc.Visible = False
  lblModalidade.Caption = "CHEQUE"
  txtValorModalidade.Enabled = True
  txtValorModalidade.SetFocus
  CodigoModalidade = "0201"
  WCodigoModalidadeCHEQUE = "0201"
  'fraFormaPagamento.Visible = False
End Sub
Private Sub chbDinheiro_Click()
  frmcond.Visible = False
  FraParcelas.Visible = False
  lblParc.Visible = False
  lblModalidade.Caption = "DINHEIRO"
  txtValorModalidade.Enabled = True
  txtValorModalidade.SetFocus

  CodigoModalidade = "0101"
  wCodigoModalidadeDINHEIRO = "0101"
 ' fraFormaPagamento.Visible = False
End Sub

Private Sub chbDinheiro_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
    If Trim(txtSerie.Text) = "CF" Then
       If txtValorPago.Text > 0 Then
          Exit Sub
       Else
          Unload Me
          Exit Sub
       End If
    Else
       Unload Me
       Exit Sub
    End If
 End If
 
End Sub

Private Sub chbDinners_Click()
  lblModalidade.Caption = "DINNERS"
  FraParcelas.Visible = True
  lblParc.Visible = True
  txtValorModalidade.Enabled = True
  txtValorModalidade.SetFocus
  CodigoModalidade = "0304"
  wCodigoModalidadeDINNERS = "0304"
  'fraFormaPagamento.Visible = False
End Sub

Private Sub chbDinners_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
     'fraFormaPagamento.Visible = True
    lblModalidade.Caption = ""
    'fraCartao.Visible = False
    FraParcelas.Visible = False
    lblParc.Visible = False
End If
End Sub

Private Sub chbMasterCard_Click()
  lblModalidade.Caption = "MASTERCARD"
  FraParcelas.Visible = True
  lblParc.Visible = True
  txtValorModalidade.Enabled = True
  txtValorModalidade.SetFocus
  CodigoModalidade = "0302"
  wCodigoModalidadeMASTERCARD = "0302"
   'fraFormaPagamento.Visible = False
End Sub


Private Sub chbMasterCard_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
     'fraFormaPagamento.Visible = True
    lblModalidade.Caption = ""
    'fraCartao.Visible = False
    FraParcelas.Visible = False
    lblParc.Visible = False
End If
End Sub

Private Sub chbNotaCredito_Click()
  frmcond.Visible = False
  'fraCartao.Visible = False
  FraParcelas.Visible = False
  lblParc.Visible = False
  lblModalidade.Caption = "NOTA CRÉDITO"
  txtValorModalidade.Enabled = True
  txtValorModalidade.SetFocus
  CodigoModalidade = "0501"
  wCodigoModalidadeNOTACREDITO = "0501"
  'fraFormaPagamento.Visible = False
End Sub

Private Sub chbOK_Click()
    
GetAsyncKeyState (vbKeyTab)

frmcond.Visible = False
wRomaneio = False

If txtTipoNota.Text = "CUPOM" Then
   Retorno = Bematech_FI_FechaCupomResumido(" ", "                                                DE MEO 115 ANOS - TRADIÇAO E QUALIDADE")
   Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
   
   NroNotaFiscal = txtNroNotaFiscal.Text
   
   EncerraVenda Val(txtPedido.Text), " ", 1
ElseIf txtTipoNota.Text = "NF" Then

      
       NroNotaFiscal = ExtraiSeqNotaControle
       rdoCNLoja.BeginTrans
       Screen.MousePointer = vbHourglass
                  
       SQL = "Update Nfcapa set Nf = " & NroNotaFiscal & ", Serie = '" & PegaSerieNota _
             & "' where NumeroPed =  " & txtPedido.Text
             
       rdoCNLoja.Execute SQL
       Screen.MousePointer = vbNormal
       rdoCNLoja.CommitTrans
    
       rdoCNLoja.BeginTrans
       Screen.MousePointer = vbHourglass
       
       SQL = "Update NfItens set Nf = " & NroNotaFiscal & ", Serie = '" & PegaSerieNota _
            & "' where NumeroPed =  " & txtPedido.Text
       rdoCNLoja.Execute SQL
       Screen.MousePointer = vbNormal
       rdoCNLoja.CommitTrans
       
       If EncerraVenda(Val(txtPedido.Text), " ", 1) = False Then
          Exit Sub
       End If
        
       EmiteNotafiscal NroNotaFiscal, txtSerie.Text
      
       frmCaixaNF.txtPedido.Text = ""
       frmCaixaNF.grdItens.Rows = 1
       frmCaixaNF.lblTotalvenda.Caption = ""
       frmCaixaNF.lblTotalitens.Caption = ""
       LimpaGrid frmCaixaNF.grdItens
    
    
ElseIf txtTipoNota.Text = "Romaneio" Then
      Call PegaNumeroRomaneio
      Call ImprimeRomaneio
      wRomaneio = True
      EncerraVenda Val(txtPedido.Text), " ", 1
      frmCaixaNF.txtPedido.Text = ""
      frmCaixaNF.grdItens.Rows = 1
      frmCaixaNF.lblTotalvenda.Caption = ""
      frmCaixaNF.lblTotalitens.Caption = ""
      LimpaGrid frmCaixaNF.grdItens
   '   txtSerie = "00"
End If


FechaVenda

GravaMovimentoCaixa

EntFaturada = 0
EntFinanciada = 0
ValDinheiro = 0
ValCheque = 0
ValCartaoVisa = 0
AvistaReceber = 0
ValCartaoMastercard = 0
ValCartaoAmex = 0
ValCartaoDiners = 0
ValorPagamentoCartao = 0
ValTEFVisaElectron = 0
ValTEFRedeShop = 0
ValTEFHiperCard = 0
ValNotaCredito = 0
TotPagar = 0

txtNaotemtroco.Visible = True
lblTootip.Visible = False
lblTootip1.Visible = False
  
  Pedido = txtPedido.Text
  Pedido = IIf(txtPedido.Text = "", 0, txtPedido.Text)
  frmStartaProcessos.txtPedido.Text = txtPedido.Text
  txtValoraPagar.Text = ""
  txtValorFalta.Text = ""
  
  txtValorModalidade.Text = ""
  txtValorTroco.Text = ""
 
  If txtIdentificadequeTelaqueveio.Text = "frmCaixaTEF" Then
     frmCaixaTEF.txtCodigoProduto = ""
     frmCaixaTEF.txtCGC_CPF.Text = ""
     LimpaGrid frmCaixaTEF.grdItens
     frmCaixaTEF.grdItens.Rows = 1
     wItens = 0
     txtIdentificadequeTelaqueveio.Text = ""
     frmCaixaTEF.lblTotalvenda.Caption = ""
     frmCaixaTEF.lblTotalitens.Caption = ""
     frmCaixaTEF.lblDescricaoProduto.Caption = ""
     frmCaixaTEF.fraProduto.Visible = False
     frmCaixaTEF.fraNFP.Visible = True
  End If
  
  If txtIdentificadequeTelaqueveio.Text = "frmCaixaTEFPedido" Then
     EncerraVenda Val(txtPedido.Text), " ", 1
     
     LimpaGrid frmCaixaTEFPedido.grdItens
     frmCaixaTEFPedido.grdItens.Rows = 1
     txtIdentificadequeTelaqueveio.Text = ""
     frmCaixaTEFPedido.lblTotalvenda.Caption = ""
     frmCaixaTEFPedido.lblTotalitens.Caption = ""
     frmCaixaTEFPedido.fraNFP.Visible = False
     frmCaixaTEFPedido.txtPedido.Text = ""
     frmCaixaTEFPedido.txtCGC_CPF.Text = ""
     frmCaixaTEFPedido.fraPedido.Visible = True
     
  End If
  
  If frmFormaPagamento.txtIdentificadequeTelaqueveio.Text = "frmCaixaNF" Then
     frmCaixaNF.lblTotalvenda.Caption = ""
     frmCaixaNF.lblTotalitens.Caption = ""
     frmCaixaNF.txtPedido.Text = ""
  End If
  If frmFormaPagamento.txtIdentificadequeTelaqueveio.Text = "frmCaixaRomaneio" Then
     frmCaixaRomaneio.lblTotalvenda.Caption = ""
     frmCaixaRomaneio.lblTotalitens.Caption = ""
     frmCaixaRomaneio.txtPedido.Text = ""
  End If
   
  fraRecebimento.Visible = False
  lblTotalPedido.Visible = False
  lblValorTotalPedido.Visible = False
  lblTootip.Text = ""
  lblTootip1.Text = ""
  chbOK.Enabled = False
  
  Unload frmCaixaTEF
  Unload frmCaixaTEFPedido
  Unload frmCaixaNF
  Unload frmCaixaRomaneio
  Unload Me
  
    
  frmStartaProcessos.Show
  frmStartaProcessos.ZOrder
  
  End Sub
  
Private Sub GravaMovimentoCaixa()

        
         'EntFinanciada = 0
         
      
             
          SQL = "Select * from controlecaixa  where " _
             & " Ctr_DataInicial between '" & Format(Date, "yyyy/mm/dd") & " 00:00:00' and  '" _
             & Format(Date, "yyyy/mm/dd") & " 23:59:59'"
          
             PegaLoja.CursorLocation = adUseClient
             PegaLoja.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
          
        Wecf = GLB_ECF
         
         If PegaLoja.EOF = False Then
                  
               If txtTipoNota.Text = "Romaneio" Then
                 rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                              & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                              & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 20105 & "," & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                              & "" & ConverteVirgula(Format(txtValoraPagar.Text, "##,###0.00")) & ", " _
                              & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
              End If
              
              If Faturada = True Then
                 'ValorFaturada
                 rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                              & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                              & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10501 & "," & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                              & "" & ConverteVirgula(Format(wTotalNotaFatFin, "##,###0.00")) & ", " _
                              & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                             
                 
              End If
        
              
              If Financiada = True Then
                 'ValorFinanciada
                 rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                              & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                              & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10601 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                              & "" & ConverteVirgula(Format(wTotalNotaFatFin, "##,###0.00")) & ", " _
                              & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
               
              End If
        
              
              If AvistaReceber <> 0 Then
                 
                 rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                              & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio,MC_ControleAVR, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                              & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10204 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                              & "" & ConverteVirgula(Format(AvistaReceber, "##,###0.00")) & ", " _
                              & "0,0,0,0," & WParcelas & ", " & 9 & ",'A','A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                       
              End If

              If WCodigoModalidadeVISA = "0301" Then
                If ValCartaoVisa > 0 Then
                 'VISA
                 If EntFaturada = "0.00" And EntFinanciada = "0.00" Then
                    rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                 & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                 & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10301 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                 & "" & ConverteVirgula(Format(ValCartaoVisa, "##,###0.00")) & ", " _
                                 & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                   
                 Else
                    rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                 & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                 & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10301 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                 & "" & ConverteVirgula(Format(ValCartaoVisa, "##,###0.00")) & ", " _
                                 & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    
                    If EntFaturada <> "0.00" Then
                       rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                              & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                              & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 11004 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                              & "" & ConverteVirgula(Format(ValCartaoVisa, "##,###0.00")) & ", " _
                              & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    ElseIf EntFinanciada <> "0.00" Then
                        rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                     & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                     & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 11005 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                     & "" & ConverteVirgula(Format(ValCartaoVisa, "##,###0.00")) & ", " _
                                     & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    End If
                 End If
               End If
              End If
        
              If wCodigoModalidadeMASTERCARD = "0302" Then
                If ValCartaoMastercard > 0 Then
                 'MASTERCARD
                 If EntFaturada = "0.00" And EntFinanciada = "0.00" Then
                    rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                 & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                 & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10302 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                 & "" & ConverteVirgula(Format(ValCartaoMastercard, "##,###0.00")) & ", " _
                                 & "0,0,0,0," & WParcelas & "," & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                   
                 Else
                    rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                 & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                 & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10302 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                 & "" & ConverteVirgula(Format(ValCartaoMastercard, "##,###0.00")) & ", " _
                                 & "0,0,0,0," & WParcelas & "," & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    
                    If EntFaturada <> "0.00" Then
                       SQL = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                              & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                              & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 11004 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                              & "" & ConverteVirgula(Format(ValCartaoMastercard, "##,###0.00")) & ", " _
                              & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                              rdoCNLoja.Execute (SQL)
                              
                    ElseIf EntFinanciada <> "0.00" Then
                        rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                     & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                     & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 11005 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                     & "" & ConverteVirgula(Format(ValCartaoMastercard, "##,###0.00")) & ", " _
                                     & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    End If
                 End If
                End If
              End If
              
              
              If WCodigoModalidadeAMEX = "0303" Then
                If ValCartaoAmex > 0 Then
                 'AMEX
                 If EntFaturada = "0.00" And EntFinanciada = "0.00" Then
                    rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                 & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                 & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10303 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                 & "" & ConverteVirgula(Format(ValCartaoAmex, "##,###0.00")) & ", " _
                                 & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                   
                 Else
                    rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                 & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                 & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10303 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                 & "" & ConverteVirgula(Format(ValCartaoAmex, "##,###0.00")) & ", " _
                                 & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    
                    If EntFaturada <> "0.00" Then
                       rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                              & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                              & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 11004 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                              & "" & ConverteVirgula(Format(ValCartaoAmex, "##,###0.00")) & ", " _
                              & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & ")"
                    ElseIf EntFinanciada <> "0.00" Then
                        rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                     & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                     & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 11005 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                     & "" & ConverteVirgula(Format(ValCartaoAmex, "##,###0.00")) & ", " _
                                     & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    End If
                 End If
                End If
              End If
        
              
              If wCodigoModalidadeDINNERS = "0304" Then
                If ValCartaoDiners > 0 Then
                 'DINNERS
                 If EntFaturada = "0.00" And EntFinanciada = "0.00" Then
                    rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                 & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                 & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10304 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                 & "" & ConverteVirgula(Format(ValCartaoDiners, "##,###0.00")) & ", " _
                                 & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    
                 Else
                 
                    rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                 & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                 & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10304 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                 & "" & ConverteVirgula(Format(ValCartaoDiners, "##,###0.00")) & ", " _
                                 & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                 
                    If EntFaturada <> "0.00" Then
                       rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                              & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                              & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 11004 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                              & "" & ConverteVirgula(Format(ValCartaoDiners, "##,###0.00")) & ", " _
                              & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    ElseIf EntFinanciada <> "0.00" Then
                        rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                     & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                     & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 11005 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                     & "" & ConverteVirgula(Format(ValCartaoDiners, "##,###0.00")) & ", " _
                                     & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    End If
                 End If
                End If
              End If
        
              
              If wTEFVisaElectron = "0401" Then
                If ValTEFVisaElectron > 0 Then
                 'TEF
                 If EntFaturada = "0.00" And EntFinanciada = "0.00" Then
                    rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                 & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                 & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10203 & ",'chbVisaElectron', " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                 & "" & ConverteVirgula(Format(ValTEFVisaElectron, "##,###0.00")) & ", " _
                                 & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    
                 Else
                 
                    rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                 & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                 & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10203 & ",'chbVisaElectron', " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                 & "" & ConverteVirgula(Format(ValTEFVisaElectron, "##,###0.00")) & ", " _
                                 & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                 
                    If EntFaturada <> "0.00" Then
                       rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                              & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                              & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 11004 & ",'chbVisaElectron', " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                              & "" & ConverteVirgula(Format(ValTEFVisaElectron, "##,###0.00")) & ", " _
                              & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    ElseIf EntFinanciada <> "0.00" Then
                        rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                     & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                     & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 11005 & ",'chbVisaElectron', " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                     & "" & ConverteVirgula(Format(ValTEFVisaElectron, "##,###0.00")) & ", " _
                                     & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    End If
                 End If
                End If
              End If
              
              If wTEFRedeShop = "0402" Then
                If ValTEFRedeShop > 0 Then
                 'TEF
                 If EntFaturada = "0.00" And EntFinanciada = "0.00" Then
                    rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                 & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                 & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10203 & ",'chbRedeShop', " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                 & "" & ConverteVirgula(Format(ValTEFRedeShop, "##,###0.00")) & ", " _
                                 & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    
                 Else
                 
                    rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                 & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                 & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10203 & ",'chbRedeShop', " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                 & "" & ConverteVirgula(Format(ValTEFRedeShop, "##,###0.00")) & ", " _
                                 & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                 
                    If EntFaturada <> "0.00" Then
                       rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                              & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                              & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 11004 & ",'chbRedeShop', " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                              & "" & ConverteVirgula(Format(ValTEFRedeShop, "##,###0.00")) & ", " _
                              & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    ElseIf EntFinanciada <> "0.00" Then
                        rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                     & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                     & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 11005 & ",'chbRedeShop', " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                     & "" & ConverteVirgula(Format(ValTEFRedeShop, "##,###0.00")) & ", " _
                                     & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    End If
                 End If
                End If
              End If
              
              If wTEFHiperCard = "0403" Then
                If ValTEFHiperCard > 0 Then
                 'TEF
                 If EntFaturada = "0.00" And EntFinanciada = "0.00" Then
                    rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                 & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                 & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10203 & ",'chbHiperCard', " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                 & "" & ConverteVirgula(Format(ValTEFHiperCard, "##,###0.00")) & ", " _
                                 & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    
                 Else
                 
                    rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                 & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                 & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10203 & ",'chbHiperCard', " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                 & "" & ConverteVirgula(Format(ValTEFHiperCard, "##,###0.00")) & ", " _
                                 & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                 
                    If EntFaturada <> "0.00" Then
                       rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                              & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                              & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 11004 & ",'chbHiperCard', " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                              & "" & ConverteVirgula(Format(ValTEFHiperCard, "##,###0.00")) & ", " _
                              & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    ElseIf EntFinanciada <> "0.00" Then
                        rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                     & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                     & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 11005 & ",'chbHiperCard', " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                     & "" & ConverteVirgula(Format(ValTEFHiperCard, "##,###0.00")) & ", " _
                                     & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    End If
                 End If
                End If
              End If
              
        
              If wCodigoModalidadeNOTACREDITO = "0501" Then
                If ValNotaCredito > 0 Then
                 'NOTA DE CREDITO
                 If EntFaturada = "0.00" And EntFinanciada = "0.00" Then
                    rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                 & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                 & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10701 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                 & "" & ConverteVirgula(Format(ValNotaCredito, "##,###0.00")) & ", " _
                                 & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    
                 Else
                 
                    rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                 & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                 & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10701 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                 & "" & ConverteVirgula(Format(ValNotaCredito, "##,###0.00")) & ", " _
                                 & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                 
                    If EntFaturada <> "0.00" Then
                       rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                              & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                              & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 11004 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                              & "" & ConverteVirgula(Format(ValNotaCredito, "##,###0.00")) & ", " _
                              & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    ElseIf EntFinanciada <> "0.00" Then
                        rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                     & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                     & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 11005 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                     & "" & ConverteVirgula(Format(ValNotaCredito, "##,###0.00")) & ", " _
                                     & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    End If
                 End If
                End If
              End If
              
              If WCodigoModalidadeCHEQUE = "0201" Then
                If ValCheque > 0 Then
                 'CHEQUE
                 If EntFaturada = "0.00" And EntFinanciada = "0.00" Then
                    rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                 & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                 & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10201 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                 & "" & ConverteVirgula(Format(ValCheque, "##,###0.00")) & ", " _
                                 & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                  
                    
                 Else
                 
                    rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                 & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                 & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 10201 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                 & "" & ConverteVirgula(Format(ValCheque, "##,###0.00")) & ", " _
                                 & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    
                    If EntFaturada <> "0.00" Then
                       rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                              & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                              & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 11004 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                              & "" & ConverteVirgula(Format(ValCheque, "##,###0.00")) & ", " _
                              & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    ElseIf EntFinanciada <> "0.00" Then
                        rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                     & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                     & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 11005 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                     & "" & ConverteVirgula(Format(ValCheque, "##,###0.00")) & ", " _
                                     & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                    End If
                 End If
                End If
              End If
              'ValDinheiro = ValDinheiro - txtValorTroco.Text
              If wCodigoModalidadeDINHEIRO = "0101" Then
                If ValDinheiro > 0 Then
                 'DINHEIRO
                 If EntFaturada = "0.00" And EntFinanciada = "0.00" Then
                    SQL = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja, MC_Data, MC_Grupo, MC_Documento,MC_Serie, MC_Valor, MC_banco, MC_Agencia," _
                                 & "MC_Contacorrente, MC_bomPara, MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                 & "'" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', 10101, " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                 & " " & ConverteVirgula(Format(ValDinheiro, "##,###0.00") - Format(txtValorTroco.Text, "##,###0.00")) & ", " _
                                 & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                        rdoCNLoja.Execute (SQL)
                   
                    
                 Else
                 
                     SQL = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja, MC_Data, MC_Grupo, MC_Documento,MC_Serie, MC_Valor, MC_banco, MC_Agencia," _
                                 & "MC_Contacorrente, MC_bomPara, MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                 & "'" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', 10101, " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                 & " " & ConverteVirgula(Format(ValDinheiro, "##,###0.00") - Format(txtValorTroco.Text, "##,###0.00")) & ", " _
                                 & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                        rdoCNLoja.Execute (SQL)
                   
                 
                    If EntFaturada <> "0.00" Then
                       rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                              & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                              & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 11004 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                              & " " & ConverteVirgula(Format(ValDinheiro, "##,###0.00") - Format(txtValorTroco.Text, "##,###0.00")) & ", " _
                              & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                       
                              
                    ElseIf EntFinanciada <> "0.00" Then
                        rdoCNLoja.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                                     & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                                     & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & 11005 & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                                     & " " & ConverteVirgula(Format(ValDinheiro, "##,###0.00") - Format(txtValorTroco.Text, "##,###0.00")) & ", " _
                                     & "0,0,0,0," & WParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                      
                    End If
                 End If
  
               
               End If
             End If
              
              wGrupo = 0
              
              If txtSerie.Text = "CF" Then
                 wGrupo = 20101
              ElseIf txtSerie.Text = PegaSerieNota Then
                 wGrupo = 20102
              ElseIf txtSerie.Text = "SF" Then
                 wGrupo = 20103
              ElseIf txtSerie.Text = "SM" Then
                 wGrupo = 20104
              ElseIf txtSerie.Text = "00" Then
                 wGrupo = 20105
              ElseIf txtSerie.Text = "C0" Then
                 wGrupo = 20106
              ElseIf txtSerie.Text = "D1" Then
                 wGrupo = 20107
              ElseIf txtSerie.Text = "S1" Then
                 wGrupo = 20108
              
              End If
             
              If wGrupo <> 0 Then
                 If txtTipoNota.Text = "CUPOM" Then
                    SQL = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                              & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                              & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & wGrupo & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                              & "" & ConverteVirgula(Format(txtValoraPagar.Text, "##,###0.00")) & ", " _
                              & "0,0,0,0,0,0,'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                              rdoCNLoja.Execute (SQL)
                              
                 ElseIf txtSerie.Text <> "00" Then
                        SQL = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                              & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa) values(" & Wecf & ",'" & PegaLoja("ctr_operador") & "','" & Trim(wlblloja) & "', " _
                              & " '" & Format(PegaLoja("ctr_datainicial"), "mm/dd/yyyy") & "', " & wGrupo & ", " & NroNotaFiscal & ",'" & txtSerie.Text & "', " _
                              & "" & ConverteVirgula(Format(wTotalNota, "##,###0.00")) & ", " _
                              & "0,0,0,0,0,9,'A'," & GLB_CTR_Protocolo & "," & wNumeroCaixa & ")"
                              rdoCNLoja.Execute (SQL)
                 End If
              End If
    
    End If
    
    PegaLoja.Close
    
 wCodigoModalidadeDINHEIRO = ""
 WCodigoModalidadeAMEX = ""
 WCodigoModalidadeCHEQUE = ""
 wCodigoModalidadeDINNERS = ""
 wCodigoModalidadeMASTERCARD = ""
 wCodigoModalidadeNOTACREDITO = ""
 wTEFVisaElectron = ""
 wTEFRedeShop = ""
 wTEFHiperCard = ""
 WCodigoModalidadeVISA = ""
 
End Sub
Private Function ProcuraPedido()
   
   Screen.MousePointer = 11
   Dim vSQL As String
   Dim Linha As Long
   Dim i As Integer
   Dim wTootip As Double
   Dim Tootip1 As Double
   
  
        
        ConsistePedido Val(txtPedido)
        
        SQL = "SELECT DISTINCT nfcapa.*,Produtoloja.*,Vende.*,Nfitens.* " _
            & "from Vende,Nfcapa,Nfitens,Produtoloja " _
            & "Where (Vende.ve_codigo = nfcapa.vendedor) " _
            & "and (Produtoloja.pr_referencia = Nfitens.referencia) " _
            & "and (nfcapa.numeroped=Nfitens.numeroped) " _
            & "and nfcapa.numeroped=" & txtPedido.Text & " " _
            & "and nfcapa.tiponota='PA' order by Nfitens.item"
        
        RsDados.CursorLocation = adUseClient
        RsDados.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
        If Not RsDados.EOF Then
         
           txtPedido.Text = Trim(RsDados("NumeroPED"))
          
           lblTotalPedido.Visible = True
           lblValorTotalPedido.Visible = True
           lblValorTotalPedido.Text = Format(CDbl(RsDados("TotalNota")), "##,##0.00")
         
           WParcelas = RsDados("parcelas")
           txtParcelas.Text = RsDados("parcelas")
           wIndicePreco = RsDados("Indicepreco")
           
           If Trim(RsDados("ModalidadeVenda")) = "Financiado" Then
              lblTootip.Text = " ATENÇÃO: Valor do contrato R$   " & Format(((lblValorTotalPedido.Text - RsDados("pgentra")) * wIndicePreco), "##,###,##0.00")
              lblTootip1.Text = WParcelas & "  Parcela(s)  de  R$   " & Format(((lblValorTotalPedido.Text - RsDados("pgentra")) * wIndicePreco) / WParcelas, "##,###,##0.00")
           Else
              lblTootip.Text = ""
              lblTootip1.Text = ""
           End If
             
         
           
           If RsDados("condpag") > 3 Then
              Faturada = True
              Financiada = False
              wVerificaAVR = False
              EntFaturada = Format(RsDados("pgentra"), "0.00")
              ValorFaturada = Format(RsDados("VlrMercadoria"), "0.00")
           ElseIf RsDados("condpag") = 3 Then
              Financiada = True
              Faturada = False
              wVerificaAVR = False
              EntFinanciada = Format(RsDados("pgentra"), "0.00")
              ValorFinanciada = Format(RsDados("VlrMercadoria"), "0.00") - Format(IIf(IsNull(RsDados("pgentra")), 0, RsDados("pgentra")), "0.00")
           ElseIf RsDados("condpag") = 1 Then
              Faturada = False
              Financiada = False
              wVerificaAVR = False
           ElseIf RsDados("condpag") = 2 Then
                  wVerificaAVR = True
                  Faturada = False
                  Financiada = False
                  AvistaReceber = Format(RsDados("totalnota"), "0.00")
           End If
           wTotalNotaFatFin = 0
           wTotalNota = 0
           '---- Quando Faturada ou fincanciada, gravar para codigo 10501 e 10601, somente o valor já descontando a entrada Adilson 08/02/2010
           If Faturada = True Or Financiada = True Then
              wTotalNotaFatFin = Format(CDbl(RsDados("TotalNota")), "##,##0.00") - Format(CDbl(RsDados("pgentra")), "##,##0.00")
              wTotalNota = Format(CDbl(RsDados("TotalNota")), "##,##0.00")
           Else
              wTotalNota = Format(CDbl(RsDados("TotalNota")), "##,##0.00")
           End If
        '---------------------------------------------------------------------------------------------------------------
           
           If RsDados("cgccli") <> "" Then
              wDocumento = Trim(RsDados("cgccli"))
              wPessoa = RsDados("Pessoacli")
           End If
        
           If wVerificaAVR = True Then
              lblApagar.Text = Format(CDbl(RsDados("TotalNota")), "##,###0.00")
           End If
           txtPedido.Enabled = False
           RsDados.Close
        Else
           MsgBox "Número de pedido inexistente.", vbInformation, "Informação"
        
        Unload Me
        End If
        
  
   Screen.MousePointer = 0
   
End Function


Private Sub chbOK_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Exit Sub
End If
 If KeyAscii = 27 Then
    Unload Me
 End If
End Sub

Private Sub chbOutros_Click()
frmcond.Visible = False
End Sub

'Private Sub chbTEF_Click()
'  frmcond.Visible = False
'  'fraCartao.Visible = False
'  FraParcelas.Visible = False
'  lblParc.Visible = False
'  lblModalidade.Caption = "TEF"
'  txtValorModalidade.Enabled = True
'  txtValorModalidade.SetFocus
'  CodigoModalidade = "0401"
'  wTEFVisaElectron = "0401"
'End Sub

Private Sub chbVisa_Click()
  lblModalidade.Caption = "VISA"
  FraParcelas.Visible = True
  lblParc.Visible = True
  txtValorModalidade.Enabled = True
  txtValorModalidade.SetFocus
  CodigoModalidade = "0301"
  WCodigoModalidadeVISA = "0301"
 ' fraFormaPagamento.Visible = False
End Sub


Private Sub chbVisa_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
     'fraFormaPagamento.Visible = True
    lblModalidade.Caption = ""
    'fraCartao.Visible = False
    FraParcelas.Visible = False
    lblParc.Visible = False
End If
End Sub

Private Sub chbVisaElectron_Click()
  frmcond.Visible = False
 ' fraCartao.Visible = False
  FraParcelas.Visible = False
  lblParc.Visible = False
  lblModalidade.Caption = "TEF"
 ' lblModalidade.Caption = "TEFVisaElectron"
  txtValorModalidade.Enabled = True
  txtValorModalidade.SetFocus
  CodigoModalidade = "0401"
  wTEFVisaElectron = "0401"
End Sub

Private Sub chbVisaElectronEst_Click()
If wControle = "De" Then
   wControle = ""
Else
   wControle = "De"
End If

If wControle = "De" Then
   lblDe = "VisaElectronTEF"
Else
   lblPara = "VisaElectronTEF"
End If
End Sub

Private Sub chbVisaEst_Click()
If wControle = "De" Then
   wControle = ""
Else
   wControle = "De"
End If

If wControle = "De" Then
   lblDe = "Visa"
Else
   lblPara = "Visa"
End If
End Sub

Private Sub cmdCondicao_Click()
If wMostraGrideCondicao = False Then
   frmcond.Visible = True
   wMostraGrideCondicao = True
Else
   frmcond.Visible = False
   wMostraGrideCondicao = False
End If
End Sub

Private Sub cmdteste_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command12_Click()

End Sub

Private Sub Form_Activate()
If frmFormaPagamento.txtIdentificadequeTelaqueveio.Text = "frmCaixaTEFPedido" Then
   txtPedido.Text = Pedido
   ProcuraPedido
   VerificaTipoModalidade
     
End If
frmcond.Visible = True
chbTroco.Visible = False
'If txtTipoNota.Text = "CUPOM" Then
    frmcond.Visible = False
  ' lblTotalPedido.Visible = True
  ' lblValorTotalPedido.Visible = True
  ' lblValorTotalPedido.Text = frmFormaPagamento.txtValoraPagar.Text
  ' lblTootip.Text = ""
  ' lblTootip1.Text = ""
   txtValoraPagar.Text = frmFormaPagamento.txtValoraPagar.Text
   chbValoraPagar.Caption = txtValoraPagar.Text
'End If
End Sub

Private Sub Form_Load()
 ' chbteste.Caption = 100
  Left = 9250
  Top = 1175
  frmcond.Left = 330
  frmcond.Top = 1305
  frmcond.Height = 6570
  frmcond.Width = 5100
  fraNModalidades.Visible = True
  txtTipoNota.Text = "CUPOM"
  txtTipoNota.Text = "NF"
  txtSerie.Text = PegaSerieNota
  txtValorModalidade.Enabled = False
  txtValorTroco.Visible = False
  txtValorPago.Text = 0
  chbValorPago.Caption = Format(txtValorPago.Text, "##,###0.00")
  txtValorFalta.Text = 0
  txtValorTroco.Text = 0
 
  If txtTipoNota.Text = "NF" Or _
     txtTipoNota.Text = "Romaneio" Then
    txtValoraPagar.Text = Format(wValoraPagarNORMAL, "0.00")
    chbValoraPagar.Caption = txtValoraPagar.Text
  End If

If frmFormaPagamento.txtSerie.Text = "00" Then
     txtPedido.Text = Pedido
     ProcuraPedido
     VerificaTipoModalidade
     GoTo Continua
 End If

If frmFormaPagamento.txtSerie.Text = PegaSerieNota Then
   txtPedido.Text = Pedido
   ProcuraPedido
   VerificaTipoModalidade
     
End If
  

Continua:

txtNaotemtroco.Top = txtValorTroco.Top
txtNaotemtroco.Left = txtValorTroco.Left
txtNaotemtroco.Height = txtValorTroco.Height
txtNaotemtroco.Width = txtValorTroco.Width

txtNaotemtroco.Visible = True
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
EntFaturada = 0
EntFinanciada = 0
ValDinheiro = 0
ValCheque = 0
ValCartaoVisa = 0
ValorPagamentoCartao = 0
AvistaReceber = 0
ValCartaoMastercard = 0
ValCartaoAmex = 0
ValCartaoDiners = 0
ValTEFVisaElectron = 0
ValTEFRedeShop = 0
ValTEFHiperCard = 0

ValNotaCredito = 0
TotPagar = 0
frmcond.Visible = False

txtNaotemtroco.Visible = True
lblTootip.Visible = False
lblTootip1.Visible = False
 If txtPedido.Text <> "" Then
    Pedido = txtPedido.Text
 End If
  Pedido = IIf(txtPedido.Text = "", 0, txtPedido.Text)
  frmStartaProcessos.txtPedido.Text = txtPedido.Text
       
  txtValoraPagar.Text = ""
  txtValorFalta.Text = ""
  txtValorModalidade.Text = ""
  txtValorTroco.Text = ""
 
 ' If txtIdentificadequeTelaqueveio.Text = "frmCaixaTEF" Then
 '    frmCaixaTEF.txtCodigoProduto = ""
 '    frmCaixaTEF.txtCGC_CPF.Text = ""
 '    LimpaGrid frmCaixaTEF.grdItens
 '    frmCaixaTEF.grdItens.Rows = 1
 '    txtIdentificadequeTelaqueveio.Text = ""
 '    frmCaixaTEF.lblTotalvenda.Caption = ""
 '    frmCaixaTEF.lblTotalitens.Caption = ""
 '    frmCaixaTEF.lblDescricaoProduto.Caption = ""
 '    frmCaixaTEF.fraProduto.Visible = False
 '    frmCaixaTEF.fraNFP.Visible = True
 ' End If
  
 ' If txtIdentificadequeTelaqueveio.Text = "frmCaixaTEFPedido" Then
 '    LimpaGrid frmCaixaTEFPedido.grdItens
 '    frmCaixaTEFPedido.grdItens.Rows = 1
 '    txtIdentificadequeTelaqueveio.Text = ""
 '    frmCaixaTEFPedido.lblTotalvenda.Caption = ""
 '    frmCaixaTEFPedido.lblTotalitens.Caption = ""
 '    frmCaixaTEFPedido.fraNFP.Visible = False
 '    frmCaixaTEFPedido.txtPedido.Text = ""
 '    frmCaixaTEFPedido.txtCGC_CPF.Text = ""
 '    frmCaixaTEFPedido.fraPedido.Visible = True
 '    frmCaixaTEFPedido.grdItens.BackColorBkg = &H80000006
 '    frmCaixaTEFPedido.grdItens.ColWidth(0) = 6500
 '    frmCaixaTEFPedido.lblTotalvenda.Caption = ""
 '    frmCaixaTEFPedido.lblTotalitens.Caption = ""
        
 ' End If
  
'  If frmFormaPagamento.txtIdentificadequeTelaqueveio.Text = "frmCaixaNF" Then
'     frmCaixaNF.lblTotalvenda.Caption = ""
'     frmCaixaNF.lblTotalitens.Caption = ""
'     frmCaixaNF.txtPedido.Text = ""
'  ElseIf frmFormaPagamento.txtIdentificadequeTelaqueveio.Text = "frmCaixaRomaneio" Then
'       frmCaixaNF.lblTotalvenda.Caption = ""
'     frmCaixaNF.lblTotalitens.Caption = ""
'     frmCaixaNF.txtPedido.Text = ""
'  End If
 

  frmCaixaTEFPedido.fraPedido.Visible = True
  'frmCaixaTEFPedido.fraNFP.Visible = False
  
  
  fraRecebimento.Visible = False
  lblTotalPedido.Visible = False
  lblValorTotalPedido.Visible = False
  lblTootip.Text = ""
  lblTootip1.Text = ""
  

End Sub

Private Sub txtParcelas_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
       If IsNumeric(txtParcelas.Text) Then
            If txtParcelas > 12 Then
               MsgBox "Quantidade de parcelas invalida", vbCritical, "Atenção"
               txtParcelas.SelStart = 0
               txtParcelas.SelLength = Len(txtParcelas.Text)
            Else
                WParcelas = Val(txtParcelas.Text)
                txtValorModalidade.SetFocus
                'fraCartao.Visible = False
                FraParcelas.Visible = False
                lblParc.Visible = False
            End If
       End If
   End If

If KeyAscii = 27 Then
   FraParcelas.Visible = False
   lblParc.Visible = False
End If


End Sub

Private Sub txtValorModalidade_GotFocus()
   txtValorModalidade.Text = ""
   txtValorModalidade.SelStart = 0
   txtValorModalidade.SelLength = Len(txtValorModalidade.Text)
End Sub

Private Sub txtValorModalidade_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
   ' fraFormaPagamento.Visible = True
    lblModalidade.Caption = "Modalidade"
    
   ' fraCartao.Visible = False
    FraParcelas.Visible = False
    lblParc.Visible = False
    txtValorModalidade.Enabled = False
    txtValorModalidade.Text = ""
    chbDinheiro.SetFocus
    Exit Sub
 End If
    
 VerteclaVirgula txtValorModalidade, KeyAscii
 If KeyAscii = 13 Then
    If txtValorModalidade.Text = "" Then
       txtValorModalidade.SelStart = 0
       txtValorModalidade.SelLength = Len(txtValorModalidade.Text)
       txtValorModalidade.SetFocus
       Exit Sub
    End If
    txtPedido.Text = Pedido
    
    If chbSair.Visible = True Then
       chbSair.Visible = False
       If txtIdentificadequeTelaqueveio.Text = "frmCaixaTEF" Or _
          txtIdentificadequeTelaqueveio.Text = "frmCaixaPedido" Then
          Retorno = Bematech_FI_IniciaFechamentoCupom("D", "$", 0)
          Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
       End If
    End If
'    wValorRomaneio = Format(txtValorModalidade.Text, "##,###,##0.00")
    
    Call GuardaValoresParaGravarMovimentoCaixa
    
    
    
    'Call GravarModalidade
    
    txtValorModalidade.Text = ""
    txtValorModalidade.Enabled = False
    lblModalidade.Caption = "Modalidade"
    'fraFormaPagamento.Visible = True
    'fraCartao.Visible = False
    FraParcelas.Visible = False
    lblParc.Visible = False
    
    
    'A pergunta abaixo é feita para que se o valor do troco for maior que o valor em dinheiro
    ' ou o valor do cartao > que o valor da nota, saia da rotina sem sumir o franModalidade.
    
    If wValorModalidadeIncorreto = True Then
       Exit Sub
    End If
    
    If txtValorFalta.Text = "" Then
       txtValorFalta.Text = 0
    End If
       
         
    If txtValorFalta.Text <= 0 Then
       chbOK.Visible = True
       chbOK.Enabled = True
       chbSair.Enabled = False
       chbOK.SetFocus
       fraNModalidades.Visible = False
    End If
    
End If

End Sub

Private Sub txtValorModalidade_LostFocus()
  txtValorModalidade.Text = ""
End Sub
 Private Sub GravarModalidade()
 
 'If wEstorno = True Then
 '   txtValorModalidade.Text = txtEstorno.Text
 'End If
 
 If Not IsNumeric(txtValorModalidade) Then
    txtValorModalidade.SelStart = 0
    txtValorModalidade.SelLength = Len(txtValorModalidade.Text)
    Exit Sub
 End If
 
 If Numeros(txtValorModalidade) = "" Then
    txtValorModalidade.SelStart = 0
    txtValorModalidade.SelLength = Len(txtValorModalidade.Text)
    Exit Sub
 End If
 
 If Numeros(txtValorModalidade) <= 0 Then
    txtValorModalidade.SelStart = 0
    txtValorModalidade.SelLength = Len(txtValorModalidade.Text)
    Exit Sub
 End If
' lblModalidade.Caption = "VisaElectron"
If UCase(txtSerie.Text) = "CF" Then
   Retorno = Bematech_FI_EfetuaFormaPagamento(lblModalidade.Caption, txtValorModalidade.Text * 100)
   Call VerificaRetornoImpressora("", "", "Emissão de Cupom Fiscal")
End If
     
     
     SQL = ("Select * from FormaPagamento where FOP_Loja= '" & Trim(wlblloja) & "'" _
       & " And FOP_NumeroPedido = " & txtPedido & " And FOP_CodigoModalidade1 = '" & CodigoModalidade & " '")
     RsDados.CursorLocation = adUseClient
     RsDados.Open SQL, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
      
     If RsDados.EOF Then
        
        
           
        ' If frmFormaPagamento.txtIdentificadequeTelaqueveio = "frmCaixaNF" Then
         
        ' End If
        
         rdoCNLoja.BeginTrans
         Screen.MousePointer = vbHourglass
         
           SQL = "Insert Into FormaPagamento (FOP_Loja,FOP_Documento," _
           & "FOP_Serie," _
           & "FOP_CodigoModalidade1," _
           & "FOP_Valor," _
           & "FOP_Parcelas," _
           & "FOP_Protocolo," _
           & "FOP_Situacao, " _
           & "FOP_NumeroPedido)" _
           & "Values ('" & Trim(wlblloja) _
           & "',0,'" & txtSerie.Text & "','" & CodigoModalidade & "'," _
           & ConverteVirgula(CDbl(txtValorModalidade.Text)) & ",01," _
           & wNumeroCaixa & ",'I', " & Pedido & " )"
           
           rdoCNLoja.Execute SQL
           Screen.MousePointer = vbNormal
           rdoCNLoja.CommitTrans
   ElseIf Not RsDados.EOF Then
              rdoCNLoja.BeginTrans
              Screen.MousePointer = vbHourglass
             
              SQL = "Update FormaPagamento set FOP_Valor=( FOP_Valor + " _
                    & ConverteVirgula(CDbl(txtValorModalidade.Text)) _
                    & ") where FOP_Loja='" & Trim(wlblloja) _
                    & "' and FOP_Numeropedido=" & Pedido & " and FOP_Serie='" _
                    & txtSerie.Text & "' and  FOP_CodigoModalidade1='" & CodigoModalidade & "'"
              
             rdoCNLoja.Execute SQL
             Screen.MousePointer = vbNormal
             rdoCNLoja.CommitTrans
   End If
   RsDados.Close
  

End Sub
Private Sub EstonaFormaPagamento()
    rdoCNLoja.BeginTrans
    Screen.MousePointer = vbHourglass
    
     SQL = "Update FormaPagamento set FOP_Valor=( FOP_Valor - " _
         & ConverteVirgula(CDbl(txtValorModalidade.Text)) _
         & ") Where FOP_Loja='" & Trim(wlblloja) & "' and FOP_NumeroPedido=" _
         & Pedido & " and FOP_Serie='" & txtSerie.Text & "' and " _
         & "FOP_CodigoModalidade1='" & CodigoModalidade & "'"
         rdoCNLoja.Execute SQL
         Screen.MousePointer = vbNormal
         rdoCNLoja.CommitTrans
End Sub

Private Sub LimpaGrid(ByRef GradeUsu)
    GradeUsu.Rows = GradeUsu.FixedRows + 1
    GradeUsu.AddItem ""
    GradeUsu.RemoveItem GradeUsu.FixedRows
End Sub

Private Sub FechaVenda()

    rdoCNLoja.BeginTrans
    Screen.MousePointer = vbHourglass
    SQL = "Update FormaPagamento set FOP_Situacao = 'O' " _
    & " where FOP_NumeroPedido =  " & txtPedido.Text
    rdoCNLoja.Execute SQL
    Screen.MousePointer = vbNormal
    rdoCNLoja.CommitTrans
            
    NroItens = 0
    wTotalVenda = 0
    wtotalitens = 0

End Sub

Private Sub VerificaTipoModalidade()
      
      lblTootip.Visible = True
      lblTootip1.Visible = True
  
             If Faturada = True Then
                If EntFaturada <> "0.00" Then
                   fraRecebimento.BackColor = &HC00000
                   
                   lblEntrada.Top = 720
                   lblEntrada.Visible = True
                   lblEntrada.Text = "ENT.FAT.        R$ "
                   txtValoraPagar.Text = Format(EntFaturada, "0.00")
                   lblApagar.Text = Format(EntFaturada, "0.00")
                   fraRecebimento.Visible = True
                   fraRecebimento.ZOrder
                  
                Else
                   lblFatFin.Top = 720
                   lblFatFin.Visible = True
                   lblFatFin.Text = "FATURADA     R$ "
                   lblValorFatFin.Top = lblFatFin.Top
                   lblValorFatFin.Text = Format(ValorFaturada, "0.00")
                   fraRecebimento.Visible = True
                   fraRecebimento.ZOrder
                   lblModalidade.Caption = "Modalidade"
                   chbOK.Enabled = True
                 
                End If
             ElseIf Financiada = True Then
                If EntFinanciada <> "0.00" Then
                   fraRecebimento.BackColor = &HC00000
                   lblModalidade.BackColor = &HC00000
                   lblEntrada.Top = 720
                   lblEntrada.Visible = True
                   lblEntrada.Text = "ENT.FIN.         R$ "
                   txtValoraPagar.Text = Format(EntFinanciada, "0.00")
                   lblApagar.Text = Format(EntFinanciada, "0.00")
                   fraRecebimento.Visible = True
                   fraRecebimento.ZOrder
                  
                Else
                   lblFatFin.Top = 720
                   lblFatFin.Visible = True
                   lblFatFin.Text = "FINANCIADA   R$ "
                   lblValorFatFin.Top = lblFatFin.Top
                   lblValorFatFin.Text = Format(ValorFinanciada, "0.00")
             
                   fraRecebimento.Visible = True
                   fraRecebimento.ZOrder
                   lblModalidade.Caption = "Modalidade"
                   chbOK.Enabled = True
                  
                End If
             ElseIf wVerificaAVR = True Then
                    lblFatFin.Top = 720
                    lblFatFin.Visible = True
                    lblFatFin.Text = "A V R       "
                    
                    chbOK.Enabled = True
                    'chbOK.Visible = True
                   ' Esperar 1
                    fraRecebimento.Visible = True
                    fraRecebimento.ZOrder
                    lblModalidade.Caption = "Modalidade"
                     fraRecebimento.Visible = True
                     fraRecebimento.ZOrder
                     txtValorModalidade.Text = lblApagar.Text
                    ' fraFormaPagamento.Enabled = False
                     lblModalidade.Caption = "Modalidade"
                     fraNModalidades.Visible = False
                     chbOK.Visible = True
                 
             Else
            
                lblEntrada.Visible = False
                fraRecebimento.Visible = True
                fraRecebimento.ZOrder
            If txtTipoNota.Text = "Romaneio" Then
               txtValorModalidade.Text = lblApagar.Text
               txtValorModalidade.Text = txtValoraPagar.Text
              ' fraFormaPagamento.Enabled = False
               chbOK.Enabled = True
             End If
                
       End If

End Sub



Private Function VerteclaVirgula(ByRef Controle As Control, ByRef Tecla As Integer)

'-- * -- Aceita apenas digitação de números e o sinal de "," -- * -- '
   If Controle.SelStart = 0 And Controle.SelLength = Len(Controle.Text) Then
      Controle.Text = ""
   End If
    
   
   
   If Tecla <> 13 Then
      If Chr(Tecla) = "," Or Chr(Tecla) = "." Then
         If InStr(Controle.Text, ",") <> 0 Or InStr(Controle.Text, ".") <> 0 Then
            Tecla = 0
         Else
            Tecla = Asc(",")
         End If
      ElseIf Not IsNumeric(Chr(Tecla)) And Tecla <> vbKeyBack Then
         Tecla = 0
      End If
   End If

End Function

