VERSION 5.00
Begin VB.Form frmFormaPagamento 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Forma de Pagamento"
   ClientHeight    =   8640
   ClientLeft      =   3060
   ClientTop       =   1635
   ClientWidth     =   13425
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   8.25
      Charset         =   0
      Weight          =   900
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   13425
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNumeroTEF 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   6045
      TabIndex        =   57
      Text            =   "0"
      Top             =   330
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame framePagamentoTEF 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   2145
      Left            =   5550
      TabIndex        =   55
      Top             =   5235
      Visible         =   0   'False
      Width           =   4755
      Begin VB.Frame framePagamentoTEFInterno 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   1515
         Left            =   0
         TabIndex        =   59
         Top             =   0
         Width           =   4755
         Begin VB.CommandButton cmdTefDebito 
            Appearance      =   0  'Flat
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
            Left            =   0
            Picture         =   "frmFormaPagamento.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   0
            Width           =   2340
         End
         Begin VB.CommandButton cmdTefCredito 
            Appearance      =   0  'Flat
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
            Left            =   2370
            Picture         =   "frmFormaPagamento.frx":1AA2
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   0
            Width           =   2340
         End
         Begin VB.Label lblMensagemTEF 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Mensagem TEF"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   570
            Left            =   15
            TabIndex        =   62
            Top             =   900
            Width           =   4725
         End
      End
      Begin Balcao2010.chameleonButton cmdRetornaOperacaoTEF 
         Height          =   465
         Left            =   1155
         TabIndex        =   58
         Top             =   1650
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   820
         BTYPE           =   14
         TX              =   "Cancelar Opera��o"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmFormaPagamento.frx":3A18
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Timer timeHabilitaTEF 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12135
      Top             =   7845
   End
   Begin Balcao2010.chameleonButton chbSair 
      Height          =   0
      Left            =   4605
      TabIndex        =   41
      Top             =   -45
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   0
      BTYPE           =   11
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFormaPagamento.frx":3A34
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraFinanciadoFaturado 
      BackColor       =   &H80000012&
      Height          =   3555
      Left            =   8055
      TabIndex        =   35
      Top             =   570
      Width           =   5070
      Begin Balcao2010.chameleonButton chbOkFat 
         Height          =   570
         Left            =   2235
         TabIndex        =   9
         Top             =   2760
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   1005
         BTYPE           =   3
         TX              =   "OK"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFormaPagamento.frx":3A50
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Balcao2010.chameleonButton chbValoraPagarFat 
         Height          =   555
         Left            =   2250
         TabIndex        =   10
         Top             =   870
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "0"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFormaPagamento.frx":3A6C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Balcao2010.chameleonButton chbValorEntrada 
         Height          =   555
         Left            =   2265
         TabIndex        =   11
         Top             =   1500
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "0"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFormaPagamento.frx":3A88
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Balcao2010.chameleonButton chbConfimaEntrada 
         Height          =   570
         Left            =   2235
         TabIndex        =   42
         Top             =   2760
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   1005
         BTYPE           =   3
         TX              =   "Confirma Entrada"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFormaPagamento.frx":3AA4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblParcelasFat 
         AutoSize        =   -1  'True
         BackColor       =   &H00AE7411&
         BackStyle       =   0  'Transparent
         Caption         =   "Parcelas"
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
         Left            =   165
         TabIndex        =   39
         Top             =   2175
         Width           =   945
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   165
         TabIndex        =   38
         Top             =   960
         Width           =   1860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00AE7411&
         BackStyle       =   0  'Transparent
         Caption         =   "Entrada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   195
         TabIndex        =   37
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblFinanciadoFaturado 
         BackStyle       =   0  'Transparent
         Caption         =   "Faturado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   195
         TabIndex        =   36
         Top             =   270
         Width           =   1665
      End
   End
   Begin VB.Frame fraPagamento 
      BackColor       =   &H00000000&
      Height          =   7515
      Left            =   150
      TabIndex        =   26
      Top             =   45
      Width           =   5220
      Begin VB.CheckBox chbPOS 
         BackColor       =   &H00000000&
         Caption         =   "POS"
         Enabled         =   0   'False
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
         Height          =   315
         Left            =   255
         MaskColor       =   &H00000000&
         TabIndex        =   44
         Top             =   4125
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   810
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
         Left            =   2640
         TabIndex        =   56
         Top             =   2925
         Width           =   2310
      End
      Begin VB.CommandButton cmdCondicao 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Condi��o "
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
         Left            =   60
         MaskColor       =   &H00C0E0FF&
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   7110
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame fraNModalidades 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
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
         Height          =   2970
         Left            =   255
         TabIndex        =   27
         Top             =   4410
         Width           =   4785
         Begin VB.Frame frameCartoes 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1635
            Left            =   675
            TabIndex        =   46
            Top             =   1755
            Width           =   4725
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
               Left            =   1185
               Picture         =   "frmFormaPagamento.frx":3AC0
               Style           =   1  'Graphical
               TabIndex        =   52
               Top             =   0
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
               Left            =   0
               Picture         =   "frmFormaPagamento.frx":6316
               Style           =   1  'Graphical
               TabIndex        =   51
               Top             =   0
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
               Left            =   2370
               Picture         =   "frmFormaPagamento.frx":720B
               Style           =   1  'Graphical
               TabIndex        =   50
               Top             =   0
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
               Left            =   0
               Picture         =   "frmFormaPagamento.frx":8296
               Style           =   1  'Graphical
               TabIndex        =   49
               Top             =   780
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
               Left            =   3555
               Picture         =   "frmFormaPagamento.frx":B398
               Style           =   1  'Graphical
               TabIndex        =   48
               Top             =   0
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
               Left            =   1185
               Picture         =   "frmFormaPagamento.frx":1010E
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   780
               Width           =   1155
            End
            Begin Balcao2010.chameleonButton cmdTrocar 
               Height          =   795
               Left            =   2340
               TabIndex        =   54
               Top             =   750
               Width           =   2400
               _ExtentX        =   4233
               _ExtentY        =   1402
               BTYPE           =   14
               TX              =   "Trocar Bandeira"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
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
               MICON           =   "frmFormaPagamento.frx":1357C
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
         End
         Begin VB.CommandButton chbCielo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            DownPicture     =   "frmFormaPagamento.frx":13598
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
            Left            =   2370
            Picture         =   "frmFormaPagamento.frx":144F3
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   825
            Width           =   2340
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
            Left            =   0
            Picture         =   "frmFormaPagamento.frx":15CC5
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   45
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
            Left            =   1185
            Picture         =   "frmFormaPagamento.frx":18BE7
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   45
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
            Left            =   3555
            Picture         =   "frmFormaPagamento.frx":1B6E9
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   45
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
            Left            =   2370
            Picture         =   "frmFormaPagamento.frx":1E06B
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   45
            Width           =   1155
         End
         Begin Balcao2010.chameleonButton chbSaiPagamento 
            Height          =   465
            Left            =   1170
            TabIndex        =   1
            Top             =   2460
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   820
            BTYPE           =   14
            TX              =   "Sair Pagamento"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
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
            MICON           =   "frmFormaPagamento.frx":20CED
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.CommandButton chbRede 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            DownPicture     =   "frmFormaPagamento.frx":20D09
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
            Left            =   0
            Picture         =   "frmFormaPagamento.frx":21B03
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   825
            Width           =   2340
         End
      End
      Begin Balcao2010.chameleonButton chbTroco 
         Height          =   855
         Left            =   150
         TabIndex        =   29
         Top             =   2115
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1508
         BTYPE           =   13
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
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
         MICON           =   "frmFormaPagamento.frx":22E3E
         PICN            =   "frmFormaPagamento.frx":22E5A
         PICH            =   "frmFormaPagamento.frx":24272
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
         Height          =   585
         Left            =   2640
         TabIndex        =   7
         Top             =   1590
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   1032
         BTYPE           =   3
         TX              =   "0"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFormaPagamento.frx":266A4
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
         Height          =   585
         Left            =   2640
         TabIndex        =   6
         Top             =   930
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   1032
         BTYPE           =   3
         TX              =   "0"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFormaPagamento.frx":266C0
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
         Height          =   585
         Left            =   2640
         TabIndex        =   0
         Top             =   270
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   1032
         BTYPE           =   3
         TX              =   "0"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFormaPagamento.frx":266DC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Balcao2010.chameleonButton chbOkPag 
         Height          =   570
         Left            =   2640
         TabIndex        =   8
         Top             =   3570
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   1005
         BTYPE           =   3
         TX              =   "OK"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFormaPagamento.frx":266F8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Balcao2010.chameleonButton chbvalortroco 
         Height          =   585
         Left            =   2640
         TabIndex        =   43
         Top             =   2250
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   1032
         BTYPE           =   3
         TX              =   "0"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFormaPagamento.frx":26714
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   -1  'True
      End
      Begin VB.Label lblParcelas 
         AutoSize        =   -1  'True
         BackColor       =   &H00AE7411&
         BackStyle       =   0  'Transparent
         Caption         =   "Parcelas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000BB&
         Height          =   300
         Left            =   255
         TabIndex        =   40
         Top             =   3840
         Width           =   4575
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
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   255
         TabIndex        =   34
         Top             =   3045
         Width           =   1650
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
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   255
         TabIndex        =   33
         Top             =   390
         Width           =   1860
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
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   255
         TabIndex        =   32
         Top             =   1050
         Width           =   1560
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
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   255
         TabIndex        =   31
         Top             =   1710
         Width           =   1125
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
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1140
         TabIndex        =   30
         Top             =   2370
         Visible         =   0   'False
         Width           =   840
      End
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
      Left            =   180
      TabIndex        =   17
      Top             =   -1545
      Visible         =   0   'False
      Width           =   4980
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         Left            =   210
         TabIndex        =   18
         Top             =   720
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
            Top             =   705
            Width           =   1350
         End
      End
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
      TabIndex        =   16
      Top             =   180
      Visible         =   0   'False
      Width           =   660
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   12
      Top             =   210
      Visible         =   0   'False
      Width           =   855
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
Dim sql As String

Dim ativaAtalho As Boolean

Dim ContadorItens As Long
Dim CodigoModalidade As String
Dim ValoraPagar As Double
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
Dim wValorGE As Double

Dim NomeCartao As String

Dim ValCartaoVisa As Double
Dim ValCartaoMastercard As Double
Dim ValCartaoAmex As Double
Dim ValCartaoBNDES As Double
Dim ValorPagamentoCartao As Double

Dim bandeiraTEFVisaElectron As String
Dim bandeiraTEFRedeShop As String
Dim bandeiraTEFHiperCard As String
Dim bandeiraCartaoVisa As String
Dim bandeiraCartaoMastercard As String
Dim bandeiraCartaoAmex As String

Dim ValTEFVisaElectron As Double
Dim ValTEFRedeShop As Double
Dim ValTEFHiperCard As Double

Dim ValNotaCredito As Double
Dim AvistaReceber As Double
Dim modalidade As Double
Dim wParcelas As Long
Dim wIndicePreco As Double
Dim wVerificaAVR As Boolean
Dim wGrupo As Long
Dim ValorFinanciada As Double

Dim wCodigoModalidadeDINHEIRO As String
Dim WCodigoModalidadeAMEX As String
Dim WCodigoModalidadeCHEQUE As String
Dim wCodigoModalidadeBNDES As String
Dim wCodigoModalidadeMASTERCARD As String
Dim wCodigoModalidadeNOTACREDITO As String
Dim wCodigoModalidadeFINANCIADO As String
Dim wCodigoModalidadeFATURADO As String
Dim wTEFVisaElectron As String
Dim wTEFRedeShop As String
Dim wTEFHiperCard As String
Dim WCodigoModalidadeVISA As String
Dim wtempo As Long
Dim wMostraGrideCondicao As Boolean

Dim wControle As String
 
Dim wvalorparcelas As Double
Dim nf As notaFiscalTEF


Dim wValorModalidadeIncorreto As Boolean


Dim wGrupoMovimento As String
Dim wSubGrupo As String
Dim wValorMovimento As Double
Dim wValorTotalItem As String * 10

Dim wDescricao As String * 29
Dim wAliquota As String * 5
Dim wPrecoVenda As String * 8
Dim wDesconto As String * 8
Dim wCodigoProduto As String * 13
Dim wQtde As String * 4
Dim wQuantidade As Integer
Dim wTipoQuantidade As String * 1
Dim wCasaDecimais As Integer
Dim wTipoDesconto As String * 1

'---------------------------------tef ----------------
Dim cFormaPGTO As String
Dim cNumeroCupom As String
Dim ValorPago() As String
Dim hora As Date
Dim pagtoCartao As Boolean
Dim acumulado As Double
Public i As Integer
Public j As Integer
Dim iConta As Integer
Dim curSubTotal As Currency
Dim curValorRestante As Currency
Dim cValorPago As String
''---------------------
Dim valValorPago As Double
Dim valValorFalta As Double
Dim valValoraPagar As Double
Dim tempoHabilitaPOS As Byte

Dim primeiroCarregamento As Boolean

Dim Agencia As String
Dim tipoNotaMovimentoCaixa As String



Private Sub GravaRegistro()

    Wecf = GLB_ECF

If EntFaturada = "0.00" And EntFinanciada = "0.00" Then

      sql = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
      & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa,MC_Pedido,MC_dataprocesso,MC_TipoNota,MC_SequenciaTEF,MC_DataBaixaAVR) values(" & Wecf & ",'" & GLB_USU_Codigo & "','" & Trim(wlblloja) & "', " _
      & " '" & Format(GLB_DataInicial, "yyyy/mm/dd") & "', " & wGrupoMovimento & ",'" & wSubGrupo & "', " & NroNotaFiscal & ",'" & txtSerie.text & "', " _
      & "" & ConverteVirgula(Format(wValorMovimento, "##,###0.00")) & ", " _
      & "0,'" & Agencia & "',0,0," & wParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & GLB_Caixa & ",'" & txtPedido.text & "','" & Format(GLB_DataInicial, "yyyy/mm/dd") & "','" & tipoNotaMovimentoCaixa & "','" & txtNumeroTEF.text & "', '" & Format(horaOperacaoTEF, "HH:MM:SS") & "')"
      rdoCNLoja.Execute (sql)
      
Else
   sql = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
      & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa,MC_Pedido,MC_DataProcesso,MC_TipoNota,MC_SequenciaTEF,MC_DataBaixaAVR) values(" & Wecf & ",'" & GLB_USU_Codigo & "','" & Trim(wlblloja) & "', " _
      & " '" & Format(GLB_DataInicial, "yyyy/mm/dd") & "', " & wGrupoMovimento & ",'" & wSubGrupo & "', " & NroNotaFiscal & ",'" & txtSerie.text & "', " _
      & "" & ConverteVirgula(Format(wValorMovimento, "##,###0.00")) & ", " _
      & "0,'" & Agencia & "',0,0," & wParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & GLB_Caixa & ",'" & txtPedido.text & "','" & Format(GLB_DataInicial, "yyyy/mm/dd") & "','" & tipoNotaMovimentoCaixa & "','" & txtNumeroTEF.text & "','" & horaOperacaoTEF & "')"
      rdoCNLoja.Execute (sql)
      
   If EntFaturada <> "0.00" Then
      sql = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
         & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa,MC_Pedido, MC_DataProcesso,MC_TipoNota,MC_SequenciaTEF,MC_DataBaixaAVR) values(" & Wecf & ",'" & GLB_USU_Codigo & "','" & Trim(wlblloja) & "', " _
         & " '" & Format(GLB_DataInicial, "yyyy/mm/dd") & "', " & 11004 & ",'" & wSubGrupo & "', " & NroNotaFiscal & ",'" & txtSerie.text & "', " _
         & "" & ConverteVirgula(Format(EntFaturada, "##,###0.00")) & ", " _
         & "0,'" & Agencia & "',0,0," & wParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & GLB_Caixa & ",'" & txtPedido.text & "','" & Format(GLB_DataInicial, "yyyy/mm/dd") & "','" & tipoNotaMovimentoCaixa & "','" & txtNumeroTEF.text & "','" & horaOperacaoTEF & "')"
         rdoCNLoja.Execute (sql)
   ElseIf EntFinanciada <> "0.00" Then
       sql = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
             & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa,MC_Pedido, MC_DataProcesso,MC_TipoNota,MC_SequenciaTEF,MC_DataBaixaAVR) values(" & Wecf & ",'" & GLB_USU_Codigo & "','" & Trim(wlblloja) & "', " _
             & " '" & Format(GLB_DataInicial, "yyyy/mm/dd") & "', " & 11005 & ",'" & wSubGrupo & "', " & NroNotaFiscal & ",'" & txtSerie.text & "', " _
             & "" & ConverteVirgula(Format(EntFinanciada, "##,###0.00")) & ", " _
             & "0,'" & Agencia & "',0,0," & wParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & GLB_Caixa & ",'" & txtPedido.text & "','" & Format(GLB_DataInicial, "yyyy/mm/dd") & "','" & tipoNotaMovimentoCaixa & "','" & txtNumeroTEF.text & "','" & horaOperacaoTEF & "')"
             rdoCNLoja.Execute (sql)
   End If
End If
   
End Sub




Private Function obterSequenciaMovimentoCaixa(pedido As String)

    Dim sql As String
    Dim rsConsulta As New ADODB.Recordset

    sql = "SELECT top 1 MC_Sequencia FROM  MovimentoCaixa" & vbNewLine & _
          "where MC_GRUPO = '99999'" & vbNewLine & _
          "and MC_PEDIDO = '" & pedido & "'" & vbNewLine & _
          ""
    
    rsConsulta.CursorLocation = adUseClient
    rsConsulta.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not rsConsulta.EOF Then
        
        obterSequenciaMovimentoCaixa = rsConsulta("MC_Sequencia")

    End If
    
    rsConsulta.Close

    
    sql = "update movimentocaixa set MC_GRUPO = '10101'" & vbNewLine & _
          "where MC_GRUPO = '99999'" & vbNewLine & _
          "and MC_PEDIDO = '" & pedido & "'" & vbNewLine & _
          ""
    
    rdoCNLoja.Execute sql
    

End Function

Private Sub GuardaValoresParaGravarMovimentoCaixa()

'wValorModalidadeIncorreto = False
      Dim nf As notaFiscalTEF
      
      tipoNotaMovimentoCaixa = "PA"
      
      If Trim(txtValorModalidade.text = "") Then
         Exit Sub
      End If
      
      If Trim(txtValorModalidade.text = ",") Then
         Exit Sub
      End If

'      Nf.pedido = pedido
      'Nf.numero = NroNotaFiscal
      'Nf.serie = txtSerie.text
      'Nf.Parcelas = wParcelas
      'Nf.dataEmissao = GLB_DataInicial
      'Nf.valor = txtValorModalidade.text

      modalidade = Format(txtValorModalidade.text, "0.00")

         If lblModalidade.Caption <> "DINHEIRO" Then
         
            'If ((modalidade + TotPago) - valValoraPagar) > ValDinheiro _
            'And valValoraPagar < (modalidade + TotPago) Then
            'If ((modalidade + TotPago) - valValoraPagar) > 0 Or modalidade <= 0 Then
                'MsgBox "N�o � permitido troco maior que pagamento em dinheiro"
                'Exit Sub
            'End If
         
         End If
         

      
      nf.pedido = txtPedido.text
      nf.serie = txtSerie.text
      nf.numero = NroNotaFiscal
      nf.dataEmissao = GLB_DataInicial
      nf.valor = txtValorModalidade
      nf.Parcelas = wParcelas
      nf.numeroTEF = 0
      
      'TEF ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
      habilitaFrameTEFOperacoes False
      
           If lblModalidade.Caption = "CREDITO" Then
      
              tipoNotaMovimentoCaixa = "TF"
      
              operacaoTEFCompleta = False
              CodigoModalidade = "0101"
              wCodigoModalidadeDINHEIRO = "0101"
              wGrupoMovimento = "99999"
              wSubGrupo = ""
              ValTEFVisaElectron = ValTEFVisaElectron + modalidade
              wValorMovimento = Format(ValTEFVisaElectron, "##,###0.00")
              Call GravaRegistro
              nf.sequenciaMovimentoCaixa = obterSequenciaMovimentoCaixa(txtPedido.text)
              ValTEFVisaElectron = 0
              txtNumeroTEF.text = ""
              
      
          If EfetuaOperacaoTEF("3", nf, lblModalidade, lblMensagemTEF) Then
            operacaoTEFCompleta = True
            TotPago = TotPago + modalidade
            txtNumeroTEF.text = nf.numeroTEF
            
            'carregaCodigoModalidade lblModalidade.Caption
            
            ComprovantePagamentoFila = ComprovantePagamentoFila & nf.comprovantePagamento
            lblModalidade.Caption = ""
          Else
            lblModalidade.Caption = ""
          End If
          
      End If

      If lblModalidade.Caption = "DEBITO" Then
      
              tipoNotaMovimentoCaixa = "TF"
      
              operacaoTEFCompleta = False
              CodigoModalidade = "0101"
              wCodigoModalidadeDINHEIRO = "0101"
              wGrupoMovimento = "99999"
              wSubGrupo = ""
              ValTEFVisaElectron = ValTEFVisaElectron + modalidade
              wValorMovimento = Format(ValTEFVisaElectron, "##,###0.00")
              Call GravaRegistro
              nf.sequenciaMovimentoCaixa = obterSequenciaMovimentoCaixa(txtPedido.text)
              ValTEFVisaElectron = 0
              txtNumeroTEF.text = ""
              
      
          If EfetuaOperacaoTEF("2", nf, lblModalidade, lblMensagemTEF) Then
          
            operacaoTEFCompleta = True
            TotPago = TotPago + modalidade
        
            txtNumeroTEF.text = nf.numeroTEF
            'carregaCodigoModalidade lblModalidade.Caption
            ComprovantePagamentoFila = ComprovantePagamentoFila & nf.comprovantePagamento
            lblModalidade.Caption = ""
          Else
            lblModalidade.Caption = ""
          End If
      End If
      
      If lblModalidade.Caption = "Opera��o 0" Then
          tipoNotaMovimentoCaixa = "PA"
          If EfetuaOperacaoTEF("0", nf, lblModalidade, lblMensagemTEF) Then
            txtNumeroTEF.text = nf.numeroTEF
            carregaCodigoModalidade lblModalidade.Caption
            ComprovantePagamentoFila = ComprovantePagamentoFila & nf.comprovantePagamento
          Else
            lblModalidade.Caption = ""
          End If
      End If
      
      habilitaFrameTEFOperacoes True
      
      
      'TEF ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            
      If lblModalidade.Caption = "DINHEIRO" Then
           TotPago = TotPago + modalidade
           ValDinheiro = ValDinheiro + modalidade
      End If
      

      If lblModalidade.Caption = "CHEQUE" Then
         ValCheque = ValCheque + modalidade
         TotPago = TotPago + modalidade
         
    
            wGrupoMovimento = "10201"
            wSubGrupo = ""
            wValorMovimento = Format(ValCheque, "##,##0.00")
            If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Or _
                UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" Then
                retorno = Bematech_FI_EfetuaFormaPagamento("CHEQUE", ValCheque * 100)
                Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
            End If
            Call GravaRegistro
            
            ValCheque = 0
            txtNumeroTEF.text = ""
         
      End If

      If lblModalidade.Caption = "VISA ELEC." Then
          TotPago = TotPago + modalidade
          ValTEFVisaElectron = ValTEFVisaElectron + modalidade
          bandeiraTEFVisaElectron = Agencia
          

            'TEF
            wGrupoMovimento = "10206"
            wSubGrupo = "Visa Elec."
            Agencia = bandeiraTEFVisaElectron
            wValorMovimento = Format(ValTEFVisaElectron, "##,###0.00")
    
            If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Or _
                UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" Then
                retorno = Bematech_FI_EfetuaFormaPagamento("TEF", ValTEFVisaElectron * 100)
                Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
            End If
    
            Call GravaRegistro
            
            ValTEFVisaElectron = 0
            txtNumeroTEF.text = ""
          
      End If

      If lblModalidade.Caption = "REDESHOP" Then
          ValTEFRedeShop = ValTEFRedeShop + modalidade
          TotPago = TotPago + modalidade
          bandeiraTEFRedeShop = Agencia
          

            'TEF
            wGrupoMovimento = "10203"
            wSubGrupo = "RedeShop"
            Agencia = bandeiraTEFRedeShop
            wValorMovimento = Format(ValTEFRedeShop, "##,###0.00")
    
            If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Or _
                UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" Then
                retorno = Bematech_FI_EfetuaFormaPagamento("TEF", ValTEFRedeShop * 100)
                Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
            End If
    
            Call GravaRegistro

            ValTEFRedeShop = 0
            txtNumeroTEF.text = ""
            
          
      End If

      If lblModalidade.Caption = "HIPERCARD" Then
          TotPago = TotPago + modalidade
          ValTEFHiperCard = ValTEFHiperCard + modalidade
          bandeiraTEFHiperCard = Agencia
          
            wGrupoMovimento = "10205"
            wSubGrupo = "HiperCard"
            Agencia = bandeiraTEFHiperCard
            wValorMovimento = Format(ValTEFHiperCard, "##,##0.00")
    
            If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Or _
                UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" Then
                retorno = Bematech_FI_EfetuaFormaPagamento("TEF", ValTEFHiperCard * 100)
                Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
            End If
            
            Call GravaRegistro
    
            ValTEFHiperCard = 0
            txtNumeroTEF.text = ""
          
      End If

      If lblModalidade.Caption = "NOTA DE CR�D." Then
          TotPago = TotPago + modalidade
          ValNotaCredito = ValNotaCredito + modalidade
          
    
            'NOTA DE CREDITO
            wGrupoMovimento = "10701"
            wSubGrupo = ""
            wValorMovimento = Format(ValNotaCredito, "##,##0.00")
            
            If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Or _
                UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" Then
                retorno = Bematech_FI_EfetuaFormaPagamento("NC", ValNotaCredito * 100)
                Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
            End If
            
            Call GravaRegistro
        
            ValNotaCredito = 0
            txtNumeroTEF.text = ""
          
      End If

      If lblModalidade.Caption = "VISA" Then
          ValCartaoVisa = ValCartaoVisa + modalidade
          TotPago = TotPago + modalidade
          bandeiraCartaoVisa = Agencia
        
            wGrupoMovimento = "10301"
            Agencia = bandeiraCartaoVisa
            wSubGrupo = ""
            wValorMovimento = Format(ValCartaoVisa, "##,##0.00")
    
                If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Or _
                    UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" Then
                    retorno = Bematech_FI_EfetuaFormaPagamento("VISA", ValCartaoVisa * 100)
                    Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
    
                    If retorno <> 1 Then
                        MsgBox "Por favor verificar se impressora est� ligada corretamente!"
                        Exit Sub
                    End If
                End If
    
            Call GravaRegistro
            
            ValCartaoVisa = 0
            txtNumeroTEF.text = ""
          
      End If

      If lblModalidade.Caption = "MASTERCARD" Then
          ValCartaoMastercard = ValCartaoMastercard + modalidade
          TotPago = TotPago + modalidade
          bandeiraCartaoMastercard = Agencia
          
          
            wGrupoMovimento = "10302"
            Agencia = bandeiraCartaoMastercard
            wSubGrupo = ""
            wValorMovimento = Format(ValCartaoMastercard, "##,##0.00")
            If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Or _
            UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" Then
            retorno = Bematech_FI_EfetuaFormaPagamento("MASTERCARD", ValCartaoMastercard * 100)
            Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
            End If
            Call GravaRegistro
                 
            ValCartaoMastercard = 0
            wCodigoModalidadeMASTERCARD = ""
            txtNumeroTEF.text = ""
          
        End If

      If lblModalidade.Caption = "AMEX" Then
          ValCartaoAmex = ValCartaoAmex + modalidade
          TotPago = TotPago + modalidade
          bandeiraCartaoAmex = Agencia
          
            wGrupoMovimento = "10303"
            Agencia = bandeiraCartaoAmex
            wSubGrupo = ""
            wValorMovimento = Format(ValCartaoAmex, "##,##0.00")

            If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Or _
            UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" Then
                retorno = Bematech_FI_EfetuaFormaPagamento("AMEX", ValCartaoAmex * 100)
                Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
            End If

            Call GravaRegistro
        
            ValCartaoAmex = 0
            txtNumeroTEF.text = ""
          
      End If

      If lblModalidade.Caption = "BNDES" Then
          ValCartaoBNDES = ValCartaoBNDES + modalidade
          TotPago = TotPago + modalidade
          

            'BNDES
            wGrupoMovimento = "10304"
            wSubGrupo = ""
            wValorMovimento = Format(ValCartaoBNDES, "##,##0.00")


            If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Or _
            UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" Then
                retorno = Bematech_FI_EfetuaFormaPagamento("AMEX", ValCartaoAmex * 100)
                Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
            End If
        
            Call GravaRegistro
            
            ValCartaoBNDES = 0
            txtNumeroTEF.text = ""
          
      End If

      
      txtValorModalidade.text = 0
      lblModalidade.Caption = " "
 
       If TotPago > valValoraPagar Then
          chbValorPago.Caption = Format(TotPago, "##,##0.00")
          chbValorFalta.Caption = Format(0, "##,##0.00")
          lblFaltaPagar.Visible = True
          chbvalortroco.top = 2280
          lblTroco.Visible = True
          chbTroco.Visible = True
          chbvalortroco.Caption = Format((TotPago - valValoraPagar), "##,##0.00")
       
       ElseIf TotPago = valValoraPagar Then
          chbValorPago.Caption = Format(TotPago, "##,##0.00")
          chbValorFalta.Caption = Format(0, "##,##0.00")
          chbvalortroco.Caption = Format((TotPago - valValoraPagar), "##,##0.00")
       Else
         chbValorPago.Caption = Format(TotPago, "##,##0.00")
         chbValorFalta.Caption = Format(valValoraPagar - TotPago, "##,##0.00")
         chbvalortroco.Caption = Format(0, "##,##0.00")
       End If
       ValTroco = chbvalortroco.Caption


End Sub


'Private Sub ZeraVariaveis()
'ValorPagamentoCartao = 0
'ValDinheiro = 0
'ValTroco = 0
'ValCheque = 0
'ValCartaoAmex = 0
'ValCartaoBNDES = 0
'ValCartaoMastercard = 0
'ValCartaoVisa = 0
'ValTEFVisaElectron = 0
'valValoraPagar = 0
'ValTEFRedeShop = 0
'ValTEFHiperCard = 0
'TotPago = 0
'Modalidade = 0
''wTEFRedeShop = 0
''wTEFHiperCard = 0
'ValNotaCredito = 0
'chbValorPago.Caption = 0
'chbValorPago.Caption = Format(chbValorPago.Caption, "##,###0.00")
'chbValoraPagar.Caption = Format(chbValorPago.Caption, "##,###0.00")
'chbValorFalta.Caption = Format(chbValoraPagar.Caption, "##,###0.00")
'txtValorModalidade.Text = ""
'
'
' wCodigoModalidadeDINHEIRO = ""
' WCodigoModalidadeAMEX = ""
' WCodigoModalidadeCHEQUE = ""
' wCodigoModalidadeBNDES = ""
' wCodigoModalidadeMASTERCARD = ""
' wCodigoModalidadeNOTACREDITO = ""
' wCodigoModalidadeFINANCIADO = ""
' wCodigoModalidadeFATURADO = ""
' wTEFVisaElectron = ""
' wTEFRedeShop = ""
' wTEFHiperCard = ""
' WCodigoModalidadeVISA = ""
'End Sub

Private Sub chameleonButton3_Click()
    fraPagamento.Visible = True
    chbDinheiro.SetFocus
    fraFinanciadoFaturado.left = 135
    fraFinanciadoFaturado.top = 510
    Call FormaPagamento
End Sub

Private Sub chameleonButton1_Click()
    
End Sub

Private Sub chbBNDES_Click()
    
 ' lblModalidade.Caption = "Cart�o"
 cmdTrocar_Click
  lblModalidade.Caption = "BNDES"
  'FraParcelas.Visible = True
  'lblParc.Visible = True
  lblParcelas.Visible = True
  txtValorModalidade.Enabled = True
  txtValorModalidade.SetFocus
  CodigoModalidade = "0304"
  wCodigoModalidadeBNDES = "0304"
 ' wPagamentoECF = 2
 ' wPagamentoECF = BuscaCodigoPagamentoTEF("TEF")
End Sub


Private Sub chbCielo_Click()
  frameCartoes.Visible = True
  Agencia = "025"
End Sub

Private Sub chbConfimaEntrada_Click()
       chbConfimaEntrada.Visible = False
       chbOkFat.Visible = True
       If lblFinanciadoFaturado.Caption = "Faturado" Then
         wCodigoModalidadeFATURADO = "0501"
       ElseIf lblFinanciadoFaturado.Caption = "Financiado" Then
         wCodigoModalidadeFINANCIADO = "0601"
       End If
       
       
End Sub



Public Sub chbOkFat_Click()
    
'If chbOkFat.Caption = "OK" Then

   'criaDuplicataBanco
    
    txtPedido.text = pedido
          
    Call GuardaValoresParaGravarMovimentoCaixa
    
    
    txtValorModalidade.text = ""
    txtValorModalidade.Enabled = False
    lblModalidade.Caption = " "
  
 '   FraParcelas.Visible = False
 '   lblParc.Visible = False
    
    
    'A pergunta abaixo � feita para que se o valor do troco for maior que o valor em dinheiro
    ' ou o valor do cartao > que o valor da nota, saia da rotina sem sumir o franModalidade.
    

    
    If wValorModalidadeIncorreto = True Then
       Exit Sub
    End If
    
    If chbValorFalta.Caption = "" Then
       chbValorFalta.Caption = 0
    End If
       
         
    If chbValorFalta.Caption <= 0 Then
       chbOkFat.Visible = True
       chbOkFat.Enabled = True
       chbSair.Enabled = False
       chbOkFat.SetFocus
       fraNModalidades.Visible = False
       txtValorModalidade.Visible = False
       lblModalidade.Visible = False
    End If
    
    Call FinalizaPagamento
'Else
'    chbOkFat.Caption = "OK"
'End If
    
End Sub

Private Sub chbOkFat_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
    If UCase(frmFormaPagamento.txtIdentificadequeTelaqueveio.text) <> "FRMCAIXATEFPEDIDO" And _
       UCase(frmFormaPagamento.txtIdentificadequeTelaqueveio.text) <> "FRMCAIXATEF" Then
       Unload Me
       Exit Sub
    End If
 End If
End Sub

Private Sub chbRede_Click()
    frameCartoes.Visible = True
    Agencia = "012"
End Sub

Private Sub chbSaiPagamento_Click()

If txtSerie.text Like GLB_SerieCF & "*" Then
 
   If MsgBox("Est� opera��o permite reiniciar o procedimento de Recebimento. Deseja Continuar?", vbQuestion + vbYesNo, "Aten��o") = vbYes Then
       wTipoQuantidade = "I"
       wCasaDecimais = 2
       wTipoDesconto = "$"
       wDesconto = 0
       wCupomAberto = False
       Call ZeraVariaveis
       Unload Me
       If txtIdentificadequeTelaqueveio.text = "FRMCAIXATEFPEDIDO" Then
           frmCaixaTEFPedido.fraNFP.Visible = True
       End If
   End If
  Else
  Unload Me
  End If
 
End Sub

Private Sub chbHiperCard_Click()
    carregaCodigoModalidade "HIPERCARD"
    frmcond.Visible = False
    lblParcelas.Visible = False
    txtValorModalidade.Enabled = True
    txtValorModalidade.SetFocus
End Sub

Private Sub chbRedeShop_Click()
    carregaCodigoModalidade "REDESHOP"
    frmcond.Visible = False
    lblParcelas.Visible = False
    txtValorModalidade.Enabled = True
    txtValorModalidade.SetFocus
End Sub
Private Sub chbSair_Click()
Unload Me
End Sub

Private Sub chbAmex_Click()
    carregaCodigoModalidade "AMEX"
    lblParcelas.Visible = True
    txtValorModalidade.Enabled = True
    txtValorModalidade.SetFocus
End Sub

Private Sub chbAmex_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    lblModalidade.Caption = ""
'    FraParcelas.Visible = False
 '   lblParc.Visible = False
 lblParcelas.Visible = False
End If
End Sub

'Private Sub chbCartao_Click()
'  frmcond.Visible = False
'  lblModalidade.Caption = "CART�O"
'  txtValorModalidade.Enabled = False
'  FraParcelas.Visible = True
'  lblParc.Visible = True
'  txtParcelas.SelStart = 0
'  txtParcelas.SelLength = Len(txtParcelas.Text)
  
'End Sub

Private Sub chbCheque_Click()
   
cmdTrocar_Click
  frmcond.Visible = False
'  FraParcelas.Visible = False
  'lblParc.Visible = False
  lblModalidade.Caption = "CHEQUE"
  txtValorModalidade.Enabled = True
  txtValorModalidade.SetFocus
  CodigoModalidade = "0201"
  WCodigoModalidadeCHEQUE = "0201"
'  wPagamentoECF = 7
'  wPagamentoECF = BuscaCodigoPagamentoTEF("cheque")

End Sub
Private Sub chbDinheiro_Click()
  
  cmdTrocar_Click
  frmcond.Visible = False
'  FraParcelas.Visible = False
' lblParcelas.Visible  = False
  lblModalidade.Caption = "DINHEIRO"
  txtValorModalidade.Enabled = True
  txtValorModalidade.SetFocus
 
  CodigoModalidade = "0101"
  wCodigoModalidadeDINHEIRO = "0101"

  'wPagamentoECF = BuscaCodigoPagamentoTEF("Dinheiro")

End Sub

Private Sub chbDinheiro_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
    If UCase(frmFormaPagamento.txtIdentificadequeTelaqueveio.text) <> "FRMCAIXATEFPEDIDO" And _
       UCase(frmFormaPagamento.txtIdentificadequeTelaqueveio.text) <> "FRMCAIXATEF" Then
       Unload Me
       Exit Sub
    End If
 End If
 
End Sub

Private Sub chbMasterCard_Click()
    carregaCodigoModalidade "MASTERCARD"
    lblParcelas.Visible = True
    txtValorModalidade.Enabled = True
    txtValorModalidade.SetFocus
End Sub

Private Sub carregaCodigoModalidade(modalidade As String)
    Select Case modalidade
    Case "MASTERCARD"
        lblModalidade.Caption = "MASTERCARD"
        CodigoModalidade = "0302"
        wCodigoModalidadeMASTERCARD = "0302"
        cFormaPGTO = "Cartao"
    Case "AMEX"
        lblModalidade.Caption = "AMEX"
        CodigoModalidade = "0303"
        WCodigoModalidadeAMEX = "0303"
        cFormaPGTO = "Cartao"
    Case "HIPERCARD"
        lblModalidade.Caption = "HIPERCARD"
        CodigoModalidade = "0405"
    Case "REDESHOP"
        lblModalidade.Caption = "REDESHOP"
        CodigoModalidade = "0402"
    Case "VISA"
        lblModalidade.Caption = "VISA"
        CodigoModalidade = "0301"
        WCodigoModalidadeVISA = "0301"
        cFormaPGTO = "Cartao"
    Case "VISA ELEC."
        lblModalidade.Caption = "VISA ELEC."
        CodigoModalidade = "0401"
    Case Else
        'MsgBox "A modalidade n�o foi reconhecida internamente no sistema. " _
             & "Essa modalidade ser� gravada como: DINHEIRO. " _
             & "Voc� poder� alterar modalidade futuramente.", vbInformation, "Modalidade desconhecida"
        lblModalidade.Caption = "DINHEIRO"
        CodigoModalidade = "0101"
        wCodigoModalidadeDINHEIRO = "0101"
    End Select
End Sub

Private Sub chbMasterCard_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    lblModalidade.Caption = ""
    lblParcelas.Visible = False
  '  FraParcelas.Visible = False
  '  lblParc.Visible = False
End If
End Sub

Private Sub chbNotaCredito_Click()
    
    cmdTrocar_Click
 
 
 
 frmcond.Visible = False
  'FraParcelas.Visible = False
  'lblParc.Visible = False
  lblParcelas.Visible = False
  lblModalidade.Caption = "NOTA DE CR�D."
  txtValorModalidade.Enabled = True
  txtValorModalidade.SetFocus
  CodigoModalidade = "0701"
  wCodigoModalidadeNOTACREDITO = "0701"
'  wPagamentoECF = 8
  'wPagamentoECF = BuscaCodigoPagamentoTEF("dinheiro")
End Sub

Private Sub FinalizaPagamento()
  
Dim rsControle As New ADODB.Recordset
  
GetAsyncKeyState (vbKeyTab)

frmcond.Visible = False
wRomaneio = False
'
'-- Colocar a mensagem na tabela parametro
'

sql = "Update Nfcapa set ECF = '" & GLB_ECF & "' where NumeroPed =  " & txtPedido.text
rdoCNLoja.Execute sql


If txtTipoNota.text = "CUPOM" Then

' ROTINA ECF (NAO APAGAR)
' Fecha o Cupom
   sql = ""
   sql = "select ve_Codigo,ve_nome,desconto,nf from nfcapa,vende where vendedor = ve_codigo and " & _
             "NumeroPed = " & txtPedido.text
             
   rsComplementoVenda.CursorLocation = adUseClient
   rsComplementoVenda.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

   If rsComplementoVenda("ve_Codigo") = 725 Then
      wAdicionaisECF = "Pedido: " & Trim(txtPedido.text) & "     " & rsComplementoVenda("ve_Codigo") & " - Caixa "
   Else
      wAdicionaisECF = "Pedido: " & Trim(txtPedido.text) & "     " & rsComplementoVenda("ve_Codigo") & " - " & _
                        rsComplementoVenda("ve_nome")
   End If


    NroNotaFiscal = rsComplementoVenda("nf")
    rsComplementoVenda.Close
    txtSerie.text = GLB_SerieCF
    GravaMovimentoCaixa
    EncerraVenda Val(txtPedido.text), " ", 1
    ' If EncerraVenda(Val(txtPedido.Text), " ", 1) = False Then
   '      Exit Sub
   '  End If

    retorno = 0
    retorno = Bematech_FI_TerminaFechamentoCupom(wAdicionaisECF)
    'Fun��o que analisa o retorno da impressora
    Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
    
    If retorno <> 1 Then
        MsgBox "Por favor verificar se impressora est� ligada corretamente!", vbCritical, "ERRO"
        Exit Sub

    End If



ElseIf txtTipoNota.text = "SAT" Then

' ROTINA ECF (NAO APAGAR)
' Fecha o Cupom
   sql = ""
   sql = "select ve_Codigo,ve_nome,desconto,nf from nfcapa,vende where vendedor = ve_codigo and " & _
             "NumeroPed = " & txtPedido.text
             
   rsComplementoVenda.CursorLocation = adUseClient
   rsComplementoVenda.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

   If rsComplementoVenda("ve_Codigo") = 725 Then
      wAdicionaisECF = "Pedido: " & Trim(txtPedido.text) & "     " & rsComplementoVenda("ve_Codigo") & " - Caixa "
   Else
      wAdicionaisECF = "Pedido: " & Trim(txtPedido.text) & "     " & rsComplementoVenda("ve_Codigo") & " - " & _
                        rsComplementoVenda("ve_nome")
   End If


    NroNotaFiscal = rsComplementoVenda("nf")
    rsComplementoVenda.Close
    txtSerie.text = GLB_SerieCF
    GravaMovimentoCaixa
    EncerraVenda Val(txtPedido.text), " ", 1
    ' If EncerraVenda(Val(txtPedido.Text), " ", 1) = False Then
   '      Exit Sub
   '  End If

    'Retorno = 0
    'Retorno = Bematech_FI_TerminaFechamentoCupom(wAdicionaisECF)
    'Fun��o que analisa o retorno da impressora
    'Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
    
    'If Retorno <> 1 Then
        'MsgBox "Por favor verificar se impressora est� ligada corretamente!", vbCritical, "ERRO"
        'Exit Sub

    'End If

   
ElseIf txtTipoNota.text = "NF" Then
     
  '************************ Verificando se Nota � Eletr�nica
        
       sql = ""
       sql = "select ce_Estado,ce_tipopessoa, cliente from fin_cliente,nfcapa where ce_CodigoCliente = Cliente and " & _
             "NumeroPed = " & txtPedido.text

            rsComplementoVenda.CursorLocation = adUseClient
            rsComplementoVenda.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                 
            sql = "select CTS_SerieNota from ControleSistema"
            rsControle.CursorLocation = adUseClient
            rsControle.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                 
                 
            If (rsControle("CTS_SerieNota") = "NE") Or RTrim(LTrim(rsComplementoVenda("ce_Estado"))) <> "SP" _
                 Or RTrim(LTrim(rsComplementoVenda("ce_TipoPessoa"))) = "O" Then
                 
                 'If RTrim(LTrim(rsComplementoVenda("cliente"))) <> "999999" Then

                 
                   txtSerie.text = "NE"
                   GLB_Pessoa = rsComplementoVenda("ce_tipopessoa")
                   
'--------------------------------------------------------------------------------------
                   NroNotaFiscal = ExtraiSeqNEControle
                   rdoCNLoja.BeginTrans
                   Screen.MousePointer = vbHourglass

                   sql = ""
                   sql = "Update NfItens set  Serie = 'NE' " & _
                         "where NumeroPed = " & pedido
                   rdoCNLoja.Execute (sql)

                   sql = ""
                   sql = "Update NfCapa set  Serie = 'NE', TM = 0 " & _
                         "where NumeroPed = " & pedido
                   rdoCNLoja.Execute (sql)
                   
                   If (rsControle("CTS_SerieNota") <> "NE") Then MsgBox "ESTE PEDIDO IR� GERAR A NOTA FISCAL ELETR�NICA N�MERO " & NroNotaFiscal & _
                           ", AVISE O CLIENTE.", vbInformation, "Aten��o"
                   rdoCNLoja.CommitTrans

                  Screen.MousePointer = vbNormal
'------------------------------------------------------K-----------------------------------------

                   'Comentado FELIPE 27/07/2015
                   'GravaMovimentoCaixa
                   
                   If EncerraVenda(Val(txtPedido.text), " ", 1) = False Then
                      rsComplementoVenda.Close
                      Exit Sub
                   End If
                   
                   GravaMovimentoCaixa
                   
            Else
            
                  NroNotaFiscal = ExtraiSeqNotaControle
                  rdoCNLoja.BeginTrans
                   Screen.MousePointer = vbHourglass
                  
''''                FELIPE
''''                 sql = "Update Nfcapa set Nf = " & NroNotaFiscal & ", Serie = '" & PegaSerieNota _
''''                        & "' where NumeroPed =  " & txtPedido.text
                        
                 sql = "Update Nfcapa set Nf = " & NroNotaFiscal & ", Serie = '" & PegaSerieNota _
                        & "' where NumeroPed =  " & txtPedido.text
             
                 rdoCNLoja.Execute sql
                 Screen.MousePointer = vbNormal
                 rdoCNLoja.CommitTrans
    
                 rdoCNLoja.BeginTrans
                 Screen.MousePointer = vbHourglass
       
                 sql = "Update NfItens set Nf = " & NroNotaFiscal & ", Serie = '" & PegaSerieNota _
                        & "' where NumeroPed =  " & txtPedido.text
                        
                 rdoCNLoja.Execute sql
                 Screen.MousePointer = vbNormal
                 rdoCNLoja.CommitTrans
                GravaMovimentoCaixa
                
                If EncerraVenda(Val(txtPedido.text), " ", 1) = False Then
                   'MsgBox "ICMS inter estadual da referencia " & (RsItensNF("PR_Referencia")) _
                   '        & " n�o encontrado" & Chr(10) & "A nota n�o pode ser impressa", vbCritical, "Aviso"
                   rsComplementoVenda.Close
                   Exit Sub
                End If
                EmiteNotafiscal NroNotaFiscal, txtSerie.text
                    
            End If
           rsComplementoVenda.Close
          
          
  
   '*********************************************
        
       If UCase(frmFormaPagamento.txtIdentificadequeTelaqueveio.text) = "FRMCAIXANF" Then
            frmCaixaNF.txtPedido.text = ""
            frmCaixaNF.grdItens.Rows = 1
            frmCaixaNF.lblTotalvenda.Caption = ""
            frmCaixaNF.lblTotalItens.Caption = ""
            limpaGrid frmCaixaNF.grdItens
       End If
       frmControlaCaixa.cmdTotalVenda.Caption = ""
       frmControlaCaixa.cmdTotalItens.Caption = ""
       frmControlaCaixa.cmdTotalPedidoGE.Caption = ""
    
ElseIf txtTipoNota.text = "Romaneio" Then
    
      'defineImpressora

      Call PegaNumeroRomaneio
      Call ImprimeRomaneio
      Call ImprimeRomaneio
      txtSerie.text = "00"
      Call GravaMovimentoCaixa
      wRomaneio = True
      EncerraVenda Val(txtPedido.text), " ", 1
      frmCaixaRomaneio.txtPedido.text = ""
      frmCaixaRomaneio.grdItens.Rows = 1
      frmCaixaRomaneio.lblTotalvenda.Caption = ""
      frmCaixaRomaneio.lblTotalItens.Caption = ""
     Call GravaValorCarrinho(frmCaixaRomaneio, frmCaixaRomaneio.lblTotalItens.Caption)
      limpaGrid frmCaixaRomaneio.grdItens


ElseIf txtTipoNota.text = "RomaneioDireto" Then
      Call PegaNumeroRomaneio
      Call ImprimeRomaneio
      Call ImprimeRomaneio
      txtSerie.text = "00"
      Call GravaMovimentoCaixa
      wRomaneio = True
      EncerraVenda Val(txtPedido.text), " ", 1
      frmCaixaRomaneioDireto.grdItens.Rows = 1
      frmCaixaRomaneioDireto.lblTotalvenda.Caption = ""
      frmCaixaRomaneioDireto.lblTotalItens.Caption = ""
      Call GravaValorCarrinho(frmCaixaRomaneioDireto, frmCaixaRomaneioDireto.lblTotalItens.Caption)
      limpaGrid frmCaixaRomaneioDireto.grdItens
      
ElseIf txtTipoNota.text = "D1" Or txtTipoNota.text = "S1" Then

          'sql = "Select * from controlecaixa  where CTR_Supervisor <> 99 and" _
          '   & "   between '" & Format(Date, "yyyy/mm/dd") & " 00:00:00' and  '" _
          '   & Format(Date, "yyyy/mm/dd") & " 23:59:59'"
          
          
          
          
          
          
             'PegaLoja.CursorLocation = adUseClient
             'PegaLoja.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
             
             'PegaLoja.Close

       sql = "Update nfcapa " & vbNewLine & _
             "set NF = " & Trim(frmCaixaNotaManual.txtNota) & ", " & vbNewLine & _
             "Serie = '" & txtTipoNota.text & "' " & vbNewLine & _
             "Where NumeroPed = " & frmControlaCaixa.txtPedido.text _
              & " and tiponota = 'PA'"
        rdoCNLoja.Execute sql
        
        sql = "Update NfItens " & vbNewLine & _
              "set Nf = " & Trim(frmCaixaNotaManual.txtNota) & ", " & vbNewLine & _
              "Serie = '" & txtTipoNota.text & "' " & vbNewLine & _
              "where NumeroPed =  " & frmControlaCaixa.txtPedido.text
        rdoCNLoja.Execute sql
        
      txtSerie.text = txtTipoNota.text
      NroNotaFiscal = frmCaixaNotaManual.txtNota.text
      Call GravaMovimentoCaixa
      EncerraVenda Val(txtPedido.text), " ", 1
      frmCaixaNotaManual.grdItens.Rows = 1
      frmCaixaNotaManual.lblTotalvenda.Caption = ""
      frmCaixaNotaManual.lblTotalItens.Caption = ""
      Call GravaValorCarrinho(frmCaixaNotaManual, frmCaixaNotaManual.lblTotalItens.Caption)
      limpaGrid frmCaixaNotaManual.grdItens

End If

'Call ZeraVariaveis

lblTootip.Visible = False
'lblTootip1.Visible = False
  
  pedido = txtPedido.text
  pedido = IIf(txtPedido.text = "", 0, txtPedido.text)
  
  'AQUI 2016 FELIPE
  'CriaNFE 0, pedido
  frmStartaProcessos.txtPedido.text = txtPedido.text

  
  txtValorModalidade.text = ""
  chbvalortroco.Caption = ""
  
 
  If (UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Or _
     UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXASAT" Or _
     UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXASATDIRETO" Or _
     UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO") And _
     txtTipoNota.text <> "NF" Then
     Exit Sub
  End If
  
  If UCase(frmFormaPagamento.txtIdentificadequeTelaqueveio.text) = "FRMCAIXANF" Then
     frmCaixaNF.lblTotalvenda.Caption = ""
     frmCaixaNF.lblTotalItens.Caption = ""
     Call GravaValorCarrinho(frmCaixaNF, frmCaixaNF.lblTotalItens.Caption)
     frmCaixaNF.txtPedido.text = ""
  End If
  
  If UCase(frmFormaPagamento.txtIdentificadequeTelaqueveio.text) = "FRMCAIXAROMANEIO" Then
     frmCaixaRomaneio.lblTotalvenda.Caption = ""
     frmCaixaRomaneio.lblTotalItens.Caption = ""
     Call GravaValorCarrinho(frmCaixaRomaneio, frmCaixaRomaneio.lblTotalItens.Caption)
     frmCaixaRomaneio.txtPedido.text = ""
  End If
  
  
  If UCase(frmFormaPagamento.txtIdentificadequeTelaqueveio.text) = "FRMCAIXAROMANEIODIRETO" Then
     frmCaixaRomaneioDireto.lblTotalvenda.Caption = ""
     frmCaixaRomaneioDireto.lblTotalItens.Caption = ""
     Call GravaValorCarrinho(frmCaixaRomaneioDireto, frmCaixaRomaneioDireto.lblTotalItens.Caption)
 '    frmCaixaRomaneioDireto.txtPedido.Text = ""
  End If
   
    fraRecebimento.Visible = False
    lblTotalPedido.Visible = False
    lblValorTotalPedido.Visible = False
    lblTootip.text = ""
'   lblTootip1.Text = ""
    chbOkPag.Enabled = False
  
    Unload Me
    Unload frmCaixaTEF
    Unload frmCaixaTEFPedido
    Unload frmCaixaNF
    Unload frmCaixaRomaneio
    Unload frmCaixaRomaneioDireto
    Unload frmCaixaNotaManual
    'Unload frmPortal
    frmStartaProcessos.Show vbModal
'   frmStartaProcessos.ZOrder
 
  End Sub
  
Private Sub GravaMovimentoCaixa()

    Wecf = GLB_ECF

    If txtTipoNota.text = "Romaneio" Or txtTipoNota.text = "RomaneioDireto" Then
        sql = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Subgrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
        & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa,MC_Pedido,MC_DataProcesso,MC_TipoNota,MC_SequenciaTEF) values(" & Wecf & ",'" & GLB_USU_Codigo & "','" & Trim(wlblloja) & "', " _
        & " '" & Format(GLB_DataInicial, "YYYY/MM/DD") & "', " & 20105 & ",''," & NroNotaFiscal & ",'" & txtSerie.text & "', " _
        & "" & ConverteVirgula(Format(frmControlaCaixa.cmdTotalVenda.Caption, "##,##0.00")) & ", " _
        & "0,'" & Agencia & "',0,0," & wParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & GLB_Caixa & ",'" & txtPedido.text & "','" & Format(Date, "yyyy/mm/dd") & "','V'," & txtNumeroTEF.text & ")"
        rdoCNLoja.Execute (sql)
    End If

    If AvistaReceber <> 0 Then
        sql = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Subgrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
        & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio,MC_ControleAVR, MC_Protocolo,MC_Nrocaixa,MC_Pedido,MC_DataProcesso,MC_TipoNota,MC_SequenciaTEF) values(" & Wecf & ",'" & GLB_USU_Codigo & "','" & Trim(wlblloja) & "', " _
        & " '" & Format(GLB_DataInicial, "YYYY/MM/DD") & "', " & 10204 & ",'', " & NroNotaFiscal & ",'" & txtSerie.text & "', " _
        & "" & ConverteVirgula(Format(AvistaReceber, "##,##0.00")) & ", " _
        & "0,'" & Agencia & "',0,0," & wParcelas & ", " & 9 & ",'A','A'," & GLB_CTR_Protocolo & "," & GLB_Caixa & ",'" & txtPedido.text & "','" & Format(Date, "yyyy/mm/dd") & "','V'," & txtNumeroTEF.text & ")"
        rdoCNLoja.Execute (sql)
    End If

'    If WCodigoModalidadeVISA = "0301" Then
'        If ValCartaoVisa > 0 Then
'            'VISA
'            wGrupoMovimento = "10301"
'            Agencia = bandeiraCartaoVisa
'            wSubGrupo = ""
'            wValorMovimento = Format(ValCartaoVisa, "##,##0.00")
'
'                If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Or _
'                    UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" Then
'                    Retorno = Bematech_FI_EfetuaFormaPagamento("VISA", ValCartaoVisa * 100)
'                    Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
'
'                    If Retorno <> 1 Then
'                        MsgBox "Por favor verificar se impressora est� ligada corretamente!"
'                        Exit Sub
'                    End If
'                End If
'
'            Call GravaRegistro
'        End If
'    End If

'    If wCodigoModalidadeMASTERCARD = "0302" Then
'        If ValCartaoMastercard > 0 Then
'            'MASTERCARD
'            wGrupoMovimento = "10302"
'            Agencia = bandeiraCartaoMastercard
'            wSubGrupo = ""
'            wValorMovimento = Format(ValCartaoMastercard, "##,##0.00")
'            If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Or _
'                UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" Then
'                Retorno = Bematech_FI_EfetuaFormaPagamento("MASTERCARD", ValCartaoMastercard * 100)
'                Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
'            End If
'            Call GravaRegistro
'        End If
'    End If


'    If WCodigoModalidadeAMEX = "0303" Then
'        If ValCartaoAmex > 0 Then
'            'AMEX
'            wGrupoMovimento = "10303"
'            Agencia = bandeiraCartaoAmex
'            wSubGrupo = ""
'            wValorMovimento = Format(ValCartaoAmex, "##,##0.00")
'
'            If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Or _
'            UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" Then
'                Retorno = Bematech_FI_EfetuaFormaPagamento("AMEX", ValCartaoAmex * 100)
'                Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
'            End If
'
'            Call GravaRegistro
'        End If
'    End If




'    If ValTEFVisaElectron > 0 Then
'        'TEF
'        wGrupoMovimento = "10206"
'        wSubGrupo = "Visa Elec."
'        Agencia = bandeiraTEFVisaElectron
'        wValorMovimento = Format(ValTEFVisaElectron, "##,###0.00")
'
'        If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Or _
'            UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" Then
'            Retorno = Bematech_FI_EfetuaFormaPagamento("TEF", ValTEFVisaElectron * 100)
'            Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
'        End If
'
'        Call GravaRegistro
'    End If

'    If ValTEFRedeShop > 0 Then
'        'TEF
'        wGrupoMovimento = "10203"
'        wSubGrupo = "RedeShop"
'        Agencia = bandeiraTEFRedeShop
'        wValorMovimento = Format(ValTEFRedeShop, "##,###0.00")
'
'        If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Or _
'            UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" Then
'            Retorno = Bematech_FI_EfetuaFormaPagamento("TEF", ValTEFRedeShop * 100)
'            Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
'        End If
'
'
'        Call GravaRegistro
'    End If

'    If ValTEFHiperCard > 0 Then
'        'TEF
'        wGrupoMovimento = "10205"
'        wSubGrupo = "HiperCard"
'        Agencia = bandeiraTEFHiperCard
'        wValorMovimento = Format(ValTEFHiperCard, "##,##0.00")
'
'        If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Or _
'            UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" Then
'            Retorno = Bematech_FI_EfetuaFormaPagamento("TEF", ValTEFHiperCard * 100)
'            Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
'        End If
'
'        Call GravaRegistro
'    End If


'    If wCodigoModalidadeNOTACREDITO = "0701" Then
'        If ValNotaCredito > 0 Then
'            'NOTA DE CREDITO
'            wGrupoMovimento = "10701"
'            wSubGrupo = ""
'            wValorMovimento = Format(ValNotaCredito, "##,##0.00")
'
'            If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Or _
'                UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" Then
'                Retorno = Bematech_FI_EfetuaFormaPagamento("NC", ValNotaCredito * 100)
'                Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
'            End If
'
'            Call GravaRegistro
'        End If
'    End If

'    If WCodigoModalidadeCHEQUE = "0201" Then
'        If ValCheque > 0 Then
'            'CHEQUE
'            wGrupoMovimento = "10201"
'            wSubGrupo = ""
'            wValorMovimento = Format(ValCheque, "##,##0.00")
'            If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Or _
'                UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" Then
'                Retorno = Bematech_FI_EfetuaFormaPagamento("CHEQUE", ValCheque * 100)
'                Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
'            End If
'            Call GravaRegistro
'        End If
'    End If


    If wCodigoModalidadeDINHEIRO = "0101" Then
        If ValDinheiro > 0 Then
            'DINHEIRO
            wGrupoMovimento = "10101"
            wSubGrupo = ""
            wValorMovimento = Format((ValDinheiro - ValTroco), "##,##0.00")
            
            If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Or _
                UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" Then
                retorno = Bematech_FI_EfetuaFormaPagamento("DINHEIRO", ValDinheiro * 100)
                Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
            End If

        Call GravaRegistro
        End If
    End If

    If wCodigoModalidadeFATURADO = "0501" Then
        wGrupoMovimento = "10501"
        wSubGrupo = ""
        wValorMovimento = Format(chbValoraPagarFat.Caption, "##,##0.00")
        Call GravaRegistro
    End If

    If wCodigoModalidadeFINANCIADO = "0601" Then
        wGrupoMovimento = "10601"
        wSubGrupo = ""
        wValorMovimento = Format(chbValoraPagarFat.Caption, "##,##0.00")
        Call GravaRegistro
    End If
    
    wGrupo = 0



    If txtSerie.text Like GLB_SerieCF & "*" Then
        wGrupo = 20101
    ElseIf txtSerie.text = PegaSerieNota Then
        wGrupo = 20102
    ElseIf txtSerie.text = "SF" Then
        wGrupo = 20103
    ElseIf txtSerie.text = "SM" Then
        wGrupo = 20104
    ElseIf txtSerie.text = "00" Then
        wGrupo = 20105
    ElseIf txtSerie.text = "C0" Then
        wGrupo = 20106
    ElseIf txtSerie.text = "D1" Then
        wGrupo = 20107
    ElseIf txtSerie.text = "S1" Then
        wGrupo = 20108
    End If


    If wGrupo <> 0 Then
        If txtTipoNota.text = "CUPOM" Then
            sql = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Subgrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
            & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa,MC_Pedido,MC_DataProcesso,MC_TipoNota,MC_SequenciaTEF) values(" & Wecf & ",'" & GLB_USU_Codigo & "','" & Trim(wlblloja) & "', " _
            & " '" & Format(GLB_DataInicial, "yyyy/mm/dd") & "', " & wGrupo & ",'" & wSubGrupo & "', " & NroNotaFiscal & ",'" & txtSerie.text & "', " _
            & "" & ConverteVirgula(Format(frmControlaCaixa.cmdTotalVenda.Caption, "##,##0.00")) & ", " _
            & "0,'" & Agencia & "',0,0,0,0,'A'," & GLB_CTR_Protocolo & "," & GLB_Caixa & ",'" & txtPedido.text & "','" & Format(GLB_DataInicial, "yyyy/mm/dd") & "','PA','" & txtNumeroTEF.text & "')"
            rdoCNLoja.Execute (sql)

        ElseIf txtSerie.text <> "00" Then

            wTotalNota = frmControlaCaixa.cmdTotalVenda.Caption

            sql = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Subgrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
            & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa,MC_Pedido,MC_DataProcesso,MC_TipoNota,MC_SequenciaTEF) values(" & Wecf & ",'" & GLB_USU_Codigo & "','" & Trim(wlblloja) & "', " _
            & " '" & Format(GLB_DataInicial, "yyyy/mm/dd") & "', " & wGrupo & ",'" & wSubGrupo & "', " & NroNotaFiscal & ",'" & txtSerie.text & "', " _
            & "" & ConverteVirgula(Format(wTotalNota, "##,##0.00")) & ", " _
            & "0,'" & "" & "',0,0,0,9,'A'," & GLB_CTR_Protocolo & "," & GLB_Caixa & ",'" & txtPedido.text & "','" & Format(GLB_DataInicial, "yyyy/mm/dd") & "','PA','" & txtNumeroTEF.text & "')"
            rdoCNLoja.Execute (sql)
            
        End If
    End If


    'Garantia estendida
    If wValorGE > 0 Then
        sql = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_SubGrupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
        & "MC_Contacorrente,MC_bomPara,MC_Parcelas,MC_Remessa,MC_SituacaoEnvio, MC_Protocolo,MC_Nrocaixa,MC_Pedido, MC_DataProcesso,MC_TipoNota,MC_SequenciaTEF) values(" & Wecf & ",'" & GLB_USU_Codigo & "','" & Trim(wlblloja) & "', " _
        & " '" & Format(GLB_DataInicial, "yyyy/mm/dd") & "', " & 11009 & ",'" & wSubGrupo & "', " & NroNotaFiscal & ",'" & txtSerie.text & "', " _
        & "" & ConverteVirgula(Format(wValorGE, "##,###0.00")) & ", " _
        & "0,'" & "" & "',0,0," & wParcelas & ", " & 9 & ",'A'," & GLB_CTR_Protocolo & "," & GLB_Caixa & ",'" & txtPedido.text & "','" & Format(GLB_DataInicial, "yyyy/mm/dd") & "','PA','" & txtNumeroTEF.text & "')"
        rdoCNLoja.Execute (sql)
    End If



End Sub

Private Function ProcuraPedido()
   
   Screen.MousePointer = 11
   Dim vSQL As String
   Dim Linha As Long
   Dim i As Integer
   Dim wTootip As Double
   Dim Tootip1 As Double
      
        ConsistePedido Val(txtPedido)
        
 If RsDados.State = 1 Then
  RsDados.Close
End If

         
sql = "SELECT DISTINCT NFCapa.NumeroPed, NFCapa.totalNota, NFCapa.pgentra, NFCapa.CondPag, NFCapa.vlrMercadoria," & _
      "parcelas, modalidadeVenda, CondicaoPagamento.CP_intervaloparcelas, fin_Cliente.CE_CGC, Fin_Cliente.CE_TipoPessoa, " & _
      "Fin_Cliente.ce_Razao, " & _
      "CondicaoPagamento.CP_tipo From nfcapa, Produtoloja, CondicaoPagamento, fin_Cliente " & _
      "Where fin_Cliente.ce_codigoCliente = nfcapa.Cliente And CondicaoPagamento.cp_codigo = nfcapa.condpag " & _
      "and nfcapa.numeroped= " & txtPedido.text & " and nfcapa.tiponota= 'PA'"
        
        RsDados.CursorLocation = adUseClient
        RsDados.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

        If Not RsDados.EOF Then
         
           txtPedido.text = Trim(RsDados("NumeroPED"))
          
           lblTotalPedido.Visible = True
           lblValorTotalPedido.Visible = True
           lblValorTotalPedido.text = Format(CDbl(RsDados("TotalNota")), "##,##0.00")
         
           wParcelas = RsDados("parcelas")
'           txtParcelas.Text = RsDados("cp_parcelas")
'?'           wIndicePreco = RsDados("Indicepreco")
           If wParcelas > 1 Then
              'wvalorparcelas = (RsDados("TotalNota") - RsDados("pgentra")) / WParcelas
              'lblParcelas.Caption = WParcelas & " x " & Format(wvalorparcelas, "##,##0.00")
              lblParcelas.Caption = "Forma Pagamento: " & RTrim(RsDados("modalidadeVenda")) & " (" & wParcelas & " parcelas)"
           Else
              lblParcelas.Caption = "Forma Pagamento: " & RTrim(RsDados("modalidadeVenda"))
           End If

           If Trim(RsDados("cp_tipo")) = "FI" Then 'Financiado
              lblTootip.text = " ATEN��O: Valor do contrato R$   " & Format(((lblValorTotalPedido.text - RsDados("pgentra")) * wIndicePreco), "##,###,##0.00")
 '             lblTootip1.Text = WParcelas & "  Parcela(s)  de  R$   " & Format(((lblValorTotalPedido.Text - RsDados("pgentra")) * wIndicePreco) / WParcelas, "##,###,##0.00")
           Else
              lblTootip.text = ""
   '           lblTootip1.Text = ""
           End If
             
           
           If Trim(RsDados("cp_tipo")) = "FA" Then
              Faturada = True
              Financiada = False
              wVerificaAVR = False
              ValorFaturada = Format(RsDados("VlrMercadoria"), "0.00")
              wTotalNotaFatFin = Format(CDbl(RsDados("TotalNota")), "##,##0.00") - Format(CDbl(RsDados("pgentra")), "##,##0.00")
              wTotalNota = Format(CDbl(RsDados("TotalNota")), "##,##0.00")
           ElseIf Trim(RsDados("cp_tipo")) = "FI" Then
              Financiada = True
              Faturada = False
              wVerificaAVR = False
              ValorFinanciada = Format(RsDados("VlrMercadoria"), "0.00") - Format(IIf(IsNull(RsDados("pgentra")), 0, RsDados("pgentra")), "0.00")
              wTotalNotaFatFin = Format(CDbl(RsDados("TotalNota")), "##,##0.00") - Format(CDbl(RsDados("pgentra")), "##,##0.00")
              wTotalNota = Format(CDbl(RsDados("TotalNota")), "##,##0.00")
           ElseIf Trim(RsDados("cp_tipo")) = "AV" Or Trim(RsDados("cp_tipo")) = "CC" Then
              If RsDados("condpag") = 2 Then
                  wVerificaAVR = True
                  Faturada = False
                  Financiada = False
                  AvistaReceber = Format(RsDados("totalnota"), "0.00")
                  wTotalNota = Format(CDbl(RsDados("TotalNota")), "##,##0.00")
              Else
                  Faturada = False
                  Financiada = False
                  wVerificaAVR = False
                  wTotalNota = Format(CDbl(RsDados("TotalNota")), "##,##0.00")
              End If
          End If

           If RsDados("ce_cgc") <> "" Then
              wDocumento = Trim(RsDados("ce_cgc"))
              If RsDados("ce_Tipopessoa") = "F" Or RsDados("ce_Tipopessoa") = "U" Then
                wPessoa = 2
              Else
                wPessoa = 1
              End If
           End If
        
           If wVerificaAVR = True Then
              lblApagar.text = Format(CDbl(RsDados("TotalNota")), "##,###0.00")
           End If
           
           'txtCliente.Text = RsDados("ce_razao")
           
           txtPedido.Enabled = False
'           RsDados.Close
        Else
           MsgBox "N�mero de pedido inexistente.", vbInformation, "Informa��o"
        
        Unload Me
        End If
        RsDados.Close
  
   Screen.MousePointer = 0
   
End Function


Private Sub chbOK_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Exit Sub
End If

 If KeyAscii = 27 Then
    If Trim(txtSerie.text) Like GLB_SerieCF & "*" Then
       If chbValorPago.Caption > 0 Then
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

Private Sub chbOutros_Click()
frmcond.Visible = False
End Sub


Private Sub chbValoraPagar_KeyDown(KeyCode As Integer, Shift As Integer)
    chamaAtalho KeyCode
End Sub

Private Sub chamaAtalho(KeyCode As Integer)
    
    If KeyCode = 98 Then
        If cmdTefCredito.Enabled And cmdTefCredito.Visible Then
            cmdTefCredito_Click
        End If
    End If
    
    If KeyCode = 97 Then
        If cmdTefDebito.Enabled And cmdTefDebito.Visible Then
            cmdTefDebito_Click
        End If
    End If
    
End Sub

Private Sub chbValoraPagar_KeyUp(KeyCode As Integer, Shift As Integer)
    ativaAtalho = False
End Sub

Private Sub chbValorFalta_KeyDown(KeyCode As Integer, Shift As Integer)
    chamaAtalho KeyCode
End Sub

Private Sub chbValorFalta_KeyUp(KeyCode As Integer, Shift As Integer)
    ativaAtalho = False
End Sub

Private Sub chbValorPago_KeyDown(KeyCode As Integer, Shift As Integer)
    chamaAtalho KeyCode
End Sub

Private Sub chbValorPago_KeyUp(KeyCode As Integer, Shift As Integer)
ativaAtalho = False
End Sub

Private Sub chbVisa_Click()
    carregaCodigoModalidade "VISA"
    lblParcelas.Visible = True
    txtValorModalidade.Enabled = True
    txtValorModalidade.SetFocus
End Sub


Private Sub chbVisa_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    lblModalidade.Caption = ""
   ' FraParcelas.Visible = False
   ' lblParc.Visible = False
   lblParcelas.Visible = False
End If
End Sub

Private Sub chbVisaElectron_Click()
    carregaCodigoModalidade "VISA ELEC."
    frmcond.Visible = False
    lblParcelas.Visible = False
    txtValorModalidade.Enabled = True
    txtValorModalidade.SetFocus
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

Function carregarFormaPagamentoAnteriorTEF() As Double

    Dim sql As String
    Dim rsConsulta As New ADODB.Recordset
    
    sql = "select sum(mc_valor) as totalvalor " & vbNewLine & _
    "from movimentocaixa " & vbNewLine & _
    "where  mc_serie = '" & txtSerie.text & "' and " & vbNewLine & _
    "mc_protocolo = " & GLB_CTR_Protocolo & " and " & vbNewLine & _
    "mc_nrocaixa = '" & GLB_Caixa & "' " & vbNewLine & _
    "and mc_pedido = '" & txtPedido.text & "'" & vbNewLine & _
    ""
    
    rsConsulta.CursorLocation = adUseClient
    rsConsulta.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not rsConsulta.EOF Then
        
        If Not IsNull(rsConsulta("totalvalor")) Then
            carregarFormaPagamentoAnteriorTEF = rsConsulta("totalvalor")
        End If
        
        rsConsulta.MoveNext
        
    End If
    
    rsConsulta.Close
   
End Function

Private Sub FormaPagamento()

'  left = 9100
 ' top = 2750
 wCodigoModalidadeDINHEIRO = ""
 WCodigoModalidadeAMEX = ""
 WCodigoModalidadeCHEQUE = ""
 wCodigoModalidadeBNDES = ""
 wCodigoModalidadeMASTERCARD = ""
 wCodigoModalidadeNOTACREDITO = ""
 wTEFVisaElectron = ""
 wTEFRedeShop = ""
 wTEFHiperCard = ""
 
 WCodigoModalidadeVISA = ""
'  frmcond.Left = 330
'  frmcond.Top = 1305
'  frmcond.Height = 6570
'  frmcond.Width = 5100
  fraNModalidades.Visible = True
  txtValorModalidade.Visible = True
  lblModalidade.Visible = True



  txtValorModalidade.Enabled = False
  chbValorPago.Caption = Format(0, "0.00")
  chbValorPago.Caption = Format(carregarFormaPagamentoAnteriorTEF, "##,###0.00")
  chbValorFalta.Caption = Format(wValoraPagarNORMAL + wtotalGarantia - chbValorPago.Caption, "##,###0.00")
  chbValoraPagar.Caption = Format(wValoraPagarNORMAL + wtotalGarantia, "##,###0.00")
  valValoraPagar = wValoraPagarNORMAL + wtotalGarantia
  TotPago = chbValorPago.Caption

If frmFormaPagamento.txtSerie.text = "00" Then
     txtPedido.text = pedido
     ProcuraPedido
     VerificaTipoModalidade
     GoTo Continua
 End If
 

If frmFormaPagamento.txtSerie.text = PegaSerieNota Then
   txtPedido.text = pedido
   ProcuraPedido
   VerificaTipoModalidade
   GoTo Continua
End If
  
If frmFormaPagamento.txtSerie.text = "NE" Then
   txtPedido.text = pedido
   ProcuraPedido
   VerificaTipoModalidade
     
End If

Continua:


    frmcond.Visible = True
    chbTroco.Visible = False
    frmcond.Visible = False
  
End Sub

Private Sub cmdRetornaOperacaoTEF_Click()
    'If MsgBox("A opera��o com o TEF ser� cancelada. Deseja Continuar?", vbQuestion + vbYesNo, "Aten��o") = vbYes Then
        cancelarOperacaoTEF = True
    'End If
End Sub

Private Sub cmdTefCredito_Click()

  lblModalidade.Caption = "CREDITO"

  'lblParcelas.Visible = True
  txtValorModalidade.Enabled = True
  txtValorModalidade.SetFocus
  'CodigoModalidade = "0302"
  'wCodigoModalidadeMASTERCARD = "0302"
  cFormaPGTO = "Cartao"

    '3 Cr�dito

End Sub

Private Sub cmdTefDebito_Click()

  carregaCodigoModalidade "VISA ELEC."
  lblModalidade.Caption = "DEBITO"

  'lblParcelas.Visible = True
  txtValorModalidade.Enabled = True
  txtValorModalidade.SetFocus
  'CodigoModalidade = "0302"
  'wCodigoModalidadeMASTERCARD = "0302"
  cFormaPGTO = "Cartao"

  '2 D�bito

End Sub

Private Sub habilitaFrameTEFOperacoes(ativa As Boolean)

    retornaOperacaoTEF = False
    cancelarOperacaoTEF = False
    
    framePagamentoTEFInterno.Enabled = ativa
    
    If Not ativa Then
        framePagamentoTEF.Height = 2145
    End If
    If ativa Then
        framePagamentoTEF.Height = 1500
    End If
    
End Sub

Private Sub cmdTefDebito_KeyDown(KeyCode As Integer, Shift As Integer)
    timeHabilitaTEF.Enabled = True
    tempoHabilitaPOS = 0
End Sub

Private Sub cmdTefDebito_KeyUp(KeyCode As Integer, Shift As Integer)
    timeHabilitaTEF.Enabled = False
    tempoHabilitaPOS = 0
End Sub

Private Sub cmdTefDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    timeHabilitaTEF.Enabled = True
    tempoHabilitaPOS = 0
End Sub

Private Sub cmdTefDebito_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    timeHabilitaTEF.Enabled = False
    tempoHabilitaPOS = 0
End Sub

Private Sub cmdTefOperacao0_Click()


  
  
End Sub

Private Sub cmdTrocar_Click()
    Agencia = ""
    frameCartoes.Visible = False
End Sub



Private Sub Form_Activate()

    If primeiroCarregamento Then

    primeiroCarregamento = False
frameCartoes.top = chbRede.top
frameCartoes.left = chbRede.left
frameCartoes.Visible = False

 Me.top = frmControlaCaixa.webPadraoTamanho.top
 Me.Height = frmControlaCaixa.webPadraoTamanho.Height - 100
 
 fraFinanciadoFaturado.left = fraPagamento.left
 fraFinanciadoFaturado.top = fraPagamento.top
 
 
'frmFormaPagamento.top = 2875
frmFormaPagamento.left = 8880
frmFormaPagamento.Width = 5550
'frmFormaPagamento.Height = 7110
 

chbOkPag.Caption = "OK"
chbOkPag.Height = 570
fraPagamento.Visible = False
fraFinanciadoFaturado.Visible = False
lblParcelasFat.Caption = ""
lblParcelas.Caption = ""
chbOkPag.Visible = False
wValorGE = 0
'chbOkFat.Caption = "Confirma Entrada"
chbvalortroco.top = chbValorFalta.top

If UCase(frmFormaPagamento.txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" _
Or UCase(frmFormaPagamento.txtIdentificadequeTelaqueveio.text) = "FRMCAIXASAT" _
Or UCase(frmFormaPagamento.txtIdentificadequeTelaqueveio.text) = "FRMCAIXASATDireto" Then
   ProcuraPedido
   VerificaTipoModalidade
End If

 If txtIdentificadequeTelaqueveio.text = "FRMCAIXANF" Then
     txtTipoNota.text = "NF"
  End If
  
 If txtIdentificadequeTelaqueveio.text = "FRMCAIXATEF" Then
     frmCaixaTEF.txtCodigoProduto.text = ""
  End If
sql = ""
sql = "Select condpag,pgentra,cp_parcelas,totalnota,cp_tipo,cp_coeficiente,GarantiaEstendida,TotalGarantia from nfcapa,CondicaoPagamento " & _
      "where cp_codigo = condpag and numeroped = " & txtPedido.text
rsComplementoVenda.CursorLocation = adUseClient
rsComplementoVenda.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic

'txtSerie.Text = rsComplementoVenda("serie")

 If txtSerie.text <> "NE" And (Not txtSerie.text Like GLB_SerieCF & "*") Then
       txtSerie.text = PegaSerieNota
 End If
 
 If Trim(rsComplementoVenda("GarantiaEstendida")) = "S" Then
       wValorGE = rsComplementoVenda("TotalGarantia")
 Else
       wValorGE = 0
       wtotalGarantia = 0
 End If


If rsComplementoVenda("cp_tipo") = "FI" Or rsComplementoVenda("cp_tipo") = "FA" Then
    fraFinanciadoFaturado.Visible = True
    chbOkFat.Visible = False
    chbConfimaEntrada.SetFocus
    fraPagamento.left = 135
    fraPagamento.top = 510
    If rsComplementoVenda("cp_tipo") = "FI" Then
        lblFinanciadoFaturado.Caption = "Financiado"
        EntFinanciada = rsComplementoVenda("pgentra")
        'wPagamentoECF = 9
    Else
        lblFinanciadoFaturado.Caption = "Faturado"
        EntFaturada = rsComplementoVenda("pgentra")
        'wPagamentoECF = 1
    End If
    chbValoraPagarFat.Caption = Format((wValoraPagarNORMAL + wtotalGarantia), "##,###0.00")
    If rsComplementoVenda("pgentra") > 0 Then
       chbValorEntrada.Caption = Format((rsComplementoVenda("pgentra")), "##,###0.00")
    End If
    
    wParcelas = rsComplementoVenda("cp_parcelas")
    If wParcelas > 1 Then
      'wvalorparcelas = ((rsComplementoVenda("TotalNota")) - rsComplementoVenda("pgentra")) / WParcelas
      lblParcelasFat.Caption = wParcelas & " Parcelas " '& Format(wvalorparcelas, "##,##0.00")
    End If

    rsComplementoVenda.Close
Else
    fraPagamento.Visible = True
    'chbValoraPagar.SetFocus
    fraFinanciadoFaturado.left = 135
    fraFinanciadoFaturado.top = 510
    
    
    wParcelas = rsComplementoVenda("cp_parcelas")
    If wParcelas > 1 Then
      wvalorparcelas = ((rsComplementoVenda("TotalNota") * rsComplementoVenda("cp_coeficiente")) - rsComplementoVenda("pgentra")) / wParcelas
      lblParcelas.Caption = wParcelas & " X " & Format(wvalorparcelas, "##,##0.00")
    End If
    
    rsComplementoVenda.Close
    Call FormaPagamento
    
    'exibirMensagemPedidoTEF txtPedido.text, WParcelas

End If

    End If
    
    limpaCamposPagamentoTotalConfirmado

End Sub

Private Sub habilitaPagamentoTEF()

    framePagamentoTEF.Visible = False

    If Not GLB_TefHabilidado Then Exit Sub
        
    framePagamentoTEF.Visible = True
    framePagamentoTEF.left = chbRede.left + fraPagamento.left + fraNModalidades.left
    framePagamentoTEF.top = chbRede.top + fraPagamento.top + fraNModalidades.top
    framePagamentoTEF.Height = 1515
    framePagamentoTEF.BackColor = vbBlack
    framePagamentoTEFInterno.BackColor = vbBlack
    lblMensagemTEF.Caption = ""
    If GLB_Administrador Then lblMensagemTEF.Caption = "Click aqui para executar fun��o 0"

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'MsgBox "oi 1"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'MsgBox "oi 2"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'MsgBox "oi 3"
End Sub

Private Sub Form_Load()
    
    primeiroCarregamento = True
    habilitaPagamentoTEF
    
End Sub

Private Sub limpaMovimentoAnteriores()
    'Limpa registros
    sql = "delete movimentocaixa where  mc_serie = '" & txtSerie.text & "' and " _
    & "mc_protocolo = " & GLB_CTR_Protocolo & " and " _
    & "mc_nrocaixa = '" & GLB_Caixa & "' and mc_pedido = '" & txtPedido.text & "'" _
    & "and mc_sequenciaTEF = 0 and mc_tiponota <> 'TF'"
    
    rdoCNLoja.Execute sql
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
    Call ZeraVariaveis
    frmcond.Visible = False
    
    lblTootip.Visible = False
    'lblTootip1.Visible = False
    
    If txtPedido.text <> "" Then
        pedido = txtPedido.text
    End If
    
    pedido = IIf(txtPedido.text = "", 0, txtPedido.text)
    frmStartaProcessos.txtPedido.text = txtPedido.text
    
    chbValoraPagar.Caption = ""
    txtValorModalidade.text = ""
    chbvalortroco.Caption = ""
    
    ' If UCase(frmFormaPagamento.txtIdentificadequeTelaqueveio) = "FRMCAIXATEFPEDIDO" Then
    '    frmCaixaTEFPedido.fraPedido.Visible = True
    ' End If
    
    fraRecebimento.Visible = False
    lblTotalPedido.Visible = False
    lblValorTotalPedido.Visible = False
    lblTootip.text = ""
    '  lblTootip1.Text = ""

End Sub
'*??????????????????

Private Sub chbOkPag_Click()

If txtTipoNota.text = "CUPOM" Then

     txtPedido.text = pedido


''********************

' Inicia o fechamento do cupom

          retorno = Bematech_FI_IniciaFechamentoCupom("D", "$", 0)
          Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal")
          
          If retorno <> 1 Then
            MsgBox "Por favor verificar se impressora est� ligada corretamente!"
            Exit Sub
          End If
''********************

    Call GuardaValoresParaGravarMovimentoCaixa

    txtValorModalidade.text = ""
    txtValorModalidade.Enabled = False
    lblModalidade.Caption = " "

    If wValorModalidadeIncorreto = True Then
       Exit Sub
    End If

    If chbValorFalta.Caption = "" Then
       chbValorFalta.Caption = 0
    End If

    If chbValorFalta.Caption <= 0 Then
       chbOkPag.Visible = True
       chbOkPag.Enabled = True
       chbSair.Enabled = False
       chbOkPag.SetFocus
       fraNModalidades.Visible = False
       txtValorModalidade.Visible = False
       lblModalidade.Visible = False
    End If


    Call FinalizaPagamento

    chbOkPag.Height = 700


       If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Then
         frmCaixaTEF.txtCodigoProduto = ""
         frmCaixaTEF.txtCGC_CPF.text = ""
         limpaGrid frmCaixaTEF.grdItens
         frmCaixaTEF.grdItens.Rows = 1
         wItens = 0
         frmCaixaTEF.lblTotalvenda.Caption = ""
         frmCaixaTEF.lblTotalItens.Caption = ""
         Call GravaValorCarrinho(frmCaixaTEF, frmCaixaTEF.lblTotalItens.Caption)
         txtIdentificadequeTelaqueveio.text = ""
         frmCaixaTEF.cmdTotalVenda.Caption = ""
         frmCaixaTEF.cmdItens.Caption = ""
         frmCaixaTEF.lblDescricaoProduto.Caption = ""
         frmCaixaTEF.fraProduto.Visible = False
         frmCaixaTEF.fraNFP.Visible = True
      End If

      If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" Then
         limpaGrid frmCaixaTEFPedido.grdItens
         frmCaixaTEFPedido.grdItens.Rows = 1
         txtIdentificadequeTelaqueveio.text = ""
         frmCaixaTEFPedido.lblTotalvenda.Caption = ""
         frmCaixaTEFPedido.lblTotalItens.Caption = ""
         Call GravaValorCarrinho(frmCaixaTEFPedido, frmCaixaTEFPedido.lblTotalItens.Caption)
         frmCaixaTEFPedido.fraNFP.Visible = False
         frmCaixaTEFPedido.txtPedido.text = ""
         frmCaixaTEFPedido.txtCGC_CPF.text = ""
         frmCaixaTEFPedido.fraPedido.Visible = True

       End If

        Call ZeraVariaveis
        fraRecebimento.Visible = False
        lblTotalPedido.Visible = False
        lblValorTotalPedido.Visible = False
        lblTootip.text = ""
        chbOkPag.Enabled = False

        If txtTipoNota.text = "PA" Then
             Unload Me
             Unload frmCaixaTEF
             Unload frmCaixaTEFPedido
             Unload frmCaixaNF
             Unload frmCaixaRomaneio
             Unload frmCaixaRomaneioDireto
             Unload frmCaixaNotaManual
            Exit Sub
        End If

        Unload Me
        Unload frmCaixaTEF
        Unload frmCaixaSAT
        Unload frmCaixaSATDireto
        Unload frmCaixaTEFPedido
        Unload frmCaixaNF
        Unload frmCaixaRomaneio
        Unload frmCaixaRomaneioDireto
        Unload frmCaixaNotaManual

        
        frmStartaProcessos.Show vbModal

ElseIf txtTipoNota.text = "SAT" Then

    txtPedido.text = pedido
    
    Call GuardaValoresParaGravarMovimentoCaixa
    
    txtValorModalidade.text = ""
    txtValorModalidade.Enabled = False
    lblModalidade.Caption = " "
    
    If wValorModalidadeIncorreto = True Then
        Exit Sub
    End If
    
    If chbValorFalta.Caption = "" Then
        chbValorFalta.Caption = 0
    End If
    
    If chbValorFalta.Caption <= 0 Then
        chbOkPag.Visible = True
        chbOkPag.Enabled = True
        chbSair.Enabled = False
        chbOkPag.SetFocus
        fraNModalidades.Visible = False
        txtValorModalidade.Visible = False
        lblModalidade.Visible = False
    End If
    
    Call FinalizaPagamento
    
    chbOkPag.Height = 700
    
    If UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXASATDIRETO" Then
    
        limpaGrid frmCaixaSATDireto.grdItens
        frmCaixaSATDireto.grdItens.Rows = 1
        txtIdentificadequeTelaqueveio.text = ""
        frmCaixaSATDireto.lblTotalvenda.Caption = ""
        frmCaixaSATDireto.lblTotalItens.Caption = ""
        Call GravaValorCarrinho(frmCaixaSATDireto, frmCaixaSATDireto.lblTotalItens.Caption)
        frmCaixaSATDireto.fraNFP.Visible = False
        frmCaixaSATDireto.txtCGC_CPF.text = ""
    
    ElseIf UCase(txtIdentificadequeTelaqueveio.text) = "FRMCAIXASAT" Then
    
        limpaGrid frmCaixaSAT.grdItens
        frmCaixaSAT.grdItens.Rows = 1
        txtIdentificadequeTelaqueveio.text = ""
        frmCaixaSAT.lblTotalvenda.Caption = ""
        frmCaixaSAT.lblTotalItens.Caption = ""
        Call GravaValorCarrinho(frmCaixaSAT, frmCaixaSAT.lblTotalItens.Caption)
        frmCaixaSAT.fraNFP.Visible = False
        frmCaixaSAT.txtPedido.text = ""
        frmCaixaSAT.txtCGC_CPF.text = ""
        frmCaixaSAT.fraPedido.Visible = True
    
    End If

    
    Call ZeraVariaveis
    fraRecebimento.Visible = False
    lblTotalPedido.Visible = False
    lblValorTotalPedido.Visible = False
    lblTootip.text = ""
    chbOkPag.Enabled = False
    
    Unload Me
    Unload frmCaixaSAT
    Unload frmCaixaSATDireto

    frmStartaProcessos.Show vbModal


 Else
    txtPedido.text = pedido
    Call GuardaValoresParaGravarMovimentoCaixa
    txtValorModalidade.text = ""
    txtValorModalidade.Enabled = False
    lblModalidade.Caption = " "

    If wValorModalidadeIncorreto = True Then
       Exit Sub
    End If
    
    If chbValorFalta.Caption = "" Then
       chbValorFalta.Caption = 0
    End If
       
    If chbValorFalta.Caption <= 0 Then
       chbOkPag.Visible = True
       chbOkPag.Enabled = True
       chbSair.Enabled = False
       chbOkPag.SetFocus
       fraNModalidades.Visible = False
       txtValorModalidade.Visible = False
       lblModalidade.Visible = False
    End If
    
    Call FinalizaPagamento
    txtPedido.text = pedido
    Call verificaGarantiaEstendida(txtPedido.text)
    txtPedido.text = Empty
    'Call verificaGarantiaEstendida(txtPedido.Text)
   Call ZeraVariaveis
 End If

End Sub

Private Sub chbOkPag_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Exit Sub
End If

 If KeyAscii = 27 Then
    If Trim(txtSerie.text) Like GLB_SerieCF & "*" Then
       If chbValorPago.Caption > 0 Then
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



Private Sub lblMensagemTEF_Click()

    If GLB_Administrador Then
      carregaCodigoModalidade "VISA ELEC."
      lblModalidade.Caption = "Opera��o 0"
    
      'lblParcelas.Visible = True
      txtValorModalidade.Enabled = True
      txtValorModalidade.SetFocus
      'CodigoModalidade = "0302"
      'wCodigoModalidadeMASTERCARD = "0302"
      cFormaPGTO = "Cartao"
  End If
End Sub

Private Sub timeHabilitaTEF_Timer()
    If tempoHabilitaPOS >= 3 Then
        MsgBox "POS Habilitado com sucesso!" & vbNewLine & _
               "Aten��o! Os administradores ser�o alertado sobre essa opera��o", vbInformation
        timeHabilitaTEF.Enabled = False
        framePagamentoTEF.Visible = False
    Else
        tempoHabilitaPOS = tempoHabilitaPOS + 1
    End If
End Sub

Private Sub txtValorModalidade_GotFocus()
   txtValorModalidade.text = ""
   txtValorModalidade.SelStart = 0
   txtValorModalidade.SelLength = Len(txtValorModalidade.text)
End Sub

Private Sub txtValorModalidade_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
    lblModalidade.Caption = " "
'    chbvalortroco.Visible = False
    txtValorModalidade.Enabled = False
    txtValorModalidade.text = ""
    chbDinheiro.SetFocus
    Exit Sub
 End If
    
 VerteclaVirgula txtValorModalidade, KeyAscii
 
 If KeyAscii = 13 Then
 
    cValorPago = txtValorModalidade.text
    
    If txtValorModalidade.text = "" Then
       txtValorModalidade.SelStart = 0
       txtValorModalidade.SelLength = Len(txtValorModalidade.text)
       txtValorModalidade.SetFocus
       Exit Sub
    Else
       txtValorModalidade.text = Format(txtValorModalidade.text, "###,###,##0.00")
    End If
    txtPedido.text = pedido
    
   
    ' ROTINA ECF(NAO APAGAR)
   

'    If chbSair.Visible = True Then
'       chbSair.Visible = False
'       If UCase(txtIdentificadequeTelaqueveio.Text) = "FRMCAIXATEF" Or _
'          UCase(txtIdentificadequeTelaqueveio.Text) = "FRMCAIXAPEDIDO" Then
'' Inicia o fechamento do cupom
'          Retorno = Bematech_FI_IniciaFechamentoCupom("D", "$", 0)
'          Call VerificaRetornoImpressora("", "", "Emiss�o de Cupom Fiscal") '
'       End If
'    End If

    

    Call GuardaValoresParaGravarMovimentoCaixa
    
    chbDinheiro.SetFocus
    txtValorModalidade.text = ""
    txtValorModalidade.Enabled = False
    lblModalidade.Caption = " "
  
  '  FraParcelas.Visible = False
  '  lblParcelas.Visible = False
    
    'A pergunta abaixo � feita para que se o valor do troco for maior que o valor em dinheiro
    ' ou o valor do cartao > que o valor da nota, saia da rotina sem sumir o franModalidade.
    
    If wValorModalidadeIncorreto = True Then
       Exit Sub
    End If
    
    If chbValorFalta.Caption = "" Then
       chbValorFalta.Caption = 0
    End If
       
    limpaCamposPagamentoTotalConfirmado

End If

End Sub

Private Sub limpaCamposPagamentoTotalConfirmado()

    If chbValorFalta.Caption <= 0 Then
    
       chbOkPag.Visible = True
       chbOkPag.Enabled = True
       chbSair.Enabled = False
       framePagamentoTEF.Visible = False
       'chbOkPag.SetFocus
       fraNModalidades.Visible = False
       txtValorModalidade.Visible = False
       lblModalidade.Visible = False
       lblParcelas.Visible = False
       exibirMensagemTEF "Pagamento Confir" & vbNewLine & "   Emitindo NF"
       
       If operacaoTEFCompleta Then atualizaTipoNotaMovimentoCaixa txtPedido.text
       ImprimeComprovanteTEF txtPedido.text
       finalizarTransacaoTEF txtPedido.text, txtSerie.text, True
       
       If GLB_TEFnaoCancelado Then
            MsgBox "Transa��o TEF efetuada. Favor reimprimir �ltimo cupom. " & _
                   "Caso Cielo utilizar apenas 6 �ltimos d�gitos. NSU:  *Numero do NSU* ", _
                   vbInformation, "TEF"
       End If
       
    End If
End Sub

Private Sub txtValorModalidade_LostFocus()
  txtValorModalidade.text = ""
End Sub



Private Sub VerificaTipoModalidade()
      
      lblTootip.Visible = True
  
             If Faturada = True Then
                If EntFaturada <> "0.00" Then
                   fraRecebimento.BackColor = &HC00000
                   
                   lblEntrada.top = 720
                   lblEntrada.Visible = True
                   lblEntrada.text = "ENT.FAT.        R$ "
                   chbValorPago.Caption = Format(EntFaturada, "0.00")
                   fraRecebimento.Visible = True
                  
                Else
                   lblFatFin.top = 720
                   lblFatFin.Visible = True
                   lblFatFin.text = "FATURADA     R$ "
                   lblValorFatFin.top = lblFatFin.top
                   lblValorFatFin.text = Format(ValorFaturada, "0.00")
                   fraRecebimento.Visible = True
                   lblModalidade.Caption = " "
                   chbOkFat.Enabled = True
                 
                End If
             ElseIf Financiada = True Then
                If EntFinanciada <> "0.00" Then
                   fraRecebimento.BackColor = &HC00000
                   lblModalidade.BackColor = &HC00000
                   lblEntrada.top = 720
                   lblEntrada.Visible = True
                   lblEntrada.text = "ENT.FIN.         R$ "
                   chbValoraPagar.Caption = Format(Val(EntFinanciada + wtotalGarantia), "0.00")
                   lblApagar.text = Format(Val(EntFinanciada), "0.00")
                   fraRecebimento.Visible = True
                   'fraRecebimento.ZOrder
                  
                Else
                   lblFatFin.top = 720
                   lblFatFin.Visible = True
                   lblFatFin.text = "FINANCIADA   R$ "
                   lblValorFatFin.top = lblFatFin.top
                   lblValorFatFin.text = Format(Val(ValorFinanciada), "0.00")
             
                   fraRecebimento.Visible = True
                   'fraRecebimento.ZOrder
                   lblModalidade.Caption = " "
 '                  chbOkFat.Enabled = True
                  
                End If
             ElseIf wVerificaAVR = True Then
                    lblFatFin.top = 720
                    lblFatFin.Visible = True
                    lblFatFin.text = "A V R       "
                    
                    fraRecebimento.Visible = True
                    lblModalidade.Caption = " "
                    fraRecebimento.Visible = True
                    txtValorModalidade.text = lblApagar.text
                    fraNModalidades.Visible = False
                    txtValorModalidade.Visible = False
                    lblModalidade.Visible = False
                 
             Else
            
                lblEntrada.Visible = False
                fraRecebimento.Visible = True
            If txtTipoNota.text = "Romaneio" Then
               txtValorModalidade.text = lblApagar.text
               txtValorModalidade.text = chbValoraPagar.Caption
             End If
                
       End If

End Sub

Private Function VerteclaVirgula(ByRef Controle As Control, ByRef Tecla As Integer)

'-- * -- Aceita apenas digita��o de n�meros e o sinal de "," -- * -- '
   If Controle.SelStart = 0 And Controle.SelLength = Len(Controle.text) Then
      Controle.text = ""
   End If
    
   
   
   If Tecla <> 13 Then
      If Chr(Tecla) = "," Or Chr(Tecla) = "." Then
         If InStr(Controle.text, ",") <> 0 Or InStr(Controle.text, ".") <> 0 Then
            Tecla = 0
         Else
            Tecla = Asc(",")
         End If
      ElseIf Not IsNumeric(Chr(Tecla)) And Tecla <> vbKeyBack Then
         Tecla = 0
      End If
   End If

End Function

Sub EstornoFormaPagtosCupom1()
      
      wCupomAberto = False
      
      sql = "Select nfitens.*,PR_Descricao,PR_icmpdv,PR_substituicaotributaria " _
          & "From nfitens,Produtoloja  " _
          & "Where PR_referencia = Referencia and NumeroPed = " _
          & pedido & " and Tiponota = 'PA' order by Item"
           RsDadosTef.CursorLocation = adUseClient
           RsDadosTef.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
         
          If Not RsDadosTef.EOF Then
             Do While Not RsDadosTef.EOF
                wCodigoProduto = RsDadosTef("Referencia")
                wDescricao = RsDadosTef("PR_Descricao")
                wQtde = Format(RsDadosTef("QTDE"), "000")
                wPrecoVenda = Format(RsDadosTef("VLUnit"), "###,###,##0.00")
                If RsDadosTef("pr_substituicaotributaria") = "S" Then
                  wAliquota = "FF"
                Else
                   wAliquota = Replace(Format(RsDadosTef("PR_icmpdv"), "00.00"), ",", "")
                
                   If Trim(wAliquota) = "0000" Then
                       wAliquota = "FF"
                   ElseIf Trim(wAliquota) <> "0560" And Trim(wAliquota) <> "0700" And Trim(wAliquota) <> "0880" And _
                       Trim(wAliquota) <> "1200" And Trim(wAliquota) <> "1800" And Trim(wAliquota) <> "2500" Then
                       wAliquota = "1200"
                   End If
                
                End If
                wTotalVenda = _
                (wTotalVenda + Format((wPrecoVenda * wQuantidade), "###,##0.00"))

             RsDadosTef.MoveNext
            Loop
            
          Else
              MsgBox "Pedido N�o Encontrado", vbCritical, "Aviso"
              Exit Sub
          End If
          RsDadosTef.Close
 
          
      If UCase(frmFormaPagamento.txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEFPEDIDO" Then
         frmFormaPagamento.chbValoraPagar.Caption = Format(frmCaixaTEFPedido.lblTotalvenda.Caption, "###,###,##0.00")
         wValoraPagarNORMAL = Format(frmCaixaTEFPedido.lblTotalvenda.Caption, "###,###,##0.00")
      ElseIf UCase(frmFormaPagamento.txtIdentificadequeTelaqueveio.text) = "FRMCAIXATEF" Then
             frmFormaPagamento.chbValoraPagar.Caption = Format((frmCaixaTEF.cmdTotalVenda.Caption), "###,###,##0.00")
             wValoraPagarNORMAL = Format(frmCaixaTEF.cmdTotalVenda.Caption, "###,###,##0.00")
      End If

      frmFormaPagamento.txtSerie = GLB_SerieCF
      frmFormaPagamento.txtPedido = txtPedido
      frmFormaPagamento.txtTipoNota.text = "CUPOM"

End Sub


Public Sub ZeraVariaveis()

EntFaturada = 0
EntFinanciada = 0
ValorPagamentoCartao = 0
ValDinheiro = 0
ValTroco = 0
ValCheque = 0
ValCartaoAmex = 0
ValCartaoBNDES = 0
ValCartaoMastercard = 0
ValCartaoVisa = 0
ValTEFVisaElectron = 0
valValoraPagar = 0
ValTEFRedeShop = 0
ValTEFHiperCard = 0
TotPago = 0
modalidade = 0
'wTEFRedeShop = 0
'wTEFHiperCard = 0
ValNotaCredito = 0
frmFormaPagamento.chbValorPago.Caption = 0
frmFormaPagamento.chbValorPago.Caption = Format(frmFormaPagamento.chbValorPago.Caption, "##,###0.00")
frmFormaPagamento.chbValoraPagar.Caption = Format(frmFormaPagamento.chbValorPago.Caption, "##,###0.00")
frmFormaPagamento.chbValorFalta.Caption = Format(frmFormaPagamento.chbValoraPagar.Caption, "##,###0.00")
frmFormaPagamento.txtValorModalidade.text = ""


 wCodigoModalidadeDINHEIRO = ""
 WCodigoModalidadeAMEX = ""
 WCodigoModalidadeCHEQUE = ""
 wCodigoModalidadeBNDES = ""
 wCodigoModalidadeMASTERCARD = ""
 wCodigoModalidadeNOTACREDITO = ""
 wCodigoModalidadeFINANCIADO = ""
 wCodigoModalidadeFATURADO = ""
 wTEFVisaElectron = ""
 wTEFRedeShop = ""
 wTEFHiperCard = ""
 WCodigoModalidadeVISA = ""
 
End Sub




