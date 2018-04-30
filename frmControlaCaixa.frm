VERSION 5.00
Begin VB.Form frmControlaCaixa 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   10440
   ClientLeft      =   2415
   ClientTop       =   585
   ClientWidth     =   15375
   Icon            =   "frmControlaCaixa.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmControlaCaixa.frx":23FA
   ScaleHeight     =   10440
   ScaleWidth      =   15375
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame frmTrans 
      BackColor       =   &H80000012&
      Height          =   1500
      Left            =   2295
      TabIndex        =   27
      Top             =   8850
      Visible         =   0   'False
      Width           =   6810
      Begin VB.Label lblLojaDestino 
         AutoSize        =   -1  'True
         BackColor       =   &H0081E8FA&
         BackStyle       =   0  'Transparent
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
         Height          =   300
         Left            =   1800
         TabIndex        =   34
         Top             =   960
         Width           =   4890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H0081E8FA&
         BackStyle       =   0  'Transparent
         Caption         =   "Loja Destino:"
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
         Height          =   300
         Left            =   135
         TabIndex        =   33
         Top             =   1100
         Width           =   1605
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nota de Transferencia:"
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
         Height          =   300
         Left            =   135
         TabIndex        =   32
         Top             =   200
         Width           =   2760
      End
      Begin VB.Label lblRetiradoPor 
         AutoSize        =   -1  'True
         BackColor       =   &H0081E8FA&
         BackStyle       =   0  'Transparent
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
         Height          =   300
         Left            =   1830
         TabIndex        =   31
         Top             =   480
         Width           =   4170
      End
      Begin VB.Label lblSolicitadoPor 
         AutoSize        =   -1  'True
         BackColor       =   &H0081E8FA&
         BackStyle       =   0  'Transparent
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
         Height          =   300
         Left            =   1965
         TabIndex        =   30
         Top             =   735
         Width           =   4050
      End
      Begin VB.Label lblNotaFiscal 
         BackColor       =   &H0081E8FA&
         BackStyle       =   0  'Transparent
         Caption         =   "Retirado Por:"
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
         Height          =   270
         Left            =   135
         TabIndex        =   29
         Top             =   500
         Width           =   1695
      End
      Begin VB.Label lblSerie 
         AutoSize        =   -1  'True
         BackColor       =   &H0081E8FA&
         BackStyle       =   0  'Transparent
         Caption         =   "Solicitado Por:"
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
         Height          =   300
         Left            =   135
         TabIndex        =   28
         Top             =   800
         Width           =   1755
      End
   End
   Begin VB.Timer Timer2 
      Left            =   240
      Top             =   4485
   End
   Begin VB.Frame fraPedido 
      BackColor       =   &H00000000&
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
      Height          =   900
      Left            =   210
      TabIndex        =   15
      Top             =   9450
      Width           =   1920
      Begin VB.TextBox txtPedido 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         MaxLength       =   6
         TabIndex        =   0
         Top             =   285
         Width           =   1680
      End
   End
   Begin Balcao2010.chameleonButton cmdProtocolo 
      Height          =   450
      Left            =   12975
      TabIndex        =   14
      Top             =   2085
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   794
      BTYPE           =   11
      TX              =   "Protocolo"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmControlaCaixa.frx":B78B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton cmdLoja 
      Height          =   450
      Left            =   165
      TabIndex        =   13
      Top             =   2085
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   794
      BTYPE           =   11
      TX              =   "Loja"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmControlaCaixa.frx":B7A7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton cmdNroCaixa 
      Height          =   450
      Left            =   1650
      TabIndex        =   12
      Top             =   2085
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   794
      BTYPE           =   11
      TX              =   "NroCaixa"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmControlaCaixa.frx":B7C3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton cmdOperador 
      Height          =   450
      Left            =   4995
      TabIndex        =   11
      Top             =   2085
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   794
      BTYPE           =   11
      TX              =   "Operador"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   255
      BCOLO           =   255
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmControlaCaixa.frx":B7DF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton cmdReRomaneio 
      Height          =   450
      Left            =   4440
      TabIndex        =   10
      Top             =   10770
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   794
      BTYPE           =   11
      TX              =   "Re Romaneio"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmControlaCaixa.frx":B7FB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton cmdReimprimeNF 
      Height          =   450
      Left            =   1995
      TabIndex        =   9
      Top             =   10815
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   794
      BTYPE           =   11
      TX              =   "ReEmiNF"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmControlaCaixa.frx":B817
      PICN            =   "frmControlaCaixa.frx":B833
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton cmdBotoesParte2 
      Height          =   450
      Index           =   4
      Left            =   13695
      TabIndex        =   7
      Top             =   10815
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   794
      BTYPE           =   11
      TX              =   "T. Numerário"
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
      MICON           =   "frmControlaCaixa.frx":BFCD
      PICN            =   "frmControlaCaixa.frx":BFE9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton cmdBotoesParte2 
      Height          =   450
      Index           =   5
      Left            =   14505
      TabIndex        =   6
      ToolTipText     =   "teste"
      Top             =   10815
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   794
      BTYPE           =   11
      TX              =   "Consulta"
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
      MICON           =   "frmControlaCaixa.frx":C5D7
      PICN            =   "frmControlaCaixa.frx":C5F3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   1335
      Top             =   3810
   End
   Begin Balcao2010.chameleonButton cmdBotoesParte2 
      Height          =   450
      Index           =   2
      Left            =   12105
      TabIndex        =   8
      Top             =   10815
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   794
      BTYPE           =   11
      TX              =   "Cancelar"
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
      MICON           =   "frmControlaCaixa.frx":CC67
      PICN            =   "frmControlaCaixa.frx":CC83
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton cmdBotoesParte2 
      Height          =   450
      Index           =   3
      Left            =   12900
      TabIndex        =   16
      Top             =   10815
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   794
      BTYPE           =   11
      TX              =   "Movimento"
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
      MICON           =   "frmControlaCaixa.frx":D2A4
      PICN            =   "frmControlaCaixa.frx":D2C0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton cmdVersao 
      Height          =   555
      Left            =   6795
      TabIndex        =   17
      Top             =   10740
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   979
      BTYPE           =   11
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
      MICON           =   "frmControlaCaixa.frx":D8F0
      PICN            =   "frmControlaCaixa.frx":D90C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton cmdBotoesParte2 
      Height          =   450
      Index           =   1
      Left            =   11295
      TabIndex        =   19
      Top             =   10815
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   794
      BTYPE           =   11
      TX              =   "Portal"
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
      MICON           =   "frmControlaCaixa.frx":F470
      PICN            =   "frmControlaCaixa.frx":F48C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton cmdBotoesParte2 
      Height          =   450
      Index           =   0
      Left            =   10500
      TabIndex        =   20
      Top             =   10815
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   794
      BTYPE           =   11
      TX              =   "Menu ECF"
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
      MICON           =   "frmControlaCaixa.frx":FB5A
      PICN            =   "frmControlaCaixa.frx":FB76
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton cmdBotoesParte2 
      Height          =   450
      Index           =   6
      Left            =   390
      TabIndex        =   21
      ToolTipText     =   "teste"
      Top             =   10815
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   794
      BTYPE           =   11
      TX              =   "Fecha Caixa"
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
      MICON           =   "frmControlaCaixa.frx":102DC
      PICN            =   "frmControlaCaixa.frx":102F8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton cmdBotoesParte2 
      Height          =   450
      Index           =   7
      Left            =   1185
      TabIndex        =   1
      ToolTipText     =   "teste"
      Top             =   10815
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   794
      BTYPE           =   11
      TX              =   "Garantia"
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
      MICON           =   "frmControlaCaixa.frx":1095D
      PICN            =   "frmControlaCaixa.frx":10979
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton cmdBotoesParte2 
      Height          =   450
      Index           =   8
      Left            =   2505
      TabIndex        =   22
      ToolTipText     =   "teste"
      Top             =   10815
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   794
      BTYPE           =   11
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
      MICON           =   "frmControlaCaixa.frx":1115C
      PICN            =   "frmControlaCaixa.frx":11178
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Balcao2010.chameleonButton cmdBotoesParte2 
      Height          =   450
      Index           =   9
      Left            =   2790
      TabIndex        =   23
      ToolTipText     =   "teste"
      Top             =   10815
      Visible         =   0   'False
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   794
      BTYPE           =   11
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
      MICON           =   "frmControlaCaixa.frx":117DD
      PICN            =   "frmControlaCaixa.frx":117F9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblMensagensTEF 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "teste teste teste teste teste teste teste teste teste teste teste teste teste teste teste "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   870
      Left            =   2490
      TabIndex        =   35
      Top             =   9540
      Width           =   10365
   End
   Begin VB.Label cmdTotalPedidoGE 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   9555
      TabIndex        =   26
      Top             =   2235
      Width           =   2565
   End
   Begin VB.Label cmdTotalVenda 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   9555
      TabIndex        =   25
      Top             =   1890
      Width           =   2565
   End
   Begin VB.Label cmdTotalItens 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   9555
      TabIndex        =   24
      Top             =   1620
      Width           =   2565
   End
   Begin VB.Image webPadraoTamanho 
      Height          =   7860
      Left            =   90
      Stretch         =   -1  'True
      Top             =   2610
      Width           =   15180
   End
   Begin VB.Label lblBotao 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Movimento Caixa"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   12480
      TabIndex        =   18
      Top             =   10575
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Image webInternet1 
      Height          =   2475
      Left            =   90
      Stretch         =   -1  'True
      Top             =   105
      Width           =   15180
   End
   Begin VB.Label lblProtocolo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "111"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0058D6F5&
      Height          =   360
      Left            =   10260
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblOperador 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Adilson"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0058D6F5&
      Height          =   330
      Left            =   2610
      TabIndex        =   4
      Top             =   4380
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Label lblNroCaixa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0058D6F5&
      Height          =   360
      Left            =   7830
      TabIndex        =   3
      Top             =   4440
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblloja 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "271"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0058D6F5&
      Height          =   405
      Left            =   12870
      TabIndex        =   2
      Top             =   4440
      Visible         =   0   'False
      WhatsThisHelpID =   825
      Width           =   1425
   End
End
Attribute VB_Name = "frmControlaCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ControlaMenu As Integer
Dim sql As String
Dim X As Integer
Dim Y As Integer
Dim Controle As Integer
Dim wVerificaRomaneio As Boolean
Dim wVerificaNotaManual As Boolean
Dim wVerificaNotaFiscal As Boolean

Public wSequencia As String
Public wTipoNota As String

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim tempoMouseParado As Double

Private Sub chameleonButton3_Click()

End Sub

Private Sub chameleonButton4_Click()
    frmCancelaCFNF.Show 1
    'frmCancelaCFNF.ZOrder
End Sub

Private Sub chameleonButton1_Click()
'Retorno = Bematech_FI_LeituraX()
'Call VerificaRetornoImpressora("", "", "Leitura X")
End Sub

Private Sub cmdBotoesParte2_Click(Index As Integer)
    Select Case Index
    Case 0
        If tipoCupomEmite = "CF" Then
          frmOperacoesECF.Show vbModal
        Else
          MsgBox "Funções ECF desativadas"
        End If
    Case 1
        frmEmissaoNFe.Show vbModal
       ' frmPortal.Show 0
    
    Case 2
         frmCancelaCFNF.Show vbModal
    Case 3
        frmReimpressaoMovimento.Show vbModal
    Case 4
        frmSangria.Show vbModal
    Case 5
        wFechamentoGeral = False
        frmFechaCaixaGeral.Show vbModal
    Case 6
        frmFechaCaixa.Show vbModal
    Case 7
        frmBilheteGarantia.Show 1
    End Select
    txtPedido.SetFocus
End Sub

Private Sub cmdECF_click()
    frmCaixaTEF.Show vbModal
End Sub

Private Sub cmdECFPedido_Click()
frmCaixaTEFPedido.Show vbModal
End Sub

Private Sub popupNomeBotao(nomeBotao As String, posicaoBotaoY)
    Timer2.Enabled = True
    Timer2.Interval = 500
    'lblBotao.Visible = True
    lblBotao.Caption = nomeBotao
    lblBotao.left = posicaoBotaoY - 430
'   lblBotao.left = nomeBotao.
End Sub

Private Sub cmdCancelaNota_MouseOver()
    tempoMouseParado = 0
End Sub

Private Sub cmdBotoesParte2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    popupNomeBotao cmdBotoesParte2(Index).Caption, cmdBotoesParte2(Index).left
End Sub


Private Sub cmdBotoesParte2_MouseOut(Index As Integer)
    Timer2.Enabled = False
    lblBotao.Visible = False
End Sub




Private Sub cmdFechaCaixa_Click()
 
End Sub

Private Sub cmdNotaFiscal_Click()
 frmCaixaNF.Show vbModal
End Sub

Private Sub cmdConsultaCaixa_Click()

End Sub

Private Sub chbFechaCaixa_Click()
 frmFechaCaixa.Show vbModal
 'frmFechaCaixa.ZOrder
End Sub

Private Sub chbConsultaCaixa_Click()

 frmMovimentoCaixa.Show vbModal
 'frmMovimentoCaixa.ZOrder
End Sub

Private Sub cnbSangria_Click()
frmSangria.Show vbModal
'frmSangria.ZOrder
End Sub

Private Sub chameleonButton8_Click()

End Sub

Private Sub cmdNroCaixa_Click()
    funcao110
End Sub

Private Sub funcao110()
    If GLB_Administrador And GLB_TefHabilidado Then
        If MsgBox("Deseja executar o função 110 do TEF?", vbQuestion + vbYesNo, "MODO ADMINISTRADOR TEF") = vbYes Then
            Dim nf As notaFiscalTEF
            PegaNumeroPedido
            
            nf.pedido = pedido
            
            Call EfetuaOperacaoTEF("110", nf, lblMensagensTEF, lblMensagensTEF)
            ImprimeComprovanteTEF nf.pedido
            finalizarTransacaoTEF nf.pedido, nf.serie, False
        End If
    End If
End Sub

Private Sub cmdReimprimeNF_Click()
frmReemissaoNotaFiscal.Show vbModal
'frmReemissaoNotaFiscal.ZOrder
End Sub

Private Sub cmdReRomaneio_Click()
frmEmissaoRomaneio.Show vbModal
'frmEmissaoRomaneio.ZOrder
End Sub

Private Sub cmdRomaneio_Click()
 frmCaixaRomaneio.Show vbModal
 'frmEmissaoRomaneio.ZOrder
End Sub



Private Sub chSair_Click()
Unload Me
End Sub

Private Sub Command1_Click()
frmReimpressaoMovimento.Show vbModal
End Sub




Private Sub cmdSangria_Click()

End Sub

Private Sub cmdVersao_Click()
    MsgBox "Versão " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Activate()
    'txtPedido.Visible = False
    'webAviso.Navigate "https://www.nfe.fazenda.gov.br/portal/disponibilidade.aspx?versao=0.00&tipoConteudo=Skeuqr8PQBY="
    
    Dim statusContingencia As Byte

    
    If wFechamentoGeral = True Then
         frmFechaCaixaGeral.Show vbModal
    End If
    
    statusContingencia = frmContingencia.verificaModoEmissaoAtual
    If (statusContingencia) <> 0 Then
        If statusContingencia = 2 Then
            webPadraoTamanho.Picture = LoadPicture("C:\Sistemas\DMAC Caixa\Imagens\contingencia2")
        Else
            webPadraoTamanho.Picture = LoadPicture("C:\Sistemas\DMAC Caixa\Imagens\contingencia1")
        End If
    End If

End Sub

Private Sub carregaQtdeTEFnaoCancelado()
    
    Dim sql As String
    Dim RsDados As New ADODB.Recordset

    If Not GLB_TefHabilidado Then Exit Sub

    sql = "SELECT count(*) qtdeTEFnaoCancelado FROM MOVIMENTOCAIXA WHERE MC_Sequenciatef > '0' AND MC_TipoNota in ('PA','TF') "
    
    RsDados.CursorLocation = adUseClient
    RsDados.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
        
        If RsDados("qtdeTEFnaoCancelado") > 0 Then
            GLB_TEFnaoCancelado = True
            operacaoTEFCompleta = True
            lblMensagensTEF.Caption = "ATENÇÃO!" & vbNewLine & " Há " & RsDados("qtdeTEFnaoCancelado") & " TEF(s) não cancelado registrado no sistema. Cancele antes de realiszar qualquer operação."
        End If
        
    RsDados.Close
        
End Sub

Private Sub abilitaFuncoesCF()
    
    VerificaSeEmiteCupom
    
    If tipoCupomEmite <> "CF" Then
        cmdBotoesParte2(0).Visible = False
    End If
    
End Sub

Private Sub Form_GotFocus()
    txtPedido.SetFocus
End Sub

Private Sub Form_Load()
        
    defineImpressora
    
    'Call criaIconeBarra(TrayAdd, Me.Hwnd, "DMAC Caixa", imgIconBandeja.Picture)
    
    lblBotao.top = 11230
    emitiNota = False
    
    resolucaoOriginal.Colunas = resolucaoTela.Colunas
    resolucaoOriginal.Linhas = resolucaoTela.Linhas
    Call AlterarResolucao(1024, 768)
    
    webInternet1.Picture = LoadPicture(endIMG("topo1024768hd"))
    'frmControlaCaixa.Picture = LoadPicture("C:\sistemas\DMAC Caixa\imagens\TelaDMAC.jpg")
    cmdVersao.Caption = ""
    ControlaMenu = 0
    
    sql = "Select * from ParametroCaixa where PAR_NroCaixa = " & GLB_Caixa
    
    rdoParametro.CursorLocation = adUseClient
    rdoParametro.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If rdoParametro.EOF Then
        rdoParametro.Close
        MsgBox "Problema com os Parametros avise ao CPD", vbCritical, "Aviso"
        Unload Me
    End If
    
    lblNroCaixa.Caption = GLB_Caixa
    cmdNroCaixa.Caption = "Caixa " & GLB_Caixa
    
    If GLB_TefHabilidado Then cmdNroCaixa.Caption = cmdNroCaixa.Caption & " (TEF)"
        lblloja.Caption = rdoParametro("PAR_Loja")
        cmdLoja.Caption = "Loja " & rdoParametro("PAR_Loja")
        lblOperador.Caption = Trim(GLB_USU_Nome)
        cmdOperador.Caption = "Operador  " & Trim(GLB_USU_Nome)
    
        tipoZero = False
        
    If VerificaSeEmiteCodigoZero = "S" Then
        tipoZero = True
    End If
    
    lblProtocolo.Caption = GLB_CTR_Protocolo
    cmdProtocolo.Caption = "Protocolo  " & GLB_CTR_Protocolo
    rdoParametro.Close
    
    sql = "select CS_SerieCF as serie from ControleSerie where CS_NroCaixa = '" & GLB_Caixa & "'"
    rdoParametro.CursorLocation = adUseClient
    rdoParametro.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
    GLB_SerieCF = rdoParametro("Serie")
    rdoParametro.Close
    
    abilitaFuncoesCF
    
    Call conectarTEF(lblMensagensTEF)
    
    carregaQtdeTEFnaoCancelado
    
    'txtPedido.SetFocus
    
End Sub


Private Sub Image1_Click()

End Sub

Private Sub imgBanner1_Click()

End Sub

Public Sub Form_LostFocus()
Debug.Print "OI"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If GLB_Administrador = True Then
        notificacaoEmail "Sessão encerrada"
        sql = "EXEc SP_Alerta_Modificacao_MovimentoCaixa ''"
        rdoCNRetaguarda.Execute (sql)
    End If
    
        exibirMensagemTEF ""
    
    End
    
End Sub

Private Sub Timer2_Timer()
    tempoMouseParado = tempoMouseParado + 1
    If tempoMouseParado >= 1 Then
        lblBotao.Visible = True
        Timer2.Enabled = False
    End If
End Sub

Private Sub txtPedido_GotFocus()
   frmControlaCaixa.cmdTotalItens.Caption = ""
   frmControlaCaixa.cmdTotalVenda.Caption = ""
   frmControlaCaixa.cmdTotalPedidoGE.Caption = ""
   frmControlaCaixa.txtPedido.text = ""
End Sub

Private Sub txtPedido_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Vendedor As Integer
If KeyCode = 115 And frmTrans.Visible = True Then

If rdoCNRetaguarda = "" Then
ConectaODBCMatriz
End If

     sql = "Select * from  Vendedor where  VE_CodigoVendedor like '4%' and ve_senha is not null and VE_Nome like 'VENDA%' AND VE_Loja='" & GLB_Loja & "'"

                            If rdoTrans.State <> 0 Then rdoTrans.Close
                            rdoTrans.CursorLocation = adUseClient
                            rdoTrans.Open sql, rdoCNRetaguarda, adOpenForwardOnly, adLockPessimistic
            If Not rdoTrans.EOF Then
            Vendedor = rdoTrans("VE_CodigoVendedor")
            End If
                 rdoTrans.Close
        sql = "update nfcapa set  TIPONOTA='PA',VENDEDOR=" & Vendedor & " ,OUTROVEND=" & Vendedor & ",VendedorLojaVenda=" & Vendedor & ",vendedorGarantia=" & Vendedor & " where NUMEROPED= " & txtPedido.text
        rdoCNLoja.Execute sql
        txtPedido.text = ""
        frmTrans.Visible = False
End If

End Sub

Public Sub txtPedido_KeyPress(KeyAscii As Integer)
Dim letra As Integer
    KeyAscii = campoCaixa(KeyAscii)

If KeyAscii = 27 Then



    If txtPedido.text = "" Then
       'Call criaIconeBarra(TrayDelete, Me.Hwnd, "DMAC Caixa", imgIconBandeja.Picture)
       Call AlterarResolucao(resolucaoOriginal.Colunas, resolucaoOriginal.Linhas)
       Unload Me
    Else

       txtPedido.text = "" 'vbKeyF2
       frmTrans.Visible = False
End If


   End If
   
   If KeyAscii = vbKeyReturn Then
        
        
        lblMensagensTEF.Caption = ""
        
        If wPermitirVenda = True Then
          txtPedido.text = UCase(txtPedido.text)
          wVerificaRomaneio = False
          wVerificaNotaManual = False
          wVerificaNotaFiscal = False
          
          If txtPedido.text = "0" Then
            If tipoCupomEmite Like "CE*" Then
              frmCaixaSATDireto.Show vbModal
              Exit Sub
            ElseIf tipoCupomEmite = "CF" Then
              frmCaixaTEF.Show vbModal
              Exit Sub
            Else
              MsgBox "Esse caixa não está abilitado para vende a cliente-consumidor", vbCritical, "Cupom Fiscal desabilitado"
            End If
          End If
          
         
         If Mid(Trim(txtPedido.text), 1, 1) = "B" Then
            If VerificaSeEmiteCodigoZero = "S" Then
                If Len(Trim(txtPedido.text)) = 1 Then
                   frmCaixaRomaneioDireto.Show vbModal
                   wVerificaRomaneio = True
                   Exit Sub
                Else
                   txtPedido.text = Mid(txtPedido.text, 2, Len(txtPedido.text))
                   wVerificaRomaneio = True
                End If
            End If
         End If
         
         If Mid(Trim(txtPedido.text), 1, 1) = "M" Then
            If Len(Trim(txtPedido.text)) < 2 Then
                  MsgBox "Informe M e Número do pedido"
                  txtPedido.SelStart = 0
                  txtPedido.SelLength = Len(txtPedido.text)
                  Exit Sub
            Else
                 txtPedido.text = Mid(txtPedido.text, 2, Len(txtPedido.text))
                 wVerificaNotaManual = True
            End If
         End If
         
         
         If Mid(Trim(txtPedido.text), 1, 1) = "N" Then
            If Len(Trim(txtPedido.text)) < 2 Then
                  MsgBox "Informe N e Número do pedido"
                  txtPedido.SelStart = 0
                  txtPedido.SelLength = Len(txtPedido.text)
                  Exit Sub
            Else
                 txtPedido.text = Mid(txtPedido.text, 2, Len(txtPedido.text))
                 wVerificaNotaFiscal = True
            End If
         End If
          
          
         If Not IsNumeric(txtPedido) Then
            txtPedido.SelStart = 0
            txtPedido.SelLength = Len(txtPedido.text)
            Exit Sub
         End If
 
         If txtPedido.text = "" Then
            Exit Sub
         End If

        sql = ""
        'sql = "Select cliente from nfcapa where tiponota = 'PA' and numeroped = " & txtPedido.Text
        sql = "Select tiponota,Cliente,totalnota,serie,numeroped from nfcapa where tiponota in ('PA','TA','T0') and numeroped = " & txtPedido.text

        If rdoParametro.State <> 0 Then rdoParametro.Close
        rdoParametro.CursorLocation = adUseClient
        rdoParametro.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
        
        If Not rdoParametro.EOF Then
        
            If Trim(rdoParametro("tiponota")) = "PA" Then
        
                 If wVerificaRomaneio = True Then
                     frmCaixaRomaneio.Show vbModal
                     
                 ElseIf wVerificaNotaManual = True Then
                     frmCaixaNotaManual.Show vbModal
                 
                 ElseIf rdoParametro("Cliente") = "999999" And wVerificaNotaFiscal = False Then
                    If tipoCupomEmite Like "CE*" Then
                        frmCaixaSAT.Show vbModal
                    ElseIf tipoCupomEmite = "CF" Then
                        frmCaixaTEFPedido.Show vbModal
                    Else
                        
                    End If
                 
                 ElseIf rdoParametro("Cliente") > "0" And rdoParametro("Cliente") <= "999999" Then
                     'Unload frmCaixaNF
                     'If aceitaGarantia = True Then Unload frmCaixaNF
                     frmCaixaNF.Show vbModal
                
                 Else
                     MsgBox "Cliente não encontrado!"
                 End If
               
            
            Else
                        If frmTrans.Visible = False Then
            
                            sql = "select LO_Numero,LO_Endereco,TipoTransporte from  nfcapa,loja where lojat=lo_loja and NUMEROPED= " & txtPedido.text

                            If rdoTrans.State <> 0 Then rdoTrans.Close
                            rdoTrans.CursorLocation = adUseClient
                            rdoTrans.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
        
                            frmTrans.Visible = True
                            lblRetiradoPor.Caption = Mid$(rdoTrans("TipoTransporte"), 1, (InStr(rdoTrans("TipoTransporte"), "/") - 1))
                            lblSolicitadoPor.Caption = Mid$(rdoTrans("TipoTransporte"), (1 + InStr(rdoTrans("TipoTransporte"), "/")))
                            lblLojaDestino.Caption = rdoTrans("LO_Numero") & "-" & rdoTrans("LO_Endereco")
                            rdoTrans.Close
                            
                            
                    Else
                            notaTrans
                            frmTrans.Visible = False
                            
                    End If
                    
                      
                     End If

         '      End If
            
            
         '************** ATUALIZA RETAGUARDA
         
                   'sql = "exec SP_Atualiza_Processos_Venda_Central"
                   'rdoCNLoja.Execute sql
                   
        
                   'GLB_ConectouOK = False
                   'ConectaODBCMatriz
          
                   'If wConectouRetaguarda = True Then
                      
             
                      'sql = "Exec SP_Est_Transferencia_destino '" & RTrim(LTrim(GLB_Loja)) & "'"
                      'rdoCNRetaguarda.Execute sql
                
                   'End If
        
        Else
            MsgBox "Pedido não Existe ou Nota Fiscal já foi emitida.", vbCritical, "Aviso"
            txtPedido.SelStart = 0
            txtPedido.SelLength = Len(txtPedido.text)
        End If
        If rdoParametro.State <> 0 Then rdoParametro.Close
        'rdoParametro.Close
      Else
        txtPedido.text = ""
        MsgBox "Data do caixa incorreta.Favor efetuar o Fechamento", vbCritical, "Atenção"
        frmFechaCaixa.Show vbModal
      End If
    End If

End Sub

Function CriaMovimentoCaixa(ByVal nf As Double, ByVal serie As String, ByVal TotalNota As Double, ByVal loja As String, ByVal Grupo As Double, ByVal NroProtocolo As Double, ByVal nroCaixa As Integer, ByVal NroPedido As Double, ByVal tiponota As String)
    
    sql = "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
        & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio, MC_Protocolo, MC_NroCaixa, MC_DataProcesso, MC_Pedido,MC_TipoNota) values(" & GLB_ECF & ",'0','" & Trim(loja) & "', " _
        & " '" & Format(Date, "yyyy/mm/dd") & "'," & Grupo & ", " & nf & ",'" & serie & "', " _
        & "" & ConverteVirgula(Format(TotalNota, "##,###0.00")) & ", " _
        & "0,0,0,0,0,9,'A'," & NroProtocolo & "," & nroCaixa & ",'" & Format(Date, "yyyy/mm/dd") & "'," & NroPedido & ",'" & tiponota & "')"
        rdoCNLoja.Execute (sql)

End Function




Private Sub txtPedido_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        frmLoginCaixa.Show 1
    End If
End Sub

Private Sub webInternet1_Click()
    If txtPedido.Visible = True And txtPedido.Enabled = True Then
        txtPedido.SetFocus
    End If
End Sub


Private Function notaTrans()

                     If MsgBox("Deseja emitir NF de transferência? Pedido --> " & rdoParametro("numeroped") & ", " _
                           & " Valor --> " & rdoParametro("totalnota") & "", vbQuestion + vbYesNo, "Atenção") = vbNo Then
                               txtPedido.SelStart = 0
                               txtPedido.SelLength = Len(txtPedido.text)
                               rdoParametro.Close
                               Screen.MousePointer = vbNormal
                               Exit Function

                     Else


                                If Trim(rdoParametro("tiponota")) = "TA" And Trim(rdoParametro("serie")) = "00" Then
                                
                                    Dim wTotalNota As Double
                                    'Dim wNumProtocolo As String
                                    'Dim wNroCaixa As String
                                    
                                    
                                    wSequencia = txtPedido.text
                                    wTipoNota = Trim(rdoParametro("tiponota"))
                                    ImprimeTransferencia00 wSequencia
                                    
                                    '***************** ATUALIZA ESTOQUE LOJA
                                    'sql = ""
                                    'sql = "exec SP_EST_Transferencia '" & Trim(frmControlaCaixa.txtPedido.Text) & "'"
                                    'rdoCNLoja.Execute sql
                                    
                                    
                                    '**************** CRIA MOVIMENTO CAIXA
                                    sql = ""
                                    sql = "Select TotalNota, Protocolo, NroCaixa From NFCapa Where TipoNota = '" & Trim(wTipoNota) & _
                                          "' and Serie = '" & "00" & "' and " & _
                                          "NF = " & wSequencia & " and NumeroPed = " & wSequencia
                                    
                                    rdocontrole.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                                    
                                    If rdocontrole.EOF = False Then
                                        wTotalNota = rdocontrole("TotalNota")
                                    End If
                                    
                                    rdocontrole.Close
                                    CriaMovimentoCaixa wSequencia, Trim(wTipoNota), wTotalNota, GLB_Loja, "20109", Trim(lblProtocolo.Caption), Trim(lblNroCaixa.Caption), Trim(txtPedido.text), Trim(wTipoNota)
                                    'CriaMovimentoCaixa "        ", PegaSerieNota , rdocontrol, GLB_Loja, "20109", Trim(lblProtocolo.Caption), Trim(lblNroCaixa.Caption), Trim(txtPedido.Text), Trim(rdocontrole("tiponota"))
                                    
                                ElseIf Trim(rdoParametro("tiponota")) = "TA" And Trim(rdoParametro("serie")) = "CT" Then
                                
                                    ''defineImpressora
                                
                                    NroNotaFiscal = ExtraiSeq00Controle
                                    wSequencia = txtPedido.text
                                    ImprimeTransferencia00 wSequencia
                                    ImprimeTransferencia00 wSequencia
                                    sql = "exec sp_totaliza_capa_nota_fiscal_Loja " & txtPedido.text
                                    rdoCNLoja.Execute sql
                                    sql = "Select TotalNota, tiponota,serie From NFCapa Where NumeroPed = " & txtPedido.text
                                    rdocontrole.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                                    
                                    If rdocontrole.EOF = False Then
                                        CriaMovimentoCaixa NroNotaFiscal, rdocontrole("serie"), rdocontrole("TotalNota"), GLB_Loja, "20109", Trim(lblProtocolo.Caption), Trim(lblNroCaixa.Caption), Trim(txtPedido.text), Trim(rdocontrole("tiponota"))
                                        rdocontrole.Close
                                    Else
                                        MsgBox "Erro na criação do Movimento Caixa de Transferencia. Informe o TI."
                                    End If
                                Else

                                               If PegaSerieNota = "NE" Then
                                                    NroNotaFiscal = ExtraiSeqNEControle
                                               Else
                                                    NroNotaFiscal = ExtraiSeqNotaControle
                                               End If
                                               
                                               'FELIPE 2015/04/04
                                               'sql = "update nfcapa set nf = " & NroNotaFiscal & _
                                               '" where numeroped = " & Trim(txtPedido.text)
                                               'rdoCNLoja.Execute sql
                                               
                                               'sql = "update nfitens set nf = " & NroNotaFiscal & _
                                               '" where numeroped = " & Trim(txtPedido.text)
                                               'rdoCNLoja.Execute sql
                            
                                               'sql = "update CarimboNotaFiscal set CNF_NF = " & NroNotaFiscal & _
                                               '" where CNF_NumeroPed = " & Trim(txtPedido.text)
                                               'rdoCNLoja.Execute sql
                                               
                                               '''''sql = "exec SP_EST_Transferencia '" & Trim(txtPedido.Text) & "'"
                                               '''''rdoCNLoja.Execute sql
                            
                                              '**************** CRIA MOVIMENTO CAIXA
                                              sql = "exec sp_totaliza_capa_nota_fiscal_Loja " & txtPedido.text
                                              rdoCNLoja.Execute sql
                                              
                                              sql = "Select TotalNota, tiponota From NFCapa Where NumeroPed = " & txtPedido.text
                                              rdocontrole.Open sql, rdoCNLoja, adOpenForwardOnly, adLockPessimistic
                            
                                            If rdocontrole.EOF = False Then
                                               CriaMovimentoCaixa NroNotaFiscal, PegaSerieNota, rdocontrole("TotalNota"), GLB_Loja, "20109", Trim(lblProtocolo.Caption), Trim(lblNroCaixa.Caption), Trim(txtPedido.text), Trim(rdocontrole("tiponota"))
                                               rdocontrole.Close
                                            Else
                                               MsgBox "Erro na criação do Movimento Caixa de Transferencia. Informe o TI."
                                            End If
                                            
                                                        'If PegaSerieNota = "NE" Then
                                                            'Call CriaNFE(NroNotaFiscal, txtPedido.Text)
                                                        'Else
                                                            'Call EmiteNotafiscalTransferencia(NroNotaFiscal, rdoParametro("serie"))
                                                        'End If
                                                    'End If
                                                  
                                        End If
                                        
                                        pedido = txtPedido.text
                                        
                                        Screen.MousePointer = vbNormal
                                        frmStartaProcessos.Show vbModal
                                        txtPedido.text = ""
                                        'txtPedido.Text = ""
                                    
                                End If
                                

End Function




