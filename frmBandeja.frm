VERSION 5.00
Begin VB.Form frmBandeja 
   BackColor       =   &H00000000&
   Caption         =   "DMAC Caixa"
   ClientHeight    =   7605
   ClientLeft      =   7650
   ClientTop       =   345
   ClientWidth     =   6585
   Icon            =   "frmBandeja.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   6585
   Begin VB.Image imgTarefas 
      Height          =   11520
      Left            =   0
      Picture         =   "frmBandeja.frx":23FA
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "frmBandeja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    
    frmControlaCaixa.Show 1
End Sub

Private Sub Form_Load()
    imgTarefas.top = 0
    imgTarefas.left = 0
    Me.Height = (imgTarefas.Height) + 500
    Me.Width = (imgTarefas.Width)
    Me.top = -500
    Me.left = -100
    
    frmTrocaVersao.Show 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
