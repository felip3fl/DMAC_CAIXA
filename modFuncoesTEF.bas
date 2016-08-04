Attribute VB_Name = "modFuncoes"
Option Explicit


Global aceitaGarantia As Boolean

Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Declaraciones para 32 bits
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal Hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long

Public Const CB_SHOWDROPDOWN = &H14F

Dim lFlag As Boolean
Public iTransacao As Integer
Dim i As Integer
Dim response As Integer
Dim linhaArquivo As String
Public naoConfirmado As Boolean


