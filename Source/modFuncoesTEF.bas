Attribute VB_Name = "modFuncoes"
Option Explicit


Global aceitaGarantia As Boolean

Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Declaraciones para 32 bits
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long

Public Const CB_SHOWDROPDOWN = &H14F

Dim lFlag As Boolean
Public iTransacao As Integer
Dim I As Integer
Dim response As Integer
Dim linhaArquivo As String
Public naoConfirmado As Boolean


