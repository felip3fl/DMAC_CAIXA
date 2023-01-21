Attribute VB_Name = "ModuloIcoBandeja"
Option Explicit

Public Declare Function Shell_NotifyIcon Lib _
       "shell32.dll" Alias "Shell_NotifyIconA" _
       (ByVal dwMessage As Long, lpData As _
       NOTIFYICONDATA) As Long

Public Type NOTIFYICONDATA
  cbSize As Long
  Hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONDBLCLK = &H206

Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4

Public Enum Actions
  TrayAdd = &H0
  TrayModify = &H1
  TrayDelete = &H2
End Enum

Public Sub criaIconeBarra(Action As Actions, Hwnd As _
       Long, ToolTip As String, Icon As _
       StdPicture)
  Dim STray As NOTIFYICONDATA
  STray.uID = vbNull
  STray.uCallbackMessage = &H200
  STray.Hwnd = Hwnd
  STray.hIcon = Icon
  STray.szTip = ToolTip & vbNullChar
  STray.uFlags = NIF_MESSAGE Or NIF_ICON Or _
                 NIF_TIP
  STray.cbSize = Len(STray)
  Select Case Action
    Case NIM_ADD
      Call Shell_NotifyIcon(NIM_ADD, STray)
    Case NIM_MODIFY
      Call Shell_NotifyIcon(NIM_MODIFY, STray)
    Case NIM_DELETE
      Call Shell_NotifyIcon(NIM_DELETE, STray)
  End Select
End Sub

Public Sub excluirIconeBarra()

    Dim Tic As NOTIFYICONDATA

    Tic.cbSize = Len(Tic)
    'Tic.Hwnd = Picture1.Hwnd
    Tic.uID = 1&
    'erg = Shell_NotifyIcon(NIM_DELETE, Tic)

End Sub






