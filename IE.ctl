VERSION 5.00
Begin VB.UserControl IEalt 
   ClientHeight    =   1545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3570
   InvisibleAtRuntime=   -1  'True
   Picture         =   "IE.ctx":0000
   ScaleHeight     =   1545
   ScaleWidth      =   3570
   ToolboxBitmap   =   "IE.ctx":0C50
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   210
      Top             =   870
   End
End
Attribute VB_Name = "IEalt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal Hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&
'if this = false then we want errors to be written and handled silently
'in the background so user is not bothered. Otherwise show errors in
'message box so we are made fully aware of all errors in design mode
Public bControlInDevelopmentMode As Boolean

'the hwnd of internet explorer in case you wish to embed in your form
Public ieHwnd As Long

'where errors will be written to
Public sErrPrintPath As String


'specify which Internet Explorer addons you want displayed.
'Setting all to false (do nothing) reduces overhead. Since IE
'is a memory pig this is probably a good idea
Public bShowIEToolbar As Boolean
Public bShowIEMenubar As Boolean
Public bShowIEStatusbar As Boolean

Private oIE As Object


Private Sub DisableIEclose()
On Error Resume Next
Dim lHndSysMenu As Long
    
    'get handle to IE system menu (close, min, max buttons)
    lHndSysMenu = GetSystemMenu(ieHwnd, 0)
    'remove close button
        'RemoveMenu lHndSysMenu, 6, MF_BYPOSITION
   'Remove seperator bar
        'RemoveMenu lHndSysMenu, 5, MF_BYPOSITION
        
    'MaximizeRestoredForm Me
End Sub
 
Private Sub UserControl_Resize()
On Error Resume Next
   'since this control is invisible at runtime
   'we make the control stay at 32 pixels if
   'attempted resize
   Dim w As Long, h As Long
   
   With Screen
       'screen resolutions vary. This code gets the value
       'of how many twips there is per pixel (a twip is much
       'smaller than a pixel) and multiply that by 32 so our
       'control always stays at 32 x 32
       w = (.TwipsPerPixelX * 32)
       h = (.TwipsPerPixelY * 32)
       UserControl.Width = w
       UserControl.Height = h
   End With
End Sub

Sub EmbedIE(ParentHwnd As Long)
  If (oIE Is Nothing) Then Exit Sub
  SetParent ieHwnd, ParentHwnd
End Sub

Sub Nav(sUrl As String)
On Error GoTo err_handler:
Dim lSetParent As Long
   
   'if user has not previously closed IE then navigate IE to
   'the new specified address [sUrl]
   If Not (oIE Is Nothing) Then
      oIE.Navigate sUrl
      Exit Sub
   End If
   
   
   'create IE without having to set references
   Set oIE = CreateObject("InternetExplorer.Application")
   
   'set IE addon options
   oIE.Toolbar = bShowIEToolbar
   oIE.MenuBar = bShowIEMenubar
   oIE.StatusBar = bShowIEStatusbar
   oIE.Silent = True
   
   'store hwnd of IE
   ieHwnd = oIE.Parent.Hwnd
   'prevent user from closing IE
   DisableIEclose
   
   'show IE
   oIE.Visible = True
   
   'Go to the web address [sUrl]
   oIE.Navigate sUrl
   
   'to demonstrate how you can place controls inside Internet Explorer
   'the power of this is only limited by your imagination
   'lSetParent = SetParent(picIE.hwnd, IEhwnd)
   
   'move to extreme upper left top corner
   'picIE.Move 0, 0
 
   

Exit Sub
err_handler:
   If Err.Number <> 0 Then
      Dim s As String
      s = "Error: IEalt.Nav." & Err.Number & "." & Err.Description & "." & Now

      If bControlInDevelopmentMode Then
          MsgBox s
      Else
          Call subSaveErrInfo(sErrPrintPath, s)
      End If

      Resume Next
   End If
End Sub


 
'print errors to file. If user of your software is having bug issues you
'can have them send this file to you in email so you can easily find and fix bugs
Private Sub subSaveErrInfo(sErrPrintPath As String, sErrorString As String)
On Error Resume Next

Dim inum As Integer

  inum = FreeFile 'get a free file from the system
  Open sErrPrintPath For Append As #inum 'append adds to the current file instead of overwriting
      Print #inum, sErrorString
  Close #inum
End Sub
