Attribute VB_Name = "ModuloResolucaoTelas"
Public ResX As Single
Public ResY As Single
Public OldX As Single
Public OldY As Single
Public resolucao As Boolean

'muda data e símbolo de R$
Public Const LOCALE_SSHORTDATE = &H1F
Public Const LOCALE_SCURRENCY = 20
Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean

Public Type resolucaoTela
    Linhas As Single
    Colunas As Single
End Type


' muda resolução do vídeo
Public Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type

Public Declare Function GetClipCursor Lib "user32.dll" (lprc As RECT) As Long

Private Declare Function EnumDisplaySettings Lib "user32" Alias _
"EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, _
lpDevMode As Any) As Boolean

Private Declare Function ChangeDisplaySettings Lib "user32" Alias _
"ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long

Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000

Private Type DEVMODE
   dmDeviceName As String * CCDEVICENAME
   dmSpecVersion As Integer
   dmDriverVersion As Integer
   dmSize As Integer
   dmDriverExtra As Integer
   dmFields As Long
   dmOrientation As Integer
   dmPaperSize As Integer
   dmPaperLength As Integer
   dmPaperWidth As Integer
   dmScale As Integer
   dmCopies As Integer
   dmDefaultSource As Integer
   dmPrintQuality As Integer
   dmColor As Integer
   dmDuplex As Integer
   dmYResolution As Integer
   dmTTOption As Integer
   dmCollate As Integer
   dmFormName As String * CCFORMNAME
   dmUnusedPadding As Integer
   dmBitsPerPel As Integer
   dmPelsWidth As Long
   dmPelsHeight As Long
   dmDisplayFlags As Long
   dmDisplayFrequency As Long
End Type

Dim DevM As DEVMODE

Public resolucaoOriginal As resolucaoTela

Public Sub AlterarResolucao(iWidth As Single, iHeight As Single)

    If Glb_AlteraResolucao = True Then
    
       Dim a As Boolean
       Dim i As Long
       Do
          a = EnumDisplaySettings(0&, i, DevM)
          i = i + 1
       Loop Until (a = False)
    
       Dim B As Long
       DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
       DevM.dmPelsWidth = iWidth
       DevM.dmPelsHeight = iHeight
       B = ChangeDisplaySettings(DevM, 0)
   
   End If
End Sub

Public Function resolucaoTela() As resolucaoTela
    resolucaoTela.Linhas = Screen.Height / Screen.TwipsPerPixelX
    resolucaoTela.Colunas = Screen.Width / Screen.TwipsPerPixelY
End Function





