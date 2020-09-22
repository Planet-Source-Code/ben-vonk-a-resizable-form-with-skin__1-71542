Attribute VB_Name = "mdlWindowLayout"
Option Explicit

' Private Constants
Private Const GWL_EXSTYLE As Long = -20
Private Const GWL_STYLE   As Long = -16

' Public Enumeration
Public Enum RoundCorners
   None
   AllRound
   TopRound
   LeftRound
   RightRound
   BottomRound
End Enum

' Private Types
Private Type PointAPI
   X                      As Long
   Y                      As Long
End Type

Private Type MinMaxInfo
   ptReserved             As PointAPI
   ptMaxSize              As PointAPI
   ptMaxPosition          As PointAPI
   ptMinTrackSize         As PointAPI
   ptMaxTrackSize         As PointAPI
End Type

Private Type Rect
   Left                   As Long
   Top                    As Long
   Right                  As Long
   Bottom                 As Long
End Type

Private Type WindowPlacement
   Length                 As Long
   flags                  As Long
   showCmd                As Long
   ptMinPosition          As PointAPI
   ptMaxPosition          As PointAPI
   rcNormalPosition       As Rect
End Type

' Public Variable
Public SysMenuOpen        As Boolean

' Private Variables
Private m_FontBold        As Boolean
Private m_FontItalic      As Boolean
Private m_Window          As Form
Private m_Icon            As Image
Private m_ForeColor       As Long
Private m_hWnd            As Long
Private m_IconArea        As Long
Private m_ShadowColor     As Long
Private SizeWindow        As Long
Private m_Skin            As PictureBox
Private m_Fontsize        As Single
Private m_FontName        As String

' Private API's
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CombineRgn Lib "GDI32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function StretchBlt Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DrawIconEx Lib "User32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyHeight As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetCursorPos Lib "User32" (lpPoint As PointAPI) As Long
Private Declare Function GetSystemMenu Lib "User32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetSystemMetrics Lib "User32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowPlacement Lib "User32" (ByVal hWnd As Long, lpwndpl As WindowPlacement) As Long
Private Declare Function MoveWindow Lib "User32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Private Declare Function SetCapture Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function SetCursorPos Lib "User32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function TrackPopupMenu Lib "User32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, ByVal lprc As Any) As Long
Private Declare Function UnhookWindowsHookEx Lib "User32" (ByVal hHook As Long) As Long

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Sub CreateWindow(ByVal Width As Long, ByVal Height As Long, ByVal Curve As Long, ByVal RoundCorner As RoundCorners, ByVal Maximized As Boolean)

Const RGN_OR     As Long = 2

Dim lngRegion(1) As Long
Dim lngStart     As Long
Dim lngSize      As Long

   If RoundCorner = None Then
      lngRegion(0) = CreateRectRgn(0, 0, Width, Height)
      
   Else
      If Maximized Then
         lngSize = 2
         
      Else
         lngStart = 2
      End If
      
      lngRegion(0) = CreateRoundRectRgn(lngStart, lngStart, Width + lngSize, Height + lngSize, Curve, Curve)
      
      If RoundCorner = TopRound Then
         lngRegion(1) = CreateRectRgn(lngStart, Curve \ 2, Width + lngSize, Height + lngSize)
         
      ElseIf RoundCorner = LeftRound Then
         lngRegion(1) = CreateRectRgn(Curve \ 2, lngStart, Width + lngSize, Height + lngSize)
         
      ElseIf RoundCorner = RightRound Then
         lngRegion(1) = CreateRectRgn(lngStart, lngStart, Width - Curve \ 2 + lngSize, Height + lngSize)
         
      ElseIf RoundCorner = BottomRound Then
         lngRegion(1) = CreateRectRgn(lngStart, lngStart, Width + lngSize, Height - Curve \ 2 + lngSize)
      End If
      
      CombineRgn lngRegion(0), lngRegion(0), lngRegion(1), RGN_OR
      DeleteObject lngRegion(1)
   End If
   
   SetWindowRgn m_hWnd, lngRegion(0), True
   DeleteObject lngRegion(0)
   Erase lngRegion

End Sub

Public Sub DockToSysTray()

Dim lngHeight  As Long
Dim lngLeft    As Long
Dim lngTop     As Long
Dim lngWidth   As Long
Dim lngWindow  As Long
Dim wpmSysTray As WindowPlacement

   lngWindow = FindWindow("Shell_traywnd", vbNullString)
   GetWindowPlacement lngWindow, wpmSysTray
   
   With wpmSysTray.rcNormalPosition
      ' Right sided
      If .Left Then
         lngTop = 0
         lngLeft = 0
         lngWidth = .Left
         lngHeight = .Bottom
         
      ' Bottom sided
      ElseIf .Top Then
         lngTop = 0
         lngLeft = 0
         lngWidth = .Right
         lngHeight = .Top
         
      ' Left sided
      ElseIf .Bottom = Screen.Height / Screen.TwipsPerPixelY Then
         lngTop = 0
         lngLeft = .Right
         lngWidth = Screen.Width / Screen.TwipsPerPixelX - lngLeft
         lngHeight = .Bottom
         
      ' Top sided
      Else
         lngTop = .Bottom
         lngLeft = 0
         lngWidth = .Right
         lngHeight = Screen.Height / Screen.TwipsPerPixelY - lngTop
      End If
   End With
   
   MoveWindow m_hWnd, lngLeft, lngTop, lngWidth, lngHeight, 1

End Sub

Public Sub DrawSkin()

Const DI_NORMAL      As Long = &H3

Dim blnFontBold      As Boolean
Dim blnFontItalic    As Boolean
Dim intCount         As Integer
Dim lngBar           As Long
Dim lngTitleBarCurve As Long
Dim lngForeColor     As Long
Dim sngFontSize      As Single
Dim strFontName      As String

   With m_Window
      lngTitleBarCurve = .ScaleWidth * 0.55
      lngBar = .ScaleWidth * 0.3
      strFontName = .FontName
      sngFontSize = .FontSize
      blnFontBold = .FontBold
      blnFontItalic = .FontItalic
      lngForeColor = .ForeColor
      .Cls
      BitBlt .hDC, 0, 0, 9, 52, m_Skin.hDC, 0, 0, vbSrcCopy
      StretchBlt .hDC, 9, 0, lngTitleBarCurve, 52, m_Skin.hDC, 10, 0, 1, 52, vbSrcCopy
      BitBlt .hDC, lngTitleBarCurve, 0, 52, 52, m_Skin.hDC, 23, 0, vbSrcCopy
      StretchBlt .hDC, lngTitleBarCurve + 52, 0, .ScaleWidth - lngTitleBarCurve - 62, 52, m_Skin.hDC, 76, 0, 1, 52, vbSrcCopy
      BitBlt .hDC, .ScaleWidth - 10, 0, 10, 52, m_Skin.hDC, 77, 0, vbSrcCopy
      StretchBlt .hDC, 0, 52, 1, .ScaleHeight - 67, m_Skin.hDC, 0, 52, 1, 1, vbSrcCopy
      StretchBlt .hDC, .ScaleWidth - 1, 52, 1, .ScaleHeight - 67, m_Skin.hDC, 86, 52, 1, 1, vbSrcCopy
      StretchBlt .hDC, 1, 58, lngBar - 12, .ScaleHeight - 64, m_Skin.hDC, 7, 60, 1, 1, vbSrcCopy
      BitBlt .hDC, 0, .ScaleHeight - 16, 14, 16, m_Skin.hDC, 0, 58, vbSrcCopy
      BitBlt .hDC, .ScaleWidth - 14, .ScaleHeight - 16, 14, 16, m_Skin.hDC, 73, 58, vbSrcCopy
      StretchBlt .hDC, 14, .ScaleHeight - 6, .ScaleWidth - 28, 6, m_Skin.hDC, 14, 68, 1, 6, vbSrcCopy
      BitBlt .hDC, lngTitleBarCurve, 52, 20, 6, m_Skin.hDC, 22, 52, vbSrcCopy
      StretchBlt .hDC, 1, 52, lngTitleBarCurve - 1, 6, m_Skin.hDC, 1, 52, 1, 6, vbSrcCopy
      BitBlt .hDC, lngBar, 52, 11, 6, m_Skin.hDC, 11, 52, vbSrcCopy
      StretchBlt .hDC, lngBar, 58, 11, .ScaleHeight - 64, m_Skin.hDC, 11, 57, 11, 1, vbSrcCopy
      StretchBlt .hDC, lngBar - 11, 58, 11, .ScaleHeight - 64, m_Skin.hDC, 13, 58, 11, 1, vbSrcCopy
      DrawIconEx .hDC, (m_IconArea - m_Icon.Width) / 2, (m_IconArea - m_Icon.Height) / 2, m_Icon.Picture.Handle, m_Icon.Width, m_Icon.Height, 0, 0, DI_NORMAL
      .FontName = m_FontName
      .FontSize = m_Fontsize
      .FontBold = m_FontBold
      .FontItalic = m_FontItalic
      .ForeColor = m_ShadowColor
      
      For intCount = 1 To 0 Step -1
         .CurrentX = m_IconArea + intCount
         .CurrentY = (m_IconArea - .TextHeight("X")) / 2 + intCount
         m_Window.Print .Caption
         .ForeColor = m_ForeColor
      Next 'intCount
      
      .FontName = strFontName
      .FontSize = sngFontSize
      .FontBold = blnFontBold
      .FontItalic = blnFontItalic
      .ForeColor = lngForeColor
   End With

End Sub

Public Sub Initialize(ByRef Window As Form, ByVal Skin As PictureBox, ByVal Icon As Image, ByVal IconArea As Long)

   Set m_Window = Window
   Set m_Skin = Skin
   Set m_Icon = Icon
   m_IconArea = IconArea
   m_hWnd = Window.hWnd
   
   Call SubclassWindowSizing

End Sub

Public Sub MakeSizable(ByVal Sizable As Boolean)

Const WS_THICKFRAME As Long = &H40000

Dim lngStyle As Long

   lngStyle = GetWindowLong(m_hWnd, GWL_STYLE)
   
   If Sizable Then
      lngStyle = lngStyle Or WS_THICKFRAME
      
   Else
      lngStyle = lngStyle And Not WS_THICKFRAME
   End If
   
   SetWindowLong m_hWnd, GWL_STYLE, lngStyle

End Sub

Public Sub MovingWindow(ByVal lParam As Long)

Const HTCLIENT     As Long = &H1
Const WM_LBUTTONUP As Long = &H202

   ReleaseCapture
   
   Call RefreshSysMenu(m_hWnd)
   
   SetCapture m_hWnd
   ReleaseCapture
   PostMessage m_hWnd, WM_LBUTTONUP, HTCLIENT, lParam

End Sub

Public Sub RemoveTitleBar()

Const WS_CAPTION As Long = &HC00000

   SetWindowLong m_hWnd, GWL_STYLE, GetWindowLong(m_hWnd, GWL_STYLE) And Not WS_CAPTION

End Sub

Public Sub SetCaptionFont(ByVal FontName As String, ByVal FontSize As Single, Optional ByVal FontBold As Boolean, Optional ByVal FontItalic As Boolean, Optional ByVal ForeColor As Long = vbWhite, Optional ByVal ShadowColor As Long = &H404040)

   m_FontName = FontName
   m_Fontsize = FontSize
   m_FontBold = FontBold
   m_FontItalic = FontItalic
   m_ForeColor = ForeColor
   m_ShadowColor = ShadowColor

End Sub

Public Sub ShowSystemMenu()

Const SC_CLOSE        As Long = &HF060&
Const SC_MAXIMIZE     As Long = &HF030&
Const SC_MINIMIZE     As Long = &HF020&
Const SC_MOVE         As Long = &HF010&
Const SC_RESTORE      As Long = &HF120&
Const SC_SIZE         As Long = &HF000&
Const TPM_NONOTIFY    As Long = &H80
Const TPM_RETURNCMD   As Long = &H100
Const TPM_RIGHTBUTTON As Long = &H2
Const TPM_TOPALIGN    As Long = &H0

Dim lngLeft           As Long
Dim lngMenuItem       As Long
Dim lngTop            As Long
Dim ptaMouseXY        As PointAPI

   If SysMenuOpen Then
      SysMenuOpen = False
      Exit Sub
   End If
   
   With m_Window
      Call RefreshSysMenu(m_hWnd)
      
      DoEvents
      SysMenuOpen = True
      GetCursorPos ptaMouseXY
      lngMenuItem = TrackPopupMenu(GetSystemMenu(m_hWnd, 0), TPM_NONOTIFY Or TPM_RETURNCMD Or TPM_RIGHTBUTTON Or TPM_TOPALIGN, ptaMouseXY.X + 2, ptaMouseXY.Y + 2, 0, m_hWnd, ByVal 0&)
      
      Select Case lngMenuItem
         Case SC_RESTORE, SC_MAXIMIZE
            SysMenuOpen = False
            
            Call .ToggleWindowState
            
         Case SC_MOVE, SC_SIZE
            SysMenuOpen = False
            lngTop = .Top / Screen.TwipsPerPixelY + 5
            
            If lngMenuItem = SC_SIZE Then lngTop = lngTop - 5 + .Height / Screen.TwipsPerPixelY / 2
            
            lngLeft = .Left / Screen.TwipsPerPixelX + .Width / Screen.TwipsPerPixelX / 2
            SetCursorPos lngLeft, lngTop + 1
            DoEvents
            SetCursorPos lngLeft, lngTop - 1
            .MousePointer = vbSizeAll
            
         Case SC_MINIMIZE
            SysMenuOpen = False
            .WindowState = vbMinimized
            
         Case SC_CLOSE
            SysMenuOpen = False
            
            Call .EndResizableSkin
      End Select
   End With

End Sub

Public Sub Terminate()

   Call SubclassWindowSizing
   
   Set m_Skin = Nothing
   Set m_Icon = Nothing
   Set m_Window = Nothing

End Sub

Private Function WinProcSizing(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Const SM_CXSCREEN      As Long = 0
Const SM_CYSCREEN      As Long = 1
Const WM_GETMINMAXINFO As Long = &H24

Dim mmiValues          As MinMaxInfo

   If wMsg = WM_GETMINMAXINFO Then
      Call CopyMemory(mmiValues, lParam, LenB(mmiValues))
      
      With mmiValues
         .ptMinTrackSize.X = 430
         .ptMinTrackSize.Y = 260
         .ptMaxPosition.X = 0
         .ptMaxPosition.Y = 0
         .ptMaxTrackSize.X = GetSystemMetrics(SM_CXSCREEN)
         .ptMaxTrackSize.Y = GetSystemMetrics(SM_CYSCREEN)
         .ptMaxSize.X = GetSystemMetrics(SM_CXSCREEN)
         .ptMaxSize.Y = GetSystemMetrics(SM_CYSCREEN)
      End With
      
      Call CopyMemory(ByVal lParam, mmiValues, LenB(mmiValues))
      
   Else
      WinProcSizing = CallWindowProc(SizeWindow, hWnd, wMsg, wParam, lParam)
   End If

End Function

Private Sub RefreshSysMenu(ByVal hWnd As Long)

Const HTCAPTION        As Long = &H2
Const WM_NCLBUTTONDOWN As Long = &HA1

   SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&

End Sub

Private Sub SubclassWindowSizing()

Const GWL_WNDPROC As Long = -4

   If SizeWindow Then
      SetWindowLong m_hWnd, GWL_WNDPROC, SizeWindow
      SizeWindow = 0
      
   Else
      SizeWindow = SetWindowLong(m_hWnd, GWL_WNDPROC, AddressOf WinProcSizing)
   End If

End Sub
