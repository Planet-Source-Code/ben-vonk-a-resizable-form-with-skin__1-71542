VERSION 5.00
Begin VB.UserControl CrystalButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   492
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   ClipBehavior    =   0  'None
   DefaultCancel   =   -1  'True
   HitBehavior     =   0  'None
   LockControls    =   -1  'True
   ScaleHeight     =   41
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   125
   ToolboxBitmap   =   "CrystalButton.ctx":0000
End
Attribute VB_Name = "CrystalButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CrystalButton Control
'
'Author Ben Vonk
'16-04-2008 First version (Based on Noel A. Dacara's 'dcButton control' at http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=65941&lngWId=1)

Option Explicit

' Public Events
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseEnter()
Public Event MouseLeave()
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)

' Private Constants
Private Const ALL_MESSAGES  As Long = -1
Private Const DT_CENTER     As Long = &H1
Private Const DT_NOCLIP     As Long = &H100
Private Const DT_WORDBREAK  As Long = &H10
Private Const GWL_WNDPROC   As Long = -4
Private Const PATCH_05      As Long = 93
Private Const PATCH_09      As Long = 137
Private Const PS_SOLID      As Long = 0
Private Const RGN_OR        As Long = 2
Private Const WM_ACTIVATE   As Long = &H6
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_NCACTIVATE As Long = &H86

' Public Enumerations
Public Enum ButtonShapes
   ShapeNone
   ShapeLeft
   ShapeRight
   ShapeSides
End Enum

Public Enum CornerAngles
   CornerThin
   CornerSmall
   CornerMedium
   CornerBig
   CornerFull
End Enum

Public Enum OLEDropTypes
   odNone
   odManual
End Enum

Public Enum PictureAlignments
   BehindText
   BottomEdge
   BottomOfCaption
   LeftEdge
   LeftOfCaption
   RightEdge
   RightOfCaption
   TopEdge
   TopOfCaption
End Enum

Public Enum PictureSizes
   SizePicture
   Size16x16
   Size24x24
   Size32x32
   Size48x48
   Size64x64
End Enum

Public Enum UserColors
   DownColor
   FocusBorder
   GrayColor
   GrayText
   HoverColor
   StartColor
End Enum

' Private Enumeration
Private Enum ButtonStates
   IsNormal
   IsHot
   IsDown
   IsDisabled
End Enum

Private Enum MsgWhen
   MSG_BEFORE = 1
   MSG_AFTER = 2
   MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER
End Enum

' Private Types
Private Type BitmapInfoHeader
   biSize                   As Long
   biWidth                  As Long
   biHeight                 As Long
   biPlanes                 As Integer
   biBitCount               As Integer
   biCompression            As Long
   biSizeImage              As Long
   biXPelsPerMeter          As Long
   biYPelsPerMeter          As Long
   biClrUsed                As Long
   biClrImportant           As Long
End Type

Private Type RGBQuad
   rgbBlue                  As Byte
   rgbGreen                 As Byte
   rgbRed                   As Byte
End Type

Private Type BitmapInfo
   bmiHeader                As BitmapInfoHeader
   bmiColors                As RGBQuad
End Type

Private Type ButtonColorType
   BackColor                As Long
   DownColor                As Long
   FocusBorder              As Long
   ForeColor                As Long
   GrayColor                As Long
   GrayText                 As Long
   HoverColor               As Long
   MaskColor                As Long
   StartColor               As Long
End Type

Private Type ButtonPropertyType
   BackColor                As Long
   Blocked                  As Boolean
   Caption                  As String
   CheckBox                 As Boolean
   CornerAngle              As CornerAngles
   Enabled                  As Boolean
   ForeColor                As Long
   HandPointer              As Boolean
   MaskColor                As Long
   Picture                  As StdPicture
   PicAlign                 As PictureAlignments
   PicOpacity               As Single
   PicSize                  As PictureSizes
   PicHeight                As Long
   PicWidth                 As Long
   Shape                    As ButtonShapes
   ShineFullLeft            As Boolean
   ShineFullRight           As Boolean
   Sound                    As Boolean
   UseMask                  As Boolean
   Value                    As Boolean
End Type

Private Type Rect
   Left                     As Long
   Top                      As Long
   Right                    As Long
   Bottom                   As Long
End Type

Private Type ButtonSettingType
   Button                   As Integer
   Caption                  As Rect
   Cursor                   As Long
   AsDefault                As Boolean
   Focus                    As Rect
   HasFocus                 As Boolean
   Height                   As Long
   Picture                  As Rect
   State                    As ButtonStates
   Width                    As Long
End Type

Private Type ColorsRGB
   Red                      As Long
   Green                    As Long
   Blue                     As Long
End Type

Private Type KeyboardInput
   dwType                   As Long
   wVk                      As Integer
   wScan                    As Integer
   dwFlags                  As Long
   dwTime                   As Long
   dwExtraInfo              As Long
   Dummy                    As Double
End Type

Private Type OSVersionInfo
   dwOSVersionInfoSize      As Long
   dwMajorVersion           As Long
   dwMinorVersion           As Long
   dwBuildNumber            As Long
   dwPlatformId             As Long
   szCSDVersion             As String * 128
End Type

Private Type PointAPI
   X                        As Long
   Y                        As Long
End Type

Private Type SubclassDataType
   hWnd                     As Long
   nAddrSclass              As Long
   nAddrOrig                As Long
   nMsgCountA               As Long
   nMsgCountB               As Long
   aMsgTabelA()             As Long
   aMsgTabelB()             As Long
End Type

Private Type TrackMouseEventType
   cbSize                   As Long
   dwFlags                  As Long
   hwndTrack                As Long
   dwHoverTime              As Long
End Type

' Private Variables
Private ButtonHasFocus      As Boolean
Private ButtonIsDown        As Boolean
Private CalculateRect       As Boolean
Private ControlHidden       As Boolean
Private IsTracking          As Boolean
Private IsNT                As Boolean
Private MouseIsDown         As Boolean
Private MouseOnButton       As Boolean
Private ParentActive        As Boolean
Private RedrawOnResize      As Boolean
Private SpacebarIsDown      As Boolean
Private TrackUser32         As Boolean
Private m_ButtonColors      As ButtonColorType
Private m_ButtonProperty    As ButtonPropertyType
Private m_ButtonSettings    As ButtonSettingType
Private m_BorderPoints()    As PointAPI
Private SubclassData()      As SubclassDataType

' Private API's
Private Declare Function TrackMouseEventComCtl Lib "ComCtl32" Alias "_TrackMouseEvent" (ByRef lpEventTrack As TrackMouseEventType) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CombineRgn Lib "GDI32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function CreatePen Lib "GDI32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateSolidBrush Lib "GDI32" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function GetDIBits Lib "GDI32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BitmapInfo, ByVal wUsage As Long) As Long
Private Declare Function GetNearestColor Lib "GDI32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function LineTo Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As PointAPI) As Long
Private Declare Function Polyline Lib "GDI32" (ByVal hDC As Long, ByRef lpPoint As PointAPI, ByVal nCount As Long) As Long
Private Declare Function PtInRegion Lib "GDI32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function Rectangle Lib "GDI32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, ByRef Bits As Any, ByRef BitsInfo As BitmapInfo, ByVal wUsage As Long) As Long
Private Declare Function SetPixelV Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "GDI32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function FreeLibrary Lib "Kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetModuleHandle Lib "Kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "Kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVersionInfo) As Long
Private Declare Function GlobalAlloc Lib "Kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function LoadLibrary Lib "Kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function OleTranslateColor Lib "OLEAut32" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Private Declare Function CopyRect Lib "User32" (ByRef lpDestRect As Rect, ByRef lpSourceRect As Rect) As Long
Private Declare Function DrawIconEx Lib "User32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DrawText Lib "User32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "User32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function GetCapture Lib "User32" () As Long
Private Declare Function GetClientRect Lib "User32" (ByVal hWnd As Long, ByRef lpRect As Rect) As Long
Private Declare Function GetCursorPos Lib "User32" (ByRef lpPoint As PointAPI) As Long
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function LoadCursor Lib "User32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function OffsetRect Lib "User32" (ByRef lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Declare Function SendInput Lib "User32" (ByVal nInputs As Long, pInputs As Any, ByVal cbSize As Long) As Long
Private Declare Function SetCapture Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function SetCursor Lib "User32" (ByVal hCursor As Long) As Long
Private Declare Function SetRect Lib "User32" (ByRef lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowLongA Lib "User32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function TrackMouseEvent Lib "User32" (ByRef lpEventTrack As TrackMouseEventType) As Long
Private Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SoundPlay Lib "WinMM" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Sub Subclass_WndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lhWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)

Dim blnFocus As Boolean
Dim blnForce As Boolean

   Select Case uMsg
      Case WM_MOUSELEAVE
         MouseOnButton = False
         IsTracking = False
         
         If Not SpacebarIsDown Then
            If m_ButtonProperty.Enabled Then Call DrawButton(IsNormal, (m_ButtonSettings.Button = 1))
            
            RaiseEvent MouseLeave
         End If
         
      Case WM_ACTIVATE, WM_NCACTIVATE
         ParentActive = wParam
         
         If ParentActive Then
            If m_ButtonProperty.Enabled Then If ButtonHasFocus Or m_ButtonSettings.AsDefault Then Call DrawButton(IsNormal, True)
            
         Else
            blnFocus = ButtonHasFocus
            blnForce = ButtonHasFocus Or m_ButtonSettings.AsDefault Or MouseOnButton
            ButtonHasFocus = False
            ButtonIsDown = False
            MouseIsDown = False
            MouseOnButton = False
            SpacebarIsDown = False
            m_ButtonSettings.AsDefault = False
                 
            If m_ButtonProperty.Enabled Then Call DrawButton(IsNormal, blnForce)
            
            m_ButtonSettings.AsDefault = Ambient.DisplayAsDefault
            ButtonHasFocus = blnFocus
         End If
   End Select

End Sub

Private Function Subclass_AddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long

   Subclass_AddrFunc = GetProcAddress(GetModuleHandle(sDLL), sProc)
   Debug.Assert Subclass_AddrFunc

End Function

Private Function Subclass_Index(ByVal lhWnd As Long, Optional ByVal bAdd As Boolean) As Long

   For Subclass_Index = UBound(SubclassData) To 0 Step -1
      If SubclassData(Subclass_Index).hWnd = lhWnd Then
         If Not bAdd Then Exit Function
         
      ElseIf SubclassData(Subclass_Index).hWnd = 0 Then
         If bAdd Then Exit Function
      End If
   Next 'Subclass_Index
   
   If Not bAdd Then Debug.Assert False

End Function

Private Function Subclass_InIDE() As Boolean

   Debug.Assert Subclass_SetTrue(Subclass_InIDE)

End Function

Private Function Subclass_Initialize(ByVal lhWnd As Long) As Long

Const CODE_LEN                  As Long = 200
Const GMEM_FIXED                As Long = 0
Const PATCH_01                  As Long = 18
Const PATCH_02                  As Long = 68
Const PATCH_03                  As Long = 78
Const PATCH_06                  As Long = 116
Const PATCH_07                  As Long = 121
Const PATCH_0A                  As Long = 186
Const FUNC_CWP                  As String = "CallWindowProcA"
Const FUNC_EBM                  As String = "EbMode"
Const FUNC_SWL                  As String = "SetWindowLongA"
Const MOD_USER                  As String = "User32"
Const MOD_VBA5                  As String = "vba5"
Const MOD_VBA6                  As String = "vba6"

Static bytBuffer(1 To CODE_LEN) As Byte
Static lngCWP                   As Long
Static lngEbMode                As Long
Static lngSWL                   As Long

Dim lngCount                    As Long
Dim lngIndex                    As Long
Dim strHex                      As String

   If bytBuffer(1) Then
      lngIndex = Subclass_Index(lhWnd, True)
      
      If lngIndex = -1 Then
         lngIndex = UBound(SubclassData) + 1
         
         ReDim Preserve SubclassData(lngIndex) As SubclassDataType
      End If
      
      Subclass_Initialize = lngIndex
      
   Else
      strHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D000000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
      
      For lngCount = 1 To CODE_LEN
         bytBuffer(lngCount) = Val("&H" & Left(strHex, 2))
         strHex = Mid(strHex, 3)
      Next 'lngCount
      
      If Subclass_InIDE Then
         bytBuffer(16) = &H90
         bytBuffer(17) = &H90
         lngEbMode = Subclass_AddrFunc(MOD_VBA6, FUNC_EBM)
         
         If lngEbMode = 0 Then lngEbMode = Subclass_AddrFunc(MOD_VBA5, FUNC_EBM)
      End If
      
      lngCWP = Subclass_AddrFunc(MOD_USER, FUNC_CWP)
      lngSWL = Subclass_AddrFunc(MOD_USER, FUNC_SWL)
      
      ReDim SubclassData(0) As SubclassDataType
   End If
   
   With SubclassData(lngIndex)
      .hWnd = lhWnd
      .nAddrSclass = GlobalAlloc(GMEM_FIXED, CODE_LEN)
      .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSclass)
      
      Call CopyMemory(ByVal .nAddrSclass, bytBuffer(1), CODE_LEN)
      Call Subclass_PatchRel(.nAddrSclass, PATCH_01, lngEbMode)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_02, .nAddrOrig)
      Call Subclass_PatchRel(.nAddrSclass, PATCH_03, lngSWL)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_06, .nAddrOrig)
      Call Subclass_PatchRel(.nAddrSclass, PATCH_07, lngCWP)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_0A, ObjPtr(Me))
   End With

End Function

Private Function Subclass_SetTrue(ByRef bValue As Boolean) As Boolean

   Subclass_SetTrue = True
   bValue = True

End Function

Private Sub Subclass_AddMsg(ByVal lhWnd As Long, ByVal uMsg As Long, Optional ByVal When As MsgWhen = MSG_AFTER)

   With SubclassData(Subclass_Index(lhWnd))
      If When And MSG_BEFORE Then Call Subclass_DoAddMsg(uMsg, .aMsgTabelB, .nMsgCountB, MSG_BEFORE, .nAddrSclass)
      If When And MSG_AFTER Then Call Subclass_DoAddMsg(uMsg, .aMsgTabelA, .nMsgCountA, MSG_AFTER, .nAddrSclass)
   End With

End Sub

Private Sub Subclass_DoAddMsg(ByVal uMsg As Long, ByRef aMsgTabel() As Long, ByRef nMsgCount As Long, ByVal When As MsgWhen, ByVal nAddr As Long)

Const PATCH_04 As Long = 88
Const PATCH_08 As Long = 132

Dim lngEntry   As Long

   ReDim lngOffset(1) As Long
   
   If uMsg = ALL_MESSAGES Then
      nMsgCount = ALL_MESSAGES
      
   Else
      For lngEntry = 1 To nMsgCount - 1
         If aMsgTabel(lngEntry) = 0 Then
            aMsgTabel(lngEntry) = uMsg
            
            GoTo ExitSub
            
         ElseIf aMsgTabel(lngEntry) = uMsg Then
            GoTo ExitSub
         End If
      Next 'lngEntry
      
      nMsgCount = nMsgCount + 1
      
      ReDim Preserve aMsgTabel(1 To nMsgCount) As Long
      
      aMsgTabel(nMsgCount) = uMsg
   End If
   
   If When = MSG_BEFORE Then
      lngOffset(0) = PATCH_04
      lngOffset(1) = PATCH_05
      
   Else
      lngOffset(0) = PATCH_08
      lngOffset(1) = PATCH_09
   End If
   
   If uMsg <> ALL_MESSAGES Then Call Subclass_PatchVal(nAddr, lngOffset(0), VarPtr(aMsgTabel(1)))
   
   Call Subclass_PatchVal(nAddr, lngOffset(1), nMsgCount)
   
ExitSub:
   Erase lngOffset

End Sub

Private Sub Subclass_PatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)

   Call CopyMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)

End Sub

Private Sub Subclass_PatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)

   Call CopyMemory(ByVal nAddr + nOffset, nValue, 4)

End Sub

Private Sub Subclass_Stop(ByVal lhWnd As Long)

   With SubclassData(Subclass_Index(lhWnd))
      SetWindowLongA .hWnd, GWL_WNDPROC, .nAddrOrig
      
      Call Subclass_PatchVal(.nAddrSclass, PATCH_05, 0)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_09, 0)
      
      GlobalFree .nAddrSclass
      .hWnd = 0
      .nMsgCountA = 0
      .nMsgCountB = 0
      Erase .aMsgTabelA, .aMsgTabelB
   End With

End Sub

Private Sub Subclass_Terminate()

Dim lngCount As Long

   For lngCount = UBound(SubclassData) To 0 Step -1
      If SubclassData(lngCount).hWnd Then Call Subclass_Stop(SubclassData(lngCount).hWnd)
   Next 'lngCount

End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used for the button."

   BackColor = m_ButtonProperty.BackColor

End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)

   m_ButtonProperty.BackColor = NewBackColor
   PropertyChanged "BackColor"
   
   Call SetButtonColors
   Call DrawButton(Force:=True)

End Property

Public Property Get Blocked() As Boolean
Attribute Blocked.VB_Description = "Determines whether an button is blocked for use."

   Blocked = m_ButtonProperty.Blocked

End Property

Public Property Let Blocked(ByVal NewBlocked As Boolean)

   m_ButtonProperty.Blocked = NewBlocked
   PropertyChanged "Blocked"

End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in the button."
Attribute Caption.VB_UserMemId = -518
Attribute Caption.VB_MemberFlags = "200"

   Caption = m_ButtonProperty.Caption

End Property

Public Property Let Caption(ByVal NewCaption As String)

   m_ButtonProperty.Caption = NewCaption
   AccessKeys = GetAccessKey(NewCaption)
   PropertyChanged "Caption"
   
   Call Refresh

End Property

Public Property Get CheckBoxMode() As Boolean
Attribute CheckBoxMode.VB_Description = "Returns/sets the type of control the button will observe."

   CheckBoxMode = m_ButtonProperty.CheckBox

End Property

Public Property Let CheckBoxMode(ByVal NewCheckBoxMode As Boolean)

   m_ButtonProperty.CheckBox = NewCheckBoxMode
   
   If Not Value And m_ButtonProperty.Value Then
      m_ButtonProperty.Value = False
      PropertyChanged "Value"
   End If
   
   PropertyChanged "CheckBox"
   
   Call DrawButton(Force:=True)

End Property

Public Property Get CornerAngle() As CornerAngles
Attribute CornerAngle.VB_Description = "Returns/sets the angle of the button corners."

   CornerAngle = m_ButtonProperty.CornerAngle

End Property

Public Property Let CornerAngle(ByVal NewCornerAngle As CornerAngles)

   m_ButtonProperty.CornerAngle = NewCornerAngle
   PropertyChanged "CornerAngle"
   
   Call Refresh

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value to determine whether the button can respond to events."

   Enabled = m_ButtonProperty.Enabled

End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)

   m_ButtonProperty.Enabled = NewEnabled
   UserControl.Enabled = NewEnabled
   
   If Not NewEnabled Then
      Call DrawButton(IsDisabled)
      
   ElseIf MouseOnButton Then
      Call DrawButton(IsHot)
      
   Else
      Call DrawButton(IsNormal)
   End If
   
   PropertyChanged "Enabled"

End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."

    Set Font = UserControl.Font

End Property

Public Property Let Font(ByRef NewFont As StdFont)

   Set Font = NewFont

End Property

Public Property Set Font(ByRef NewFont As StdFont)

   Set UserControl.Font = NewFont
   PropertyChanged "Font"
   
   Call Refresh

End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the text color of the button caption."

   ForeColor = m_ButtonProperty.ForeColor

End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)

   m_ButtonProperty.ForeColor = NewForeColor
   
   If m_ButtonProperty.Enabled Then
      Call SetButtonColors
      Call DrawButton(Force:=True)
   End If
   
   PropertyChanged "ForeColor"

End Property

Public Property Get HandPointer() As Boolean
Attribute HandPointer.VB_Description = "Returns/sets a value to determine whether the control uses the system's hand pointer as its cursor."

   HandPointer = m_ButtonProperty.HandPointer

End Property

Public Property Let HandPointer(ByVal NewHandPointer As Boolean)

   m_ButtonProperty.HandPointer = NewHandPointer
   PropertyChanged "HandPointer"

End Property

Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets a color in a button's picture to be transparent."

   MaskColor = m_ButtonProperty.MaskColor

End Property

Public Property Let MaskColor(ByVal NewMaskColor As OLE_COLOR)

   m_ButtonProperty.MaskColor = NewMaskColor
   m_ButtonColors.MaskColor = TranslateColor(NewMaskColor)
   PropertyChanged "MaskColor"
   
   Call DrawButton(Force:=True)

End Property

Public Property Get MouseIcon() As IPictureDisp
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon for the button."

   Set MouseIcon = UserControl.MouseIcon

End Property

Public Property Set MouseIcon(ByRef NewMouseIcon As IPictureDisp)

   Set UserControl.MouseIcon = NewMouseIcon
   
   If NewMouseIcon Is Nothing Then
      MousePointer = vbDefault
      
   Else
      MousePointer = vbCustom
   End If
   
   PropertyChanged "MouseIcon"

End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when cursor over the button."

   MousePointer = UserControl.MousePointer

End Property

Public Property Let MousePointer(ByVal NewMousePointer As MousePointerConstants)

   If (NewMousePointer <> vbCustom) And m_ButtonProperty.HandPointer Then HandPointer = False
   
   UserControl.MousePointer = NewMousePointer
   PropertyChanged "MousePointer"

End Property

Public Property Get OLEDropMode() As OLEDropTypes
Attribute OLEDropMode.VB_Description = "Returns/sets whether this object can act as an OLE drop target."

   OLEDropMode = UserControl.OLEDropMode

End Property

Public Property Let OLEDropMode(ByVal NewOLEDropMode As OLEDropTypes)

   UserControl.OLEDropMode = NewOLEDropMode
   PropertyChanged "OLEDropMode"

End Property

Public Property Get Picture() As StdPicture
Attribute Picture.VB_Description = "Returns/sets the picture displayed on a normal state button."

   Set Picture = m_ButtonProperty.Picture

End Property

Public Property Let Picture(ByRef NewPicture As StdPicture)

   Set Picture = NewPicture

End Property

Public Property Set Picture(ByRef NewPicture As StdPicture)

   With m_ButtonProperty
      If NewPicture Is Nothing Then
         .PicOpacity = 0
         
      ElseIf .PicOpacity = 0 Then
         .PicOpacity = 1
      End If
      
      Set .Picture = NewPicture
      PictureSize = .PicSize
      PropertyChanged "Picture"
      PropertyChanged "PicOpacity"
   End With
   
   Call Refresh

End Property

Public Property Get PictureAlignment() As PictureAlignments
Attribute PictureAlignment.VB_Description = "Returns/sets a value to determine where to draw the picture in the button."

   PictureAlignment = m_ButtonProperty.PicAlign

End Property

Public Property Let PictureAlignment(ByVal NewPictureAlignment As PictureAlignments)

   m_ButtonProperty.PicAlign = NewPictureAlignment
   PropertyChanged "PicAlign"
   
   Call Refresh

End Property

Public Property Get PictureSize() As PictureSizes
Attribute PictureSize.VB_Description = "Returns/sets a value to determine the size of the picture to draw."

   PictureSize = m_ButtonProperty.PicSize

End Property

Public Property Let PictureSize(ByVal NewPictureSize As PictureSizes)

   With m_ButtonProperty
      If .Picture Is Nothing Then
         .PicSize = SizePicture
         .PicHeight = 0
         .PicWidth = 0
         
      Else
         .PicSize = NewPictureSize
         
         Call SetPictureSize
      End If
   End With
   
   PropertyChanged "PicSize"
   
   Call Refresh

End Property

Public Property Get PictureOpacity() As Long
Attribute PictureOpacity.VB_Description = "Returns/sets a value in percent how the pictures will be blended to the button."

   PictureOpacity = m_ButtonProperty.PicOpacity * 100

End Property

Public Property Let PictureOpacity(ByVal NewPictureOpacity As Long)

   m_ButtonProperty.PicOpacity = TranslateNumber(NewPictureOpacity, 10, 100) / 100
   PropertyChanged "PicOpacity"
   
   Call DrawButton(Force:=True)

End Property

Public Property Get Shape() As ButtonShapes
Attribute Shape.VB_Description = "Returns/sets a value to determine the shape used to draw the button."

   Shape = m_ButtonProperty.Shape

End Property

Public Property Let Shape(ByVal NewShape As ButtonShapes)

   m_ButtonProperty.Shape = NewShape
   PropertyChanged "Shape"
   
   Call Refresh

End Property

Public Property Get ShineFullLeft() As Boolean
Attribute ShineFullLeft.VB_Description = "Determines whether the shine of the button will go completly to the left side."

   ShineFullLeft = m_ButtonProperty.ShineFullLeft

End Property

Public Property Let ShineFullLeft(ByVal NewShineFullLeft As Boolean)

   m_ButtonProperty.ShineFullLeft = NewShineFullLeft
   PropertyChanged "ShineLeft"
   
   Call Refresh

End Property

Public Property Get ShineFullRight() As Boolean
Attribute ShineFullRight.VB_Description = "Determines whether the shine of the button will go completly to the right side."

   ShineFullRight = m_ButtonProperty.ShineFullRight

End Property

Public Property Let ShineFullRight(ByVal NewShineFullRight As Boolean)

   m_ButtonProperty.ShineFullRight = NewShineFullRight
   PropertyChanged "ShineRight"
   
   Call Refresh

End Property

Public Property Get Sound() As Boolean
Attribute Sound.VB_Description = "Determins whether a sound will be played when a button is clicked."

   Sound = m_ButtonProperty.Sound

End Property

Public Property Let Sound(ByVal NewSound As Boolean)

   m_ButtonProperty.Sound = NewSound
   PropertyChanged "Sound"

End Property

Public Property Get UseMaskColor() As Boolean
Attribute UseMaskColor.VB_Description = "Returns/sets a value to determine whether to use MaskColor to create transparent areas of the picture."

   UseMaskColor = m_ButtonProperty.UseMask

End Property

Public Property Let UseMaskColor(ByVal NewUseMaskColor As Boolean)

   m_ButtonProperty.UseMask = NewUseMaskColor
   PropertyChanged "UseMask"
   
   Call DrawButton(Force:=True)
    
End Property

Public Property Get Value() As Boolean
Attribute Value.VB_Description = "Returns/sets the value or state of the button."

   Value = m_ButtonProperty.Value

End Property

Public Property Let Value(ByVal NewValue As Boolean)

   m_ButtonProperty.Value = NewValue
   
   If Value And Not m_ButtonProperty.CheckBox Then
      If Ambient.UserMode Then
         m_ButtonSettings.Button = vbLeftButton
         
         Call UserControl_Click
         
      Else
         m_ButtonProperty.Value = False
      End If
      
   Else
      Call DrawButton(Force:=True)
   End If
   
   PropertyChanged "Value"

End Property

Public Function hWnd() As Long

   hWnd = UserControl.hWnd

End Function

Public Sub OverrideColor(ByRef Property As UserColors, ByVal Color As Long, Optional ByVal NoRedraw As Boolean)

   Color = TranslateColor(Color)
   
   Select Case Property
      Case DownColor
         m_ButtonColors.DownColor = Color
         
      Case FocusBorder
         m_ButtonColors.FocusBorder = Color
         
      Case GrayColor
         m_ButtonColors.GrayColor = Color
         
      Case GrayText
         m_ButtonColors.GrayText = Color
         
      Case HoverColor
         m_ButtonColors.HoverColor = Color
         
      Case StartColor
         m_ButtonColors.StartColor = Color
         
      Case Else
         Err.Raise 5
         Exit Sub
   End Select
   
   If Not NoRedraw Then Call DrawButton(Force:=True)

End Sub

Public Sub Refresh()

   CalculateRect = True
   
   Call DrawButton(Force:=True)

End Sub

Private Function BlendColors(ByVal Color1 As Long, ByVal Color2 As Long, Optional ByVal Percent As Single = 0.5) As Long

Dim RGBColor(1) As ColorsRGB

   RGBColor(0) = GetRGB(Color1)
   RGBColor(1) = GetRGB(Color2)
   
   With RGBColor(0)
      BlendColors = GetColor(.Red + (RGBColor(1).Red - .Red) * Percent, .Green + (RGBColor(1).Green - .Green) * Percent, .Blue + (RGBColor(1).Blue - .Blue) * Percent)
   End With
   
   Erase RGBColor

End Function

Private Function BlendRGBQuad(RGB1 As RGBQuad, RGB2 As RGBQuad, Optional PercentInDecimal As Single = 0.5) As RGBQuad

Dim RGBColor As ColorsRGB

   With RGBColor
      .Red = RGB2.rgbRed
      .Green = RGB2.rgbGreen
      .Blue = RGB2.rgbBlue
      .Red = .Red - RGB1.rgbRed
      .Green = .Green - RGB1.rgbGreen
      .Blue = .Blue - RGB1.rgbBlue
      BlendRGBQuad.rgbRed = RGB1.rgbRed + .Red * PercentInDecimal
      BlendRGBQuad.rgbGreen = RGB1.rgbGreen + .Green * PercentInDecimal
      BlendRGBQuad.rgbBlue = RGB1.rgbBlue + .Blue * PercentInDecimal
   End With

End Function

Private Function GetAccessKey(Caption As String) As String

Dim intMax       As Integer
Dim intPosition  As Integer
Dim strCharacter As String * 1

   intMax = Len(Caption)
   
   If intMax < 2 Then Exit Function
   
   intMax = intMax - 1
   
   Do
      intPosition = InStrRev(Caption, "&", intMax)
      
      If intPosition = 0 Then Exit Do
      
      If intPosition = 1 Then
         GetAccessKey = Mid(Caption, intPosition + 1, 1)
         Exit Do
         
      Else
         strCharacter = Mid(Caption, intPosition - 1, 1)
         
         If StrComp(strCharacter, "&") Then
            GetAccessKey = Mid(Caption, intPosition + 1, 1)
            Exit Do
         End If
         
         intMax = intPosition - 2
      End If
   Loop While intMax > 0
   
   GetAccessKey = LCase$(GetAccessKey)

End Function

Private Function GetColor(ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long) As Long

   GetColor = Red + &H100& * Green + &H10000 * Blue

End Function

Private Function GetOSVersion() As Boolean

Const VER_PLATFORM_WIN32_NT As Long = 2

Dim osvInfo                 As OSVersionInfo

   osvInfo.dwOSVersionInfoSize = Len(osvInfo)
   
   If GetVersionEx(osvInfo) Then GetOSVersion = (osvInfo.dwPlatformId = VER_PLATFORM_WIN32_NT)

End Function

Private Function GetRGB(ByVal Color As Long) As ColorsRGB

   With GetRGB
      .Red = Color And 255
      .Green = (Color \ 256) And 255
      .Blue = (Color \ 65536) And 255
   End With

End Function

Private Function IsFunctionSupported(ByVal sFunction As String, ByVal sModule As String) As Boolean

Dim lngModule As Long

   lngModule = GetModuleHandle(sModule)
   
   If lngModule = 0 Then lngModule = LoadLibrary(sModule)
   
   If lngModule Then
      If GetProcAddress(lngModule, sFunction) Then IsFunctionSupported = True
      
      FreeLibrary lngModule
   End If

End Function

Private Function ShiftColor(ByVal Color As Long, ByVal Percent As Single) As Long

Dim RGBColor As ColorsRGB

   With RGBColor
      RGBColor = GetRGB(Color)
      .Red = .Red + Percent * 255
      .Green = .Green + Percent * 255
      .Blue = .Blue + Percent * 255
      
      If Percent > 0 Then
         If .Red > 255 Then .Red = 255
         If .Green > 255 Then .Green = 255
         If .Blue > 255 Then .Blue = 255
         
      Else
         If .Red < 0 Then .Red = 0
         If .Green < 0 Then .Green = 0
         If .Blue < 0 Then .Blue = 0
      End If
      
      ShiftColor = GetColor(.Red, .Green, .Blue)
   End With

End Function

Private Function TranslateColor(ByVal Value As Long) As Long

   OleTranslateColor Value, 0, TranslateColor

End Function

Private Function TranslateNumber(ByVal Value As Single, ByVal Minimum As Long, ByVal Maximum As Long) As Single

   If Value > Maximum Then
      TranslateNumber = Maximum
      
   ElseIf Value < Minimum Then
      TranslateNumber = Minimum
      
   Else
      TranslateNumber = Value
   End If

End Function

Private Sub CalculateRectangle()

Const BORDER_SPACE  As Long = 4
Const DT_CALCRECT   As Long = &H400
Const DT_CALCFLAG   As Long = DT_WORDBREAK Or DT_CALCRECT Or DT_NOCLIP Or DT_CENTER

Dim lngButtonHeight As Long
Dim lngButtonWidth  As Long
Dim lngTextLenght   As Long
Dim rctPicture      As Rect
Dim rctText         As Rect
Dim strText         As String

   CalculateRect = False
   
   With m_ButtonSettings
      lngButtonHeight = .Height
      lngButtonWidth = .Width
      strText = m_ButtonProperty.Caption
      lngTextLenght = Len(strText)
      SetRect .Focus, BORDER_SPACE, BORDER_SPACE, .Width - BORDER_SPACE, .Height - BORDER_SPACE
   End With
   
   If lngTextLenght > 0 Then
      SetRect rctText, 0, 0, lngButtonWidth, lngButtonHeight
      
      If IsNT Then
         DrawTextW hDC, StrPtr(strText), lngTextLenght, rctText, DT_CALCFLAG
         
      Else
         DrawText hDC, strText, lngTextLenght, rctText, DT_CALCFLAG
      End If
   End If
   
   OffsetRect rctText, (lngButtonWidth - rctText.Right) \ 2, (lngButtonHeight - rctText.Bottom) \ 2
   CopyRect m_ButtonSettings.Caption, rctText
   
   If m_ButtonProperty.Picture Is Nothing Then Exit Sub
   
   SetRect rctPicture, 0, 0, m_ButtonProperty.PicWidth, m_ButtonProperty.PicHeight
   
   If lngTextLenght > 0 Then
      lngTextLenght = 2
      
   Else
      lngTextLenght = 0
   End If
   
   Select Case m_ButtonProperty.PicAlign
      Case BehindText
         OffsetRect rctPicture, (lngButtonWidth - rctPicture.Right) \ 2, (lngButtonHeight - rctPicture.Bottom) \ 2
         
      Case BottomEdge, BottomOfCaption
         OffsetRect rctPicture, (lngButtonWidth - rctPicture.Right) \ 2, 0
         OffsetRect rctText, 0, -rctText.Top
         
         If m_ButtonProperty.PicAlign = BottomEdge Then
            OffsetRect rctPicture, 0, lngButtonHeight - rctPicture.Bottom - BORDER_SPACE
            
            If lngTextLenght > 0 Then
               OffsetRect rctText, 0, rctPicture.Top - rctText.Bottom
               lngTextLenght = rctText.Top - BORDER_SPACE
               
               If lngTextLenght > 1 Then OffsetRect rctText, 0, -(lngTextLenght \ 2)
            End If
            
         ElseIf lngTextLenght = 0 Then
            OffsetRect rctPicture, 0, (lngButtonHeight - rctPicture.Bottom) \ 2
            
         Else
            OffsetRect rctPicture, 0, rctText.Bottom
            lngTextLenght = (lngButtonHeight - rctPicture.Bottom) \ 2
            OffsetRect rctPicture, 0, lngTextLenght
            OffsetRect rctText, 0, lngTextLenght
         End If
         
   Case LeftEdge, LeftOfCaption
      OffsetRect rctPicture, 0, (lngButtonHeight - rctPicture.Bottom) \ 2
      OffsetRect rctText, -rctText.Left, 0
                                                              
      If m_ButtonProperty.PicAlign = LeftEdge Then
         OffsetRect rctPicture, BORDER_SPACE, 0
         
         If lngTextLenght > 0 Then
            OffsetRect rctText, rctPicture.Right + lngTextLenght, 0
            lngTextLenght = lngButtonWidth - BORDER_SPACE - rctText.Right
            
            If lngTextLenght > 1 Then OffsetRect rctText, lngTextLenght \ 2, 0
         End If
         
      ElseIf lngTextLenght = 0 Then
         OffsetRect rctPicture, (lngButtonWidth - rctPicture.Right) \ 2, 0
         
      Else
         OffsetRect rctText, rctPicture.Right + lngTextLenght, 0
         lngTextLenght = (lngButtonWidth - rctText.Right) \ 2
         OffsetRect rctText, lngTextLenght, 0
         OffsetRect rctPicture, lngTextLenght, 0
      End If
      
   Case RightEdge, RightOfCaption
      OffsetRect rctPicture, 0, (lngButtonHeight - rctPicture.Bottom) \ 2
      OffsetRect rctText, -rctText.Left, 0
      
      If m_ButtonProperty.PicAlign = RightEdge Then
         OffsetRect rctPicture, lngButtonWidth - rctPicture.Right - BORDER_SPACE, 0
         
         If lngTextLenght > 0 Then
            OffsetRect rctText, rctPicture.Left - rctText.Right - lngTextLenght, 0
            lngTextLenght = rctText.Left - BORDER_SPACE
            
            If lngTextLenght > 1 Then OffsetRect rctText, -(lngTextLenght \ 2), 0
         End If
         
      ElseIf lngTextLenght = 0 Then
         OffsetRect rctPicture, (lngButtonWidth - rctPicture.Right) \ 2, 0
         
      Else
         OffsetRect rctPicture, rctText.Right + lngTextLenght, 0
         lngTextLenght = (lngButtonWidth - rctPicture.Right) \ 2
         OffsetRect rctPicture, lngTextLenght, 0
         OffsetRect rctText, lngTextLenght, 0
      End If
      
   Case TopEdge, TopOfCaption
      OffsetRect rctPicture, (lngButtonWidth - rctPicture.Right) \ 2, 0
      OffsetRect rctText, 0, -rctText.Top
      
      If m_ButtonProperty.PicAlign = TopEdge Then
         OffsetRect rctPicture, 0, BORDER_SPACE
         
         If lngTextLenght > 0 Then
            OffsetRect rctText, 0, rctPicture.Bottom
            lngTextLenght = lngButtonHeight - rctText.Bottom - BORDER_SPACE
            
            If lngTextLenght > 1 Then OffsetRect rctText, 0, lngTextLenght \ 2
         End If
         
      ElseIf lngTextLenght = 0 Then
         OffsetRect rctPicture, 0, (lngButtonHeight - rctPicture.Bottom) \ 2
         OffsetRect rctText, 0, rctPicture.Bottom
         lngTextLenght = (lngButtonHeight - rctText.Bottom) \ 2
         OffsetRect rctText, 0, lngTextLenght
         OffsetRect rctPicture, 0, lngTextLenght
      End If
   End Select
   
   CopyRect m_ButtonSettings.Picture, rctPicture
   CopyRect m_ButtonSettings.Caption, rctText

End Sub

Private Sub CalculateRegionBorder(ByVal Region As Long, ByVal EllipseWidth As Long, ByVal EllipseHeight As Long)

Dim lngCount  As Long
Dim lngHeight As Long
Dim lngWidth  As Long
Dim lngX      As Long
Dim lngY      As Long

   lngHeight = m_ButtonSettings.Height
   lngWidth = m_ButtonSettings.Width
   EllipseHeight = TranslateNumber((EllipseHeight / 2), 0, lngHeight / 2)
   EllipseWidth = TranslateNumber((EllipseWidth / 2), 0, lngWidth / 2)
   
   If (m_ButtonProperty.Shape = ShapeLeft) Or (m_ButtonProperty.Shape = ShapeSides) Then
      ReDim m_BorderPoints(2) As PointAPI
      
      m_BorderPoints(lngCount).X = 0
      m_BorderPoints(lngCount).Y = 0
      lngCount = 1
      m_BorderPoints(lngCount).X = 0
      m_BorderPoints(lngCount).Y = lngHeight - 1
      lngCount = lngCount + 1
      
   Else
      For lngY = 0 To EllipseHeight - 1
         For lngX = EllipseWidth To -1 Step -1
            If PtInRegion(Region, lngX, lngY) = 0 Then
               ReDim Preserve m_BorderPoints(lngCount + 1) As PointAPI
               
               m_BorderPoints(lngCount).X = lngX + 1
               m_BorderPoints(lngCount).Y = lngY
               lngCount = lngCount + 1
               
               If lngX > 0 Then
                  m_BorderPoints(lngCount).X = lngX
                  m_BorderPoints(lngCount).Y = lngY + 1
                  lngCount = lngCount + 1
               End If
               
               lngX = -1
            End If
         Next 'lngX
      Next 'lngY
      
      For lngY = lngHeight - EllipseHeight To lngHeight - 1
         For lngX = EllipseWidth To -1 Step -1
            If PtInRegion(Region, lngX, lngY) = 0 Then
               ReDim Preserve m_BorderPoints(lngCount + 1) As PointAPI
               
               If lngX > 0 Then
                  m_BorderPoints(lngCount).X = lngX
                  m_BorderPoints(lngCount).Y = lngY - 1
                  lngCount = lngCount + 1
               End If
               
               m_BorderPoints(lngCount).X = lngX + 1
               m_BorderPoints(lngCount).Y = lngY
               lngCount = lngCount + 1
               lngX = -1
            End If
         Next 'lngX
      Next 'lngY
   End If
   
   If (m_ButtonProperty.Shape = ShapeRight) Or (m_ButtonProperty.Shape = ShapeSides) Then
      ReDim Preserve m_BorderPoints(lngCount + 2) As PointAPI
      
      m_BorderPoints(lngCount).X = lngWidth - 1
      m_BorderPoints(lngCount).Y = lngHeight - 1
      lngCount = lngCount + 1
      m_BorderPoints(lngCount).X = lngWidth - 1
      m_BorderPoints(lngCount).Y = 0
      lngCount = lngCount + 1
      
   Else
      For lngY = lngHeight - 1 To lngHeight - EllipseHeight + 1 Step -1
         For lngX = lngWidth - EllipseWidth To lngWidth - 1
            If PtInRegion(Region, lngX, lngY) = 0 Then
               ReDim Preserve m_BorderPoints(lngCount + 1) As PointAPI
               
               m_BorderPoints(lngCount).X = lngX - 1
               m_BorderPoints(lngCount).Y = lngY
               lngCount = lngCount + 1
               m_BorderPoints(lngCount).X = lngX
               m_BorderPoints(lngCount).Y = lngY - 1
               lngCount = lngCount + 1
               lngX = lngWidth
            End If
         Next 'lngX
      Next 'lngY
      
      For lngY = EllipseHeight To 0 Step -1
         For lngX = lngWidth - EllipseWidth To lngWidth
            If PtInRegion(Region, lngX, lngY) = 0 Then
               ReDim Preserve m_BorderPoints(lngCount + 1) As PointAPI
               
               m_BorderPoints(lngCount).X = lngX
               m_BorderPoints(lngCount).Y = lngY + 1
               lngCount = lngCount + 1
               m_BorderPoints(lngCount).X = lngX - 1
               m_BorderPoints(lngCount).Y = lngY
               lngCount = lngCount + 1
               lngX = lngWidth
            End If
         Next 'lngX
      Next 'lngY
   End If
   
   ReDim Preserve m_BorderPoints(lngCount) As PointAPI
   
   m_BorderPoints(lngCount) = m_BorderPoints(0)

End Sub

Private Sub CreateButtonRegion(ByVal EllipseWidth As Long, ByVal EllipseHeight As Long)

Dim lngRegion(1) As Long

   With m_ButtonProperty
      If .CornerAngle = CornerThin Then
         EllipseHeight = 6
         EllipseWidth = 6
         
      ElseIf .CornerAngle = CornerSmall Then
         EllipseHeight = 16
         EllipseWidth = 16
         
      ElseIf .CornerAngle = CornerMedium Then
         EllipseHeight = 40
         EllipseWidth = 40
         
      ElseIf .CornerAngle = CornerBig Then
         EllipseHeight = 52
         EllipseWidth = 52
      End If
   End With
   
   With m_ButtonSettings
      lngRegion(0) = CreateRoundRectRgn(0, 0, .Width + 1, .Height + 1, EllipseWidth, EllipseHeight)
      
      If (m_ButtonProperty.Shape = ShapeLeft) Or (m_ButtonProperty.Shape = ShapeSides) Then
         lngRegion(1) = CreateRectRgn(0, 0, .Width / 2, .Height + 1)
         CombineRgn lngRegion(0), lngRegion(0), lngRegion(1), RGN_OR
         DeleteObject lngRegion(1)
      End If
      
      If (m_ButtonProperty.Shape = ShapeRight) Or (m_ButtonProperty.Shape = ShapeSides) Then
         lngRegion(1) = CreateRectRgn(.Width / 2, 0, .Width + 1, .Height + 1)
         CombineRgn lngRegion(0), lngRegion(0), lngRegion(1), RGN_OR
         DeleteObject lngRegion(1)
      End If
   End With
   
   Call CalculateRegionBorder(lngRegion(0), EllipseWidth, EllipseHeight)
   
   SetWindowRgn hWnd, lngRegion(0), True
   DeleteObject lngRegion(0)
   Erase lngRegion

End Sub

Private Sub DrawButton(Optional ByVal State As ButtonStates = -1, Optional ByVal Force As Boolean)

Dim intReduce   As Integer
Dim lngColor(1) As Long
Dim lngOldPen   As Long
Dim lngPen      As Long
Dim lngX(1)     As Long
Dim lngY(1)     As Long

   If ControlHidden Then Exit Sub
   If State = -1 Then State = m_ButtonSettings.State
   If Not Force Then If State = m_ButtonSettings.State Then If m_ButtonSettings.HasFocus = ButtonHasFocus Then Exit Sub
   
   With m_ButtonSettings
      Cls
      .HasFocus = ButtonHasFocus
      .State = State
      
      If CalculateRect Then
         Call CreateButtonRegion(.Height, .Width)
         
         With m_ButtonProperty
            If .CornerAngle = CornerThin Then
               intReduce = 7
               
            ElseIf .CornerAngle = CornerSmall Then
               intReduce = 6
               
            ElseIf .CornerAngle = CornerMedium Then
               intReduce = 4
               
            ElseIf .CornerAngle = CornerBig Then
               intReduce = 3
            End If
         End With
         
         lngX(0) = .Width / 4 - intReduce
         
         If lngX(0) > 10 Then lngX(0) = 10 - intReduce
         
         lngX(1) = .Width - lngX(0) - 1
         lngY(0) = 1
         lngY(1) = .Height / 2
         
         If m_ButtonProperty.ShineFullLeft Then lngX(0) = 0
         
         Call CalculateRectangle
         
         SetRect .Focus, lngX(0), lngY(0), lngX(1), lngY(1)
      End If
   End With
   
   With m_ButtonProperty
      If .CheckBox And .Value And .Enabled Then
         If State = IsNormal Then
            State = IsDown
            
         ElseIf State = IsHot Then
            State = IsDown
            
         ElseIf State = IsDown Then
            State = IsDown
         End If
         
      ElseIf MouseIsDown And (State = IsNormal) Then
         State = IsHot
      End If
   End With
   
   With m_ButtonColors
      If State = IsNormal Then
         lngColor(0) = .BackColor
         
      ElseIf State = IsHot Then
         lngColor(0) = .HoverColor
         
      ElseIf State = IsDown Then
         lngColor(0) = .DownColor
         
      ElseIf State = IsDisabled Then
         lngColor(0) = .GrayColor
      End If
      
      If .StartColor = -1 Then
         lngColor(1) = BlendColors(lngColor(0), &HFFFFFF, 0.8)
         
      Else
         lngColor(1) = .StartColor
      End If
      
      Call DrawGradientEx(lngColor(0), lngColor(1), 60)
      Call DrawShineEffect(lngColor(1), BlendColors(lngColor(0), lngColor(1), 0.6))
      
      lngColor(1) = ShiftColor(lngColor(0), 0.2)
   End With
   
   With m_ButtonSettings
      If .State = IsDisabled Then
         If m_ButtonProperty.Value Then
            Call DrawIcon(1, 1)
            Call DrawCaption(lngColor(1), 2, 2)
            Call DrawCaption(m_ButtonColors.GrayText, 1, 1)
            
         Else
            m_ButtonProperty.PicOpacity = 0.2
            
            Call DrawIcon
            Call DrawCaption(lngColor(1), 1, 1)
            Call DrawCaption(m_ButtonColors.GrayText)
            
            m_ButtonProperty.PicOpacity = 1
         End If
         
      Else
         If (.State = IsDown) Or m_ButtonProperty.Value Then
            If (.State = IsDown) And m_ButtonProperty.Value Then
               Call DrawIconEffect(lngColor(1), 2, 2)
               Call DrawCaptionEffect(m_ButtonColors.ForeColor, lngColor(1), 2, 2)
               
            Else
               Call DrawIconEffect(lngColor(1), 1, 1)
               Call DrawCaptionEffect(m_ButtonColors.ForeColor, lngColor(1), 1, 1)
            End If
            
         Else
            Call DrawIconEffect(lngColor(1))
            Call DrawCaptionEffect(m_ButtonColors.ForeColor, lngColor(1))
         End If
      End If
      
      lngPen = CreatePen(PS_SOLID, 1, BlendColors(lngColor(0), lngColor(1), 0.8))
      lngOldPen = SelectObject(hDC, lngPen)
      Polyline hDC, m_BorderPoints(0), UBound(m_BorderPoints) + 1
      SelectObject hDC, lngOldPen
      DeleteObject lngPen
   End With
   
   Erase lngColor, lngX, lngY

End Sub

Private Sub DrawCaption(ByVal Color As Long, Optional ByVal MoveX As Long, Optional ByVal MoveY As Long)

Const DT_DRAWFLAG As Long = DT_WORDBREAK Or DT_NOCLIP Or DT_CENTER

Dim lngTextLenght As Long
Dim rctText       As Rect
Dim strText       As String

   strText = m_ButtonProperty.Caption
   lngTextLenght = Len(strText)
   
   If lngTextLenght = 0 Then Exit Sub
   
   CopyRect rctText, m_ButtonSettings.Caption
   
   If MoveX Or MoveY Then OffsetRect rctText, MoveX, MoveY
                                              
   SetTextColor hDC, Color
   
   If IsNT Then
      DrawTextW hDC, StrPtr(strText), lngTextLenght, rctText, DT_DRAWFLAG
      
   Else
      DrawText hDC, strText, lngTextLenght, rctText, DT_DRAWFLAG
   End If

End Sub

Private Sub DrawCaptionEffect(ByVal TextColor As Long, ByVal BackColor As Long, Optional MoveX As Long, Optional MoveY As Long)

   Call DrawCaption(TextColor, MoveX, MoveY)

End Sub

Private Sub DrawGradientEx(ByVal StartColor As Long, ByVal EndColor As Long, Optional ByVal Center As Single = 50, Optional ByVal X1 As Long, Optional ByVal Y1 As Long, Optional ByVal X2 As Long = -1, Optional ByVal Y2 As Long = -1)

Dim RGBColor(3) As ColorsRGB
Dim lngColor    As Long
Dim sngStep     As Single

   If X2 = -1 Then X2 = m_ButtonSettings.Width - 1
   If Y2 = -1 Then Y2 = m_ButtonSettings.Height - 1
   
   X2 = TranslateNumber(X2, X1, X2)
   Y2 = TranslateNumber(Y2, Y1, Y2)
   Center = TranslateNumber(Center, 0, 100)
   RGBColor(0) = GetRGB(StartColor)
   RGBColor(1) = GetRGB(EndColor)
   
   With RGBColor(0)
      RGBColor(2).Red = .Red + (RGBColor(1).Red - .Red) * 0.5
      RGBColor(2).Green = .Green + (RGBColor(1).Green - .Green) * 0.5
      RGBColor(2).Blue = .Blue + (RGBColor(1).Blue - .Blue) * 0.5
   End With
   
   Center = (Y2 - Y1 - 1) * Center / 100
   Center = (Y1 + Center)
   
   If Center = 0 Then Center = 1
   
   For Y1 = 0 To Y2
      If Y1 <= Center Then
         With RGBColor(0)
            sngStep = Y1 / Center
            RGBColor(3).Red = .Red + (RGBColor(2).Red - .Red) * sngStep
            RGBColor(3).Green = .Green + (RGBColor(2).Green - .Green) * sngStep
            RGBColor(3).Blue = .Blue + (RGBColor(2).Blue - .Blue) * sngStep
         End With
         
      Else
         With RGBColor(2)
            sngStep = (Y1 - Center) / (Y2 - Center)
            RGBColor(3).Red = .Red + (RGBColor(1).Red - .Red) * sngStep
            RGBColor(3).Green = .Green + (RGBColor(1).Green - .Green) * sngStep
            RGBColor(3).Blue = .Blue + (RGBColor(1).Blue - .Blue) * sngStep
         End With
      End If
      
      With RGBColor(3)
         lngColor = GetColor(.Red, .Green, .Blue)
      End With
      
      If X1 = X2 Then
         SetPixelV hDC, X1, Y1, lngColor
         
      Else
         Call DrawLine(lngColor, X1, Y1, X2, Y1)
      End If
   Next 'Y1
   
   Erase RGBColor

End Sub

Private Sub DrawIcon(Optional ByVal MoveX As Long, Optional ByVal MoveY As Long, Optional ByVal BrushColor As Long = -1)

Const BI_RGB         As Long = 0
Const DI_NORMAL      As Long = &H3
Const DIB_RGB_COLORS As Long = 0

Dim bmiBitmap        As BitmapInfo
Dim RGBColor         As ColorsRGB
Dim lngBackBMP       As Long
Dim lngBackDC        As Long
Dim lngBackObject    As Long
Dim lngBrush         As Long
Dim lngCount         As Long
Dim lngDstBMP        As Long
Dim lngDstDC         As Long
Dim lngDstObject     As Long
Dim lngMask          As Long
Dim lngOpacity       As Long
Dim lngPicBMP        As Long
Dim lngPicDC         As Long
Dim lngPicHeight     As Long
Dim lngPicObject     As Long
Dim lngPicWidth      As Long
Dim lngSrcBMP        As Long
Dim lngSrcDC         As Long
Dim lngSrcObject     As Long
Dim lngDrawHeight    As Long
Dim lngDrawWidth     As Long
Dim lngX             As Long
Dim lngY             As Long
Dim ptaPicture       As PointAPI
Dim rctPicture       As Rect

   If m_ButtonProperty.Picture Is Nothing Then Exit Sub
   
   CopyRect rctPicture, m_ButtonSettings.Picture
   
   If MoveX Or MoveY Then OffsetRect rctPicture, MoveX, MoveY
   
   With rctPicture
      If m_ButtonProperty.PicSize = Size32x32 Then
         If .Left < 0 Then
            ptaPicture.X = -.Left
            .Left = 0
         End If
         
         If .Top < 0 Then
            ptaPicture.Y = -.Top
            .Top = 0
         End If
         
         If .Bottom > m_ButtonSettings.Height Then .Bottom = m_ButtonSettings.Height
         If .Right > m_ButtonSettings.Width Then .Right = m_ButtonSettings.Width
      End If
      
      lngDrawHeight = .Bottom - .Top
      lngDrawWidth = .Right - .Left
   End With
   
   If (lngDrawHeight < 1) Or (lngDrawWidth < 1) Then Exit Sub
   
   With m_ButtonProperty.Picture
      lngPicWidth = ScaleX(.Width, vbHimetric, vbPixels)
      lngPicHeight = ScaleY(.Height, vbHimetric, vbPixels)
      lngSrcDC = CreateCompatibleDC(hDC)
      
      If (.Type = vbPicTypeBitmap) Or ((.Type > vbPicTypeBitmap) And Not m_ButtonProperty.UseMask) Then lngSrcObject = SelectObject(lngSrcDC, .Handle)
      
      If m_ButtonProperty.UseMask Then
         lngMask = m_ButtonColors.MaskColor
         
      ElseIf .Type > vbPicTypeBitmap Then
         lngMask = GetPixel(lngSrcDC, 0, 0)
         DeleteObject SelectObject(lngSrcDC, lngSrcObject)
         
      Else
         lngMask = -1
      End If
      
      If .Type > vbPicTypeBitmap Then
         lngSrcBMP = CreateCompatibleBitmap(hDC, lngPicWidth, lngPicHeight)
         lngSrcObject = SelectObject(lngSrcDC, lngSrcBMP)
         lngBrush = CreateSolidBrush(lngMask)
         DrawIconEx lngSrcDC, 0, 0, .Handle, lngPicWidth, lngPicHeight, 0, lngBrush, DI_NORMAL
         DeleteObject lngBrush
      End If
   End With
   
   lngBackDC = CreateCompatibleDC(lngSrcDC)
   lngDstDC = CreateCompatibleDC(lngSrcDC)
   lngPicDC = CreateCompatibleDC(lngSrcDC)
   lngBackBMP = CreateCompatibleBitmap(hDC, lngDrawWidth, lngDrawHeight)
   lngDstBMP = CreateCompatibleBitmap(hDC, lngDrawWidth, lngDrawHeight)
   lngPicBMP = CreateCompatibleBitmap(hDC, lngDrawWidth, lngDrawHeight)
   lngBackObject = SelectObject(lngBackDC, lngBackBMP)
   lngDstObject = SelectObject(lngDstDC, lngDstBMP)
   lngPicObject = SelectObject(lngPicDC, lngPicBMP)
   
   With ptaPicture
      If (.X > 0) Or (rctPicture.Right = m_ButtonSettings.Width) Then lngPicWidth = lngDrawWidth + .X
      If (.Y > 0) Or (rctPicture.Bottom = m_ButtonSettings.Height) Then lngPicHeight = lngDrawHeight + .Y
      
      StretchBlt lngDstDC, 0, 0, lngDrawWidth, lngDrawHeight, lngSrcDC, .X, .Y, lngPicWidth - .X, lngPicHeight - .Y, vbSrcCopy
   End With
   
   If m_ButtonProperty.Picture.Type <> vbPicTypeBitmap Then DeleteObject SelectObject(lngSrcDC, lngSrcObject)
   
   DeleteDC lngSrcDC
   
   ReDim rgbBackground(lngDrawWidth * lngDrawHeight * 1.5) As RGBQuad
   ReDim rgbPicture(UBound(rgbBackground)) As RGBQuad
   
   BitBlt lngBackDC, 0, 0, lngDrawWidth, lngDrawHeight, hDC, rctPicture.Left, rctPicture.Top, vbSrcCopy
   BitBlt lngPicDC, 0, 0, lngDrawWidth, lngDrawHeight, lngDstDC, 0, 0, vbSrcCopy
   
   With bmiBitmap.bmiHeader
      .biBitCount = 24
      .biCompression = BI_RGB
      .biHeight = lngDrawHeight
      .biPlanes = 1
      .biSize = Len(bmiBitmap.bmiHeader)
      .biWidth = lngDrawWidth
   End With
   
   GetDIBits lngBackDC, lngBackBMP, 0, lngDrawHeight, rgbBackground(0), bmiBitmap, DIB_RGB_COLORS
   GetDIBits lngPicDC, lngPicBMP, 0, lngDrawHeight, rgbPicture(0), bmiBitmap, DIB_RGB_COLORS
   DeleteObject SelectObject(lngBackDC, lngBackObject)
   DeleteObject SelectObject(lngPicDC, lngPicObject)
   DeleteDC lngBackDC
   DeleteDC lngPicDC
   
   If BrushColor > -1 Then RGBColor = GetRGB(BrushColor)
   
   If m_ButtonSettings.State = IsDisabled Then
      lngOpacity = m_ButtonProperty.PicOpacity
      m_ButtonProperty.PicOpacity = 0.2
   End If
   
   If (lngMask <> -1) Or (BrushColor <> -1) Or (m_ButtonProperty.PicOpacity <> 1) Then
      For lngY = 0 To lngDrawHeight - 1
         For lngX = 0 To lngDrawWidth - 1
            With rgbPicture(lngCount)
               If GetNearestColor(lngDstDC, GetColor(.rgbRed, .rgbGreen, .rgbBlue)) = lngMask Then
                  rgbPicture(lngCount) = rgbBackground(lngCount)
                  
               Else
                  If BrushColor > -1 Then
                     .rgbRed = RGBColor.Red
                     .rgbGreen = RGBColor.Green
                     .rgbBlue = RGBColor.Blue
                     
                  ElseIf m_ButtonProperty.PicOpacity <> 1 Then
                     rgbPicture(lngCount) = BlendRGBQuad(rgbBackground(lngCount), rgbPicture(lngCount), m_ButtonProperty.PicOpacity)
                  End If
               End If
            End With
            
            lngCount = lngCount + 1
         Next 'lngX
      Next 'lngY
   End If
   
   If m_ButtonSettings.State = IsDisabled Then m_ButtonProperty.PicOpacity = lngOpacity
   
   DeleteObject SelectObject(lngDstDC, lngDstObject)
   DeleteDC lngDstDC
   SetDIBitsToDevice hDC, rctPicture.Left, rctPicture.Top, lngDrawWidth, lngDrawHeight, 0, 0, 0, lngDrawHeight, rgbPicture(0), bmiBitmap, DIB_RGB_COLORS
   Erase rgbPicture, rgbBackground

End Sub

Private Sub DrawIconEffect(ByVal BackColor As Long, Optional ByVal MoveX As Long, Optional ByVal MoveY As Long, Optional BrushColor As Long = -1)

   If m_ButtonProperty.Picture Is Nothing Then Exit Sub
   
   Call DrawIcon(MoveX, MoveY, BrushColor)

End Sub

Private Sub DrawLine(ByVal Color As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)

Dim lngOldPen As Long
Dim lngPen    As Long
Dim ptaTemp   As PointAPI

   lngPen = CreatePen(PS_SOLID, 1, Color)
   lngOldPen = SelectObject(hDC, lngPen)
   
   If X1 = X2 Then Y2 = Y2 + 1
   If Y1 = Y2 Then X2 = X2 + 1
   
   MoveToEx hDC, X1, Y1, ptaTemp
   LineTo hDC, X2, Y2
   SelectObject hDC, lngOldPen
   DeleteObject lngPen

End Sub

Private Sub DrawShineEffect(ByVal StartColor As Long, ByVal EndColor As Long)

Dim RGBColor(2)   As ColorsRGB
Dim intCorner     As Integer
Dim lngCount      As Long
Dim lngGradient() As Long
Dim lngHeight     As Long
Dim lngLeft       As Long
Dim lngRegion(1)  As Long
Dim lngTop        As Long
Dim lngWidth      As Long
Dim lngX(2)       As Long
Dim lngY          As Long
Dim sngPercent    As Single

   With m_ButtonSettings.Focus
      lngHeight = .Bottom - .Top + 1
      lngWidth = .Right - .Left + 1
      lngLeft = .Left
      lngTop = .Top
      
      If lngHeight = 0 Then Exit Sub
      
      If (m_ButtonProperty.Shape = ShapeLeft) Or (m_ButtonProperty.Shape = ShapeSides) Then
         lngWidth = lngWidth + lngLeft
         lngLeft = 0
      End If
      
      If (m_ButtonProperty.Shape = ShapeRight) Or (m_ButtonProperty.Shape = ShapeSides) Then lngWidth = lngWidth + .Left
   End With
   
   RGBColor(0) = GetRGB(StartColor)
   RGBColor(1) = GetRGB(EndColor)
   
   ReDim lngGradient(lngHeight - 1) As Long
   
   For lngCount = 0 To lngHeight - 1
      With RGBColor(2)
         sngPercent = lngCount / lngHeight
         .Red = RGBColor(0).Red + (RGBColor(1).Red - RGBColor(0).Red) * sngPercent
         .Green = RGBColor(0).Green + (RGBColor(1).Green - RGBColor(0).Green) * sngPercent
         .Blue = RGBColor(0).Blue + (RGBColor(1).Blue - RGBColor(0).Blue) * sngPercent
         lngGradient(lngCount) = GetColor(.Red, .Green, .Blue)
      End With
   Next 'lngCount
   
   With m_ButtonProperty
      If .CornerAngle = CornerThin Then
         intCorner = 6
         
      ElseIf .CornerAngle = CornerSmall Then
         intCorner = 5
         
      ElseIf .CornerAngle = CornerMedium Then
         intCorner = 3
         
      ElseIf .CornerAngle = CornerBig Then
         intCorner = 2
         
      ' Full corner
      Else
         intCorner = 1
      End If
      
      If .ShineFullRight Then lngWidth = m_ButtonSettings.Focus.Right
      
      lngRegion(0) = CreateRoundRectRgn(0, 0, lngWidth, lngHeight, lngHeight \ intCorner, lngWidth \ intCorner)
      
      If .ShineFullLeft Or (.Shape = ShapeLeft) Or (.Shape = ShapeSides) Then
         lngRegion(1) = CreateRoundRectRgn(0, 0, lngWidth / 2, lngHeight, 0, 0)
         CombineRgn lngRegion(0), lngRegion(0), lngRegion(1), RGN_OR
         DeleteObject lngRegion(1)
      End If
      
      If .ShineFullRight Or (.Shape = ShapeRight) Or (.Shape = ShapeSides) Then
         lngRegion(1) = CreateRoundRectRgn(lngWidth / 2, 0, lngWidth, lngHeight, 0, 0)
         CombineRgn lngRegion(0), lngRegion(0), lngRegion(1), RGN_OR
         DeleteObject lngRegion(1)
      End If
   End With
   
   For lngY = 0 To lngHeight - 1
      lngX(0) = -1
      lngX(1) = -1
      
      For lngX(2) = 0 To lngWidth - 1
         If PtInRegion(lngRegion(0), lngX(2), lngY) Then
            If lngX(0) = -1 Then lngX(0) = lngX(2)
            
         ElseIf (lngX(0) <> -1) And (lngX(1) = -1) Then
            lngX(1) = lngX(2)
            lngX(2) = lngWidth
         End If
      Next 'lngX(2)
      
      If (lngX(0) <> -1) And (lngX(1) <> -1) Then Call DrawLine(lngGradient(lngY), lngLeft + lngX(0), lngTop + lngY, lngLeft + lngX(1), lngTop + lngY)
   Next 'lngY
   
   DeleteObject lngRegion(0)
   Erase RGBColor, lngGradient, lngRegion, lngX

End Sub

Private Sub PlaySound()

Const SND_ASYNC            As Long = &H1
Const SND_MEMORY           As Long = &H4
Const SND_NODEFAULT        As Long = &H2
Const SOUND_BUTTON_CLICKED As Long = 1

Dim strSoundBuffer         As String

   If Not m_ButtonProperty.Sound Then Exit Sub
   
   On Local Error Resume Next
   strSoundBuffer = StrConv(LoadResData(SOUND_BUTTON_CLICKED, "Sounds"), vbUnicode)
   SoundPlay strSoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
   On Local Error GoTo 0
   DoEvents

End Sub

Private Sub SendKey(ByVal TabKey As String)

Const INPUT_KEYBOARD  As Long = 1
Const KEYEVENTF_KEYUP As Long = 2

Dim intCount          As Integer
Dim intKeys           As Integer

   If Left(TabKey, 1) = "+" Then
      ReDim intKey(1) As Integer
      
      intKey(0) = vbKeyShift
      intKey(1) = vbKeyTab
      intKeys = 3
      
   Else
      ReDim intKey(0) As Integer
      
      intKey(0) = vbKeyTab
      intKeys = 1
   End If
   
   ReDim kbiEvents(intKeys) As KeyboardInput
   
   For intCount = 0 To intKeys
      With kbiEvents(intCount)
         .dwType = INPUT_KEYBOARD
         
         If intCount >= intKeys / 2 Then
            .wVk = intKey(intKeys - intCount)
            .dwFlags = KEYEVENTF_KEYUP
            
         Else
            .wVk = intKey(intCount)
         End If
      End With
   Next 'intCount
   
   SendInput intCount, kbiEvents(0), Len(kbiEvents(0))
   Erase intKey, kbiEvents

End Sub

Private Sub SetButtonColors()

   With m_ButtonColors
      .BackColor = TranslateColor(m_ButtonProperty.BackColor)
      .DownColor = ShiftColor(.BackColor, -0.15)
      .BackColor = ShiftColor(.BackColor, -0.05)
      .ForeColor = TranslateColor(m_ButtonProperty.ForeColor)
      .GrayColor = .BackColor
      .GrayText = BlendColors(.ForeColor, &HFFFFFF, 0.6)
      .HoverColor = ShiftColor(.BackColor, 0.1)
      .MaskColor = TranslateColor(m_ButtonProperty.MaskColor)
      .StartColor = -1
   End With

End Sub

Private Sub SetPictureSize()

   With m_ButtonProperty
      Select Case .PicSize
         Case SizePicture
            If Not m_ButtonProperty.Picture Is Nothing Then
               .PicWidth = ScaleX(m_ButtonProperty.Picture.Width, vbHimetric, vbPixels)
               .PicHeight = ScaleY(m_ButtonProperty.Picture.Height, vbHimetric, vbPixels)
            End If
            
         Case Size16x16
            .PicWidth = 16
            .PicHeight = 16
            
         Case Size24x24
            .PicWidth = 24
            .PicHeight = 24
            
         Case Size32x32
            .PicWidth = 32
            .PicHeight = 32
            
         Case Size48x48
            .PicWidth = 48
            .PicHeight = 48
            
         Case Size64x64
            .PicWidth = 64
            .PicHeight = 64
      End Select
   End With

End Sub

Private Sub TrackMouseTracking(ByVal hWnd As Long)

Const TME_LEAVE As Long = &H2

Dim tmeMouse    As TrackMouseEventType

   With tmeMouse
      .cbSize = Len(tmeMouse)
      .dwFlags = TME_LEAVE
      .hwndTrack = hWnd
   End With
   
   If TrackUser32 Then
      TrackMouseEvent tmeMouse
      
   Else
      TrackMouseEventComCtl tmeMouse
   End If

End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)

   If m_ButtonProperty.Enabled Then
      If SpacebarIsDown Then If GetCapture = UserControl.hWnd Then ReleaseCapture
      If m_ButtonProperty.CheckBox Then If (KeyAscii = vbKeyReturn) Or (KeyAscii = vbKeyEscape) Then Exit Sub
      
      ButtonIsDown = False
      m_ButtonSettings.Button = vbLeftButton
      
      Call UserControl_Click
   End If

End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)

   m_ButtonSettings.AsDefault = Ambient.DisplayAsDefault
   
   If StrComp(PropertyName, "DisplayAsDefault") = 0 Then
      ButtonIsDown = False
      MouseIsDown = False
      SpacebarIsDown = False
      
      If m_ButtonProperty.Enabled And Not MouseOnButton Then If Not ButtonHasFocus Then Call DrawButton(Force:=True)
   End If

End Sub

Private Sub UserControl_Click()

   If m_ButtonProperty.Blocked Or ButtonIsDown Or (m_ButtonSettings.Button <> vbLeftButton) Then Exit Sub
   
   MouseIsDown = False
   SpacebarIsDown = False
   
   With m_ButtonProperty
      If .CheckBox Then .Value = Not .Value
      
      If Not MouseOnButton Then
         Call DrawButton(IsNormal, True)
         
      Else
         Call DrawButton(IsHot, .CheckBox)
      End If
      
      Call PlaySound
      
      If Not .Blocked Then RaiseEvent Click
      If Not .CheckBox Then .Value = False
   End With
    
End Sub

Private Sub UserControl_DblClick()

   If m_ButtonProperty.Blocked Then Exit Sub
   If m_ButtonProperty.HandPointer Then SetCursor m_ButtonSettings.Cursor
   
   If m_ButtonSettings.Button = vbLeftButton Then
      ButtonIsDown = True
      MouseIsDown = True
      
      Call DrawButton(IsDown)
      
      m_ButtonSettings.Button = 8
       
      If GetCapture <> UserControl.hWnd Then SetCapture UserControl.hWnd
      
      Call PlaySound
      
      RaiseEvent DblClick
   End If

End Sub

Private Sub UserControl_GotFocus()

   ButtonHasFocus = True
   
   If Not ButtonIsDown Then Call DrawButton(IsNormal)

End Sub

Private Sub UserControl_Hide()

   ControlHidden = True

End Sub

Private Sub UserControl_Initialize()

   IsNT = GetOSVersion

End Sub

Private Sub UserControl_InitProperties()

   With m_ButtonProperty
      .BackColor = &HBA9EA0
      .Caption = Ambient.DisplayName
      .CornerAngle = CornerFull
      .Enabled = True
       UserControl.Font = Ambient.Font
      .ForeColor = Ambient.ForeColor
      .MaskColor = &HC0C0C0
      .PicAlign = LeftOfCaption
      .PicOpacity = 1
      .PicSize = SizePicture
      .UseMask = True
   End With
   
   Call SetButtonColors
   
   RedrawOnResize = True

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

   If m_ButtonProperty.Blocked Then Exit Sub
   
   Select Case KeyCode
      Case vbKeySpace
         If Shift <> vbAltMask Then
            ButtonIsDown = True
            SpacebarIsDown = True
            
            If GetCapture <> UserControl.hWnd Then SetCapture UserControl.hWnd
            If Not MouseIsDown Then Call DrawButton(IsDown)
         End If
         
         RaiseEvent KeyDown(KeyCode, Shift)
         
      Case vbKeyLeft, vbKeyUp, vbKeyRight, vbKeyDown
         If Shift = vbDefault Then
            If (KeyCode = vbKeyLeft) Or (KeyCode = vbKeyUp) Then
               Call SendKey("+{TAB}")
               
            Else
               Call SendKey("{TAB}")
            End If
         End If
         
         If SpacebarIsDown Then
            If GetCapture = UserControl.hWnd Then ReleaseCapture
            
            DoEvents
            m_ButtonSettings.Button = vbLeftButton
            
            Call UserControl_Click
         End If
          
       Case Else
         If SpacebarIsDown Then
            ButtonIsDown = False
            SpacebarIsDown = False
            
            If GetCapture = UserControl.hWnd Then ReleaseCapture
            
            If MouseOnButton Then
               Call DrawButton(IsHot)
               
            Else
               Call DrawButton(IsNormal)
            End If
         End If
         
         RaiseEvent KeyDown(KeyCode, Shift)
   End Select

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

   If m_ButtonProperty.Blocked Then Exit Sub
   
   RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

   If m_ButtonProperty.Blocked Then
      Exit Sub
      
   ElseIf KeyCode = vbKeySpace Then
      ButtonIsDown = MouseIsDown
      
      If SpacebarIsDown Then
         SpacebarIsDown = False
         m_ButtonSettings.Button = vbLeftButton
         
         Call UserControl_Click
      End If
      
      If ButtonIsDown Then
         If GetCapture <> UserControl.hWnd Then SetCapture UserControl.hWnd
         
      Else
         If GetCapture = UserControl.hWnd Then ReleaseCapture
      End If
      
   Else
      Call PlaySound
   End If
   
   RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub UserControl_LostFocus()

   ButtonHasFocus = False
   ButtonIsDown = False
   MouseIsDown = False
   SpacebarIsDown = False
   
   If m_ButtonProperty.Enabled Then If ParentActive Then Call DrawButton(IsNormal, True)

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If m_ButtonProperty.Blocked Then Exit Sub
   If m_ButtonProperty.HandPointer Then SetCursor m_ButtonSettings.Cursor
   
   m_ButtonSettings.Button = Button
                                         
   If Button = vbLeftButton Then
      ButtonHasFocus = True
      ButtonIsDown = True
      MouseIsDown = True
      
      If Not SpacebarIsDown Then
         Call DrawButton(IsDown)
         
      ElseIf Not MouseOnButton Then
         MouseIsDown = False
         
         Call DrawButton(IsNormal)
         
         MouseIsDown = True
      End If
   End If
   
   RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim ptaMouse As PointAPI

   If m_ButtonProperty.Blocked Then
      Exit Sub
      
   ElseIf m_ButtonProperty.HandPointer Then
      SetCursor m_ButtonSettings.Cursor
      
   Else
      UserControl.MousePointer = UserControl.MousePointer
   End If
   
   GetCursorPos ptaMouse
   
   If WindowFromPoint(ptaMouse.X, ptaMouse.Y) <> UserControl.hWnd Then
      If MouseOnButton Then
         MouseOnButton = False
         
         If Not SpacebarIsDown And Not MouseIsDown Then Call DrawButton(IsNormal)
         
         RaiseEvent MouseLeave
         
      ElseIf MouseIsDown Then
         Call DrawButton(IsHot)
      End If
      
   Else
      MouseOnButton = True
      
      If Not SpacebarIsDown Then
         If ButtonIsDown Then
            Call DrawButton(IsDown)
            
         Else
            Call DrawButton(IsHot)
         End If
         
      ElseIf MouseIsDown Then
         Call DrawButton(IsDown)
      End If
      
      If Not IsTracking Then
         IsTracking = True
         
         Call TrackMouseTracking(UserControl.hWnd)
         
         RaiseEvent MouseEnter
         
      Else
         RaiseEvent MouseMove(Button, Shift, X, Y)
      End If
   End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If m_ButtonProperty.Blocked Then Exit Sub
   If m_ButtonProperty.HandPointer And MouseOnButton Then SetCursor m_ButtonSettings.Cursor
   
   If Button = vbLeftButton Then
      MouseIsDown = False
      ButtonIsDown = SpacebarIsDown
      
      If SpacebarIsDown And Not MouseOnButton Then
         Call DrawButton(IsDown)
         
      ElseIf m_ButtonSettings.Button = 8 Then
         If MouseOnButton Then
            Call DrawButton(IsHot)
            
         Else
            Call DrawButton(IsNormal)
         End If
      End If
      
      If GetCapture = UserControl.hWnd Then ReleaseCapture
      
      Call PlaySound
      
      RaiseEvent MouseUp(Button, Shift, X, Y)
   End If

End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)

   ' not used only send the event
   RaiseEvent OLECompleteDrag(Effect)

End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

   ' not used only send the event
   RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)

End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)

   ' not used only send the event
   RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)

End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)

   ' not used only send the event
   RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)

End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)

   ' not used only send the event
   RaiseEvent OLESetData(Data, DataFormat)

End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)

   ' not used only send the event
   RaiseEvent OLEStartDrag(Data, AllowedEffects)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

Const IDC_HAND As Long = 32649

   With PropBag
      m_ButtonProperty.BackColor = .ReadProperty("BackColor", &HBA9EA0)
      m_ButtonProperty.Blocked = .ReadProperty("Blocked", False)
      m_ButtonProperty.Caption = .ReadProperty("Caption", Ambient.DisplayName)
      m_ButtonProperty.CheckBox = .ReadProperty("CheckBox", False)
      m_ButtonProperty.Enabled = .ReadProperty("Enabled", True)
      m_ButtonProperty.CornerAngle = .ReadProperty("CornerAngle", CornerFull)
      m_ButtonProperty.ForeColor = .ReadProperty("ForeColor", Ambient.ForeColor)
      m_ButtonProperty.HandPointer = .ReadProperty("HandPointer", False)
      m_ButtonProperty.MaskColor = .ReadProperty("MaskColor", &HC0C0C0)
      Set UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
      UserControl.MousePointer = .ReadProperty("MousePointer", vbDefault)
      UserControl.OLEDropMode = .ReadProperty("OLEDropMode", odNone)
      Set m_ButtonProperty.Picture = .ReadProperty("Picture", Nothing)
      m_ButtonProperty.PicAlign = .ReadProperty("PicAlign", LeftOfCaption)
      m_ButtonProperty.PicSize = .ReadProperty("PicSize", SizePicture)
      m_ButtonProperty.PicOpacity = .ReadProperty("PicOpacity", 1)
      m_ButtonProperty.Shape = .ReadProperty("Shape", ShapeNone)
      m_ButtonProperty.ShineFullLeft = .ReadProperty("ShineLeft", False)
      m_ButtonProperty.ShineFullRight = .ReadProperty("ShineRight", False)
      m_ButtonProperty.Sound = .ReadProperty("Sound", False)
      m_ButtonProperty.UseMask = .ReadProperty("UseMask", True)
      m_ButtonProperty.Value = .ReadProperty("Value", False)
      Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
   End With
   
   Call SetPictureSize
   
   AccessKeys = GetAccessKey(m_ButtonProperty.Caption)
   UserControl.Enabled = m_ButtonProperty.Enabled
   
   If Ambient.UserMode Then
      If m_ButtonProperty.HandPointer Then
         m_ButtonSettings.Cursor = LoadCursor(0, IDC_HAND)
         m_ButtonProperty.HandPointer = m_ButtonSettings.Cursor
      End If
      
      Call Subclass_Initialize(hWnd)
      
      TrackUser32 = IsFunctionSupported("TrackMouseEvent", "User32")
      
      If TrackUser32 Or IsFunctionSupported("TrackMouseEvent", "User32") Then Call Subclass_AddMsg(hWnd, WM_MOUSELEAVE)
   End If
   
   CalculateRect = True
   RedrawOnResize = True
   
   Call SetButtonColors
   Call DrawButton(Force:=True)

End Sub

Private Sub UserControl_Resize()

Const MIN_HPX As Long = 15
Const MIN_WPX As Long = 15

Dim lpRect As Rect

   With m_ButtonSettings
      GetClientRect UserControl.hWnd, lpRect
      .Height = lpRect.Bottom
      .Width = lpRect.Right
      
      If (.Height < MIN_HPX) Or (.Width < MIN_WPX) Then
          If .Height < MIN_HPX Then Height = MIN_HPX * Screen.TwipsPerPixelY
          If .Width < MIN_WPX Then Width = MIN_WPX * Screen.TwipsPerPixelX
          
          Exit Sub
      End If
      
      CalculateRect = True
      
      If Ambient.UserMode Then
         Call DrawButton(Force:=True)
         
      ElseIf RedrawOnResize Then
         Call DrawButton(Force:=True)
      End If
   End With

End Sub

Private Sub UserControl_Show()

   ControlHidden = False

End Sub

Private Sub UserControl_Terminate()

   If m_ButtonProperty.HandPointer Then DeleteObject m_ButtonSettings.Cursor
   
   On Local Error GoTo ExitSub
   
   Call Subclass_Terminate
   
ExitSub:
   On Local Error GoTo 0
   Erase SubclassData

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "BackColor", m_ButtonProperty.BackColor, &HBA9EA0
      .WriteProperty "Blocked", m_ButtonProperty.Blocked, False
      .WriteProperty "Caption", m_ButtonProperty.Caption, Ambient.DisplayName
      .WriteProperty "CheckBox", m_ButtonProperty.CheckBox, False
      .WriteProperty "CornerAngle", m_ButtonProperty.CornerAngle, CornerFull
      .WriteProperty "Enabled", m_ButtonProperty.Enabled, True
      .WriteProperty "Font", UserControl.Font, Ambient.Font
      .WriteProperty "ForeColor", m_ButtonProperty.ForeColor, Ambient.ForeColor
      .WriteProperty "HandPointer", m_ButtonProperty.HandPointer, False
      .WriteProperty "MaskColor", m_ButtonProperty.MaskColor, &HC0C0C0
      .WriteProperty "MouseIcon", UserControl.MouseIcon, Nothing
      .WriteProperty "MousePointer", UserControl.MousePointer, vbDefault
      .WriteProperty "OLEDropMode", UserControl.OLEDropMode, odNone
      .WriteProperty "Picture", m_ButtonProperty.Picture, Nothing
      .WriteProperty "PicAlign", m_ButtonProperty.PicAlign, LeftOfCaption
      .WriteProperty "PicSize", m_ButtonProperty.PicSize, SizePicture
      .WriteProperty "PicOpacity", m_ButtonProperty.PicOpacity, 1
      .WriteProperty "Shape", m_ButtonProperty.Shape, ShapeNone
      .WriteProperty "ShineLeft", m_ButtonProperty.ShineFullLeft, False
      .WriteProperty "ShineRight", m_ButtonProperty.ShineFullRight, False
      .WriteProperty "Sound", m_ButtonProperty.Sound, False
      .WriteProperty "UseMask", m_ButtonProperty.UseMask, True
      .WriteProperty "Value", m_ButtonProperty.Value, False
   End With
    
End Sub
