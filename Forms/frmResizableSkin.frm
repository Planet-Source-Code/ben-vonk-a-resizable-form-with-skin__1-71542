VERSION 5.00
Begin VB.Form frmResizableSkin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Resizable Skin Demo"
   ClientHeight    =   2652
   ClientLeft      =   1260
   ClientTop       =   1560
   ClientWidth     =   6132
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmResizableSkin.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   221
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   511
   Begin ResizableSkin.CrystalButton cbtControl 
      Height          =   240
      Index           =   1
      Left            =   5280
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   48
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   423
      BackColor       =   12582912
      Caption         =   ""
      CornerAngle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
      Shape           =   3
   End
   Begin VB.PictureBox picSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   888
      Left            =   5040
      Picture         =   "frmResizableSkin.frx":08CA
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   87
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   1044
   End
   Begin VB.Timer tmrSysTrayPosition 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   5040
      Top             =   2160
   End
   Begin ResizableSkin.CrystalButton cbtControl 
      Height          =   240
      Index           =   0
      Left            =   4920
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Minimize"
      Top             =   48
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   423
      BackColor       =   12582912
      Caption         =   ""
      CornerAngle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmResizableSkin.frx":555C
      Shape           =   2
   End
   Begin ResizableSkin.CrystalButton cbtControl 
      Height          =   240
      Index           =   2
      Left            =   5640
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Close"
      Top             =   48
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   423
      BackColor       =   192
      Caption         =   ""
      CornerAngle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmResizableSkin.frx":56B6
      Shape           =   1
   End
   Begin VB.Image imgControl 
      Height          =   192
      Index           =   1
      Left            =   5760
      Picture         =   "frmResizableSkin.frx":5810
      ToolTipText     =   "Restore"
      Top             =   1680
      Width           =   192
   End
   Begin VB.Image imgControl 
      Height          =   192
      Index           =   0
      Left            =   5520
      Picture         =   "frmResizableSkin.frx":595A
      ToolTipText     =   "Maximize"
      Top             =   1680
      Width           =   192
   End
   Begin VB.Image imgIcon 
      Height          =   384
      Left            =   5040
      Top             =   1680
      Visible         =   0   'False
      Width           =   384
   End
End
Attribute VB_Name = "frmResizableSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Private Constant
Private Const ICON_SIZE As Integer = 52

' Private Variable
Private MouseX          As Single
Private MouseY          As Single

Public Sub EndResizableSkin()

   Unload Me

End Sub

Public Sub ToggleWindowState()

   If WindowState = vbMaximized Then
      WindowState = vbNormal
      
   Else
      WindowState = vbMaximized
   End If

End Sub

Private Sub cbtControl_Click(Index As Integer)

   Select Case Index
      Case 0
         WindowState = vbMinimized
         
      Case 1
         Call ToggleWindowState
         
      Case 2
         Unload Me
   End Select

End Sub

Private Sub Form_DblClick()

   If MouseY < ICON_SIZE Then
      If MouseX < ICON_SIZE Then
         Call EndResizableSkin
         
      Else
         Call ToggleWindowState
      End If
   End If

End Sub

Private Sub Form_Load()

   imgIcon.Picture = Icon
   
   Call Initialize(Me, picSkin, imgIcon, ICON_SIZE)
   Call SetCaptionFont("Arial", 10, True, , &H80FFFF)
   Call RemoveTitleBar

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If MousePointer = vbSizeAll Then MousePointer = vbNormal
   If (X < ICON_SIZE) And (Y < ICON_SIZE) Then If Not SysMenuOpen Then Call ShowSystemMenu
   If (X > ICON_SIZE) Or (Y > ICON_SIZE) Then SysMenuOpen = False

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Static blnButtonDown As Boolean

   MouseX = X
   MouseY = Y
   
   If Button = vbLeftButton Then
      If Not blnButtonDown And ((X > ICON_SIZE) And (Y < ICON_SIZE)) Then
         Call MovingWindow(X + Y * &H10000)
         
      Else
         blnButtonDown = True
      End If
      
   Else
      blnButtonDown = False
   End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   SysMenuOpen = False

End Sub

Private Sub Form_Resize()

Static blnMaximized As Boolean

   If WindowState = vbMaximized Then
      If blnMaximized Then Exit Sub
      
      blnMaximized = True
      
      Call MakeSizable(False)
      Call DockToSysTray
      
      tmrSysTrayPosition.Enabled = True
      
   Else
      tmrSysTrayPosition.Enabled = False
   End If
   
   Call CreateWindow(Width / Screen.TwipsPerPixelX, Height / Screen.TwipsPerPixelX, 40, AllRound, blnMaximized)
   
   If blnMaximized Then
      Call MakeSizable(True)
      
      blnMaximized = False
   End If
   
   Call DrawSkin
   
   With cbtControl.Item(2)
      .Left = ScaleWidth - .Width - 10
      cbtControl.Item(1).Left = .Left - .Width + 1
   End With
   
   With cbtControl.Item(1)
      cbtControl.Item(0).Left = .Left - .Width + 1
      .Picture = imgControl.Item(Abs(tmrSysTrayPosition.Enabled)).Picture
      .ToolTipText = imgControl.Item(Abs(tmrSysTrayPosition.Enabled)).ToolTipText
   End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

   Call Terminate
   
   Set frmResizableSkin = Nothing

End Sub

Private Sub tmrSysTrayPosition_Timer()

   Call DockToSysTray

End Sub
