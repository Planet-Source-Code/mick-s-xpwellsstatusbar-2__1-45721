VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "XP StatusBar Control"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrKey 
      Interval        =   100
      Left            =   1200
      Top             =   480
   End
   Begin StatusBarTest.xpWellsStatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   2415
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   609
      BackColor       =   14215660
      ForeColor       =   -2147483630
      ForeColorDissabled=   9915703
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumberOfPanels  =   6
      MaskColor       =   16711935
      ShowGripper     =   -1  'True
      Apperance       =   1
      PWidth1         =   0
      pText1          =   ""
      pTTText1        =   "Hello"
      pEnabled1       =   -1  'True
      AutoSize1       =   2
      PanelPicture1   =   "frmTest.frx":0000
      MinWidth1       =   100
      PWidth2         =   188
      pText2          =   "My Computer"
      pTTText2        =   ""
      pEnabled2       =   -1  'True
      AutoSize2       =   1
      PanelPicture2   =   "frmTest.frx":001C
      MinWidth2       =   100
      PWidth3         =   188
      pText3          =   "Internet"
      pTTText3        =   ""
      pEnabled3       =   -1  'True
      AutoSize3       =   1
      PanelPicture3   =   "frmTest.frx":036E
      MinWidth3       =   80
      PWidth4         =   22
      pText4          =   ""
      pTTText4        =   "Privacy Report"
      pEnabled4       =   -1  'True
      AutoSize4       =   2
      PanelPicture4   =   "frmTest.frx":06C0
      MinWidth4       =   25
      PWidth5         =   22
      pText5          =   ""
      pTTText5        =   "You Have New Mail"
      pEnabled5       =   -1  'True
      AutoSize5       =   2
      PanelPicture5   =   "frmTest.frx":0A12
      MinWidth5       =   25
      PWidth6         =   40
      pText6          =   "CAPS"
      pTTText6        =   ""
      pEnabled6       =   -1  'True
      AutoSize6       =   2
      PanelPicture6   =   "frmTest.frx":0D64
      MinWidth6       =   40
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   720
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   7110
      TabIndex        =   2
      Top             =   0
      Width           =   7140
      Begin VB.CheckBox Check1 
         Caption         =   "Show Size Gripper"
         Height          =   255
         Left            =   1530
         TabIndex        =   4
         Top             =   90
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Office XP"
         Height          =   345
         Left            =   60
         TabIndex        =   3
         Top             =   30
         Width           =   1245
      End
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Width           =   1755
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Keyboard API
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long
Private Declare Function GetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long
Private Declare Function SetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long

Private kbArray As KeyboardBytes
Private Const VK_CAPITAL = &H14
Private Type KeyboardBytes
    kbByte(0 To 255) As Byte
End Type

Private Sub Check1_Click()
   sb.ShowGripper = Not sb.ShowGripper
End Sub

Private Sub Command1_Click()
   If sb.Apperance = [Office XP] Then
      sb.Apperance = [Windows XP]
      Command1.Caption = "Office XP"
   Else
      sb.Apperance = [Office XP]
      Command1.Caption = "Windows XP"
   End If
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   Text1.Width = Me.ScaleWidth
   Text1.Height = Me.ScaleHeight - sb.Height
End Sub

Private Sub sb_Click(iPanelNumber As Variant)
    If iPanelNumber = 1 Then
        MsgBox "Panel 1 Click"
    End If
End Sub

Private Sub sb_DblClick(iPanelNumber As Variant)
    If iPanelNumber = 2 Then
        sb.PanelCaption(2) = InputBox("Change Caption")
    End If
End Sub

Private Sub sb_MouseDownInPanel(iPanel As Long)
    If iPanel = 5 Then
        MsgBox "Mouse Down In Panel Number " & iPanel
    End If
End Sub

Private Sub Timer1_Timer()
    sb.PanelCaption(1) = Time
End Sub

Public Function GetCapsLockState() As Boolean
'True if caps lock is on
Dim i As Long
    GetKeyboardState kbArray
    i = kbArray.kbByte(VK_CAPITAL)
    If i = 1 Then
        GetCapsLockState = True
    End If
End Function

Private Sub tmrKey_Timer()
    sb.PanelEnabled(6) = GetCapsLockState
End Sub
