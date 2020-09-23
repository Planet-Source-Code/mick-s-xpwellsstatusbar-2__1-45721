VERSION 5.00
Begin VB.UserControl xpWellsStatusBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5940
   ControlContainer=   -1  'True
   PropertyPages   =   "xpWellsStatusBar.ctx":0000
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   396
   ToolboxBitmap   =   "xpWellsStatusBar.ctx":0035
End
Attribute VB_Name = "xpWellsStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'Control Name:          xpWellsStatusBar
'Created:               01/02/2003
'Author:                Richard Wells.

'Acknowledgements:

'Ariad Software.
'For letting me look through there ToolBar code
'to see how they use Property Pages

'Manjula Dharmawardhana at www.manjulapra.com
'For his simple Common Dialog without the .OCX sample

'Special Thanks:
'Steve McMahon ( The Man ) at www.vbaccelerator.com
'for showing us mere mortals how to make quality ActiveX controls.
'Without his generosity and skills, this control would not have happened.

Public Enum Apperance
    [Office XP] = 0
    [Windows XP] = 1
End Enum


'API Stuff.
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function ReleaseCapture Lib "user32" () As Long
    'GDI and reigons.
        Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
        Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
    '//

'Gripper Stuff
    Private Const WM_NCLBUTTONDOWN = &HA1
    Private Const HTBOTTOMRIGHT = 17
    Private bDrawGripper            As Boolean
    Private bShowGripper            As Boolean
    Private frm                     As Form
    Private WithEvents eForm        As Form
Attribute eForm.VB_VarHelpID = -1
    Private rcGripper               As RECT
'//
    
'Panel Stuff.
    Private m_Panels()              As New cPanels
    Private m_PanelCount            As Long
    Private rcPanel()               As RECT
    
    'Used for Click and DblClick Events
    Private PanelNum                As Long
    '//
'//

'Colors
    'Panel colors and global mask color.
        Private oBackColor          As OLE_COLOR
        Private oForeColor          As OLE_COLOR
        Private oMaskColor          As OLE_COLOR
        Private oDissColor          As OLE_COLOR
    '//
'//

Dim cPic                        As cImageManipulation
Event MouseDownInPanel(iPanel As Long)
Event Click(iPanelNumber)
Event DblClick(iPanelNumber)
'Default Property Values:
Const m_def_Apperance = 0
Dim m_Apperance As Apperance



Private Sub UserControl_Click()
    If m_Panels(PanelNum).pEnabled = True Then
        RaiseEvent Click(PanelNum)
    End If
End Sub

Private Sub UserControl_DblClick()
    If m_Panels(PanelNum).pEnabled = True Then
        RaiseEvent DblClick(PanelNum)
    End If
End Sub

Private Sub UserControl_InitProperties()
    oBackColor = vbButtonFace
    oForeColor = vbButtonText
    oDissColor = vbGrayText
    oMaskColor = RGB(255, 0, 255)
    bShowGripper = True
   m_Apperance = m_def_Apperance
End Sub

Private Sub UserControl_Terminate()
    Set frm = Nothing
    Erase rcPanel
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim pt As POINTAPI
Dim hRgn As Long
Dim i As Long
    PanelNum = 0
    If ShowGripper = True Then
        hRgn = CreateRectRgnIndirect(rcGripper)
        If PtInRegion(hRgn, CLng(X), CLng(Y)) Then
            If Button = vbLeftButton Then
                SizeByGripper frm.hwnd
                DeleteObjectReference hRgn
                Exit Sub
            End If
        End If
        
    End If
    For i = 1 To m_PanelCount
        hRgn = CreateRectRgnIndirect(rcPanel(i))
        If PtInRegion(hRgn, CLng(X), CLng(Y)) Then
            If Button = vbLeftButton Then
                If m_Panels(i).pEnabled = True Then
                    PanelNum = i
                    RaiseEvent MouseDownInPanel(i)
                End If
                DeleteObjectReference hRgn
            End If
        End If
    Next i
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim hRgn As Long
Dim i As Long
    If ShowGripper = True Then
        hRgn = CreateRectRgnIndirect(rcGripper)
        If PtInRegion(hRgn, CLng(X), CLng(Y)) Then
            UserControl.MousePointer = 8
            DeleteObjectReference hRgn
            Exit Sub
        Else
            UserControl.MousePointer = 0
        End If

    Else
        UserControl.MousePointer = 0
    End If
    For i = 1 To m_PanelCount
        hRgn = CreateRectRgnIndirect(rcPanel(i))
        If PtInRegion(hRgn, CLng(X), CLng(Y)) Then
            Extender.ToolTipText = m_Panels(i).ToolTipTxt
            DeleteObjectReference hRgn
        Else
            DeleteObjectReference hRgn
        End If
    Next i
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
End Sub

Public Property Get BackColor() As OLE_COLOR
    BackColor = oBackColor
End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
    oBackColor = NewBackColor
    UserControl.BackColor = NewBackColor
    DrawStatusBar
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = oForeColor
End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)
    oForeColor = NewForeColor
    PropertyChanged "ForeColor"
    DrawStatusBar False
End Property

Public Property Get NumberOfPanels() As Long
    NumberOfPanels = m_PanelCount
End Property

Public Property Get PanelWidth(ByVal Index As Long) As Long
    PanelWidth = m_Panels(Index).ClientWidth
End Property

Public Property Let PanelWidth(ByVal Index As Long, ByVal PanelWidth As Long)
    m_Panels(Index).ClientWidth = PanelWidth
    DrawStatusBar
    PropertyChanged "PWidth"
End Property

Public Property Get PanelCaption(ByVal Index As Long) As String
    PanelCaption = m_Panels(Index).PanelText
End Property

Public Property Let PanelCaption(ByVal Index As Long, ByVal NewPanelCaption As String)
    Dim tmpTxtWidth As Long
    tmpTxtWidth = m_Panels(Index).PanelTxtWidth
    m_Panels(Index).PanelText = NewPanelCaption
    m_Panels(Index).PanelTxtWidth = UserControl.TextWidth(NewPanelCaption)
    If tmpTxtWidth <> m_Panels(Index).PanelTxtWidth Then
      DrawStatusBar
    Else
      DrawStatusBar False
    End If
    PropertyChanged "pText"
End Property

Public Property Get PanelAutoSize(ByVal Index As Long) As enmAutoSize
    PanelAutoSize = m_Panels(Index).AutoSize
End Property

Public Property Let PanelAutoSize(ByVal Index As Long, ByVal NewAutoSize As enmAutoSize)
    m_Panels(Index).AutoSize = NewAutoSize
    DrawStatusBar
    PropertyChanged "AutoSize"
End Property

Public Property Get ToolTipText(ByVal Index As Long) As String
    ToolTipText = m_Panels(Index).ToolTipTxt
End Property

Public Property Let ToolTipText(ByVal Index As Long, ByVal NewToolTipText As String)
    m_Panels(Index).ToolTipTxt = NewToolTipText
    PropertyChanged "pTTText"
End Property

Public Property Get PanelPicture(ByVal Index As Long) As StdPicture
    Set PanelPicture = m_Panels(Index).PanelPicture
End Property

Public Property Set PanelPicture(ByVal Index As Long, ByVal NewPanelPicture As StdPicture)
    Set m_Panels(Index).PanelPicture = NewPanelPicture
    DrawStatusBar
    PropertyChanged "PanelPicture"
End Property

Public Property Get PanelEnabled(ByVal Index As Long) As Boolean
    PanelEnabled = m_Panels(Index).pEnabled
End Property

Public Property Let PanelEnabled(ByVal Index As Long, ByVal NewEnabled As Boolean)
    m_Panels(Index).pEnabled = NewEnabled
    DrawStatusBar False
    PropertyChanged "pEnabled"
End Property

Public Property Get MaskColor() As OLE_COLOR
    MaskColor = oMaskColor
End Property

Public Property Let MaskColor(ByVal NewMaskColor As OLE_COLOR)
    oMaskColor = NewMaskColor
    PropertyChanged "MaskColor"
    DrawStatusBar False
End Property

Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal NewFont As Font)
    Set UserControl.Font = NewFont
    PropertyChanged "Font"
    DrawStatusBar
End Property

Public Property Get ShowGripper() As Boolean
Attribute ShowGripper.VB_ProcData.VB_Invoke_Property = "General"
    ShowGripper = bShowGripper
End Property

Public Property Let ShowGripper(ByVal NewValue As Boolean)
    bShowGripper = NewValue
    PropertyChanged "ShowGripper"
    DrawStatusBar
    If bShowGripper = True Then
        With UserControl
            If TypeOf .Parent Is Form Then
                If Not TypeOf .Parent Is MDIForm Then
                Set frm = .Parent
                    If Ambient.UserMode Then
                        Set eForm = frm
                    End If
                End If
            End If
        End With
    Else
        ReleaseCapture
    End If
End Property

Public Property Get ForeColorDissabled() As OLE_COLOR
    ForeColorDissabled = oDissColor
End Property

Public Property Let ForeColorDissabled(ByVal NewDissColor As OLE_COLOR)
    oDissColor = NewDissColor
    PropertyChanged "ForeColorDissabled"
    DrawStatusBar False
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim i As Long
On Error GoTo ERH:
    With PropBag
        BackColor = .ReadProperty("BackColor", vbButtonFace)
        ForeColor = .ReadProperty("ForeColor", vbButtonText)
        ForeColorDissabled = .ReadProperty("ForeColorDissabled", vbGrayText)
        Set UserControl.Font = .ReadProperty("Font", UserControl.Ambient.Font)
        m_PanelCount = .ReadProperty("NumberOfPanels", 0)
        ReDim m_Panels(m_PanelCount) As New cPanels
        MaskColor = .ReadProperty("MaskColor", RGB(255, 0, 255))
        ShowGripper = .ReadProperty("ShowGripper", True)
        m_Apperance = .ReadProperty("Apperance", m_def_Apperance)
    End With
    For i = 1 To m_PanelCount
        With m_Panels(i)
            .MinWidth = PropBag.ReadProperty("MinWidth" & i)
            .ClientWidth = PropBag.ReadProperty("PWidth" & i)
            .ToolTipTxt = PropBag.ReadProperty("pTTText" & i)
            .PanelText = PropBag.ReadProperty("pText" & i)
            .pEnabled = PropBag.ReadProperty("pEnabled" & i)
            .AutoSize = PropBag.ReadProperty("AutoSize" & i, 0)
            Set .PanelPicture = PropBag.ReadProperty("PanelPicture" & i)
        End With
    Next i
Exit Sub
ERH:
If Err.Number = 327 Then
    Err.Clear
End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim i As Long
    With PropBag
        .WriteProperty "BackColor", oBackColor
        .WriteProperty "ForeColor", oForeColor
        .WriteProperty "ForeColorDissabled", oDissColor
        .WriteProperty "Font", UserControl.Font
        .WriteProperty "NumberOfPanels", m_PanelCount
        .WriteProperty "MaskColor", oMaskColor
        .WriteProperty "ShowGripper", bShowGripper
        .WriteProperty "Apperance", m_Apperance, m_def_Apperance
    End With

    For i = 1 To m_PanelCount
        With m_Panels(i)
            PropBag.WriteProperty "PWidth" & i, .ClientWidth
            PropBag.WriteProperty "pText" & i, .PanelText
            PropBag.WriteProperty "pTTText" & i, .ToolTipTxt
            PropBag.WriteProperty "pEnabled" & i, .pEnabled
            PropBag.WriteProperty "AutoSize" & i, .AutoSize, 0
            PropBag.WriteProperty "PanelPicture" & i, .PanelPicture
            PropBag.WriteProperty "MinWidth" & i, .MinWidth
        End With
    Next i
End Sub

Private Sub UserControl_Resize()
   If UserControl.Parent.WindowState <> 2 Then
      bDrawGripper = True
   Else
      bDrawGripper = False
   End If
   DrawStatusBar
End Sub

Public Sub DrawGripper()
    With rcGripper
        .Left = UserControl.ScaleWidth - 15
        .Right = UserControl.ScaleWidth
        .Bottom = UserControl.ScaleHeight
        .Top = UserControl.ScaleHeight '- 15
    End With
    If m_Apperance = [Office XP] Then
      With UserControl
          'Retain the area
          DrawASquare .hdc, rcGripper, .BackColor, True
          
          DrawALine .hdc, .ScaleWidth - 2, .ScaleHeight, .ScaleWidth, .ScaleHeight - 2, TranslateColorToRGB(GetSysColor(COLOR_BTNSHADOW), 0, 0, 0)
          DrawALine .hdc, .ScaleWidth - 3, .ScaleHeight, .ScaleWidth, .ScaleHeight - 3, TranslateColorToRGB(GetSysColor(COLOR_BTNSHADOW), 0, 0, 0)
          DrawALine .hdc, .ScaleWidth - 4, .ScaleHeight, .ScaleWidth, .ScaleHeight - 4, TranslateColorToRGB(GetSysColor(COLOR_BTNSHADOW), 0, 0, 0)
          DrawALine .hdc, .ScaleWidth - 5, .ScaleHeight, .ScaleWidth, .ScaleHeight - 5, TranslateColorToRGB(GetSysColor(COLOR_BTNHIGHLIGHT), 0, 0, 0)
          
          DrawALine .hdc, .ScaleWidth - 7, .ScaleHeight, .ScaleWidth, .ScaleHeight - 7, TranslateColorToRGB(GetSysColor(COLOR_BTNSHADOW), 0, 0, 0)
          DrawALine .hdc, .ScaleWidth - 8, .ScaleHeight, .ScaleWidth, .ScaleHeight - 8, TranslateColorToRGB(GetSysColor(COLOR_BTNSHADOW), 0, 0, 0)
          DrawALine .hdc, .ScaleWidth - 9, .ScaleHeight, .ScaleWidth, .ScaleHeight - 9, TranslateColorToRGB(GetSysColor(COLOR_BTNSHADOW), 0, 0, 0)
          DrawALine .hdc, .ScaleWidth - 10, .ScaleHeight, .ScaleWidth, .ScaleHeight - 10, TranslateColorToRGB(GetSysColor(COLOR_BTNHIGHLIGHT), 0, 0, 0)
          
          DrawALine .hdc, .ScaleWidth - 12, .ScaleHeight, .ScaleWidth, .ScaleHeight - 12, TranslateColorToRGB(GetSysColor(COLOR_BTNSHADOW), 0, 0, 0)
          DrawALine .hdc, .ScaleWidth - 13, .ScaleHeight, .ScaleWidth, .ScaleHeight - 13, TranslateColorToRGB(GetSysColor(COLOR_BTNSHADOW), 0, 0, 0)
          DrawALine .hdc, .ScaleWidth - 14, .ScaleHeight, .ScaleWidth, .ScaleHeight - 14, TranslateColorToRGB(GetSysColor(COLOR_BTNSHADOW), 0, 0, 0)
          DrawALine .hdc, .ScaleWidth - 15, .ScaleHeight, .ScaleWidth, .ScaleHeight - 15, TranslateColorToRGB(GetSysColor(COLOR_BTNHIGHLIGHT), 0, 0, 0)
          
          UserControl.Refresh
      End With
    Else
      With UserControl
          'Retain the area
          DrawASquare .hdc, rcGripper, .BackColor, True
          DrawALine .hdc, rcGripper.Left, rcGripper.Bottom - 1, rcGripper.Right, rcGripper.Bottom - 1, TranslateColorToRGB(oBackColor, 0, 0, 0, -15), 2
          DrawALine .hdc, rcGripper.Left, rcGripper.Bottom - 3, rcGripper.Right, rcGripper.Bottom - 3, TranslateColorToRGB(oBackColor, 0, 0, 0, -8), 2
          
          DrawALine .hdc, .ScaleWidth - 3, .ScaleHeight - 3, .ScaleWidth - 3, .ScaleHeight - 3, TranslateColorToRGB(.BackColor, 0, 0, 0, 50), 2
          DrawALine .hdc, .ScaleWidth - 7, .ScaleHeight - 3, .ScaleWidth - 7, .ScaleHeight - 3, TranslateColorToRGB(.BackColor, 0, 0, 0, 50), 2
          DrawALine .hdc, .ScaleWidth - 11, .ScaleHeight - 3, .ScaleWidth - 11, .ScaleHeight - 3, TranslateColorToRGB(.BackColor, 0, 0, 0, 50), 2
      
          DrawALine .hdc, .ScaleWidth - 3, .ScaleHeight - 7, .ScaleWidth - 3, .ScaleHeight - 7, TranslateColorToRGB(.BackColor, 0, 0, 0, 50), 2
          DrawALine .hdc, .ScaleWidth - 7, .ScaleHeight - 7, .ScaleWidth - 7, .ScaleHeight - 7, TranslateColorToRGB(.BackColor, 0, 0, 0, 50), 2
      
          DrawALine .hdc, .ScaleWidth - 3, .ScaleHeight - 11, .ScaleWidth - 3, .ScaleHeight - 11, TranslateColorToRGB(.BackColor, 0, 0, 0, 50), 2
      
      
          DrawALine .hdc, .ScaleWidth - 4, .ScaleHeight - 4, .ScaleWidth - 4, .ScaleHeight - 4, TranslateColorToRGB(.BackColor, 0, 0, 0, -50), 2
          DrawALine .hdc, .ScaleWidth - 8, .ScaleHeight - 4, .ScaleWidth - 8, .ScaleHeight - 4, TranslateColorToRGB(.BackColor, 0, 0, 0, -50), 2
          DrawALine .hdc, .ScaleWidth - 12, .ScaleHeight - 4, .ScaleWidth - 12, .ScaleHeight - 4, TranslateColorToRGB(.BackColor, 0, 0, 0, -50), 2
      
          DrawALine .hdc, .ScaleWidth - 4, .ScaleHeight - 8, .ScaleWidth - 4, .ScaleHeight - 8, TranslateColorToRGB(.BackColor, 0, 0, 0, -50), 2
          DrawALine .hdc, .ScaleWidth - 8, .ScaleHeight - 8, .ScaleWidth - 8, .ScaleHeight - 8, TranslateColorToRGB(.BackColor, 0, 0, 0, -50), 2
      
          DrawALine .hdc, .ScaleWidth - 4, .ScaleHeight - 12, .ScaleWidth - 4, .ScaleHeight - 12, TranslateColorToRGB(.BackColor, 0, 0, 0, -50), 2
          UserControl.Refresh
      End With
    End If

End Sub

Public Function AddPanel(Optional iPanelWidth As Long = 100, _
                        Optional sPanelText As String = "", _
                        Optional sToolTip As String = "", _
                        Optional bEnabled As Boolean = True, _
                        Optional eAutoSize As enmAutoSize = NoAutoSize, _
                        Optional pPanelPicture As StdPicture = Nothing) As Long
    m_PanelCount = m_PanelCount + 1
    ReDim Preserve m_Panels(m_PanelCount) As New cPanels
        With m_Panels(m_PanelCount)
            .ClientWidth = iPanelWidth
            .ToolTipTxt = sToolTip
            .PanelText = sPanelText
            .pEnabled = bEnabled
            .AutoSize = eAutoSize
            Set .PanelPicture = pPanelPicture
        End With
        PropertyChanged "NumberOfPanels"
        AddPanel = m_PanelCount
        DrawStatusBar
End Function

Public Function DeletePanel()
    If m_PanelCount > 1 Then
        m_PanelCount = m_PanelCount - 1
    End If
    PropertyChanged "NumberOfPanels"
    DrawStatusBar
End Function

Public Sub DrawStatusBar(Optional FullRedraw As Boolean = True)
Dim i                   As Long
Dim rc                  As RECT
Dim rcTemp              As RECT
Dim X                   As Long
Dim Y                   As Long
Dim X1                  As Long
Dim Y1                  As Long
Dim iOffset             As Long
Dim pX                  As Long
Dim pY                  As Long
'FullRedraw = True
iOffset = -36
If FullRedraw = True Then
   With UserControl
       'Control Shading Lines.
       Cls
       
       'Top lines
       If m_Apperance = [Office XP] Then
       
       Else
         DrawALine .hdc, 0, 0, .ScaleWidth, 0, TranslateColorToRGB(oBackColor, 0, 0, 0, -45)
         For i = 1 To 4
             DrawALine .hdc, 0, i, .ScaleWidth, i, TranslateColorToRGB(oBackColor, 0, 0, 0, iOffset)
             iOffset = iOffset + 9
         Next i
         'Bottom Lines
         DrawALine .hdc, 0, .ScaleHeight - 1, .ScaleWidth, .ScaleHeight - 1, TranslateColorToRGB(oBackColor, 0, 0, 0, -15), 2
         DrawALine .hdc, 0, .ScaleHeight - 3, .ScaleWidth, .ScaleHeight - 3, TranslateColorToRGB(oBackColor, 0, 0, 0, -8), 2
       End If
   End With
End If
'The Panels.
    '******************* Dimentions. **********************
    'X = Left of the panel
    'Y = Top of the panel
    'X1 = Width of the panel
    'Y1 = Height of the panel
    '******************************************************
    
    If m_Apperance = [Office XP] Then
       'Start the panel 5 pixels down from the top edge.
      Y = 1
      'Height of the panel
      Y1 = UserControl.ScaleHeight - 1
    Else
      'Start the panel 5 pixels down from the top edge.
      Y = 5
      'Height of the panel
      Y1 = UserControl.ScaleHeight - 4
    End If
    'Loop through the panels
    For i = 1 To m_PanelCount
        With m_Panels(i)
         SetPanelDementions
        'Position the panel.
            .ClientLeft = X
            .ClientTop = Y
            'X1 is taken from property "PanelWidth"
            X1 = .ClientWidth
            '//
            .ClientHeight = Y1
        '//
        'Create a RECT area using the above dimentions to draw into.
            If m_Apperance = [Office XP] Then
               With rc
                   .Left = X
                   .Top = Y
                   .Right = .Left + X1 - 1
                   .Bottom = Y1
               End With
               ReDim Preserve rcPanel(i)
               rcPanel(i) = rc

            Else
               With rc
                   .Left = X
                   .Top = Y
                   .Right = .Left + X1
                   .Bottom = Y1
               End With
               ReDim Preserve rcPanel(i)
               rcPanel(i) = rc
               ResizeRect rcPanel(i), -2, 0
            End If

        If FullRedraw = True Then
            If m_Apperance = [Windows XP] Then
               'Draw the seperators taking into acount the first and last
               'panel seperators are different.
               If i <> 1 Then
                  'This will draw the left line ( The lighter shade )
                  'so the first panel does not need one
                  DrawALine UserControl.hdc, X, Y, X, Y1, TranslateColorToRGB(oBackColor, 0, 0, 0, 50)
               End If
               If i <> m_PanelCount Then
                  'This will draw the right line ( The darker shade )
                  'Every panel will have this line exept the last
                  'panel has this line positioned differently.
                   DrawALine UserControl.hdc, rc.Right - 1, Y, rc.Right - 1, Y1, TranslateColorToRGB(oBackColor, 0, 0, 0, -50)
               Else
                  If i = m_PanelCount Then
                  'Lines for the last panel.
                     If bDrawGripper = True And bShowGripper = True Then
                        DrawALine UserControl.hdc, rc.Right - 1, Y, rc.Right - 1, Y1, TranslateColorToRGB(oBackColor, 0, 0, 0, 50)
                        DrawALine UserControl.hdc, rc.Right - 2, Y, rc.Right - 2, Y1, TranslateColorToRGB(oBackColor, 0, 0, 0, -50)
                     End If
                  End If
               End If
            End If
        End If
        
        If m_Apperance = [Office XP] Then
            DrawASquare UserControl.hdc, rcPanel(i), oBackColor, True
'            DrawGradient UserControl.hdc, vbWhite, vbBlue, rcPanel(i), .ClientHeight
            DrawASquare UserControl.hdc, rcPanel(i), vbButtonShadow, False
        Else
            DrawASquare UserControl.hdc, rcPanel(i), oBackColor, True
        End If
        'Get the size of the picture
        'even if there is not one set.
            GetPanelPictureSize i, pX, pY
        '//

        'Create a temporary RECT to draw some text into.
            rcTemp = GetRect(UserControl.hwnd)
            GetTextRect UserControl.hdc, .PanelText, Len(.PanelText), rcTemp
        '//
        'Copy the temporary RECT
            CopyTheRect rc, rcTemp
        '//
        'Position our RECT
            rc.Left = X
            rc.Right = ((rc.Left + X1) - 6) - pX
        '//
        'Draw the text into our new panel.
            If .pEnabled = True Then
                SetTheTextColor UserControl.hdc, oForeColor
            Else
                SetTheTextColor UserControl.hdc, oDissColor
            End If
            PositionRect rc, 2 + pX + 4, (ScaleHeight - rc.Bottom) / 2
            DrawTheText UserControl.hdc, .PanelText, Len(.PanelText), rc, [Use Ellipsis]
        '//
        'Add a PanelPicture if required.
        'TODO :
        'Picture will spill into the next panel if for some
        'reason someone sets the PanelWidth to
        'a smaller width than the image.
            Set cPic = New cImageManipulation
            cPic.PaintTransparentPicture UserControl.hdc, .PanelPicture, X + 3, (ScaleHeight - pY) / 2, pX, pY, 0, 0, oMaskColor
            Refresh
            Set cPic = Nothing
        '//
        'Dont forget to move the X ( Or left )
        'for the next panel.
            X = X + .ClientWidth
        '//
        End With
    Next i
'//
    If bDrawGripper = True And bShowGripper = True Then
        DrawGripper
    End If
End Sub

Private Sub GetPanelPictureSize(Index As Long, X As Long, Y As Long)
    If m_Panels(Index).PanelPicture Is Nothing Then Exit Sub
    X = 0
    Y = 0
    X = UserControl.ScaleX(m_Panels(Index).PanelPicture.Width, 8, UserControl.ScaleMode)
    Y = UserControl.ScaleY(m_Panels(Index).PanelPicture.Height, 8, UserControl.ScaleMode)
End Sub

Private Sub SizeByGripper(ByVal iHwnd As Long)
  ReleaseCapture
  SendMessage iHwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Apperance() As Apperance
   Apperance = m_Apperance
End Property

Public Property Let Apperance(ByVal New_Apperance As Apperance)
   m_Apperance = New_Apperance
   DrawStatusBar
   PropertyChanged "Apperance"
End Property


Private Sub SetPanelDementions()
   Dim i As Long
   Dim tmpSpringCount As Long
   Dim tmpwidth As Long
   For i = 1 To m_PanelCount
      With m_Panels(i)
         If .AutoSize = Contents Then
            If Len(.PanelText) = 0 Then
               If UserControl.ScaleX(.PanelPicture.Width, 8, UserControl.ScaleMode) = 0 Then
                  .ClientWidth = UserControl.ScaleX(.PanelPicture.Width, 8, UserControl.ScaleMode)
               Else
                  .ClientWidth = UserControl.ScaleX(.PanelPicture.Width, 8, UserControl.ScaleMode) + 6
               End If
            Else
               .ClientWidth = UserControl.TextWidth(.PanelText) + (UserControl.ScaleX(.PanelPicture.Width, 8, UserControl.ScaleMode)) + 12
            End If
         ElseIf .AutoSize = NoAutoSize Then
            .ClientWidth = .MinWidth
         End If
      End With
   Next i
   For i = 1 To m_PanelCount
      With m_Panels(i)
         If .AutoSize = Spring Then
            tmpSpringCount = tmpSpringCount + 1
         Else
            tmpwidth = tmpwidth + .ClientWidth
         End If
      End With
   Next i
   For i = 1 To m_PanelCount
      With m_Panels(i)
         If .AutoSize = Spring Then
            If bDrawGripper = True And bShowGripper = True Then
               .ClientWidth = (UserControl.ScaleWidth - tmpwidth - 15) / tmpSpringCount
            Else
               .ClientWidth = (UserControl.ScaleWidth - tmpwidth) / tmpSpringCount
            End If
            If .ClientWidth < UserControl.ScaleX(.PanelPicture.Width, 8, UserControl.ScaleMode) Then .ClientWidth = UserControl.ScaleX(.PanelPicture.Width, 8, UserControl.ScaleMode)
         End If
      End With
   Next i
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,50
Public Property Get MinWidth(ByVal Index As Long) As Long
Attribute MinWidth.VB_ProcData.VB_Invoke_Property = "StatusBarPanels"
   MinWidth = m_Panels(Index).MinWidth
End Property

Public Property Let MinWidth(ByVal Index As Long, ByVal New_MinWidth As Long)
   m_Panels(Index).MinWidth = New_MinWidth
   DrawStatusBar
   PropertyChanged "MinWidth"
End Property

