VERSION 5.00
Begin VB.UserControl XDWButton 
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3030
   ScaleWidth      =   3855
   ToolboxBitmap   =   "dwbutton.ctx":0000
End
Attribute VB_Name = "XDWButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' DWButton
' Author: David Crowell (davidc@qtm.net)
' See http://www.qtm.net/~davidc for updates
' Released to the public domain
'
' Last update: October 6, 1999
'
' Use this code at your own risk.
' I assume no liability for the use of this code.
'

' SetCapture & ReleaseCapture are to track the mouse
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

' TextOut & DrawEdge are used to draw the button
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

' DrawFocusRect is to draw the focus rectangle
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long

' GetTextExtentPoint32 is used to determine the size of a text string
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZEL) As Long

' needed UDTs for the API calls
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' The following are from PaintEffects.cls, a Microsoft sample
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(1) As RGBQUAD
End Type

Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As RGBQUAD) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long
Private Declare Function GetMapMode Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyHeight As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

'DrawIconEx Flags
Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = &H3
Private Const DI_COMPAT = &H4
Private Const DI_DEFAULTSIZE = &H8

'DIB Section constants
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

'Raster Operation Codes
Private Const DSna = &H220326 '0x00220326

'VB Errors
Private Const giINVALID_PICTURE As Integer = 481

Private m_hpalHalftone As Long  'Halftone created for default palette use

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' End of Microsoft code declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Type SIZEL
    cX As Long
    cY As Long
End Type

' constant types used with DrawEdge
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDOUTER = &H1
Private Const BF_SOFT = &H1000
Private Const BF_RECT = &HF
Private Const EDGE_BUMP = &H9&
Private Const EDGE_ETCHED = &H6&
Private Const BDR_RAISEDINNER = &H4

' private variables
Private rRect As RECT
Private rFocus As RECT
Private nControlHeight As Long
Private nControlWidth As Long
Private bHovering As Boolean
Private bPressed As Boolean
Private bHasFocus As Boolean
Private nTextLeft As Long
Private nTextTop As Long
Private nTextWidth As Long
Private nTextHeight As Long
Private nPicLeft As Long
Private nPicTop As Long
Private nPicWidth As Long
Private nPicHeight As Long

' Public Enum for Border type
Public Enum fbButtonBorderStyles
    fbbNone
    fbbEtched
    fbbBump
    fbbRaised
End Enum

' Public Enum for Picture Orientation
Public Enum fbPictureOrientation
    fbpoTop
    fbpoBottom
    fbpoLeft
    fbpoRight
End Enum

' property member variables
Private mToolTip As String
Private mCaption As String
Private mEnabled As Boolean
Private mShowFocus As Boolean
Private mTextColor As OLE_COLOR
Private mHoverColor As OLE_COLOR
Private mButtonBorderStyle As fbButtonBorderStyles
Private mPicture As StdPicture
Private mPictureMaskColor As OLE_COLOR
Private mPictureOrientation As fbPictureOrientation

' events
Public Event Click()
Attribute Click.VB_Description = "Occurs when button is clicked"
Attribute Click.VB_UserMemId = -600
Public Event MouseEnter()
Attribute MouseEnter.VB_Description = "Occurs when the mouse moves over the button."
Public Event MouseLeave()
Attribute MouseLeave.VB_Description = "Occurs when the mouse moves off of the button."

' Property Procedures
'
Public Property Get ToolTip() As String
Attribute ToolTip.VB_Description = "Text displayed on the button"
Attribute ToolTip.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ToolTip.VB_UserMemId = -518
    ToolTip = mToolTip
End Property

Public Property Let ToolTip(val As String)
    mToolTip = val
    UserControl_Resize
    UserControl_Paint
    PropertyChanged "ToolTip"
End Property

Public Property Get Caption() As String
    Caption = mCaption
End Property

Public Property Let Caption(val As String)
    mCaption = val
    UserControl_Resize
    UserControl_Paint
    PropertyChanged "Caption"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Determines whether or not the control will receive user events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = mEnabled
End Property

Public Property Let Enabled(val As Boolean)
    UserControl.Enabled = val
    mEnabled = val
    UserControl_Paint
    PropertyChanged "Enabled"
End Property

Public Property Get ShowFocus() As Boolean
Attribute ShowFocus.VB_Description = "Determines whether or not a focus rectangle is drawn on the control when necessary."
Attribute ShowFocus.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ShowFocus = mShowFocus
End Property

Public Property Let ShowFocus(val As Boolean)
    mShowFocus = val
    If Not val Then bHasFocus = False
    UserControl_Paint
    PropertyChanged "ShowFocus"
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Font to be used to draw the caption on the button."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(fnt As StdFont)
    Set UserControl.Font = fnt
    UserControl_Resize
    PropertyChanged "Font"
End Property

Public Property Get TextColor() As OLE_COLOR
Attribute TextColor.VB_Description = "Color the caption is drawn in in the button's normal state."
Attribute TextColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute TextColor.VB_UserMemId = -513
    TextColor = mTextColor
End Property

Public Property Let TextColor(Color As OLE_COLOR)
    mTextColor = Color
    UserControl_Paint
    PropertyChanged "TextColor"
End Property

Public Property Get HoverColor() As OLE_COLOR
Attribute HoverColor.VB_Description = "Color of the caption when the mouse is hovering over the button."
Attribute HoverColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HoverColor = mHoverColor
End Property

Public Property Let HoverColor(Color As OLE_COLOR)
    mHoverColor = Color
    UserControl_Paint
    PropertyChanged "HoverColor"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Background color of the object"
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(Color As OLE_COLOR)
    UserControl.BackColor = Color
    UserControl_Paint
    PropertyChanged "BackColor"
End Property

Public Property Get ButtonBorderStyle() As fbButtonBorderStyles
Attribute ButtonBorderStyle.VB_Description = "Type of border around button"
Attribute ButtonBorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ButtonBorderStyle.VB_UserMemId = -504
    ButtonBorderStyle = mButtonBorderStyle
End Property

Public Property Let ButtonBorderStyle(style As fbButtonBorderStyles)
    mButtonBorderStyle = style
    UserControl_Paint
    PropertyChanged "ButtonBorderStyle"
End Property

Public Property Get ButtonPicture() As StdPicture
Attribute ButtonPicture.VB_Description = "Picture to display on the button"
Attribute ButtonPicture.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set ButtonPicture = mPicture
End Property

Public Property Set ButtonPicture(pic As StdPicture)
    Set mPicture = pic
    UserControl_Resize
    PropertyChanged "ButtonPicture"
End Property

Public Property Get PictureMaskColor() As OLE_COLOR
Attribute PictureMaskColor.VB_Description = "Determines the color in the picture that is the background color."
Attribute PictureMaskColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PictureMaskColor = mPictureMaskColor
    UserControl_Paint
    PropertyChanged "PictureMaskColor"
End Property

Public Property Let PictureMaskColor(val As OLE_COLOR)
    mPictureMaskColor = val
    UserControl_Paint
End Property

Public Property Get PictureOrientation() As fbPictureOrientation
Attribute PictureOrientation.VB_Description = "Determines where the picture is drawn on the button in relation to the caption."
Attribute PictureOrientation.VB_ProcData.VB_Invoke_Property = ";Appearance"
    PictureOrientation = mPictureOrientation
End Property

Public Property Let PictureOrientation(val As fbPictureOrientation)
    mPictureOrientation = val
    UserControl_Resize
    PropertyChanged "PictureOrientation"
End Property

Private Sub UserControl_Initialize()
    Dim hdcScreen As Long
    hdcScreen = GetDC(0&)
    m_hpalHalftone = CreateHalftonePalette(hdcScreen)
    ReleaseDC 0&, hdcScreen
End Sub

Private Sub UserControl_InitProperties()
' Setup default property values
    Set Font = UserControl.Ambient.Font
    Caption = UserControl.Ambient.DisplayName
    Enabled = True
    ShowFocus = False
    TextColor = vbButtonText
    HoverColor = vbHighlight
    BackColor = vbButtonFace
    ButtonBorderStyle = fbbNone
    Set ButtonPicture = Nothing
    PictureMaskColor = vbRed
    PictureOrientation = fbpoTop
    UserControl_Load
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
' Read saved property values
    On Error Resume Next
    Set UserControl.Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
    Caption = PropBag.ReadProperty("Caption", UserControl.Ambient.DisplayName)
    ToolTip = PropBag.ReadProperty("ToolTip", UserControl.Ambient.ToolTip)
    Enabled = PropBag.ReadProperty("Enabled", True)
    ShowFocus = PropBag.ReadProperty("ShowFocus", False)
    TextColor = PropBag.ReadProperty("TextColor", vbButtonText)
    HoverColor = PropBag.ReadProperty("HoverColor", vbHighlight)
    BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    ButtonBorderStyle = PropBag.ReadProperty("ButtonBorderStyle", fbbNone)
    Set ButtonPicture = PropBag.ReadProperty("ButtonPicture", Nothing)
    PictureMaskColor = PropBag.ReadProperty("PictureMaskColor", vbRed)
    PictureOrientation = PropBag.ReadProperty("PictureOrientation", fbpoTop)
    UserControl_Load
End Sub

Private Sub UserControl_Load()
    '
End Sub

Private Sub UserControl_Terminate()
    Set mPicture = Nothing
    DeleteObject m_hpalHalftone
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
' Save design-time property values
    On Error Resume Next
    PropBag.WriteProperty "Font", UserControl.Font, UserControl.Ambient.Font
    PropBag.WriteProperty "Caption", mCaption, UserControl.Ambient.DisplayName
    PropBag.WriteProperty "ToolTip", mToolTip, UserControl.Ambient.ToolTip
    PropBag.WriteProperty "Enabled", mEnabled, True
    PropBag.WriteProperty "ShowFocus", mShowFocus, False
    PropBag.WriteProperty "TextColor", mTextColor, vbButtonText
    PropBag.WriteProperty "HoverColor", mHoverColor, vbHighlight
    PropBag.WriteProperty "BackColor", UserControl.BackColor, vbButtonFace
    PropBag.WriteProperty "ButtonBorderStyle", mButtonBorderStyle, fbbNone
    PropBag.WriteProperty "ButtonPicture", mPicture, vbNull
    PropBag.WriteProperty "PictureMaskColor", mPictureMaskColor, vbRed
    PropBag.WriteProperty "PictureOrientation", mPictureOrientation, fbpoTop
End Sub

Private Sub UserControl_GotFocus()
    If mShowFocus Then
        bHasFocus = True
        UserControl_Paint
    End If
End Sub

Private Sub UserControl_LostFocus()
    If mShowFocus Then
        bHasFocus = False
        UserControl_Paint
    End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
' allow the space bar to click the button
    If KeyCode = 32 Then UserControl_MouseDown vbLeftButton, 0, 0, 0
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
' keep track of when mouse button is pressed
    bPressed = True
    UserControl_Paint
End Sub

Private Sub UserControl_DblClick()
' in the double click event too
    bPressed = True
    UserControl_Paint
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    UserControl_MouseUp vbLeftButton, 0, 0, 0
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
' mouse button released
    Dim bTemp As Boolean
    bTemp = bPressed
    bPressed = False
    UserControl_Paint
    
    ' only raise the event if the button was pressed before being released
    If bTemp = True Then
        ' we don't want the button to remain raised after clicking
        bHovering = False
        UserControl_Paint
        RaiseEvent MouseLeave
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
' track the mouse
    Dim Temp As Boolean
    
    ' always call ReleaseCapture or it will cause problems for other controls
    Call ReleaseCapture
    
    ' is the mouse off of the button
    If (x < 0) Or (Y < 0) Or (x > nControlWidth) Or (Y > nControlHeight) Then
        Temp = False
    Else
        ' the mouse is still over the button, so be sure to call SetCapture again
        Temp = True
        Call SetCapture(UserControl.hwnd)
    End If
    
    ' only paint if necessary
    If bHovering <> Temp Then
        If Button <> 0 Then bPressed = True
        bHovering = Temp
        If bHovering = False Then bPressed = False
        UserControl_Paint
        If bHovering Then
            RaiseEvent MouseEnter
        Else
            RaiseEvent MouseLeave
        End If
    End If
    
End Sub

Private Sub UserControl_Paint()

    UserControl.Cls
    
    ' if we are in design mode, we'll want the raised look
    If Not UserControl.Ambient.UserMode Then bHovering = True
    
    If Not mEnabled Then
        DrawDisabled
    Else
    
        If bPressed Then
            DrawPressed
        Else
            If bHovering Then
                DrawRaised
            Else
                DrawFlat
            End If
        End If
        
        If bHasFocus Then DrawFocusRect UserControl.hDC, rFocus
        
    End If
    
    DrawPic mEnabled
    
End Sub

Private Sub DrawDisabled()

    Select Case mButtonBorderStyle
        Case fbbEtched
            DrawEdge UserControl.hDC, rRect, EDGE_ETCHED, BF_RECT
        Case fbbBump
            DrawEdge UserControl.hDC, rRect, EDGE_BUMP, BF_RECT
        Case fbbRaised
            DrawEdge UserControl.hDC, rRect, BDR_RAISEDOUTER, BF_RECT Or BF_SOFT
    End Select
    
    ' Draw disabled looking text
    UserControl.ForeColor = vb3DHighlight
    Call TextOut(UserControl.hDC, nTextLeft + 1, nTextTop + 1, mCaption, Len(mCaption))
    UserControl.ForeColor = vb3DShadow
    Call TextOut(UserControl.hDC, nTextLeft, nTextTop, mCaption, Len(mCaption))
    
End Sub

Private Sub DrawPressed()

    ' Draw the pressed button
    DrawEdge UserControl.hDC, rRect, BDR_SUNKENOUTER, BF_RECT Or BF_SOFT
    
    ' Print the caption
    UserControl.ForeColor = mHoverColor
    Call TextOut(UserControl.hDC, nTextLeft + 1, nTextTop + 1, mCaption, Len(mCaption))
    
End Sub

Private Sub DrawRaised()

    ' Draw the raised button
    DrawEdge UserControl.hDC, rRect, BDR_RAISEDOUTER, BF_RECT Or BF_SOFT
    
    ' Print the caption
    UserControl.ForeColor = mHoverColor
    Call TextOut(UserControl.hDC, nTextLeft, nTextTop, mCaption, Len(mCaption))
    
End Sub

Private Sub DrawFlat()

    Select Case mButtonBorderStyle
        Case fbbEtched
            DrawEdge UserControl.hDC, rRect, EDGE_ETCHED, BF_RECT
        Case fbbBump
            DrawEdge UserControl.hDC, rRect, EDGE_BUMP, BF_RECT
        Case fbbRaised
            DrawEdge UserControl.hDC, rRect, BDR_RAISEDOUTER, BF_RECT Or BF_SOFT
    End Select
    
    ' Print the caption
    UserControl.ForeColor = mTextColor
    Call TextOut(UserControl.hDC, nTextLeft, nTextTop, mCaption, Len(mCaption))
    
End Sub

Private Sub DrawPic(Enabled As Boolean)
    If nPicWidth Then
        If mEnabled Then
            If bPressed Then
                PaintTransparentStdPic UserControl.hDC, nPicLeft + 1, nPicTop + 1, nPicWidth, nPicHeight, mPicture, 0, 0, mPictureMaskColor
            Else
                PaintTransparentStdPic UserControl.hDC, nPicLeft, nPicTop, nPicWidth, nPicHeight, mPicture, 0, 0, mPictureMaskColor
            End If
        Else
            PaintDisabledStdPic UserControl.hDC, nPicLeft, nPicTop, nPicWidth, nPicHeight, mPicture, 0, 0, mPictureMaskColor
        End If
        
    End If
End Sub

Private Sub UserControl_Resize()
' store the size for later use when control is resized
    rRect.Left = 0
    rRect.Top = 0
    rRect.Bottom = UserControl.Height \ Screen.TwipsPerPixelY
    rRect.Right = UserControl.Width \ Screen.TwipsPerPixelX
    rFocus.Left = 3
    rFocus.Top = 3
    rFocus.Bottom = rRect.Bottom - 3
    rFocus.Right = rRect.Right - 3
    nControlHeight = UserControl.Height
    nControlWidth = UserControl.Width
    CalcPosition
    UserControl_Paint
End Sub

Private Sub CalcPosition()
    Dim slTemp As SIZEL
    Dim ntemp As Long
    Dim nSpacing As Long
    
    ' figure size of picture in pixels
    If Not mPicture Is Nothing Then
        nPicHeight = ScaleY(mPicture.Height) \ Screen.TwipsPerPixelY
        nPicWidth = ScaleX(mPicture.Width) \ Screen.TwipsPerPixelX
    Else
        nPicWidth = 0
        nPicHeight = 0
    End If
    
    ' get size of text in pixels
    If Len(mCaption) > 0 Then
        Call GetTextExtentPoint32(hDC, mCaption, Len(mCaption), slTemp)
        nTextWidth = slTemp.cX
        nTextHeight = slTemp.cY
    Else
        nTextWidth = 0
        nTextHeight = 0
    End If
    
    ' Determine whether to use extra spacing or not
    If (nTextWidth > 0) And (nPicWidth > 0) Then
        nSpacing = 5
    Else
        nSpacing = 0
    End If
    
    ' Determine picture & caption position based
    ' upon size & orientation
    Select Case mPictureOrientation
    
    Case fbpoTop
        nTextLeft = (rRect.Right - nTextWidth) \ 2
        nPicLeft = (rRect.Right - nPicWidth) \ 2
        ntemp = nTextHeight + nPicHeight + nSpacing
        nPicTop = (rRect.Bottom - ntemp) \ 2
        nTextTop = nPicTop + nPicHeight + nSpacing
        
    Case fbpoBottom
        nTextLeft = (rRect.Right - nTextWidth) \ 2
        nPicLeft = (rRect.Right - nPicWidth) \ 2
        ntemp = nTextHeight + nPicHeight + nSpacing
        nTextTop = (rRect.Bottom - ntemp) \ 2
        nPicTop = nTextTop + nTextHeight + nSpacing
    
    Case fbpoLeft
        nTextTop = (rRect.Bottom - nTextHeight) \ 2
        nPicTop = (rRect.Bottom - nPicHeight) \ 2
        ntemp = nTextWidth + nPicWidth + nSpacing
        nPicLeft = (rRect.Right - ntemp) \ 2
        nTextLeft = nPicLeft + nPicWidth + nSpacing
        
    Case fbpoRight
        nTextTop = (rRect.Bottom - nTextHeight) \ 2
        nPicTop = (rRect.Bottom - nPicHeight) \ 2
        ntemp = nTextWidth + nPicWidth + nSpacing
        nTextLeft = (rRect.Right - ntemp) \ 2
        nPicLeft = nTextLeft + nTextWidth + nSpacing
        
    End Select
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' The following code is from Microsoft PaintEffects.cls sample code.
' I've copied the needed portions for my own use.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub PaintDisabledStdPic(ByVal hdcDest As Long, _
                                ByVal xDest As Long, _
                                ByVal yDest As Long, _
                                ByVal Width As Long, _
                                ByVal Height As Long, _
                                ByVal picSource As StdPicture, _
                                ByVal xSrc As Long, _
                                ByVal ySrc As Long, _
                                Optional ByVal clrMask As OLE_COLOR = vbWhite, _
                                Optional ByVal clrHighlight As OLE_COLOR = vb3DHighlight, _
                                Optional ByVal clrShadow As OLE_COLOR = vb3DShadow, _
                                Optional ByVal hPal As Long = 0)
    Dim hdcSrc As Long         'HDC that the source bitmap is selected into
    Dim hbmMemSrcOld As Long
    Dim hbmMemSrc As Long
    Dim udtRect As RECT
    Dim hbrMask As Long
    Dim lMaskColor As Long
    Dim hdcScreen As Long
    Dim hPalOld As Long
    
    'Verify that the passed picture is not nothing
    If picSource Is Nothing Then GoTo PaintDisabledDC_InvalidParam
    Select Case picSource.Type
        Case vbPicTypeBitmap
            'Select passed picture into an HDC
            hdcScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                hPal = m_hpalHalftone
            End If
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrcOld = SelectObject(hdcSrc, picSource.handle)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            
            'Draw the bitmap
            PaintDisabledDC hdcDest, xDest, yDest, Width, Height, hdcSrc, xSrc, ySrc, clrMask, clrHighlight, clrShadow, hPal
            
            SelectObject hdcSrc, hbmMemSrcOld
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
        Case vbPicTypeIcon
            'Create a bitmap and select it into a DC
            hdcScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                hPal = m_hpalHalftone
            End If
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrc = CreateCompatibleBitmap(hdcScreen, Width, Height)
            hbmMemSrcOld = SelectObject(hdcSrc, hbmMemSrc)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw Icon onto DC
            udtRect.Bottom = Height
            udtRect.Right = Width
            OleTranslateColor clrMask, 0&, lMaskColor
            SetBkColor hdcSrc, lMaskColor
            hbrMask = CreateSolidBrush(lMaskColor)
            FillRect hdcSrc, udtRect, hbrMask
            DeleteObject hbrMask
            DrawIcon hdcSrc, 0, 0, picSource.handle
            'Draw Disabled image
            PaintDisabledDC hdcDest, xDest, yDest, Width, Height, hdcSrc, 0&, 0&, clrMask, clrHighlight, clrShadow, hPal
            'Clean up
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteObject SelectObject(hdcSrc, hbmMemSrcOld)
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
        Case Else
            GoTo PaintDisabledDC_InvalidParam
    End Select
    Exit Sub
PaintDisabledDC_InvalidParam:
    Error.Raise giINVALID_PICTURE
    Exit Sub
End Sub

Private Sub PaintDisabledDC(ByVal hdcDest As Long, _
                                ByVal xDest As Long, _
                                ByVal yDest As Long, _
                                ByVal Width As Long, _
                                ByVal Height As Long, _
                                ByVal hdcSrc As Long, _
                                ByVal xSrc As Long, _
                                ByVal ySrc As Long, _
                                Optional ByVal clrMask As OLE_COLOR = vbWhite, _
                                Optional ByVal clrHighlight As OLE_COLOR = vb3DHighlight, _
                                Optional ByVal clrShadow As OLE_COLOR = vb3DShadow, _
                                Optional ByVal hPal As Long = 0)
    Dim hdcScreen As Long
    Dim hbmMonoSection As Long
    Dim hbmMonoSectionSav As Long
    Dim hdcMonoSection As Long
    Dim hdcColor As Long
    Dim hdcDisabled As Long
    Dim hbmDisabledSav As Long
    Dim lpbi As BITMAPINFO
    Dim hbmMono As Long
    Dim hdcMono As Long
    Dim hbmMonoSav As Long
    Dim lMaskColor As Long
    Dim lMaskColorCompare As Long
    Dim hdcMaskedSource As Long
    Dim hbmMasked As Long
    Dim hbmMaskedOld As Long
    Dim hpalMaskedOld As Long
    Dim hpalDisabledOld As Long
    Dim hpalMonoOld As Long
    Dim rgbBlack As RGBQUAD
    Dim rgbWhite As RGBQUAD
    Dim dwSys3dShadow As Long
    Dim dwSys3dHighlight As Long
    Dim pvBits As Long
    Dim rgbnew(1) As RGBQUAD
    Dim hbmDisabled As Long
    Dim lMonoBkGrnd As Long
    Dim lMonoBkGrndChoices(2) As Long
    Dim lIndex As Long  'For ... Next index
    Dim hbrWhite As Long
    Dim udtRect As RECT
    
    'TODO: handle pictures with dark masks
    If hPal = 0 Then
        hPal = m_hpalHalftone
    End If
  ' Define some colors
    OleTranslateColor clrShadow, hPal, dwSys3dShadow
    OleTranslateColor clrHighlight, hPal, dwSys3dHighlight
    
    hdcScreen = GetDC(0&)
    With rgbBlack
        .rgbBlue = 0
        .rgbGreen = 0
        .rgbRed = 0
        .rgbReserved = 0
    End With
    With rgbWhite
        .rgbBlue = 255
        .rgbGreen = 255
        .rgbRed = 255
        .rgbReserved = 255
    End With

    ' The first step is to create a monochrome bitmap with two colors:
    ' white where colors in the original are light, and black
    ' where the original is dark.  We can't simply bitblt to a bitmap.
    ' Instead, we create a monochrome (bichrome?) DIB section and bitblt
    ' to that.  Windows will do the conversion automatically based on the
    ' DIB section's palette.  (I.e. using a DIB section, Windows knows how
    ' to map "light" colors and "dark" colors to white/black, respectively.
    With lpbi.bmiHeader
        .biSize = LenB(lpbi.bmiHeader)
        .biWidth = Width
        .biHeight = -Height
        .biPlanes = 1
        .biBitCount = 1         ' monochrome
        .biCompression = BI_RGB
        .biSizeImage = 0
        .biXPelsPerMeter = 0
        .biYPelsPerMeter = 0
        .biClrUsed = 0          ' max colors used (2^1 = 2)
        .biClrImportant = 0     ' all (both :-]) colors are important
    End With
    With lpbi
        .bmiColors(0) = rgbBlack
        .bmiColors(1) = rgbWhite
    End With

    hbmMonoSection = CreateDIBSection(hdcScreen, lpbi, DIB_RGB_COLORS, pvBits, 0&, 0)
    
    hdcMonoSection = CreateCompatibleDC(hdcScreen)
    hbmMonoSectionSav = SelectObject(hdcMonoSection, hbmMonoSection)
    
    'Bitblt to the Monochrome DIB section
    'If a mask color is provided, create a new bitmap and copy the source
    'to it transparently.  If we don't do this, a dark mask color will be
    'turned into the outline part of the monochrome DIB section
    'Convert mask color and white before comparing
    'because the Mask color might be a system color that would be evaluated
    'to white.
    OleTranslateColor vbWhite, hPal, lMaskColorCompare
    OleTranslateColor clrMask, hPal, lMaskColor
    If lMaskColor = lMaskColorCompare Then
        BitBlt hdcMonoSection, 0, 0, Width, Height, hdcSrc, xSrc, ySrc, vbSrcCopy
    Else
        hbmMasked = CreateCompatibleBitmap(hdcScreen, Width, Height)
        hdcMaskedSource = CreateCompatibleDC(hdcScreen)
        hbmMaskedOld = SelectObject(hdcMaskedSource, hbmMasked)
        hpalMaskedOld = SelectPalette(hdcMaskedSource, hPal, True)
        RealizePalette hdcMaskedSource
        'Fill the bitmap with white
        With udtRect
            .Left = 0
            .Top = 0
            .Right = Width
            .Bottom = Height
        End With
        hbrWhite = CreateSolidBrush(vbWhite)
        FillRect hdcMaskedSource, udtRect, hbrWhite
        DeleteObject hbrWhite
        'Do the transparent paint
        PaintTransparentDC hdcMaskedSource, 0, 0, Width, Height, hdcSrc, xSrc, ySrc, lMaskColor, hPal
        'BitBlt to the Mono DIB section.  The mask color has been turned to white.
        BitBlt hdcMonoSection, 0, 0, Width, Height, hdcMaskedSource, 0, 0, vbSrcCopy
        'Clean up
        SelectPalette hdcMaskedSource, hpalMaskedOld, True
        RealizePalette hdcMaskedSource
        DeleteObject SelectObject(hdcMaskedSource, hbmMaskedOld)
        DeleteDC hdcMaskedSource
    End If
      
    ' Okay, we've got our B&W DIB section.
    ' Now that we have our monochrome bitmap, the final appearance that we
    ' want is this:  First, think of the black portion of the monochrome
    ' bitmap as our new version of the original bitmap.  We want to have a dark
    ' gray version of this with a light version underneath it, shifted down and
    ' to the right.  The light acts as a highlight, and it looks like the original
    ' image is a gray inset.
    
    ' First, create a copy of the destination.  Draw the light gray transparently,
    ' and then draw the dark gray transparently
    
    hbmDisabled = CreateCompatibleBitmap(hdcScreen, Width, Height)
    
    hdcDisabled = CreateCompatibleDC(hdcScreen)
    hbmDisabledSav = SelectObject(hdcDisabled, hbmDisabled)
    hpalDisabledOld = SelectPalette(hdcDisabled, hPal, True)
    RealizePalette hdcDisabled
    'We used to fill the background with gray, instead copy the
    'destination to memory DC.  This will allow a disabled image
    'to be drawn over a background image.
    BitBlt hdcDisabled, 0, 0, Width, Height, hdcDest, xDest, yDest, vbSrcCopy
    
    'When painting the monochrome bitmaps transparently onto the background
    'we need a background color that is not the light color of the dark color
    'Provide three choices to ensure a unique color is picked.
    OleTranslateColor vbBlack, hPal, lMonoBkGrndChoices(0)
    OleTranslateColor vbRed, hPal, lMonoBkGrndChoices(1)
    OleTranslateColor vbBlue, hPal, lMonoBkGrndChoices(2)
    
    'Pick a background color choice that doesn't match
    'the shadow or highlight color
    For lIndex = 0 To 2
        If lMonoBkGrndChoices(lIndex) <> dwSys3dHighlight And _
                lMonoBkGrndChoices(lIndex) <> dwSys3dShadow Then
            'This color can be used for a mask
            lMonoBkGrnd = lMonoBkGrndChoices(lIndex)
            Exit For
        End If
    Next

    ' Now paint a the light color shifted and transparent over the background
    ' It is not necessary to change the DIB section's color table
    ' to equal the highlight color and mask color.  In fact, setting
    ' the color table to anything besides black and white causes unpredictable
    ' results (seen in win95 with IE4, using 256 colors).
    ' Setting the Back and Text colors of the Monochrome bitmap, ensure
    ' that the desired colors are produced.
    With rgbnew(0)
        .rgbRed = (vbWhite \ 2 ^ 16) And &HFF
        .rgbGreen = (vbWhite \ 2 ^ 8) And &HFF
        .rgbBlue = vbWhite And &HFF
    End With
    With rgbnew(1)
        .rgbRed = (vbBlack \ 2 ^ 16) And &HFF
        .rgbGreen = (vbBlack \ 2 ^ 8) And &HFF
        .rgbBlue = vbBlack And &HFF
    End With
        
    SetDIBColorTable hdcMonoSection, 0, 2, rgbnew(0)
    
    '...We can't pass a DIBSection to PaintTransparentDC(), so we need to
    ' make a copy of our mono DIBSection.  Notice that we only need a monochrome
    ' bitmap, but we must set its back/fore colors to the monochrome colors we
    ' want (light gray and black), and PaintTransparentDC() will honor them.
    hbmMono = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    hdcMono = CreateCompatibleDC(hdcScreen)
    hbmMonoSav = SelectObject(hdcMono, hbmMono)
    SetMapMode hdcMono, GetMapMode(hdcSrc)
    SetBkColor hdcMono, dwSys3dHighlight
    SetTextColor hdcMono, lMonoBkGrnd
    hpalMonoOld = SelectPalette(hdcMono, hPal, True)
    RealizePalette hdcMono
    BitBlt hdcMono, 0, 0, Width, Height, hdcMonoSection, 0, 0, vbSrcCopy

    '...We can go ahead and call PaintTransparentDC with our monochrome
    ' copy
    ' Draw this transparently over the disabled bitmap
    '...Don't forget to shift right and left....
    PaintTransparentDC hdcDisabled, 1, 1, Width, Height, hdcMono, 0, 0, lMonoBkGrnd, hPal
    
    ' Now draw a transparent copy, using dark gray where the monochrome had
    ' black, and transparent elsewhere.  We'll use a transparent color of black.

    '...We can't pass a DIBSection to PaintTransparentDC(), so we need to
    ' make a copy of our mono DIBSection.  Notice that we only need a monochrome
    ' bitmap, but we must set its back/fore colors to the monochrome colors we
    ' want (dark gray and black), and PaintTransparentDC() will honor them.
    ' Use hbmMono and hdcMono; already created for first color
    SetBkColor hdcMono, dwSys3dShadow
    SetTextColor hdcMono, lMonoBkGrnd
    BitBlt hdcMono, 0, 0, Width, Height, hdcMonoSection, 0, 0, vbSrcCopy

    '...We can go ahead and call PaintTransparentDC with our monochrome
    ' copy
    ' Draw this transparently over the disabled bitmap
    PaintTransparentDC hdcDisabled, 0, 0, Width, Height, hdcMono, 0, 0, lMonoBkGrnd, hPal
    BitBlt hdcDest, xDest, yDest, Width, Height, hdcDisabled, 0, 0, vbSrcCopy
    ' Okay, we're done!
    SelectPalette hdcDisabled, hpalDisabledOld, True
    RealizePalette hdcDisabled
    DeleteObject SelectObject(hdcMonoSection, hbmMonoSectionSav)
    DeleteDC hdcMonoSection
    DeleteObject SelectObject(hdcDisabled, hbmDisabledSav)
    DeleteDC hdcDisabled
    DeleteObject SelectObject(hdcMono, hbmMonoSav)
    SelectPalette hdcMono, hpalMonoOld, True
    RealizePalette hdcMono
    DeleteDC hdcMono
    ReleaseDC 0&, hdcScreen
End Sub

Private Sub PaintTransparentDC(ByVal hdcDest As Long, _
                                    ByVal xDest As Long, _
                                    ByVal yDest As Long, _
                                    ByVal Width As Long, _
                                    ByVal Height As Long, _
                                    ByVal hdcSrc As Long, _
                                    ByVal xSrc As Long, _
                                    ByVal ySrc As Long, _
                                    ByVal clrMask As OLE_COLOR, _
                                    Optional ByVal hPal As Long = 0)
    Dim hdcMask As Long        'HDC of the created mask image
    Dim hdcColor As Long       'HDC of the created color image
    Dim hbmMask As Long        'Bitmap handle to the mask image
    Dim hbmColor As Long       'Bitmap handle to the color image
    Dim hbmColorOld As Long
    Dim hbmMaskOld As Long
    Dim hPalOld As Long
    Dim hdcScreen As Long
    Dim hdcScnBuffer As Long         'Buffer to do all work on
    Dim hbmScnBuffer As Long
    Dim hbmScnBufferOld As Long
    Dim hPalBufferOld As Long
    Dim lMaskColor As Long
    
    hdcScreen = GetDC(0&)
    'Validate palette
    If hPal = 0 Then
        hPal = m_hpalHalftone
    End If
    OleTranslateColor clrMask, hPal, lMaskColor
    
    'Create a color bitmap to server as a copy of the destination
    'Do all work on this bitmap and then copy it back over the destination
    'when it's done.
    hbmScnBuffer = CreateCompatibleBitmap(hdcScreen, Width, Height)
    'Create DC for screen buffer
    hdcScnBuffer = CreateCompatibleDC(hdcScreen)
    hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)
    hPalBufferOld = SelectPalette(hdcScnBuffer, hPal, True)
    RealizePalette hdcScnBuffer
    'Copy the destination to the screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcDest, xDest, yDest, vbSrcCopy
    
    'Create a (color) bitmap for the cover (can't use CompatibleBitmap with
    'hdcSrc, because this will create a DIB section if the original bitmap
    'is a DIB section)
    hbmColor = CreateCompatibleBitmap(hdcScreen, Width, Height)
    'Now create a monochrome bitmap for the mask
    hbmMask = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    'First, blt the source bitmap onto the cover.  We do this first
    'and then use it instead of the source bitmap
    'because the source bitmap may be
    'a DIB section, which behaves differently than a bitmap.
    '(Specifically, copying from a DIB section to a monochrome bitmap
    'does a nearest-color selection rather than painting based on the
    'backcolor and forecolor.
    hdcColor = CreateCompatibleDC(hdcScreen)
    hbmColorOld = SelectObject(hdcColor, hbmColor)
    hPalOld = SelectPalette(hdcColor, hPal, True)
    RealizePalette hdcColor
    'In case hdcSrc contains a monochrome bitmap, we must set the destination
    'foreground/background colors according to those currently set in hdcSrc
    '(because Windows will associate these colors with the two monochrome colors)
    SetBkColor hdcColor, GetBkColor(hdcSrc)
    SetTextColor hdcColor, GetTextColor(hdcSrc)
    BitBlt hdcColor, 0, 0, Width, Height, hdcSrc, xSrc, ySrc, vbSrcCopy
    'Paint the mask.  What we want is white at the transparent color
    'from the source, and black everywhere else.
    hdcMask = CreateCompatibleDC(hdcScreen)
    hbmMaskOld = SelectObject(hdcMask, hbmMask)

    'When bitblt'ing from color to monochrome, Windows sets to 1
    'all pixels that match the background color of the source DC.  All
    'other bits are set to 0.
    SetBkColor hdcColor, lMaskColor
    SetTextColor hdcColor, vbWhite
    BitBlt hdcMask, 0, 0, Width, Height, hdcColor, 0, 0, vbSrcCopy
    'Paint the rest of the cover bitmap.
    '
    'What we want here is black at the transparent color, and
    'the original colors everywhere else.  To do this, we first
    'paint the original onto the cover (which we already did), then we
    'AND the inverse of the mask onto that using the DSna ternary raster
    'operation (0x00220326 - see Win32 SDK reference, Appendix, "Raster
    'Operation Codes", "Ternary Raster Operations", or search in MSDN
    'for 00220326).  DSna [reverse polish] means "(not SRC) and DEST".
    '
    'When bitblt'ing from monochrome to color, Windows transforms all white
    'bits (1) to the background color of the destination hdc.  All black (0)
    'bits are transformed to the foreground color.
    SetTextColor hdcColor, vbBlack
    SetBkColor hdcColor, vbWhite
    BitBlt hdcColor, 0, 0, Width, Height, hdcMask, 0, 0, DSna
    'Paint the Mask to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcMask, 0, 0, vbSrcAnd
    'Paint the Color to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcColor, 0, 0, vbSrcPaint
    'Copy the screen buffer to the screen
    BitBlt hdcDest, xDest, yDest, Width, Height, hdcScnBuffer, 0, 0, vbSrcCopy
    'All done!
    DeleteObject SelectObject(hdcColor, hbmColorOld)
    SelectPalette hdcColor, hPalOld, True
    RealizePalette hdcColor
    DeleteDC hdcColor
    DeleteObject SelectObject(hdcScnBuffer, hbmScnBufferOld)
    SelectPalette hdcScnBuffer, hPalBufferOld, True
    RealizePalette hdcScnBuffer
    DeleteDC hdcScnBuffer
    
    DeleteObject SelectObject(hdcMask, hbmMaskOld)
    DeleteDC hdcMask
    ReleaseDC 0&, hdcScreen
End Sub

Private Sub PaintTransparentStdPic(ByVal hdcDest As Long, _
                                    ByVal xDest As Long, _
                                    ByVal yDest As Long, _
                                    ByVal Width As Long, _
                                    ByVal Height As Long, _
                                    ByVal picSource As Picture, _
                                    ByVal xSrc As Long, _
                                    ByVal ySrc As Long, _
                                    ByVal clrMask As OLE_COLOR, _
                                    Optional ByVal hPal As Long = 0)
    Dim hdcSrc As Long         'HDC that the source bitmap is selected into
    Dim hbmMemSrcOld As Long
    Dim hbmMemSrc As Long
    Dim udtRect As RECT
    Dim hbrMask As Long
    Dim lMaskColor As Long
    Dim hdcScreen As Long
    Dim hPalOld As Long
    'Verify that the passed picture is a Bitmap
    If picSource Is Nothing Then GoTo PaintTransparentStdPic_InvalidParam
    
    Select Case picSource.Type
        Case vbPicTypeBitmap
            hdcScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                hPal = m_hpalHalftone
            End If
            'Select passed picture into an HDC
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrcOld = SelectObject(hdcSrc, picSource.handle)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw the bitmap
            PaintTransparentDC hdcDest, xDest, yDest, Width, Height, hdcSrc, xSrc, ySrc, clrMask, hPal
            
            SelectObject hdcSrc, hbmMemSrcOld
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
        Case vbPicTypeIcon
            'Create a bitmap and select it into an DC
            hdcScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                hPal = m_hpalHalftone
            End If
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrc = CreateCompatibleBitmap(hdcScreen, Width, Height)
            hbmMemSrcOld = SelectObject(hdcSrc, hbmMemSrc)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw Icon onto DC
            udtRect.Bottom = Height
            udtRect.Right = Width
            OleTranslateColor clrMask, 0&, lMaskColor
            hbrMask = CreateSolidBrush(lMaskColor)
            FillRect hdcSrc, udtRect, hbrMask
            DeleteObject hbrMask
            DrawIcon hdcSrc, 0, 0, picSource.handle
            'Draw Transparent image
            PaintTransparentDC hdcDest, xDest, yDest, Width, Height, hdcSrc, 0, 0, lMaskColor, hPal
            'Clean up
            DeleteObject SelectObject(hdcSrc, hbmMemSrcOld)
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
        Case Else
            GoTo PaintTransparentStdPic_InvalidParam
    End Select
    Exit Sub
PaintTransparentStdPic_InvalidParam:
    Err.Raise giINVALID_PICTURE
    Exit Sub
End Sub


