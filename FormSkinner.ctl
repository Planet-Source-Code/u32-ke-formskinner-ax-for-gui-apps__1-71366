VERSION 5.00
Begin VB.UserControl FormSkinner 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   315
   ControlContainer=   -1  'True
   PropertyPages   =   "FormSkinner.ctx":0000
   ScaleHeight     =   330
   ScaleWidth      =   315
   ToolboxBitmap   =   "FormSkinner.ctx":0035
   Begin VB.Image imgLogo 
      Height          =   330
      Left            =   0
      Top             =   600
      Width           =   765
   End
   Begin VB.Image imgClose 
      Height          =   270
      Left            =   6360
      Top             =   0
      Width           =   450
   End
   Begin VB.Image imgHand2 
      Height          =   300
      Left            =   0
      Picture         =   "FormSkinner.ctx":0347
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image closeDis 
      Height          =   270
      Left            =   1800
      Top             =   4080
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image closeOnlyDis 
      Height          =   270
      Left            =   2520
      Top             =   4080
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image leftBorder 
      Height          =   4200
      Left            =   360
      Top             =   1440
      Width           =   105
   End
   Begin VB.Image bottomBorder 
      Height          =   60
      Left            =   120
      Top             =   2160
      Width           =   5250
   End
   Begin VB.Image rightBorder 
      Height          =   4500
      Left            =   720
      Top             =   1440
      Width           =   105
   End
   Begin VB.Image closeOnlyDown 
      Height          =   270
      Left            =   2520
      Top             =   3720
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image closeOnlyOver 
      Height          =   330
      Left            =   2520
      Top             =   3360
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image closeOnlyMain 
      Height          =   270
      Left            =   2520
      Top             =   3000
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgHand 
      Height          =   300
      Left            =   0
      Picture         =   "FormSkinner.ctx":1011
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label lblWeb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000004&
      Height          =   195
      Left            =   4560
      TabIndex        =   1
      Top             =   600
      Width           =   45
   End
   Begin VB.Image imgWeb 
      Height          =   300
      Left            =   2040
      Stretch         =   -1  'True
      Top             =   840
      Width           =   300
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   210
      TabIndex        =   0
      Top             =   90
      Width           =   75
   End
   Begin VB.Image imgIcon 
      Height          =   300
      Left            =   0
      Stretch         =   -1  'True
      Top             =   840
      Width           =   300
   End
   Begin VB.Image closeDown 
      Height          =   270
      Left            =   1800
      Top             =   3720
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image minDown 
      Height          =   270
      Left            =   1320
      Top             =   3720
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image closeOver 
      Height          =   330
      Left            =   1800
      Top             =   3360
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image minOver 
      Height          =   330
      Left            =   1320
      Top             =   3360
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image closeMain 
      Height          =   270
      Left            =   1800
      Top             =   3000
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image minMain 
      Height          =   270
      Left            =   1320
      Top             =   3000
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgMin 
      Height          =   270
      Left            =   5880
      Top             =   0
      Width           =   420
   End
   Begin VB.Image imgRight 
      Height          =   795
      Left            =   3600
      Top             =   0
      Width           =   4245
   End
   Begin VB.Image imgMiddle 
      Height          =   795
      Left            =   2640
      Top             =   0
      Width           =   2130
   End
   Begin VB.Image imgLeft 
      Height          =   795
      Left            =   0
      Top             =   0
      Width           =   2520
   End
End
Attribute VB_Name = "FormSkinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'                   By u32. October, 08                   '
'                    © JRE SOFT.  2008                    '
'==========================================================
' FormSkinner UserControl lets you create and load-in your
' own graphics to skin your form with custom Titlebar and
' borders. You can also put your Logo up along with the
' caption. It has more useful propertys.

' Check it out!
'==========================================================

' Programming skills...
' You should know what Property GET, LET and SET is!
' Else it's pretty easy coding.

' TIP:
' If you use only rectangular images for Min - Max - Close,
' then just leave CloseOnlyUp/Over/Down propertys blanc.
' Those images is only necesarry if you'r using rounded
' images with transparency... Like this example!


'Enum propertys
Enum TitleStyles
    Normal = 0
    [Company Logo] = 1
    Both = 2
End Enum

Enum TitleIconSizes
    [16x16 px] = 0
    [20x20 px] = 1
    [24x24 px] = 2
    [32x32 px] = 3
    [48x48 px] = 4
End Enum

Enum LinkIconSizes
    [16x16 px] = 0
    [20x20 px] = 1
    [24x24 px] = 2
    [32x32 px] = 3
    [48x48 px] = 4
End Enum

Enum LinkMouseIcons
    Default = 0
    [IE Hand] = 1
    [Other Hand] = 2
End Enum

Enum MinCloseCursors
    Default = 0
    [IE Hand] = 1
    [Other Hand] = 2
End Enum

Enum MinCloseButtonsTop
    [Modern 0px] = 0
    [Standard 4px] = 1
    [Lower 8px] = 2
    [Lower 12px] = 3
    [Bar Centered] = 4
End Enum


'Defaults
Const mDefTitleIconVisible = True
Const mDefTitleIconSize = 1
Const mDefTitleCaption = "KE TitleBar and FormSkinner"
Const mDefMinButton = True
Const mDefColorTransp1R = 255
Const mdefColorTransp2G = 255
Const mDefColorTransp3B = 222
Const mDefTitleFontColor = vbWhite
Const mDefLinkForeColorUp = vbWhite
Const mDefLinkForeColorOver = vbBlue
Const mDefToolTipTextMin = " Minimize "
Const mDefToolTipTextClose = " Close "
Const mDefTitleStyle = 0
Const mDefLink = ""
Const mDefLinkVisible = False
Const mDefLinkCaption = ""
Const mDefLinkMouseIcon = 1
Const mDefLinkLeft = 5060
Const mDefLinkTop = 475
Const mDefLinkIconSize = 0
Const mDefMinCloseTop = 0
Const mDefMinCloseCursor = 0
Const mDefLinkUnderlined = True


'Privates
Private mTitleIconVisible As Boolean
Private mTitleIconSize As TitleIconSizes
Private mTitleCaption As String
Private mMinButton As Boolean
Private mColorTransp1R As Integer
Private mColorTransp2G As Integer
Private mColorTransp3B As Integer
Private mTitleFontColor As OLE_COLOR
Private mLinkForeColorUp As OLE_COLOR
Private mLinkForeColorOver As OLE_COLOR
Private mToolTipTextMin As String
Private mToolTipTextClose As String
Private mTitleStyle As TitleStyles
Private mLink As String
Private mLinkVisible As Boolean
Private mLinkCaption As String
Private mLinkMouseIcon As LinkMouseIcons
Private mLinkLeft As Integer
Private mLinkTop As Integer
Private mLinkIconSize As LinkIconSizes
Private mMinCloseTop As MinCloseButtonsTop
Private mMinCloseCursor As MinCloseCursors
Private mLinkUnderlined As Boolean

Private Down As Boolean


Public Property Get LinkIconSize() As LinkIconSizes
    LinkIconSize = mLinkIconSize
End Property

Public Property Let LinkIconSize(ByVal NewLinkIconSize As LinkIconSizes)
    mLinkIconSize = NewLinkIconSize
    PropertyChanged "LinkIconSize"
    UserControl_Resize
End Property

Public Property Get LinkUnderlined() As Boolean
    LinkUnderlined = mLinkUnderlined
End Property

Public Property Let LinkUnderlined(ByVal NewLinkUnderlined As Boolean)
    mLinkUnderlined = NewLinkUnderlined
    PropertyChanged "LinkUnderlined"
End Property

Public Property Get MinCloseCursor() As MinCloseCursors
    MinCloseCursor = mMinCloseCursor
End Property

Public Property Let MinCloseCursor(ByVal NewMinCloseCursor As MinCloseCursors)
    mMinCloseCursor = NewMinCloseCursor
    PropertyChanged "MinCloseCursor"
    UserControl_Resize
End Property

Public Property Get MinCloseTop() As MinCloseButtonsTop
    MinCloseTop = mMinCloseTop
End Property

Public Property Let MinCloseTop(ByVal NewMinCloseTop As MinCloseButtonsTop)
    mMinCloseTop = NewMinCloseTop
    PropertyChanged "MinCloseTop"
    UserControl_Resize
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
    UserControl.BackColor = NewBackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ToolTipTextMin() As String
    ToolTipTextMin = mToolTipTextMin
End Property

Public Property Let ToolTipTextMin(ByVal NewToolTipTextMin As String)
    mToolTipTextMin = NewToolTipTextMin
    PropertyChanged "ToolTipTextMin"
    imgMin.ToolTipText = mToolTipTextMin
End Property

Public Property Get ToolTipTextClose() As String
    ToolTipTextClose = mToolTipTextClose
End Property

Public Property Let ToolTipTextClose(ByVal newToolTipTextClose As String)
    mToolTipTextClose = newToolTipTextClose
    PropertyChanged "ToolTipTextClose"
    imgClose.ToolTipText = mToolTipTextClose
End Property

Public Property Get ToolTipTextWeblink() As String
    ToolTipTextWeblink = lblWeb.ToolTipText
End Property

Public Property Let ToolTipTextWeblink(ByVal NewToolTipTextWeblink As String)
    lblWeb.ToolTipText = NewToolTipTextWeblink
    PropertyChanged "ToolTipTextWeblink"
End Property

Public Property Get LinkMouseIcon() As LinkMouseIcons
    LinkMouseIcon = mLinkMouseIcon
End Property

Public Property Let LinkMouseIcon(ByVal NewLinkMouseIcon As LinkMouseIcons)
    mLinkMouseIcon = NewLinkMouseIcon
    PropertyChanged "LinkMouseIcon"
    UserControl_Resize
End Property

Public Property Get LinkVisible() As Boolean
    LinkVisible = mLinkVisible
End Property

Public Property Let LinkVisible(ByVal NewLinkVisible As Boolean)
    mLinkVisible = NewLinkVisible
    PropertyChanged "LinkVisible"
    UserControl_Resize
End Property

Public Property Get Link() As String
    Link = mLink
End Property

Public Property Let Link(ByVal NewLink As String)
    mLink = NewLink
    PropertyChanged "Link"
End Property

Public Property Get LinkFont() As Font
    Set LinkFont = lblWeb.Font
End Property

Public Property Set LinkFont(ByVal NewLinkFont As Font)
    Set lblWeb.Font = NewLinkFont
    PropertyChanged "LinkFont"
    UserControl_Resize
End Property

Public Property Get LinkForeColorUp() As OLE_COLOR
    LinkForeColorUp = mLinkForeColorUp
End Property

Public Property Let LinkForeColorUp(ByVal NewLinkForeColorUp As OLE_COLOR)
    mLinkForeColorUp = NewLinkForeColorUp
    PropertyChanged "LinkForeColorUp"
    lblWeb.ForeColor = mLinkForeColorUp
End Property

Public Property Get LinkForeColorOver() As OLE_COLOR
    LinkForeColorOver = mLinkForeColorOver
End Property

Public Property Let LinkForeColorOver(ByVal NewLinkForeColorOver As OLE_COLOR)
    mLinkForeColorOver = NewLinkForeColorOver
    PropertyChanged "LinkForeColorOver"
End Property

Public Property Get LinkIcon() As Picture
    Set LinkIcon = imgWeb.Picture
End Property

Public Property Set LinkIcon(ByVal NewLinkIcon As Picture)
    Set imgWeb.Picture = NewLinkIcon
    PropertyChanged "LinkIcon"
End Property

Public Property Get LinkLeft() As Integer
    LinkLeft = mLinkLeft
End Property

Public Property Let LinkLeft(ByVal NewLinkLeft As Integer)
    mLinkLeft = NewLinkLeft
    PropertyChanged "LinkLeft"
    UserControl_Resize
End Property

Public Property Get LinkTop() As Integer
    LinkTop = mLinkTop
End Property

Public Property Let LinkTop(ByVal NewLinkTop As Integer)
    mLinkTop = NewLinkTop
    PropertyChanged "LinkTop"
    UserControl_Resize
End Property

Public Property Get Logo() As Picture
    Set Logo = imgLogo.Picture
End Property

Public Property Set Logo(ByVal NewLogo As Picture)
    Set imgLogo.Picture = NewLogo
    PropertyChanged "Logo"
    UserControl_Resize
End Property

Public Property Get TitleStyle() As TitleStyles
    TitleStyle = mTitleStyle
End Property

Public Property Let TitleStyle(ByVal NewTitleStyle As TitleStyles)
    mTitleStyle = NewTitleStyle
    PropertyChanged "TitleStyle"
    UserControl_Resize
End Property

Public Property Get LinkCaption() As String
    LinkCaption = mLinkCaption
End Property

Public Property Let LinkCaption(ByVal newLinkCaption As String)
    mLinkCaption = newLinkCaption
    PropertyChanged "LinkCaption"
    lblWeb.Caption = mLinkCaption
End Property

Public Property Get TitleFontColor() As OLE_COLOR
    TitleFontColor = mTitleFontColor
End Property

Public Property Let TitleFontColor(ByVal NewTitleFontColor As OLE_COLOR)
    mTitleFontColor = NewTitleFontColor
    PropertyChanged "TitleFontColor"
    lblCaption.ForeColor = mTitleFontColor
End Property

Public Property Get ColorTransp1R() As Integer
    ColorTransp1R = mColorTransp1R
End Property

Public Property Let ColorTransp1R(ByVal NewColorTransp1R As Integer)
    mColorTransp1R = NewColorTransp1R
    PropertyChanged "ColorTransp1R"
    If mColorTransp1R > 255 Then
        mColorTransp1R = 255
        MsgBox "You can't set a higher value than 255", vbInformation, "Usercontrol"
    End If
End Property

Public Property Get ColorTransp2G() As Integer
    ColorTransp2G = mColorTransp2G
End Property

Public Property Let ColorTransp2G(ByVal NewColorTransp2G As Integer)
    mColorTransp2G = NewColorTransp2G
    PropertyChanged "ColorTransp2G"
    If mColorTransp2G > 255 Then
        mColorTransp2G = 255
        MsgBox "You can't set a higher value than 255", vbInformation, "Usercontrol"
    End If
End Property

Public Property Get ColorTransp3B() As Integer
    ColorTransp3B = mColorTransp3B
End Property

Public Property Let ColorTransp3B(ByVal NewColorTransp3B As Integer)
    mColorTransp3B = NewColorTransp3B
    PropertyChanged "ColorTransp3B"
    If mColorTransp3B > 255 Then
        mColorTransp3B = 255
        MsgBox "You can't set a higher value than 255", vbInformation, "Usercontrol"
    End If
End Property

Public Property Get MinButton() As Boolean
    MinButton = mMinButton
End Property

Public Property Let MinButton(ByVal NewMinButton As Boolean)
    mMinButton = NewMinButton
    PropertyChanged "MinButton"
    UserControl_Resize
End Property

Public Property Get TitleCaption() As String
    TitleCaption = mTitleCaption
End Property

Public Property Let TitleCaption(ByVal NewTitleCaption As String)
    mTitleCaption = NewTitleCaption
    PropertyChanged "TitleCaption"
    lblCaption.Caption = mTitleCaption
    UserControl.Parent.Caption = mTitleCaption
End Property

Public Property Get TitleFont() As Font
    Set TitleFont = lblCaption.Font
End Property

Public Property Set TitleFont(ByVal NewTitleFont As Font)
    Set lblCaption.Font = NewTitleFont
    PropertyChanged "TitleFont"
End Property

Public Property Get TitleIconVisible() As Boolean
    TitleIconVisible = mTitleIconVisible
End Property

Public Property Let TitleIconVisible(ByVal NewTitleIconVisible As Boolean)
    mTitleIconVisible = NewTitleIconVisible
    PropertyChanged "TitleIconVisible"
    UserControl_Resize
End Property

Public Property Get TitleIconSize() As TitleIconSizes
    TitleIconSize = mTitleIconSize
End Property

Public Property Let TitleIconSize(ByVal NewTitleIconSize As TitleIconSizes)
    mTitleIconSize = NewTitleIconSize
    PropertyChanged "TitleIconSize"
    UserControl_Resize
End Property

Public Property Get TitleIcon() As Picture
    Set TitleIcon = imgIcon.Picture
End Property

Public Property Set TitleIcon(ByVal NewTitleIcon As Picture)
    Set imgIcon.Picture = NewTitleIcon
    PropertyChanged "TitleIcon"
    If imgIcon.Picture = 0 Then
        mTitleIconVisible = False
    Else
        mTitleIconVisible = True
    End If
    Set UserControl.Parent.Icon = NewTitleIcon
    UserControl_Resize
End Property

Private Sub imgClose_Click()
    Unload UserControl.Parent
End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mMinButton = True Then
        imgClose = closeDown
    Else
        If closeOnlyMain.Picture = 0 Then
            imgClose = closeDown
        Else
            imgClose = closeOnlyDown
        End If
    End If
    Down = True
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Close-button uses different pics depending on if
    'minimize button is visible or not.
    If mMinButton = True Then
        If Down = False Then
            'Over state
            imgClose = closeOver
            'Up
            imgMin = minMain
        Else
            'Down
            imgClose = closeDown
            'Up
            imgMin = minMain
        End If
    Else
        If Down = False Then
            'Over state, but Close only set of images
            If closeOnlyMain.Picture = 0 Then
            imgClose = closeOver
        Else
            imgClose = closeOnlyOver
        End If
            'Up
            imgMin = minMain
        Else
            'Down
            If closeOnlyMain.Picture = 0 Then
            imgClose = closeDown
        Else
            imgClose = closeOnlyDown
        End If
            'Up
            imgMin = minMain
        End If
    End If
    imgClose.ToolTipText = mToolTipTextClose
End Sub

Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Down = False
    imgRight_MouseMove Button, Shift, X, Y
End Sub

Private Sub imgLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveForm UserControl.Parent
End Sub

Private Sub imgLogo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgRight_MouseDown Button, Shift, X, Y
End Sub

Private Sub imgMiddle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveForm UserControl.Parent
End Sub

Private Sub imgMiddle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgRight_MouseMove Button, Shift, X, Y
End Sub

Private Sub imgMin_Click()
    UserControl.Parent.WindowState = 1
End Sub

Private Sub imgMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgMin = minDown
    Down = True
End Sub

Private Sub imgMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Down = False Then
        imgMin = minOver
        imgClose = closeMain
    Else
        imgMin = minDown
        imgClose = closeMain
    End If
    imgMin.ToolTipText = mToolTipTextMin
End Sub

Private Sub imgMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Down = False
    imgRight_MouseMove Button, Shift, X, Y
End Sub

Private Sub imgRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveForm UserControl.Parent
End Sub

Private Sub imgRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mMinButton = True Then
        imgMin = minMain
        imgClose = closeMain
    Else
        If closeOnlyMain.Picture = 0 Then
            imgClose = closeMain
        Else
            imgClose = closeOnlyMain
        End If
    End If
    lblWeb.ForeColor = mLinkForeColorUp
    lblWeb.FontUnderline = False
End Sub

Private Sub imgWeb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveForm UserControl.Parent
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveForm UserControl.Parent
End Sub

Private Sub lblWeb_Click()
    ExecuteURL UserControl.Parent, mLink, Normal
End Sub

Private Sub lblWeb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblWeb.ForeColor = mLinkForeColorOver
    lblWeb.FontUnderline = mLinkUnderlined
End Sub

Private Sub lblWeb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblWeb.ForeColor = mLinkForeColorUp
    lblWeb.FontUnderline = False
End Sub

Private Sub UserControl_InitProperties()
    mTitleIconVisible = mDefTitleIconVisible
    mTitleCaption = mDefTitleCaption
    mMinButton = mDefMinButton
    lblCaption.FontName = "Ms Sans Serif"
    lblCaption.FontSize = 10
    lblCaption.FontBold = True
    mColorTransp1R = mDefColorTransp1R
    mColorTransp2G = mdefColorTransp2G
    mColorTransp3B = mDefColorTransp3B
    mTitleFontColor = mDefTitleFontColor
    mLinkCaption = mDefLinkCaption
    mTitleStyle = mDefTitleStyle
    mLinkForeColorUp = mDefLinkForeColorUp
    mLinkForeColorOver = mDefLinkForeColorOver
    lblWeb.FontName = "MS Sans Serif"
    lblWeb.FontSize = 8
    lblWeb.FontBold = True
    mLink = mDefLink
    mLinkVisible = mDefLinkVisible
    mLinkMouseIcon = mDefLinkMouseIcon
    mToolTipTextClose = mDefToolTipTextClose
    mToolTipTextMin = mDefToolTipTextMin
    mTitleIconSize = mDefTitleIconSize
    mMinCloseTop = mDefMinCloseTop
    mMinCloseCursor = mDefMinCloseCursor
    mLinkLeft = mDefLinkLeft
    mLinkTop = mDefLinkTop
    mLinkUnderlined = mDefLinkUnderlined
    mLinkIconSize = mDefLinkIconSize
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgRight_MouseMove Button, Shift, X, Y
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mLinkIconSize = PropBag.ReadProperty("LinkIconSize", mDefLinkIconSize)
    mLinkUnderlined = PropBag.ReadProperty("LinkUnderlined", mDefLinkUnderlined)
    mLinkTop = PropBag.ReadProperty("LinkTop", mDefLinkTop)
    mLinkLeft = PropBag.ReadProperty("LinkLeft", mDefLinkLeft)
    mMinCloseCursor = PropBag.ReadProperty("MinCloseCursor", mDefMinCloseCursor)
    mMinCloseTop = PropBag.ReadProperty("MinCloseTop", mDefMinCloseTop)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor)
    mToolTipTextMin = PropBag.ReadProperty("ToolTipTextMin", mDefToolTipTextMin)
    mToolTipTextClose = PropBag.ReadProperty("ToolTipTextClose", mDefToolTipTextClose)
    lblWeb.ToolTipText = PropBag.ReadProperty("ToolTipTextWeblink", "")
    mLinkMouseIcon = PropBag.ReadProperty("LinkMouseIcon", mDefLinkMouseIcon)
    mLinkVisible = PropBag.ReadProperty("LinkVisible", mDefLinkVisible)
    mLink = PropBag.ReadProperty("Link", mDefLink)
    Set lblWeb.Font = PropBag.ReadProperty("LinkFont", Ambient.Font)
    mLinkForeColorUp = PropBag.ReadProperty("LinkForeColorUp", mDefLinkForeColorUp)
    mLinkForeColorOver = PropBag.ReadProperty("LinkForeColorOver", mDefLinkForeColorOver)
    Set imgWeb.Picture = PropBag.ReadProperty("LinkIcon", Nothing)
    Set imgLogo.Picture = PropBag.ReadProperty("Logo", Nothing)
    mTitleStyle = PropBag.ReadProperty("TitleStyle", mDefTitleStyle)
    mLinkCaption = PropBag.ReadProperty("LinkCaption", mDefLinkCaption)
    Set imgIcon.Picture = PropBag.ReadProperty("TitleIcon", Nothing)
    mTitleIconVisible = PropBag.ReadProperty("TitleIconVisible", mDefTitleIconVisible)
    mTitleIconSize = PropBag.ReadProperty("TitleIconSize", mDefTitleIconSize)
    mTitleCaption = PropBag.ReadProperty("TitleCaption", mDefTitleCaption)
    mMinButton = PropBag.ReadProperty("MinButton", mDefMinButton)
    Set lblCaption.Font = PropBag.ReadProperty("TitleFont", Ambient.Font)
    mColorTransp1R = PropBag.ReadProperty("ColorTransp1R", mDefColorTransp1R)
    mColorTransp2G = PropBag.ReadProperty("ColorTransp2G", mdefColorTransp2G)
    mColorTransp3B = PropBag.ReadProperty("ColorTransp3B", mDefColorTransp3B)
    mTitleFontColor = PropBag.ReadProperty("TitleFontColor", mDefTitleFontColor)
    'Skinning pictures
    Set leftBorder.Picture = PropBag.ReadProperty("ImageBorderLeft", Nothing)
    Set bottomBorder.Picture = PropBag.ReadProperty("ImageBorderBottom", Nothing)
    Set rightBorder.Picture = PropBag.ReadProperty("ImageBorderRight", Nothing)
    Set closeOnlyMain.Picture = PropBag.ReadProperty("ImageCloseOnlyUp", Nothing)
    Set closeOnlyOver.Picture = PropBag.ReadProperty("ImageCloseOnlyOver", Nothing)
    Set closeOnlyDown.Picture = PropBag.ReadProperty("ImageCloseOnlyDown", Nothing)
    Set closeMain.Picture = PropBag.ReadProperty("ImageClose1Up", Nothing)
    Set closeOver.Picture = PropBag.ReadProperty("ImageClose1Over", Nothing)
    Set closeDown.Picture = PropBag.ReadProperty("ImageClose1Down", Nothing)
    Set minMain.Picture = PropBag.ReadProperty("ImageMinUp", Nothing)
    Set minOver.Picture = PropBag.ReadProperty("ImageMinOver", Nothing)
    Set minDown.Picture = PropBag.ReadProperty("ImageMinDown", Nothing)
    Set imgRight.Picture = PropBag.ReadProperty("ImageTBRight", Nothing)
    Set imgMiddle.Picture = PropBag.ReadProperty("ImageTBMiddle", Nothing)
    Set imgLeft.Picture = PropBag.ReadProperty("ImageTBLeft", Nothing)
End Sub

Private Sub UserControl_Resize()

    'Most of the property changing is set up here.!.
    UserControl.ScaleMode = UserControl.Parent.ScaleMode
    UserControl.Width = UserControl.Parent.Width
    UserControl.Height = UserControl.Parent.Height
    imgLeft.Left = 0
    imgMiddle.Stretch = True
    imgMiddle.Left = imgLeft.Width
    'Form is so slim that imgMiddle goes in minus.
    On Error Resume Next
    imgRight.Left = UserControl.Width - imgRight.Width
    imgMiddle.Width = UserControl.Width - imgLeft.Width - imgRight.Width
    imgMiddle.Height = imgLeft.Height
    imgMin = minMain
    imgMin.Visible = mMinButton
    Select Case mTitleIconSize
        Case 0
            imgIcon.Width = 240  ' 16x16 px
            imgIcon.Height = 240
        Case 1
            imgIcon.Width = 300  ' 20x20 px
            imgIcon.Height = 300
        Case 2
            imgIcon.Width = 360  ' 24x24 px
            imgIcon.Height = 360
        Case 3
            imgIcon.Width = 480  ' 32x32 px
            imgIcon.Height = 480
        Case 4
            imgIcon.Width = 720  ' 48x48 px
            imgIcon.Height = 720
    End Select
    ' Changes
    If imgIcon.Picture = 0 Or mTitleIconVisible = False Then
        imgIcon.Left = 120
        imgIcon.Top = 80
        lblCaption.Left = imgIcon.Left + 80
        imgIcon.Visible = mTitleIconVisible
        lblCaption.Top = 60
    ElseIf mTitleIconVisible = True Then
        imgIcon.Left = 120
        imgIcon.Top = 80
        imgIcon.Visible = True
        lblCaption.Left = imgIcon.Left + imgIcon.Width + 80
        lblCaption.Top = 60
    End If
    
    If mMinButton = False Then
        ' Using only rectangular images, so
        ' the "CloseOnly" is not not needed.
        If closeOnlyMain.Picture = 0 Then
            imgClose = closeMain
            imgClose.Left = UserControl.Width - imgClose.Width - 110
        Else
            imgClose = closeOnlyMain
            imgClose.Left = UserControl.Width - imgClose.Width - 140
        End If
    Else
        imgClose = closeMain
        imgClose.Left = UserControl.Width - imgClose.Width - 110
    End If
    Select Case mMinCloseTop
        Case 0
            imgMin.Top = 0
            imgClose.Top = 0
        Case 1
            imgMin.Top = 60 '4px
            imgClose.Top = 60
        Case 2
            imgMin.Top = 120 '8px
            imgClose.Top = 120
        Case 3
            imgMin.Top = 180 '12px
            imgClose.Top = 180
        Case 4
            imgMin.Top = imgRight.Height / 2 - imgMin.Height / 2
            imgClose.Top = imgMin.Top
    End Select
    imgMin.Left = imgClose.Left - imgMin.Width
    Select Case mTitleStyle
    Case 0
        imgLogo.Visible = False
        If mTitleIconVisible = False Then
            imgIcon.Visible = False
        Else
            imgIcon.Visible = True
        End If
        lblCaption.Visible = True
    Case 1
        imgLogo.Left = 200
        imgLogo.Top = imgLeft.Height / 2 - imgLogo.Height / 2
        imgLogo.Visible = True
        imgIcon.Visible = False
        lblCaption.Visible = False
    Case 2
        imgLogo.Left = lblCaption.Left + 50
        imgLogo.Top = lblCaption.Top + lblCaption.Height + 50
        imgLogo.Visible = True
        lblCaption.Visible = True
    End Select
    Select Case mLinkIconSize
        Case 0
            imgWeb.Width = 240  ' 16x16 px
            imgWeb.Height = 240
        Case 1
            imgWeb.Width = 300  ' 20x20 px
            imgWeb.Height = 300
        Case 2
            imgWeb.Width = 360  ' 24x24 px
            imgWeb.Height = 360
        Case 3
            imgWeb.Width = 480  ' 32x32 px
            imgWeb.Height = 480
        Case 4
            imgWeb.Width = 720  ' 48x48 px
            imgWeb.Height = 720
    End Select
    imgWeb.Left = mLinkLeft
    imgWeb.Top = mLinkTop
    lblWeb.Top = imgWeb.Top + imgWeb.Height / 2 - lblWeb.Height / 2
    lblWeb.Left = imgWeb.Left + imgWeb.Width + 120
    lblWeb.ForeColor = mLinkForeColorUp
    lblWeb.Visible = mLinkVisible
    imgWeb.Visible = mLinkVisible
    Select Case mLinkMouseIcon
        Case 0
            lblWeb.MousePointer = 0
            Set lblWeb.MouseIcon = Nothing
        Case 1
            lblWeb.MousePointer = 99
            ' Cursor with hotspot
            lblWeb.MouseIcon = imgHand.Picture
        Case 2
            lblWeb.MousePointer = 99
            ' Cursor with hotspot
            lblWeb.MouseIcon = imgHand2.Picture
    End Select
    Select Case mMinCloseCursor
        Case 0
            imgMin.MousePointer = 0
            Set imgMin.MouseIcon = Nothing
            imgClose.MousePointer = 0
            Set imgClose.MouseIcon = Nothing
        Case 1
            imgMin.MousePointer = 99
            ' Cursor with hotspot
            imgMin.MouseIcon = imgHand.Picture
            imgClose.MousePointer = 99
            imgClose.MouseIcon = imgHand.Picture
        Case 2
            imgMin.MousePointer = 99
            ' Cursor with hotspot
            imgMin.MouseIcon = imgHand2.Picture
            imgClose.MousePointer = 99
            imgClose.MouseIcon = imgHand2.Picture
    End Select
    leftBorder.Stretch = True
    leftBorder.Left = 0
    leftBorder.Top = imgLeft.Height
    leftBorder.Height = UserControl.Height - imgLeft.Height
    rightBorder.Stretch = True
    rightBorder.Left = UserControl.Width - rightBorder.Width
    rightBorder.Top = leftBorder.Top
    rightBorder.Height = leftBorder.Height
    bottomBorder.Stretch = True
    bottomBorder.Left = leftBorder.Width
    bottomBorder.Top = UserControl.Height - bottomBorder.Height
    bottomBorder.Width = UserControl.Width - leftBorder.Width - rightBorder.Width
    
End Sub

Private Sub UserControl_Show()
    TitleCaption = mTitleCaption
    LinkCaption = mLinkCaption
    TitleIconVisible = mTitleIconVisible
    'MinCloseTop = mMinCloseTop
    TitleFontColor = mTitleFontColor
    Transparency UserControl.Parent, mColorTransp1R, mColorTransp2G, ColorTransp3B
    'Give the form the same
    'Icon and Caption as your Titlebar
    UserControl.Parent.Caption = TitleCaption
    UserControl.Parent.Icon = imgIcon.Picture
    UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "LinkIconSize", mLinkIconSize, mDefLinkIconSize
    PropBag.WriteProperty "LinkUnderlined", mLinkUnderlined, mDefLinkUnderlined
    PropBag.WriteProperty "LinkTop", mLinkTop, mDefLinkTop
    PropBag.WriteProperty "LinkLeft", mLinkLeft, mDefLinkLeft
    PropBag.WriteProperty "MinCloseCursor", mMinCloseCursor, mDefMinCloseCursor
    PropBag.WriteProperty "MinCloseTop", mMinCloseTop, mDefMinCloseTop
    PropBag.WriteProperty "BackColor", UserControl.BackColor, Ambient.BackColor
    PropBag.WriteProperty "ToolTipTextMin", mToolTipTextMin, mDefToolTipTextMin
    PropBag.WriteProperty "ToolTipTextClose", mToolTipTextClose, mDefToolTipTextClose
    PropBag.WriteProperty "ToolTipTextWeblink", lblWeb.ToolTipText, ""
    PropBag.WriteProperty "LinkMouseIcon", mLinkMouseIcon, mDefLinkMouseIcon
    PropBag.WriteProperty "LinkVisible", mLinkVisible, mDefLinkVisible
    PropBag.WriteProperty "Link", mLink, mDefLink
    PropBag.WriteProperty "LinkFont", lblWeb.Font, Ambient.Font
    PropBag.WriteProperty "LinkForeColorUp", mLinkForeColorUp, mDefLinkForeColorUp
    PropBag.WriteProperty "LinkForeColorOver", mLinkForeColorOver, mDefLinkForeColorOver
    PropBag.WriteProperty "LinkIcon", imgWeb.Picture, Nothing
    PropBag.WriteProperty "Logo", imgLogo.Picture, Nothing
    PropBag.WriteProperty "TitleStyle", mTitleStyle, mDefTitleStyle
    PropBag.WriteProperty "LinkCaption", mLinkCaption, mDefLinkCaption
    PropBag.WriteProperty "TitleIcon", imgIcon.Picture, Nothing
    PropBag.WriteProperty "TitleIconVisible", mTitleIconVisible, mDefTitleIconVisible
    PropBag.WriteProperty "TitleIconSize", mTitleIconSize, mDefTitleIconSize
    PropBag.WriteProperty "TitleCaption", mTitleCaption, mDefTitleCaption
    PropBag.WriteProperty "MinButton", mMinButton, mDefMinButton
    PropBag.WriteProperty "TitleFont", lblCaption.Font, Ambient.Font
    PropBag.WriteProperty "ColorTransp1R", mColorTransp1R, mDefColorTransp1R
    PropBag.WriteProperty "ColorTransp2G", mColorTransp2G, mdefColorTransp2G
    PropBag.WriteProperty "ColorTransp3B", mColorTransp3B, mDefColorTransp3B
    PropBag.WriteProperty "TitleFontColor", mTitleFontColor, mDefTitleFontColor
    'Skinning pictures
    PropBag.WriteProperty "ImageBorderLeft", leftBorder.Picture, Nothing
    PropBag.WriteProperty "ImageBorderBottom", bottomBorder.Picture, Nothing
    PropBag.WriteProperty "ImageBorderRight", rightBorder.Picture, Nothing
    PropBag.WriteProperty "ImageCloseOnlyUp", closeOnlyMain.Picture, Nothing
    PropBag.WriteProperty "ImageCloseOnlyOver", closeOnlyOver.Picture, Nothing
    PropBag.WriteProperty "ImageCloseOnlyDown", closeOnlyDown.Picture, Nothing
    PropBag.WriteProperty "ImageClose1Up", closeMain.Picture, Nothing
    PropBag.WriteProperty "ImageClose1Over", closeOver.Picture, Nothing
    PropBag.WriteProperty "ImageClose1Down", closeDown.Picture, Nothing
    PropBag.WriteProperty "ImageMinUp", minMain.Picture, Nothing
    PropBag.WriteProperty "ImageMinOver", minOver.Picture, Nothing
    PropBag.WriteProperty "ImageMinDown", minDown.Picture, Nothing
    PropBag.WriteProperty "ImageTBRight", imgRight.Picture, Nothing
    PropBag.WriteProperty "ImageTBMiddle", imgMiddle.Picture, Nothing
    PropBag.WriteProperty "ImageTBLeft", imgLeft.Picture, Nothing
End Sub


'Skinning Pictures Propertys
Public Property Get ImageTBLeft() As Picture
    Set ImageTBLeft = imgLeft.Picture
End Property

Public Property Set ImageTBLeft(ByVal NewImageTBLeft As Picture)
    Set imgLeft.Picture = NewImageTBLeft
    PropertyChanged "ImageTBLeft"
    UserControl_Resize
End Property

Public Property Get ImageTBMiddle() As Picture
    Set ImageTBMiddle = imgMiddle.Picture
End Property

Public Property Set ImageTBMiddle(ByVal NewImageTBMiddle As Picture)
    Set imgMiddle.Picture = NewImageTBMiddle
    PropertyChanged "ImageTBMiddle"
    UserControl_Resize
End Property

Public Property Get ImageTBRight() As Picture
    Set ImageTBRight = imgRight.Picture
End Property

Public Property Set ImageTBRight(ByVal NewImageTBRight As Picture)
    Set imgRight.Picture = NewImageTBRight
    PropertyChanged "ImageTBRight"
    UserControl_Resize
End Property

Public Property Get ImageMinUp() As Picture
    Set ImageMinUp = minMain.Picture
End Property

Public Property Set ImageMinUp(ByVal NewImageMinUp As Picture)
    Set minMain.Picture = NewImageMinUp
    PropertyChanged "ImageMinUp"
    UserControl_Resize
End Property

Public Property Get ImageMinOver() As Picture
    Set ImageMinOver = minOver.Picture
End Property

Public Property Set ImageMinOver(ByVal NewImageMinOver As Picture)
    Set minOver.Picture = NewImageMinOver
    PropertyChanged "ImageMinOver"
End Property

Public Property Get ImageMinDown() As Picture
    Set ImageMinDown = minDown.Picture
End Property

Public Property Set ImageMinDown(ByVal NewImageMinDown As Picture)
    Set minDown.Picture = NewImageMinDown
    PropertyChanged "ImageMinDown"
End Property

Public Property Get ImageClose1Up() As Picture
    Set ImageClose1Up = closeMain.Picture
End Property

Public Property Set ImageClose1Up(ByVal NewImageClose1Up As Picture)
    Set closeMain.Picture = NewImageClose1Up
    PropertyChanged "ImageClose1Up"
    UserControl_Resize
End Property

Public Property Get ImageClose1Over() As Picture
    Set ImageClose1Over = closeOver.Picture
End Property

Public Property Set ImageClose1Over(ByVal NewImageClose1Over As Picture)
    Set closeOver.Picture = NewImageClose1Over
    PropertyChanged "ImageClose1Over"
End Property

Public Property Get ImageClose1Down() As Picture
    Set ImageClose1Down = closeDown.Picture
End Property

Public Property Set ImageClose1Down(ByVal NewImageClose1Down As Picture)
    Set closeDown.Picture = NewImageClose1Down
    PropertyChanged "ImageClose1Down"
End Property

Public Property Get ImageCloseOnlyUp() As Picture
    Set ImageCloseOnlyUp = closeOnlyMain.Picture
End Property

Public Property Set ImageCloseOnlyUp(ByVal NewImageCloseOnlyUp As Picture)
    Set closeOnlyMain.Picture = NewImageCloseOnlyUp
    PropertyChanged "ImageCloseOnlyUp"
    UserControl_Resize
End Property

Public Property Get ImageCloseOnlyOver() As Picture
    Set ImageCloseOnlyOver = closeOnlyOver.Picture
End Property

Public Property Set ImageCloseOnlyOver(ByVal NewImageCloseOnlyOver As Picture)
    Set closeOnlyOver = NewImageCloseOnlyOver
    PropertyChanged "ImageCloseOnlyOver"
End Property

Public Property Get ImageCloseOnlyDown() As Picture
    Set ImageCloseOnlyDown = closeOnlyDown.Picture
End Property

Public Property Set ImageCloseOnlyDown(ByVal NewImageCloseOnlyDown As Picture)
    Set closeOnlyDown.Picture = NewImageCloseOnlyDown
    PropertyChanged "ImageCloseOnlyDown"
End Property

Public Property Get ImageBorderLeft() As Picture
    Set ImageBorderLeft = leftBorder.Picture
End Property

Public Property Set ImageBorderLeft(ByVal NewImageBorderLeft As Picture)
    Set leftBorder.Picture = NewImageBorderLeft
    PropertyChanged "ImageBorderLeft"
    UserControl_Resize
End Property

Public Property Get ImageBorderBottom() As Picture
    Set ImageBorderBottom = bottomBorder.Picture
End Property

Public Property Set ImageBorderBottom(ByVal NewImageBorderBottom As Picture)
    Set bottomBorder.Picture = NewImageBorderBottom
    PropertyChanged "ImageBorderBottom"
    UserControl_Resize
End Property

Public Property Get ImageBorderRight() As Picture
    Set ImageBorderRight = rightBorder.Picture
End Property

Public Property Set ImageBorderRight(ByVal NewImageBorderRight As Picture)
    Set rightBorder.Picture = NewImageBorderRight
    PropertyChanged "ImageBorderRight"
    UserControl_Resize
End Property

Sub ShowAboutBox()
Attribute ShowAboutBox.VB_UserMemId = -552
    frmAbout.Show vbModal
    Unload frmAbout
    Set frmAbout = Nothing
End Sub
