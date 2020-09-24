VERSION 5.00
Object = "*\AprjFormSkinner.vbp"
Begin VB.Form frmGreyRounded 
   BorderStyle     =   0  'None
   Caption         =   "Grey example"
   ClientHeight    =   4905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6075
   Icon            =   "frmGreyRounded.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin KEFormSkinner.FormSkinner KEFS2 
      Align           =   1  'Align Top
      Height          =   4905
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   8652
      LinkUnderlined  =   0   'False
      LinkTop         =   400
      LinkLeft        =   2400
      MinCloseCursor  =   2
      MinCloseTop     =   2
      LinkMouseIcon   =   2
      LinkVisible     =   -1  'True
      Link            =   "http://www.something.com"
      BeginProperty LinkFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LinkForeColorOver=   15183227
      Logo            =   "frmGreyRounded.frx":000C
      TitleStyle      =   2
      LinkCaption     =   "Look for updates"
      TitleIconVisible=   0   'False
      TitleCaption    =   "Grey example"
      MinButton       =   0   'False
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorTransp1R   =   246
      ColorTransp2G   =   239
      ColorTransp3B   =   128
      ImageBorderLeft =   "frmGreyRounded.frx":049F
      ImageBorderBottom=   "frmGreyRounded.frx":2B31
      ImageBorderRight=   "frmGreyRounded.frx":3F03
      ImageClose1Up   =   "frmGreyRounded.frx":5025
      ImageClose1Over =   "frmGreyRounded.frx":5497
      ImageClose1Down =   "frmGreyRounded.frx":5907
      ImageMinUp      =   "frmGreyRounded.frx":5D71
      ImageMinOver    =   "frmGreyRounded.frx":61DE
      ImageMinDown    =   "frmGreyRounded.frx":664D
      ImageTBRight    =   "frmGreyRounded.frx":6ABA
      ImageTBMiddle   =   "frmGreyRounded.frx":F1B8
      ImageTBLeft     =   "frmGreyRounded.frx":16A5E
      Begin VB.CheckBox Check1 
         Caption         =   "Show / Hide"
         Height          =   255
         Left            =   4440
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   10
         Top             =   3690
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         ScaleHeight     =   225
         ScaleWidth      =   345
         TabIndex        =   9
         Top             =   3330
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   3840
         Top             =   1155
         Width           =   255
      End
      Begin VB.Label Label8 
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label7 
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label6 
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label Label4 
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   1200
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmGreyRounded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    KEFS2.MinButton = Check1.Value
End Sub

Private Sub Form_Load()
    Label1.Caption = "Title Icon Visible = " & KEFS2.TitleIconVisible
    Label2.Caption = "Title Style = " & KEFS2.TitleStyle
    Label3.Caption = "WebLink Visible = " & KEFS2.LinkVisible
    Label4.Caption = "WebLink Underlined = " & KEFS2.LinkUnderlined
    Label5.Caption = "WebLink Mouse Icon = " & KEFS2.LinkMouseIcon
    Label6.Caption = "Minimize Button = " & KEFS2.MinButton
    Label7.Caption = "Link ForeColor Up    "
    Label8.Caption = "Link ForeColor Over  "
    Picture1.BackColor = KEFS2.LinkForeColorUp
    Picture2.BackColor = KEFS2.LinkForeColorOver
    KEFS2.ToolTipTextWeblink = KEFS2.Link
    Image1.Picture = KEFS2.ImageMinUp
End Sub
