VERSION 5.00
Object = "*\AprjFormSkinner.vbp"
Begin VB.Form frmXPYellow 
   BorderStyle     =   0  'None
   Caption         =   "XP Yellow example"
   ClientHeight    =   2325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8385
   Icon            =   "frmXPYellow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin KEFormSkinner.FormSkinner KEFS1 
      Align           =   1  'Align Top
      Height          =   2325
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   4101
      MinCloseTop     =   1
      BeginProperty LinkFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleIcon       =   "frmXPYellow.frx":030A
      TitleIconSize   =   0
      TitleCaption    =   "XP Yellow example"
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
      ColorTransp1R   =   244
      ColorTransp2G   =   198
      ColorTransp3B   =   164
      TitleFontColor  =   0
      ImageBorderLeft =   "frmXPYellow.frx":0624
      ImageBorderBottom=   "frmXPYellow.frx":09B8
      ImageBorderRight=   "frmXPYellow.frx":0D44
      ImageClose1Up   =   "frmXPYellow.frx":10DE
      ImageClose1Over =   "frmXPYellow.frx":1511
      ImageClose1Down =   "frmXPYellow.frx":1956
      ImageMinUp      =   "frmXPYellow.frx":1D92
      ImageMinOver    =   "frmXPYellow.frx":219B
      ImageMinDown    =   "frmXPYellow.frx":259E
      ImageTBRight    =   "frmXPYellow.frx":29A5
      ImageTBMiddle   =   "frmXPYellow.frx":80AF
      ImageTBLeft     =   "frmXPYellow.frx":ABA9
      Begin VB.CheckBox Check2 
         Caption         =   "Show / Hide Title Icon"
         Height          =   255
         Left            =   3240
         TabIndex        =   3
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show / Hide Minimize Button"
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   1560
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmXPYellow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    KEFS1.MinButton = Check1.Value
End Sub

Private Sub Check2_Click()
    KEFS1.TitleIconVisible = Check2.Value
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

