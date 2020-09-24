VERSION 5.00
Object = "*\AprjFormSkinner.vbp"
Begin VB.Form frmTest 
   BackColor       =   &H00E66916&
   BorderStyle     =   0  'None
   Caption         =   "KE TitleBar and FormSkinner"
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8700
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin KEFormSkinner.FormSkinner KEFS 
      Align           =   1  'Align Top
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   9975
      LinkVisible     =   -1  'True
      Link            =   "http://www.planet-source-code.com"
      BeginProperty LinkFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LinkForeColorOver=   564446
      LinkIcon        =   "frmTest.frx":038A
      Logo            =   "frmTest.frx":0A84
      TitleStyle      =   2
      LinkCaption     =   "Update Wizard"
      TitleIcon       =   "frmTest.frx":0F17
      TitleIconSize   =   0
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImageBorderLeft =   "frmTest.frx":12B1
      ImageBorderBottom=   "frmTest.frx":18C1
      ImageBorderRight=   "frmTest.frx":2983
      ImageCloseOnlyUp=   "frmTest.frx":300B
      ImageCloseOnlyOver=   "frmTest.frx":34E3
      ImageCloseOnlyDown=   "frmTest.frx":39F3
      ImageClose1Up   =   "frmTest.frx":3F3A
      ImageClose1Over =   "frmTest.frx":443F
      ImageClose1Down =   "frmTest.frx":4955
      ImageMinUp      =   "frmTest.frx":4E7E
      ImageMinOver    =   "frmTest.frx":52AD
      ImageMinDown    =   "frmTest.frx":573B
      ImageTBRight    =   "frmTest.frx":5BD8
      ImageTBMiddle   =   "frmTest.frx":10C8E
      ImageTBLeft     =   "frmTest.frx":1657C
      Begin VB.CommandButton Command4 
         Caption         =   "Example 4"
         Height          =   375
         Left            =   5520
         TabIndex        =   27
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Example 3"
         Height          =   375
         Left            =   4440
         TabIndex        =   21
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Example 2"
         Height          =   375
         Left            =   3360
         TabIndex        =   20
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   375
         Left            =   7320
         TabIndex        =   19
         Top             =   5040
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E66916&
         Caption         =   "Title Icon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   3240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E66916&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   480
         ScaleHeight     =   615
         ScaleWidth      =   3015
         TabIndex        =   12
         Top             =   2160
         Width           =   3015
         Begin VB.OptionButton Option3 
            BackColor       =   &H00E66916&
            Caption         =   "48x48"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Index           =   4
            Left            =   1080
            TabIndex        =   17
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00E66916&
            Caption         =   "32x32"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   16
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00E66916&
            Caption         =   "24x24"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Index           =   2
            Left            =   2160
            TabIndex        =   15
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00E66916&
            Caption         =   "20x20"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   14
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00E66916&
            Caption         =   "16x16"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E66916&
         Caption         =   "Minimize Button"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   3240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3510
         Left            =   4320
         Picture         =   "frmTest.frx":1CE26
         ScaleHeight     =   3510
         ScaleWidth      =   3945
         TabIndex        =   8
         Top             =   1080
         Width           =   3945
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "The set of pictures being used is zipped along to give you some hints about sizes and such to best fit and give the best results."
            ForeColor       =   &H00E66916&
            Height          =   615
            Left            =   360
            TabIndex        =   10
            Top             =   2040
            Width           =   3255
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "FormSkinner UserControl lets you load-in your own graphics to skin your form with custom Titlebar and borders."
            ForeColor       =   &H0010A5FF&
            Height          =   615
            Left            =   360
            TabIndex        =   9
            Top             =   720
            Width           =   3255
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E66916&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   480
         ScaleHeight     =   375
         ScaleWidth      =   3615
         TabIndex        =   4
         Top             =   4320
         Width           =   3615
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E66916&
            Caption         =   "Other Hand"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   7
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E66916&
            Caption         =   "Default"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E66916&
            Caption         =   "IE Hand"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   5
            Top             =   0
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E66916&
         Caption         =   "Both"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   3
         Top             =   1440
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E66916&
         Caption         =   "Logo Only"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   2
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E66916&
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   1440
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00089CDE&
         X1              =   360
         X2              =   8280
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Show / Hide"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0010A5FF&
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Title Icon Size"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0010A5FF&
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "-------- Main Example --------"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   360
         TabIndex        =   24
         Top             =   5040
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Weblink  MouseIcon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0010A5FF&
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Title Styles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0010A5FF&
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   1080
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'         FormSkinner UserControl v1.0  QUICK DEMO        '
'                                                         '
'                  By u32. October, 08                    '
'                   © JRE SOFT.  2008                     '
'==========================================================
' FormSkinner UserControl lets you create and load-in your
' own graphics to skin your form with custom Titlebar and
' borders. You can also put your Logo up along with the
' caption. It has pretty neat propertys.

' Check it and explore it!
'==========================================================

Private Sub Check1_Click()
    KEFS.MinButton = Check1.Value
End Sub

Private Sub Check2_Click()
    KEFS.TitleIconVisible = Check2.Value
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    frmXPYellow.Show , frmTest
End Sub

Private Sub Command3_Click()
    frmGreyRounded.Show , frmTest
End Sub

Private Sub Command4_Click()
    frmFancy.Show , frmTest
End Sub

Private Sub Option1_Click(Index As Integer)
    KEFS.TitleStyle = Option1(Index).Index
End Sub

Private Sub Option2_Click(Index As Integer)
    KEFS.LinkMouseIcon = Option2(Index).Index
End Sub

Private Sub Option3_Click(Index As Integer)
    KEFS.TitleIconSize = Option3(Index).Index
End Sub

