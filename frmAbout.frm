VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2475
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4170
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1708.289
   ScaleMode       =   0  'User
   ScaleWidth      =   3915.846
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1440
      TabIndex        =   0
      Top             =   1920
      Width           =   1260
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Free to use in none commercial apps."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   1440
      Width           =   2640
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KE FormSkinner ActiveX Control v1.0"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   2640
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "frmAbout.frx":08CA
      Stretch         =   -1  'True
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "© Copyright  JRE SOFT. 2005 - 2008"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   2640
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
  Unload Me
End Sub

