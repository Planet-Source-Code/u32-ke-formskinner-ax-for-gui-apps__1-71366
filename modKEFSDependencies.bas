Attribute VB_Name = "modKEFSDependencies"
Option Explicit

' UserControl dependency module
'   1  Move the form by it's titlebar
'   2  Create transparency
'   3  URL Execution

Public Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const LWA_COLORKEY = &H1
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000

Public Enum T_WindowStyle
    Maximized = 3
    Normal = 1
    ShowOnly = 5
End Enum


Public Sub MoveForm(frm As Object)

    ReleaseCapture
    SendMessage frm.hwnd, &HA1, 2, 0&
    
End Sub

Public Function Transparency(frm As Form, R As Integer, G As Integer, B As Integer)
    
    Dim RetVal, Color As Long
    
    Color = RGB(R, G, B)
    RetVal = RetVal Or WS_EX_LAYERED
    SetWindowLong frm.hwnd, GWL_EXSTYLE, RetVal
    SetLayeredWindowAttributes frm.hwnd, Color, 0, LWA_COLORKEY
    
End Function

Public Sub ExecuteURL(Parent As Form, URL As String, WindowStyle As T_WindowStyle)
    
    ShellExecute Parent.hwnd, "Open", URL, "", "", WindowStyle
    
End Sub

