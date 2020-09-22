Attribute VB_Name = "HotKey_Functions"
Public Declare Function RegisterHotKey Lib "user32" (ByVal HWND As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal HWND As Long, ByVal ID As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Const VK_F12& = &H7B
Public Const VK_F11& = &H7A
Public Const VK_CAPITAL = &H14
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_EREOF = &HF9
Public Const VK_HOME = &H24
Public Const VK_LCONTROL = &HA2
Public Const VK_LSHIFT = &HA0
Public Const VK_NUMLOCK = &H90
Public Const VK_RCONTROL = &HA3
Public Const VK_RSHIFT = &HA1
Public Const VK_SCROLL = &H91
Public Const VK_SHIFT = &H10

Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4
Public Const MOD_ALT = &H1


Public Const ID = 0
Public Const WM_HOTKEY = &H312
Public Const GWL_WNDPROC = (-4)
Public WinProc As Long

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal HWND As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal _
lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias _
"SetWindowLongA" (ByVal HWND As Long, ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias _
"GetWindowLongA" (ByVal HWND As Long, ByVal nIndex _
As Long) As Long

Public Function ProcessWin(ByVal wnd As Long, ByVal umsg As Long, ByVal wp As Long, ByVal lp As Long) As Long
    
On Error GoTo Err_Handler
    
    If umsg = WM_HOTKEY Then
        HotKeyPressed
        ProcessWin = 1
        Exit Function
    End If
    ProcessWin = CallWindowProc(WinProc, wnd, umsg, wp, lp)
    
Exit_Sub:
    Exit Function

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub
    
End Function
Sub HotKeyPressed()
        
On Error GoTo Err_Handler
    
        Call ER_Converter.Do_Get_ClipBoard_Click
        Call ER_Converter.DO_Convert_Click
        Call ER_Converter.DO_ReWrite_CLipboard_Click
        
Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub

End Sub


