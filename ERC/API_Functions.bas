Attribute VB_Name = "API_Functions"
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Declare Function GetUserNameA Lib "advapi32.dll" _
    (ByVal lpBuffer As String, _
    nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" _
    (ByVal lpBuffer As String, nSize As Long) As Long
Private Const MAX_PATH = 260

Public Declare Sub ReleaseCapture Lib "user32" ()
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2



Public Declare Function SetForegroundWindow Lib "user32" (ByVal HWND As Long) As Long

Declare Function SetWindowPos Lib "user32" (ByVal HWND As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE


Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal HWND As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal HWND As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal HWND As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal HWND As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000


Public Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, _
    ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, _
    ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long

Private Declare Function SetWindowRgn Lib "user32" (ByVal HWND As Long, _
ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Public Declare Function CreateRoundRectRgn Lib "gdi32" _
    (ByVal RectX1 As Long, ByVal RectY1 As Long, ByVal RectX2 As Long, _
    ByVal RectY2 As Long, ByVal EllipseWidth As Long, _
    ByVal EllipseHeight As Long) As Long


Public Sub RoundCorners(ByRef FRM As Form)
    FRM.ScaleMode = vbPixels
    mlWidth = FRM.ScaleWidth
    mlHeight = FRM.ScaleHeight
    
    
    SetWindowRgn FRM.HWND, CreateRoundRectRgn(1, 1, _
                (FRM.Width / Screen.TwipsPerPixelX), (FRM.Height / Screen.TwipsPerPixelY), _
                15, 15), _
                True
    FRM.ScaleMode = vbTwips
End Sub


Public Function Get_Current_User_Name() As String
    
On Error GoTo Err_Handler
    
    Dim Temp_Return As Long
    Dim Temp_Buffer As String
    Dim Temp_Size As Long
    
    Temp_Size = MAX_PATH
    
    Temp_Buffer = Space$(MAX_PATH)
    Temp_Return = GetUserNameA(Temp_Buffer, Temp_Size)
    
    If r Then
        Get_Current_User_Name = Left$(Temp_Buffer, Temp_Size - 1&)
    End If

Exit_Function:
    Exit Function

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Function

End Function


Public Sub Make_On_Top(ByVal HWND As Long, Optional OnTop As Boolean = True)
    
On Error GoTo Err_Handler
    
    Dim r As Long
    
    If OnTop = True Then
        r = SetWindowPos(HWND, HWND_TOPMOST, _
            0&, 0&, 0&, 0&, TOPMOST_FLAGS)
    Else
        r = SetWindowPos(HWND, HWND_NOTOPMOST, _
            0&, 0&, 0&, 0&, TOPMOST_FLAGS)
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub

End Sub


Public Sub Make_Form_Transparent(ByVal HWND As Long, Optional Perc = -1)

Dim Temp_Value As Long

On Error Resume Next

    If Perc < 0 Or Perc > 255 Then
        Exit Sub
    Else
        Temp_Value = GetWindowLong(HWND, GWL_EXSTYLE)
        Temp_Value = Temp_Value Or WS_EX_LAYERED
        SetWindowLong HWND, GWL_EXSTYLE, Temp_Value
        SetLayeredWindowAttributes HWND, 0, Perc, LWA_ALPHA
    End If

End Sub






Sub Lock_Window(Temp_HWND As Long, YN As Boolean)

    Dim bLocked As Boolean
    
    If YN = True Then
            LockWindowUpdate (Temp_HWND)
    Else
        LockWindowUpdate 0
    End If

End Sub


