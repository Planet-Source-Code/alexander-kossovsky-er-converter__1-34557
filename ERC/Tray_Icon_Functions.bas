Attribute VB_Name = "Tray_Icon_Functions"
'Code for Tray Icon

Declare Function SetForegroundWindow& Lib "user32" (ByVal HWND As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Declare Sub PostMessageBynum Lib "user32" Alias "PostMessageA" (ByVal HWND As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long)

'Tray Icon API
Private Type NOTIFYICONDATA

cbSize As Long
HWND As Long
uID As Long
uFlags As Long
uCallbackMessage As Long
hIcon As Long
szTip As String * 64
End Type

Const NIM_ADD = 0
Const NIM_MODIFY = 1
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200

Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Public Icon_Tooltip As String

Private Function setNOTIFYICONDATA(HWND As Long, ID As Long, Flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA

On Error GoTo Err_Handler
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.HWND = HWND
    nidTemp.uID = ID
    nidTemp.uFlags = Flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp

Exit_Sub:
    Exit Function

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub

End Function

Public Sub DoHide()

On Error GoTo Err_Handler
    
    'Add an icon.
    '                                                                                                                       Icon to Hide -- Tray Tool Caption
    Call Shell_NotifyIconA(NIM_ADD, setNOTIFYICONDATA(ER_Converter.HWND, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, ER_Converter.Icon, Icon_Tooltip))
    ER_Converter.Visible = False
    IS_Minimized = True
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub
    
End Sub

Public Sub UnHide()

On Error GoTo Err_Handler

    'Delete an existing icon.
    Call Shell_NotifyIconA(NIM_DELETE, setNOTIFYICONDATA(HWND:=ER_Converter.HWND, ID:=vbNull, Flags:=NIF_MESSAGE Or NIF_ICON Or NIF_TIP, CallbackMessage:=vbNull, Icon:=ER_Converter.Icon, Tip:=""))
    ER_Converter.WindowState = 0
    ER_Converter.Show
    IS_Minimized = False

Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub

End Sub

Public Sub Change_Tray_Icon(Optional Icon_Tip = "ER Converter")

On Error GoTo Err_Handler
    
    
    Icon_Tooltip = Icon_Tip

    Call Shell_NotifyIconA(NIM_MODIFY, setNOTIFYICONDATA(ER_Converter.HWND, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, ER_Converter.Icon, Icon_Tooltip))

Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub

End Sub

Public Sub Destroy_Tray_Icon()
    
On Error GoTo Err_Handler
    
    Call Shell_NotifyIconA(NIM_DELETE, setNOTIFYICONDATA(HWND:=ER_Converter.HWND, ID:=vbNull, Flags:=NIF_MESSAGE Or NIF_ICON Or NIF_TIP, CallbackMessage:=vbNull, Icon:=ER_Converter.Icon, Tip:=""))
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub
    
End Sub
