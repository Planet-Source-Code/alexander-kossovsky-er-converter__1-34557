VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form ER_Converter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   -165
   ClientWidth     =   6135
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "ER_Converter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "ER_Converter.frx":014A
   NegotiateMenus  =   0   'False
   Picture         =   "ER_Converter.frx":0454
   ScaleHeight     =   5640
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox C_No_Unicode 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   4440
      TabIndex        =   20
      Top             =   3000
      Width           =   255
   End
   Begin ERC.C_Progress P_Bar 
      Height          =   135
      Left            =   2880
      TabIndex        =   17
      Top             =   5280
      Width           =   2415
      _ExtentX        =   3836
      _ExtentY        =   238
   End
   Begin ERC.C_Scroll C_Scroll 
      Height          =   195
      Left            =   4380
      Top             =   45
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   344
   End
   Begin VB.CheckBox C_Real_Time 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "No Unicode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   4800
      TabIndex        =   21
      Tag             =   "Label"
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label L_AutoConversion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Disabled"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   1545
      TabIndex        =   19
      Top             =   5250
      Width           =   990
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "AutoConversion :"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   5250
      Width           =   1335
   End
   Begin MSForms.TextBox Converted_Text 
      Height          =   1575
      Left            =   360
      TabIndex        =   16
      Top             =   3360
      Width           =   5415
      VariousPropertyBits=   -1466939365
      ScrollBars      =   3
      Size            =   "9551;2778"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox Original_Text 
      Height          =   1575
      Left            =   360
      TabIndex        =   15
      Top             =   1200
      Width           =   5415
      VariousPropertyBits=   -1466939365
      ScrollBars      =   3
      Size            =   "9551;2778"
      SpecialEffect   =   3
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
   End
   Begin VB.Label C_Clear 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   3360
      MouseIcon       =   "ER_Converter.frx":FBC7
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Tag             =   "Label"
      Top             =   3000
      Width           =   360
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Real-Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Tag             =   "Label"
      Top             =   840
      Width           =   855
   End
   Begin VB.Label B_Direction 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Translit -> Cyrillic"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2760
      MouseIcon       =   "ER_Converter.frx":FED1
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   435
      Width           =   2535
   End
   Begin VB.Label DO_Convert 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Convert"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   240
      MouseIcon       =   "ER_Converter.frx":101DB
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   435
      Width           =   735
   End
   Begin VB.Label DO_ReWrite_CLipboard 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Write"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   1920
      MouseIcon       =   "ER_Converter.frx":104E5
      MousePointer    =   99  'Custom
      TabIndex        =   9
      ToolTipText     =   "Write to Clipboard"
      Top             =   435
      Width           =   615
   End
   Begin VB.Label Do_Get_ClipBoard 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Get"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   1200
      MouseIcon       =   "ER_Converter.frx":107EF
      MousePointer    =   99  'Custom
      TabIndex        =   8
      ToolTipText     =   "Get from Clipboard"
      Top             =   435
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "CB"
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   1680
      TabIndex        =   7
      Tag             =   "Label"
      Top             =   630
      Width           =   285
   End
   Begin VB.Label DO_Help 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   5520
      MouseIcon       =   "ER_Converter.frx":10AF9
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   435
      Width           =   375
   End
   Begin VB.Label O_Clear 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   3360
      MouseIcon       =   "ER_Converter.frx":10E03
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Tag             =   "Label"
      Top             =   840
      Width           =   360
   End
   Begin VB.Label Label_T 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cyrillic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Tag             =   "Label"
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label_O 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Translit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Tag             =   "Label"
      Top             =   840
      Width           =   975
   End
   Begin VB.Label DO_UnLoad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   5880
      MouseIcon       =   "ER_Converter.frx":1110D
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "Close"
      Top             =   30
      Width           =   135
   End
   Begin VB.Label B_Minimize 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   5640
      MouseIcon       =   "ER_Converter.frx":11417
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Minimize"
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Top_Caption 
      BackStyle       =   0  'Transparent
      Caption         =   "ER Converter"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   30
      Width           =   3855
   End
   Begin VB.Menu Tray_Menu 
      Caption         =   "ER Converter"
      Visible         =   0   'False
      Begin VB.Menu ER_Converter_ID 
         Caption         =   "ER Converter"
      End
      Begin VB.Menu Empty_1 
         Caption         =   "-"
      End
      Begin VB.Menu Convert_ID 
         Caption         =   "Convert Clipboard"
         Visible         =   0   'False
      End
      Begin VB.Menu Clear_ID 
         Caption         =   "Clear Clipboard"
         Visible         =   0   'False
      End
      Begin VB.Menu UnLoad_ID 
         Caption         =   "UnLoad"
      End
   End
End
Attribute VB_Name = "ER_Converter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long


Private Function Convert_To_Normal(STR As String) As String
    Dim X  As Integer
    Dim P  As String
    For X = 1 To Len(STR)
        On Error Resume Next
            If Asc(StrConv(Mid(STR, X, 1), vbUnicode)) = 81 Then
                P = P & Chr(184)
            ElseIf Asc(StrConv(Mid(STR, X, 1), vbUnicode)) = 1 Then
                P = P & Chr(168)
            Else
                P = P & Chr(Asc(StrConv(Mid(STR, X, 1), vbUnicode)) - 132)
            End If
    Next
    
    Convert_To_Normal = P
End Function


Private Function Convert_To_Russian(STR As String) As String
    Dim X  As Integer
    Dim P  As String
    For X = 1 To Len(STR)
        If Asc(StrConv(Mid(STR, X, 1), vbUnicode)) <> Asc(Mid(STR, X, 1)) Or Asc(Mid(STR, X, 1)) = 63 Then
            If Asc(StrConv(Mid(STR, X, 1), vbUnicode)) = 81 Then
                P = P & Chr(184)
            ElseIf Asc(StrConv(Mid(STR, X, 1), vbUnicode)) = 1 Then
                P = P & Chr(168)
            Else
                P = P & Chr(Asc(StrConv(Mid(STR, X, 1), vbUnicode)) + 176)
            End If
        Else
            P = P & Mid(STR, X, 1)
        End If
    Next
    
    Convert_To_Russian = P
End Function



Private Sub E_Click()

On Error GoTo Err_Handler

    UnregisterHotKey HWND, ID
    SetWindowLong HWND, GWL_WNDPROC, WinProc
    End

Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub

End Sub


Private Sub B_Direction_Click()
    If B_Direction.Caption = "Translit -> Cyrillic" Then
        B_Direction.Caption = "Cyrillic -> Translit"
        Call Change_Tray_Icon("ERC [ " & "Cyrillic -> Translit" & " ]")
        Label_O.Caption = "Cyrillic"
        Label_T.Caption = "Translit"
    Else
        B_Direction.Caption = "Translit -> Cyrillic"
        Call Change_Tray_Icon("ERC [ " & "Translit -> Cyrillic" & " ]")
        Label_O.Caption = "Translit"
        Label_T.Caption = "Cyrillic"
    End If

End Sub

Public Sub B_Minimize_Click()

On Error GoTo Err_Handler

    Unload Me

Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub

End Sub


Private Sub C_Clear_Click()
Converted_Text = ""
End Sub


Private Sub C_Scroll1_Changed(Value As Long)

End Sub

Private Sub C_Scroll_Changed(Value As Long)
On Error GoTo Err_Handler

    Transparent_Value = Value
    Call Make_Form_Transparent(Me.HWND, Transparent_Value)

Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub
End Sub

Private Sub Clear_ID_Click()
       
On Error GoTo Err_Handler

       Call Do_Clear_ClipBoard_Click

Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub

End Sub





Private Sub Convert_ID_Click()
        
On Error GoTo Err_Handler
        
        Call Do_Get_ClipBoard_Click
        Call DO_Convert_Click
        Call DO_ReWrite_CLipboard_Click

Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub

End Sub

Private Sub Do_Clear_ClipBoard_Click()

On Error GoTo Err_Handler

    Clipboard.SetText ""
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub

End Sub


Public Sub DO_Convert_Click()
    
On Error GoTo Err_Handler

    Dim Temp_Text As String
                
    Temp_Text = Original_Text.Text
    
    If Len(Temp_Text) < 1 Then Exit Sub
    
    If B_Direction = "Translit -> Cyrillic" Then
        P_Bar.Value = 0
        P_Bar.Max = UBound(L_ENG_RUS) + 1
    
        For T = LBound(L_ENG_RUS) To UBound(L_ENG_RUS)
            DoEvents
            Temp_Text = Replace(Temp_Text, Trim(L_ENG_RUS(T).EC), Trim(L_ENG_RUS(T).RC), , , vbBinaryCompare)
            
            If IS_Minimized = False Then
                If C_Real_Time <> 1 Then
                    P_Bar.Value = P_Bar.Value + 1
                End If
            End If
        
        Next
    Else
        Temp_Text = Original_Text.Text
        P_Bar.Value = 0
        P_Bar.Max = UBound(L_RUS_ENG) + 1
        
        For T = LBound(L_RUS_ENG) To UBound(L_RUS_ENG)
            DoEvents
            Temp_Text = Replace(Temp_Text, Trim(L_RUS_ENG(T).RC), Trim(L_RUS_ENG(T).EC), , , vbBinaryCompare)
            
            If IS_Minimized = False Then
                If C_Real_Time <> 1 Then
                    P_Bar.Value = P_Bar.Value + 1
                End If
            End If

        Next
    End If
    
    
   Converted_Text = Temp_Text
   
Exit_Sub:
    P_Bar.Value = 0
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub

End Sub

Public Sub Do_Get_ClipBoard_Click()
    
On Error GoTo Err_Handler

        Original_Text = ""
        Original_Text.Paste

Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub

End Sub


Private Sub DO_Help_Click()
On Error GoTo Err_Handler

 f_Help.Show vbModal, Me
   
Exit_Sub:
    P_Bar.Value = 0
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub
End Sub

Public Sub DO_ReWrite_CLipboard_Click()
    
On Error GoTo Err_Handler
    
    If C_No_Unicode.Value = 1 Then
        Clipboard.SetText Convert_To_Russian(Converted_Text)
    Else
        Converted_Text.SelStart = 0
        Converted_Text.SelLength = Len(Converted_Text)
        Converted_Text.Copy
        Converted_Text.SelLength = 0
    End If
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub
End Sub

Public Sub DO_UnLoad_Click()
    
On Error GoTo Err_Handler
    
        If Unload_Status = False Then
            Unload Me
        Else
            Call f_Msg_YN(" What would you like to do ?")
        End If
        
Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub
    
End Sub


Private Sub Form_Load()

On Error GoTo Err_Handler

    Call RoundCorners(Me)
    
    C_Scroll.Max = 256
    C_Scroll.Min = 150
    C_Scroll.Value = 256

    Icon_Tooltip = "ERC [ " & "Translit -> Cyrillic" & " ]"
    
    Dim i As Boolean
    
    
    Top_Caption.Caption = "ER Converter  v. " & App.Major & "." & App.Minor & "." & App.Revision & " by " & App.CompanyName
    Transparent_Value = 255
    Call Make_On_Top(Me.HWND, True)
    Call Make_Form_Transparent(Me.HWND, Transparent_Value)
    Unload_Status = True

    
    Dim T As Integer
    Dim X As Integer

    Dim DB As DAO.Database
    Set DB = DAO.OpenDatabase(App.Path & "\ER_Converter.mdb")
                
    Dim RS As DAO.Recordset
    Dim QD As DAO.QueryDef
    
    Set QD = DB.CreateQueryDef("", "SELECT Len([E_Char]) AS LEN_CHAR, t_ER.E_Char, t_ER.R_Char " & _
                                          "FROM t_ER " & _
                                          "WHERE (((t_ER.Set_ID)=1001)) " & _
                                          "ORDER BY Len([E_Char]) DESC,t_ER.E_Char;")
    
    Set RS = QD.OpenRecordset()
    
    ReDim L_ENG_RUS(RS.RecordCount) As FL
    
    T = 0
    With RS
        Do While Not .EOF
           L_ENG_RUS(T).EC = !E_char
           L_ENG_RUS(T).RC = !R_Char
            .MoveNext
            T = T + 1
        Loop
    End With
    
    RS.Close
    Set RS = Nothing
    QD.Close
    Set QD = Nothing
    
    Set QD = DB.CreateQueryDef("", "SELECT Len([E_Char]) AS LEN_CHAR, t_ER.E_Char, t_ER.R_Char " & _
                                          "FROM t_ER " & _
                                          "WHERE (((t_ER.Set_ID)=2001)) " & _
                                          "ORDER BY Len([E_Char]) DESC, t_ER.R_Char;")
    
    Set RS = QD.OpenRecordset()
    
    ReDim L_RUS_ENG(RS.RecordCount) As FL
    
    T = 0
    With RS
        Do While Not .EOF
           L_RUS_ENG(T).EC = !E_char
           L_RUS_ENG(T).RC = !R_Char
            .MoveNext
            T = T + 1
        Loop
    End With
    
    RS.Close
    Set RS = Nothing
    QD.Close
    Set QD = Nothing
    DB.Close
    Set DB = Nothing
    
    
    If UCase(Command$) = "HOTKEY" Then
        L_AutoConversion.Caption = "CTRL + F12"
        i = RegisterHotKey(HWND, ID, MOD_CONTROL, VK_F12)
        If Not i Then
            Unload Me
        Else
            WinProc = SetWindowLong(HWND, GWL_WNDPROC, AddressOf ProcessWin)
        End If
        DoHide
    End If
   
   
Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

On Error GoTo Err_Handler
    
    If IS_Minimized = True Then
        Select Case X
            Case 7755
                ' Right MouseUp
                PopupMenu Me.Tray_Menu, vbPopupMenuRightButton
        End Select
    End If


    If Button = 1 Then
        Call ReleaseCapture
        Temp_Return = SendMessage(Me.HWND, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub

End Sub






Private Sub Form_Unload(Cancel As Integer)

On Error GoTo Err_Handler

    Cancel = Unload_Status
    If Unload_Status = True Then
        DoHide
        Exit Sub
    End If
    
    If Command$ = "HOTKEY" Then
        'unregister hotkey to free up resource
        UnregisterHotKey HWND, ID
        'return control of the messages back to windows before the program exits
        SetWindowLong HWND, GWL_WNDPROC, WinProc
    End If
    
    Call Destroy_Tray_Icon
    End

Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub

End Sub




Private Sub Label1_Click()
    Text1 = Get_Clipboard
End Sub



Private Sub O_Clear_Click()
    Original_Text = ""
End Sub



Private Sub Original_Text_Change()

On Error GoTo Err_Handler
    
    If C_Real_Time = 1 Then
        Call DO_Convert_Click
    End If
        
Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub
    
End Sub








Private Sub Top_Caption_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

On Error GoTo Err_Handler

    If Button = 1 Then
        Call ReleaseCapture
        Temp_Return = SendMessage(Me.HWND, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub

End Sub


Private Sub UnLoad_ID_Click()
        
On Error GoTo Err_Handler

        Unload_Status = False
        Call DO_UnLoad_Click
 
Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub
 
End Sub

Private Sub ER_Converter_ID_Click()
        
On Error GoTo Err_Handler

    UnHide
 
Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub
 
End Sub

