VERSION 5.00
Begin VB.Form f_YN 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   Picture         =   "f_YN.frx":0000
   ScaleHeight     =   1605
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   4320
      MouseIcon       =   "f_YN.frx":18B0
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "Close"
      Top             =   30
      Width           =   255
   End
   Begin VB.Label Msg_Text 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Unload_Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Unload"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   1920
      MouseIcon       =   "f_YN.frx":1BBA
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1230
      Width           =   1215
   End
   Begin VB.Label Top_Caption 
      BackStyle       =   0  'Transparent
      Caption         =   "ER Converter [ Confirmation ]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   30
      Width           =   3855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   1680
      X2              =   120
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Image Msg_Icon 
      Height          =   480
      Left            =   240
      Picture         =   "f_YN.frx":1EC4
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Minimize_Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Minimize"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   3240
      MouseIcon       =   "f_YN.frx":2606
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1230
      Width           =   1215
   End
End
Attribute VB_Name = "f_YN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Error_Text_Click()

End Sub

Private Sub DO_UnLoad_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call RoundCorners(Me)
    Call Make_On_Top(Me.HWND, True)
    Call Make_Form_Transparent(Me.HWND, Transparent_Value)
    Msg_Text.Caption = Temp_Error_Text_String
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo Err_Handler

    If Button = 1 Then
        Call ReleaseCapture
        Temp_Return = SendMessage(Me.HWND, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Resume Exit_Sub
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call Make_On_Top(Me.HWND, False)
End Sub

Private Sub Minimize_Label_Click()
    On Error Resume Next
    Call ER_Converter.B_Minimize_Click
    Unload Me
End Sub

Private Sub Msg_Icon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo Err_Handler

    If Button = 1 Then
        Call ReleaseCapture
        Temp_Return = SendMessage(Me.HWND, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Resume Exit_Sub
    
End Sub


Private Sub Msg_Text_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Err_Handler

    If Button = 1 Then
        Call ReleaseCapture
        Temp_Return = SendMessage(Me.HWND, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Resume Exit_Sub

End Sub

Private Sub Top_Caption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo Err_Handler

    If Button = 1 Then
        Call ReleaseCapture
        Temp_Return = SendMessage(Me.HWND, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Resume Exit_Sub
    
End Sub

Private Sub Unload_Label_Click()
    On Error Resume Next
    Unload_Status = False
    Unload ER_Converter
    Unload Me
End Sub
