VERSION 5.00
Begin VB.Form f_Error 
   BackColor       =   &H00E5E5E5&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1575
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "f_Error.frx":0000
   ScaleHeight     =   1575
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Error_Text 
      Appearance      =   0  'Flat
      BackColor       =   &H00E5E5E5&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Height          =   495
      Left            =   1200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label OK_Lable 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
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
      MouseIcon       =   "f_Error.frx":41FB
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1230
      Width           =   1215
   End
   Begin VB.Image Msg_Icon 
      Height          =   480
      Left            =   240
      Picture         =   "f_Error.frx":4505
      Top             =   480
      Width           =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   3120
      X2              =   120
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Top_Caption 
      BackStyle       =   0  'Transparent
      Caption         =   "ER Converter [ Error ]"
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
      TabIndex        =   0
      Top             =   30
      Width           =   3855
   End
End
Attribute VB_Name = "f_Error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Error_Text_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

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

Private Sub Form_Load()
    Call RoundCorners(Me)
    Call Make_On_Top(Me.HWND, True)
    Call Make_Form_Transparent(Me.HWND, Transparent_Value)
    Error_Text.Text = Temp_Error_Text_String
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

Private Sub Image1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Make_On_Top(Me.HWND, False)
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

Private Sub OK_Lable_Click()
    Unload Me
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
