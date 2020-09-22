VERSION 5.00
Begin VB.UserControl C_Scroll 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipBehavior    =   0  'None
   HitBehavior     =   2  'Use Paint
   KeyPreview      =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "C_Scroll.ctx":0000
   Begin VB.Image I_B_B 
      Height          =   120
      Left            =   3600
      Picture         =   "C_Scroll.ctx":0312
      Top             =   2880
      Width           =   120
   End
   Begin VB.Image I_B_G 
      Height          =   120
      Left            =   3360
      Picture         =   "C_Scroll.ctx":2348
      Top             =   2880
      Width           =   120
   End
   Begin VB.Image I_B 
      Appearance      =   0  'Flat
      Height          =   120
      Left            =   2565
      MouseIcon       =   "C_Scroll.ctx":438E
      MousePointer    =   99  'Custom
      Picture         =   "C_Scroll.ctx":4698
      Top             =   1785
      Width           =   120
   End
   Begin VB.Image I_S 
      Height          =   165
      Left            =   1845
      Picture         =   "C_Scroll.ctx":66CE
      Top             =   1755
      Width           =   1110
   End
End
Attribute VB_Name = "C_Scroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Is_M_Down As Boolean
Private S_Min As Long
Private S_Max As Long
Private S_Value As Long

Event Changed(Value As Long)

Public Property Get Max() As Long
  Max = S_Max
End Property

Public Property Let Max(ByVal M As Long)
  S_Max = M
End Property

Public Property Get Min() As Long
  Min = S_Min
End Property

Public Property Let Min(ByVal M As Long)
  S_Min = M
End Property

Public Property Get Value() As Long
  Value = S_Value
End Property

Public Property Let Value(ByVal M As Long)
  I_B.Left = (M - 155 + 5) / ((S_Max - S_Min) / (I_S.Width - 164)) + 22
End Property




Private Sub I_B_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)

    On Error GoTo Err_Handler

    I_B.Picture = I_B_G.Picture
    Is_M_Down = True
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub
    
End Sub
 
Private Sub I_B_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

On Error GoTo Err_Handler

    Dim Temp_S As Double
    Dim S_Value As Double
    If Button = vbLeftButton And Is_M_Down = True Then
        I_B.Left = I_B.Left + X - I_B.Width / 2
        If I_B.Left < 22 Then I_B.Left = 22
        If I_B.Left > I_S.Width - 142 Then I_B.Left = I_S.Width - 142
        S_Value = (I_B.Left - 22) * ((S_Max - S_Min) / (I_S.Width - 164)) - 5
        RaiseEvent Changed(150 + Fix(S_Value))
    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub

End Sub

Private Sub I_B_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)

On Error GoTo Err_Handler

    I_B.Picture = I_B_B.Picture
    Is_M_Down = False

Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub

End Sub


Private Sub UserControl_Initialize()
    
On Error GoTo Err_Handler
    
    I_S.Left = 0
    I_S.Top = 0
    I_B.Top = 22
    I_B.Left = 22
    UserControl.Height = I_S.Height
    UserControl.Width = I_S.Width
    Is_M_Down = False

Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub

End Sub


