VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.UserControl C_Progress 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "C_Progress.ctx":0000
   Begin MSForms.Image I_B 
      Height          =   120
      Left            =   1320
      Top             =   2880
      Visible         =   0   'False
      Width           =   840
      BorderStyle     =   0
      Size            =   "1482;212"
      Picture         =   "C_Progress.ctx":0312
      PictureAlignment=   0
      PictureTiling   =   -1  'True
      VariousPropertyBits=   19
   End
   Begin VB.Image I_S 
      Height          =   165
      Left            =   1440
      Picture         =   "C_Progress.ctx":2349
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1110
   End
End
Attribute VB_Name = "C_Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private S_Max As Long
Private S_Value As Long

Event Changed(Value As Long)

Public Property Get Max() As Long
  Max = S_Max
End Property

Public Property Let Max(ByVal M As Long)
  S_Max = M
End Property


Public Property Get Value() As Long
  Value = S_Value
End Property

Public Property Let Value(ByVal M As Long)
    On Error Resume Next
    
    S_Value = M
    I_B.Width = (I_S.Width / 100) * ((100 * M) / (S_Max))
    If M > 0 Then I_B.Visible = True Else I_B.Visible = False

End Property








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



Private Sub UserControl_Resize()
    On Error GoTo Err_Handler
        
        I_S.Width = UserControl.Width
        
Exit_Sub:
    Exit Sub

Err_Handler:
    Call f_Error_Msg(Err.Description, "Error " & Err.Number & " / " & Err.Source)
    Resume Exit_Sub
        
End Sub
