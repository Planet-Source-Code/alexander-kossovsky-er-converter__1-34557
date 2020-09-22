VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form f_Help 
   BackColor       =   &H00E5E5E5&
   BorderStyle     =   0  'None
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "f_Help.frx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView ER_List 
      Height          =   2295
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   8421504
      BackColor       =   15066597
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial CYR"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Rus."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Eng."
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ListView RE_List 
      Height          =   2295
      Left            =   3120
      TabIndex        =   4
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   8421504
      BackColor       =   15066597
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial CYR"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Eng."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Rus"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Translit -> Cyrillic"
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
      Left            =   3000
      TabIndex        =   6
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cyrillic -> Translit"
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
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label DO_UnLoad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   5280
      MouseIcon       =   "f_Help.frx":607C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Close"
      Top             =   30
      Width           =   135
   End
   Begin VB.Label Top_Caption 
      BackStyle       =   0  'Transparent
      Caption         =   "ER Converter [ Help ]"
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
      TabIndex        =   2
      Top             =   30
      Width           =   3855
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
      Left            =   4200
      MouseIcon       =   "f_Help.frx":6386
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   4080
      X2              =   120
      Y1              =   3360
      Y2              =   3360
   End
End
Attribute VB_Name = "f_Help"
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

Private Sub DO_UnLoad_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim LI As ListItem
    
    Call RoundCorners(Me)
    Call Make_On_Top(Me.HWND, True)
    Call Make_Form_Transparent(Me.HWND, Transparent_Value)
    
    Dim T As Integer
    Dim X As Integer

    Dim DB As DAO.Database
    Set DB = DAO.OpenDatabase(App.Path & "\ER_Converter.mdb")
                
    Dim RS As DAO.Recordset
    Dim QD As DAO.QueryDef
    
    Set QD = DB.CreateQueryDef("", "SELECT Len([E_Char]) AS LEN_CHAR, t_ER.E_Char, t_ER.R_Char " & _
                                          "FROM t_ER " & _
                                          "WHERE (((t_ER.Set_ID)=1001)) " & _
                                          "ORDER by t_ER.R_Char")
    
    Set RS = QD.OpenRecordset()
    
    With RS
        Do While Not .EOF
            Set LI = ER_List.ListItems.Add(, , Convert_To_Russian(!R_Char))
            LI.SubItems(1) = !E_char
            .MoveNext
        Loop
    End With
    
    RS.Close
    Set RS = Nothing
    QD.Close
    Set QD = Nothing
    
    Set QD = DB.CreateQueryDef("", "SELECT Len([E_Char]) AS LEN_CHAR, t_ER.E_Char, t_ER.R_Char " & _
                                          "FROM t_ER " & _
                                          "WHERE (((t_ER.Set_ID)=2001)) " & _
                                          "ORDER by t_ER.E_Char")
    
    Set RS = QD.OpenRecordset()
    
    With RS
        Do While Not .EOF
            Set LI = RE_List.ListItems.Add(, , !E_char)
            LI.SubItems(1) = Convert_To_Russian(!R_Char)
            .MoveNext
        Loop
    End With

    RS.Close
    Set RS = Nothing
    QD.Close
    Set QD = Nothing
    DB.Close
    Set DB = Nothing
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

