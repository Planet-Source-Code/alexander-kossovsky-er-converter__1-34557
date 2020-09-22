Attribute VB_Name = "Main_Module"
Public IS_Minimized As Boolean
Public Unload_Status As Boolean


Public Temp_Error_Text_String As String
Public Temp_Error_Number_String As String
Public Transparent_Value As Long

Public Type FL
    EC As String * 5
    RC As String * 5
End Type

Public L_ENG_RUS() As FL
Public L_RUS_ENG() As FL


Public Function f_Error_Msg(Temp_Error As String, Temp_Error_Number As String)
    On Error Resume Next
    Temp_Error_Text_String = Temp_Error
    Temp_Error_Number_String = Temp_Error_Number
    f_Error.Show vbModal, ER_Converter
End Function

Public Function f_Msg_YN(Temp_Error As String)
    On Error Resume Next
    Temp_Error_Text_String = Temp_Error
    f_YN.Show vbModal, ER_Converter
End Function




Public Function Convert_To_Russian(STR As String) As String
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



