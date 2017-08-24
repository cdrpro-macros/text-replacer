Attribute VB_Name = "mod_menu"
Option Explicit

Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As LongPtr
Declare PtrSafe Function CreatePopupMenu Lib "user32" () As LongPtr
Declare PtrSafe Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As LongPtr, ByVal wFlags As LongPtr, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As LongPtr
Declare PtrSafe Function TrackPopupMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal wFlags As LongPtr, ByVal x As LongPtr, ByVal y As LongPtr, ByVal nReserved As LongPtr, ByVal hwnd As LongPtr, lprc As RECT) As LongPtr
Declare PtrSafe Function DestroyMenu Lib "user32" (ByVal hMenu As LongPtr) As LongPtr
Declare PtrSafe Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As msg, ByVal hwnd As LongPtr, ByVal wMsgFilterMin As LongPtr, ByVal wMsgFilterMax As LongPtr) As LongPtr
Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As LongPtr
Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As LongPtr

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type POINTAPI
    x As Long
    y As Long
End Type

Type msg
    hwnd As LongPtr
    message As LongPtr
    wParam As LongPtr
    lParam As LongPtr
    time As LongPtr
    pt As POINTAPI
End Type

Const MF_ENABLED = &H0
Const TPM_LEFTALIGN = &H0
Const MF_SEPARATOR = &H800
Const MF_GRAYED = &H1


    
Function PopMenuList(str1 As String, mx As LongPtr, my As LongPtr) As Long
    Dim msgdata As msg
    Dim rectdata As RECT
    Dim Cursor As POINTAPI
    Dim i&, j&, last&
    Dim hMenu As LongPtr
    Dim id%
    Dim junk As LongPtr
    Dim s$
    
    hMenu = CreatePopupMenu() 'Создание объекта окна меню
    id = 1 ' Счетчик, задающий значение, которое вернет функция при выборе соответствующего пункта меню
    For i = 1 To countSubString(str1, "|") + 1 ' Добавление в меню пунктов
            s = CStr(Split(str1, "|")(i - 1))
            If s = "" Then
                junk = AppendMenu(hMenu, MF_SEPARATOR, i, "") ' Добавление сепаратора
            ElseIf CStr(Split(str1, "|")(i - 1)) Like "\*" Then
                s = Right(s, Len(s) - 1)
                junk = AppendMenu(hMenu, MF_GRAYED, i, s) ' Добавление нового неактивного пункта меню
            Else
                junk = AppendMenu(hMenu, MF_ENABLED, i, CStr(Split(str1, "|")(i - 1))) ' Добавление нового пункта меню
            End If
    Next i
    If mx = 0 And my = 0 Then
        Call GetCursorPos(Cursor) ' Получение текущих координат курсора мыши
        mx = Cursor.x + 10: my = Cursor.y + 10 ' Поправка
    End If
    junk = TrackPopupMenu(hMenu, TPM_LEFTALIGN, mx, my, 0, GetActiveWindow(), rectdata) ' Визуализация объекта
    junk = GetMessage(msgdata, GetActiveWindow(), 0, 0) ' Ожидание события
    i = Abs(CLng(msgdata.wParam))
    If msgdata.message = 273 Then PopMenuList = i ' Присвоение возвращаемого значения
    Call DestroyMenu(hMenu)
End Function

Function countSubString(str1$, str2$) As Long
    Dim i&
    If Len(str1) <> 0 And Len(str2) <> 0 Then
        countSubString = 0
        For i = 1 To Len(str1)
            If Mid(str1, i, Len(str2)) = str2 Then countSubString = countSubString + 1
        Next i
    End If
End Function
