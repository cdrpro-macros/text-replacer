Attribute VB_Name = "mod_menu"
Option Explicit

Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Long
Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Type RECT: Left As Long: Top As Long: Right As Long: Bottom As Long: End Type
Type POINTAPI: x As Long: y As Long: End Type
Type msg: hwnd As Long: message As Long: wParam As Long: lParam As Long: time As Long: pt As POINTAPI: End Type


    
Function PopMenuList(str1 As String, mx As Long, my As Long) As Long
    Const MF_ENABLED = &H0, TPM_LEFTALIGN = &H0, MF_SEPARATOR = &H800, MF_GRAYED = &H1
    Dim msgdata As msg, rectdata As RECT, Cursor As POINTAPI, i&, j&, last&, hMenu&, id%, junk&, s$
    
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
    i = Abs(msgdata.wParam)
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
