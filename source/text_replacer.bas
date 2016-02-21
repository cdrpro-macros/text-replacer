Attribute VB_Name = "text_replacer"
Option Explicit

Public Const trSep$ = "#"
Public Const myVer& = 1 'Версия пресета (Должна быть в каждом поиске/замене)
Public Const macroName$ = "TextReplacer (GREP)"
Public Const macroVersion$ = "1.1 beta"

Public sPath$ 'Путь до списков
Public myListName$ 'Активный список
Public myFormatForC$
Public FindFormat$(5)
Public ChangeFormat$(5)


Sub doReplace()
    If CorelDRAW.ActiveDocument Is Nothing Then Exit Sub
    Select Case CorelDRAW.VersionMajor
        Case 15: If CorelDRAW.VersionBuild < 486 Then _
                MsgBox "Error" & vbCr & _
                "CorelDraw " & CorelDRAW.Version & vbCr & _
                "Need Version 15.0.0.486", vbCritical: Exit Sub
        Case 14: If CorelDRAW.VersionBuild < 653 Then _
                MsgBox "SP1 not found! Please install." & vbCr & _
                "CorelDraw " & CorelDRAW.Version & vbCr & _
                "Need Version 14.0.0.653", vbCritical: Exit Sub
        Case 13
        Case Else: MsgBox "Macros only for 13 and 14 version", vbCritical: Exit Sub
    End Select
    mainForm.Show 0
End Sub

