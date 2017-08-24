Attribute VB_Name = "text_replacer"
Option Explicit

Public Const trSep$ = "#"
Public Const myVer& = 1 'Версия пресета (Должна быть в каждом поиске/замене)
Public Const macroName$ = "TextReplacer (GREP)"
Public Const macroVersion$ = "1.2 beta"

Public sPath$ 'Путь до списков
Public myListName$ 'Активный список
Public myFormatForC$
Public FindFormat$(5)
Public ChangeFormat$(5)


Sub doReplace()
    If ActiveDocument Is Nothing Then Exit Sub
    mainForm.Show 0
End Sub

