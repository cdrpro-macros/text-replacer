VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} mainForm 
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   OleObjectBlob   =   "mainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sCount&


Private Sub UserForm_Initialize()
    Me.Caption = macroName & " " & macroVersion
    Dim sPos$
'    sPos = GetSetting("TextReplacer", "Options", "Pos")
'    If Len(sPos) Then: StartUpPosition = 0: Move CSng(Split(sPos, " ")(0)), CSng(Split(sPos, " ")(1))
    
    sPath = GMSManager.UserGMSPath
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    myListName = ""
    listZamFileLOAD
    cb_Mode.AddItem "Document"
    cb_Mode.AddItem "Current page"
    cb_Mode.AddItem "Selection"
    
    If ActiveSelectionRange.Count > 0 Then _
        cb_Mode.Text = "Selection" Else cb_Mode.Text = "Document"
    labCopy.Caption = Chr(169) & " 2010 by Sancho" & vbCr & "cdrpro.ru"
    
    FindFormat(0) = "0"
    ChangeFormat(0) = "0"
End Sub

'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    SaveSetting "TextReplacer", "Options", "Pos", Left & " " & Top
'    End Sub

'====================================================================================
'==========================    «агрузка наборов замен     ===========================
'====================================================================================
Private Sub listZamFileLOAD()
    Dim sFile$, i&
    i = 1: sFile = Dir(sPath & "*.tlz")
    If sFile <> "" Then cb_listZamFile.AddItem i & "| " & sFile
    Do While sFile <> ""
        i = i + 1: sFile = Dir
        If sFile <> "" Then cb_listZamFile.AddItem i & "| " & sFile
    Loop
    End Sub

'====================================================================================
'=========================    «агрузка набора в список     ==========================
'====================================================================================
Private Sub cb_listZamFile_Change()
    If cb_listZamFile.SelText = "" Then Exit Sub
    Dim a$()
    a = Split(cb_listZamFile.SelText, "|")
    List2.Clear
    myListName = Trim(a(1))
    Dim hF&, s$
    hF = FreeFile()
    Open sPath & Trim(a(1)) For Input As #hF
    s = Input(LOF(hF), #hF)
    Close hF
    a = Split(s, vbCrLf)
    Dim i&
    For i = 0 To UBound(a)
        If Len(Trim(a(i))) > 2 Then
            Dim ab$()
            ab = Split(a(i), trSep)
            'ab(1) - number of preset
            s = ab(2) & " (" & SetFomatString(ab(4)) & ")  -  " & ab(3) & " (" & SetFomatString(ab(5)) & ")"
            List2.AddItem (s)
        End If
    Next 'i
    List2.Height = 129.1
    End Sub
    
        

'====================================================================================
'=========================    добавление поиска/замены     ==========================
'====================================================================================
Private Sub cm_add_Click()
    If myListName = "" Then Exit Sub
    If tb_find.Text = "" Then Exit Sub
    Dim hF&, s$
    hF = FreeFile()
    'открывает дл€ сравнени€
    Open sPath & myListName For Input As #hF
    s = Input(LOF(hF), #hF)
    Close hF
    Dim a$(), b$(), sn$, i&
    a = Split(s, vbCrLf)
    sn = myVer & trSep & IIf(cb_useGREP, "1", "0") & trSep & _
        tb_find.Text & trSep & tb_chenge.Text & trSep & _
        makeStringFormat(FindFormat) & trSep & makeStringFormat(ChangeFormat)
        'добавить форматирование!!!!!!!!!!!!!!!!!!!!!
    For i = 0 To UBound(a)
        If a(i) = sn Then MsgBox "Find / Change is already in the list!   ", vbCritical, macroName & " " & macroVersion: Exit Sub
    Next i
    Open sPath & myListName For Append As #hF
    Print #hF, sn
    Close hF
    cb_listZamFile_Change
    End Sub
Private Function makeStringFormat(f$()) As String
    makeStringFormat = f(0) & "|" & f(1) & "|" & f(2) & "|" & f(3) & "|" & f(4) & "|" & f(5)
    End Function
        
        
        
'====================================================================================
'===========================    удаление поиска/замены     ==========================
'====================================================================================
Private Sub cm_del_Click()
    If List2.Text = "" Then Exit Sub
    Dim msg&
    msg = MsgBox("Are you sure you want to delete item from the List?   ", vbQuestion + vbOKCancel, macroName & " " & macroVersion)
    If msg <> 1 Then Exit Sub
    Dim hF&, s$
    hF = FreeFile()
    Open sPath & myListName For Input As #hF
    s = Input(LOF(hF), #hF)
    Close hF
    Dim a$()
    a = Split(s, vbCrLf)
    s = Replace(s, a(List2.ListIndex) & vbCrLf, "", , , vbTextCompare)
    If Len(s) > 2 Then
        s = Left$(s, Len(s) - 2) 'обрезаем строку с конца
        FileSystem.Kill sPath & myListName
        Open sPath & myListName For Append As #hF
        Print #hF, s
        Close hF
    Else
        FileSystem.Kill sPath & myListName
        Open sPath & myListName For Append As #hF: Close hF
    End If
    cb_listZamFile_Change
    End Sub
        



Private Sub List2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim hF&, s$
    hF = FreeFile()
    Open sPath & myListName For Input As #hF
    s = Input(LOF(hF), #hF)
    Close hF
    Dim a$(), a2$()
    a = Split(s, vbCrLf)
    a2 = Split(a(List2.ListIndex), trSep)
    '«аполнение данных (предусмотреть версии пресетов)
    cb_useGREP.Value = a2(1)
    tb_find.Text = a2(2)
    tb_chenge.Text = a2(3)

    labFindFormat.Caption = SetFomatString(a2(4)): labFindFormat.ControlTipText = SetFomatString(a2(4))
    SetQFomat FindFormat, a2(4)
    labChangeFormat.Caption = SetFomatString(a2(5)): labChangeFormat.ControlTipText = SetFomatString(a2(5))
    SetQFomat ChangeFormat, a2(5)
    End Sub
Private Function SetFomatString(str$) As String
    Dim f$(), s$
    f = Split(str, "|")
    If f(1) <> "" Then s = f(1) & " + "
    If f(2) <> "" Then s = s & f(2) & " + "
    If f(3) <> "" Then s = s & f(3) & " + "
    If f(4) <> "" Then s = s & f(4) & " + "
    If f(5) <> "" Then s = s & f(5) & " + "
    If s <> "" Then s = Left$(s, Len(s) - 3): SetFomatString = s Else SetFomatString = "None"
    End Function
Private Sub SetQFomat(f$(), str$)
    Dim i&, s$()
    s = Split(str, "|")
    For i = 0 To UBound(s)
        f(i) = s(i)
    Next
    End Sub



Private Sub cm_newList_Click()
    Dim hF&, sName$, i&
    sName = InputBox("Name for a New List", "Name...")
    If sName = "" Then Exit Sub
    i = cb_listZamFile.ListCount + 1
    hF = FreeFile()
    Open sPath & sName & ".tlz" For Append As #hF: Close hF
    cb_listZamFile.AddItem i & "| " & sName & ".tlz"
    cb_listZamFile.Text = i & "| " & sName & ".tlz"
    End Sub
Private Sub cm_EditList_Click()
    If myListName = "" Then Exit Sub
    Dim h#
    h = Shell("Notepad.exe " & sPath & myListName, vbNormalFocus)
    MsgBox "Need to update the list after editing!", vbExclamation, macroName & " " & macroVersion
    End Sub
Private Sub cm_delList_Click()
    If cb_listZamFile.SelLength = 0 Then Exit Sub
    Dim msg&
    msg = MsgBox("Are you sure you want to delete the List?   ", vbQuestion + vbOKCancel, macroName & " " & macroVersion)
    If msg <> 1 Then Exit Sub
    FileSystem.Kill sPath & myListName
    cb_listZamFile.Clear
    listZamFileLOAD
    List2.Clear
    End Sub
    
    
    
Private Sub cm_Setfformat_Click()
    myFormatForC = "f"
    uf_SetFormat.Show
    End Sub
Private Sub cm_Delfformat_Click()
    FindFormat(0) = "0"
    labFindFormat.Caption = "None"
    labFindFormat.ControlTipText = ""
    End Sub

Private Sub cm_Setcformat_Click()
    myFormatForC = "c"
    uf_SetFormat.Show
    End Sub
Private Sub cm_Delcformat_Click()
    ChangeFormat(0) = "0"
    labChangeFormat.Caption = "None"
    labChangeFormat.ControlTipText = ""
    End Sub
    
    

        





Private Sub cm_findMenu_Click()
    Dim c&
    'On Error Resume Next
    c = PopMenuList("Beginning of line|End of line|End of Paragraph|Forced Line Break|Tab||Any Whitespace|Non-breaking Space|Any Digit|Any Word Character||Ellipsis|Em Dash|En Dash|Optional Hyphen||Double Quotation Mark Before Word|Double Quotation Mark After Word", 0, 0)
    Select Case c
    Case 1: tb_find.Text = tb_find.Text & "^"
    Case 2: tb_find.Text = tb_find.Text & "$"
    Case 3: tb_find.Text = tb_find.Text & "\r"
    Case 4: tb_find.Text = tb_find.Text & "\n"
    Case 5: tb_find.Text = tb_find.Text & "\t"
    Case 7: tb_find.Text = tb_find.Text & "\s"
    Case 8: tb_find.Text = tb_find.Text & "~s"
    Case 9: tb_find.Text = tb_find.Text & "\d"
    Case 10: tb_find.Text = tb_find.Text & "\w"
    Case 12: tb_find.Text = tb_find.Text & "~e"
    Case 13: tb_find.Text = tb_find.Text & "~_"
    Case 14: tb_find.Text = tb_find.Text & "~="
    Case 15: tb_find.Text = tb_find.Text & "~-"
    Case 17: tb_find.Text = tb_find.Text & "~{"
    Case 18: tb_find.Text = tb_find.Text & "~}"
    Case 0: Exit Sub
    End Select
    End Sub
Private Sub cm_chengeMenu_Click()
    Dim c&
    'On Error Resume Next
    c = PopMenuList("End of Paragraph|Forced Line Break|Tab||Non-breaking Space||Ellipsis|Em Dash|En Dash|Optional Hyphen||Double Quotation Mark Ђ|Double Quotation Mark ї|Double Quotation Mark У|Double Quotation Mark Ф|Double Quotation Mark " & Chr(34), 0, 0)
    Select Case c
    Case 1: tb_chenge.Text = tb_chenge.Text & "\r"
    Case 2: tb_chenge.Text = tb_chenge.Text & "\n"
    Case 3: tb_chenge.Text = tb_chenge.Text & "\t"
    Case 5: tb_chenge.Text = tb_chenge.Text & "~s"
    Case 7: tb_chenge.Text = tb_chenge.Text & "~e"
    Case 8: tb_chenge.Text = tb_chenge.Text & "~_"
    Case 9: tb_chenge.Text = tb_chenge.Text & "~="
    Case 10: tb_chenge.Text = tb_chenge.Text & "~-"
    Case 12: tb_chenge.Text = tb_chenge.Text & "Ђ"
    Case 13: tb_chenge.Text = tb_chenge.Text & "ї"
    Case 14: tb_chenge.Text = tb_chenge.Text & "У"
    Case 15: tb_chenge.Text = tb_chenge.Text & "Ф"
    Case 16: tb_chenge.Text = tb_chenge.Text & Chr(34)
    Case 0: Exit Sub
    End Select
    End Sub
Private Sub cb_useGREP_Click()
    If cb_useGREP Then
    cm_findMenu.Enabled = True: cm_chengeMenu.Enabled = True
    Else
    cm_findMenu.Enabled = False: cm_chengeMenu.Enabled = False
    End If
    End Sub








Private Sub cm_ApplyRepList_Click()
    If myListName = "" Then Exit Sub
    Dim hF&, s$
    sCount = 0
    hF = FreeFile()
    Open sPath & myListName For Input As #hF: s = Input(LOF(hF), #hF): Close hF
    Dim a$(), i&, sMsg$
    a = Split(s, vbCrLf)
    boostStart "Text replace"
    For i = 0 To UBound(a)
        If Len(Trim(a(i))) > 2 Then
            Dim ab$()
            ab = Split(a(i), trSep)
            
            SetQFomat FindFormat, ab(4)
            SetQFomat ChangeFormat, ab(5)
            'If ab(1) = "1" Then myReplaceGREP ab(2), ab(3) Else myReplaceTxt ab(2), ab(3)
            sMsg = myReplaceGREP(ab(2), ab(3))
        End If
    Next 'i
    boostFinish True
    If sMsg = "done!" Then sMsg = sCount & " replacements done!"
    MsgBox sMsg, vbInformation, macroName & " " & macroVersion
    End Sub
Private Sub cm_QuickReplace_Click()
    Dim sMsg$
    boostStart "Text replace"
    sCount = 0
    sMsg = myReplaceGREP(tb_find.Text, tb_chenge.Text)
    boostFinish True
    If sMsg = "done!" Then sMsg = sCount & " replacements done!"
    MsgBox sMsg, vbInformation, macroName & " " & macroVersion
'    If cb_useGREP Then _
'        myReplaceGREP tb_find.Text, tb_chenge.Text Else _
'        myReplaceTxt tb_find.Text, tb_chenge.Text
    End Sub
    
    
    
Private Function myFindShapes() As ShapeRange
    Select Case cb_Mode.Text
        Case "Document"
            Dim p As Page
            Dim startPage As Page
            Set startPage = ActivePage
            Dim sr As New ShapeRange
            For Each p In ActiveDocument.Pages
                p.Activate
                sr.AddRange ActivePage.FindShapes(, cdrTextShape)
            Next
            startPage.Activate
            Set myFindShapes = sr
        Case "Current page"
            Set myFindShapes = ActivePage.FindShapes(, cdrTextShape)
        Case "Selection"
            Set myFindShapes = ActiveSelectionRange
    End Select
    End Function
        
        
        
        
        
'====================================================================================
'================================    ReplaceGREP     ================================
'====================================================================================
Private Function myReplaceGREP(sFind$, sRep$) As String
    Dim sr As ShapeRange, reg As Object
    Set sr = New ShapeRange
    Set sr = myFindShapes
    If sr.Count < 1 Then myReplaceGREP = "No objects!": Exit Function
    
    If Trim(sFind) = "" Then myReplaceGREP = "Find field is empty!": Exit Function
    sFind = ConvStr(sFind)
        
    Set reg = CreateObject("VBScript.RegExp")
    reg.Pattern = sFind: reg.IgnoreCase = True: reg.MultiLine = True: reg.Global = True
    
    If InStr(1, sRep, "\\r", vbTextCompare) > 0 Then _
        sRep = Replace(sRep, "\\r", "\r") Else sRep = Replace(sRep, "\r", Chr(13))
    sRep = ConvStr(sRep, True)
    
    Dim sh As Shape
    
    '=====================================================
    For Each sh In sr
        If sh.Type = cdrTextShape Then
            Dim txt As TextRange, matches As Object, i&, str$, fi&, sRepIns$, doRep As Boolean
            Set txt = sh.Text.Story
            sRepIns = sRep
            Set matches = reg.Execute(txt)
            For i = matches.Count - 1 To 0 Step -1
                With matches(i)
                    If .FirstIndex = 0 Then fi = 0 Else fi = .FirstIndex
                    Dim tr As TextRange
                    Set tr = txt.Range(fi, fi + .Length)
                    
                    If FindFormat(0) Then
                    doRep = myCheckFF(tr, FindFormat)
                    Else
                    doRep = True
                    End If
                    
                    'изменени€ строки замены дл€ разного типа текста
                    sRepIns = ConvStrP(sRepIns, sh.Text.Type)
                    If doRep Then
                        If ChangeFormat(0) Then
                            If ChangeFormat(1) <> "" Then tr.Font = ChangeFormat(1)
                            If ChangeFormat(2) <> "" Then
                                Select Case ChangeFormat(2)
                                Case "Normal": tr.Bold = False: tr.Italic = False
                                Case "Bold": tr.Bold = True: tr.Italic = False
                                Case "Italic": tr.Bold = False: tr.Italic = True
                                Case "BoldItalic": tr.Bold = True: tr.Italic = True
                                End Select
                            End If
                            If ChangeFormat(3) <> "" Then tr.Size = CSng(ChangeFormat(3))
                            If ChangeFormat(4) <> "" Then
                                Select Case ChangeFormat(4)
                                Case "Normal": tr.Case = cdrNormalFontCase
                                Case "All Caps": tr.Case = cdrAllCapsFontCase
                                Case "Small Caps": tr.Case = cdrSmallCapsFontCase
                                End Select
                            End If
                            If ChangeFormat(5) <> "" Then tr.LineSpacing = CSng(ChangeFormat(5))
                        End If
                        tr.Text = reg.Replace(tr, sRepIns)
                        sCount = sCount + 1
                        doRep = False
                    End If
                    Set tr = Nothing
                End With
            Next 'i
            sRepIns = ""
        End If 'cdrTextShape
    Next 'sh
    myReplaceGREP = "done!"
    End Function
Private Function ConvStr(str$, Optional ByVal isReplace% = False) As String
    str = Replace(str, "~s", Chr(160))
    str = Replace(str, "~_", Chr(151))
    str = Replace(str, "~=", Chr(150))
    str = Replace(str, "~-", Chr(173))
    str = Replace(str, "~e", Chr(133))
    If isReplace Then
        'some code for Replace
    Else
        'ƒл€ поиска ===================================
        str = Replace(str, "\n", Chr(11))
        str = Replace(str, "\w", "[A-Za-zј-яа-€0-9_]")
        str = Replace(str, "~{", "[ЂУ" & Chr(34) & "]")
        str = Replace(str, "~}", "[їФ" & Chr(34) & "]")
    End If
    ConvStr = str
    End Function
Private Function ConvStrP(str$, sType As cdrTextType) As String
    If sType = cdrParagraphText Then
        If InStr(1, str, "\\t", vbTextCompare) > 0 Then _
            str = Replace(str, "\\t", "\t") Else _
            str = Replace(str, "\t", Chr(9))
        If InStr(1, str, "\\n", vbTextCompare) > 0 Then _
            str = Replace(str, "\\n", "\n") Else _
            str = Replace(str, "\n", Chr(11))
    ElseIf sType = cdrArtisticText Then
        If InStr(1, str, "\\t", vbTextCompare) > 0 Then _
            str = Replace(str, "\\t", "\t") Else _
            str = Replace(str, "\t", " ")
        If InStr(1, str, "\\n", vbTextCompare) > 0 Then _
            str = Replace(str, "\\n", "\n") Else _
            str = Replace(str, "\n", Chr(13))
    End If
    ConvStrP = str
    End Function
    
Private Function myCheckFF(tr As TextRange, f$()) As Boolean
    myCheckFF = False
    If f(1) <> "" Then If tr.Font <> f(1) Then myCheckFF = False: Exit Function
    If f(2) <> "" Then
        Select Case f(2)
        Case "Normal"
            If tr.Bold Or tr.Italic Then myCheckFF = False: Exit Function
        Case "Bold"
            If tr.Bold = False Then myCheckFF = False: Exit Function
            If tr.Bold And tr.Italic Then myCheckFF = False: Exit Function
        Case "Italic"
            If tr.Italic = False Then myCheckFF = False: Exit Function
            If tr.Bold And tr.Italic Then myCheckFF = False: Exit Function
        Case "BoldItalic"
            If tr.Bold = False And tr.Italic = False Then myCheckFF = False: Exit Function
        End Select
    End If
    If f(3) <> "" Then If tr.Size <> CSng(f(3)) Then myCheckFF = False: Exit Function
    If f(4) <> "" Then
        Select Case f(4)
        Case "Normal": If tr.Case <> cdrNormalFontCase Then myCheckFF = False: Exit Function
        Case "All Caps": If tr.Case <> cdrAllCapsFontCase Then myCheckFF = False: Exit Function
        Case "Small Caps": If tr.Case <> cdrSmallCapsFontCase Then myCheckFF = False: Exit Function
        End Select
    End If
    If f(5) <> "" Then If tr.LineSpacing <> CSng(f(5)) Then myCheckFF = False: Exit Function
    myCheckFF = True
    End Function

