VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_SetFormat 
   Caption         =   "Format"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6375
   OleObjectBlob   =   "uf_SetFormat.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_SetFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    cb_FontStyle.List = Array("Normal", "Bold", "Italic", "BoldItalic")
    cb_FontCase.List = Array("Normal", "All Caps", "Small Caps")
    
    Dim col As New Collection
    Dim vFont As Variant
    
    For Each vFont In FontList
        If Left(vFont, 1) <> "@" Then col.Add vFont
    Next vFont
    
    Set col = SortCollection(col)
    
    For Each vFont In col
        cb_Font.AddItem vFont
    Next vFont
    End Sub
Private Sub cb_Cancel_Click(): Unload Me: End Sub

Private Sub cb_OK_Click()
    Select Case myFormatForC
    Case "f"
        FindFormat(0) = "1"
        FindFormat(1) = cb_Font.Text
        FindFormat(2) = cb_FontStyle.Text
        FindFormat(3) = tb_Size.Text
        FindFormat(4) = cb_FontCase.Text
        FindFormat(5) = tb_LineSp.Text
        mainForm.labFindFormat.Caption = SetFomatString(FindFormat)
        mainForm.labFindFormat.ControlTipText = SetFomatString(FindFormat)
    Case "c"
        ChangeFormat(0) = "1"
        ChangeFormat(1) = cb_Font.Text
        ChangeFormat(2) = cb_FontStyle.Text
        ChangeFormat(3) = tb_Size.Text
        ChangeFormat(4) = cb_FontCase.Text
        ChangeFormat(5) = tb_LineSp.Text
        mainForm.labChangeFormat.Caption = SetFomatString(ChangeFormat)
        mainForm.labChangeFormat.ControlTipText = SetFomatString(ChangeFormat)
    End Select
    Call Unload(uf_SetFormat)
    End Sub
Private Function SetFomatString(f$()) As String
    Dim s$
    If f(1) <> "" Then s = f(1) & " + "
    If f(2) <> "" Then s = s & f(2) & " + "
    If f(3) <> "" Then s = s & f(3) & " + "
    If f(4) <> "" Then s = s & f(4) & " + "
    If f(5) <> "" Then s = s & f(5) & " + "
    If s <> "" Then s = Left$(s, Len(s) - 2): SetFomatString = s Else SetFomatString = "None"
    End Function
    
    
Public Function SortCollection(ByVal col As Collection) As Collection
    Dim nc As New Collection
    Dim v1 As Variant
    Dim v2 As Variant
    Dim bAdd As Boolean
    Dim n As Long
    
    For Each v1 In col
        bAdd = False
        n = 1
        For Each v2 In nc
            If UCase(v2) > UCase(v1) Then
                nc.Add v1, , n
                bAdd = True
                Exit For
            End If
            n = n + 1
        Next v2
        If Not bAdd Then nc.Add v1
    Next v1
    Set SortCollection = nc
End Function
