Attribute VB_Name = "MainModule"
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Const HKEY_CLASSES_ROOT = &H80000000
Const REG_SZ = 1
Public fMainForm As frmMain
Public MyWord As String
Type POINTAPI ' Declare types

X As Long

Y As Long

End Type

Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long ' Declare API

Dim QQDir As String
Dim QQNumber As String

Private fSplash As frmSplash
Sub Main()
    Set fSplash = New frmSplash
    Load fSplash
    fSplash.Show
    fSplash.Refresh
    Set fMainForm = New frmMain
    Load fMainForm
    Unload fSplash
    Set fSplash = Nothing
    fMainForm.Show
    App.HelpFile = App.Path & "\qpdhelp.chm"
    
    Dim sKeyName As String
    Dim sKeyValue As String
    Dim MyReturn As Long
    Dim keyhandle As Long
    sKeyName = "test"
    sKeyValue = "QQpicDiy"
    MyReturn& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, keyhandle&)
    MyReturn& = RegSetValue&(keyhandle, "", REG_SZ, sKeyValue, 0&)
    'MsgBox MyReturn&
    sKeyName = ".qpd"
    sKeyValue = "test"
    MyReturn& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, keyhandle&)
    MyReturn& = RegSetValue&(keyhandle, "", REG_SZ, sKeyValue, 0&)
    sKeyName = "test"
    sKeyValue = App.Path & "\" & App.EXEName & " %1"
    MyReturn& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, keyhandle&)
    MyReturn& = RegSetValue&(keyhandle, "shell\open\command", REG_SZ, sKeyValue, MAX_PATH)
End Sub
Public Function isQQrun() As Long

    Dim strqq As String
    strqq = String(50, " ")
    j = 0
    Do
        hqq = FindWindowEx(0, j, "#32770", vbNullString)
        If hqq <> 0 Then
            'htext = GetDlgItem(hqq, &H321)
            i = GetWindowText(hqq, strqq, Len(strqq))
            qqno = Val(strqq)
            If IsNumeric(qqno) Then
            Else
                qqno = 0
            End If
        End If
        j = hqq
    Loop While ((qqno = 0) And (hqq <> 0))
            
    isQQrun = qqno

End Function

Public Function DelQQRec()

    Dim QQpos As Long
    
    QQDir = InputBox("请输入QQ安装目录:", "欢迎使用", "C:\Program Files\Tencent")
    If QQDir = "" Then End
    QQNumber = InputBox("请输入QQ号码:", "删除登录记录")
    If QQNumber = "" Then End
    QQpos = SearchQQ(QQNumber)
    If QQpos <> -1 Then
        If MsgBox("要同时删除" & QQNumber & "的聊天记录吗？", vbYesNo + vbInformation, "聊天记录") = vbYes Then
            DeleteChat QQNumber
            DeleteQQ QQpos
        Else
            MsgBox "没有找到" & QQNumber & "在本地的登录信息", vbOKOnly + vbQuestion, "找不到"
        End
        End If
filenotfound:
    If Err = "76" Then MsgBox "QQ的文件夹输入有误", vbOKOnly + vbCritical, "有点问题": End
    End If
    
End Function

Private Function SearchQQ(QQNumber As String) As Long
Dim QQLen As Integer, BeginPos As Long, SingleNum As String * 1, GetedNum As String
Dim i As Integer
Open QQDir & "\dat\oicq2000.cfg" For Binary As #1
BeginPos = 13
While Not EOF(1)
Get #1, BeginPos, QQLen
If QQLen = Len(QQNumber) Then
BeginPos = BeginPos + 4
For i = 1 To QQLen
Get #1, BeginPos, SingleNum
BeginPos = BeginPos + 1
If Mid(QQNumber, i, 1) <> SingleNum Then Exit For
Next
If i > Len(QQNumber) Then
SearchQQ = BeginPos - QQLen - 4
Close #1
Exit Function
Else: BeginPos = BeginPos + QQLen - i
End If
Else
BeginPos = BeginPos + 4 + QQLen
End If
Wend
SearchQQ = -1
Close #1
End Function


Private Sub DeleteQQ(WritePos As Long)
Dim Temp As Byte, QQLen As Integer, TotalNum As Byte, i As Long, FileLong As Long
Open QQDir & "\dat\oicq2000.cfg" For Binary As #2
Open QQDir & "\dat\oicq2000.tmp" For Binary As #3
Get #2, WritePos, QQLen
Get #2, 9, TotalNum
FileLong = LOF(2) - QQLen - 4
For i = 1 To WritePos - 1
Get #2, i, Temp
Put #3, i, Temp
Next
For i = WritePos To FileLong
Get #2, i + 4 + QQLen, Temp
Put #3, i, Temp
Next
Put #3, 9, TotalNum - 1
Close #2, #3
Kill QQDir & "\dat\oicq2000.cfg"
Name QQDir & "\dat\oicq2000.tmp" As QQDir & "\dat\oicq2000.cfg"
MsgBox "QQ号码为:" & QQNumber & "的信息已经清除", vbOKOnly + vbInformation, "谢谢使用"

End Sub

Private Sub DeleteChat(QQNum As String)
Kill QQDir & "\" & QQNum & "\" & "*.*"
RmDir QQDir & "\" & QQNum

End Sub
