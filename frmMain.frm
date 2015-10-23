VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QQ贴图DIY"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8205
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   8205
   StartUpPosition =   2  '屏幕中心
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   420
      Left            =   5520
      TabIndex        =   21
      Top             =   165
      Visible         =   0   'False
      Width           =   465
      ExtentX         =   820
      ExtentY         =   741
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Frame frm_qqol 
      BackColor       =   &H00FFC0C0&
      Caption         =   "她/他/它在线么？"
      Height          =   630
      Left            =   2355
      TabIndex        =   19
      Top             =   495
      Width           =   3450
      Begin VB.CommandButton cmdQq 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "搜索"
         Height          =   295
         Left            =   1815
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "搜索此号码是否在线"
         Top             =   225
         Width           =   540
      End
      Begin VB.TextBox txtqq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   165
         TabIndex        =   20
         Text            =   "输入要查询的QQ号"
         Top             =   225
         Width           =   1620
      End
      Begin VB.Label lblqqsta 
         BackColor       =   &H00FFC0C0&
         Caption         =   "状态..."
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   2550
         TabIndex        =   23
         Top             =   255
         Width           =   720
      End
   End
   Begin VB.PictureBox PicToolbar 
      BackColor       =   &H00C0C0C0&
      Height          =   385
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   8145
      TabIndex        =   11
      Top             =   0
      Width           =   8205
      Begin VB.CommandButton CmdToolbar 
         Height          =   330
         Index           =   5
         Left            =   2070
         Picture         =   "frmMain.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "帮助"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton CmdToolbar 
         Height          =   330
         Index           =   4
         Left            =   1485
         Picture         =   "frmMain.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "切换<移动/点击>模式"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton CmdToolbar 
         Height          =   330
         Index           =   3
         Left            =   1170
         Picture         =   "frmMain.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "复制到剪贴板"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton CmdToolbar 
         Height          =   330
         Index           =   2
         Left            =   855
         Picture         =   "frmMain.frx":0B78
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "重新画过"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton CmdToolbar 
         Height          =   330
         Index           =   1
         Left            =   315
         Picture         =   "frmMain.frx":10AA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "保存"
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton CmdToolbar 
         Height          =   330
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":15DC
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "打开"
         Top             =   0
         Width           =   330
      End
   End
   Begin VB.Timer timQQrun 
      Interval        =   5000
      Left            =   6135
      Top             =   180
   End
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   4395
      Top             =   2850
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "qpd"
      DialogTitle     =   "保存为..."
      Filter          =   "QQpicDiyFile(*.qpd)|*.qpd"
   End
   Begin VB.TextBox Word 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   180
      TabIndex        =   0
      Text            =   "*"
      Top             =   720
      Width           =   690
   End
   Begin VB.Frame Frame_word 
      BackColor       =   &H00FFC0C0&
      Caption         =   "请输入填入的(中/英)字符"
      Height          =   645
      Left            =   0
      TabIndex        =   3
      Top             =   495
      Width           =   2280
      Begin VB.CommandButton CmdWord 
         Appearance      =   0  'Flat
         Height          =   295
         Left            =   855
         Picture         =   "frmMain.frx":1B0E
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "按此选择特殊字符"
         Top             =   225
         Width           =   285
      End
   End
   Begin VB.Frame Frame_en 
      BackColor       =   &H00FFC0C0&
      Height          =   2415
      Left            =   0
      TabIndex        =   1
      Top             =   1170
      Width           =   8200
      Begin VB.Label cmdPos 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   250
         Index           =   0
         Left            =   120
         MousePointer    =   10  'Up Arrow
         TabIndex        =   4
         Top             =   240
         Width           =   125
      End
   End
   Begin VB.Frame Frame_cn 
      BackColor       =   &H00FFC0C0&
      Height          =   2415
      Left            =   0
      TabIndex        =   2
      Top             =   1170
      Width           =   8200
      Begin VB.Label cmdPos_cn 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Caption         =   "  "
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   250
         Index           =   0
         Left            =   120
         MousePointer    =   10  'Up Arrow
         TabIndex        =   5
         Top             =   240
         Width           =   250
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   465
      Left            =   6030
      Picture         =   "frmMain.frx":2000
      Stretch         =   -1  'True
      ToolTipText     =   "你的QQ的运行状态"
      Top             =   630
      Width           =   495
   End
   Begin VB.Label lblStat_3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ling     Antee"
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   6210
      TabIndex        =   10
      Top             =   3600
      Width           =   1965
   End
   Begin VB.Label lblStat_2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "操作提示栏"
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   1755
      TabIndex        =   9
      Top             =   3600
      Width           =   4440
   End
   Begin VB.Label lblStat_1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<鼠标移动>画图模式"
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   15
      TabIndex        =   8
      Top             =   3600
      Width           =   1725
   End
   Begin VB.Label lblQQNo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   6750
      TabIndex        =   7
      Top             =   900
      Width           =   1335
   End
   Begin VB.Label lblQQ 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   6750
      TabIndex        =   6
      Top             =   585
      Width           =   1335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuOpen 
         Caption         =   "打开...(&O)"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "保存(&S)"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "另存为...(&A)"
      End
      Begin VB.Menu mnuExample 
         Caption         =   "贴图例子"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuRedraw 
         Caption         =   "重新画过(&R)"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "复制到剪贴板(&C)"
      End
      Begin VB.Menu mnuDraw 
         Caption         =   "单击画图模式(&D)"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "工具(&T)"
      Begin VB.Menu mnuQQDel 
         Caption         =   "删除QQ登录记录"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuComt 
         Caption         =   "帮助主题"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "关于(&A)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sel_word As String
Dim ChsEng As String
Dim FileName As String
Dim boolDrawing As Boolean
Dim isDown As Boolean
Dim z As POINTAPI
Dim qqurl As String

Private Sub cmdQq_Click()
    qqurl = "http://home.meetchinese.com/qqol/qqol.php?s=10&qq=" & txtqq.Text
    Web.Navigate qqurl
    lblqqsta.Caption = "搜索中..."
End Sub

Private Sub CmdWord_Click()
    Dim fText As frmText
    Set fText = New frmText
    'Load fText
    fText.Show 1
    'fText.Refresh
    If MyWord <> "" Then
        Word.Text = MyWord
    End If
End Sub



Private Sub cmdPos_cn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If boolDrawing Then
        If Button = 1 Then
            isDown = True
            lblStat_2.Caption = "单击<右键><暂停>鼠标移动画图"
        ElseIf Button = 2 Then
            isDown = False
            lblStat_2.Caption = "单击<左键><开始>鼠标移动画图"
        End If
    Else
        If cmdPos_cn(Index).Caption = "  " Then
            cmdPos_cn(Index).Caption = Word.Text
        Else: cmdPos_cn(Index).Caption = "  "
        End If
    End If
End Sub




Private Sub cmdPos_cn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If boolDrawing Then
        If isDown Then
            cmdPos_cn(Index).Caption = Word.Text
        End If
    End If
End Sub

Private Sub cmdPos_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If boolDrawing Then
        If Button = 1 Then
            isDown = True
            lblStat_2.Caption = "单击<右键><暂停>鼠标移动画图"
        ElseIf Button = 2 Then
            isDown = False
            lblStat_2.Caption = "单击<左键><开始>鼠标移动画图"
        End If
    Else
        If cmdPos(Index).Caption = " " Then
            cmdPos(Index).Caption = Word.Text
        Else: cmdPos(Index).Caption = " "
        End If
    End If
End Sub
Private Sub cmdPos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If boolDrawing Then
        If isDown Then
            cmdPos(Index).Caption = Word.Text
        End If
    End If
End Sub






Private Sub CmdToolbar_Click(Index As Integer)
    Select Case Index
        Case 0
            mnuOpen_Click
        Case 1
            mnuSave_Click
        Case 2
            mnuRedraw_Click
        Case 3
            mnuCopy_Click
        Case 4
            mnuDraw_Click
        Case 5
            mnuComt_Click
    End Select
End Sub

Private Sub lblStat_3_Click()
    MsgBox ("Hello")
End Sub

Private Sub mnuAbout_Click()
    Dim fAbout As frmAbout
    Set fAbout = New frmAbout
    Load fAbout
    fAbout.Show 1
    fAbout.Refresh
    Unload fAbout
End Sub

Private Sub mnuComt_Click()
    Shell "hh " & App.Path & "\qpdhelp.chm", vbNormalFocus
End Sub

Private Sub mnuCopy_Click()
    'mnuCopy_Click
    Dim TempWord As String
    Dim i As Integer
    If ChsEng = "eng" Then
        word_join
    End If
        
    TempWord = ""

    For i = 0 To 242
        TempWord = TempWord & cmdPos_cn(i).Caption
        If (i Mod 32) = 31 Then
            TempWord = TempWord + Chr(13) + Chr(10)
        End If
    Next i
    Clipboard.Clear
    Clipboard.SetText TempWord
End Sub

Private Sub mnuDraw_Click()
    boolDrawing = Not boolDrawing
    mnuDraw.Checked = Not mnuDraw.Checked
    If boolDrawing Then
        lblStat_1.Caption = "<鼠标移动>画图模式"
        lblStat_2.Caption = "单击<左键><开始>鼠标移动画图"
    Else
        lblStat_1.Caption = "<左键单击>画图模式"
        lblStat_2.Caption = "单击<左键>画一个字符"

    End If
End Sub

Private Sub mnuExample_Click()
    Dim TempFile As Long
    Dim TempWord As String
    Dim i As Integer
    FileName = App.Path & "\Example.qpd"
    TempFile = FreeFile

    word_join
    Open FileName For Input As #TempFile
    Frame_en.Visible = False
    Frame_cn.Visible = True
    For j = 0 To 6
        k = 1
        Line Input #TempFile, databuff$
        For i = 0 To 31
            TempWord = Mid$(databuff$, k, 1)
            If TempWord <> "" Then
                If Asc(TempWord) < 0 Then
                    cmdPos_cn(32 * j + i).Caption = Mid$(databuff$, k, 1)
                    k = k + 1
                Else
                    cmdPos_cn(32 * j + i).Caption = Mid$(databuff$, k, 2)
                    k = k + 2
                End If
            End If
        Next i
    Next j
    Close TempFile
    Me.Caption = "QQ贴图DIY  v" & App.Major & "." & App.Minor & "." & App.Revision & "-----" & FileName
    If Asc(Word.Text) > 0 Then
        Frame_en.Visible = True
        Frame_cn.Visible = False
    End If
    word_depart
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim fso
    Dim strCmdLine As String
    Dim iCmdLine As Integer

    'test
    For i = 1 To 485 Step 1
        Load cmdPos(i)
        cmdPos(i).Visible = True
        cmdPos(i).Caption = " "
        cmdPos(i).Top = cmdPos(0).Height * (i \ 64) + cmdPos(0).Top
        cmdPos(i).Left = cmdPos(0).Left + (i Mod 64) * cmdPos(0).Width
        If (i Mod 2) <> ((i \ 64) Mod 2) Then
            cmdPos(i).BackColor = &HFFFFFF
        Else
            cmdPos(i).BackColor = &HC0C0FF
        End If
    Next i
    
    For i = 1 To 242 Step 1
        Load cmdPos_cn(i)
        cmdPos_cn(i).Visible = True
        cmdPos_cn(i).Caption = "  "
        cmdPos_cn(i).Top = cmdPos_cn(0).Height * (i \ 32) + cmdPos_cn(0).Top
        cmdPos_cn(i).Left = cmdPos_cn(0).Left + (i Mod 32) * cmdPos_cn(0).Width
        If (i Mod 2) <> ((i \ 32) Mod 2) Then
            cmdPos_cn(i).BackColor = &HFFFFFF
        Else
            cmdPos_cn(i).BackColor = &HC0C0FF
        End If
    Next i
    
    sel_word = "*"
    Word.Text = "*"
    ChsEng = "eng"
    Me.Caption = "QQ贴图DIY  v" & App.Major & "." & App.Minor & "." & App.Revision
    FileName = ""
    
    Dim qqstat As String
    If isQQrun() = 0 Then
        qqstat = "未运行"
    Else
        qqstat = "已运行"
    End If
    lblQQ.Caption = "QQ状态:" & qqstat
    lblQQNo.Caption = "No:" & Str(isQQrun())
    boolDrawing = True
    isDown = False
    strCmdLine = Command
    'Set fso = CreateObject("Scripting.FileSystemObject")
    'If fso.FileExists(strCmdLine) Then
    '    FileName = strCmdLine
    '    FileOpen
    'End If

End Sub


Private Sub mnuOpen_Click()
    Dim TempFile As Long
    Dim TempWord As String
    Dim i As Integer
    'Dim databuff() As Byte
    
        cdMain.ShowSave
        If Len(cdMain.FileName) < 1 Then
            Exit Sub
        End If
        FileName = cdMain.FileName
    TempFile = FreeFile

    Open FileName For Input As #TempFile
    Frame_en.Visible = False
    Frame_cn.Visible = True
    For j = 0 To 6
        k = 1

        Line Input #TempFile, databuff$
        For i = 0 To 31
            TempWord = Mid$(databuff$, k, 1)
            If Asc(TempWord) < 0 Then
                cmdPos_cn(32 * j + i).Caption = Mid$(databuff$, k, 1)
                k = k + 1
            Else
                cmdPos_cn(32 * j + i).Caption = Mid$(databuff$, k, 2)
                k = k + 2
            End If
        Next i
        
    Next j
    
    Close TempFile
    Me.Caption = "QQ贴图DIY  v" & App.Major & "." & App.Minor & "." & App.Revision & "-----" & FileName
End Sub

Private Sub mnuQQDel_Click()
    DelQQRec
End Sub

Private Sub mnuRedraw_Click()
    'mnuRedraw_Click
    Dim i As Integer
    For i = 0 To 485 Step 1
        cmdPos(i).Caption = " "
    Next i
    For i = 0 To 242 Step 1
        cmdPos_cn(i).Caption = "  "
    Next i
End Sub

Private Sub mnuSave_Click()
    'mnuSave_Click
    Dim TempFile As Long
    Dim TempWord As String
    Dim i As Integer
    If FileName = "" Then
        cdMain.ShowSave
        If Len(cdMain.FileName) < 1 Then
            Exit Sub
        End If
        FileName = cdMain.FileName
    End If
    TempFile = FreeFile
    If ChsEng = "eng" Then
        word_join
    End If

    Open FileName For Output As #TempFile
        
    TempWord = ""
    
    For i = 0 To 242
        TempWord = TempWord & cmdPos_cn(i).Caption
        If (i Mod 32) = 31 Then
            TempWord = TempWord + Chr(13) + Chr(10)
        End If
    Next i
    
    Print #TempFile, TempWord
    Close TempFile
    Me.Caption = "QQ贴图DIY  v" & App.Major & "." & App.Minor & "." & App.Revision & "-----" & FileName
End Sub

Private Sub mnuSaveAs_Click()
    'mnuSaveAs_Click
    Dim TempFile As Long
    Dim TempWord As String
    Dim i As Integer
    cdMain.DialogTitle = "另存为..."
    cdMain.ShowSave
    If Len(cdMain.FileName) < 1 Then
        Exit Sub
    End If
    FileName = cdMain.FileName
    TempFile = FreeFile
    If ChsEng = "eng" Then
        word_join
    End If

    Open FileName For Output As #TempFile
        
    TempWord = ""

    For i = 0 To 242
        TempWord = TempWord & cmdPos_cn(i).Caption
        If (i Mod 32) = 31 Then
            TempWord = TempWord + Chr(13) + Chr(10)
        End If
    Next i
    
    Print #TempFile, TempWord
    Close TempFile
    Me.Caption = "QQ贴图DIY  v" & App.Major & "." & App.Minor & "." & App.Revision & "-----" & FileName

End Sub

Private Sub timQQrun_Timer()
    Dim qqstat As String
    If isQQrun() = 0 Then
        qqstat = "未运行"
    Else
        qqstat = "已运行"
    End If
    lblQQ.Caption = "QQ状态:" & qqstat
    lblQQNo.Caption = "No:" & Str(isQQrun())
End Sub






Private Sub txtqq_GotFocus()
    txtqq.Text = ""
End Sub

Private Sub Web_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
If URL = "http://img.tencent.com/face/l/0-1.gif" Then
    lblqqsta.Caption = "离线或隐身"
ElseIf URL = "http://img.tencent.com/face/l/0-0.gif" Then
    lblqqsta.Caption = "在线"
Else
    lblqqsta.Caption = "未连接或错误"
End If
End Sub



Private Sub Word_Change()
        Do While Len(Word.Text) <> 1
            MsgBox "只能输入一个中/英文字符！", vbOKOnly, "请注意"
            Word.SetFocus
            GoTo error
        Loop
        
        If Asc(Word.Text) > 0 And Asc(sel_word) < 0 Then
            ChsEng = "eng"
            Frame_en.Visible = True
            Frame_cn.Visible = False
            word_depart
        End If
        If Asc(Word.Text) < 0 And Asc(sel_word) > 0 Then
            ChsEng = "chs"
            Frame_cn.Visible = True
            Frame_en.Visible = False
            word_join
        End If
        sel_word = Word.Text
error:
End Sub

Private Sub word_join()
        Dim i As Integer
        For i = 0 To 242
            If Asc(cmdPos(2 * (i + 1) - 2).Caption) < 0 Or Asc(cmdPos(2 * (i + 1) - 1).Caption) < 0 Then
                cmdPos_cn(i).Caption = cmdPos(2 * (i + 1) - 2).Caption
            Else: cmdPos_cn(i).Caption = cmdPos(2 * (i + 1) - 2).Caption & cmdPos(2 * (i + 1) - 1).Caption
            End If
        Next i
    
End Sub
Private Sub word_depart()
        Dim i As Integer
        For i = 0 To 484 Step 2
            cmdPos(i).Caption = Left(cmdPos_cn(i \ 2).Caption, 1)
        Next i
        For i = 1 To 485 Step 2
            cmdPos(i).Caption = Right(cmdPos_cn(i \ 2).Caption, 1)
        Next i
End Sub
Private Sub FileOpen()
    Dim TempFile As Long
    Dim TempWord As String
    word_join
    TempFile = FreeFile
    Open FileName For Input As #TempFile
    Frame_en.Visible = False
    Frame_cn.Visible = True
    For j = 0 To 6
        k = 1
        Line Input #TempFile, databuff$
        For i = 0 To 31
            TempWord = Mid$(databuff$, k, 1)
            If TempWord <> "" Then
                If Asc(TempWord) < 0 Then
                    cmdPos_cn(32 * j + i).Caption = Mid$(databuff$, k, 1)
                    k = k + 1
                Else
                    cmdPos_cn(32 * j + i).Caption = Mid$(databuff$, k, 2)
                    k = k + 2
                End If
            End If
        Next i
    Next j
    Close TempFile
    Me.Caption = "QQ贴图DIY  v" & App.Major & "." & App.Minor & "." & App.Revision & "-----" & FileName
    If Asc(Word.Text) > 0 Then
        Frame_en.Visible = True
        Frame_cn.Visible = False
    End If
    word_depart
End Sub

