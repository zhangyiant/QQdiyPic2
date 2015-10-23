VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于我的应用程序"
   ClientHeight    =   3330
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2298.425
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraContactUs 
      BackColor       =   &H0080C0FF&
      Caption         =   "联系我们:"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   5535
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "zhangyiant@citiz.net"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         MouseIcon       =   "frmAbout.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "forpromise@163.com"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         MouseIcon       =   "frmAbout.frx":074C
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "张 毅"
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "凌 华"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "33422404"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "24412058"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail"
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "作 者"
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "QQ"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0A56
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   345
      Left            =   4200
      MaskColor       =   &H0080C0FF&
      TabIndex        =   0
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5337.57
      Y1              =   1905.001
      Y2              =   1905.001
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "应用程序描述"
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   1725
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   1605
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   5323.484
      Y1              =   1905.001
      Y2              =   1905.001
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "版本"
      Height          =   225
      Left            =   3000
      TabIndex        =   4
      Top             =   240
      Width           =   2205
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "关于 " & App.Title
    lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDescription.Caption = App.FileDescription
End Sub


Private Sub Label8_Click()
    ShellExecute hWnd, "open", "mailto:forpromise@163.com", vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub Label9_Click()
    ShellExecute hWnd, "open", "mailto:zhangyiant@citiz.net", vbNullString, vbNullString, vbNormalFocus
End Sub
