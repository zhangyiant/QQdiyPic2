VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3480
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5520
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   3315
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5280
      Begin VB.Image imgLogo 
         Height          =   1185
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "版权所有"
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         BackStyle       =   0  'Transparent
         Caption         =   "公司"
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "警告:This is a free software."
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   3000
         Width           =   6165
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "版本"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3480
         TabIndex        =   5
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "产品"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   1920
         TabIndex        =   6
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "授权"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    lblCompany = "公司:" & App.CompanyName
    Me.lblCopyright = "版权所有:" & App.LegalCopyright
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub
