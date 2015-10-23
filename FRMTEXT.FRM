VERSION 5.00
Begin VB.Form frmText 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "特殊字符"
   ClientHeight    =   3390
   ClientLeft      =   3750
   ClientTop       =   570
   ClientWidth     =   2985
   ClipControls    =   0   'False
   Icon            =   "frmText.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2339.838
   ScaleMode       =   0  'User
   ScaleWidth      =   2803.069
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "请在此点击所需的字符"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.Label selc_word 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         Caption         =   "§"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   250
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Word = "§№☆★○●◎◇◆□■△▲※→←↑↓〓‰℃¤＾＿αβγδεζηθικλμνξοπρστυφχψωΓΘΛΞΠΣΥΦΨΩⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ≈≡≠≤≥＜＞≮≯∷±∫∮∝∞∧∨∑∏∪∩∈∵∴⊥∥∠⌒⊙≌∽√＄￡￥￠"
    For i = 1 To 105 Step 1
        Load selc_word(i)
        selc_word(i).Visible = True

        selc_word(i).Top = selc_word(0).Height * (i \ 10) + selc_word(0).Top
        selc_word(i).Left = selc_word(0).Left + (i Mod 10) * selc_word(0).Width

        If (i Mod 2) <> ((i \ 10) Mod 2) Then
            selc_word(i).BackColor = &HFFFFFF
        Else
            selc_word(i).BackColor = &HFFC0FF
        End If
        
        selc_word(i).Caption = Mid$(Word, i + 1, 1)
    Next i
End Sub

Private Sub selc_word_Click(Index As Integer)
    MyWord = selc_word(Index).Caption
    Unload Me
End Sub
