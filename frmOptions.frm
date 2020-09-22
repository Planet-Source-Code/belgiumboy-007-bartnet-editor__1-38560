VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3727
      TabIndex        =   2
      Top             =   3487
      Width           =   1575
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   233
      TabIndex        =   1
      Top             =   3487
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   75
      TabIndex        =   0
      Top             =   82
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   5741
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmOptions.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Default Document Name"
      TabPicture(1)   =   "frmOptions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Default Font"
      TabPicture(2)   =   "frmOptions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Other"
      TabPicture(3)   =   "frmOptions.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame1"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Indent Size"
         Height          =   855
         Left            =   120
         TabIndex        =   32
         Top             =   2280
         Width           =   5175
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   285
            Left            =   1336
            TabIndex        =   34
            Top             =   360
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            Value           =   10
            AutoBuddy       =   -1  'True
            BuddyControl    =   "Text2"
            BuddyDispid     =   196612
            OrigLeft        =   1560
            OrigTop         =   360
            OrigRight       =   1800
            OrigBottom      =   645
            Increment       =   10
            Max             =   5000
            Min             =   10
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Default Alignment"
         Height          =   855
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   5175
         Begin VB.OptionButton optLeft 
            Caption         =   "Left"
            Height          =   195
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   615
         End
         Begin VB.OptionButton optCenter 
            Caption         =   "Center"
            Height          =   195
            Left            =   2040
            TabIndex        =   30
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optRight 
            Caption         =   "Right"
            Height          =   195
            Left            =   3720
            TabIndex        =   29
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   24
         Top             =   360
         Width           =   5175
         Begin VB.CheckBox Check1 
            Caption         =   "Show StartUp Screen At StartUp"
            Height          =   255
            Left            =   360
            TabIndex        =   27
            Top             =   600
            Width           =   2655
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Show ToolBar (Default Setting)"
            Height          =   255
            Left            =   360
            TabIndex        =   26
            Top             =   1200
            Width           =   3015
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Show StatusBar (Default Setting)"
            Height          =   255
            Left            =   360
            TabIndex        =   25
            Top             =   1800
            Width           =   3015
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   19
         Top             =   360
         Width           =   5175
         Begin VB.OptionButton Option1 
            Caption         =   "NewPage"
            Height          =   195
            Left            =   360
            TabIndex        =   23
            Top             =   600
            Width           =   2295
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Untitled"
            Height          =   195
            Left            =   360
            TabIndex        =   22
            Top             =   1200
            Width           =   2775
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Other"
            Height          =   195
            Left            =   360
            TabIndex        =   21
            Top             =   1800
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   360
            TabIndex        =   20
            Top             =   2160
            Width           =   3015
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   3
         Top             =   360
         Width           =   5175
         Begin VB.TextBox txtBackgroundColor 
            Height          =   285
            Left            =   2520
            TabIndex        =   14
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txtFontColor 
            Height          =   285
            Left            =   2520
            TabIndex        =   13
            Top             =   600
            Width           =   2535
         End
         Begin VB.CheckBox ckBold 
            Caption         =   "Bold"
            Height          =   255
            Left            =   2520
            TabIndex        =   12
            Top             =   1680
            Width           =   615
         End
         Begin VB.CheckBox ckItalic 
            Caption         =   "Italic"
            Height          =   255
            Left            =   2520
            TabIndex        =   11
            Top             =   1920
            Width           =   735
         End
         Begin VB.CheckBox ckUnderline 
            Caption         =   "Underline"
            Height          =   255
            Left            =   2520
            TabIndex        =   10
            Top             =   2160
            Width           =   1095
         End
         Begin VB.CheckBox ckStrikeThru 
            Caption         =   "StrikeThru"
            Height          =   255
            Left            =   2520
            TabIndex        =   9
            Top             =   2400
            Width           =   1095
         End
         Begin VB.TextBox txtFontSize 
            Height          =   285
            Left            =   2520
            TabIndex        =   8
            Top             =   960
            Width           =   2535
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Change"
            Height          =   615
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtFont 
            Height          =   285
            Left            =   2520
            TabIndex        =   6
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Change"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Change"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   735
         End
         Begin MSComDlg.CommonDialog cdChange 
            Left            =   840
            Top             =   1920
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Font:"
            Height          =   255
            Left            =   2040
            TabIndex        =   18
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Font Color:"
            Height          =   255
            Left            =   1560
            TabIndex        =   17
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Background Color:"
            Height          =   255
            Left            =   1080
            TabIndex        =   16
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Font Size:"
            Height          =   255
            Left            =   1680
            TabIndex        =   15
            Top             =   960
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    cmdApply.Enabled = True
End Sub

Private Sub Check2_Click()
    cmdApply.Enabled = True
End Sub

Private Sub Check3_Click()
    cmdApply.Enabled = True
End Sub

Private Sub ckBold_Click()
    cmdApply.Enabled = True
End Sub

Private Sub ckItalic_Click()
    cmdApply.Enabled = True
End Sub

Private Sub ckStrikeThru_Click()
    cmdApply.Enabled = True
End Sub

Private Sub ckUnderline_Click()
    cmdApply.Enabled = True
End Sub

Private Sub cmdApply_Click()
Dim strNaam As String
Dim a As TextStream
Dim fsf As New FileSystemObject

    With fsf
        strNaam = .BuildPath(App.Path, "Info.BartNet")
        Set a = .OpenTextFile(strNaam, ForWriting)
    End With
    If Check1.Value = vbChecked Then a.WriteLine "True" Else a.WriteLine "False"
    If Check2.Value = vbChecked Then a.WriteLine "True" Else a.WriteLine "False"
    If Check3.Value = vbChecked Then a.WriteLine "True" Else a.WriteLine "False"
    If Option1.Value = True Then a.WriteLine "NewPage" Else If Option2.Value = True Then a.WriteLine "Untitled" Else a.WriteLine Text1.Text
    a.WriteLine txtFont.Text
    a.WriteLine txtFontSize.Text
    a.WriteLine txtFontColor.Text
    a.WriteLine txtBackgroundColor.Text
    If ckBold.Value = vbChecked Then a.WriteLine "True" Else a.WriteLine "False"
    If ckItalic.Value = vbChecked Then a.WriteLine "True" Else a.WriteLine "False"
    If ckUnderline.Value = vbChecked Then a.WriteLine "True" Else a.WriteLine "False"
    If ckStrikeThru.Value = vbChecked Then a.WriteLine "True" Else a.WriteLine "False"
    If optLeft.Value = True Then a.WriteLine "Left" Else If optRight.Value = True Then a.WriteLine "Right" Else a.WriteLine "Center"
    a.WriteLine Text2.Text
    
    IndentSize = Text2.Text
    cmdApply.Enabled = False
End Sub

Private Sub cmdClose_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub Command3_Click()
    cdChange.Flags = cdlCFBoth Or cdlCFTTOnly And cdlCFEffects
    cdChange.ShowFont
    txtFont.Text = cdChange.FontName
    txtFontSize.Text = cdChange.FontSize
    txtFontColor.Text = cdChange.Color
    
    If cdChange.FontBold = True Then ckBold.Value = vbChecked Else ckBold.Value = vbUnchecked
    If cdChange.FontItalic = True Then ckItalic.Value = vbChecked Else ckItalic.Value = vbUnchecked
    If cdChange.FontUnderline = True Then ckUnderline.Value = vbChecked Else ckUnderline.Value = vbUnchecked
    If cdChange.FontStrikethru = True Then ckStrikeThru.Value = vbChecked Else ckStrikeThru.Value = vbUnchecked
End Sub

Private Sub Command4_Click()
    cdChange.ShowColor
    txtBackgroundColor.Text = cdChange.Color
End Sub

Private Sub Command5_Click()
    cdChange.ShowColor
    txtFontColor.Text = cdChange.Color
End Sub

Private Sub Form_Load()
    Select Case DefaultAlignment
        Case "Left"
            optLeft = True
            optCenter = False
            optRight = False
        Case "Center"
            optLeft = False
            optCenter = True
            optRight = False
        Case "Right"
            optLeft = False
            optCenter = False
            optRight = True
    End Select
    Label1.Top = Label1.Top + 35
    Label2.Top = Label2.Top + 35
    Label3.Top = Label3.Top + 35
    Label4.Top = Label4.Top + 35
        
    SSTab1.Tab = 0
    If StartUpScreenShow = True Then Check1.Value = vbChecked Else Check1.Value = vbUnchecked
    If ShowToolBar = True Then Check2.Value = vbChecked Else Check2.Value = vbUnchecked
    If ShowStatusBar = True Then Check3.Value = vbChecked Else Check3.Value = vbUnchecked
    
    If DocumentName = "NewPage" Then
        Option1.Value = True
        Option2.Value = False
        Option3.Value = False
        Text1.Text = ""
        Text1.Enabled = False
        Text1.Locked = True
        Text1.BackColor = &H80000013
    Else
        If DocumentName = "Untitled" Then
            Option1.Value = False
            Option2.Value = True
            Option3.Value = False
            Text1.Text = ""
            Text1.Enabled = False
            Text1.Locked = True
            Text1.BackColor = &H80000013
        Else
            Option1.Value = False
            Option2.Value = False
            Option3.Value = True
            Text1.Enabled = True
            Text1.Locked = False
            Text1.BackColor = &H80000005
            Text1.Text = DocumentName
        End If
    End If
    txtFont = DefaultFont
    txtFontSize = DefaultFontSize
    txtFontColor = DefaultTextColor
    txtBackgroundColor = DefaultBackgroundColor
    If DefaultBold = True Then ckBold.Value = vbChecked Else ckBold.Value = vbUnchecked
    If DefaultItalic = True Then ckItalic.Value = vbChecked Else ckItalic.Value = vbUnchecked
    If DefaultUnderline = True Then ckUnderline.Value = vbChecked Else ckUnderline.Value = vbUnchecked
    If DefaultStrikeThru = True Then ckStrikeThru.Value = vbChecked Else ckStrikeThru.Value = vbUnchecked
    Text2.Text = IndentSize
    
    cmdApply.Enabled = False
End Sub





Private Sub optCenter_Click()
    cmdApply.Enabled = True
End Sub

Private Sub Option1_Click()
    cmdApply.Enabled = True
    Text1.Text = ""
    Text1.Enabled = False
    Text1.Locked = True
    Text1.BackColor = &H80000013
End Sub

Private Sub Option2_Click()
    cmdApply.Enabled = True
    Text1.Text = ""
    Text1.Enabled = False
    Text1.Locked = True
    Text1.BackColor = &H80000013
End Sub

Private Sub Option3_Click()
    cmdApply.Enabled = True
    Text1.Locked = False
    Text1.Enabled = True
    Text1.BackColor = &H80000005
    Text1.Text = ""
End Sub

Private Sub optLeft_Click()
    cmdApply.Enabled = True
End Sub

Private Sub optRight_Click()
    cmdApply.Enabled = True
End Sub

Private Sub Text2_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtBackgroundColor_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtFont_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtFontColor_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtFontSize_Change()
    cmdApply.Enabled = True
End Sub
