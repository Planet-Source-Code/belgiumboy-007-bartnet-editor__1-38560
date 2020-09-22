VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStartUp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Document"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Show this screen at startup."
      Height          =   255
      Left            =   124
      TabIndex        =   7
      Top             =   3951
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog cdPath 
      Left            =   4092
      Top             =   1795
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ils1 
      Left            =   4092
      Top             =   1075
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   125
      ImageHeight     =   168
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartUp.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStartUp.frx":5852
            Key             =   "Open"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "What would you like to do?"
      Height          =   3720
      Left            =   132
      TabIndex        =   2
      Top             =   115
      Width           =   3615
      Begin VB.OptionButton optOpen 
         Height          =   1575
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1920
         Width           =   1335
      End
      Begin VB.OptionButton optNew 
         Height          =   1575
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label2 
         Caption         =   "Open existing document"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Create new document"
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   3852
      TabIndex        =   1
      Top             =   3467
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3852
      TabIndex        =   0
      Top             =   2987
      Width           =   1335
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strFilter As String
Private fso As New FileSystemObject
Private strNaam As String
Private a As TextStream

Public OpenedFromStartUpScreen As Boolean
Private Sub Command1_Click()
    If Check1.Value = vbUnchecked Then
        With fso
            strNaam = .BuildPath(App.Path, "Info.BartNet")
            Set a = .OpenTextFile(strNaam, ForWriting)
        End With
        a.WriteLine "False"
        a.WriteLine ShowToolBar
        a.WriteLine ShowStatusBar
        a.WriteLine DocumentName
        a.WriteLine DefaultFont
        a.WriteLine DefaultFontSize
        a.WriteLine DefaultTextColor
        a.WriteLine DefaultBackgroundColor
        a.WriteLine DefaultBold
        a.WriteLine DefaultItalic
        a.WriteLine DefaultUnderline
        a.WriteLine DefaultStrikeThru
        a.WriteLine DefaultAlignment
        a.WriteLine IndentSize
    End If
    If optNew.Value = True Then
        Me.Hide
        frmBartNetEditor.mnuNew_Click
    Else
        Me.Hide
        OpenedFromStartUpScreen = True
        frmBartNetEditor.mnuOpen_Click
    End If
    Exit Sub
End Sub

Private Sub Command2_Click()
    Unload Me
    End
End Sub

Private Sub Form_Load()
    frmNewPage1.Hide
    optNew.Picture = ils1.ListImages("New").Picture
    optOpen.Picture = ils1.ListImages("Open").Picture
    optNew.Value = True
    optOpen.Value = False
    strFilter = _
        "Text Documents|*.txt;*.bn" _
        & "|BartNet Editor Documents|*.bn" _
        & "|Notepad Documents|*.txt"
End Sub

