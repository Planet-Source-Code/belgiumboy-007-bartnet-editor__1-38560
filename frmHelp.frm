VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "Help"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&CLose"
      Height          =   375
      Left            =   3930
      TabIndex        =   2
      Top             =   2670
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Display"
      Default         =   -1  'True
      Height          =   375
      Left            =   90
      TabIndex        =   1
      Top             =   2670
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   2520
      Left            =   -15
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   30
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   2535
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private HelpIndex As New FileSystemObject
Private HelpFiles As TextStream

Private Sub Command1_Click()
    Set HelpFiles = HelpIndex.OpenTextFile(App.Path & "\Help\" & Combo1.Text & ".BartNet", ForReading)
    With HelpFiles
        Do Until .AtEndOfStream
            Label1.Caption = Label1.Caption & .ReadLine & vbCrLf
        Loop
    End With
End Sub

Private Sub Command2_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub Form_Load()
    Set HelpFiles = HelpIndex.OpenTextFile(App.Path & "\Help\Index.BartNet", ForReading)
    With HelpFiles
        Do Until .AtEndOfStream
            Combo1.AddItem .ReadLine
        Loop
    End With

    Label1.Caption = ""
End Sub
