VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1095
   ClientLeft      =   2520
   ClientTop       =   6270
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Find"
      Height          =   375
      Left            =   4316
      TabIndex        =   4
      Top             =   112
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   4309
      TabIndex        =   3
      Top             =   607
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Next"
      Height          =   375
      Left            =   4309
      TabIndex        =   2
      Top             =   127
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1069
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   127
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Find What:"
      Height          =   255
      Left            =   109
      TabIndex        =   0
      Top             =   127
      Width           =   855
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Path As New FileSystemObject
Private Dir As TextStream
Private strName As String

Private Sub cmdCancel_Click()
    Me.Hide
    Combo1.Text = ""
    cmdFind.Visible = False
    Command1.Visible = True
End Sub

Private Sub cmdFind_Click()
Dim Found As Integer

    Found = frmBartNetEditor.ActiveForm.rtb1.Find(Combo1.Text, frmBartNetEditor.ActiveForm.rtb1.SelStart + Len(Combo1.Text))
    If Found = True Then
        MsgBox "Search text not found"
        cmdFind.Visible = False
        Command1.Visible = True
    Else
        frmBartNetEditor.ActiveForm.rtb1.SetFocus
    End If
End Sub

Private Sub Command1_Click()
    Combo1.AddItem (Combo1.Text)
    Command1.Visible = False
    cmdFind.Visible = True
    cmdFind_Click
End Sub

Private Sub Form_Load()
    Call FormOnTop(Me.hWnd, True)
    Label1.Top = Label1.Top + 35
    Combo1.Text = ""
    cmdFind.Visible = False
End Sub
