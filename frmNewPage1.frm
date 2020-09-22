VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmNewPage1 
   Caption         =   "frmNewPage1"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12120
   Icon            =   "frmNewPage1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   12120
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   9495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   16748
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmNewPage1.frx":0442
   End
End
Attribute VB_Name = "frmNewPage1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Saved As Boolean
Public SavedBefore As Boolean
Public SavedBeforePath As String
Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(1000) As String
Private Sub Form_GotFocus()
    If Me.WindowState = vbMaximized Then
        frmBartNetEditor.Caption = "BartNet Editor"
    Else
        frmBartNetEditor.Caption = "BartNet Editor"
    End If
End Sub
Public Sub Undo()
    'This says that if the Index is = to 0, then It shouldn't undo anymore
    If gintIndex = 0 Then Exit Sub
    
    'This is the basic undo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
    rtb1.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub
Public Sub Redo()
    'This is the basic redo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex + 1
    On Error Resume Next
    rtb1.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub
Private Sub Form_Load()
    Me.WindowState = vbMaximized
        If frmBartNetEditor.tlbStandard.Visible = True Then
            If frmBartNetEditor.StatusBar1.Visible = True Then
                rtb1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
            Else
                rtb1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
            End If
        Else
            rtb1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        End If
    With rtb1
        .SelFontSize = DefaultFontSize
        .SelFontName = DefaultFont
        .SelColor = DefaultTextColor
        .BackColor = DefaultBackgroundColor
        .SelBold = DefaultBold
        .SelItalic = DefaultItalic
        .SelUnderline = DefaultUnderline
        .SelStrikeThru = DefaultStrikeThru
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim a%
    If Me.Caption = "frmNewPage1" Then End
    If Saved = True Then
    Beep
        a = MsgBox("Are You Sure You Want To Close '" & Me.Caption & " ' ?", 4132, "BartNet Editor")
        Select Case a
            Case 6
                frmBartNetEditor.FormsCount = frmBartNetEditor.FormsCount - 1
            Case 7
                frmBartNetEditor.Cancelled = True
                Cancel = 1
        End Select
    Else
        If SavedBefore = True Then
        Beep
            a = MsgBox("Do You Want To Save The Changes You Made To '" & Me.Caption & " ' ?", 4131, "BartNet Editor")
            Select Case a
                Case 6
                    rtb1.SaveFile (SavedBeforePath)
                    Saved = True
                    frmBartNetEditor.FormsCount = frmBartNetEditor.FormsCount - 1
                    End
                Case 7
                    frmBartNetEditor.FormsCount = frmBartNetEditor.FormsCount - 1
                Case 2
                    frmBartNetEditor.Cancelled = True
                    Cancel = 1
            End Select
        Else
            Beep
            a = MsgBox("Do You Want To Save '" & Me.Caption & " ' ?", 4131, "BartNet Editor")
            Select Case a
                Case 6
                    On Error GoTo errorHandling
                    With CommonDialog1
                        .Filter = "Rich Text Format|*.rtf"
                        .ShowSave
                    End With
                    rtb1.SaveFile (CommonDialog1.FileName)
                    frmBartNetEditor.FormsCount = frmBartNetEditor.FormsCount - 1
                    End
                Case 7
                    frmBartNetEditor.FormsCount = frmBartNetEditor.FormsCount - 1
                Case 2
                    frmBartNetEditor.Cancelled = True
                    Cancel = 1
            End Select
        End If
    End If
    Exit Sub

errorHandling:
    If Err.Number = cdlCancel Then
        Me.Saved = False
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMaximized Then
        frmBartNetEditor.Caption = "BartNet Editor"
    Else
        frmBartNetEditor.Caption = "BartNet Editor"
    End If
    If Me.WindowState <> vbMinimized Then
        If frmBartNetEditor.tlbStandard.Visible = True Then
            If frmBartNetEditor.StatusBar1.Visible = True Then
                rtb1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
            Else
                rtb1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
            End If
        Else
            rtb1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        End If
    End If
End Sub

Private Sub rtb1_Change()
    If Not gblnIgnoreChange Then
        gintIndex = gintIndex + 1
        gstrStack(gintIndex) = rtb1.TextRTF
    End If
    Saved = False
    frmBartNetEditor.tlbStandard.Buttons("Save").Enabled = True
    frmBartNetEditor.mnuSave.Enabled = True
End Sub

Private Sub rtb1_Click()
    frmBartNetEditor.Check
End Sub

Private Sub rtb1_KeyUp(KeyCode As Integer, Shift As Integer)
    If rtb1.SelText <> "" Then
        frmBartNetEditor.mnuCopy.Enabled = True
        frmBartNetEditor.mnuCut.Enabled = True
        frmBartNetEditor.tlbStandard.Buttons("Copy").Enabled = True
        frmBartNetEditor.tlbStandard.Buttons("Cut").Enabled = True
    Else
        frmBartNetEditor.mnuCopy.Enabled = False
        frmBartNetEditor.mnuCut.Enabled = False
        frmBartNetEditor.tlbStandard.Buttons("Copy").Enabled = False
        frmBartNetEditor.tlbStandard.Buttons("Cut").Enabled = False
    End If
    
    frmBartNetEditor.Check
End Sub

Private Sub rtb1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If rtb1.SelText <> "" Then
        frmBartNetEditor.mnuCopy.Enabled = True
        frmBartNetEditor.mnuCut.Enabled = True
        frmBartNetEditor.tlbStandard.Buttons("Copy").Enabled = True
        frmBartNetEditor.tlbStandard.Buttons("Cut").Enabled = True
    Else
        frmBartNetEditor.mnuCopy.Enabled = False
        frmBartNetEditor.mnuCut.Enabled = False
        frmBartNetEditor.tlbStandard.Buttons("Copy").Enabled = False
        frmBartNetEditor.tlbStandard.Buttons("Cut").Enabled = False
    End If
    
    frmBartNetEditor.Check
End Sub
