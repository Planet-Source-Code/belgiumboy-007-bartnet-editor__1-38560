VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmBartNetEditor 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "BartNet Editor"
   ClientHeight    =   10770
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   14730
   Icon            =   "frmBartNetEditor.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ilsToolbar 
      Left            =   9000
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBartNetEditor.frx":0442
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBartNetEditor.frx":0554
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBartNetEditor.frx":0666
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBartNetEditor.frx":0778
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBartNetEditor.frx":088A
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBartNetEditor.frx":099C
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBartNetEditor.frx":0AAE
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBartNetEditor.frx":0BC0
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBartNetEditor.frx":0CD2
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBartNetEditor.frx":0DE4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBartNetEditor.frx":0EF6
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBartNetEditor.frx":1008
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBartNetEditor.frx":111A
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBartNetEditor.frx":122C
            Key             =   "StrikeThru"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBartNetEditor.frx":133E
            Key             =   "Color"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBartNetEditor.frx":1660
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBartNetEditor.frx":1772
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBartNetEditor.frx":1884
            Key             =   "Right"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBartNetEditor.frx":1996
            Key             =   "DecreaseIndent"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBartNetEditor.frx":1DA0
            Key             =   "IncreaseIndent"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbStandard 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14730
      _ExtentX        =   25982
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ilsToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   28
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Object.ToolTipText     =   "Undo"
            ImageKey        =   "Undo"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Redo"
            Object.ToolTipText     =   "Redo"
            ImageKey        =   "Redo"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageKey        =   "Help"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "StrikeThru"
            Object.ToolTipText     =   "StrikeThru"
            ImageKey        =   "StrikeThru"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Color"
            Object.ToolTipText     =   "Font Color"
            ImageKey        =   "Color"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Left"
            Object.ToolTipText     =   "Align Left"
            ImageKey        =   "Left"
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageKey        =   "Center"
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Right"
            Object.ToolTipText     =   "Align Right"
            ImageKey        =   "Right"
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DecreaseIndent"
            Object.ToolTipText     =   "Decrease Indent"
            ImageKey        =   "DecreaseIndent"
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "IncreaseIndent"
            Object.ToolTipText     =   "Increase Indent"
            ImageKey        =   "IncreaseIndent"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   10395
      Width           =   14730
      _ExtentX        =   25982
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17595
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1402
            MinWidth        =   1411
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1402
            MinWidth        =   1411
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "9/11/2002"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "15:39"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnu 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mn 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Shortcut        =   ^P
      End
      Begin VB.Menu b 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu c 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuStandardToolbar 
         Caption         =   "Toolbar"
      End
      Begin VB.Menu d 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatusBar 
         Caption         =   "StatusBar"
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "&Format"
      Begin VB.Menu mnuFont 
         Caption         =   "Font"
         Begin VB.Menu mnuBold 
            Caption         =   "Bold"
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuItalic 
            Caption         =   "Italic"
            Shortcut        =   ^I
         End
         Begin VB.Menu mnuUnderline 
            Caption         =   "Underline"
            Shortcut        =   ^U
         End
         Begin VB.Menu l 
            Caption         =   "-"
         End
         Begin VB.Menu mnuStrikeThru 
            Caption         =   "StrikeThru"
         End
         Begin VB.Menu k 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMore 
            Caption         =   "More ..."
         End
      End
      Begin VB.Menu m 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSize 
         Caption         =   "Size"
         Begin VB.Menu mnuSize8 
            Caption         =   "8"
         End
         Begin VB.Menu mnuSize9 
            Caption         =   "9"
         End
         Begin VB.Menu mnuSize10 
            Caption         =   "10"
         End
         Begin VB.Menu mnuSize11 
            Caption         =   "11"
         End
         Begin VB.Menu mnuSize12 
            Caption         =   "12"
         End
         Begin VB.Menu mnuSize14 
            Caption         =   "14"
         End
         Begin VB.Menu mnuSize16 
            Caption         =   "16"
         End
         Begin VB.Menu mnuSize18 
            Caption         =   "18"
         End
         Begin VB.Menu mnuSize20 
            Caption         =   "20"
         End
         Begin VB.Menu mnuSize22 
            Caption         =   "22"
         End
         Begin VB.Menu mnuSize24 
            Caption         =   "24"
         End
         Begin VB.Menu mnuSize26 
            Caption         =   "26"
         End
         Begin VB.Menu mnuSize28 
            Caption         =   "28"
         End
         Begin VB.Menu mnuSize30 
            Caption         =   "30"
         End
         Begin VB.Menu mnuSize36 
            Caption         =   "36"
         End
         Begin VB.Menu mnuSize40 
            Caption         =   "40"
         End
         Begin VB.Menu mnuSize48 
            Caption         =   "48"
         End
         Begin VB.Menu mnuSize72 
            Caption         =   "72"
         End
      End
      Begin VB.Menu hgfhfgh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFontColor 
         Caption         =   "Font Color"
      End
      Begin VB.Menu mnuBackgroundColor 
         Caption         =   "Background Color"
      End
      Begin VB.Menu ghdhgfh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAlignment 
         Caption         =   "ALignment"
         Begin VB.Menu mnuLeft 
            Caption         =   "Left"
         End
         Begin VB.Menu mnuCenter 
            Caption         =   "Center"
         End
         Begin VB.Menu mnuRight 
            Caption         =   "Right"
         End
      End
      Begin VB.Menu gfsd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIndent 
         Caption         =   "Indent"
         Begin VB.Menu mnuDecreaseIndent 
            Caption         =   "Decrease Indent"
         End
         Begin VB.Menu mnuIncreaseIndent 
            Caption         =   "Increase Indent"
         End
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascadeWindows 
         Caption         =   "Cascade Windows"
      End
      Begin VB.Menu mnuTileHorizontally 
         Caption         =   "Tile Windows Horizontally"
      End
      Begin VB.Menu mnuTileVertically 
         Caption         =   "Tile Windows Vertically"
      End
   End
   Begin VB.Menu mnuasf 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu sgfsdgdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBartNetOnline 
         Caption         =   "Visit BartNet Online"
      End
      Begin VB.Menu gfdgsdgmnu 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmBartNetEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PageNumber As Integer
Public frmNewPage As New frmNewPage1
Private fff As New FileSystemObject
Private strNaam As String
Private strm As TextStream
Public FormsCount As Integer
Public Cancelled As Boolean

Sub FontSelect()
    CommonDialog1.Flags = cdlCFBoth Or cdlCFTTOnly Or cldCFEffects
    CommonDialog1.ShowFont
    With Me.ActiveForm.rtb1
        .SelFontName = CommonDialog1.FontName
        .SelFontSize = CommonDialog1.FontSize
        .SelBold = CommonDialog1.FontBold
        .SelItalic = CommonDialog1.FontItalic
        .SelStrikeThru = CommonDialog1.FontStrikethru
        .SelUnderline = CommonDialog1.FontUnderline
        .SelColor = CommonDialog1.Color
    End With
End Sub

Private Sub MDIForm_Load()
    Saved = False
    Cancelled = False
    FormsCount = 0
    Unload frmNewPage1
    
    With tlbStandard
    .Buttons("Cut").Enabled = False
    .Buttons("Copy").Enabled = False
    .Buttons("Undo").Enabled = False
    .Buttons("Redo").Enabled = False
    .Buttons("Paste").Enabled = False
    .Buttons("Find").Enabled = False
    .Buttons("Save").Enabled = False
    End With
    mnuCut.Enabled = False
    mnuCopy.Enabled = False
    mnuUndo.Enabled = False
    mnuRedo.Enabled = False
    mnuPaste.Enabled = False
    mnuFind.Enabled = False
    mnuSave.Enabled = False
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim a%
    If FormsCount > 0 Then
        Do Until FormsCount = 0
            If Cancelled = True Then
                Cancelled = False
                Cancel = 1
                Exit Sub
            Else
                Unload Me.ActiveForm
            End If
        Loop
    Else
        a = MsgBox("Are You Sure You Want To Exit?", vbYesNo + vbInformation, "BartNet Editor")
        Select Case a
        Case 6
            Unload Me
            End
        Case 7
            Cancel = 1
        End Select
    End If
End Sub

Private Sub MDIForm_Resize()
    If Me.WindowState = vbMinimized Then frmFind.Hide
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuBackgroundColor_Click()
    CommonDialog1.ShowColor
    Me.ActiveForm.rtb1.BackColor = CommonDialog1.Color
End Sub

Private Sub mnuBartNetOnline_Click()
    With frmBrowser
        .Show
        .WindowState = vbMaximized
        .Caption = "http://www.BartNet.freeservers.com"
    End With
    Me.Caption = "BartNet Editor"
End Sub

Private Sub mnuBold_Click()
    If mnuBold.Checked = True Then
        Me.ActiveForm.rtb1.SelBold = False
        mnuBold.Checked = False
        tlbStandard.Buttons("Bold").Value = tbrUnpressed
    Else
        Me.ActiveForm.rtb1.SelBold = True
        mnuBold.Checked = True
        tlbStandard.Buttons("Bold").Value = tbrPressed
    End If
End Sub

Private Sub mnuCascadeWindows_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuCenter_Click()
    Me.ActiveForm.rtb1.SelAlignment = 2
    
    tlbStandard.Buttons("Left").Value = tbrUnpressed
    tlbStandard.Buttons("Center").Value = tbrPressed
    tlbStandard.Buttons("Right").Value = tbrUnpressed
    
    mnuLeft.Checked = False
    mnuCenter.Checked = True
    mnuRight.Checked = False
End Sub

Public Sub mnuClose_Click()
    Unload Me.ActiveForm
End Sub

Private Sub mnuCopy_Click()
    Clipboard.Clear
    Clipboard.SetText Me.ActiveForm.rtb1.SelText, 1
    
    tlbStandard.Buttons("Paste").Enabled = True
    mnuPaste.Enabled = True
End Sub

Private Sub mnuCut_Click()
    Clipboard.Clear
    Clipboard.SetText Me.ActiveForm.rtb1.SelText, 1
    Me.ActiveForm.rtb1.SelText = ""
    
    tlbStandard.Buttons("Paste").Enabled = True
    mnuPaste.Enabled = True
End Sub

Private Sub mnuDecreaseIndent_Click()
    Me.ActiveForm.rtb1.SelIndent = Me.ActiveForm.rtb1.SelIndent - IndentSize
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Sub SaveNow()
    If Me.ActiveForm.Saved = False Then
        Dim blnAnnuleren As Boolean
    
    
        On Error GoTo Foutafhandeling
        blnAnnuleren = False
        With CommonDialog1
            .Filter = "BartNet Editor Documents|*.bn"
            .ShowSave
        End With
        
        If Not blnAnnuleren Then
            Dim strm As TextStream
            Dim str As String
            With fff
                Set strm = .CreateTextFile(CommonDialog1.FileName, False)
            End With
            strm.Write Me.ActiveForm.rtb1.TextRTF
            Me.Caption = CommonDialog1.FileTitle
            Me.ActiveForm.Saved = True
        End If

        Exit Sub
Foutafhandeling:
    If Err.Number = cdlCancel Then
    Me.ActiveForm.Saved = False
    Resume Next
    Else
        MsgBox Err.Description
    End If
    End If
End Sub

Private Sub mnuFind_Click()
    frmFind.Show
End Sub

Private Sub mnuFontColor_Click()
    CommonDialog1.ShowColor
    Me.ActiveForm.rtb1.SelColor = CommonDialog1.Color
End Sub

Private Sub mnuHelp_Click()
    frmHelp.Show vbModal
End Sub

Private Sub mnuIncreaseIndent_Click()
    Me.ActiveForm.rtb1.SelIndent = Me.ActiveForm.rtb1.SelIndent + IndentSize
End Sub

Private Sub mnuItalic_Click()
    If mnuItalic.Checked = True Then
        Me.ActiveForm.rtb1.SelItalic = False
        mnuItalic.Checked = False
        tlbStandard.Buttons("Italic").Value = tbrUnpressed
    Else
        Me.ActiveForm.rtb1.SelItalic = True
        mnuItalic.Checked = True
        tlbStandard.Buttons("Italic").Value = tbrPressed
    End If
End Sub

Private Sub mnuLeft_Click()
    Me.ActiveForm.rtb1.SelAlignment = 0
    
    tlbStandard.Buttons("Left").Value = tbrPressed
    tlbStandard.Buttons("Center").Value = tbrUnpressed
    tlbStandard.Buttons("Right").Value = tbrUnpressed
    
    mnuLeft.Checked = True
    mnuCenter.Checked = False
    mnuRight.Checked = False
End Sub

Private Sub mnuMore_Click()
    FontSelect
End Sub

Public Sub mnuNew_Click()
    With tlbStandard
    .Buttons("Cut").Enabled = True
    .Buttons("Copy").Enabled = True
    .Buttons("Undo").Enabled = True
    .Buttons("Redo").Enabled = True
    .Buttons("Paste").Enabled = True
    .Buttons("Find").Enabled = True
    .Buttons("Save").Enabled = True
    End With
    mnuCut.Enabled = True
    mnuCopy.Enabled = True
    mnuUndo.Enabled = True
    mnuRedo.Enabled = True
    mnuPaste.Enabled = True
    mnuFind.Enabled = True
    mnuSave.Enabled = True
    
    PageNumber = PageNumber + 1
    Set frmNewPage = New frmNewPage1
    With frmNewPage
        .Caption = DocumentName & " " & PageNumber
        .Show
        .WindowState = vbMaximized
        .Saved = False
        .SavedBefore = False
    End With
    Me.Caption = "BartNet Editor"
    FormsCount = FormsCount + 1
End Sub

Public Sub mnuOpen_Click()
    With tlbStandard
    .Buttons("Cut").Enabled = True
    .Buttons("Copy").Enabled = True
    .Buttons("Undo").Enabled = True
    .Buttons("Redo").Enabled = True
    .Buttons("Paste").Enabled = True
    .Buttons("Find").Enabled = True
    .Buttons("Save").Enabled = False
    End With
    mnuCut.Enabled = True
    mnuCopy.Enabled = True
    mnuUndo.Enabled = True
    mnuRedo.Enabled = True
    mnuPaste.Enabled = True
    mnuFind.Enabled = True
    mnuSave.Enabled = False
    
    
    frmNewPage.rtb1.BackColor = 16777215
    Dim blnAnnuleren As Boolean
    
    On Error GoTo errorHandling

    With CommonDialog1
        .Flags = cdlOFNFileMustExist
        .Filter = "Text Documents|*.txt;*.rtf"
        .ShowOpen
    End With
        
    Me.Hide
    Set frmNewPage = New frmNewPage1
    With fff
        strNaam = CommonDialog1.FileName
    End With
    With frmNewPage
        .Show
        .rtb1.LoadFile (strNaam)
        .Saved = True
        .SavedBefore = True
        .SavedBeforePath = CommonDialog1.FileName
        .Caption = CommonDialog1.FileTitle
        .WindowState = vbMaximized
    End With
    Me.Caption = "BartNet Editor"
    FormsCount = FormsCount + 1
    Check
    Exit Sub

errorHandling:
    If Err.Number = cdlCancel Then
        If frmStartUp.OpenedFromStartUpScreen = True Then
            frmStartUp.Show vbModal
        Else
        
        End If
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show vbModal
End Sub

Private Sub mnuPaste_Click()
    Me.ActiveForm.rtb1.SelText = Clipboard.GetText(1)
    frmNewPage1.Saved = False
    mnuSave.Enabled = True
    tlbStandard.Buttons("Save").Enabled = True
End Sub

Private Sub mnuRedo_Click()
    Me.ActiveForm.Redo
    
    mnuUndo.Enabled = True
    tlbStandard.Buttons("Undo").Enabled = True
End Sub

Private Sub mnuRight_Click()
    Me.ActiveForm.rtb1.SelAlignment = 1
    
    tlbStandard.Buttons("Left").Value = tbrUnpressed
    tlbStandard.Buttons("Center").Value = tbrUnpressed
    tlbStandard.Buttons("Right").Value = tbrPressed
    
    mnuLeft.Checked = False
    mnuCenter.Checked = False
    mnuRight.Checked = True
End Sub

Private Sub mnuSave_Click()
    Dim a%
    If Me.ActiveForm.SavedBefore = True Then
        Me.ActiveForm.rtb1.SaveFile (Me.ActiveForm.SavedBeforePath)
        Me.ActiveForm.Saved = True
        Me.ActiveForm.SavedBefore = True
        mnuSave.Enabled = False
        tlbStandard.Buttons("Save").Enabled = False
    Else
        mnuSaveAs_Click
    End If
End Sub

Private Sub mnuSaveAs_Click()
    On Error GoTo ErrorHandling2
    BeforeCaption = Me.ActiveForm.Caption
    With CommonDialog1
        .Flags = &H8000 And &H2
        .Filter = "Rich Text Format|*.rtf"
        .ShowSave
    End With
    With Me.ActiveForm
        .rtb1.SaveFile (CommonDialog1.FileName)
        .Saved = True
        .SavedBefore = True
        .SavedBeforePath = CommonDialog1.FileName
        .Caption = CommonDialog1.FileTitle
    End With
    Me.Caption = "BartNet Editor"
    tlbStandard.Buttons("Save").Enabled = False
    mnuSave.Enabled = False
    Exit Sub

ErrorHandling2:
    If Err.Number = cdlCancel Then
        
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuSelectAll_Click()
    Me.ActiveForm.rtb1.SelStart = 0
    Me.ActiveForm.rtb1.SelLength = Len(Me.ActiveForm.rtb1.Text)
    Me.ActiveForm.rtb1.SetFocus
End Sub

Private Sub mnuSize10_Click()
    Me.ActiveForm.rtb1.SelFontSize = 10
    mnuSize8.Checked = False
    mnuSize9.Checked = False
    mnuSize10.Checked = True
    mnuSize11.Checked = False
    mnuSize12.Checked = False
    mnuSize14.Checked = False
    mnuSize16.Checked = False
    mnuSize18.Checked = False
    mnuSize20.Checked = False
    mnuSize22.Checked = False
    mnuSize24.Checked = False
    mnuSize26.Checked = False
    mnuSize28.Checked = False
    mnuSize30.Checked = False
    mnuSize36.Checked = False
    mnuSize40.Checked = False
    mnuSize48.Checked = False
    mnuSize72.Checked = False
End Sub

Private Sub mnuSize8_Click()
    Me.ActiveForm.rtb1.SelFontSize = 8
    mnuSize8.Checked = True
    mnuSize9.Checked = False
    mnuSize10.Checked = False
    mnuSize11.Checked = False
    mnuSize12.Checked = False
    mnuSize14.Checked = False
    mnuSize16.Checked = False
    mnuSize18.Checked = False
    mnuSize20.Checked = False
    mnuSize22.Checked = False
    mnuSize24.Checked = False
    mnuSize26.Checked = False
    mnuSize28.Checked = False
    mnuSize30.Checked = False
    mnuSize36.Checked = False
    mnuSize40.Checked = False
    mnuSize48.Checked = False
    mnuSize72.Checked = False
End Sub
Private Sub mnuSize9_Click()
    Me.ActiveForm.rtb1.SelFontSize = 9
    mnuSize8.Checked = False
    mnuSize9.Checked = True
    mnuSize10.Checked = False
    mnuSize11.Checked = False
    mnuSize12.Checked = False
    mnuSize14.Checked = False
    mnuSize16.Checked = False
    mnuSize18.Checked = False
    mnuSize20.Checked = False
    mnuSize22.Checked = False
    mnuSize24.Checked = False
    mnuSize26.Checked = False
    mnuSize28.Checked = False
    mnuSize30.Checked = False
    mnuSize36.Checked = False
    mnuSize40.Checked = False
    mnuSize48.Checked = False
    mnuSize72.Checked = False
End Sub

Private Sub mnuStandardToolbar_Click()
    If mnuStandardToolbar.Checked = True Then
        mnuStandardToolbar.Checked = False
        tlbStandard.Visible = False
    Else
        mnuStandardToolbar.Checked = True
        tlbStandard.Visible = True
    End If
End Sub

Private Sub mnuStatusBar_Click()
    If mnuStatusBar.Checked = True Then
        mnuStatusBar.Checked = False
        StatusBar1.Visible = False
    Else
        mnuStatusBar.Checked = True
        StatusBar1.Visible = True
    End If
End Sub

Private Sub mnuStrikeThru_Click()
    If mnuStrikeThru.Checked = True Then
        Me.ActiveForm.rtb1.SelStrikeThru = False
        mnuStrikeThru.Checked = False
        tlbStandard.Buttons("StrikeThru").Value = tbrUnpressed
    Else
        Me.ActiveForm.rtb1.SelStrikeThru = True
        mnuStrikeThru.Checked = True
        tlbStandard.Buttons("StrikeThru").Value = tbrPressed
    End If
End Sub

Private Sub mnuTileHorizontally_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileVertically_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuUnderline_Click()
    If mnuUnderline.Checked = True Then
        Me.ActiveForm.rtb1.SelUnderline = False
        mnuUnderline.Checked = False
        tlbStandard.Buttons("Underline").Value = tbrUnpressed
    Else
        Me.ActiveForm.rtb1.SelUnderline = True
        mnuUnderline.Checked = True
        tlbStandard.Buttons("Underline").Value = tbrPressed
    End If
End Sub
Sub Check()
    If Me.ActiveForm.rtb1.SelBold = True Then
        mnuBold.Checked = True
        tlbStandard.Buttons("Bold").Value = tbrPressed
    Else
        mnuBold.Checked = False
        tlbStandard.Buttons("Bold").Value = tbrUnpressed
    End If
    
    If Me.ActiveForm.rtb1.SelItalic = True Then
        mnuItalic.Checked = True
        tlbStandard.Buttons("Italic").Value = tbrPressed
    Else
        mnuItalic.Checked = False
        tlbStandard.Buttons("Italic").Value = tbrUnpressed
    End If
    
    If Me.ActiveForm.rtb1.SelUnderline = True Then
        mnuUnderline.Checked = True
        tlbStandard.Buttons("Underline").Value = tbrPressed
    Else
        mnuUnderline.Checked = False
        tlbStandard.Buttons("Underline").Value = tbrUnpressed
    End If
    
    If Me.ActiveForm.rtb1.SelStrikeThru = True Then
        mnuStrikeThru.Checked = True
        tlbStandard.Buttons("StrikeThru").Value = tbrPressed
    Else
        mnuStrikeThru.Checked = False
        tlbStandard.Buttons("StrikeThru").Value = tbrUnpressed
    End If
    
    If Me.ActiveForm.rtb1.SelAlignment = 0 Then
        tlbStandard.Buttons("Left").Value = tbrPressed
        tlbStandard.Buttons("Center").Value = tbrUnpressed
        tlbStandard.Buttons("Right").Value = tbrUnpressed
        
        mnuLeft.Checked = True
        mnuCenter.Checked = False
        mnuRight.Checked = False
    Else
        If Me.ActiveForm.rtb1.SelAlignment = 1 Then
            tlbStandard.Buttons("Left").Value = tbrUnpressed
            tlbStandard.Buttons("Center").Value = tbrUnpressed
            tlbStandard.Buttons("Right").Value = tbrPressed
        
            mnuLeft.Checked = False
            mnuCenter.Checked = False
            mnuRight.Checked = True
        Else
            tlbStandard.Buttons("Left").Value = tbrUnpressed
            tlbStandard.Buttons("Center").Value = tbrPressed
            tlbStandard.Buttons("Right").Value = tbrUnpressed
        
            mnuLeft.Checked = False
            mnuCenter.Checked = True
            mnuRight.Checked = False
        End If
        
    End If
    
End Sub

Private Sub mnuSize72_Click()
    Me.ActiveForm.rtb1.SelFontSize = 72
    mnuSize8.Checked = False
    mnuSize9.Checked = False
    mnuSize10.Checked = False
    mnuSize11.Checked = False
    mnuSize12.Checked = False
    mnuSize14.Checked = False
    mnuSize16.Checked = False
    mnuSize18.Checked = False
    mnuSize20.Checked = False
    mnuSize22.Checked = False
    mnuSize24.Checked = False
    mnuSize26.Checked = False
    mnuSize28.Checked = False
    mnuSize30.Checked = False
    mnuSize36.Checked = False
    mnuSize40.Checked = False
    mnuSize48.Checked = False
    mnuSize72.Checked = True
End Sub
Private Sub mnuSize48_Click()
    Me.ActiveForm.rtb1.SelFontSize = 48
    mnuSize8.Checked = False
    mnuSize9.Checked = False
    mnuSize10.Checked = False
    mnuSize11.Checked = False
    mnuSize12.Checked = False
    mnuSize14.Checked = False
    mnuSize16.Checked = False
    mnuSize18.Checked = False
    mnuSize20.Checked = False
    mnuSize22.Checked = False
    mnuSize24.Checked = False
    mnuSize26.Checked = False
    mnuSize28.Checked = False
    mnuSize30.Checked = False
    mnuSize36.Checked = False
    mnuSize40.Checked = False
    mnuSize48.Checked = True
    mnuSize72.Checked = False
End Sub
Private Sub mnuSize40_Click()
    Me.ActiveForm.rtb1.SelFontSize = 40
    mnuSize8.Checked = False
    mnuSize9.Checked = False
    mnuSize10.Checked = False
    mnuSize11.Checked = False
    mnuSize12.Checked = False
    mnuSize14.Checked = False
    mnuSize16.Checked = False
    mnuSize18.Checked = False
    mnuSize20.Checked = False
    mnuSize22.Checked = False
    mnuSize24.Checked = False
    mnuSize26.Checked = False
    mnuSize28.Checked = False
    mnuSize30.Checked = False
    mnuSize36.Checked = False
    mnuSize40.Checked = True
    mnuSize48.Checked = False
    mnuSize72.Checked = False
End Sub
Private Sub mnuSize36_Click()
    Me.ActiveForm.rtb1.SelFontSize = 36
    mnuSize8.Checked = False
    mnuSize9.Checked = False
    mnuSize10.Checked = False
    mnuSize11.Checked = False
    mnuSize12.Checked = False
    mnuSize14.Checked = False
    mnuSize16.Checked = False
    mnuSize18.Checked = False
    mnuSize20.Checked = False
    mnuSize22.Checked = False
    mnuSize24.Checked = False
    mnuSize26.Checked = False
    mnuSize28.Checked = False
    mnuSize30.Checked = False
    mnuSize36.Checked = True
    mnuSize40.Checked = False
    mnuSize48.Checked = False
    mnuSize72.Checked = False
End Sub
Private Sub mnuSize30_Click()
    Me.ActiveForm.rtb1.SelFontSize = 30
    mnuSize8.Checked = False
    mnuSize9.Checked = False
    mnuSize10.Checked = False
    mnuSize11.Checked = False
    mnuSize12.Checked = False
    mnuSize14.Checked = False
    mnuSize16.Checked = False
    mnuSize18.Checked = False
    mnuSize20.Checked = False
    mnuSize22.Checked = False
    mnuSize24.Checked = False
    mnuSize26.Checked = False
    mnuSize28.Checked = False
    mnuSize30.Checked = True
    mnuSize36.Checked = False
    mnuSize40.Checked = False
    mnuSize48.Checked = False
    mnuSize72.Checked = False
End Sub
Private Sub mnuSize28_Click()
    Me.ActiveForm.rtb1.SelFontSize = 28
    mnuSize8.Checked = False
    mnuSize9.Checked = False
    mnuSize10.Checked = False
    mnuSize11.Checked = False
    mnuSize12.Checked = False
    mnuSize14.Checked = False
    mnuSize16.Checked = False
    mnuSize18.Checked = False
    mnuSize20.Checked = False
    mnuSize22.Checked = False
    mnuSize24.Checked = False
    mnuSize26.Checked = False
    mnuSize28.Checked = True
    mnuSize30.Checked = False
    mnuSize36.Checked = False
    mnuSize40.Checked = False
    mnuSize48.Checked = False
    mnuSize72.Checked = False
End Sub
Private Sub mnuSize26_Click()
    Me.ActiveForm.rtb1.SelFontSize = 26
    mnuSize8.Checked = False
    mnuSize9.Checked = False
    mnuSize10.Checked = False
    mnuSize11.Checked = False
    mnuSize12.Checked = False
    mnuSize14.Checked = False
    mnuSize16.Checked = False
    mnuSize18.Checked = False
    mnuSize20.Checked = False
    mnuSize22.Checked = False
    mnuSize24.Checked = False
    mnuSize26.Checked = True
    mnuSize28.Checked = False
    mnuSize30.Checked = False
    mnuSize36.Checked = False
    mnuSize40.Checked = False
    mnuSize48.Checked = False
    mnuSize72.Checked = False
End Sub
Private Sub mnuSize24_Click()
    Me.ActiveForm.rtb1.SelFontSize = 24
    mnuSize8.Checked = False
    mnuSize9.Checked = False
    mnuSize10.Checked = False
    mnuSize11.Checked = False
    mnuSize12.Checked = False
    mnuSize14.Checked = False
    mnuSize16.Checked = False
    mnuSize18.Checked = False
    mnuSize20.Checked = False
    mnuSize22.Checked = False
    mnuSize24.Checked = True
    mnuSize26.Checked = False
    mnuSize28.Checked = False
    mnuSize30.Checked = False
    mnuSize36.Checked = False
    mnuSize40.Checked = False
    mnuSize48.Checked = False
    mnuSize72.Checked = False
End Sub
Private Sub mnuSize22_Click()
    Me.ActiveForm.rtb1.SelFontSize = 22
    mnuSize8.Checked = False
    mnuSize9.Checked = False
    mnuSize10.Checked = False
    mnuSize11.Checked = False
    mnuSize12.Checked = False
    mnuSize14.Checked = False
    mnuSize16.Checked = False
    mnuSize18.Checked = False
    mnuSize20.Checked = False
    mnuSize22.Checked = True
    mnuSize24.Checked = False
    mnuSize26.Checked = False
    mnuSize28.Checked = False
    mnuSize30.Checked = False
    mnuSize36.Checked = False
    mnuSize40.Checked = False
    mnuSize48.Checked = False
    mnuSize72.Checked = False
End Sub
Private Sub mnuSize20_Click()
    Me.ActiveForm.rtb1.SelFontSize = 20
    mnuSize8.Checked = False
    mnuSize9.Checked = False
    mnuSize10.Checked = False
    mnuSize11.Checked = False
    mnuSize12.Checked = False
    mnuSize14.Checked = False
    mnuSize16.Checked = False
    mnuSize18.Checked = False
    mnuSize20.Checked = True
    mnuSize22.Checked = False
    mnuSize24.Checked = False
    mnuSize26.Checked = False
    mnuSize28.Checked = False
    mnuSize30.Checked = False
    mnuSize36.Checked = False
    mnuSize40.Checked = False
    mnuSize48.Checked = False
    mnuSize72.Checked = False
End Sub
Private Sub mnuSize18_Click()
    Me.ActiveForm.rtb1.SelFontSize = 18
    mnuSize8.Checked = False
    mnuSize9.Checked = False
    mnuSize10.Checked = False
    mnuSize11.Checked = False
    mnuSize12.Checked = False
    mnuSize14.Checked = False
    mnuSize16.Checked = False
    mnuSize18.Checked = True
    mnuSize20.Checked = False
    mnuSize22.Checked = False
    mnuSize24.Checked = False
    mnuSize26.Checked = False
    mnuSize28.Checked = False
    mnuSize30.Checked = False
    mnuSize36.Checked = False
    mnuSize40.Checked = False
    mnuSize48.Checked = False
    mnuSize72.Checked = False
End Sub
Private Sub mnuSize16_Click()
    Me.ActiveForm.rtb1.SelFontSize = 16
    mnuSize8.Checked = False
    mnuSize9.Checked = False
    mnuSize10.Checked = False
    mnuSize11.Checked = False
    mnuSize12.Checked = False
    mnuSize14.Checked = False
    mnuSize16.Checked = True
    mnuSize18.Checked = False
    mnuSize20.Checked = False
    mnuSize22.Checked = False
    mnuSize24.Checked = False
    mnuSize26.Checked = False
    mnuSize28.Checked = False
    mnuSize30.Checked = False
    mnuSize36.Checked = False
    mnuSize40.Checked = False
    mnuSize48.Checked = False
    mnuSize72.Checked = False
End Sub
Private Sub mnuSize14_Click()
    Me.ActiveForm.rtb1.SelFontSize = 14
    mnuSize8.Checked = False
    mnuSize9.Checked = False
    mnuSize10.Checked = False
    mnuSize11.Checked = False
    mnuSize12.Checked = False
    mnuSize14.Checked = True
    mnuSize16.Checked = False
    mnuSize18.Checked = False
    mnuSize20.Checked = False
    mnuSize22.Checked = False
    mnuSize24.Checked = False
    mnuSize26.Checked = False
    mnuSize28.Checked = False
    mnuSize30.Checked = False
    mnuSize36.Checked = False
    mnuSize40.Checked = False
    mnuSize48.Checked = False
    mnuSize72.Checked = False
End Sub
Private Sub mnuSize12_Click()
    Me.ActiveForm.rtb1.SelFontSize = 12
    mnuSize8.Checked = False
    mnuSize9.Checked = False
    mnuSize10.Checked = False
    mnuSize11.Checked = False
    mnuSize12.Checked = True
    mnuSize14.Checked = False
    mnuSize16.Checked = False
    mnuSize18.Checked = False
    mnuSize20.Checked = False
    mnuSize22.Checked = False
    mnuSize24.Checked = False
    mnuSize26.Checked = False
    mnuSize28.Checked = False
    mnuSize30.Checked = False
    mnuSize36.Checked = False
    mnuSize40.Checked = False
    mnuSize48.Checked = False
    mnuSize72.Checked = False
End Sub
Private Sub mnuSize11_Click()
    Me.ActiveForm.rtb1.SelFontSize = 11
    mnuSize8.Checked = False
    mnuSize9.Checked = False
    mnuSize10.Checked = False
    mnuSize11.Checked = True
    mnuSize12.Checked = False
    mnuSize14.Checked = False
    mnuSize16.Checked = False
    mnuSize18.Checked = False
    mnuSize20.Checked = False
    mnuSize22.Checked = False
    mnuSize24.Checked = False
    mnuSize26.Checked = False
    mnuSize28.Checked = False
    mnuSize30.Checked = False
    mnuSize36.Checked = False
    mnuSize40.Checked = False
    mnuSize48.Checked = False
    mnuSize72.Checked = False
End Sub

Private Sub mnuUndo_Click()
    Me.ActiveForm.Undo

    mnuRedo.Enabled = True
    tlbStandard.Buttons("Redo").Enabled = True
End Sub

Private Sub StatusBar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuView
    End If
End Sub

Private Sub tlbStandard_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "New"
            mnuNew_Click
        Case "Open"
            mnuOpen_Click
        Case "Save"
            mnuSave_Click
        Case "Cut"
            mnuCut_Click
        Case "Copy"
            mnuCopy_Click
        Case "Paste"
            mnuPaste_Click
        Case "Find"
            mnuFind_Click
        Case "Undo"
            mnuUndo_Click
        Case "Redo"
            mnuRedo_Click
        Case "Help"
            mnuHelp_Click
        Case "Color"
            mnuFontColor_Click
        Case "StrikeThru"
            mnuStrikeThru_Click
        Case "Underline"
            mnuUnderline_Click
        Case "Italic"
            mnuItalic_Click
        Case "Bold"
            mnuBold_Click
        Case "Left"
            mnuLeft_Click
        Case "Center"
            mnuCenter_Click
        Case "Right"
            mnuRight_Click
        Case "IncreaseIndent"
            mnuIncreaseIndent_Click
        Case "DecreaseIndent"
            mnuDecreaseIndent_Click
        End Select
End Sub

Private Sub tlbStandard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuView
    End If
End Sub
