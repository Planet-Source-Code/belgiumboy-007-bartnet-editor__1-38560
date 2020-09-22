VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timSplash 
      Interval        =   1000
      Left            =   2040
      Top             =   1320
   End
   Begin MSComctlLib.ImageList ilsSplash 
      Left            =   3240
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   400
      ImageHeight     =   221
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSplash.frx":0000
            Key             =   "Splash"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgSplash 
      Height          =   3315
      Left            =   0
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Seconds%
Private Sub Form_Load()
    imgSplash.Picture = ilsSplash.ListImages("Splash").Picture
    frmStartUp.OpenedFromStartUpScreen = False
End Sub

Private Sub timSplash_Timer()
    If Seconds = 3 Then
        frmSplash.Hide
        frmBartNetEditor.Show
        If ShowToolBar = True Then
            frmBartNetEditor.tlbStandard.Visible = True
            frmBartNetEditor.mnuStandardToolbar.Checked = True
        Else
            frmBartNetEditor.tlbStandard.Visible = False
            frmBartNetEditor.mnuStandardToolbar.Checked = False
        End If
        If ShowStatusBar = True Then
            frmBartNetEditor.StatusBar1.Visible = True
            frmBartNetEditor.mnuStatusBar.Checked = True
        Else
            frmBartNetEditor.StatusBar1.Visible = False
            frmBartNetEditor.mnuStatusBar.Checked = False
        End If
        frmBartNetEditor.WindowState = vbMaximized
        If StartUpScreenShow = True Then
            frmStartUp.Show vbModal, frmBartNetEditor
        Else
            frmBartNetEditor.PageNumber = frmBartNetEditor.PageNumber + 1
            Set frmNewPage = New frmNewPage1
            With frmNewPage
            .Caption = DocumentName & " " & frmBartNetEditor.PageNumber
            .Show
            .WindowState = vbMaximized
            .Saved = True
            .SavedBefore = False
        End With
        End If
        Unload Me
        
        Select Case DefaultFontSize
            Case 8
                frmBartNetEditor.mnuSize8.Checked = True
            Case 9
                frmBartNetEditor.mnuSize9.Checked = True
            Case 10
                frmBartNetEditor.mnuSize10.Checked = True
            Case 11
                frmBartNetEditor.mnuSize11.Checked = True
            Case 12
                frmBartNetEditor.mnuSize12.Checked = True
            Case 14
                frmBartNetEditor.mnuSize14.Checked = True
            Case 16
                frmBartNetEditor.mnuSize16.Checked = True
            Case 18
                frmBartNetEditor.mnuSize18.Checked = True
            Case 20
                frmBartNetEditor.mnuSize20.Checked = True
            Case 22
                frmBartNetEditor.mnuSize22.Checked = True
            Case 24
                frmBartNetEditor.mnuSize24.Checked = True
            Case 26
                frmBartNetEditor.mnuSize26.Checked = True
            Case 28
                frmBartNetEditor.mnuSize28.Checked = True
            Case 30
                frmBartNetEditor.mnuSize30.Checked = True
            Case 36
                frmBartNetEditor.mnuSize36.Checked = True
            Case 40
                frmBartNetEditor.mnuSize40.Checked = True
            Case 48
                frmBartNetEditor.mnuSize48.Checked = True
            Case 72
                frmBartNetEditor.mnuSize72.Checked = True
            End Select
            
            If DefaultBold = True Then
                frmBartNetEditor.mnuBold.Checked = True
                frmBartNetEditor.tlbStandard.Buttons("Bold").Value = tbrPressed
            End If
            If DefaultItalic = True Then
                frmBartNetEditor.mnuItalic.Checked = True
                frmBartNetEditor.tlbStandard.Buttons("Italic").Value = tbrPressed
            End If
            If DefaultUnderline = True Then
                frmBartNetEditor.mnuUnderline.Checked = True
                frmBartNetEditor.tlbStandard.Buttons("Underline").Value = tbrPressed
            End If
            If DefaultStrikeThru = True Then
                frmBartNetEditor.mnuStrikeThru.Checked = True
                frmBartNetEditor.tlbStandard.Buttons("StrikeThru").Value = tbrPressed
            End If
            
            Select Case DefaultAlignment
                Case "Left"
                    frmBartNetEditor.ActiveForm.rtb1.SelAlignment = 0
                    frmBartNetEditor.mnuLeft.Checked = True
                    frmBartNetEditor.mnuCenter.Checked = False
                    frmBartNetEditor.mnuRight.Checked = False
                    frmBartNetEditor.tlbStandard.Buttons("Left").Value = tbrPressed
                    frmBartNetEditor.tlbStandard.Buttons("Center").Value = tbrUnpressed
                    frmBartNetEditor.tlbStandard.Buttons("Right").Value = tbrUnpressed
                Case "Center"
                    frmBartNetEditor.ActiveForm.rtb1.SelAlignment = 2
                    frmBartNetEditor.mnuLeft.Checked = False
                    frmBartNetEditor.mnuCenter.Checked = True
                    frmBartNetEditor.mnuRight.Checked = False
                    frmBartNetEditor.tlbStandard.Buttons("Left").Value = tbrUnpressed
                    frmBartNetEditor.tlbStandard.Buttons("Center").Value = tbrPressed
                    frmBartNetEditor.tlbStandard.Buttons("Right").Value = tbrUnpressed
                Case "Right"
                    frmBartNetEditor.ActiveForm.rtb1.SelAlignment = 1
                    frmBartNetEditor.mnuLeft.Checked = False
                    frmBartNetEditor.mnuCenter.Checked = False
                    frmBartNetEditor.mnuRight.Checked = True
                    frmBartNetEditor.tlbStandard.Buttons("Left").Value = tbrUnpressed
                    frmBartNetEditor.tlbStandard.Buttons("Center").Value = tbrUnpressed
                    frmBartNetEditor.tlbStandard.Buttons("Right").Value = tbrPressed
            End Select

    Else
        Seconds = Seconds + 1
    End If
End Sub
