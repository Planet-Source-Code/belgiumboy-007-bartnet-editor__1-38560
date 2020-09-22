VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmBrowser 
   Caption         =   "http://www.BartNet.freeservers.com"
   ClientHeight    =   7695
   ClientLeft      =   165
   ClientTop       =   660
   ClientWidth     =   9045
   Icon            =   "frmBrowser2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7695
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboURL 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   840
      Visible         =   0   'False
      Width           =   8535
   End
   Begin MSComctlLib.Toolbar tlbBrowser 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   900
      ButtonWidth     =   820
      ButtonHeight    =   794
      AllowCustomize  =   0   'False
      ImageList       =   "imlIcons(0)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Terug1"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Vooruit"
            Object.ToolTipText     =   "Vooruit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Annuleren"
            Object.ToolTipText     =   "Annuleren"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bijwerken"
            Object.ToolTipText     =   "Bijwerken"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser brwBrowser 
      Height          =   3255
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   7335
      ExtentX         =   12938
      ExtentY         =   5741
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer timBrowser 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   7320
      Top             =   1560
   End
   Begin MSComctlLib.StatusBar stbBrowser 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7320
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser2.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser2.frx":0724
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser2.frx":0A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser2.frx":0CE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser2.frx":0FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser2.frx":12AC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsBrowser 
      Index           =   1
      Left            =   6360
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser2.frx":158E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser2.frx":1870
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser2.frx":1B52
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser2.frx":1E34
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser2.frx":2116
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser2.frx":23F8
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Saved As Boolean
Public SavedBefore As Boolean
Dim mblnNavigeren As Boolean

Private Sub brwBrowser_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, headers As Variant, Cancel As Boolean)
     If InStr(URL, "smut") > 0 Then
        MsgBox "Deze URL is geblokkeerd.", _
            vbOKOnly Or vbExclamation Or vbMsgBoxSetForeground
        Cancel = True
     Else
        timBrowser.Enabled = True
        Me.MousePointer = vbHourglass
    End If
End Sub

Private Sub Form_Load()
    Saved = True
    SavedBefore = True
    On Error Resume Next

    mblnNavigeren = True
    Me.WindowState = vbMaximized
    cboURL.Text = ""
    Me.Show
    brwBrowser.Navigate "http://www.bartnet.freeservers.com"
    
    Me.Caption = "http://www.BartNet.freeservers.com"

    If Err.Number <> 0 Then stbBrowser.SimpleText = "FOUT: " & Err.Description
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        cboURL.Move 0, tlbBrowser.Height, Me.ScaleWidth
        brwBrowser.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    End If
End Sub

Private Sub timBrowser_Timer()
    brwBrowser.Stop
    Me.MousePointer = vbDefault
    MsgBox "aanvraag geannuleerd, duurde te lang.", vbOKOnly Or vbExclamation Or vbMsgBoxSetForeground
    timBrowser.Enabled = False
End Sub

Private Sub brwBrowser_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    Dim intURLindex As Integer
    Dim blnReedsInlijst As Boolean
    
    timBrowser.Enabled = False
    blnReedsInlijst = False
    
    For intURLindex = 0 To cboURL.ListCount - 1
        If cboURL.List(intURLindex) = brwBrowser.LocationURL Then
            blnReedsInlijst = True
            Exit For
        End If
    Next
    mblnNavigeren = False
    If blnReedsInlijst Then cboURL.RemoveItem intURLindex
    cboURL.AddItem brwBrowser.LocationURL, 0
    cboURL.ListIndex = 0
    mblnNavigeren = True
       
End Sub



Private Sub brwBrowser_DownloadComplete()
    MousePointer = vbDefault
End Sub

Private Sub brwBrowser_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    If Progress = -1 Then
        stbBrowser.SimpleText = "klaar"
    ElseIf ProgressMax = 0 Then
        stbBrowser.SimpleText = ""
    Else
        stbBrowser.SimpleText = Progress / ProgressMax * 100 & "%"
    End If
End Sub

Private Sub brwBrowser_StatusTextChange(ByVal Text As String)
stbBrowser.SimpleText = Text
End Sub

Private Sub tlbBrowser_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error Resume Next

    Select Case Button.Key
        Case "Terug"
            brwBrowser.GoBack
        Case "Vooruit"
            brwBrowser.GoForward
        Case "annleren"
            brwBrowser.Stop
        Case "Bijwerken"
            brwBrowser.Refresh
        End Select
        
        If Err.Number <> 0 Then _
            stbBrowser.SimpleText = "fout: " & Err.Description

End Sub

