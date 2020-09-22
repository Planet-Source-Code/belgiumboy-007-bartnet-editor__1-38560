Attribute VB_Name = "modStartUp"
Public StartUpScreenShow As Boolean
Public ShowToolBar As Boolean
Public ShowStatusBar As Boolean
Public DocumentName As String
Public DefaultFont As String
Public DefaultFontSize As Integer
Public DefaultTextColor As String
Public DefaultBackgroundColor As String
Public DefaultBold As Boolean
Public DefaultItalic As Boolean
Public DefaultUnderline As Boolean
Public DefaultStrikeThru As Boolean
Public DefaultAlignment As String
Public IndentSize As Integer

Private fso As New FileSystemObject

Private s As TextStream

Private m As String
Private mr As String
Private mrt As String
Private mrts As String
Private mrtsm As String
Private mrtsmr As String
Private mrtsmrt As String
Private mrtsmrts As String
Private mrtsmrtsm As String
Private mrtsmrtsmr As String
Private mrtsmrtsmrt As String
Private mrtsmrtsmrts As String
Private mrtsmrtsmrtsm As String
Private mrtsmrtsmrtsmr As String
Sub main()
    Set s = fso.OpenTextFile(App.Path & "\" & "Info.BartNet")

    With s
        m = .ReadLine
        mr = .ReadLine
        mrt = .ReadLine
        mrts = .ReadLine
        mrtsm = .ReadLine
        mrtsmr = .ReadLine
        mrtsmrt = .ReadLine
        mrtsmrts = .ReadLine
        mrtsmrtsm = .ReadLine
        mrtsmrtsmr = .ReadLine
        mrtsmrtsmrt = .ReadLine
        mrtsmrtsmrts = .ReadLine
        mrtsmrtsmrtsm = .ReadLine
        mrtsmrtsmrtsmr = .ReadLine
    End With
    
    StartUpScreenShow = m
    ShowToolBar = mr
    ShowStatusBar = mrt
    DocumentName = mrts
    DefaultFont = mrtsm
    DefaultFontSize = mrtsmr
    DefaultTextColor = mrtsmrt
    DefaultBackgroundColor = mrtsmrts
    DefaultBold = mrtsmrtsm
    DefaultItalic = mrtsmrtsmr
    DefaultUnderline = mrtsmrtsmrt
    DefaultStrikeThru = mrtsmrtsmrts
    DefaultAlignment = mrtsmrtsmrtsm
    IndentSize = mrtsmrtsmrtsmr
    
    frmSplash.Show
End Sub

