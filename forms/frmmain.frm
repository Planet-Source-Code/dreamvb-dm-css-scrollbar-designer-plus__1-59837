VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmmain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DM CSS Scrollbar Designer Plus"
   ClientHeight    =   4800
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9525
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "CSS Code"
      Height          =   1875
      Left            =   4680
      TabIndex        =   32
      Top             =   2130
      Width           =   4725
      Begin VB.TextBox txtCode 
         Height          =   1485
         Left            =   150
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   255
         Width           =   4455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Preview Area"
      Height          =   1935
      Left            =   4680
      TabIndex        =   27
      Top             =   75
      Width           =   4725
      Begin SHDocVwCtl.WebBrowser WebView 
         Height          =   1500
         Left            =   165
         TabIndex        =   28
         Top             =   285
         Width           =   4425
         ExtentX         =   7805
         ExtentY         =   2646
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
   End
   Begin VB.PictureBox PicA 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3630
      Left            =   195
      Picture         =   "frmmain.frx":1042
      ScaleHeight     =   242
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Frame Frame1 
      Caption         =   "Design Area"
      Height          =   4290
      Left            =   30
      TabIndex        =   0
      Top             =   75
      Width           =   4560
      Begin VB.PictureBox PicCol 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   8
         Left            =   3390
         ScaleHeight     =   195
         ScaleWidth      =   885
         TabIndex        =   31
         Top             =   3525
         Width           =   915
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "...."
         Height          =   270
         Index           =   8
         Left            =   2715
         TabIndex        =   30
         Top             =   3525
         Width           =   330
      End
      Begin VB.PictureBox PicCol 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   3390
         ScaleHeight     =   195
         ScaleWidth      =   885
         TabIndex        =   25
         Top             =   315
         Width           =   915
      End
      Begin VB.PictureBox PicCol 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   7
         Left            =   3390
         ScaleHeight     =   195
         ScaleWidth      =   885
         TabIndex        =   24
         Top             =   3105
         Width           =   915
      End
      Begin VB.PictureBox PicCol 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   3390
         ScaleHeight     =   195
         ScaleWidth      =   885
         TabIndex        =   23
         Top             =   2685
         Width           =   915
      End
      Begin VB.PictureBox PicCol 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   5
         Left            =   3390
         ScaleHeight     =   195
         ScaleWidth      =   885
         TabIndex        =   22
         Top             =   2265
         Width           =   915
      End
      Begin VB.PictureBox PicCol 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   4
         Left            =   3390
         ScaleHeight     =   195
         ScaleWidth      =   885
         TabIndex        =   21
         Top             =   1860
         Width           =   915
      End
      Begin VB.PictureBox PicCol 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   3390
         ScaleHeight     =   195
         ScaleWidth      =   885
         TabIndex        =   20
         Top             =   1455
         Width           =   915
      End
      Begin VB.PictureBox PicCol 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   3390
         ScaleHeight     =   195
         ScaleWidth      =   885
         TabIndex        =   19
         Top             =   1050
         Width           =   915
      End
      Begin VB.PictureBox PicCol 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   3390
         ScaleHeight     =   195
         ScaleWidth      =   885
         TabIndex        =   18
         Top             =   690
         Width           =   915
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "...."
         Height          =   270
         Index           =   7
         Left            =   2715
         TabIndex        =   10
         Top             =   3105
         Width           =   330
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "...."
         Height          =   270
         Index           =   6
         Left            =   2715
         TabIndex        =   9
         Top             =   2685
         Width           =   330
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "...."
         Height          =   270
         Index           =   5
         Left            =   2715
         TabIndex        =   8
         Top             =   2280
         Width           =   330
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "...."
         Height          =   270
         Index           =   4
         Left            =   2715
         TabIndex        =   7
         Top             =   1860
         Width           =   330
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "...."
         Height          =   270
         Index           =   3
         Left            =   2715
         TabIndex        =   6
         Top             =   1455
         Width           =   330
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "...."
         Height          =   270
         Index           =   2
         Left            =   2715
         TabIndex        =   5
         Top             =   1050
         Width           =   330
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "...."
         Height          =   270
         Index           =   1
         Left            =   2715
         TabIndex        =   4
         Top             =   690
         Width           =   330
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "...."
         Height          =   270
         Index           =   0
         Left            =   2715
         TabIndex        =   3
         Top             =   315
         Width           =   330
      End
      Begin VB.PictureBox PicB 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3630
         Left            =   165
         MouseIcon       =   "frmmain.frx":54CC
         MousePointer    =   99  'Custom
         ScaleHeight     =   242
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   66
         TabIndex        =   2
         Top             =   300
         Width           =   990
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         Caption         =   "Page Font Color"
         Height          =   195
         Index           =   8
         Left            =   1410
         TabIndex        =   29
         Top             =   3525
         Width           =   1140
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         Caption         =   "Arrow"
         Height          =   195
         Index           =   0
         Left            =   1410
         TabIndex        =   26
         Top             =   315
         Width           =   405
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         Caption         =   "Page Back Color"
         Height          =   195
         Index           =   7
         Left            =   1410
         TabIndex        =   17
         Top             =   3105
         Width           =   1200
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         Caption         =   "Track Bar"
         Height          =   195
         Index           =   6
         Left            =   1410
         TabIndex        =   16
         Top             =   2685
         Width           =   705
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         Caption         =   "Dark Shadow"
         Height          =   195
         Index           =   5
         Left            =   1410
         TabIndex        =   15
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         Caption         =   "Shadow"
         Height          =   195
         Index           =   4
         Left            =   1410
         TabIndex        =   14
         Top             =   1860
         Width           =   585
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         Caption         =   "3D Light"
         Height          =   195
         Index           =   3
         Left            =   1410
         TabIndex        =   13
         Top             =   1455
         Width           =   600
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         Caption         =   "Highlight"
         Height          =   195
         Index           =   2
         Left            =   1410
         TabIndex        =   12
         Top             =   1050
         Width           =   615
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         Caption         =   "Face"
         Height          =   195
         Index           =   1
         Left            =   1410
         TabIndex        =   11
         Top             =   690
         Width           =   360
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DM CSS Scrollbar Designer Plus"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   4500
      Width           =   2760
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   330
      Left            =   30
      Top             =   4440
      Width           =   9390
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   300
      Left            =   4680
      Top             =   4065
      Width           =   4740
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   35
      Y1              =   1
      Y2              =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   35
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblCopy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copy To Clipboard"
      Height          =   195
      Left            =   4740
      MouseIcon       =   "frmmain.frx":561E
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Top             =   4110
      Width           =   1305
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnunew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnublank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuproj 
         Caption         =   "&Project"
         Begin VB.Menu mnuopen 
            Caption         =   "&Open Project"
         End
         Begin VB.Menu mnusave 
            Caption         =   "&Save Project"
         End
      End
      Begin VB.Menu mnucss 
         Caption         =   "Save CSS Code"
      End
      Begin VB.Menu mnublank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnublank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Type SBarType ' Type information for the project file
    Sig As String
    Arrow As Long
    Light3D As Long
    Face As Long
    Highlight As Long
    Shadow As Long
    DarkShadow As Long
    Trackbar As Long
    HtmlBackColor As Long
    HtmlFontColor As Long
End Type

Private Type tRGB 'RGB Type
    Red As Long
    Green As Long
    Blue As Long
End Type

Private TDialog As New CDialog
Private WebBar As SBarType
Private mRgb As tRGB
Private sTempFile As String 'Temp heml file for preview

Public Function LngToHex(hDecCol As Long) As String
Dim StrHex As String
    ' Used to convert a long vb color to a HTML hex value
    StrHex = Hex(hDecCol)
    
    Do While Len(StrHex) < 6 ' while length of string is lower than 6
        StrHex = "0" & StrHex ' keep adding a zero to left side
        DoEvents ' let system carray our the tasks
    Loop
    
    LngToHex = "#" & Right(StrHex, 2) & Mid(StrHex, 3, 2) & Left(StrHex, 2)
    StrHex = ""
    
End Function

Private Function BuildCSS() As String
Dim StrA As String
    'Build CSS Style sheet code
    StrA = "<STYLE TYPE=" & Chr(34) & "text/css" & Chr(34) & ">" & vbCrLf
    StrA = StrA & "<!--" & vbCrLf
    StrA = StrA & "BODY{" & vbCrLf
    StrA = StrA & "scrollbar-arrow-color:" & LngToHex(WebBar.Arrow) & ";" & vbCrLf
    StrA = StrA & "scrollbar-highlight-color:" & LngToHex(WebBar.Highlight) & ";" & vbCrLf
    StrA = StrA & "scrollbar-face-color:" & LngToHex(WebBar.Face) & ";" & vbCrLf
    StrA = StrA & "scrollbar-3dlight-color:" & LngToHex(WebBar.Light3D) & ";" & vbCrLf
    StrA = StrA & "scrollbar-track-color:" & LngToHex(WebBar.Trackbar) & ";" & vbCrLf
    StrA = StrA & "scrollbar-darkshadow-color:" & LngToHex(WebBar.DarkShadow) & ";" & vbCrLf
    StrA = StrA & "scrollbar-shadow-color:" & LngToHex(WebBar.Shadow) & ";" & vbCrLf
    StrA = StrA & "background-color:" & LngToHex(WebBar.HtmlBackColor) & ";" & vbCrLf
    StrA = StrA & "color:" & LngToHex(WebBar.HtmlFontColor) & ";" & vbCrLf
    StrA = StrA & "}" & vbCrLf
    StrA = StrA & "-->" & vbCrLf
    StrA = StrA & "</STYLE>" & vbCrLf
    
    BuildCSS = StrA
    
End Function

Public Function IsFileHere(lzFileName As String) As Boolean
    ' check if a given file is found
    If Dir(lzFileName) = "" Then IsFileHere = False: Exit Function Else IsFileHere = True
End Function

Function FixPath(lzPath As String) As String
    'add a bacl slash to a given path if required
    If Right(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Private Function DeleteTemp()
On Error GoTo TempDel:
    If IsFileHere(sTempFile) Then Kill sTempFile ' check if the temp file is found if so delete it
    Exit Function
TempDel:
    ' show an error if delete was not possiable
    If Err Then MsgBox Err.Description, vbExclamation, "Error removeing temp file"
    
End Function
Private Sub MakeTmpFile()
Dim HtmlRes As String

      HtmlRes = StrConv(LoadResData(101, "CUSTOM"), vbUnicode) ' get the html template from the resource file
      HtmlRes = Replace(HtmlRes, "<!--CSS_TAG-->", BuildCSS, , , vbTextCompare) ' find tag and replace with CSS code
      
      Open sTempFile For Output As #1 'create a temp file
        Print #1, HtmlRes ' save the html data to the file
      Close #1 'close file
      
      HtmlRes = "" 'clean up
      txtCode.Text = BuildCSS ' output CSS text to textbox
      WebView.Navigate sTempFile ' display the temp page in the webbroswer control
End Sub

Private Sub MakeBoldFont(Index As Integer)
Dim I As Integer

    For I = 0 To lblA.Count - 1
        lblA(I).FontBold = False ' Remove bold style of all labels
    Next
    
    I = 0
    lblA(Index).FontBold = True ' add bold to style to selected label
    
End Sub

Sub TrackColor(X As Single, Y As Single)
    GetRGB GetPixel(PicA.hdc, X, Y)
    ' this sub find out what color we selected in the scroolbar
    ' and turns on bold text for that color value
    Select Case RGB(mRgb.Red, mRgb.Green, mRgb.Blue)
        Case vbBlack: MakeBoldFont 0
        Case vbYellow: MakeBoldFont 1
        Case vbWhite: MakeBoldFont 2
        Case vbBlue: MakeBoldFont 3
        Case vbCyan: MakeBoldFont 4
        Case vbGreen: MakeBoldFont 5
        Case vbRed: MakeBoldFont 6
    End Select
    
End Sub

Sub GetRGB(lnCol As Long)
Dim mByte(2) As Byte
    ' Extract the rgb color from a long color
    CopyMemory mByte(0), lnCol, Len(lnCol)
    mRgb.Red = mByte(0)
    mRgb.Green = mByte(1)
    mRgb.Blue = mByte(2)
    
End Sub

Sub ColorBar()
Dim X As Long, Y As Long

    BitBlt PicB.hdc, 0, 0, PicA.Width, PicA.Height, PicA.hdc, 0, 0, vbSrcCopy
    
    For X = 0 To PicB.ScaleWidth - 1
        For Y = 0 To PicB.ScaleHeight - 1
            GetRGB GetPixel(PicB.hdc, X, Y) ' Get the rgb color
            
            Select Case RGB(mRgb.Red, mRgb.Green, mRgb.Blue)
                Case vbBlack: SetPixel PicB.hdc, X, Y, WebBar.Arrow
                Case vbBlue: SetPixel PicB.hdc, X, Y, WebBar.Light3D '3D Light
                Case vbYellow: SetPixel PicB.hdc, X, Y, WebBar.Face 'Face
                Case vbWhite: SetPixel PicB.hdc, X, Y, WebBar.Highlight ' Highlight
                Case vbCyan: SetPixel PicB.hdc, X, Y, WebBar.Shadow ' Shadow
                Case vbGreen: SetPixel PicB.hdc, X, Y, WebBar.DarkShadow ' DarkShadow
                Case vbRed: SetPixel PicB.hdc, X, Y, WebBar.Trackbar ' Trackbar
            End Select
            DoEvents
        Next
    Next
    'update the pictureboxes with the correct color
    PicCol(0).BackColor = WebBar.Arrow
    PicCol(1).BackColor = WebBar.Face
    PicCol(2).BackColor = WebBar.Highlight
    PicCol(3).BackColor = WebBar.Light3D
    PicCol(4).BackColor = WebBar.Shadow
    PicCol(5).BackColor = WebBar.DarkShadow
    PicCol(6).BackColor = WebBar.Trackbar
    PicCol(7).BackColor = WebBar.HtmlBackColor
    PicCol(8).BackColor = WebBar.HtmlFontColor
    X = 0: Y = 0
    PicB.Refresh
    ' Generate the CSS code
    Call MakeTmpFile
    
End Sub

Sub DoDefault()
Dim X As Long, Y As Long, Col As Long
    ' Fill WebBar type with default data
    WebBar.Arrow = 0
    WebBar.Light3D = 12632256
    WebBar.Face = 12632256
    WebBar.Highlight = 16777215
    WebBar.Shadow = 6250335
    WebBar.DarkShadow = 4144959
    WebBar.Trackbar = 16777215
    WebBar.HtmlBackColor = 16777215
    WebBar.HtmlFontColor = 0
End Sub

Private Sub cmdButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicCol_MouseDown Index, Button, Shift, X, Y
End Sub

Private Sub Form_Load()
    'fill in our dialog type
    With TDialog
        .hInst = App.hInstance
        .DlgHwnd = frmmain.hWnd
        .flags = 0
        .InitialDir = FixPath(App.Path)
    End With
    
    WebBar.Sig = "sbp" 'sig for scrollbar project file
    sTempFile = FixPath(App.Path) & "temp.html" ' used as a temp file
    DoDefault ' set up the default scrollbar color
    ColorBar ' update scrollbar picture
    
    Line1(0).X2 = frmmain.Width
    Line1(1).X2 = frmmain.Width
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCopy.FontUnderline = False
    lblCopy.ForeColor = vbBlack
End Sub

Private Sub Form_Terminate()
    DeleteTemp ' delete html temp file
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmmain = Nothing
    Form_Terminate
    End
End Sub

Private Sub lblCopy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText txtCode.Text
    MsgBox "CSS code has now been copiyed to the clipboard.", vbInformation, "Copy"
    
End Sub

Private Sub lblCopy_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCopy.FontUnderline = True
    lblCopy.ForeColor = vbRed
End Sub

Private Sub mnuabout_Click()
    frmabout.Show vbModal, frmmain
End Sub

Private Sub mnucss_Click()
Dim nFile As Long
' open a scrollbar project file
    TDialog.DialogTitle = "Save Style sheet code"
    TDialog.Filter = "Style sheet (*.css)" + Chr$(0) + "*.css" + Chr$(0)
    TDialog.ShowSave
    If Not TDialog.CancelError Then Exit Sub
    
    nFile = FreeFile
    Open TDialog.FileName & ".css" For Output As #nFile
        Print #nFile, txtCode.Text
    Close #nFile
    
End Sub

Private Sub mnuexit_Click()
    Unload frmmain
End Sub

Private Sub mnunew_Click()
    If MsgBox("Do you want to reset the scrollbar to it's default style?", vbYesNo Or vbQuestion) = vbNo Then Exit Sub
    DoDefault
    ColorBar
End Sub

Private Sub mnuopen_Click()
Dim nFile As Long
' open a scrollbar project file
    TDialog.DialogTitle = "Open Scrollbar Project"
    TDialog.Filter = "Scrollbar projects (*.sbp)" + Chr$(0) + "*.sbp" + Chr$(0)
    TDialog.ShowOpen
    If Not TDialog.CancelError Then Exit Sub
    
    nFile = FreeFile
    Open TDialog.FileName For Binary As #nFile
        Get #nFile, , WebBar
    Close #nFile
    
    If WebBar.Sig <> "sbp" Then
        MsgBox "Unable to open project.", vbExclamation, "error"
        Exit Sub
    Else
        ColorBar
    End If
    
End Sub

Private Sub mnusave_Click()
Dim nFile As Long

    TDialog.DialogTitle = "Save Scrollbar Project"
    TDialog.Filter = "Scrollbar projects (*.sbp)" + Chr$(0) + "*.sbp" + Chr$(0)
    TDialog.ShowSave
    If Not TDialog.CancelError Then Exit Sub
    
    nFile = FreeFile
    'save the scrollbar project file
    Open TDialog.FileName & ".sbp" For Binary As #nFile
        Put #nFile, , WebBar
    Close #nFile
    
End Sub

Private Sub PicB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    Call TrackColor(X, Y)
End Sub

Private Sub PicCol_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TheColor As Long
    If Button <> vbLeftButton Then Exit Sub
    
    TDialog.ShowColor
    TheColor = TDialog.Color
    If TDialog.CancelError Then Exit Sub
    PicCol(Index).BackColor = TheColor

    ' set the colors
    WebBar.Arrow = PicCol(0).BackColor
    WebBar.Face = PicCol(1).BackColor
    WebBar.Highlight = PicCol(2).BackColor
    WebBar.Light3D = PicCol(3).BackColor
    WebBar.Shadow = PicCol(4).BackColor
    WebBar.DarkShadow = PicCol(5).BackColor
    WebBar.Trackbar = PicCol(6).BackColor
    WebBar.HtmlBackColor = PicCol(7).BackColor
    WebBar.HtmlFontColor = PicCol(8).BackColor
    ColorBar
End Sub
