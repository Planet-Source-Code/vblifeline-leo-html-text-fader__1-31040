VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C07800&
   BorderStyle     =   0  'None
   ClientHeight    =   6705
   ClientLeft      =   2385
   ClientTop       =   1185
   ClientWidth     =   5085
   Icon            =   "fader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   447
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   339
   StartUpPosition =   2  'CenterScreen
   Tag             =   "made by leo/jason"
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8F0&
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   11
      Left            =   4575
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   24
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Use These To Fade Your Text From Color To Color"
      Top             =   1275
      Width           =   420
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8F0&
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   10
      Left            =   4125
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   23
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Use These To Fade Your Text From Color To Color"
      Top             =   1275
      Width           =   420
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8F0&
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   9
      Left            =   3675
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   22
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Use These To Fade Your Text From Color To Color"
      Top             =   1275
      Width           =   420
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8F0&
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   8
      Left            =   3225
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   21
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Use These To Fade Your Text From Color To Color"
      Top             =   1275
      Width           =   420
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8F0&
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   7
      Left            =   2775
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   20
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Use These To Fade Your Text From Color To Color"
      Top             =   1275
      Width           =   420
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8F0&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   9825
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   17
      Top             =   6600
      Width           =   135
   End
   Begin VB.CheckBox Check4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Italic"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1575
      TabIndex        =   15
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click this option to make the faded text Italic"
      Top             =   900
      Width           =   690
   End
   Begin VB.CheckBox Check3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "UnderLine"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2325
      TabIndex        =   14
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click this option to make the faded text Underlined"
      Top             =   900
      Width           =   1065
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Bold"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   825
      TabIndex        =   13
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click this option to make the faded text bold"
      Top             =   900
      Width           =   765
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Wavy"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   75
      TabIndex        =   12
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click This Option To Make the Text Come Out Faded"
      Top             =   900
      Width           =   765
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3600
      Tag             =   "made by leo/jason"
      Top             =   3075
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8F0&
      ForeColor       =   &H00000000&
      Height          =   3990
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Tag             =   "made by leo/jason"
      Text            =   "fader.frx":08CA
      ToolTipText     =   "After you fade he text the code will apere in this box"
      Top             =   1800
      Width           =   4890
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8F0&
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   6
      Left            =   2325
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   6
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Use These To Fade Your Text From Color To Color"
      Top             =   1275
      Width           =   420
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8F0&
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   5
      Left            =   1875
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   5
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Use These To Fade Your Text From Color To Color"
      Top             =   1275
      Width           =   420
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8F0&
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   4
      Left            =   1425
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   4
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Use These To Fade Your Text From Color To Color"
      Top             =   1275
      Width           =   420
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8F0&
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   3
      Left            =   975
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   3
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Use These To Fade Your Text From Color To Color"
      Top             =   1275
      Width           =   420
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8F0&
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   2
      Left            =   525
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   2
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Use These To Fade Your Text From Color To Color"
      Top             =   1275
      Width           =   420
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8F0&
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   75
      TabIndex        =   1
      Tag             =   "made by leo/jason"
      Text            =   "Text To Fade"
      ToolTipText     =   "Use this Box To Define thew text that you want faded"
      Top             =   450
      Width           =   4890
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8F0&
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   1
      Left            =   75
      ScaleHeight     =   375
      ScaleWidth      =   390
      TabIndex        =   0
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Use These To Fade Your Text From Color To Color"
      Top             =   1275
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   75
      Tag             =   "Program Made By Jason. Can Contact Me At Isurftheweb2001@msn.com"
      Top             =   4575
      Width           =   4890
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Color Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2175
      TabIndex        =   25
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click this button To Bring up the color control panel"
      Top             =   6300
      Width           =   2790
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00C07800&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   2175
      Top             =   6300
      Width           =   2790
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copy Code To ClipBoard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   75
      TabIndex        =   19
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click this button to copy the code into the clipboard"
      Top             =   5925
      Width           =   2865
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "More Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3450
      TabIndex        =   18
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Bring Up Adtional Option"
      Top             =   900
      Width           =   1515
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C07800&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   3450
      Top             =   900
      Width           =   1515
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New Fade Job"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   75
      TabIndex        =   16
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click this button to start all over againg"
      Top             =   6300
      Width           =   2040
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C07800&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   75
      Top             =   6300
      Width           =   2040
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4650
      TabIndex        =   11
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click this button to exit the program"
      Top             =   75
      Width           =   315
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4350
      TabIndex        =   10
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click here to minimize the program"
      Top             =   75
      Width           =   315
   End
   Begin VB.Shape max 
      BackColor       =   &H00C07800&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   4650
      Top             =   75
      Width           =   315
   End
   Begin VB.Shape min 
      BackColor       =   &H00C07800&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   4350
      Top             =   75
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3075
      TabIndex        =   8
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click here To Generate the faded text code"
      Top             =   5925
      Width           =   1890
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C07800&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   3075
      Top             =   5925
      Width           =   1890
   End
   Begin VB.Label tt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Leo Text Fader..... LeoTown.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   75
      TabIndex        =   9
      Tag             =   "made by leo/jason"
      ToolTipText     =   "This Program was Made By leo inc.  Visit Leotown.com For more programs and source code"
      Top             =   75
      Width           =   4290
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C07800&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   75
      Top             =   75
      Width           =   4890
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C07800&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   75
      Top             =   5925
      Width           =   2940
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private red(1000)
Private green(1000)
Private blue(1000)
Private FadeTT(43)
Private color(55)
Private lenth
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private MouseDownForm
Private MouseDownFormX
Private MouseDownFormY
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Sub tt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseDownForm = 1
MouseDownFormX = X
MouseDownFormY = Y
End Sub
Private Sub tt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseDownForm = 0
End Sub
Private Sub tt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoEvents
If MouseDownForm <> 1 Then
Exit Sub
End If
Dim Z As POINTAPI
Call GetCursorPos(Z)
Form1.Top = (Z.Y * 15) - MouseDownFormY
Form1.Left = (Z.X * 15) - MouseDownFormX
End Sub
Private Sub Form_Load()
On Error GoTo ErrorHandler
ReloadInterface
Text2.text = "Code Will Appere Here"
MakeTopMost Form1.hwnd
optiona(1) = 4
optiona(2) = 0
optiona(3) = "http://www.leotown.com"
Form1.Width = 5070
FadeTT(0) = 0
FadeTT(1) = "&HFF0000"
FadeTT(2) = "&HFF1212"
FadeTT(3) = "&HFF2424"
FadeTT(4) = "&HFF3636"
FadeTT(5) = "&HFF4848"
FadeTT(6) = "&HFF5A5A"
FadeTT(7) = "&HFF6C6C"
FadeTT(8) = "&HFF7E7E"
FadeTT(9) = "&HFF9090"
FadeTT(10) = "&HFFA2A2"
FadeTT(11) = "&HFFB4B4"
FadeTT(12) = "&HFFC6C6"
FadeTT(13) = "&HFFD8D8"
FadeTT(14) = "&HFFEAEA"
FadeTT(15) = "&HFFFCFC"
FadeTT(16) = "&HEDFFFF"
FadeTT(17) = "&HDBEDFF"
FadeTT(18) = "&HC9DBFF"
FadeTT(19) = "&HB7C9FF"
FadeTT(20) = "&HA5B7FF"
FadeTT(21) = "&H93A5FF"
FadeTT(22) = "&H8193FF"
FadeTT(23) = "&H6F81FF"
FadeTT(24) = "&H5D6FFF"
FadeTT(25) = "&H4B5DFF"
FadeTT(26) = "&H394BFF"
FadeTT(27) = "&H2739FF"
FadeTT(28) = "&H1527FF"
FadeTT(29) = "&H0315FF"
FadeTT(30) = "&H0003FF"
FadeTT(31) = "&H1200ED"
FadeTT(32) = "&H2400DB"
FadeTT(33) = "&H3600C9"
FadeTT(34) = "&H4800B7"
FadeTT(35) = "&H5A00A5"
FadeTT(36) = "&H6C0093"
FadeTT(37) = "&H7E0081"
FadeTT(38) = "&H90006F"
FadeTT(39) = "&HA2005D"
FadeTT(40) = "&HB4004B"
FadeTT(41) = "&HC60039"
FadeTT(42) = "&HD80027"
FadeTT(43) = "&HEA0015"
For X = 1 To Picture1.UBound
Picture1(X).BackColor = CbBgColor
Picture1(X).Tag = "a"
Next
Me.Show
Form3.Hide
Exit Sub
ErrorHandler:
ErrorHandler
End Sub
Private Sub Label1_Click()
fadetext
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape2.BackColor = MdBgColor
Label1.ForeColor = MdFgColor
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape2.BackColor = BgColor
Label1.ForeColor = FgColor
End Sub
Private Sub Label2_Click()
Form1.WindowState = 1
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
min.BackColor = MdBgColor
Label2.ForeColor = MdFgColor
End Sub
Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
min.BackColor = BgColor
Label2.ForeColor = FgColor
End Sub
Private Sub Label3_Click()
End
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
max.BackColor = MdBgColor
Label3.ForeColor = MdFgColor
End Sub
Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
max.BackColor = BgColor
Label3.ForeColor = FgColor
End Sub
Private Sub Label4_Click()
For X = 1 To Picture1.UBound
Picture1(X).BackColor = CbBgColor
Picture1(X).Tag = "n"
Next
Text1.text = "Text To Fade"
Text2.text = "Code Will Appere Here"
For X = 1 To 4
Me("Check" & X).Value = 0
Next
optiona(1) = 4
optiona(2) = 0
optiona(3) = "http://www.leotown.com"
Form2.Text1.text = "4"
Form2.Text2.text = "http://www.leotown.com"
Form2.Check1.Value = 0
Form5.Text1.text = ""
Form5.Text2.text = ""
Form5.Text3.text = ""
Form5.Picture3.BackColor = 14215408
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape3.BackColor = MdBgColor
Label4.ForeColor = MdFgColor
End Sub
Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape3.BackColor = BgColor
Label4.ForeColor = FgColor
End Sub
Private Sub Label5_Click()
Form2.Show
MakeTopMost Form2.hwnd
End Sub
Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape4.BackColor = MdBgColor
Label5.ForeColor = MdFgColor
End Sub
Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape4.BackColor = BgColor
Label5.ForeColor = FgColor
End Sub
Private Sub Label6_Click()
Text2.SetFocus
Text2.SelStart = 0
Text2.SelLength = Len(Text2.text)

SendKeys "^c"
End Sub
Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape5.BackColor = MdBgColor
Label6.ForeColor = MdFgColor
End Sub
Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape5.BackColor = BgColor
Label6.ForeColor = FgColor
End Sub
Private Sub Label7_Click()
Form6.Show
End Sub
Private Sub Picture1_Click(Index As Integer)
On Error Resume Next
Picture1(Index).BackColor = SetColor
If ColorC <> "yes" Then
Picture1(Index).Tag = "t"
End If
Exit Sub
End Sub
Private Sub Timer1_Timer()
DoEvents
FadeTT(0) = FadeTT(0) + 1
If FadeTT(0) = 44 Then
FadeTT(0) = 1
End If
tt.ForeColor = FadeTT(FadeTT(0))
Form1.caption = FadeTT(FadeTT(0))
App.Title = FadeTT(FadeTT(0))
End Sub
Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.text)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Text1.text = "302836512332412" Then
MyMSG "Info", Image1.Tag
End If
fadetext
End If
End Sub
Private Function fadetext()
On Error GoTo ErrorHandler
Dim nTemp
red(0) = 0
green(0) = 0
blue(0) = 0
color(0) = Picture1.UBound
nTemp = 0
For X = 1 To Picture1.UBound
If Picture1(X).Tag = "a" Then
color(0) = color(0) - 1
Else
nTemp = nTemp + 1
color(nTemp) = Hex(Picture1(X).BackColor)
If Len(color(nTemp)) = 4 Then
color(nTemp) = "00" & color(nTemp)
End If
If Len(color(nTemp)) = 5 Then
color(nTemp) = "0" & color(nTemp)
End If
If Len(color(nTemp)) = 2 Then
color(nTemp) = "0000" & color(nTemp)
End If
End If
If color(0) = 1 Then
MyMSG "Two Colors", "Must Have At Lest To Colors"
Exit Function
End If
Next
Text2.text = ""
lenth = Len(Text1.text) / (color(0) - 1)
lenth = Round(lenth, 0)
For ntemocolor = 1 To color(0) - 1
xc = Left(color(ntemocolor), 2)
xb = Right(Left(color(ntemocolor), 4), 2)
xa = Right(color(ntemocolor), 2)
yc = Left(color(ntemocolor + 1), 2)
yb = Right(Left(color(ntemocolor + 1), 4), 2)
ya = Right(color(ntemocolor + 1), 2)
ncount = red(0)
nstep = Round((HexCon(ya) - HexCon(xa)) / lenth, 0)
If nstep = 0 Then
ncount = ncount + lenth
For X = red(0) + 1 To ncount
red(X) = ya
Next
GoTo reda
End If
For nTemp = HexCon(xa) To HexCon(ya) Step nstep
ncount = ncount + 1
DoEvents
ntempb = Hex(Round(nTemp, 0))
If Len(ntempb) = 1 Then
ntempb = "0" & ntempb
End If
red(ncount) = ntempb
Next
reda:
red(0) = ncount
ncount = green(0)
nstep = Round((HexCon(yb) - HexCon(xb)) / lenth, 0)
If nstep = 0 Then
ncount = ncount + lenth
For X = green(0) + 1 To ncount
green(X) = yb
Next
GoTo greena
End If
For nTemp = HexCon(xb) To HexCon(yb) Step nstep
ncount = ncount + 1
DoEvents
ntempb = Hex(Round(nTemp, 0))
If Len(ntempb) = 1 Then
ntempb = "0" & ntempb
End If
green(ncount) = ntempb
Next
greena:
green(0) = ncount
ncount = blue(0)
nstep = Round((HexCon(yc) - HexCon(xc)) / lenth, 0)
If nstep = 0 Then
ncount = ncount + lenth
For X = blue(0) + 1 To ncount
blue(X) = yc
Next
GoTo bluea
End If
For nTemp = HexCon(xc) To HexCon(yc) Step nstep
ncount = ncount + 1
DoEvents
ntempb = Hex(Round(nTemp, 0))
If Len(ntempb) = 1 Then
ntempb = "0" & ntempb
End If
blue(ncount) = ntempb
Next
bluea:
blue(0) = ncount
Next
v = "Jason Is the coolest person in the world o ya i know he is he is the coolest"
For X = red(0) + 1 To 1000
red(X) = red(red(0))
Next
For X = green(0) + 1 To 1000
green(X) = green(green(0))
Next
For X = blue(0) + 1 To 1000
blue(X) = blue(blue(0))
Next
If Check1.Value = 1 Then
ntempb = 0
For X = 1 To Len(Text1.text)
ntempb = ntempb + 1
If ntempb = 5 Then
ntempb = 1
End If
If ntempb = 1 Then
ntempc = "<sup>"
End If
If ntempb = 2 Then
ntempc = "</sup>"
End If
If ntempb = 3 Then
ntempc = "<sub>"
End If
If ntempb = 4 Then
ntempc = "</sub>"
End If
If X = 1 Then
ntempgg = " size=" & optiona(1)
Else
ntempgg = ""
End If
nTemp = Right(Left(Text1.text, X), 1)
Text2.text = Text2.text & "<font color=" & red(X) & green(X) & blue(X) & ntempgg & ">" & ntempc & nTemp
Next
Else
For X = 1 To Len(Text1.text)
If X = 1 Then
ntempgg = " size=" & optiona(1)
Else
ntempgg = ""
End If
nTemp = Right(Left(Text1.text, X), 1)
Text2.text = Text2.text & "<font color=" & red(X) & green(X) & blue(X) & ntempgg & ">" & nTemp
Next
End If
If Check3.Value = 1 Then
Text2.text = "<u>" & Text2.text & "</u>"
End If
If Check4.Value = 1 Then
Text2.text = "<i>" & Text2.text & "</i>"
End If
If Check2.Value = 1 Then
Text2.text = "<b>" & Text2.text & "</b>"
End If
If optiona(2) = 1 Then
Text2.text = "<a href=" & Chr(34) & optiona(3) & Chr(34) & ">" & Text2.text & "</a>"
End If
Exit Function
ErrorHandler:
ErrorHandler
End Function
Public Sub Delay(HowLong As Date)
TempTime = DateAdd("s", HowLong, Now)
While TempTime > Now
DoEvents
Wend
End Sub
Public Sub ReloadInterface()
On Error Resume Next
Dim ntempb
ntempb = ""
Set a54339 = CreateObject("WScript.Shell")
X = a54339.RegRead("HKEY_LOCAL_MACHINE\software\leoinc\fader\there")
If X = "leo" Then
Set a11313 = CreateObject("WScript.Shell")
nTemp = a11313.RegRead("HKEY_LOCAL_MACHINE\software\leoinc\fader\info")
varray = Split(nTemp, ":")
changecolor varray(0), varray(1), varray(2), varray(3), varray(4), varray(5), varray(6), varray(7)
Else
changecolor 12613632, 16777215, 0, 16777215, 0, 16777215, 0, 16777215
End If
Exit Sub
End Sub
Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape6.BackColor = MdBgColor
Label7.ForeColor = MdFgColor
End Sub
Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape6.BackColor = BgColor
Label7.ForeColor = FgColor
End Sub
