VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C07800&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2715
   LinkTopic       =   "Form2"
   ScaleHeight     =   2415
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "made by leo/jason"
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8F0&
      Height          =   315
      Left            =   150
      TabIndex        =   5
      Tag             =   "made by leo/jason"
      Text            =   "http://www.leotown.com"
      ToolTipText     =   "Put The location of the link in to this text box"
      Top             =   1500
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Make Fade Link"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   150
      TabIndex        =   3
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click this option to turn the faded text in to a link to another webPage"
      Top             =   825
      Width           =   1965
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8F0&
      Height          =   315
      Left            =   675
      TabIndex        =   2
      Tag             =   "made by leo/jason"
      Text            =   "4"
      ToolTipText     =   "Put in the size of the faded text you want"
      Top             =   450
      Width           =   1890
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
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
      Left            =   1500
      TabIndex        =   7
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click this button to reject the changeg and exit the Options dialog"
      Top             =   1950
      Width           =   1065
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
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
      Left            =   150
      TabIndex        =   6
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click this button to accept changes"
      Top             =   1950
      Width           =   1065
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C07800&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   150
      Top             =   1950
      Width           =   1065
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C07800&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   1500
      Top             =   1950
      Width           =   1065
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Link Target:"
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
      Left            =   150
      TabIndex        =   4
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Loction of the link"
      Top             =   1125
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
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
      Left            =   150
      TabIndex        =   1
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Size of the faded text"
      Top             =   450
      Width           =   540
   End
   Begin VB.Label tt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      Left            =   75
      TabIndex        =   0
      Tag             =   "made by leo/jason"
      ToolTipText     =   "More Options Dialog Box"
      Top             =   75
      Width           =   2565
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C07800&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   75
      Top             =   75
      Width           =   2565
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C07800&
      FillStyle       =   0  'Solid
      Height          =   1965
      Left            =   75
      Tag             =   "made by leo/jason"
      Top             =   375
      Width           =   2565
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Form2.Top = (Z.Y * 15) - MouseDownFormY
Form2.Left = (Z.X * 15) - MouseDownFormX
End Sub
Private Sub Label3_Click()
Form2.Hide
End Sub
Private Sub Label4_Click()
If IsNumeric(Text1.text) Then
optiona(1) = Text1.text
Else
optiona(1) = 4
End If
optiona(2) = Check1.Value
optiona(3) = Text2.text
Form2.Hide
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape3.BackColor = MdBgColor
Label4.ForeColor = MdFgColor
End Sub
Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape3.BackColor = BgColor
Label4.ForeColor = FgColor
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape4.BackColor = MdBgColor
Label3.ForeColor = MdFgColor
End Sub
Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape4.BackColor = BgColor
Label3.ForeColor = FgColor
End Sub
