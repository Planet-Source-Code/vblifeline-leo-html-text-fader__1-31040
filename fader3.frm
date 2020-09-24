VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00C07800&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   1650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5700
   LinkTopic       =   "Form3"
   ScaleHeight     =   1650
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "made by leo/jason"
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
      Left            =   2025
      TabIndex        =   2
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click here to exit the dialog"
      Top             =   1200
      Width           =   1515
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C07800&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   2025
      Top             =   1200
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   765
      Left            =   150
      TabIndex        =   1
      Tag             =   "made by leo/jason"
      Top             =   450
      Width           =   5415
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C07800&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   75
      Tag             =   "made by leo/jason"
      Top             =   375
      Width           =   5565
   End
   Begin VB.Label tt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
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
      Top             =   75
      Width           =   5565
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C07800&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   75
      Top             =   75
      Width           =   5565
   End
End
Attribute VB_Name = "Form3"
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
Form3.Top = (Z.Y * 15) - MouseDownFormY
Form3.Left = (Z.X * 15) - MouseDownFormX
End Sub
Private Sub Label4_Click()
If PTemp = 1 Then
End
End If
Form3.Hide
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape3.BackColor = MdBgColor
Label4.ForeColor = MdFgColor
End Sub
Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape3.BackColor = BgColor
Label4.ForeColor = FgColor
End Sub
