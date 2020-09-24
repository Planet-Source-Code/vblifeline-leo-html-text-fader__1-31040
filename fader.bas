Attribute VB_Name = "Module1"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public optiona(3)
Public MdFgColor
Public MdBgColor
Public TxBgColor
Public TxFgColor
Public CbBgColor
Public BgColor
Public FgColor
Public BdColor
Public ColorC
Public PTemp
Dim CurRgn, TempRgn As Long
Public Sub UnLoadMe(Frm As Form)
Unload Frm
Load Frm
Frm.Show
End Sub
Public Sub MakeNormal(lngHwnd As Long)
    SetWindowPos lngHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
Public Sub MakeTopMost(lngHwnd As Long)
    SetWindowPos lngHwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
Public Function ErrorHandler()
Beep
Form4.Show
Form4.Label3.caption = "Details>"
Form4.Height = 1605
MakeTopMost Form4.hwnd
Form4.tt.caption = "Error:" & Err.Number & " has been created"
Form4.Label5 = Err.Description
Err.Clear
End Function
Public Function MyMSG(caption, text)
Beep
Form3.Show
Form3.tt.caption = caption
Form3.Label1.caption = text
MakeTopMost Form3.hwnd
End Function
Public Function SetColor()
ColorC = "no"
Form5.Width = 4290
Form5.Show
Form5.Label1.caption = "Custom>"
Do
DoEvents
If ColorC = "yes" Then
SetColor = CbBgColor
Exit Function
End If
Loop Until ColorC <> "no"
SetColor = ColorC
End Function
Public Function HexCon(HexText)
X = LCase(Right(HexText, 1))
Y = LCase(Left(HexText, 1))
If X = "a" Then
X = 10
ElseIf X = "b" Then
X = 11
ElseIf X = "c" Then
X = 12
ElseIf X = "d" Then
X = 13
ElseIf X = "e" Then
X = 14
ElseIf X = "f" Then
X = 15
End If
If Y = "a" Then
Y = 10
ElseIf Y = "b" Then
Y = 11
ElseIf Y = "c" Then
Y = 12
ElseIf Y = "d" Then
Y = 13
ElseIf Y = "e" Then
Y = 14
ElseIf Y = "f" Then
Y = 15
End If
nTemp = (Y * 16) + X
HexCon = nTemp
End Function
Public Function changecolor(nBgColor, nFgColor, nBdColor, nMdBgColor, nMdFgColor, nTxBgColor, nTxFgColor, nCbBgColor)
nTemp = nBgColor & ":" & nFgColor & ":" & nBdColor & ":" & nMdBgColor & ":" & nMdFgColor & ":" & nTxBgColor & ":" & nTxFgColor & ":" & nCbBgColor & ":"
Set a1041 = CreateObject("WScript.Shell")
a1041.RegWrite "HKEY_LOCAL_MACHINE\software\leoinc\fader\info", nTemp
Set a57128 = CreateObject("WScript.Shell")
a57128.RegWrite "HKEY_LOCAL_MACHINE\software\leoinc\fader\there", "leo"
BgColor = nBgColor
FgColor = nFgColor
BdColor = nBdColor
MdBgColor = nMdBgColor
MdFgColor = nMdFgColor
TxBgColor = nTxBgColor
TxFgColor = nTxFgColor
CbBgColor = nCbBgColor
If nInvis = "kktrue" Then
End
End If
UpdateColors
End Function
Public Sub UpdateColors()
For X = 0 To Form6.Check1.UBound
Form6.Check1(X).BackColor = BgColor
Form6.Check1(X).ForeColor = FgColor
Next
Form5.Picture3.BackColor = CbBgColor
Form1.Text1.BackColor = TxBgColor
Form1.Text1.ForeColor = TxFgColor
Form5.Text1.BackColor = TxBgColor
Form5.Text1.ForeColor = TxFgColor
Form1.Text2.BackColor = TxBgColor
Form1.Text2.ForeColor = TxFgColor
Form5.Text2.BackColor = TxBgColor
Form5.Text2.ForeColor = TxFgColor
Form5.Text3.BackColor = TxBgColor
Form5.Text3.ForeColor = TxFgColor
Form2.Text1.BackColor = TxBgColor
Form2.Text1.ForeColor = TxFgColor
Form2.Text2.BackColor = TxBgColor
Form2.Text2.ForeColor = TxFgColor
Form1.Check1.BackColor = BgColor
Form2.Check1.BackColor = BgColor
Form1.Check1.ForeColor = FgColor
Form2.Check1.ForeColor = FgColor
Form1.Check2.BackColor = BgColor
Form1.Check2.ForeColor = FgColor
Form1.Check3.BackColor = BgColor
Form1.Check3.ForeColor = FgColor
Form1.Check4.BackColor = BgColor
Form1.Check4.ForeColor = FgColor
Form1.Label1.ForeColor = FgColor
Form2.Label1.ForeColor = FgColor
Form3.Label1.ForeColor = FgColor
Form4.Label1.ForeColor = FgColor
Form5.Label1.ForeColor = FgColor
Form1.Label2.ForeColor = FgColor
Form2.Label2.ForeColor = FgColor
Form4.Label2.ForeColor = FgColor
Form5.Label2.ForeColor = FgColor
Form1.Label3.ForeColor = FgColor
Form2.Label3.ForeColor = FgColor
Form4.Label3.ForeColor = FgColor
Form5.Label3.ForeColor = FgColor
Form1.Label4.ForeColor = FgColor
Form2.Label4.ForeColor = FgColor
Form3.Label4.ForeColor = FgColor
Form4.Label4.ForeColor = FgColor
Form5.Label4.ForeColor = FgColor
Form1.Label5.ForeColor = FgColor
Form4.Label5.ForeColor = FgColor
Form5.Label5.ForeColor = FgColor
Form1.Label6.ForeColor = FgColor
Form5.Label6.ForeColor = FgColor
Form1.Label7.ForeColor = FgColor
Form6.Label1.ForeColor = FgColor
Form6.Label2.ForeColor = FgColor
Form6.Label3.ForeColor = FgColor
Form6.Label4.ForeColor = FgColor
Form6.Label5.ForeColor = FgColor
Form6.Label6.ForeColor = FgColor
Form6.Label7.ForeColor = FgColor
Form6.Label8.ForeColor = FgColor
Form1.Shape1.FillColor = BgColor
Form2.Shape1.FillColor = BgColor
Form3.Shape1.FillColor = BgColor
Form4.Shape1.FillColor = BgColor
Form5.Shape1.FillColor = BgColor
Form6.Shape1.FillColor = BgColor
Form1.Shape2.BackColor = BgColor
Form2.Shape2.FillColor = BgColor
Form3.Shape2.FillColor = BgColor
Form4.Shape2.FillColor = BgColor
Form6.Shape2.FillColor = BgColor
Form1.Shape3.BackColor = BgColor
Form2.Shape3.BackColor = BgColor
Form3.Shape3.BackColor = BgColor
Form4.Shape3.BackColor = BgColor
Form3.Shape3.BackColor = BgColor
Form5.Shape3.BackColor = BgColor
Form1.Shape4.BackColor = BgColor
Form2.Shape4.BackColor = BgColor
Form4.Shape4.BackColor = BgColor
Form5.Shape4.BackColor = BgColor
Form1.Shape5.BackColor = BgColor
Form4.Shape5.BackColor = BgColor
Form5.Shape5.BackColor = BgColor
Form6.Shape5.BackColor = BgColor
Form4.Shape6.FillColor = BgColor
Form1.Shape6.BackColor = BgColor
Form6.Shape6.BackColor = BgColor
Form6.Shape7.FillColor = BgColor
Form6.Shape9.FillColor = BgColor
Form1.Shape1.BorderColor = BdColor
Form2.Shape1.BorderColor = BdColor
Form3.Shape1.BorderColor = BdColor
Form4.Shape1.BorderColor = BdColor
Form5.Shape1.BorderColor = BdColor
Form6.Shape1.BorderColor = BdColor
Form1.Shape2.BorderColor = BdColor
Form2.Shape2.BorderColor = BdColor
Form3.Shape2.BorderColor = BdColor
Form4.Shape2.BorderColor = BdColor
Form6.Shape2.BorderColor = BdColor
Form1.Shape3.BorderColor = BdColor
Form2.Shape3.BorderColor = BdColor
Form3.Shape3.BorderColor = BdColor
Form4.Shape3.BorderColor = BdColor
Form5.Shape3.BorderColor = BdColor
Form1.Shape4.BorderColor = BdColor
Form2.Shape4.BorderColor = BdColor
Form4.Shape4.BorderColor = BdColor
Form5.Shape4.BorderColor = BdColor
Form1.Shape5.BorderColor = BdColor
Form4.Shape5.BorderColor = BdColor
Form5.Shape5.BorderColor = BdColor
Form6.Shape5.BorderColor = BdColor
Form4.Shape6.BorderColor = BdColor
Form1.Shape6.BorderColor = BdColor
Form6.Shape6.BorderColor = BdColor
Form6.Shape7.BorderColor = BdColor
Form6.Shape9.BorderColor = BdColor
Form1.min.BorderColor = BdColor
Form1.max.BorderColor = BdColor
Form1.min.BackColor = BgColor
Form1.max.BackColor = BgColor
Form1.BackColor = BgColor
Form2.BackColor = BgColor
Form3.BackColor = BgColor
Form4.BackColor = BgColor
Form5.BackColor = BgColor
Form6.BackColor = BgColor
Form2.tt.ForeColor = FgColor
Form3.tt.ForeColor = FgColor
Form4.tt.ForeColor = FgColor
Form5.tt.ForeColor = FgColor
Form1.Picture1(1).BackColor = CbBgColor
Form1.Picture1(2).BackColor = CbBgColor
Form1.Picture1(3).BackColor = CbBgColor
Form1.Picture1(4).BackColor = CbBgColor
Form1.Picture1(5).BackColor = CbBgColor
Form1.Picture1(6).BackColor = CbBgColor
Form1.Picture1(7).BackColor = CbBgColor
Form1.Picture1(8).BackColor = CbBgColor
Form1.Picture1(9).BackColor = CbBgColor
Form1.Picture1(10).BackColor = CbBgColor
Form1.Picture1(11).BackColor = CbBgColor
End Sub
