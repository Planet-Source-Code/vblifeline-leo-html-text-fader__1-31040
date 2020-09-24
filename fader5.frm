VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00C07800&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   3990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4290
   LinkTopic       =   "Form5"
   ScaleHeight     =   3990
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8F0&
      Height          =   315
      Left            =   6300
      TabIndex        =   56
      ToolTipText     =   "Put a number 0 to 255 in here to define the amount of blue you want"
      Top             =   3525
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8F0&
      Height          =   315
      Left            =   6300
      TabIndex        =   55
      ToolTipText     =   "Put a number 0 to 255 in here to define the amount of green you want"
      Top             =   3225
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8F0&
      Height          =   315
      Left            =   6300
      TabIndex        =   54
      ToolTipText     =   "Put a number 0 to 255 in here to define the amount of red you want"
      Top             =   2925
      Width           =   615
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E8F0&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   4275
      ScaleHeight     =   885
      ScaleWidth      =   960
      TabIndex        =   53
      ToolTipText     =   "The color or cutom color that you define appers here"
      Top             =   2925
      Width           =   990
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   47
      Left            =   3825
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   48
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   3150
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   46
      Left            =   3300
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   47
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   3150
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   45
      Left            =   2775
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   46
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   3150
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   44
      Left            =   2250
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   45
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   3150
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004040&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   43
      Left            =   1725
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   44
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   3150
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   42
      Left            =   1200
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   43
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   3150
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   41
      Left            =   675
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   42
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   3150
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   40
      Left            =   150
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   41
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   3150
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800080&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   39
      Left            =   3825
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   40
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   2625
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   38
      Left            =   3300
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   39
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   2625
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   37
      Left            =   2775
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   38
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   2625
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   36
      Left            =   2250
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   37
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   2625
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   35
      Left            =   1725
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   36
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   2625
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   34
      Left            =   1200
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   35
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   2625
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   33
      Left            =   675
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   34
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   2625
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   32
      Left            =   150
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   33
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   2625
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   31
      Left            =   3825
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   32
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   2100
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   30
      Left            =   3300
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   31
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   2100
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   29
      Left            =   2775
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   30
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   2100
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   28
      Left            =   2250
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   29
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   2100
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   27
      Left            =   1725
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   28
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   2100
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   26
      Left            =   1200
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   27
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   2100
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   25
      Left            =   675
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   26
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   2100
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   24
      Left            =   150
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   25
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   2100
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   23
      Left            =   3825
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   24
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   1575
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   22
      Left            =   3300
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   23
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   1575
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   21
      Left            =   2775
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   22
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   1575
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   20
      Left            =   2250
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   21
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   1575
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   19
      Left            =   1725
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   20
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   1575
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   18
      Left            =   1200
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   19
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   1575
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   17
      Left            =   675
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   18
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   1575
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   16
      Left            =   150
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   17
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   1575
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   15
      Left            =   3825
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   16
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   1050
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   14
      Left            =   3300
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   15
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   1050
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   13
      Left            =   2775
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   14
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   1050
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   12
      Left            =   2250
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   13
      Tag             =   "made by leo/jasonmade by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   1050
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   11
      Left            =   1725
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   12
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   1050
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   10
      Left            =   1200
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   11
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   1050
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   9
      Left            =   675
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   10
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   1050
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   8
      Left            =   150
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   9
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   1050
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   7
      Left            =   3825
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   8
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   525
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   6
      Left            =   3300
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   7
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   525
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   5
      Left            =   2775
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   6
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   525
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   4
      Left            =   2250
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   5
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   525
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   3
      Left            =   1725
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   4
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   525
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   1200
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   3
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   525
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   675
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   2
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   525
      Width           =   315
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   150
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   1
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click Here To Select This Color"
      Top             =   525
      Width           =   315
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2790
      Left            =   4275
      Picture         =   "fader5.frx":0000
      ScaleHeight     =   2760
      ScaleWidth      =   2610
      TabIndex        =   52
      ToolTipText     =   "Click somewhere on the picture to difine your cutom color"
      Top             =   75
      Width           =   2640
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Blue"
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
      Left            =   5550
      TabIndex        =   59
      Top             =   3525
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Green"
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
      Left            =   5550
      TabIndex        =   58
      Top             =   3225
      Width           =   690
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Red"
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
      Left            =   5550
      TabIndex        =   57
      Top             =   2925
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Custom>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2925
      TabIndex        =   51
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click here to define your own cutom colors"
      Top             =   3600
      Width           =   1290
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C07800&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   2925
      Top             =   3600
      Width           =   1290
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
      Height          =   315
      Left            =   1500
      TabIndex        =   49
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Never mind you dont want a color just hit this button"
      Top             =   3600
      Width           =   1365
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
      Height          =   315
      Left            =   75
      TabIndex        =   50
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click here to acept the color you have defined"
      Top             =   3600
      Width           =   1365
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C07800&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   1500
      Top             =   3600
      Width           =   1365
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C07800&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   75
      Top             =   3600
      Width           =   1365
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   47
      Left            =   3750
      Top             =   3075
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   46
      Left            =   3225
      Top             =   3075
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   45
      Left            =   2700
      Top             =   3075
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   44
      Left            =   2175
      Top             =   3075
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   43
      Left            =   1650
      Top             =   3075
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   42
      Left            =   1125
      Top             =   3075
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   41
      Left            =   600
      Top             =   3075
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   40
      Left            =   75
      Top             =   3075
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   39
      Left            =   3750
      Top             =   2550
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   38
      Left            =   3225
      Top             =   2550
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   37
      Left            =   2700
      Top             =   2550
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   36
      Left            =   2175
      Top             =   2550
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   35
      Left            =   1650
      Top             =   2550
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   34
      Left            =   1125
      Top             =   2550
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   33
      Left            =   600
      Top             =   2550
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   32
      Left            =   75
      Top             =   2550
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   31
      Left            =   3750
      Top             =   2025
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   30
      Left            =   3225
      Top             =   2025
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   29
      Left            =   2700
      Top             =   2025
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   28
      Left            =   2175
      Top             =   2025
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   27
      Left            =   1650
      Top             =   2025
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   26
      Left            =   1125
      Top             =   2025
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   25
      Left            =   600
      Top             =   2025
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   24
      Left            =   75
      Top             =   2025
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   23
      Left            =   3750
      Top             =   1500
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   22
      Left            =   3225
      Top             =   1500
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   21
      Left            =   2700
      Top             =   1500
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   20
      Left            =   2175
      Top             =   1500
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   19
      Left            =   1650
      Top             =   1500
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   18
      Left            =   1125
      Top             =   1500
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   17
      Left            =   600
      Top             =   1500
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   16
      Left            =   75
      Top             =   1500
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   15
      Left            =   3750
      Top             =   975
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   14
      Left            =   3225
      Top             =   975
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   13
      Left            =   2700
      Top             =   975
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   12
      Left            =   2175
      Top             =   975
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   11
      Left            =   1650
      Top             =   975
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   10
      Left            =   1125
      Top             =   975
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   9
      Left            =   600
      Top             =   975
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   8
      Left            =   75
      Top             =   975
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   7
      Left            =   3750
      Top             =   450
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   6
      Left            =   3225
      Top             =   450
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   5
      Left            =   2700
      Top             =   450
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   4
      Left            =   2175
      Top             =   450
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   3
      Left            =   1650
      Top             =   450
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   2
      Left            =   1125
      Top             =   450
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   1
      Left            =   600
      Top             =   450
      Width           =   465
   End
   Begin VB.Shape Shape2 
      Height          =   465
      Index           =   0
      Left            =   75
      Top             =   450
      Width           =   465
   End
   Begin VB.Label tt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Colors"
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
      TabIndex        =   0
      Tag             =   "made by leo/jason"
      ToolTipText     =   "This Is the dialog Where you define your own colors"
      Top             =   75
      Width           =   4140
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C07800&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   75
      Top             =   75
      Width           =   4140
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Last
Private TempColor
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
Form5.Top = (Z.Y * 15) - MouseDownFormY
Form5.Left = (Z.X * 15) - MouseDownFormX
End Sub
Private Sub Form_Load()
MakeTopMost Form5.hwnd
TempColor = "14215408"
Last = "n/a"
For X = 0 To 47
Shape2(X).Visible = False
Shape2(X).BorderColor = BdColor
Next
End Sub
Private Sub mapcolor()
On Error GoTo ErrorHandler
DoEvents
nTemp = Hex(Picture3.BackColor)
If Len(nTemp) = 4 Then
nTemp = "00" & nTemp
End If
If Len(nTemp) = 5 Then
nTemp = "0" & nTemp
End If
If Len(nTemp) = 2 Then
nTemp = "0000" & nTemp
End If
Text3.text = HexCon(Left(nTemp, 2))
Text2.text = HexCon(Right(Left(nTemp, 4), 2))
Text1.text = HexCon(Right(nTemp, 2))
Exit Sub
ErrorHandler:
ErrorHandler
End Sub
Private Sub Label1_Click()
If Label1.caption = "Custom>" Then
Label1.caption = "<Custom"
Form5.Width = 6990
Else
Form5.Width = 4290
Label1.caption = "Custom>"
End If
End Sub
Private Sub Label3_Click()
ColorC = "yes"
Me.Hide
End Sub
Private Sub Label4_Click()
ColorC = Picture3.BackColor
Me.Hide
End Sub
Private Sub Picture1_Click(Index As Integer)
On Error GoTo ErrorHandler
If Last = "n/a" Then
Last = Index
Shape2(Index).Visible = True
TempColor = Picture1(Index).BackColor
Picture3.BackColor = Picture1(Index).BackColor
mapcolor
Exit Sub
End If
Shape2(Last).Visible = False
Last = Index
Shape2(Index).Visible = True
TempColor = Picture1(Index).BackColor
Picture3.BackColor = Picture1(Index).BackColor
mapcolor
Exit Sub
ErrorHandler:
ErrorHandler
End Sub
Private Sub Picture1_DblClick(Index As Integer)
Label4_Click
End Sub
Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
DoEvents
Picture3.BackColor = Picture2.Point(X, Y)
mapcolor
Exit Sub
ErrorHandler:
ErrorHandler
End Sub
Private Sub Text1_Change()
On Error GoTo ErrorHandler
If Text1.text <> "" And Text2.text <> "" And Text3.text <> "" Then
Picture3.BackColor = "&H" & HexCheck(Text3.text) & HexCheck(Text2.text) & HexCheck(Text1.text)
End If
Exit Sub
ErrorHandler:
ErrorHandler
End Sub
Private Sub Text2_Change()
On Error GoTo ErrorHandler
If Text1.text <> "" And Text2.text <> "" And Text3.text <> "" Then
Picture3.BackColor = "&H" & HexCheck(Text3.text) & HexCheck(Text2.text) & HexCheck(Text1.text)
End If
Exit Sub
ErrorHandler:
ErrorHandler
End Sub
Private Sub Text3_Change()
On Error GoTo ErrorHandler
If Text1.text <> "" And Text2.text <> "" And Text3.text <> "" Then
Picture3.BackColor = "&H" & HexCheck(Text3.text) & HexCheck(Text2.text) & HexCheck(Text1.text)
End If
Exit Sub
ErrorHandler:
ErrorHandler
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
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape5.BackColor = MdBgColor
Label1.ForeColor = MdFgColor
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape5.BackColor = BgColor
Label1.ForeColor = FgColor
End Sub
Private Function HexCheck(ntext)
X = Hex(ntext)
If Len(X) = 1 Then
X = "0" & X
End If
HexCheck = X
End Function
