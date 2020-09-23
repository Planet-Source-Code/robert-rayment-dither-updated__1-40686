VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   -510
   ClientWidth     =   11730
   DrawWidth       =   2
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   497
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   782
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVBASM 
      BackColor       =   &H00FFFFFF&
      Caption         =   "VB"
      Height          =   300
      Left            =   10275
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Toggle VB <-> ASM"
      Top             =   75
      Width           =   585
   End
   Begin VB.PictureBox picLims 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H000040C0&
      Height          =   270
      Left            =   6300
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   346
      TabIndex        =   40
      Top             =   405
      Width           =   5250
   End
   Begin VB.CommandButton cmdExit 
      Height          =   345
      Index           =   2
      Left            =   75
      Picture         =   "Main.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "EXIT"
      Top             =   75
      Width           =   390
   End
   Begin VB.CommandButton cmdPatterns 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Patterns"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4095
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   75
      Width           =   810
   End
   Begin VB.CommandButton cmdFile 
      BackColor       =   &H00FFFFFF&
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   495
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   75
      Width           =   495
   End
   Begin VB.CommandButton cmdExit 
      Height          =   345
      Index           =   1
      Left            =   10935
      Picture         =   "Main.frx":0584
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "EXIT"
      Top             =   75
      Width           =   390
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008000&
      Height          =   2235
      Left            =   510
      TabIndex        =   28
      Top             =   225
      Width           =   1320
      Begin VB.CommandButton cmdSave32 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Save 32bpp"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   1305
         Width           =   1170
      End
      Begin VB.CommandButton cmdExit 
         Height          =   345
         Index           =   0
         Left            =   450
         Picture         =   "Main.frx":0B08
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "EXIT"
         Top             =   1725
         Width           =   390
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print picture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   90
         TabIndex        =   31
         Top             =   915
         Width           =   1170
      End
      Begin VB.CommandButton cmdSavePic 
         Caption         =   "Save picture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   30
         Top             =   555
         Width           =   1170
      End
      Begin VB.CommandButton cmdLoadPic 
         Caption         =   "Load picture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   105
         TabIndex        =   29
         Top             =   195
         Width           =   1170
      End
   End
   Begin VB.CheckBox chkResizer 
      BackColor       =   &H000000C0&
      Caption         =   "Resize"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4935
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "After screen res change"
      Top             =   75
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Height          =   6510
      Left            =   4065
      TabIndex        =   14
      Top             =   255
      Width           =   1575
      Begin VB.CommandButton cmdContour 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Contour"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   4320
         Width           =   945
      End
      Begin VB.PictureBox picQBColor 
         Height          =   1020
         Left            =   450
         Picture         =   "Main.frx":108C
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   40
         TabIndex        =   48
         Top             =   5310
         Width           =   660
      End
      Begin VB.CommandButton cmdInvert 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Invert"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4770
         Width           =   945
      End
      Begin VB.CommandButton cmdLine2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Line 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3705
         Width           =   945
      End
      Begin VB.CommandButton cmdHalfTone2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "HalfTone 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   3105
         Width           =   945
      End
      Begin VB.CommandButton cmdRandom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Random"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4005
         Width           =   945
      End
      Begin VB.CommandButton cmdLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Line"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3405
         Width           =   945
      End
      Begin VB.CommandButton cmdHalfTone 
         BackColor       =   &H00FFFFFF&
         Caption         =   "HalfTone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2805
         Width           =   945
      End
      Begin VB.CommandButton cmdJarvisEtAl 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Jarvis et al"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   375
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2070
         Width           =   885
      End
      Begin VB.CommandButton cmdSierra 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sierra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   510
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1770
         Width           =   600
      End
      Begin VB.CommandButton cmdBurkes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Burkes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   510
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1470
         Width           =   600
      End
      Begin VB.CommandButton cmdStucki 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Stucki"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   510
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1170
         Width           =   600
      End
      Begin VB.CommandButton cmdFloydSteinbergB 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Floyd-SteinbergB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   870
         Width           =   1410
      End
      Begin VB.CommandButton cmdFloydSteinbergA 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Floyd-SteinbergA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   570
         Width           =   1410
      End
      Begin VB.CommandButton cmdBlackWhite 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Black && White"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   270
         Width           =   1200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   30
         X2              =   1545
         Y1              =   5175
         Y2              =   5175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   45
         X2              =   1560
         Y1              =   4650
         Y2              =   4650
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   45
         X2              =   1560
         Y1              =   2715
         Y2              =   2715
      End
   End
   Begin VB.Timer Timer1 
      Left            =   6285
      Top             =   6630
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00008000&
      Height          =   5625
      Left            =   6300
      ScaleHeight     =   371
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   8
      Top             =   705
      Width           =   5235
      Begin VB.PictureBox pic2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   705
         Left            =   120
         ScaleHeight     =   47
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   9
         Top             =   60
         Width           =   960
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008000&
      Height          =   5685
      Left            =   555
      ScaleHeight     =   375
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   336
      TabIndex        =   6
      Top             =   750
      Width           =   5100
      Begin VB.Frame Frame3 
         BackColor       =   &H00008000&
         Height          =   1410
         Left            =   1170
         TabIndex        =   43
         Top             =   255
         Width           =   705
         Begin VB.CommandButton cmdPrint 
            Caption         =   "x4"
            Height          =   270
            Index           =   4
            Left            =   90
            TabIndex        =   47
            Top             =   1065
            Width           =   555
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "x3"
            Height          =   270
            Index           =   3
            Left            =   90
            TabIndex        =   46
            Top             =   780
            Width           =   555
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "x2"
            Height          =   270
            Index           =   2
            Left            =   90
            TabIndex        =   45
            Top             =   495
            Width           =   555
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "x1"
            Height          =   270
            Index           =   1
            Left            =   90
            TabIndex        =   44
            Top             =   210
            Width           =   555
         End
      End
      Begin VB.PictureBox pic1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3840
         Left            =   150
         Picture         =   "Main.frx":2ECE
         ScaleHeight     =   256
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   256
         TabIndex        =   7
         Top             =   150
         Width           =   3840
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         Caption         =   "Dither  by  Robert Rayment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   345
         TabIndex        =   26
         Top             =   4980
         Width           =   2580
      End
   End
   Begin VB.Frame fraScr 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   540
      TabIndex        =   1
      Top             =   6465
      Width           =   3030
      Begin VB.CommandButton cmdScr 
         BackColor       =   &H0000C000&
         Height          =   240
         Index           =   3
         Left            =   2745
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   180
         Width           =   255
      End
      Begin VB.CommandButton cmdScr 
         BackColor       =   &H0000C000&
         Height          =   240
         Index           =   2
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   180
         Width           =   255
      End
      Begin VB.Label LabThumb 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   360
         TabIndex        =   13
         Top             =   180
         Width           =   210
      End
      Begin VB.Label LabBack 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   330
         TabIndex        =   12
         Top             =   165
         Width           =   2400
      End
   End
   Begin VB.Frame fraScr 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   3150
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   630
      Width           =   375
      Begin VB.CommandButton cmdScr 
         BackColor       =   &H0000C000&
         Height          =   250
         Index           =   1
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2865
         Width           =   270
      End
      Begin VB.CommandButton cmdScr 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   0
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   195
         Width           =   270
      End
      Begin VB.Label LabThumb 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   1020
         Width           =   210
      End
      Begin VB.Label LabBack 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   2400
         Index           =   0
         Left            =   75
         TabIndex        =   2
         Top             =   450
         Width           =   255
      End
   End
   Begin VB.Label LabFileName 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   1035
      TabIndex        =   36
      Top             =   75
      Width           =   3060
   End
   Begin VB.Label LabPatternName 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   6315
      TabIndex        =   35
      Top             =   75
      Width           =   4020
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   7170
      Left            =   -8835
      Top             =   -6690
      Width           =   11730
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Dither by  Robert Rayment

' Updated 26/11/02 Using a Jump table for ASN (see Dither.asm)
'                  & Scroll bars improved

' See Notes.txt for info on dithering.

Option Base 1

DefLng A-W
DefSng X-Z

'Public pic1mem() As Byte   ' For Loaded picture
'Public pic2mem() As Byte   ' For Output dither picture

Dim picsave() As Byte   ' For saving 2-color picture

Dim zHSc, zVSc    ' Resizing factors
Dim FW, FH        ' Reduced Form width & height

' Scroll values
Dim VVal '= nTVal(0)
Dim HVal '= nTVal(1)

' For using OSDialog(FileDlg2.cls)
Dim CommonDialog1 As New OSDialog

Dim Pathspec$, CurrPath$


Private Sub cmdContour_Click()
Frame2.Visible = False

LabPatternName = "Black and White"
Screen.MousePointer = vbHourglass

InitBits
'picLims.Visible = True
'cmdPrint(0).Enabled = True
'cmdSavePic.Enabled = True
'bm.Colors(1).rgbBlue = 255
'bm.Colors(1).rgbGreen = 255
'bm.Colors(1).rgbRed = 255

ReDim pic2mem(4, picWd, picHt)
DoEvents

'If ASMVB Then
'   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   'Set at picture load
'   DStruc.picHt = picHt
'   DStruc.picWd = picWd
'   DStruc.Ptrpic1mem = VarPtr(pic1mem(1, 1, 1))
'   DStruc.Ptrpic2mem = VarPtr(pic2mem(1, 1, 1))
'
'   DStruc.BlackThreshHold = BlackThreshHold
'   ASM_BlackWhite
'   ShowDitheredPicture
'   Screen.MousePointer = vbDefault
'   Exit Sub
'   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'End If

' VB VB VB
ReDim pic2mem(4, picWd, picHt)
For j = 2 To picHt - 1
For i = 2 To picWd - 1
   
   LBlue = (1& * pic1mem(1, i - 1, j - 1) + pic1mem(1, i, j - 1) + pic1mem(1, i + 1, j - 1))
   LBlue = LBlue + _
   (1& * pic1mem(1, i - 1, j) + pic1mem(1, i + 1, j))
   LBlue = LBlue + _
   (1& * pic1mem(1, i - 1, j + 1) + pic1mem(1, i, j + 1) + pic1mem(1, i + 1, j + 1))
   
   LGreen = (1& * pic1mem(2, i - 1, j - 1) + pic1mem(2, i, j - 1) + pic1mem(2, i + 1, j - 1))
   LGreen = LGreen + _
   (1& * pic1mem(2, i - 1, j) + pic1mem(2, i + 1, j))
   LGreen = LGreen + _
   (1& * pic1mem(2, i - 1, j + 1) + pic1mem(2, i, j + 1) + pic1mem(2, i + 1, j + 1))
   
   LRed = (1& * pic1mem(3, i - 1, j - 1) + pic1mem(3, i, j - 1) + pic1mem(3, i + 1, j - 1))
   LRed = LRed + _
   (1& * pic1mem(3, i - 1, j) + pic1mem(3, i + 1, j))
   LRed = LRed + _
   (1& * pic1mem(3, i - 1, j + 1) + pic1mem(3, i, j + 1) + pic1mem(3, i + 1, j + 1))
   
   LBlue = 8 * pic1mem(1, i, j) - LBlue + 255
   LGreen = 8 * pic1mem(2, i, j) - LGreen + 255
   LRed = 8 * pic1mem(3, i, j) - LRed + 255
   
   If LBlue > 255 Then LBlue = 255
   If LGreen > 255 Then LGreen = 255
   If LRed > 255 Then LRed = 255
  
   If LBlue < 0 Then LBlue = 0
   If LGreen < 0 Then LGreen = 0
   If LRed < 0 Then LRed = 0
   
   pic2mem(1, i, j) = LBlue
   pic2mem(2, i, j) = LGreen
   pic2mem(3, i, j) = LRed

Next i
Next j

ShowDitheredPicture
Screen.MousePointer = vbDefault

End Sub

Private Sub cmdSave32_Click()
On Error GoTo Save32Error

Title$ = "Save 32bpp bmp"
Filt$ = "Save bmp|*.bmp"
InDir$ = CurrPath$ 'Pathspec$

CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, ""

If Len(FileSpec$) = 0 Then
   Close
   Exit Sub
End If

SavePicture pic2.Image, FileSpec$


Exit Sub
Save32Error:
End Sub

Private Sub Form_Load()

If App.PrevInstance Then End

Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False

pic1.ScaleMode = vbpixel
pic2.ScaleMode = vbpixel

Pathspec$ = App.Path
If Right$(Pathspec$, 1) <> "\" Then Pathspec$ = Pathspec$ & "\"
CurrPath$ = Pathspec$

pic1.Move 0, 0
'pic1.Picture = LoadPicture("lena.jpg")
LabFileName.Caption = "Lena"

picHt = pic1.Height
picWd = pic1.Width

pic2.Move 0, 0, pic1.Width, pic1.Height

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
GETDIBS pic1.Image  ' BPP Fills pic1mem and sizes pic2mem

'------------------------------
' Scroll bars set up in calling Form:-
'Timer1.Enabled = False
'Timer1.Interval = 60
Set frm = Form1
ReDim scrMax(0 To 1)
ReDim scrMin(0 To 1)
ReDim SmallChange(0 To 1)
ReDim LargeChange(0 To 1)
'------------------------------

'fraVert.Visible = False
'fraHorz.Visible = False

RESIZER

cmdPrint(0).Enabled = False
cmdSavePic.Enabled = False
picLims.Visible = False
ASMVB = False

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Load Machine code
Loadmcode Pathspec$ & "Dither.bin", DitherMC()
ptrStruc = VarPtr(DStruc.picWd)
ptrMC = VarPtr(DitherMC(1))
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
DStruc.picHt = picHt
DStruc.picWd = picWd
DStruc.Ptrpic1mem = VarPtr(pic1mem(1, 1, 1))
DStruc.Ptrpic2mem = VarPtr(pic2mem(1, 1, 1))
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Show
End Sub

'Private Sub cmdGrey_Click()  ' NOT USED
''NB: 1/3 of RGB taken for grey rather than other weightings
''    good enough here.
'
'Frame2.Visible = False
'
'LabPatternName = "Grey"
'
'Screen.MousePointer = vbHourglass
'For j = 1 To picHt
'For i = 1 To picWd
'   red = (1& * pic1mem(1, i, j) + pic1mem(2, i, j) + pic1mem(3, i, j)) \ 3
'   For k = 1 To 3
'      pic2mem(k, i, j) = red
'   Next k
'Next i
'Next j
'ShowDitheredPicture
'Screen.MousePointer = vbDefault
'End Sub

'###### DITHERING SUBS #############################################

Private Sub cmdBlackWhite_Click()
Frame2.Visible = False

LabPatternName = "Black and White"
Screen.MousePointer = vbHourglass

InitBits
'picLims.Visible = True
'cmdPrint(0).Enabled = True
'cmdSavePic.Enabled = True
'bm.Colors(1).rgbBlue = 255
'bm.Colors(1).rgbGreen = 255
'bm.Colors(1).rgbRed = 255

ReDim pic2mem(4, picWd, picHt)
DoEvents

If ASMVB Then
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   'Set at picture load
   DStruc.picHt = picHt
   DStruc.picWd = picWd
   DStruc.Ptrpic1mem = VarPtr(pic1mem(1, 1, 1))
   DStruc.Ptrpic2mem = VarPtr(pic2mem(1, 1, 1))
   
   DStruc.BlackThreshHold = BlackThreshHold
   ASM_BlackWhite
   ShowDitheredPicture
   Screen.MousePointer = vbDefault
   Exit Sub
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End If

' VB VB VB
ReDim pic2mem(4, picWd, picHt)
For j = 1 To picHt
For i = 1 To picWd
   red = (1& * pic1mem(1, i, j) + pic1mem(2, i, j) + pic1mem(3, i, j)) \ 3
   If red > BlackThreshHold Then  ' BlackThreshHold default=128
      pic2mem(1, i, j) = 255
      pic2mem(2, i, j) = 255
      pic2mem(3, i, j) = 255
   End If
Next i
Next j

ShowDitheredPicture
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdFloydSteinbergA_Click()

Frame2.Visible = False

LabPatternName = "Floyd-SteinbergA"

'Simplest Floyd-SteinbergA
'   x 3
'   3 2  /8

Screen.MousePointer = vbHourglass

InitBits

' Zero Intensity Array
ReDim IArr(-1 To picWd + 2, -1 To picHt + 2)
DoEvents

If ASMVB Then
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   DStruc.picHt = picHt
   DStruc.picWd = picWd
   DStruc.Ptrpic1mem = VarPtr(pic1mem(1, 1, 1))
   DStruc.Ptrpic2mem = VarPtr(pic2mem(1, 1, 1))
   
   DStruc.PtrIArr = VarPtr(IArr(-1, -1))
   DStruc.DithDIV = FSADIV
   ASM_FloydSteinbergA
   ShowDitheredPicture
   Screen.MousePointer = vbDefault
   Exit Sub
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End If

FillIntensityArray
zDiv = FSADIV '8
zMul = 1 / zDiv

For j = 1 To picHt '- 1
For i = 1 To picWd '- 1
   cul = 0
   If IArr(i, j) > greysum Then
      cul = 255
      pic2mem(1, i, j) = cul
      pic2mem(2, i, j) = cul
      pic2mem(3, i, j) = cul
   End If
   zErr = (IArr(i, j) - cul) * zMul

   ' Spread error
   'j+1   3 2
   IArr(i, j + 1) = IArr(i, j + 1) + 3 * zErr
   IArr(i + 1, j + 1) = IArr(i + 1, j + 1) + 2 * zErr
   'j   x 3
   IArr(i + 1, j) = IArr(i + 1, j) + 3 * zErr
Next i
Next j
ShowDitheredPicture
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdFloydSteinbergB_Click()

Frame2.Visible = False

LabPatternName = "Floyd-SteinbergB"

'Floyd-SteinbergB
'   x 7
' 3 5 1  /16

Screen.MousePointer = vbHourglass

InitBits

' Zero Intensity Array
ReDim IArr(-1 To picWd + 2, -1 To picHt + 2)
DoEvents

If ASMVB Then
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   DStruc.picHt = picHt
   DStruc.picWd = picWd
   DStruc.Ptrpic1mem = VarPtr(pic1mem(1, 1, 1))
   DStruc.Ptrpic2mem = VarPtr(pic2mem(1, 1, 1))
   
   DStruc.PtrIArr = VarPtr(IArr(-1, -1))
   DStruc.DithDIV = FSBDIV
   ASM_FloydSteinbergB
   ShowDitheredPicture
   Screen.MousePointer = vbDefault
   Exit Sub
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End If

FillIntensityArray
zDiv = FSBDIV '16
zMul = 1 / zDiv

For j = 1 To picHt
For i = 1 To picWd
   cul = 0
   If IArr(i, j) > greysum Then
      cul = 255
      pic2mem(1, i, j) = cul
      pic2mem(2, i, j) = cul
      pic2mem(3, i, j) = cul
   End If
   zErr = (IArr(i, j) - cul) * zMul
   
   ' Spread error
   'j+1   3 5 1
   IArr(i - 1, j + 1) = IArr(i - 1, j + 1) + 3 * zErr
   IArr(i, j + 1) = IArr(i, j + 1) + 5 * zErr
   IArr(i + 1, j + 1) = IArr(i + 1, j + 1) + zErr
   'j   x 7
   IArr(i + 1, j) = IArr(i + 1, j) + 7 * zErr
Next i
Next j
ShowDitheredPicture
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdStucki_Click()

Frame2.Visible = False

LabPatternName = "Stucki"

'Stucki
'    x 8 4
'2 4 8 4 2
'1 2 4 2 1  / 42

Screen.MousePointer = vbHourglass

InitBits

' Zero Intensity Array
ReDim IArr(-1 To picWd + 2, -1 To picHt + 2)
DoEvents

If ASMVB Then
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   DStruc.picHt = picHt
   DStruc.picWd = picWd
   DStruc.Ptrpic1mem = VarPtr(pic1mem(1, 1, 1))
   DStruc.Ptrpic2mem = VarPtr(pic2mem(1, 1, 1))
   
   DStruc.PtrIArr = VarPtr(IArr(-1, -1))
   DStruc.DithDIV = StuckiDiv
   ASM_Stucki
   ShowDitheredPicture
   Screen.MousePointer = vbDefault
   Exit Sub
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End If

FillIntensityArray
zDiv = StuckiDiv '42
zMul = 1 / zDiv

For j = 1 To picHt
For i = 1 To picWd
   cul = 0
   If IArr(i, j) > greysum Then
      cul = 255
      pic2mem(1, i, j) = cul
      pic2mem(2, i, j) = cul
      pic2mem(3, i, j) = cul
   End If
   zErr = (IArr(i, j) - cul) * zMul
   
   ' Spread error
      ' j+2 '1 2 4 2 1
   IArr(i - 2, j + 2) = IArr(i - 2, j + 2) + zErr
   IArr(i - 1, j + 2) = IArr(i - 1, j + 2) + 2 * zErr
   IArr(i, j + 2) = IArr(i, j + 2) + 4 * zErr
   IArr(i + 1, j + 2) = IArr(i + 1, j + 2) + 2 * zErr
   IArr(i + 2, j + 2) = IArr(i + 2, j + 2) + zErr
      'j+1  '2 4 8 4 2
   IArr(i - 2, j + 1) = IArr(i - 2, j + 1) + 2 * zErr
   IArr(i - 1, j + 1) = IArr(i - 1, j + 1) + 4 * zErr
   IArr(i, j + 1) = IArr(i, j + 1) + 8 * zErr
   IArr(i + 1, j + 1) = IArr(i + 1, j + 1) + 4 * zErr
   IArr(i + 2, j + 1) = IArr(i + 2, j + 1) + 2 * zErr
   'j x 8 4
   IArr(i + 1, j) = IArr(i + 1, j) + 8 * zErr
   IArr(i + 2, j) = IArr(i + 2, j) + 4 * zErr

Next i
Next j
ShowDitheredPicture
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdBurkes_Click()
Frame2.Visible = False

LabPatternName = "Burkes"

' Burkes
'    x 8 4
'2 4 8 4 2  / 32

Screen.MousePointer = vbHourglass

InitBits

' Zero Intensity Array
ReDim IArr(-1 To picWd + 2, -1 To picHt + 2)
DoEvents

If ASMVB Then
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   DStruc.picHt = picHt
   DStruc.picWd = picWd
   DStruc.Ptrpic1mem = VarPtr(pic1mem(1, 1, 1))
   DStruc.Ptrpic2mem = VarPtr(pic2mem(1, 1, 1))
   
   DStruc.PtrIArr = VarPtr(IArr(-1, -1))
   DStruc.DithDIV = BurkesDiv
   ASM_Burkes
   ShowDitheredPicture
   Screen.MousePointer = vbDefault
   Exit Sub
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End If

FillIntensityArray
zDiv = BurkesDiv '32
zMul = 1 / zDiv

For j = 1 To picHt
For i = 1 To picWd
   cul = 0
   If IArr(i, j) > greysum Then
      cul = 255
      pic2mem(1, i, j) = cul
      pic2mem(2, i, j) = cul
      pic2mem(3, i, j) = cul
   End If
   zErr = (IArr(i, j) - cul) * zMul

   ' Spread error
      'j+1  2 4 8 4 2
   IArr(i - 2, j + 1) = IArr(i - 2, j + 1) + 2 * zErr
   IArr(i - 1, j + 1) = IArr(i - 1, j + 1) + 4 * zErr
   IArr(i, j + 1) = IArr(i, j + 1) + 8 * zErr
   IArr(i + 1, j + 1) = IArr(i + 1, j + 1) + 4 * zErr
   IArr(i + 2, j + 1) = IArr(i + 2, j + 1) + 2 * zErr
   'j   x 8 4
   IArr(i + 1, j) = IArr(i + 1, j) + 8 * zErr
   IArr(i + 2, j) = IArr(i + 2, j) + 4 * zErr

Next i
Next j
ShowDitheredPicture
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSierra_Click()
Frame2.Visible = False

LabPatternName = "Sierra"

'Sierra
'    x 5 3
'2 4 5 4 2
'  2 3 2    / 32

Screen.MousePointer = vbHourglass

InitBits

' Zero Intensity Array
ReDim IArr(-1 To picWd + 2, -1 To picHt + 2)
DoEvents

If ASMVB Then
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   DStruc.picHt = picHt
   DStruc.picWd = picWd
   DStruc.Ptrpic1mem = VarPtr(pic1mem(1, 1, 1))
   DStruc.Ptrpic2mem = VarPtr(pic2mem(1, 1, 1))
   
   DStruc.PtrIArr = VarPtr(IArr(-1, -1))
   DStruc.DithDIV = SierraDiv
   ASM_Sierra
   ShowDitheredPicture
   Screen.MousePointer = vbDefault
   Exit Sub
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End If

FillIntensityArray
zDiv = SierraDiv '32
zMul = 1 / zDiv

For j = 1 To picHt
For i = 1 To picWd
   cul = 0
   If IArr(i, j) > greysum Then
      cul = 255
      pic2mem(1, i, j) = cul
      pic2mem(2, i, j) = cul
      pic2mem(3, i, j) = cul
   End If
   zErr = (IArr(i, j) - cul) * zMul
   
   ' Spread error
      ' j+2 ' - 2 3 2 -
   IArr(i - 1, j + 2) = IArr(i - 1, j + 2) + 2 * zErr
   IArr(i, j + 2) = IArr(i, j + 2) + 3 * zErr
   IArr(i + 1, j + 2) = IArr(i + 1, j + 2) + 2 * zErr
   'j+1  2 4 5 4 2
   IArr(i - 2, j + 1) = IArr(i - 2, j + 1) + 2 * zErr
   IArr(i - 1, j + 1) = IArr(i - 1, j + 1) + 4 * zErr
   IArr(i, j + 1) = IArr(i, j + 1) + 5 * zErr
   IArr(i + 1, j + 1) = IArr(i + 1, j + 1) + 4 * zErr
   IArr(i + 2, j + 1) = IArr(i + 2, j + 1) + 2 * zErr
   'j x 5 3
   IArr(i + 1, j) = IArr(i + 1, j) + 5 * zErr
   IArr(i + 2, j) = IArr(i + 2, j) + 3 * zErr
Next i
Next j
ShowDitheredPicture
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdJarvisEtAl_Click()
Frame2.Visible = False

LabPatternName = "Jarvis, Judice and Ninke"

'Jarvis, Judice and Ninke filter:
'    x 7 5
'3 5 7 5 3
'1 3 5 3 1  / 48

Screen.MousePointer = vbHourglass

InitBits

' Zero Intensity Array
ReDim IArr(-1 To picWd + 2, -1 To picHt + 2)
DoEvents

If ASMVB Then
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   DStruc.picHt = picHt
   DStruc.picWd = picWd
   DStruc.Ptrpic1mem = VarPtr(pic1mem(1, 1, 1))
   DStruc.Ptrpic2mem = VarPtr(pic2mem(1, 1, 1))
   
   DStruc.PtrIArr = VarPtr(IArr(-1, -1))
   DStruc.DithDIV = JarvisDiv
   ASM_Jarvis
   ShowDitheredPicture
   Screen.MousePointer = vbDefault
   Exit Sub
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End If

FillIntensityArray
zDiv = JarvisDiv '48
zMul = 1 / zDiv

For j = 1 To picHt
For i = 1 To picWd
   cul = 0
   If IArr(i, j) > greysum Then
      cul = 255
      pic2mem(1, i, j) = cul
      pic2mem(2, i, j) = cul
      pic2mem(3, i, j) = cul
   End If
   zErr = (IArr(i, j) - cul) * zMul
   
   ' Spread error
      ' j+2  1 3 5 3 1
   IArr(i - 2, j + 2) = IArr(i - 2, j + 2) + zErr
   IArr(i - 1, j + 2) = IArr(i - 1, j + 2) + 3 * zErr
   IArr(i, j + 2) = IArr(i, j + 2) + 5 * zErr
   IArr(i + 1, j + 2) = IArr(i + 1, j + 2) + 3 * zErr
   IArr(i + 2, j + 2) = IArr(i + 2, j + 2) + zErr
   'j+1  3 5 7 5 3
   IArr(i - 2, j + 1) = IArr(i - 2, j + 1) + 3 * zErr
   IArr(i - 1, j + 1) = IArr(i - 1, j + 1) + 5 * zErr
   IArr(i, j + 1) = IArr(i, j + 1) + 7 * zErr
   IArr(i + 1, j + 1) = IArr(i + 1, j + 1) + 5 * zErr
   IArr(i + 2, j + 1) = IArr(i + 2, j + 1) + 3 * zErr
   'j x 7 5
   IArr(i + 1, j) = IArr(i + 1, j) + 7 * zErr
   IArr(i + 2, j) = IArr(i + 2, j) + 5 * zErr
Next i
Next j
ShowDitheredPicture
Screen.MousePointer = vbDefault
End Sub

'=======================================================================
'=======================================================================

Private Sub DO_PATTERN(Limits(), Ptr)

' Zero Intensity Array
ReDim IArr(-1 To picWd + 6, -1 To picHt + 6)
DoEvents

If ASMVB Then
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   DStruc.picHt = picHt
   DStruc.picWd = picWd
   DStruc.Ptrpic1mem = VarPtr(pic1mem(1, 1, 1))
   DStruc.Ptrpic2mem = VarPtr(pic2mem(1, 1, 1))
   
   DStruc.PtrIArr = VarPtr(IArr(-1, -1))
   
   DStruc.NX = NX
   DStruc.NY = NY
   DStruc.NT = NT
   DStruc.PtrLimits = Ptr  ' VarPtr(HalftoneLimits(1)) etc
   DStruc.PtrTM = VarPtr(TM(1, 1, 1))
   ASM_DOPATTERN
   ShowDitheredPicture
   Screen.MousePointer = vbDefault
   Exit Sub
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End If



'NB Written to match mcode

FillIntensityArray
For jj = 1 To picHt Step NY
jup = jj + NY - 1

For ii = 1 To picWd Step NX
iup = ii + NX - 1
   
   GetAvGrey jj, ii, NX, NY, AvGrey
   
   If AvGrey = 0 Then GoTo NexJJ
   
   For n = 1 To NT
       If AvGrey < Limits(n) Then GoTo NexN
         
         For j = jj To jup
            If j > picHt Then GoTo NexN
               For i = ii To iup
                  If i > picWd Then GoTo NexJ
                     If TM(i - ii + 1, j - jj + 1, n) = 49 Then
                        pic2mem(1, i, j) = 255
                        pic2mem(2, i, j) = 255
                        pic2mem(3, i, j) = 255
'                     Else ' Alternative *  NB Not in ASM
'                        pic2mem(1, i, j) = 0
'                        pic2mem(2, i, j) = 0
'                        pic2mem(3, i, j) = 0
                     End If
               Next i
NexJ:
         Next j
NexN:
   Next n

Next ii
NexJJ:
Next jj

ShowDitheredPicture
Screen.MousePointer = vbDefault
End Sub

'=======================================================================
Private Sub cmdHalfTone_Click()
Frame2.Visible = False
LabPatternName = "Half tone"
'Half tone
' (i,j) 0 0 0
'       0 0 0
'       0 0 0 (i+2,j+2)

' 0 0 0  0 0 0  W 0 0  W 0 W  W 0 W  W 0 W  W 0 W  W W W  W W W  W W W
' 0 0 0  0 0 0  0 0 0  0 0 0  0 0 0  0 0 0  W 0 0  W 0 0  W 0 W  W W W
' 0 0 0  0 0 W  0 0 W  0 0 W  W 0 W  W W W  W W W  W W W  W W W  W W W
'          1      2      3      4      5      6      7      8      9
'         25,    50,    75,   100,   125,   150,   175,   200,   230   >230

Screen.MousePointer = vbHourglass

InitBits

DoEvents

NX = 3
NY = 3
NT = 9
ReDim TM(NX, NY, NT)
ReDim TS$(NT)
' These patterns can be changed
' to give other effects. There are
' NT strings of length NX*NY
' NB 1's in a string will mask
'    1's in later strings
TS$(1) = "000000001"
TS$(2) = "100000000"
TS$(3) = "001000000"
TS$(4) = "000000100"
TS$(5) = "000000010"
TS$(6) = "000100000"
TS$(7) = "010000000"
TS$(8) = "000001000"
TS$(9) = "000010000"

Sum$ = ""
For k = 1 To NT
   Sum$ = Sum$ & TS$(k)
Next k
CopyMemory TM(1, 1, 1), ByVal Sum$, Len(Sum$)

Ptr = VarPtr(HalftoneLimits(1))
DO_PATTERN HalftoneLimits(), Ptr

End Sub

Private Sub cmdHalfTone2_Click() ' Fuzzier half-tone
Frame2.Visible = False

LabPatternName = "Half tone 2"

'Half tone 2
' (i,j) 0 0 0 0 0
'       0 0 0 0 0
'       0 0 0 0 0
'       0 0 0 0 0
'       0 0 0 0 0(i+4,j+4)

'   1      2      3      4      5      6      7      8      9
'   25,    50,    75,   100,   125,   150,   175,   200,   230   >230

Screen.MousePointer = vbHourglass

InitBits

DoEvents

NX = 5
NY = 5
NT = 9
ReDim TM(NX, NY, NT)
ReDim TS$(NT)
TS$(1) = "0000000000001000000000000"
TS$(2) = "0000000000010100000000000"
TS$(3) = "0000000100000000010000000"
TS$(4) = "0000001010000000101000000"
TS$(5) = "0010000000100010000000100"
TS$(6) = "0010010001000001000100000"
TS$(7) = "0101000000000000000001010"
TS$(8) = "0000100000000000000010000"
TS$(9) = "1000000000000000000000001"

Sum$ = ""
For k = 1 To NT
   Sum$ = Sum$ & TS$(k)
Next k
CopyMemory TM(1, 1, 1), ByVal Sum$, Len(Sum$)

Ptr = VarPtr(HalftoneLimits(1))
DO_PATTERN HalftoneLimits(), Ptr

End Sub

Private Sub cmdLine_Click()
Frame2.Visible = False

LabPatternName = "Line"

'Line
' 0 0 0     0 0 0    1 1 1    1 1 1
' 0 0 0     1 1 1    0 0 0    1 1 1
' 0 0 0     0 0 0    1 1 1    1 1 1
'  <60      <120      <180     >=180
'   1         2         3
Screen.MousePointer = vbHourglass

InitBits

DoEvents

NX = 3
NY = 3
NT = 3
ReDim TM(NX, NY, NT)
ReDim TS$(NT)
TS$(1) = "000000000"
TS$(2) = "000111000"
TS$(3) = "111000111"
Sum$ = ""
For k = 1 To NT
   Sum$ = Sum$ & TS$(k)
Next k
CopyMemory TM(1, 1, 1), ByVal Sum$, Len(Sum$)

Ptr = VarPtr(LineLimits(1))
DO_PATTERN LineLimits(), Ptr

End Sub


Private Sub cmdLine2_Click()  ' Coarse lines
Frame2.Visible = False

LabPatternName = "Line 2"

'Line 2
' 0 0 0  0 0 0  0 0 0  0 0 0  0 0 W  0 0 0  0 W 0  0 0 0  W 0 0  0 0 0
' 0 W 0  W 0 W  0 0 0  0 0 0  0 0 0  0 0 0  0 0 0  0 0 0  0 0 0  0 0 0
' 0 0 0  0 0 0  0 0 0  0 0 0  W 0 0  0 0 0  0 W 0  0 0 0  0 0 W  0 0 0
'          1      2      3      4      5      6      7      8      9
'         25,    50,    75,   100,   125,   150,   175,   200,   230   >230

Screen.MousePointer = vbHourglass

InitBits

DoEvents

NX = 3
NY = 3
NT = 9
ReDim TM(NX, NY, NT)
ReDim TS$(NT)

TS$(1) = "000010000"
TS$(2) = "000101000"
TS$(3) = "000000000"
TS$(4) = "001000100"
TS$(5) = "000000000"
TS$(6) = "010000010"
TS$(7) = "100000001"
TS$(8) = "100000001"
TS$(9) = "000000000"

Sum$ = ""
For k = 1 To NT
   Sum$ = Sum$ & TS$(k)
Next k
CopyMemory TM(1, 1, 1), ByVal Sum$, Len(Sum$)

Ptr = VarPtr(LineLimits2(1))
DO_PATTERN LineLimits2(), Ptr


End Sub

Private Sub cmdRandom_Click()
Frame2.Visible = False

LabPatternName = "Random"

InitBits

DoEvents

'Random
' 0 0 0
' 0 0 0
' 0 0 0

'   1      2      3      4      5      6      7      8      9
'   25,    50,    75,   100,   125,   150,   175,   200,   230   >230

Randomize 8

NY = 3
NX = 3
NT = 9

Screen.MousePointer = vbHourglass

' Zero Intensity Array
ReDim IArr(-1 To picWd + 6, -1 To picHt + 6)
DoEvents

'' Ensure same sequence each time
'Rnd -1
'Randomize 1
'DStruc.Ranseed = 255 * Rnd

If ASMVB Then
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   DStruc.picHt = picHt
   DStruc.picWd = picWd
   DStruc.Ptrpic1mem = VarPtr(pic1mem(1, 1, 1))
   DStruc.Ptrpic2mem = VarPtr(pic2mem(1, 1, 1))
   DStruc.PtrLimits = VarPtr(RandomLimits(1))
   DStruc.NT = NT
   DStruc.NX = NX
   DStruc.NY = NY
   
   DStruc.PtrIArr = VarPtr(IArr(-1, -1))
   DStruc.RandSeed = Int(255 * Rnd)
   
   ASM_Random
   ShowDitheredPicture
   Screen.MousePointer = vbDefault
   Exit Sub
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End If

FillIntensityArray

For jj = 1 To picHt Step NY
jup = jj + NY - 1
For ii = 1 To picWd Step NX
iup = ii + NX - 1

   GetAvGrey jj, ii, NX, NY, AvGrey
   
   If AvGrey = 0 Then Exit For

   For n = 1 To 9
      If AvGrey < RandomLimits(n) Then
         RandEval n, jj, ii
         Exit For
      End If
   Next n
   If n = 10 Then RandEval 10, jj, ii
Next ii
Next jj
ShowDitheredPicture
Screen.MousePointer = vbDefault
End Sub

Private Sub RandEval(ByVal k, jj, ii)
' VB Code to roughly match mcode

jjup = jj + NY - 1
iiup = ii + NX - 1

If k < 10 Then
   If RandomLimits(k) >= 128 Then
      For j = jj To jjup
         If j > picHt Then GoTo TCulr
            For i = ii To iiup
               If i > picWd Then GoTo NexJ
                  pic2mem(1, i, j) = 255
                  pic2mem(2, i, j) = 255
                  pic2mem(3, i, j) = 255
            Next i
NexJ:
      Next j
      Exit Sub
   End If
End If

TCulr:

If k > 0 And k < 9 Then
   For k = 1 To 8
      ' Variations
      'j = Int(((jjup - jj + 1) * Rnd) + jj)
      'i = Int(((iiup - ii + 1) * Rnd) + ii)
      
      'j = Int(((jjup - jj) * Rnd) + jj)
      'i = Int(((iiup - ii) * Rnd) + ii)
      
      j = Int(((jjup) * Rnd) + jj)
      i = Int(((iiup) * Rnd) + ii)
      
      If i > picWd Or j > picHt Then GoTo Nexk
      pic2mem(1, i, j) = 255
      pic2mem(2, i, j) = 255
      pic2mem(3, i, j) = 255
Nexk:
   Next k
End If
End Sub

Private Sub cmdInvert_Click()
res = SetRect(IR, 0, 0, picWd, picHt)
res = InvertRect(pic2.hdc, IR)
pic2.Refresh

' Invert for saving
bm.Colors(0).rgbBlue = 255 - bm.Colors(0).rgbBlue
bm.Colors(0).rgbGreen = 255 - bm.Colors(0).rgbGreen
bm.Colors(0).rgbRed = 255 - bm.Colors(0).rgbRed

bm.Colors(1).rgbBlue = 255 - bm.Colors(1).rgbBlue
bm.Colors(1).rgbGreen = 255 - bm.Colors(1).rgbGreen
bm.Colors(1).rgbRed = 255 - bm.Colors(1).rgbRed
End Sub

Private Sub picQBColor_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
LongCul = GetPixel(picQBColor.hdc, x, y)

If LongCul = -1 Or LongCul = 0 Then Exit Sub
   
Screen.MousePointer = vbHourglass
   
red = LongCul And &HFF&
green = (LongCul And &HFF00&) / &H100&
blue = (LongCul And &HFF0000) / &H10000

' Put in new color where pic2mem is white
For j = 1 To picHt
For i = 1 To picWd
   If pic2mem(1, i, j) = 255 Then
      pic2mem(1, i, j) = blue
      pic2mem(2, i, j) = green
      pic2mem(3, i, j) = red
   End If
Next i
Next j

ShowDitheredPicture

' Restore pic2mem for further color changing
For j = 1 To picHt
For i = 1 To picWd
   If pic2mem(1, i, j) = blue Then
      If pic2mem(2, i, j) = green Then
         If pic2mem(3, i, j) = red Then
            pic2mem(1, i, j) = 255
            pic2mem(2, i, j) = 255
            pic2mem(3, i, j) = 255
         End If
      End If
   End If
Next i
Next j

' For saving & printing
bm.Colors(1).rgbBlue = blue
bm.Colors(1).rgbGreen = green
bm.Colors(1).rgbRed = red

Screen.MousePointer = vbDefault
End Sub


'###### DITHERED COMMON SUBS #######################################

Private Sub InitBits()
picLims.Visible = True
cmdPrint(0).Enabled = True
cmdSavePic.Enabled = True
bm.Colors(0).rgbBlue = 0
bm.Colors(0).rgbGreen = 0
bm.Colors(0).rgbRed = 0
bm.Colors(1).rgbBlue = 255
bm.Colors(1).rgbGreen = 255
bm.Colors(1).rgbRed = 255
End Sub

Private Sub FillIntensityArray()
' Find average grey value
' Public greysum
'ASM_FillIntensityArray  ' Test Checked out
'Exit Sub

greysum = 0
greycount = picHt * picWd

For j = 1 To picHt
For i = 1 To picWd
   red = (1& * pic1mem(1, i, j) + pic1mem(2, i, j) + pic1mem(3, i, j)) \ 3
   IArr(i, j) = red
   greysum = greysum + red 'IArr(i, j)
Next i
Next j
greysum = greysum / greycount
' Zero pic2mem
ReDim pic2mem(4, picWd, picHt)
End Sub

Private Sub GetAvGrey(jj, ii, NX, NY, AvGrey)
' In: y,x,template size NX by NY Out: AvGrey
AvCount = 0
AvGrey = 0
For j = jj To jj + NY - 1
   For i = ii To ii + NX - 1
      AvGrey = AvGrey + IArr(i, j)
      AvCount = AvCount + 1
   Next i
Next j
If AvCount = 0 Then AvGrey = 0 Else AvGrey = AvGrey \ AvCount
End Sub

Private Sub ShowDitheredPicture()

'Erase IArr

'Public Const DIB_PAL_COLORS = 1 '  uses system colors
'Public Const DIB_RGB_COLORS = 0 '  uses RGBQUAD colors

If StretchDIBits(pic2.hdc, _
   0, 0, picWd, picHt, _
   0, 0, picWd, picHt, _
   pic2mem(1, 1, 1), bm, _
   DIB_RGB_COLORS, vbSrcCopy) = 0 Then
      MsgBox "Blit Error", , "Dither"
      Done = True
      Erase pic1mem, pic2mem
      Unload Me
      End
End If
pic2.Refresh
End Sub

'###### LOAD PICTURE ###############################################

Private Sub cmdLoadPic_Click()

On Error GoTo LoadError

Title$ = "Load a picture file"
Filt$ = "Pics bmp,jpg,gif,ico,cur,wmf,emf|*.bmp;*.jpg;*.gif;*.ico;*.cur;*.wmf;*.emf"
InDir$ = CurrPath$ 'Pathspec$

CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, ""

If Len(FileSpec$) = 0 Then
   Close
   Exit Sub
End If

Erase pic1mem, pic2mem
pic1.Picture = LoadPicture
pic2.Picture = LoadPicture
DoEvents

pic1.Picture = LoadPicture(FileSpec$)
LabFileName.Caption = ExtractFileName$(FileSpec$)

CurrPath$ = FileSpec$

picWd = pic1.Width
picHt = pic1.Height

pic2.Move 0, 0, pic1.Width, pic1.Height

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
GETDIBS pic1.Image  ' BPP Fills pic1mem and sizes pic2mem

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
DStruc.picWd = picWd
DStruc.picHt = picHt
DStruc.Ptrpic1mem = VarPtr(pic1mem(1, 1, 1))
DStruc.Ptrpic2mem = VarPtr(pic2mem(1, 1, 1))
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Timer1.Enabled = False

fraScr(0).Visible = False
fraScr(1).Visible = False

Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False

RESIZER
On Error GoTo 0
Exit Sub
'==========
LoadError:
Erase pic1mem, pic2mem
DoEvents
pic1.Picture = LoadPicture
pic2.Picture = LoadPicture
MsgBox "VB doesn't like this picture ??", vbCritical, "Dither"
Exit Sub
End Sub

'###### SAVE FILE ##################################################

Private Sub cmdSavePic_Click()

On Error GoTo SaveError

Title$ = "Save 2-color bmp"
Filt$ = "Save bmp|*.bmp"
InDir$ = CurrPath$ 'Pathspec$

CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, ""

If Len(FileSpec$) = 0 Then
   Close
   Exit Sub
End If

' For bpp=1
BScanLine = (picWd + 7) \ 8
' Expand to 4B boundary
BScanLine = ((BScanLine + 3) \ 4) * 4

ReDim picsave(BScanLine, picHt)  ' BScanLine bytes

' Transfer pic2mem to 2-color picsave
' (1,1,1)(2,1,1)(3,1,1)(4,1,1)... (1,picWd,1)(2,picWd,1)(3,picWd,1)(4,picWd,1)
'   0 (< 20)or 255(>248)
' to (building bits from 4-bytes)
' picsave(1,1)....(picWd\8,1)
'
' picsave(1,picHt)....(picWd\8,pivHt)

For j = 1 To picHt
   n = 1
   b = 1
   For i = 1 To picWd
      If pic2mem(1, i, j) < 20 Then
         If b < 8 Then picsave(n, j) = picsave(n, j) * 2
      Else
         picsave(n, j) = picsave(n, j) Or 1
         If b < 8 Then picsave(n, j) = picsave(n, j) * 2
      End If
      b = b + 1
      If b = 9 Then
         n = n + 1
         b = 1
      End If
   Next i
   
   If b <> 1 Then
      For k = b To 7
         picsave(n, j) = picsave(n, j) * 2
      Next k
   End If

Next j
   
''''''''''''''''''''''''''''''''''''''''''''''''''
' Save 32 bpp & image size for further dithering
svbiBitCount = bm.bmiH.biBitCount
svbiSizeImage = bm.bmiH.biSizeImage
''''''''''''''''''''''''''''''''''''''''''''''''''

bm.bmiH.biBitCount = 1
bm.bmiH.biSizeImage = BScanLine * picHt

Open FileSpec$ For Binary As #1
TheBM% = 19778: Put #1, , TheBM%
FileSize& = 62 + BScanLine * picHt: Put #1, , FileSize&
ires% = 0: Put #1, , ires%: Put #1, , ires%
off& = 62: Put #1, , off&
'' BITMAPINFOHEADER =
Put #1, , bm
Put #1, , picsave
Close
Erase picsave
''''''''''''''''''''''''''''''''''''''''''''''''''
' Restore 32 bpp & image size
bm.bmiH.biBitCount = svbiBitCount
bm.bmiH.biSizeImage = svbiSizeImage
''''''''''''''''''''''''''''''''''''''''''''''''''
Frame2.Visible = False
Exit Sub
'=============
SaveError:
MsgBox "Error in saving", , "Dither"
Close
Erase picsave
''''''''''''''''''''''''''''''''''''''''''''''''''
' Restore 32 bpp & image size
bm.bmiH.biBitCount = svbiBitCount
bm.bmiH.biSizeImage = svbiSizeImage
''''''''''''''''''''''''''''''''''''''''''''''''''
Frame2.Visible = False
Frame3.Visible = False
End Sub

'###### PRINTING ###################################################

Private Sub cmdPrint_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Select Case Index
Case 0
   Frame3.Visible = True
   Exit Sub
Case 1, 2, 3, 4
Case Else
   Exit Sub
End Select

' Print Index x real size, offset 400,400 printer units
Screen.MousePointer = vbHourglass

Printer.Print " ";

If StretchDIBits(Printer.hdc, 400, 400, Index * picWd, Index * picHt, 0, 0, _
   picWd, picHt, pic2mem(1, 1, 1), bm, DIB_RGB_COLORS, vbSrcCopy) = 0 Then
      MsgBox "Printing failed", , "Dither"
      Printer.EndDoc
      Frame2.Visible = False
      Frame3.Visible = False
      Screen.MousePointer = vbDefault
      Exit Sub
   End If

Printer.NewPage
Printer.EndDoc

Frame2.Visible = False
Frame3.Visible = False
Screen.MousePointer = vbDefault

End Sub

Private Sub cmdVBASM_Click()
ASMVB = Not ASMVB
If ASMVB Then
   cmdVBASM.Caption = "ASM"
   cmdVBASM.BackColor = RGB(255, 200, 0)
   picLims.CurrentX = picLims.Width \ 6
   picLims.CurrentY = 0
   picLims.Print "STROKE ME"
   picLims.Refresh
Else
   cmdVBASM.Caption = "VB"
   cmdVBASM.BackColor = RGB(255, 255, 255)
   picLims.CurrentX = picLims.Width \ 6
   picLims.CurrentY = 0
   picLims.Print "CLICK ME "
   picLims.Refresh
End If
DoEvents
End Sub


'###### INITIALIZE #################################################

Private Sub Form_Initialize()

'Dim zHSc, zVSc

' Get screen res
FWW = GetSystemMetrics(SM_CXSCREEN)
FHH = GetSystemMetrics(SM_CYSCREEN)
' To show Desktop border
FW = 0.95 * FWW
FH = 0.9 * FHH
' Scaling for screen res changes
zHSc = (350 / 800) * (FWW / FHH) * 3 / 4
zVSc = (500 / 600) * (FWW / FHH) * 3 / 4

SetDefaultFactors ' Thresholds & Divisors

End Sub

'###### RESIZING ###################################################

Private Sub chkResizer_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Frame1.Visible = False
Frame2.Visible = False
chkResizer.Value = Unchecked
RESIZER
End Sub

Private Sub RESIZER()

' Screen res resizer

If WindowState = vbMinimized Then Exit Sub

' Get screen res
FWW = GetSystemMetrics(SM_CXSCREEN)
FHH = GetSystemMetrics(SM_CYSCREEN)
' To show Desktop border
FW = 0.95 * FWW
FH = 0.9 * FHH
   

If WindowState <> vbMaximized Then
   Form1.Top = 250
   Form1.Left = 250
   Form1.Width = FW * Screen.TwipsPerPixelX
   Form1.Height = FH * Screen.TwipsPerPixelY
End If

Picture1.Width = FW * zHSc
Picture1.Height = FH * zVSc

Picture1.Top = 50
pic1.Top = 0

With fraScr(0)
   .Left = 10
   .Top = Picture1.Top + 2
   .Width = 20
   .Height = Picture1.Height '+ 7
End With

Picture1.Left = fraScr(0).Left + fraScr(0).Width + 5
pic1.Left = 0 'Picture1.Left

With fraScr(1)
   .Top = Picture1.Top + Picture1.Height + 2
   .Left = Picture1.Left
   .Height = 20
   .Width = Picture1.Width
End With
'
'
'With fraVert
'   .Left = 10
'   .Top = Picture1.Top - 8
'   .Width = 25
'   '.Height = Picture1.Height + 7
'End With
'
'Picture1.Left = fraVert.Left + fraVert.Width + 5
'pic1.Left = 0 'Picture1.Left
'
'With fraHorz
'   .Top = Picture1.Top + Picture1.Height - 5
'   .Left = Picture1.Left
'   .Height = 30
'   '.Width = Picture1.Width
'End With

With Picture2
   .Top = Picture1.Top
   .Left = Picture1.Left + Picture1.Width + 10
   .Height = Picture1.Height
   .Width = Picture1.Width
End With
pic2.Top = 0
pic2.Left = 0

With picLims
   .Top = Picture1.Top - picLims.Height - 6
   .Left = Picture2.Left
   .Width = Picture2.Width
End With
zincr = picLims.Width / 256
cul = 255
For px = 0 To picLims.Width - 1
   picLims.Line (px, 0)-(px, picLims.Height), RGB(cul, cul, cul)
   cul = cul - zincr
   If cul < 0 Then cul = 0
Next px
picLims.CurrentX = picLims.Width \ 6
picLims.CurrentY = 0
If ASMVB Then
   picLims.Print "STROKE ME"
Else
   picLims.Print "CLICK ME "
End If
picLims.Refresh
picLims.Visible = True

' Other controls
chkResizer.Left = Picture1.Left + Picture1.Width - chkResizer.Width - 2
LabPatternName.Left = Picture2.Left + 2
cmdExit(1).Left = Form1.Width / Screen.TwipsPerPixelX - cmdExit(1).Width - 6
cmdVBASM.Left = picLims.Left + picLims.Width - cmdVBASM.Width + 1

Shape1.Move 3, 3, Form1.Width / Screen.TwipsPerPixelX - 6, Form1.Height / Screen.TwipsPerPixelY - 6
cmdPatterns.Left = chkResizer.Left - cmdPatterns.Width - 2
Frame1.Left = cmdPatterns.Left

'--- Vertical Scroll Bar Settings -------------
scrMax(0) = picHt:
scrMin(0) = 0
SmallChange(0) = 1
LargeChange(0) = 1 ' NB > 0
'--- Horizontal Scroll Bar Settings -----------
scrMax(1) = picWd:
scrMin(1) = 0
SmallChange(1) = 1
LargeChange(1) = 1 ' NB > 0
'----------------------------------------------

If picHt < 4 Or picWd < 4 Then
   MsgBox "Very small or thin picture", , "Dither"
End If

ModReDim

If picHt > Picture1.Height Then
   fraScr(0).Visible = True
   scrMax(0) = picHt - Picture1.Height + 6
   LargeChange(0) = scrMax(0) \ 10
   If LargeChange(0) < 1 Then LargeChange(0) = 1
Else
   fraScr(0).Visible = False
End If

If picWd > Picture1.Width Then
   fraScr(1).Visible = True
   scrMax(1) = picWd - Picture1.Width + 6
   LargeChange(1) = scrMax(1) \ 10
   If LargeChange(1) < 1 Then LargeChange(1) = 1
Else
   fraScr(1).Visible = False
End If
'ModReDim
VVal = nTVal(0)
HVal = nTVal(1)
'----------------------------------------------

End Sub

'###### EXIT #######################################################

Private Sub Form_Unload(Cancel As Integer)
Erase pic1mem, pic2mem
Unload Me
End
End Sub

Private Sub cmdExit_Click(Index As Integer)
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Form_Unload 0
End Sub

'###### TOGGLE FRAMES ##############################################

Private Sub cmdFile_Click()
Frame1.Visible = False
Frame2.Visible = Not Frame2.Visible
Frame3.Visible = False
End Sub

Private Sub cmdPatterns_Click()
Frame1.Visible = Not Frame1.Visible
Frame2.Visible = False
Frame3.Visible = False
End Sub

'###### CLEAR FRAMES ###############################################

Private Sub Form_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
End Sub

Private Sub pic1_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
End Sub

Private Sub pic2_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
End Sub

Private Sub Picture1_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
End Sub

Private Sub pic2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
End Sub

'###### CHANGE WEIGHTING ###########################################

Private Sub picLims_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If ASMVB Then
   'Increase divisors
   zFrac = (x + picLims.Width / 2) / (picLims.Width / 2)
   ReShow zFrac
Else
   'SetCursor LoadCursor(0, IDC_HAND)
   SetCursor LoadCursor(0, IDC_LRTRI)
End If
End Sub

Private Sub picLims_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Not ASMVB Then
   SetCursor LoadCursor(0, IDC_ARROW)
   'Increase divisors (at x = 0 zFrac=1 at
   ' x = picLims.Width zFrac = 3)
   zFrac = (x + picLims.Width / 2) / (picLims.Width / 2)
   ReShow zFrac
End If
End Sub

Private Sub ReShow(zFrac)

picLims.Enabled = False
Select Case LabPatternName.Caption

Case "Black and White"
'Default BlackThreshHold = 128
   BlackThreshHold = BlackThreshHold * zFrac * 0.5
   If BlackThreshHold > 255 Then BlackThreshHold = 255
   cmdBlackWhite_Click
   SetDefaultFactors

Case "Floyd-SteinbergA"
'Default FSADiv = 8
   FSADIV = FSADIV * zFrac
   cmdFloydSteinbergA_Click
   SetDefaultFactors

Case "Floyd-SteinbergB"
'Default FSBDiv = 16
   FSBDIV = FSBDIV * zFrac
   cmdFloydSteinbergB_Click
   SetDefaultFactors

Case "Stucki"
'Default StuckiDiv = 42
   StuckiDiv = StuckiDiv * zFrac
   cmdStucki_Click
   SetDefaultFactors

Case "Burkes"
'Default BurkesDiv = 32
   BurkesDiv = BurkesDiv * zFrac
   cmdBurkes_Click
   SetDefaultFactors

Case "Sierra"
'Default SierraDiv = 32
   SierraDiv = SierraDiv * zFrac
   cmdSierra_Click
   SetDefaultFactors

Case "Jarvis, Judice and Ninke"
'Default JarvisDiv = 48
   JarvisDiv = JarvisDiv * zFrac
   cmdJarvisEtAl_Click
   SetDefaultFactors

'------------------------------------------------
Case "Half tone"
   For k = 1 To 9
      HalftoneLimits(k) = HalftoneLimits(k) * zFrac * 0.5
   Next k
   cmdHalfTone_Click
   SetDefaultFactors

Case "Half tone 2"
   For k = 1 To 9
      HalftoneLimits(k) = HalftoneLimits(k) * zFrac * 0.5
   Next k
   cmdHalfTone2_Click
   SetDefaultFactors

Case "Line"
   For k = 1 To 3
      LineLimits(k) = LineLimits(k) * zFrac * 0.5
   Next k
   cmdLine_Click
   SetDefaultFactors

Case "Line 2"
   For k = 1 To 7
      LineLimits2(k) = LineLimits2(k) * zFrac * 0.5
   Next k
   cmdLine2_Click
   SetDefaultFactors

Case "Random"
   For k = 1 To 9
      RandomLimits(k) = RandomLimits(k) * zFrac * 0.5
   Next k
   cmdRandom_Click
   SetDefaultFactors

End Select
DoEvents
picLims.Enabled = True
End Sub

'###################################################################
'####### SCROLL BARS ###############################################
'###################################################################

Private Sub LabThumb_MouseDown(Index As Integer, Button As Integer, Shift As Integer, _
x As Single, y As Single)
Frame1.Visible = False
Frame2.Visible = False
   LabThumbMouseDown Index, x, y
   
   If nTVal(0) <> VVal Or nTVal(1) <> HVal Then
      MovePics
      VVal = nTVal(0)
      HVal = nTVal(1)
   End If
   
End Sub

Private Sub LabThumb_MouseMove(Index As Integer, Button As Integer, Shift As Integer, _
x As Single, y As Single)
   LabThumbMouseMove Index, x, y
   
   If nTVal(0) <> VVal Or nTVal(1) <> HVal Then
      MovePics
      VVal = nTVal(0)
      HVal = nTVal(1)
   End If

End Sub

Private Sub LabThumb_MouseUp(Index As Integer, Button As Integer, Shift As Integer, _
x As Single, y As Single)
   ' Snap thumb position
   LabThumbMouseUp Index, x, y
   
   If nTVal(0) <> VVal Or nTVal(1) <> HVal Then
      MovePics
      VVal = nTVal(0)
      HVal = nTVal(1)
   End If

End Sub

'###################################################################

Private Sub LabBack_MouseDown(Index As Integer, Button As Integer, Shift As Integer, _
x As Single, y As Single)
Frame1.Visible = False
Frame2.Visible = False
   
   LabBackMouseDown Index, x, y
   
   If nTVal(0) <> VVal Or nTVal(1) <> HVal Then
      MovePics
      VVal = nTVal(0)
      HVal = nTVal(1)
   End If

End Sub

Private Sub LabBack_MouseUp(Index As Integer, Button As Integer, Shift As Integer, _
x As Single, y As Single)
   LabBackMouseUp
   
   If nTVal(0) <> VVal Or nTVal(1) <> HVal Then
      MovePics
      VVal = nTVal(0)
      HVal = nTVal(1)
   End If

End Sub

'###################################################################

Private Sub cmdScr_MouseDown(Index As Integer, Button As Integer, Shift As Integer, _
x As Single, y As Single)
Frame1.Visible = False
Frame2.Visible = False
   
   cmdScrMouseDown Index, Button, x, y
End Sub

Private Sub cmdScr_MouseUp(Index As Integer, Button As Integer, Shift As Integer, _
x As Single, y As Single)
Frame1.Visible = False
Frame2.Visible = False
   
   cmdScrMouseUp
End Sub

Private Sub cmdScr_Click(Index As Integer)

   cmdScrMouseUp

End Sub


'###################################################################

Private Sub Timer1_Timer()
   Timer1Timer
   'LabVal(Index).Caption = nTVal(Index)  ' Out
   
   If nTVal(0) <> VVal Or nTVal(1) <> HVal Then
      MovePics
      VVal = nTVal(0)
      HVal = nTVal(1)
   End If

End Sub

'###################################################################
'###################################################################

Private Sub MovePics()
' Vert
pic1.Top = -nTVal(0)
pic1.Refresh
pic2.Top = -nTVal(0)
pic2.Refresh
' Horz
pic1.Left = -nTVal(1)
pic1.Refresh
pic2.Left = -nTVal(1)
pic2.Refresh
End Sub

