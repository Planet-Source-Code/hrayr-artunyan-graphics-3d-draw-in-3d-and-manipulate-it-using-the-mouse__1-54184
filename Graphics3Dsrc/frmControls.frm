VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmControl 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Computer Graphics - Control Pannel"
   ClientHeight    =   5730
   ClientLeft      =   8925
   ClientTop       =   2385
   ClientWidth     =   6330
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6330
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   255
      Left            =   3360
      TabIndex        =   58
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reset"
      Height          =   255
      Left            =   120
      TabIndex        =   57
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmdLines 
      Caption         =   "Show Lines"
      Height          =   375
      Left            =   4560
      TabIndex        =   54
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdvanced 
      Caption         =   "SHOW  ADVANCED   CONTROLS"
      Height          =   4575
      Left            =   4200
      TabIndex        =   53
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox picAdvanced 
      BackColor       =   &H8000000C&
      Height          =   5175
      Left            =   4440
      ScaleHeight     =   5115
      ScaleWidth      =   1755
      TabIndex        =   38
      Top             =   480
      Width           =   1815
      Begin VB.CommandButton cmdShowCanvas 
         BackColor       =   &H000000FF&
         Caption         =   "Show Canvas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaskColor       =   &H0080C0FF&
         TabIndex        =   41
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdGrid 
         Caption         =   "Hide Grid"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmdBackColor 
         Caption         =   "Canvas Color"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdGridColor 
         Caption         =   "Grid Color"
         Height          =   255
         Left            =   360
         TabIndex        =   56
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdpzHide 
         Caption         =   "Hide behind xy"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton tD 
         Caption         =   "3D OFF"
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton cmdPoints 
         Caption         =   "Show Points"
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Load"
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdFillPoint 
         Caption         =   "Hollow"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton cmdPerspective 
         Caption         =   "Perspective OFF"
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   600
         Width           =   1575
      End
      Begin VB.HScrollBar scrlPerspective 
         Enabled         =   0   'False
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   20
         TabIndex        =   46
         Top             =   4080
         Value           =   10
         Width           =   1095
      End
      Begin VB.HScrollBar scrl3DAngle 
         Enabled         =   0   'False
         Height          =   255
         LargeChange     =   6
         Left            =   120
         Max             =   360
         TabIndex        =   49
         Top             =   4800
         Value           =   6
         Width           =   1095
      End
      Begin VB.Label lblPersp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10"
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
         Height          =   255
         Left            =   1200
         TabIndex        =   51
         Top             =   4080
         Width           =   495
      End
      Begin VB.Label lblLeft 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "6"
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
         Height          =   255
         Left            =   1200
         TabIndex        =   50
         Top             =   4800
         Width           =   495
      End
   End
   Begin VB.PictureBox picScroll 
      BackColor       =   &H8000000C&
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4515
      ScaleWidth      =   4035
      TabIndex        =   1
      Top             =   480
      Width           =   4095
      Begin VB.PictureBox picTr 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1635
         Index           =   0
         Left            =   0
         ScaleHeight     =   1605
         ScaleWidth      =   4005
         TabIndex        =   24
         Top             =   0
         Width           =   4035
         Begin VB.HScrollBar scrlScale 
            Height          =   255
            Index           =   2
            LargeChange     =   10
            Left            =   720
            Max             =   10000
            TabIndex        =   28
            Top             =   900
            Value           =   1500
            Width           =   2475
         End
         Begin VB.HScrollBar scrlScale 
            Height          =   255
            Index           =   1
            LargeChange     =   10
            Left            =   720
            Max             =   10000
            TabIndex        =   27
            Top             =   615
            Value           =   1500
            Width           =   2475
         End
         Begin VB.HScrollBar scrlScale 
            Height          =   255
            Index           =   0
            LargeChange     =   10
            Left            =   720
            Max             =   10000
            TabIndex        =   26
            Top             =   330
            Value           =   1500
            Width           =   2475
         End
         Begin VB.HScrollBar scrlScale 
            Height          =   255
            Index           =   3
            LargeChange     =   10
            Left            =   720
            Max             =   10000
            TabIndex        =   25
            Top             =   1185
            Value           =   1500
            Width           =   2475
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "x-axis"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   330
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "y-axis"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   615
            Width           =   495
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "z-axis"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   900
            Width           =   495
         End
         Begin VB.Label lblScale 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Index           =   2
            Left            =   3255
            TabIndex        =   34
            Top             =   885
            Width           =   675
         End
         Begin VB.Label lblScale 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Index           =   1
            Left            =   3255
            TabIndex        =   33
            Top             =   600
            Width           =   675
         End
         Begin VB.Label lblScale 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   3255
            TabIndex        =   32
            Top             =   315
            Width           =   675
         End
         Begin VB.Label lblPicTr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "SCALING"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   1545
            TabIndex        =   31
            Top             =   -15
            Width           =   915
         End
         Begin VB.Label lblScale 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Index           =   3
            Left            =   3255
            TabIndex        =   30
            Top             =   1170
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "x,y,z"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1185
            Width           =   405
         End
      End
      Begin VB.PictureBox picTr 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1350
         Index           =   1
         Left            =   90
         ScaleHeight     =   1320
         ScaleWidth      =   4005
         TabIndex        =   13
         Top             =   585
         Width           =   4035
         Begin VB.HScrollBar scrlRotate 
            Height          =   255
            Index           =   0
            LargeChange     =   157
            Left            =   720
            Max             =   628
            TabIndex        =   16
            Top             =   330
            Width           =   2475
         End
         Begin VB.HScrollBar scrlRotate 
            Height          =   255
            Index           =   1
            LargeChange     =   157
            Left            =   720
            Max             =   628
            TabIndex        =   15
            Top             =   615
            Width           =   2475
         End
         Begin VB.HScrollBar scrlRotate 
            Height          =   255
            Index           =   2
            LargeChange     =   157
            Left            =   720
            Max             =   628
            TabIndex        =   14
            Top             =   900
            Width           =   2475
         End
         Begin VB.Label lblPicTr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ROTATION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   1455
            TabIndex        =   23
            Top             =   -15
            Width           =   1095
         End
         Begin VB.Label lblRotate 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   3255
            TabIndex        =   22
            Top             =   315
            Width           =   735
         End
         Begin VB.Label lblRotate 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Index           =   1
            Left            =   3255
            TabIndex        =   21
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lblRotate 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Index           =   2
            Left            =   3255
            TabIndex        =   20
            Top             =   885
            Width           =   735
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "z-axis"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   900
            Width           =   495
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "y-axis"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   615
            Width           =   495
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "x-axis"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   330
            Width           =   495
         End
      End
      Begin VB.PictureBox picTr 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1350
         Index           =   2
         Left            =   180
         ScaleHeight     =   1320
         ScaleWidth      =   4005
         TabIndex        =   2
         Top             =   1860
         Width           =   4035
         Begin VB.HScrollBar scrlTranslate 
            Height          =   255
            Index           =   2
            LargeChange     =   10
            Left            =   720
            Max             =   5000
            TabIndex        =   5
            Top             =   900
            Value           =   2500
            Width           =   2475
         End
         Begin VB.HScrollBar scrlTranslate 
            Height          =   255
            Index           =   1
            LargeChange     =   10
            Left            =   720
            Max             =   5000
            TabIndex        =   4
            Top             =   615
            Value           =   2500
            Width           =   2475
         End
         Begin VB.HScrollBar scrlTranslate 
            Height          =   255
            Index           =   0
            LargeChange     =   10
            Left            =   720
            Max             =   5000
            TabIndex        =   3
            Top             =   330
            Value           =   2500
            Width           =   2475
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "x-axis"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   330
            Width           =   495
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "y-axis"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   615
            Width           =   495
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "z-axis"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   900
            Width           =   495
         End
         Begin VB.Label lblTranslate 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Index           =   2
            Left            =   3255
            TabIndex        =   9
            Top             =   885
            Width           =   735
         End
         Begin VB.Label lblTranslate 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Index           =   1
            Left            =   3255
            TabIndex        =   8
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lblTranslate 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   3255
            TabIndex        =   7
            Top             =   315
            Width           =   735
         End
         Begin VB.Label lblPicTr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TRANSLATION"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   1275
            TabIndex        =   6
            Top             =   -15
            Width           =   1455
         End
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3720
      Top             =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CONTROLS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   4560
      TabIndex        =   52
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TRANSFORMATIONS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   1095
      TabIndex        =   0
      Top             =   120
      Width           =   2130
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Author: Hrayr Artunyan
' Date:   May 20, 2004
' You may use this software any way you wish but please give credit to me as well.
' If you have any questions or comments please email me at hrayr@artunyan.com or
' you may send me feedback from my website at www.artunyan.com/feedback.htm
' You may also find this software at www.artunyan.com
' If you make any modifications, please send me the code so that I can post it on my website
' and give you appropriate credit for it

Dim oldS, degS, indx, d
Dim picTrHeight(1000)
Dim picTrOpen(1000) As Boolean

Private Sub cmdAnimate_Click()
    If cmdAnimate.Caption = "Animate" Then
        cmdAnimate.Caption = "Pause"
        frmCanvas.Timer1.Enabled = True
    Else
        cmdAnimate.Caption = "Animate"
        frmCanvas.Timer1.Enabled = False
    End If
End Sub

Private Sub cmdAdvanced_Click()
    If cmdAdvanced.Caption = "SHOW  ADVANCED   CONTROLS" Then
        Me.Width = picAdvanced.Left + picAdvanced.Width + picScroll.Left * 2
        cmdAdvanced.Caption = "HIDE  ADVANCED  CONTROLS"
        picAdvanced.SetFocus
    Else
        Me.Width = cmdAdvanced.Left + cmdAdvanced.Width + 85
        cmdAdvanced.Caption = "SHOW  ADVANCED   CONTROLS"
        picAdvanced.SetFocus
    End If
End Sub

Private Sub cmdBackColor_Click()
On Error GoTo ERRCANCEL
    cd.Color = frmCanvas.picCanvas.BackColor
    cd.ShowColor
        frmCanvas.picCanvas.BackColor = cd.Color
        frmCanvas.BackColor = cd.Color
        frmCanvas.redraw
ERRCANCEL:
    Exit Sub
End Sub

Private Sub cmdFillPoint_Click()
    If cmdFillPoint.Caption = "Hollow" Then
        frmCanvas.setpFill False
        cmdFillPoint.Caption = "Solid"
    Else
        frmCanvas.setpFill True
        cmdFillPoint.Caption = "Hollow"
    End If
    frmCanvas.redraw
End Sub

Private Sub cmdGrid_Click()
    If cmdGrid.Caption = "Hide Grid" Then
        cmdGrid.Caption = "Show Grid"
        frmCanvas.setShowGrid False
    Else
        cmdGrid.Caption = "Hide Grid"
        frmCanvas.setShowGrid True
    End If
    frmCanvas.redraw
End Sub

Private Sub cmdGridColor_Click()
On Error GoTo ERRCANCEL
    cd.Color = RGB(255, 255, 210)
    cd.ShowColor
        frmCanvas.setGridColor cd.Color
        frmCanvas.redraw
ERRCANCEL:
    Exit Sub
End Sub

Private Sub cmdLines_Click()
    If cmdLines.Caption = "Show Lines" Then
        frmCanvas.setLines True
        cmdLines.Caption = "Hide Lines"
    Else
        frmCanvas.setLines False
        cmdLines.Caption = "Show Lines"
    End If
    frmCanvas.redraw
End Sub

Private Sub cmdOpen_Click()
On Error GoTo ERRCANCEL
    cd.DefaultExt = ".txt"
    cd.InitDir = App.Path
    cd.ShowOpen
    If Not cd.FileName = "" Then
        frmCanvas.initFile (cd.FileName)
    End If
    frmCanvas.redraw
ERRCANCEL:
    Exit Sub
End Sub

Private Sub cmdPerspective_Click()
    If cmdPerspective.Caption = "Perspective OFF" Then
        cmdPerspective.Caption = "Perspective ON"
        scrlPerspective.Enabled = True
        frmCanvas.setPerspective scrlPerspective.Value
        frmCanvas.setPersp True
    Else
        cmdPerspective.Caption = "Perspective OFF"
        scrlPerspective.Enabled = False
        frmCanvas.setPersp False
    End If
    frmCanvas.redraw
End Sub

Private Sub cmdPoints_Click()
    frmCanvas.showPoints True
    If cmdPoints.Caption = "Show Points" Then
        cmdPoints.Caption = "Hide Points"
        frmCanvas.showPoints True
        frmCanvas.redraw
    Else
        cmdPoints.Caption = "Show Points"
        frmCanvas.showPoints False
        frmCanvas.redraw
    End If
End Sub

Private Sub cmdpzHide_Click()
    If cmdpzHide.Caption = "Hide behind xy" Then
        cmdpzHide.Caption = "Show behind xy"
        frmCanvas.setpzHide True
    Else
        frmCanvas.setpzHide False
        cmdpzHide.Caption = "Hide behind xy"
    End If
End Sub

Private Sub cmdReset_Click()
    'Reset Rotateion scroll bars and labels
    For i = 0 To scrlRotate.Count - 1
        scrlRotate(i).Value = 0
        lblRotate(i).Caption = "0.00" & degS
    Next i
    For i = 0 To scrlScale.Count - 1
        scrlScale(i).Value = scrlScale(i).Max / 2
        lblScale(i).Caption = "100%"
    Next i
    For i = 0 To scrlTranslate.Count - 1
        scrlTranslate(i).Value = scrlTranslate(i).Max / 2
        lblTranslate(i).Caption = "0"
    Next i
    frmCanvas.resetCanvas
    frmCanvas.redraw
End Sub


Private Sub cmdShowCanvas_Click()
    If cmdShowCanvas.Caption = "Show Canvas" Then
        cmdShowCanvas.Caption = "Hide Canvas"
        frmCanvas.Show
    Else
        cmdShowCanvas.Caption = "Show Canvas"
        frmCanvas.Hide
    End If
End Sub

Private Sub Command1_Click()
cmdReset_Click
End Sub

Private Sub Command2_Click()
frmAbout.Show 1, Me
End Sub

Private Sub Form_Load()

    Me.Width = cmdAdvanced.Left + cmdAdvanced.Width + 85
    picTrHeight(0) = picTr(0).Height
    picTr(0).Height = lblPicTr(0).Height

    For i = 1 To picTr.Count - 1
        picTr(i).Top = picTr(i - 1).Top + lblPicTr(i - 1).Height - 15
        picTr(i).Left = picTr(0).Left
        picTr(i).Width = picTr(0).Width
'        lblPicTr(i).Width = lblPicTr(0).Width
        picTrHeight(i) = picTr(i).Height
        picTr(i).Height = lblPicTr(i).Height
    Next i
    
    degS = Chr(176)         ' degree symbol
    oldS = 1
    pi = 3.14159265358979
    oldX = oldY = oldZ = 0

    lblLeft.Caption = scrl3DAngle.Value & degS
    cmdReset_Click
    cmdShowCanvas_Click
    cmdLines_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmCanvas
End Sub

Private Sub HScroll1_Change()
frmCanvas.setPerspective (HScroll1.Value)
lblPersp.Caption = HScroll1.Value
frmCanvas.redraw
End Sub

Private Sub HScroll1_Scroll()
HScroll1_Change
End Sub

Private Sub lblPicTr_Click(Index As Integer)

    indx = Index
    d = lblPicTr(Index).Height
    If Not picTrOpen(Index) Then
        picTrOpen(Index) = True
        Timer1.Enabled = True
    Else
        picTrOpen(Index) = False
        Timer1.Enabled = True
    End If
    For Each i In lblPicTr
        If picTrOpen(i.Index) Then
            i.BackColor = RGB(255, 150, 150)
        Else
            'i.BackColor = RGB(150, 150, 255)
            i.BackColor = RGB(255, 255, 255)
        End If
    Next i
End Sub

Private Sub lblPicTr_DblClick(Index As Integer)
lblPicTr_Click (Index)
End Sub

Private Sub scrl3DAngle_Change()
    frmCanvas.set3D_Angle (scrl3DAngle.Value)
    lblLeft.Caption = scrl3DAngle.Value & degS
    frmCanvas.redraw
End Sub

Private Sub scrl3DAngle_Scroll()
    scrl3DAngle_Change
End Sub

Private Sub scrlPerspective_Change()
    frmCanvas.setPerspective scrlPerspective.Value
    lblPersp.Caption = scrlPerspective.Value
    frmCanvas.redraw
End Sub

Private Sub scrlPerspective_Scroll()
    scrlPerspective_Change
End Sub

Private Sub scrlRotate_Change(Index As Integer)
    Static oldX, oldY, oldZ
    Dim rad As Double, deg As Double
    
    rad = (scrlRotate(Index)) / 100
    deg = (rad * 180) / pi
    Select Case Index
        Case 0:
            frmCanvas.MatrixRotateMa rad - oldX, 0, 0
            oldX = rad
        Case 1:
            frmCanvas.MatrixRotateMa 0, rad - oldY, 0
            oldY = rad
        Case 2:
            frmCanvas.MatrixRotateMa 0, 0, rad - oldZ
            oldZ = rad
    End Select
    lblRotate(Index).Caption = Format(deg, "0.00" & degS)
    frmCanvas.redraw
End Sub

Private Sub scrlRotate_Scroll(Index As Integer)
    scrlRotate_Change (Index)
End Sub

Private Sub scrlScale_Change(Index As Integer)
    'frmCanvas.MatrixScale1 scrlScale(0).Value / 1000, scrlScale(1).Value / 1000, scrlScale(2).Value / 1000
    'frmCanvas.redraw
    
   Static oldX, oldY, oldZ, oldS
   Dim scl As Double, scll As Double
'
    scl = (scrlScale(Index).Value - scrlScale(Index).Max / 2) / scrlScale(Index).Max

    lblScale(Index).Caption = Format(scl + 1, "0.0%")
    Select Case Index
    Case 0:
        scll = scl - oldX
        If (scll < 0) Then
            frmCanvas.MatrixScale1 1 / (1 - scll), 1, 1
        Else: frmCanvas.MatrixScale1 1 + scll, 1, 1
        End If
        oldX = scl
    Case 1:
        scll = scl - oldY
        If (scll < 0) Then
            frmCanvas.MatrixScale1 1, 1 / (1 - scll), 1
        Else: frmCanvas.MatrixScale1 1, 1 + scll, 1
        End If
        oldY = scl
    Case 2:
        scll = scl - oldZ
        If (scll < 0) Then
            frmCanvas.MatrixScale1 1, 1, 1 / (1 - scll)
        Else: frmCanvas.MatrixScale1 1, 1, 1 + scll
        End If
        oldZ = scl
    Case 3:
        scll = scl - oldS
        If (scll < 0) Then
            frmCanvas.MatrixScale1 1 / (1 - scll), 1 / (1 - scll), 1 / (1 - scll)
        Else: frmCanvas.MatrixScale1 1 + scll, 1 + scll, 1 + scll
        End If
        oldS = scl
    End Select
    frmCanvas.redraw
End Sub

Private Sub scrlScale_Scroll(Index As Integer)
    scrlScale_Change Index
End Sub

Private Sub scrlTranslate_Change(Index As Integer)
    Static oldX, oldY, oldZ

    scl = (scrlTranslate(Index).Value - scrlTranslate(Index).Max / 2) / 500
    Select Case Index
        Case 0:
            scll = scl - oldX
                frmCanvas.mtrxTranslate1 scll * 1#, 0, 0
            oldX = scl
        Case 1:
            scll = scl - oldY
                frmCanvas.mtrxTranslate1 0, scll * 1#, 0
            oldY = scl
        Case 2:
            scll = scl - oldZ
                frmCanvas.mtrxTranslate1 0, 0, scll * 1#
            oldZ = scl
    End Select
    lblTranslate(Index).Caption = scl
    frmCanvas.redraw
End Sub

Private Sub scrlTranslate_Scroll(Index As Integer)
    scrlTranslate_Change Index
End Sub

Private Sub tD_Click()
    If tD.Caption = "3D OFF" Then
        tD.Caption = "3D ON"
        frmCanvas.set3D True
        scrl3DAngle.Enabled = True
        frmCanvas.redraw
    Else
        tD.Caption = "3D OFF"
        frmCanvas.set3D False
        scrl3DAngle.Enabled = False
        frmCanvas.redraw
    End If
End Sub

Private Sub Timer1_Timer()
    f = 60
    If d <= picTrHeight(indx) Then
        If picTrOpen(indx) = True Then
            picTr(indx).Height = picTr(indx).Height + f
            For i = indx + 1 To picTr.Count - 1
                picTr(i).Top = picTr(i).Top + f
            Next i
        Else
            picTr(indx).Height = picTr(indx).Height - f
            For i = indx + 1 To picTr.Count - 1
                picTr(i).Top = picTr(i).Top - f
            Next i
        End If
    '    Me.Height = picTr(picTr.Count - 1).Top + picTr(picTr.Count - 1).Height + 600
        d = d + f
    Else
        Timer1.Enabled = False
    End If
End Sub
