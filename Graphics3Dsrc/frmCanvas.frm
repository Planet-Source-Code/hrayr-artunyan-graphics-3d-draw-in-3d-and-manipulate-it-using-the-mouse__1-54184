VERSION 5.00
Begin VB.Form frmCanvas 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "3D Graphics Form"
   ClientHeight    =   8670
   ClientLeft      =   120
   ClientTop       =   1530
   ClientWidth     =   8685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   90
      Top             =   2895
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3795
      Left            =   1560
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   367
      TabIndex        =   0
      Top             =   1320
      Width           =   5505
   End
End
Attribute VB_Name = "frmCanvas"
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

Dim orig() As Double
Dim cMR() As Double                         ' Coordinate Matrix
Dim cML() As Double                         ' Coordinate Matrix, used for 3D mode
Dim cDraw() As Integer

Dim midx As Double, midy As Double          ' mid values for the picturebox
Dim unitx As Double, unity As Double        ' unit length
Dim unitPx As Double, unitPy As Double      ' unit for perspective

Dim angl3D As Double                        ' angle difference of left and right eye images around y-axis
Dim threeD As Boolean                       ' used to tuggle between 3D and 2D views
Dim showP As Boolean                        ' used to tuggle between perspective and non-perspective views
Dim pzHide As Boolean
Dim pFill As Boolean
Dim persp As Boolean
Dim lns As Boolean
Dim showGrid As Boolean
Dim fileOpened As Boolean
Dim gridColor

Dim xx, yy, p                               ' old values to keep track of. used for mouse possition
Dim x1 As Double, y1 As Double              ' used for mouse possition

Public Sub MatrixCopy(ma() As Double, MB() As Double)
    ' copy matrix element by element.
    ' I don't use this function anymore but I left it just in case I need it.
    For i = 1 To UBound(ma, 1)
        For j = 1 To UBound(ma, 2)
            MB(i, j) = ma(i, j)
        Next j
    Next i
End Sub

Public Sub set3D_Angle(d)
    angl3D = d * pi / 180
End Sub

Public Sub set3D(t As Boolean)
    threeD = t
End Sub

Public Sub showPoints(sp As Boolean)
    showP = sp
End Sub

Public Sub setPerspective(perp As Double)
    p = perp
End Sub

Public Sub setpzHide(a As Boolean)
    pzHide = a
End Sub

Public Sub setpFill(a As Boolean)
    pFill = a
End Sub

Public Sub setPersp(p As Boolean)
    persp = p
End Sub

Public Sub setLines(l As Boolean)
    lns = l
End Sub

Public Sub setShowGrid(sg As Boolean)
    showGrid = sg
End Sub

Public Sub setGridColor(gc)
    gridColor = gc
End Sub

Public Sub resetCanvas()
    xx = 0
    yy = 0
    x1 = 0
    y1 = 0
    cMR = orig
    redraw
End Sub

Public Sub initFile(fn As String)
On Error GoTo ERRNOFILE
    Open fn For Input As #1
    Dim i  As Integer
    Do While Not EOF(1) And Not (ax = "a" And ay = "b" And az = "c")   ' Loop until end of file.
        Input #1, ax, ay, az ' Read data into two variables.
        If Not ax = "a" Then n = n + 1
    Loop
 
    Do While Not EOF(1)   ' Loop until end of file.
        Input #1, p1, p2  ' Read data into two variables.
        l = l + 1
    Loop
    Close #1
    
    ReDim cMR(3, n)        ' allocate size for right eye image
    ReDim cML(3, n)        ' allocate size for left eye image, to be used for 3D
    ReDim orig(3, n)
    ReDim cDraw(l, 2)
    
    Open fn For Input As #1
    Do
        Input #1, ax, ay, az ' Read data into two variables.
        If Not ax = "a" Then
            i = i + 1
            cMR(1, i) = Val(ax): cMR(2, i) = Val(ay): cMR(3, i) = Val(az)
        End If
    Loop While Not EOF(1) And Not (ax = "a" And ay = "b" And az = "c")  ' Loop until end of file.
    orig = cMR
    i = 0
    Do While Not EOF(1)   ' Loop until end of file.
        Input #1, p1, p2 ' Read data into two variables.
        i = i + 1
        cDraw(i, 1) = p1: cDraw(i, 2) = p2
    Loop
    Close #1
    fileOpened = True
    Exit Sub
ERRNOFILE:
    MsgBox "File could not be opened."
    createNullGraphic
    fileOpened = False
    Exit Sub
End Sub
Public Sub createNullGraphic()
    ReDim cMR(3, 0)        ' allocate size for right eye image
    ReDim cML(3, 0)        ' allocate size for left eye image, to be used for 3D
    ReDim orig(3, 0)
    ReDim cDraw(0, 0)
End Sub
Private Sub Form_Load()
    pzHide = False
    pFill = True
    threeD = False
    showGrid = True
    picCanvas.FillStyle = 0
    picCanvas.Appearance = 1
    picCanvas.DrawMode = 9                      ' this is critical for the 3D view.
                                                ' try changing the value to 13(default) you'll see what I mean
    pi = 3.14159265358979
    angl3D = 6 * pi / 180                       ' by default rotate the left eye image 6 degrees. used for 3D mode
    picCanvas.BackColor = RGB(255, 255, 215)    ' default background image
    Me.BackColor = picCanvas.BackColor
    gridColor = RGB(200, 200, 200)
    initFile ("pyramid.txt")                    ' read in the file and create coordinate matrix
    redraw
End Sub

Public Function MatrixMult(ByVal matrixA As Variant, ByVal matrixB As Variant) As Double()
    Dim temp() As Double
    Dim msum As Double
    
    ' get dimension of matricies
    m1 = UBound(matrixA, 1)
    n1 = UBound(matrixA, 2)
    m2 = UBound(matrixB, 1)
    n2 = UBound(matrixB, 2)
    
    ' oooooo wrong matrix dimension.
    If Not n1 = m2 Then
        MatrixMult = matrixB
        MsgBox "ERROR: Matrix dimensions are not valid for multiplication", vbCritical, "MatrixMult()"
        Exit Function
    End If
    
    ' create temp matrix to store the values in
    ReDim temp(m1, n2)
    
    ' multiply matricies and store in temp
    For m = 1 To m1
        For n = 1 To n2
            For k = 1 To n1
                msum = msum + matrixA(m, k) * matrixB(k, n)
            Next k
            temp(m, n) = msum
            msum = 0
        Next n
    Next m
    MatrixMult = temp
End Function

Public Sub redraw()
    DoEvents
    ' grid size.
    xn = 2
    yn = 2
    ' draw the canvas
    With picCanvas
        .Cls
        ' get the middle location in terms of the picture box. The grid is always centered
        midx = (.ScaleWidth - 1) / 2
        midy = (.ScaleHeight - 1) / 2
        ' calculate the unit length of x and y
        'unitx = (.ScaleWidth - 1) / (xn * 2)
        unity = (.ScaleHeight - 1) / (yn * 2)
        unitx = unity
        If showGrid Then
            ' color of the grid. (gray)
            .ForeColor = gridColor
            ' draw the vertical lines of the grid
            For i = 1 To Me.ScaleWidth / (unitx * 2)
                picCanvas.Line (midx + unitx * i, 0)-(midx + unitx * i, .ScaleHeight)
                picCanvas.Line (midx - unitx * i, 0)-(midx - unitx * i, .ScaleHeight)
            Next i
            ' draw the horizontal lines of the grid
            For i = 1 To yn
                picCanvas.Line (0, midy + unity * i)-(.ScaleWidth, midy + unity * i)
                picCanvas.Line (0, midy - unity * i)-(.ScaleWidth, midy - unity * i)
            Next i
            ' color of the axis (black)
            .ForeColor = RGB(150, 150, 150)
            picCanvas.Line (midx, 0)-(midx, .ScaleHeight)
            picCanvas.Line (0, midy)-(.ScaleWidth, midy)
        End If
        .ForeColor = RGB(0, 0, 0)
        .CurrentX = 5
        .CurrentY = .ScaleHeight - 45
        picCanvas.Print "by Hrayr Artunyan"
    End With

    ' draw 3D
    If threeD = True Then
        cML = cMR                               ' create an exact copy of the coordinate matrix
        MatrixRotate 0, -angl3D, 0, cML            ' rotate the copy matrix angl3D degrees where angl3D is decided at another location
        picCanvas.FillColor = RGB(255, 0, 0)
        picCanvas.ForeColor = RGB(255, 0, 0)    ' color the original matrix red, this is what the right eye sees
        Draw cDraw, cMR                         ' draw the original
        picCanvas.FillColor = RGB(0, 255, 255)
        picCanvas.ForeColor = RGB(0, 255, 255)  ' color the copy bluegreen, this is what the left eye sees
        Draw cDraw, cML                                ' draw the copy
    Else
        picCanvas.FillColor = RGB(0, 0, 0)
        picCanvas.ForeColor = RGB(0, 0, 0)      ' if not 3D draw only the original black
        Draw cDraw, cMR
    End If
End Sub


Private Sub Form_Resize()
    sz = 0
'***********************************
    ' don't keep aspect ratio.
    'picCanvas.Left = 0
    'picCanvas.Width = Me.ScaleWidth
    'picCanvas.Height = Me.Height
    'picCanvas.Top = 0
'***********************************

'***************************************************************************
    ' always keep the canvas' aspect ration to be a square
    
    ' hide the picturebox if it gets too small ( when the width is negative )
    If Me.ScaleHeight - sz < 0 Or Me.ScaleWidth - sz < 0 Then
        picCanvas.Visible = False
        Exit Sub
    Else
        picCanvas.Visible = True
    End If

    If Me.ScaleWidth > Me.ScaleHeight Then
        picCanvas.Width = Me.ScaleHeight - sz
        picCanvas.Height = Me.ScaleHeight - sz
    Else
        picCanvas.Width = Me.ScaleWidth - sz
        picCanvas.Height = Me.ScaleWidth - sz
    End If
    
   ' center the picture box horizontally on screen
    picCanvas.Left = Me.ScaleWidth / 2 - picCanvas.Width / 2
    picCanvas.Top = 0
'*******************************************************************************

    redraw
End Sub

Public Sub drawLine(X As Double, Y As Double, Z As Double, xx As Double, yy As Double, zz As Double)
    ' this function only takes in two coordinates, components, and draws a line connecting them.
    picCanvas.Line (midx + X * unitx, midy - Y * unity)-(midx + xx * unitx, midy - yy * unity)
End Sub

Public Sub drawLineP(p1 As Integer, p2 As Integer, mtrx() As Double)
    If persp Then
        ' this function takes a coordinate matrix and two column possitions and gets the coordinates from there
        z1 = mtrx(3, p1) * p        ' calculate perspective
        z2 = mtrx(3, p2) * p        ' I don't think this formula is really accurate but it does the job (somewhat).
    End If
    picCanvas.Line (midx + mtrx(1, p1) * (unitx + z1), midy - mtrx(2, p1) * (unity + z1)) _
                  -(midx + mtrx(1, p2) * (unitx + z2), midy - mtrx(2, p2) * (unity + z2))
End Sub

Public Sub drawPoint(X As Double, Y As Double, Z As Double)
    picCanvas.PSet (midx + X * unitx, midy - Y * unity)
End Sub

Public Sub drawCircle(X As Double, Y As Double, Z As Double, s)
    c = 1
    d = Z + c
    If (Z + c) < 0 Then d = 1 / (c - Z)
    
    ' this is a test
    'If Z < 0 Then d = -Z
    'If Z < 0 Then d = RGB(255, 0, 0) Else d = RGB(0, 0, 255)
    'picCanvas.Circle (midx + X * (unitx + Z * p), midy - Y * (unity + Z * p)), s, d
    
    's = s + s * d
    If pFill Then picCanvas.FillStyle = 0 Else picCanvas.FillStyle = 1
    If persp Then
        pr = Z * p
        s = s + s * d
    Else
        pr = 0
    End If
    picCanvas.Circle (midx + X * (unitx + pr), midy - Y * (unity + pr)), s
End Sub

Public Sub Draw(md() As Integer, ma() As Double)

    ' connect the dots. The numbers are column positions in the matrix ma
    If lns = True Then
        For i = 1 To UBound(cDraw, 1)
            drawLineP md(i, 1), md(i, 2), ma
        Next i
    End If
    ' draw circles around the points
    If showP Then
         For i = 1 To UBound(ma, 2)
            If pzHide Then
                If ma(3, i) > 0 Then     ' only show the points whos z component is positive (in front of the xy plane)
                    drawCircle ma(1, i), ma(2, i), ma(3, i), 2
                End If
            Else
                drawCircle ma(1, i), ma(2, i), ma(3, i), 2
            End If
        Next i
    End If
    
End Sub

Public Function MatrixScale1(X As Double, Y As Double, Z As Double)
    ' scale the cMR coordinate matrix
    MatrixScale X, Y, Z, cMR
End Function

Public Function MatrixScale(X As Double, Y As Double, Z As Double, ma() As Double)
    Dim mtrxS(3, 3)
    ' create scaleing matrix
    mtrxS(1, 1) = X: mtrxS(2, 1) = 0: mtrxS(3, 1) = 0
    mtrxS(1, 2) = 0: mtrxS(2, 2) = Y: mtrxS(3, 2) = 0
    mtrxS(1, 3) = 0: mtrxS(2, 3) = 0: mtrxS(3, 3) = Z
    Me.Caption = " "
    ma = MatrixMult(mtrxS, ma)

End Function

Public Sub MatrixRotateMa(X As Double, Y As Double, Z As Double)
    MatrixRotate X, Y, Z, cMR
End Sub

Public Sub MatrixRotate(X As Double, Y As Double, Z As Double, ma() As Double)
    Dim mtrxX(3, 3) As Double
    Dim mtrxY(3, 3) As Double
    Dim mtrxZ(3, 3) As Double
    
    ' create rotation matrix for x-axis
    mtrxX(1, 1) = 1: mtrxX(2, 1) = 0:      mtrxX(3, 1) = 0
    mtrxX(1, 2) = 0: mtrxX(2, 2) = Cos(X): mtrxX(3, 2) = -Sin(X)
    mtrxX(1, 3) = 0: mtrxX(2, 3) = Sin(X): mtrxX(3, 3) = Cos(X)

    ' create rotation matrix for y-axis
    mtrxY(1, 1) = Cos(Y):  mtrxY(2, 1) = 0: mtrxY(3, 1) = Sin(Y)
    mtrxY(1, 2) = 0:       mtrxY(2, 2) = 1: mtrxY(3, 2) = 0
    mtrxY(1, 3) = -Sin(Y): mtrxY(2, 3) = 0: mtrxY(3, 3) = Cos(Y)
    
    ' create rotation matrix for z-axis
    mtrxZ(1, 1) = Cos(Z): mtrxZ(2, 1) = -Sin(Z): mtrxZ(3, 1) = 0
    mtrxZ(1, 2) = Sin(Z): mtrxZ(2, 2) = Cos(Z):  mtrxZ(3, 2) = 0
    mtrxZ(1, 3) = 0:      mtrxZ(2, 3) = 0:       mtrxZ(3, 3) = 1

    ma = MatrixMult(mtrxX, ma)
    ma = MatrixMult(mtrxY, ma)
    ma = MatrixMult(mtrxZ, ma)
End Sub

Public Sub mtrxTranslate(X As Double, Y As Double, Z As Double, ma() As Double)
    ' translation is easy, just add the X, Y and Z values to each coresponding element
  '  MsgBox X & ", " & Y & ", " & Z
    For j = 1 To UBound(ma, 2)
        ma(1, j) = ma(1, j) + X
        ma(2, j) = ma(2, j) + Y
        ma(3, j) = ma(3, j) + Z
    Next j
End Sub

Public Sub mtrxTranslate1(X As Double, Y As Double, Z As Double)
    ' translation is easy, just add the X, Y and Z values to each coresponding element
    mtrxTranslate X, Y, Z, cMR

End Sub


Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' get coordinate of where the mouse was clicked and store for later use
    Timer1.Enabled = False
 
    xx = X
    yy = Y
    If Button = 1 Then
        y1 = (xx - X) / 100
        x1 = (yy - Y) / 100
    If Shift = 1 Then z1 = (xx - X) / 100
    End If
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    If Button = 1 Then              ' left click ( rotate )
        Timer1.Enabled = False      ' disable animation while still holding the mouse button
        If Shift = 1 Then
          z1 = (xx - X) / 100
          
        '  If y1 < 0.009 And y1 > -0.009 Then y1 = 0
        '  If x1 < 0.009 And x1 > -0.009 Then x1 = 0
          
          MatrixRotate 0, 0, (xx - X) / 100, cMR
          xx = X
          yy = Y
          redraw
        Else
          y1 = (xx - X) / 100
          x1 = (yy - Y) / 100
          
        '  If y1 < 0.009 And y1 > -0.009 Then y1 = 0
        '  If x1 < 0.009 And x1 > -0.009 Then x1 = 0
          
          MatrixRotate (yy - Y) / 100, (xx - X) / 100, 0, cMR
          xx = X
          yy = Y
          redraw
        End If
    ElseIf Button = 2 Then          ' right click ( x- and y-axis translate )
        Timer1.Enabled = False      ' disable animation while still holding the mouse button
        mtrxTranslate (X - xx) / unitx, (yy - Y) / unity, 0, cMR
        yy = Y
        xx = X
        redraw
    ElseIf Button = 3 Then          ' left and right click at the same time ( translate z-axis )
        Timer1.Enabled = False      ' disable animation while still holding the mouse button
        mtrxTranslate 0, 0, -((yy - Y) / unity), cMR
        yy = Y
        xx = X
        redraw
    Else
        Me.Caption = "Canvas: (" & Format((X - midx) / unitx, "0.000") & "," & Format((midy - Y) / unity, "0.000") & ")"
    End If
End Sub

Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If y1 < 0.01 And y1 > -0.01 Then y1 = 0
    If x1 < 0.01 And x1 > -0.01 Then x1 = 0
    If Not (x1 = 0 And y1 = 0 And z1) Then
        Timer1.Enabled = True
    End If
End Sub

Private Sub Timer1_Timer()
    If x1 < 0.001 And y1 < 0.001 Then Timer1.Enabled = False
    MatrixRotate x1, y1, 0, cMR
    redraw
End Sub

