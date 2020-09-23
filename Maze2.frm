VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Maze"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7590
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Maze2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   426
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   506
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pBuff 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FF0000&
      Height          =   1500
      Left            =   240
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start with"
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cheat"
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox FinishFloor 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   5160
      Picture         =   "Maze2.frx":030A
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox StartFloor 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   3480
      Picture         =   "Maze2.frx":11DE
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox Wall 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   1800
      Picture         =   "Maze2.frx":201D
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox Floor 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   240
      Picture         =   "Maze2.frx":2EBB
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   1
      Top             =   4680
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox TheMaze 
      BackColor       =   &H0000C000&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5775
      Left            =   0
      ScaleHeight     =   381
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   501
      TabIndex        =   0
      Top             =   600
      Width           =   7575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "="
      Height          =   240
      Left            =   3840
      TabIndex        =   10
      Top             =   240
      Width           =   105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "6x6"
      Height          =   240
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Maze Size"
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   930
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim mX As Long
Dim mY As Long

Dim FirstLoad As Boolean
Dim Winner As Boolean

Function DrawTheMaze(xad As Long, yad As Long) As Boolean
On Error Resume Next

Dim xx As Long
Dim yy As Long
Dim sX As Long
Dim sY As Long
Dim sxAdd As Long
Dim syAdd As Long
Dim xCentre As Long
Dim yCentre As Long

Dim OldX As Integer
Dim OldY As Integer

'xad = xad - 100
OldX = xC
OldY = yC

sX = Int((xad + 50) / 100)
sY = Int((yad + 50) / 100)

'If xC <> sX + 1 And yC <> sY + 1 Then
    xC = sX + 1
    yC = sY + 1

'End If

xCentre = xad - (sX * 100)
yCentre = yad - (sY * 100)

Debug.Print sX, sY
If GetMaze(sX + 1, sY + 1) Then
    xC = OldX
    yC = OldY
    DrawTheMaze = True
    Exit Function
End If

    If Form2.Visible Then
        Form2.RedrawTheMaze
    Else
        Unload Form2
    End If

If sX + 1 = MazeX - 1 And sY + 1 = MazeY - 1 Then
    Winner = True
    Exit Function
End If

xad = xad - Int(TheMaze.ScaleWidth / 2) + 150
yad = yad - Int(TheMaze.ScaleHeight / 2) + 150

sX = Int(xad / 100)
sY = Int(yad / 100)

sxAdd = xad - (sX * 100)
syAdd = yad - (sY * 100)

'If sX + Int(TheMaze.ScaleWidth / 100) > MazeX Then
    'sX = MazeX - Int(TheMaze.ScaleWidth / 100) + 1
'ElseIf sX < 0 Then
    'sX = 0
'End If
'If sY + Int(TheMaze.ScaleHeight / 100) > MazeY Then
    'sY = MazeY - Int(TheMaze.ScaleHeight / 100) + 1
'ElseIf sY < 0 Then
    'sY = 0
'End If

Dim PlyX As Long
Dim PlyY As Long

PlyX = (xC - sX) * Wall.ScaleWidth - sxAdd + 50 + xCentre
PlyY = (yC - sY) * Wall.ScaleHeight - syAdd + 50 + yCentre

For yy = sY To Int(TheMaze.ScaleHeight / 100) + 1 + sY
    For xx = sX To Int(TheMaze.ScaleWidth / 100) + 1 + sX
        If xx >= 0 And yy >= 0 And xx <= MazeX And yy <= MazeY Then
            If GetMaze(xx, yy) Then
                BitBlt TheMaze.hDC, (xx - sX) * Wall.ScaleWidth - sxAdd, (yy - sY) * Wall.ScaleHeight - syAdd, Wall.ScaleWidth, Wall.ScaleHeight, Wall.hDC, 0, 0, vbSrcCopy
            Else
                
            If PlyX > (xx - sX) * Wall.ScaleWidth - sxAdd And _
                        PlyX < (xx - sX) * Wall.ScaleWidth - sxAdd + Wall.ScaleWidth And _
                        PlyY > (yy - sY) * Wall.ScaleHeight - syAdd And _
                        PlyY < (yy - sY) * Wall.ScaleHeight - syAdd + Wall.ScaleHeight Then
                
                If xx = 1 And yy = 1 Then
'                    BitBlt TheMaze.hDC, (xx - sX) * Wall.ScaleWidth - sxAdd, (yy - sY) * Wall.ScaleHeight - syAdd, Wall.ScaleWidth, Wall.ScaleHeight, StartFloor.hDC, 0, 0, vbSrcCopy
                    BitBlt pBuff.hDC, 0, 0, Wall.ScaleWidth, Wall.ScaleHeight, StartFloor.hDC, 0, 0, vbSrcCopy
                    pBuff.Circle (50 + xCentre, 50 + yCentre), 10
                    BitBlt TheMaze.hDC, (xx - sX) * Wall.ScaleWidth - sxAdd, (yy - sY) * Wall.ScaleHeight - syAdd, Wall.ScaleWidth, Wall.ScaleHeight, pBuff.hDC, 0, 0, vbSrcCopy
                ElseIf xx = MazeX - 1 And yy = MazeY - 1 Then
'                    BitBlt TheMaze.hDC, (xx - sX) * Wall.ScaleWidth - sxAdd, (yy - sY) * Wall.ScaleHeight - syAdd, Wall.ScaleWidth, Wall.ScaleHeight, FinishFloor.hDC, 0, 0, vbSrcCopy
                    BitBlt pBuff.hDC, 0, 0, Wall.ScaleWidth, Wall.ScaleHeight, FinishFloor.hDC, 0, 0, vbSrcCopy
                    pBuff.Circle (50 + xCentre, 50 + yCentre), 10
                    BitBlt TheMaze.hDC, (xx - sX) * Wall.ScaleWidth - sxAdd, (yy - sY) * Wall.ScaleHeight - syAdd, Wall.ScaleWidth, Wall.ScaleHeight, pBuff.hDC, 0, 0, vbSrcCopy
                Else
                    BitBlt pBuff.hDC, 0, 0, Wall.ScaleWidth, Wall.ScaleHeight, Floor.hDC, 0, 0, vbSrcCopy
                    pBuff.Circle (50 + xCentre, 50 + yCentre), 10
                    BitBlt TheMaze.hDC, (xx - sX) * Wall.ScaleWidth - sxAdd, (yy - sY) * Wall.ScaleHeight - syAdd, Wall.ScaleWidth, Wall.ScaleHeight, pBuff.hDC, 0, 0, vbSrcCopy
                    'Else
'                       BitBlt TheMaze.hDC, (xx - sX) * Wall.ScaleWidth - sxAdd, (yy - sY) * Wall.ScaleHeight - syAdd, Wall.ScaleWidth, Wall.ScaleHeight, Floor.hDC, 0, 0, vbSrcCopy
                    'End If
                End If
            Else
                If xx = 1 And yy = 1 Then
                    BitBlt TheMaze.hDC, (xx - sX) * Wall.ScaleWidth - sxAdd, (yy - sY) * Wall.ScaleHeight - syAdd, Wall.ScaleWidth, Wall.ScaleHeight, StartFloor.hDC, 0, 0, vbSrcCopy
                ElseIf xx = MazeX - 1 And yy = MazeY - 1 Then
                    BitBlt TheMaze.hDC, (xx - sX) * Wall.ScaleWidth - sxAdd, (yy - sY) * Wall.ScaleHeight - syAdd, Wall.ScaleWidth, Wall.ScaleHeight, FinishFloor.hDC, 0, 0, vbSrcCopy
                Else
                    'BitBlt pBuff.hDC, 0, 0, Wall.ScaleWidth, Wall.ScaleHeight, Floor.hDC, 0, 0, vbSrcCopy
                    'pBuff.Circle (50 + xCentre, 50 + yCentre), 10
                    'BitBlt TheMaze.hDC, (xx - sX) * Wall.ScaleWidth - sxAdd, (yy - sY) * Wall.ScaleHeight - syAdd, Wall.ScaleWidth, Wall.ScaleHeight, pBuff.hDC, 0, 0, vbSrcCopy
                    'Else
                        BitBlt TheMaze.hDC, (xx - sX) * Wall.ScaleWidth - sxAdd, (yy - sY) * Wall.ScaleHeight - syAdd, Wall.ScaleWidth, Wall.ScaleHeight, Floor.hDC, 0, 0, vbSrcCopy
                    'End If
                End If
            End If
            End If
        Else
            'BitBlt TheMaze.hDC, (xx - sX) * Wall.ScaleWidth - sxAdd, (yy - sY) * Wall.ScaleHeight - syAdd, Wall.ScaleWidth, Wall.ScaleHeight, Wall.hDC, 0, 0, vbSrcCopy
            TheMaze.Line ((xx - sX) * Wall.ScaleWidth - sxAdd, (yy - sY) * Wall.ScaleHeight - syAdd)-((xx - sX) * Wall.ScaleWidth - sxAdd + 100, (yy - sY) * Wall.ScaleHeight - syAdd + 100), QBColor(3), BF
        End If
    Next
Next

TheMaze.Circle ((xC - sX) * Wall.ScaleWidth - sxAdd + 50 + xCentre, (yC - sY) * Wall.ScaleHeight - syAdd + 50 + yCentre), 10

xad = xad + Int(TheMaze.ScaleWidth / 2) - 150
yad = yad + Int(TheMaze.ScaleHeight / 2) - 150

'xad = xad + 100
'TheMaze.Refresh
End Function

Private Sub Command1_Click()
Form2.Show 0, Me
TheMaze.SetFocus
End Sub

Private Sub Command2_Click()
MazeX = Val(Mid(Combo1.Text, 1, InStr(1, Combo1.Text, "X", vbTextCompare) - 1))
MazeY = MazeX

Call ResetMaze

DrawMaze TheMaze, MazeX, MazeY, , , , False

Form_Resize
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 37 Then
    'Debug.Print "Left"
    mX = mX - 20
ElseIf KeyCode = 39 Then
    'Debug.Print "Right"
    mX = mX + 20
ElseIf KeyCode = 38 Then
    'Debug.Print "Up"
    mY = mY - 20
ElseIf KeyCode = 40 Then
    'Debug.Print "Down"
    mY = mY + 20
End If

If DrawTheMaze(mX, mY) Then

    If KeyCode = 37 Then
        'Debug.Print "Left"
        mX = mX + 20
    ElseIf KeyCode = 39 Then
        'Debug.Print "Right"
        mX = mX - 20
    ElseIf KeyCode = 38 Then
        'Debug.Print "Up"
        mY = mY + 20
    ElseIf KeyCode = 40 Then
        'Debug.Print "Down"
        mY = mY - 20
    End If

End If

If Winner Then
    MsgBox "Finished, time for a bigger maze"
    MazeX = MazeX + 2
    MazeY = MazeY + 2
    Maze.DrawMaze TheMaze, MazeX, MazeY, , , , False
       
    Call ResetMaze
       
    Form_Resize
    
    If Form2.Visible Then
        Form2.RedrawTheMaze
    Else
        Unload Form2
    End If
End If
'Debug.Print mX, mY
End Sub

Private Sub Form_Load()
Randomize Timer

Dim cc As Integer

cc = 2
Do
    cc = cc + 2
    Combo1.AddItem cc & " X " & cc
Loop Until cc = 200

Combo1.ListIndex = 0

MazeX = 4
MazeY = 4

Maze.DrawMaze TheMaze, MazeX, MazeY, , , , False

Label2.Caption = MazeX & " X " & MazeY

xC = 1
yC = 1
End Sub

Private Sub Form_Resize()
'TheMaze.Top = 0

If WindowState <> vbMinimized Then
    TheMaze.Left = 0
    TheMaze.Height = ScaleHeight - TheMaze.Top
    TheMaze.Width = ScaleWidth

    Call DrawTheMaze(mX, mY)
End If

If Not FirstLoad Then
    Call DrawTheMaze(mX, mY)
    FirstLoad = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Maze.CancelDrawing
End Sub

Private Sub TheMaze_Paint()
Call DrawTheMaze(mX, mY)
End Sub

Sub ResetMaze()
    mX = 0
    mY = 0
    xC = 1
    yC = 1
    Winner = False
    Label2.Caption = MazeX & " X " & MazeY
End Sub
