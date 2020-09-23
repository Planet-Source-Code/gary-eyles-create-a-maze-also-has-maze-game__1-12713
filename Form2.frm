VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Maze"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3900
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Solve"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Redraw"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3735
      Left            =   0
      ScaleHeight     =   3675
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   600
      Width           =   3855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FirstLoad As Boolean

Private Const mRight = 1
Private Const mLeft = 2
Private Const mUp = 3
Private Const mDown = 4

Dim TmpMaze() As Boolean
Dim CancelSolve As Boolean

Public Sub RedrawTheMaze()
Dim xSize As Integer
Dim ySize As Integer
Dim txSize As Integer
Dim tySize As Integer

txSize = MazeX
tySize = MazeY

If (txSize Mod 2) Then
    xSize = txSize
Else
    xSize = txSize + 1
End If
    
If (tySize Mod 2) Then
    ySize = tySize
Else
    ySize = tySize + 1
End If

Dim xP As Long
Dim yP As Long
Dim xAdd As Long
Dim yAdd As Long

xP = Int(Picture1.ScaleWidth / xSize)
yP = Int(Picture1.ScaleHeight / ySize)
xAdd = Int(((Picture1.ScaleWidth) - (xP * xSize)) / 2)
yAdd = Int(((Picture1.ScaleHeight) - (yP * ySize)) / 2)

Dim Wall As Boolean

Dim xx As Long
Dim yy As Long

Dim xd As Long
Dim yd As Long

Picture1.BackColor = QBColor(0)
Picture1.Cls
For xx = 0 To xSize - 1
    For yy = 0 To ySize - 1
        Wall = Maze.GetMaze(xx, yy)
        If Not Wall Then
            If Not GetMaze(xx - 1, yy) Then
                xd = 15
            Else
                xd = 0
            End If
            If Not GetMaze(xx, yy - 1) Then
                yd = 15
            Else
                yd = 0
            End If

            Picture1.Line (xP * xx + xAdd + 15 - xd, yP * yy + yAdd + 15 - yd)-(xP * xx + xAdd + xP - 15, yP * yy + yAdd + yP - 15), QBColor(15), BF
        End If
    Next
Next

Picture1.Line (xP * xC + xAdd + 15 - xd + 15, yP * yC + yAdd + 15 - yd)-(xP * xC + xAdd + xP - 15, yP * yC + yAdd + yP), QBColor(12), BF

Picture1.Refresh
End Sub

Private Sub Command1_Click()
Call RedrawTheMaze
End Sub

Private Sub Command2_Click()
Dim xSize As Integer
Dim ySize As Integer
Dim txSize As Integer
Dim tySize As Integer

txSize = MazeX
tySize = MazeY

If (txSize Mod 2) Then
    xSize = txSize
Else
    xSize = txSize + 1
End If
    
If (tySize Mod 2) Then
    ySize = tySize
Else
    ySize = tySize + 1
End If


ReDim TmpMaze(xSize + 1, ySize + 1)

Dim xP As Long
Dim yP As Long
Dim xAdd As Long
Dim yAdd As Long

For xP = 0 To xSize - 1
    For yP = 0 To ySize - 1
        TmpMaze(xP, yP) = GetMaze(xP, yP)
    Next
Next

xP = Int(Picture1.ScaleWidth / xSize)
yP = Int(Picture1.ScaleHeight / ySize)
xAdd = Int(((Picture1.ScaleWidth) - (xP * xSize)) / 2)
yAdd = Int(((Picture1.ScaleHeight) - (yP * ySize)) / 2)

'Picture1.Line (xP + xAdd, yP + yAdd)-(xP * 2 + xAdd, yP * 2 + yAdd), QBColor(10), BF
Picture1.Line ((xP * (xSize - 1)) + xAdd - 15, _
    yP * (ySize - 2) + yAdd)- _
    ((xP * (xSize - 2)) + xAdd + 15, _
    yP * (ySize - 1) + yAdd - 15), _
    QBColor(13), BF

Dim sX As Long
Dim sY As Long
Dim Moving As Integer
Dim Tim As Long
Dim OldMove As Integer

sX = 1: sY = 1

If GetMaze(sX + 1, sY) Then
    Moving = mDown
Else
    Moving = mRight
End If

Dim aColor As Long

aColor = QBColor(13)

Do
CarryOn:
    
If CancelSolve Then Exit Do
    
If Moving = mLeft Then
    sX = sX - 1
ElseIf Moving = mRight Then
    sX = sX + 1
ElseIf Moving = mUp Then
    sY = sY - 1
ElseIf Moving = mDown Then
    sY = sY + 1
End If
    
If TmpMaze(sX, sY) Then
    aColor = QBColor(15)
Else
    aColor = QBColor(13)
End If
      
If Moving = mLeft Then
    sX = sX + 1
ElseIf Moving = mRight Then
    sX = sX - 1
ElseIf Moving = mUp Then
    sY = sY + 1
ElseIf Moving = mDown Then
    sY = sY - 1
End If
      
If sX = xSize - 2 And sY = ySize - 2 Then Exit Do
    
    If Moving = mRight Then
'            sX = sX + 1

Picture1.Line (xP * sX + xAdd + 15, _
(yP * sY + yAdd) + 15)- _
(xP * sX + xAdd + xP, _
(yP * sY + yAdd) + yP - 15), _
aColor, BF
            
            TmpMaze(sX, sY) = True
            sX = sX + 1
    End If

    If Moving = mDown Then
'            sY = sY + 1

Picture1.Line ((xP * sX) + xAdd + 15, _
(yP * sY + yAdd) + 15)- _
((xP * sX) + xAdd + xP - 15, _
(yP * sY + yAdd) + yP), _
aColor, BF
            
            TmpMaze(sX, sY) = True
            sY = sY + 1
    End If

    If Moving = mUp Then
'            sY = sY - 1

Picture1.Line (xP * sX + xAdd + 15, _
(yP * sY + yAdd) - 15)- _
(xP * sX + xAdd + xP - 15, _
(yP * sY + yAdd) + yP - 15), _
aColor, BF
            
            TmpMaze(sX, sY) = True
            sY = sY - 1
    End If

    If Moving = mLeft Then
'            sX = sX - 1

Picture1.Line (xP * sX + xAdd, _
(yP * sY + yAdd) + 15)- _
(xP * sX + xAdd + xP - 15, _
(yP * sY + yAdd) + yP - 15), _
aColor, BF
            
            TmpMaze(sX, sY) = True
            sX = sX - 1
    End If

'If CBool(Check2.Value) Then
'    Picture1.Refresh
'    For Tim = 0 To 2000: DoEvents: Next
'End If
    
OldMove = Moving
    
    If Moving = mRight Then
        If Not GetMaze(sX, sY + 1) Then Moving = mDown: GoTo CarryOn
        If Not GetMaze(sX + 1, sY) Then Moving = mRight: GoTo CarryOn
        If Not GetMaze(sX, sY - 1) Then Moving = mUp: GoTo CarryOn
        Moving = mLeft: aColor = QBColor(14): GoTo CarryOn
    End If
    
    If Moving = mLeft Then
        If Not GetMaze(sX, sY - 1) Then Moving = mUp: GoTo CarryOn
        If Not GetMaze(sX - 1, sY) Then Moving = mLeft: GoTo CarryOn
        If Not GetMaze(sX, sY + 1) Then Moving = mDown: GoTo CarryOn
        Moving = mRight: aColor = QBColor(14): GoTo CarryOn
    End If
    
    If Moving = mUp Then
        If Not GetMaze(sX + 1, sY) Then Moving = mRight: GoTo CarryOn
        If Not GetMaze(sX, sY - 1) Then Moving = mUp: GoTo CarryOn
        If Not GetMaze(sX - 1, sY) Then Moving = mLeft: GoTo CarryOn
        Moving = mDown: aColor = QBColor(14): GoTo CarryOn
    End If
    
    If Moving = mDown Then
        If Not GetMaze(sX - 1, sY) Then Moving = mLeft: GoTo CarryOn
        If Not GetMaze(sX, sY + 1) Then Moving = mDown: GoTo CarryOn
        If Not GetMaze(sX + 1, sY) Then Moving = mRight: GoTo CarryOn
        Moving = mUp: aColor = QBColor(14): GoTo CarryOn
    End If

Loop

CancelSolve = False
Erase TmpMaze
End Sub

Private Sub Form_Resize()
On Error Resume Next

Picture1.Width = ScaleWidth
Picture1.Height = ScaleHeight - Picture1.Top

If Not FirstLoad Then
    Call RedrawTheMaze
    FirstLoad = True
End If
End Sub
