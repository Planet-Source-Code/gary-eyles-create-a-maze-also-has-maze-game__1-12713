VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Maze"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7605
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
   LinkTopic       =   "Form2"
   ScaleHeight     =   297
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   507
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "See it being solved"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Redraw"
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Solve"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "See it being drawn"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1800
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Left            =   1200
      TabIndex        =   3
      Text            =   "30"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   1200
      TabIndex        =   2
      Text            =   "30"
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Maze"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   3120
      ScaleHeight     =   4395
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Columns"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Rows"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mRight = 1
Private Const mLeft = 2
Private Const mUp = 3
Private Const mDown = 4

Dim TmpMaze() As Boolean
Dim CancelSolve As Boolean

Private Sub Command1_Click()
Picture1.Cls
Call DrawMaze(Picture1, Val(Combo1.Text), Val(Combo2.Text), , , CBool(Check1.Value))
End Sub

Private Sub Command2_Click()
Maze.CancelDrawing
CancelSolve = True
End Sub

Private Sub Command3_Click()
Dim xSize As Integer
Dim ySize As Integer
Dim txSize As Integer
Dim tySize As Integer

txSize = Val(Combo1.Text)
tySize = Val(Combo2.Text)

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

Picture1.Line (xP + xAdd, yP + yAdd)-(xP * 2 + xAdd, yP * 2 + yAdd), QBColor(10), BF
Picture1.Line (xP * (xSize - 1) + xAdd, yP * (ySize - 2) + yAdd)-(xP * (xSize - 2) + xAdd + 5, yP * (ySize - 1) + yAdd), QBColor(12), BF

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

If CBool(Check2.Value) Then
    Picture1.Refresh
    For Tim = 0 To 2000: DoEvents: Next
End If
    
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

Private Sub Command4_Click()
Dim xSize As Integer
Dim ySize As Integer
Dim txSize As Integer
Dim tySize As Integer

txSize = Val(Combo1.Text)
tySize = Val(Combo2.Text)

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

Picture1.Refresh
End Sub

Private Sub Form_Load()
Dim c As Integer

For c = 4 To 200
    Combo1.AddItem c
    Combo2.AddItem c
Next
End Sub

Private Sub Form_Resize()
Picture1.Top = 0
Picture1.Left = Combo1.Left + Combo1.Width + 10
Picture1.Height = ScaleHeight
Picture1.Width = ScaleWidth - Picture1.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
Maze.CancelDrawing
End Sub
