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
   LinkTopic       =   "Form2"
   ScaleHeight     =   297
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   507
   StartUpPosition =   2  'CenterScreen
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
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "See it being drawn"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   2040
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

Private Sub Command1_Click()
Picture1.Cls
Call DrawMaze(Picture1, Val(Combo1.Text), Val(Combo2.Text), , , CBool(Check1.Value))
End Sub

Private Sub Command2_Click()
Maze.CancelDrawing
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

Dim xP As Long
Dim yP As Long
Dim xAdd As Long
Dim yAdd As Long

xP = Int(Picture1.ScaleWidth / xSize)
yP = Int(Picture1.ScaleHeight / ySize)
xAdd = Int(((Picture1.ScaleWidth) - (xP * xSize)) / 2)
yAdd = Int(((Picture1.ScaleHeight) - (yP * ySize)) / 2)

Picture1.Line (xP + xAdd, yP + yAdd)-(xP * 2 + xAdd, yP * 2 + yAdd), QBColor(10), BF
Picture1.Line (xP * (xSize - 1) + xAdd, yP * (ySize - 2) + yAdd)-(xP * (xSize - 2) + xAdd + 5, yP * (ySize - 1) + yAdd), QBColor(12), BF

Dim sX As Integer
Dim sY As Integer
Dim Moving As Integer
Dim Tim As Long

sX = 1: sY = 1

If GetMaze(sX + 1, sY) Then
    Moving = mDown
Else
    Moving = mRight
End If

'Debug.Print IIf(Moving = mRight, "Right", "Down")

Picture1.DrawWidth = 2
Picture1.DrawMode = 7

Dim aColor As Long
'3,6,10
aColor = QBColor(10)

Do
CarryOn:
    
If sX = xSize - 2 And sY = ySize - 2 Then Exit Do
    
    If Moving = mRight Then
        'Do


'Picture1.Line (((sX) * xP) + xAdd + 15, (sY * yP) + yAdd + 15)- _
            (((sX) * xP) + xAdd + xP - 15, (sY * yP) + yAdd + 15), QBColor(4)

Picture1.Line (xP * sX + xAdd + Int(xP / 2), _
(yP * sY + yAdd) + Int(yP / 2))- _
(xP * (sX) + xAdd + Int(xP) + Int(xP / 2), _
(yP * sY + yAdd) + Int(yP / 2)), _
aColor, BF
            
            sX = sX + 1
        'Loop Until GetMaze(sX, sY)
        'sX = sX - 1
    End If

    If Moving = mDown Then
        'Do

Picture1.Line ((xP * sX) + xAdd + Int(xP / 2), _
(yP * sY + yAdd) + Int(yP / 2))- _
((xP * sX) + xAdd + Int(xP / 2), _
(yP * sY + yAdd) + yP + Int(yP / 2)), _
aColor, BF
            
            sY = sY + 1
        'Loop Until GetMaze(sX, sY)
        'sY = sY - 1
    End If

    If Moving = mUp Then
        'Do

Picture1.Line (xP * sX + xAdd + Int(xP / 2), _
(yP * sY + yAdd) + Int(yP / 2))- _
(xP * (sX) + xAdd + Int(xP / 2), _
(yP * sY + yAdd) - yP + Int(yP / 2)), _
aColor, BF
            
            sY = sY - 1
        'Loop Until GetMaze(sX, sY)
        'sY = sY + 1
    End If

    If Moving = mLeft Then
        'Do

Picture1.Line (xP * sX + xAdd + Int(xP / 2), _
(yP * sY + yAdd) + Int(yP / 2))- _
(xP * sX + xAdd - xP + Int(xP / 2), _
(yP * sY + yAdd) + Int(yP / 2)), _
aColor, BF
            
            sX = sX - 1
        'Loop Until GetMaze(sX, sY)
        'sX = sX + 1
    End If

    Picture1.Refresh
    'For Tim = 0 To 1000: DoEvents: Next
    
    If Moving = mRight Then
        If Not GetMaze(sX, sY + 1) Then Moving = mDown: GoTo CarryOn
        If Not GetMaze(sX + 1, sY) Then Moving = mRight: GoTo CarryOn
        If Not GetMaze(sX, sY - 1) Then Moving = mUp: GoTo CarryOn
        Moving = mLeft: GoTo CarryOn
    End If
    
    If Moving = mLeft Then
        If Not GetMaze(sX, sY - 1) Then Moving = mUp: GoTo CarryOn
        If Not GetMaze(sX - 1, sY) Then Moving = mLeft: GoTo CarryOn
        If Not GetMaze(sX, sY + 1) Then Moving = mDown: GoTo CarryOn
        Moving = mRight: GoTo CarryOn
    End If
    
    If Moving = mUp Then
        If Not GetMaze(sX + 1, sY) Then Moving = mRight: GoTo CarryOn
        If Not GetMaze(sX, sY - 1) Then Moving = mUp: GoTo CarryOn
        If Not GetMaze(sX - 1, sY) Then Moving = mLeft: GoTo CarryOn
        Moving = mDown: GoTo CarryOn
    End If
    
    If Moving = mDown Then
        If Not GetMaze(sX - 1, sY) Then Moving = mLeft: GoTo CarryOn
        If Not GetMaze(sX, sY + 1) Then Moving = mDown: GoTo CarryOn
        If Not GetMaze(sX + 1, sY) Then Moving = mRight: GoTo CarryOn
        Moving = mUp: GoTo CarryOn
    End If

Loop

Picture1.DrawWidth = 1
Picture1.DrawMode = 13
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
