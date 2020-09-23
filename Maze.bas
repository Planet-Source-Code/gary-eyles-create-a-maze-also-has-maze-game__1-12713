Attribute VB_Name = "Maze"
Option Explicit

Dim xSize As Integer
Dim ySize As Integer

Dim Maze() As Boolean
Dim aMaze As New CBitArray

Dim xCellStack() As Long
Dim yCellStack() As Long

Dim tWallColor As Long
Dim tFloorColor As Long
Dim tDst
Dim tSeeDrawing As Boolean
Dim CancelDraw As Boolean

Dim xP As Long
Dim yP As Long
Dim xAdd As Long
Dim yAdd As Long

Dim d() As Integer

Function GetMaze(x As Long, y As Long) As Boolean
    'GetMaze = Maze(x, y)
    GetMaze = aMaze.Value((ySize * y) + x)

GetMaze = Not GetMaze
'    If GetMaze = True Then
'        GetMaze = False
'    ElseIf GetMaze = False Then
'        GetMaze = True
'    End If
End Function

Function MakeMaze(x As Integer, y As Integer) As Boolean
'On Error Resume Next

Dim tx As Integer
Dim ty As Integer

Dim p As Integer
Dim t As Long

StartAgain:
If tSeeDrawing Then
    For t = 0 To 1000
        DoEvents
        If CancelDraw Then Exit Function
    Next
End If

ReDim d(4)
p = 0
Maze(x, y) = True
aMaze.Value((CLng(y) * CLng(ySize)) + x) = True

If x > 1 Then If (Maze(x - 2, y) = False) Then d(p) = 1: p = p + 1
If y > 1 Then If (Maze(x, y - 2) = False) Then d(p) = 2: p = p + 1
If (x < (xSize - 2)) Then If (Maze(x + 2, y) = False) Then d(p) = 3: p = p + 1
If (y < (ySize - 2)) Then If (Maze(x, y + 2)) = False Then d(p) = 4: p = p + 1
    
If p = 0 And UBound(xCellStack) > 0 Then
    tx = xCellStack(UBound(xCellStack))
    ty = yCellStack(UBound(yCellStack))
    ReDim Preserve xCellStack(UBound(xCellStack) - 1)
    ReDim Preserve yCellStack(UBound(yCellStack) - 1)
    
    x = tx
    y = ty
    GoTo StartAgain
    'Call MakeMaze(tx, ty)
Else
    ReDim Preserve xCellStack(UBound(xCellStack) + 1)
    xCellStack(UBound(xCellStack)) = x
    ReDim Preserve yCellStack(UBound(yCellStack) + 1)
    yCellStack(UBound(yCellStack)) = y
End If
     
Randomize Timer
p = d(Int(p * Rnd))
    
If p = 1 Then
    Maze(x - 1, y) = True
    aMaze.Value(CLng(y * CLng(ySize)) + x - 1) = True
    
    'If tSeeDrawing Then tDst.Line (xP * (x - 2) + xAdd, yP * y + yAdd)-(xP * x + xAdd + xP, yP * y + yAdd + yP), tFloorColor, BF
    If tSeeDrawing Then tDst.Line (xP * (x - 2) + xAdd + 15, yP * y + yAdd + 15)-(xP * x + xAdd + xP - 15, yP * y + yAdd + yP - 15), tFloorColor, BF
    x = x - 2
    GoTo StartAgain
ElseIf p = 2 Then
    Maze(x, y - 1) = True
    aMaze.Value((CLng(y - 1) * CLng(ySize)) + x) = True
    
    'If tSeeDrawing Then tDst.Line (xP * x + xAdd, yP * (y - 2) + yAdd)-(xP * x + xAdd + xP, yP * y + yAdd + yP), tFloorColor, BF
    If tSeeDrawing Then tDst.Line (xP * x + xAdd + 15, yP * (y - 2) + yAdd + 15)-(xP * x + xAdd + xP - 15, yP * y + yAdd + yP - 15), tFloorColor, BF
    y = y - 2
    GoTo StartAgain
ElseIf p = 3 Then
    Maze(x + 1, y) = True
    aMaze.Value(CLng(y * CLng(ySize)) + x + 1) = True

    'If tSeeDrawing Then tDst.Line (xP * x + xAdd, yP * y + yAdd)-(xP * (x + 2) + xAdd + xP, yP * y + yAdd + yP), tFloorColor, BF
    If tSeeDrawing Then tDst.Line (xP * x + xAdd + 15, yP * y + yAdd + 15)-(xP * (x + 2) + xAdd + xP - 15, yP * y + yAdd + yP - 15), tFloorColor, BF
    x = x + 2
    GoTo StartAgain
ElseIf p = 4 Then
    Maze(x, y + 1) = True
    aMaze.Value((CLng(y + 1) * CLng(ySize)) + x) = True

    'If tSeeDrawing Then tDst.Line (xP * x + xAdd, yP * y + yAdd)-(xP * x + xAdd + xP, yP * (y + 2) + yAdd + yP), tFloorColor, BF
    If tSeeDrawing Then tDst.Line (xP * x + xAdd + 15, yP * y + yAdd + 15)-(xP * x + xAdd + xP - 15, yP * (y + 2) + yAdd + yP - 15), tFloorColor, BF
    y = y + 2
    GoTo StartAgain
End If
     
End Function

Sub DrawMaze(Dst, ByVal txSize As Integer, ByVal tySize As Integer, Optional WallColor As Long = vbBlack, Optional FloorColor As Long = vbWhite, Optional SeeDrawing As Boolean = False, Optional tDrawMaze As Boolean = True)
'On Error GoTo errCor

If (txSize Mod 2) Then
    xSize = txSize
Else
    xSize = txSize + 1
    txSize = txSize + 1
End If
    
If (tySize Mod 2) Then
    ySize = tySize
Else
    ySize = tySize + 1
    tySize = tySize + 1
End If

tFloorColor = FloorColor
tWallColor = WallColor
Set tDst = Dst
tSeeDrawing = SeeDrawing
CancelDraw = False

xP = Int(Dst.ScaleWidth / txSize)
yP = Int(Dst.ScaleHeight / tySize)
xAdd = Int(((Dst.ScaleWidth) - (xP * txSize)) / 2)
yAdd = Int(((Dst.ScaleHeight) - (yP * tySize)) / 2)

If tDrawMaze Then
    If tSeeDrawing Then
        Dst.Line (0, 0)-(Dst.ScaleWidth, Dst.ScaleHeight), WallColor, BF
    Else
        Dst.Line (0, 0)-(Dst.ScaleWidth, Dst.ScaleHeight), WallColor, BF
        Dst.Line (xAdd, yAdd)-(xP * xSize + xAdd, yP * ySize + yAdd), FloorColor, BF
    End If
End If

ReDim Maze(xSize, ySize)
Dim tmpN As Long
tmpN = (CLng(xSize + 1) * CLng(ySize + 1))
Call aMaze.RedimArray(tmpN, True)

ReDim xCellStack(0)
ReDim yCellStack(0)

Call MakeMaze(1, 1)

If Not CancelDraw And tDrawMaze = True Then
    Call ActualDraw
End If

Erase Maze
Erase xCellStack
Erase yCellStack

Exit Sub

errCor:
Debug.Print Err.Description & " " & tmpN
End Sub

Sub ActualDraw()
Dim x As Long
Dim y As Long

tDst.Line (0, 0)-(tDst.ScaleWidth, tDst.ScaleHeight), tWallColor, BF
tDst.Line (xAdd, yAdd)-(xP * xSize + xAdd, yP * ySize + yAdd), tFloorColor, BF

For y = 0 To ySize
    For x = 0 To xSize
        If Not Maze(x, y) Then
            tDst.Line (xP * x + xAdd, yP * y + yAdd)-(xP * x + xAdd + xP, yP * y + yAdd + yP), tWallColor, BF
        End If
    Next
Next

tDst.Refresh
End Sub

Sub CancelDrawing()
CancelDraw = True
End Sub
