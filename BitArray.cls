VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBitArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'BitArray - Bit Array Demo Program
'Copyright (c) 1997 SoftCircuits Programming (R)
'Redistributed by Permission.
'
'This Visual Basic 5.0 example program demonstrates a bit array
'class. A bit array is an array of yes/no, 0/1, or true/false values.
'It behaves like any other array but since it only stores a simple yes
'or no value, it is designed to only use a single bit for each value.
'
'For example, if you were to create an array of 100 Boolean values,
'that array would consume 200 bytes of memory (Boolean variables use
'two bytes of memory each). The bit array, however, can store the same
'100 values in only about 13 bytes resulting in tremendous memory
'savings, particularly if you need to store a very large number of
'items. The bit array class also takes advantage of data abstraction
'to make this functionality available using syntax that is virtually
'identical to the syntax used to access an ordinary array.
'
'The bit array class also includes routines to save the data to disk
'and read it back. Keeping with the spirit of object-oriented
'programming, disk access is handled internal to the class such that
'using the class does not require knowledge of how the data is stored.
'
'This program may be distributed on the condition that it is
'distributed in full and unchanged, and that no fee is charged for
'such distribution with the exception of reasonable shipping and media
'charged. In addition, the code in this program may be incorporated
'into your own programs and the resulting programs may be distributed
'without payment of royalties.
'
'This example program was provided by:
' SoftCircuits Programming
' http://www.softcircuits.com
' P.O. Box 16262
' Irvine, CA 92623
Option Explicit

'Private dynamic array to hold data
Private BitArray() As Integer

'Resizes the bit array to hold the number of
'bits specified by nCount
Public Function RedimArray(nCount As Long, Optional ClearUp As Boolean = False) As Long
On Error GoTo errCor
Dim TotalCount As Long

TotalCount = CLng(nCount \ 16)
RedimArray = TotalCount

If ClearUp Then
    ReDim BitArray(0 To TotalCount)
Else
    ReDim Preserve BitArray(0 To TotalCount)
End If

Exit Function

errCor:
Debug.Print "Error " & Err.Description & " " & TotalCount
End Function

'Gets the value of the bit specified by nIndex
Property Get Value(nIndex As Long) As Boolean
    Dim nBit As Long

    nBit = 2 ^ (nIndex Mod 16)
    If nBit = &H8000& Then  'Prevent overflow on high bit
        Value = BitArray(nIndex \ 16) And &H8000
    Else
        Value = BitArray(nIndex \ 16) And nBit
    End If
End Property

'Sets the value of the bit specified by nIndex
Property Let Value(nIndex As Long, bValue As Boolean)
    Dim nBit As Long

    nBit = 2 ^ (CLng(nIndex) Mod 16)
    If bValue Then
        If nBit = &H8000& Then  'Prevent overflow on high bit
            BitArray(CLng(nIndex) \ 16) = (BitArray(CLng(nIndex) \ 16) Or &H8000)
        Else
            BitArray(CLng(nIndex) \ 16) = (BitArray(CLng(nIndex) \ 16) Or nBit)
        End If
    Else
        If nBit = &H8000& Then  'Prevent overflow on high bit
            BitArray(CLng(nIndex) \ 16) = (BitArray(CLng(nIndex) \ 16) And Not &H8000)
        Else
            BitArray(CLng(nIndex) \ 16) = (BitArray(CLng(nIndex) \ 16) And Not nBit)
        End If
    End If
End Property

'Reads the data from the specified file number
'The file should be open in binary mode
Public Sub FileRead(nFileNum As Integer)
    Dim i As Long

    'Get array size and redimension array
    Get #nFileNum, , i
    ReDim BitArray(0 To i)
    'Read actual data
    For i = 0 To UBound(BitArray)
        Get #nFileNum, , BitArray(i)
    Next i
End Sub

'Writes the data to the specified file number
'The file should be open in binary mode
Public Sub FileWrite(nFileNum As Integer)
    Dim i As Long

    'Save array size
    i = UBound(BitArray)
    Put #nFileNum, , i
    'Save actual data
    For i = 0 To UBound(BitArray)
        Put #nFileNum, , BitArray(i)
    Next i
End Sub

