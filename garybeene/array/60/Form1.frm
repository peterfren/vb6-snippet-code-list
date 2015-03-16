VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'set AutoRedraw = True
'set StartUpPosition = 1 - CenterOwner


Private Sub Form_Activate()
'Cycle through an array
Dim i As Byte: i = 100
Dim MyArray(100)    'default lower limit is zero in VB

For i = LBound(MyArray) To UBound(MyArray)
    'do something with MyArray(i)
    MyArray(i) = i
Next i

'to decrease glitch
Me.AutoRedraw = True

'if you know the upper/lower limits of the array, just use them
For i = 0 To 100
    'do something to MyArray(i)
    MyArray(i) = "Array " & i
    
    'Print MyArray on form
    Print MyArray(i)
    
    'Print MyArray on Immediate window
    Debug.Print MyArray(i)
Next i
End Sub

