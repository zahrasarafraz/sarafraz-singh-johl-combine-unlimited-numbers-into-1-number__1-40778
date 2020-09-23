VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Combine Numbers - by Sarafraz Johl"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDecode 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   90
      TabIndex        =   7
      Text            =   "ANSWER WILL BE DISPLAYED HERE"
      Top             =   1110
      Width           =   5295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Help"
      Height          =   1335
      Left            =   105
      TabIndex        =   5
      Top             =   1485
      Width           =   5310
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000080&
         Height          =   960
         Left            =   405
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "Form1.frx":0E42
         Top             =   255
         Width           =   4830
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2745
      Left            =   5490
      ScaleHeight     =   2685
      ScaleWidth      =   0
      TabIndex        =   3
      Top             =   105
      Width           =   60
   End
   Begin VB.TextBox txtMaxVal 
      Height          =   300
      Left            =   5940
      TabIndex        =   2
      Text            =   "255"
      Top             =   420
      Width           =   1320
   End
   Begin VB.TextBox txtEncode 
      Height          =   300
      Left            =   75
      TabIndex        =   1
      Text            =   "255,0,255"
      Top             =   135
      Width           =   5310
   End
   Begin VB.CommandButton cmdEncode 
      Caption         =   "Encode Numbers"
      Default         =   -1  'True
      Height          =   345
      Left            =   90
      TabIndex        =   0
      Top             =   510
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   6300
      Picture         =   "Form1.frx":108B
      Top             =   1155
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Max Range:"
      Height          =   195
      Left            =   5685
      TabIndex        =   4
      Top             =   195
      Width           =   870
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function HighValue(ByRef ObjNumber As Long, ByRef MaxVal As Long) As Variant
    If ObjNumber = 0 Then
        HighValue = MaxVal
    Else
        Dim i As Long
        For i = ObjNumber - 1 To 0 Step -1
            HighValue = HighValue + HighValue(i, MaxVal)
        Next
        HighValue = (HighValue + 1) * MaxVal
    End If
End Function

Private Function Combine(arr As Variant, ByVal MaxVal As Long) As Variant
'On Error GoTo errH
Dim i As Long
Dim Count As Long
Dim lngObjectValue As Variant
    Count = 0
    For i = LBound(arr) To UBound(arr)
        'Error Control
        If IsNumeric(arr(i)) = False Then
            arr(i) = 0
        ElseIf arr(i) < 0 Then
            arr(i) = 0
        ElseIf arr(i) > MaxVal Then
            arr(i) = MaxVal
        Else
            arr(i) = Round(arr(i))
        End If
        'Combine objects
        If Count = 0 Then
            lngObjectValue = Val(arr(i))
        Else
            Dim y As Long
            For y = Count - 1 To 0 Step -1
                lngObjectValue = lngObjectValue + HighValue(y, MaxVal)
            Next
            lngObjectValue = (lngObjectValue + 1) * arr(i)
        End If
        Combine = Combine + lngObjectValue
        lngObjectValue = 0
        Count = Count + 1
    Next
errH:
    Exit Function
    MsgBox Err.Description
End Function

Private Sub cmdEncode_Click()
    txtDecode = Combine(Split(txtEncode, ","), Val(txtMaxVal))
End Sub
