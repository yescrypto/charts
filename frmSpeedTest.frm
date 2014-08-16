VERSION 5.00
Begin VB.Form frmSpeedTest 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   1140
   ClientLeft      =   7080
   ClientTop       =   1935
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   1140
   ScaleWidth      =   7065
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   1000
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   500
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1000
      Index           =   0
      Left            =   60
      ScaleHeight     =   975
      ScaleWidth      =   465
      TabIndex        =   0
      Top             =   60
      Width           =   500
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1005
      Index           =   0
      Left            =   60
      Top             =   60
      Width           =   495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   495
   End
End
Attribute VB_Name = "frmSpeedTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub testSpeed()

 Dim tempstring As String, timeSpent As Double, i As Integer
    
    Me.Height = 1725
    Me.Width = 7305
    timeSpent = 0
    For i = 1 To 40
        Load Picture1(i)
        Picture1(i).Visible = True
        timeSpent = timeSpent + speedTestControl(Picture1(i), Picture1(i - 1), True)
        DoEvents
    Next
    For i = 1 To 40
        Unload Picture1(i)
    Next i
    Me.Height = 1725
    Me.Width = 7305
    tempstring = tempstring & "doevents" & vbCrLf & vbCrLf & "pictureBox:" & Chr(9) & CStr(timeSpent) & vbCrLf
    timeSpent = 0
    For i = 1 To 40
        Load Text1(i)
        Text1(i).Visible = True
        timeSpent = timeSpent + speedTestControl(Text1(i), Text1(i - 1), True)
        DoEvents
    Next
    For i = 1 To 40
        Unload Text1(i)
    Next i
    Me.Height = 1725
    Me.Width = 7305
    tempstring = tempstring & "textBox:" & Chr(9) & Chr(9) & CStr(timeSpent) & vbCrLf
    timeSpent = 0
    For i = 1 To 40
        Load Label1(i)
        Label1(i).Visible = True
        timeSpent = timeSpent + speedTestControl(Label1(i), Label1(i - 1), True)
        DoEvents
    Next
    For i = 1 To 40
        Unload Label1(i)
    Next i
    Me.Height = 1725
    Me.Width = 7305
    tempstring = tempstring & "label:" & Chr(9) & Chr(9) & CStr(timeSpent) & vbCrLf
    timeSpent = 0
    For i = 1 To 40
        Load Shape1(i)
        Shape1(i).Visible = True
        timeSpent = timeSpent + speedTestControl(Shape1(i), Shape1(i - 1), True)
        DoEvents
    Next
    For i = 1 To 40
        Unload Shape1(i)
    Next i
    Me.Height = 1725
    Me.Width = 7305
    tempstring = tempstring & "shape:" & Chr(9) & Chr(9) & CStr(timeSpent)
    
    MsgBox tempstring, , "results"
End Sub

