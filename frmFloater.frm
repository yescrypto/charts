VERSION 5.00
Begin VB.Form frmFloater 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "floater"
   ClientHeight    =   1845
   ClientLeft      =   17145
   ClientTop       =   420
   ClientWidth     =   3120
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      MousePointer    =   2  'Cross
      TabIndex        =   0
      Top             =   -90
      Width           =   2835
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "01:31"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   5
         Left            =   60
         MousePointer    =   2  'Cross
         TabIndex        =   6
         Top             =   1260
         Width           =   450
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "volume: 0"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   4
         Left            =   60
         MousePointer    =   2  'Cross
         TabIndex        =   5
         Top             =   960
         Width           =   810
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "high: 0"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   3
         Left            =   60
         MousePointer    =   2  'Cross
         TabIndex        =   4
         Top             =   540
         Width           =   630
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "low: 0"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   2
         Left            =   60
         MousePointer    =   2  'Cross
         TabIndex        =   3
         Top             =   720
         Width           =   540
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "close: 0"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   1
         Left            =   60
         MousePointer    =   2  'Cross
         TabIndex        =   2
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "open: 0"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   120
         Index           =   0
         Left            =   60
         MousePointer    =   2  'Cross
         TabIndex        =   1
         Top             =   131
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmFloater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub fillValues(lowPrice As Double, highPrice As Double, openPrice As Double, closePrice As Double, theVolume As Double, theTime As String)
 Dim i As Integer, widestLabel As Long, currentWidth As Long
    
    lblInfo(0).Caption = "open:    " & CStr(openPrice)
    lblInfo(1).Caption = "close:   " & CStr(closePrice)
    lblInfo(2).Caption = "low:     " & CStr(lowPrice)
    lblInfo(3).Caption = "high:    " & CStr(highPrice)
    lblInfo(4).Caption = "volume:  " & CStr(theVolume)
    lblInfo(5).Caption = theTime
    
    widestLabel = lblInfo(0).Width
    For i = 1 To 5
        currentWidth = lblInfo(i).Width
        If currentWidth > widestLabel Then widestLabel = currentWidth
    Next i
    
    Frame1.Width = lblInfo(0).Left + widestLabel + 60
    Me.Width = lblInfo(0).Left + widestLabel + 60
End Sub

Private Sub Form_Load()
    Me.Width = Frame1.Width
    Frame1.Height = lblInfo(5).Top + lblInfo(5).Height + 20
    Me.Height = Frame1.Height + Frame1.Top
    stayOnTop Me.hWnd
End Sub

'Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Dim mousePosition As POINTAPI, floaterTop As Long, floaterLeft As Long, i As Long
'
'    'If frmFloater.Visible = False Then frmFloater.Visible = True
'
'    GetCursorPos mousePosition
'    floaterTop = ScaleY(mousePosition.Y, vbPixels, frmChart.ScaleMode) + 20
'    floaterLeft = ScaleX(mousePosition.X, vbPixels, frmChart.ScaleMode) + 20
'
'    If floaterTop < frmChart.Top Or floaterTop > frmChart.Top + frmChart.Height Or _
'       floaterLeft < frmChart.Left Or floaterLeft > frmChart.Left + frmChart.Width Then
'
'        Me.Visible = False
'        Exit Sub
'
'    End If
'
'    Me.Top = floaterTop
'    Me.Left = floaterLeft
'End Sub
