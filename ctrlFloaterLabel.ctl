VERSION 5.00
Begin VB.UserControl ctrlFloaterLabel 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BackStyle       =   0  'Transparent
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2205
   Enabled         =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   2205
   Begin VB.Label lblFloater 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   0
      MousePointer    =   2  'Cross
      TabIndex        =   0
      Top             =   0
      Width           =   105
   End
End
Attribute VB_Name = "ctrlFloaterLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub setCaption(theCaption As String)
    lblFloater.Caption = theCaption
    UserControl.Height = lblFloater.Height
    UserControl.Width = lblFloater.Width
End Sub
