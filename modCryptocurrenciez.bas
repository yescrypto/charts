Attribute VB_Name = "modCryptocurrenciez"
Option Explicit

Public currentMarket As String
Public currentMarketSymbol As String
Public currentMarketID As String

Public Const defaultSections = 5

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

    Type POINTAPI
          X As Long
          Y As Long
    End Type

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOZORDER = &H4
Private Const WS_THICKFRAME = &H40000

Public candleSpacing As Long



Public Sub makeNormal(hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
Public Sub stayOnTop(hWnd As Long)
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Function mouseY() As Long
 Dim mousePosition As POINTAPI
    GetCursorPos mousePosition
    mouseY = mousePosition.Y
End Function

Public Function mouseX() As Long
 Dim mousePosition As POINTAPI
    GetCursorPos mousePosition
    mouseX = mousePosition.X
End Function

Public Function speedTestControl(theControl As Control, lastControl As Control, Optional doTheEvents As Boolean = False) As Double
 Dim startTime As Double, endTime As Double, i As Long
    
    startTime = Timer
    
    If doTheEvents = True Then DoEvents
    
    For i = 1 To 10000
        theControl.BackColor = RGB(255, 0, 0)
        If doTheEvents = True Then DoEvents
        theControl.Left = lastControl.Left + lastControl.Width + randomValue(10, 100)
        If theControl.Parent.ScaleWidth < theControl.Left + theControl.Width Then theControl.Parent.Width = theControl.Parent.Width + theControl.Width + 50
        If doTheEvents = True Then DoEvents
        theControl.Height = lastControl.Height - randomValue(1, 100)
        If doTheEvents = True Then DoEvents
        theControl.Height = theControl.Height + randomValue(1, 100)
        If theControl.Parent.scaleHeight < theControl.Top + theControl.Height Then theControl.Parent.Height = theControl.Parent.Height + theControl.Height + 50
        If doTheEvents = True Then DoEvents
        theControl.BackColor = RGB(0, 255, 0)
        If doTheEvents = True Then DoEvents
        theControl.Top = theControl.Top - randomValue(1, 100)
        If doTheEvents = True Then DoEvents
        theControl.Top = theControl.Top + randomValue(1, 101)
        If theControl.Parent.scaleHeight < theControl.Top + theControl.Height Then theControl.Parent.Height = theControl.Parent.Height + theControl.Height + 50
        If doTheEvents = True Then DoEvents
    Next i
    
    If theControl.Parent.ScaleWidth < theControl.Left + theControl.Width Then theControl.Parent.Width = theControl.Parent.Width + theControl.Width + 50
    If theControl.Parent.scaleHeight < theControl.Top + theControl.Height Then theControl.Parent.Height = theControl.Parent.Height + theControl.Height + 50
    
    If doTheEvents = True Then DoEvents
    endTime = Timer
    
    speedTestControl = CStr(endTime - startTime)
    
End Function

Private Function randomValue(minValue As Long, maxValue As Long) As Long
    Randomize Timer * Rnd + Timer
    
    randomValue = Int(Rnd * (maxValue + 1 - (minValue))) + (minValue)

End Function

Public Sub hideMe(yesOrNo As Boolean)
    If yesOrNo = True Then
        frmChart.lblLoading.FontSize = Int(frmChart.ScaleWidth / 186)
        frmChart.priceChart.Visible = False
        frmChart.tmrHideChart.Enabled = False
        frmChart.tmrHideChart.Enabled = True
    Else
        frmChart.priceChart.Visible = True
    End If
    
    Const swpFlags As Long = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE
    
    SetWindowPos frmChart.hWnd, 0, 0, 0, 0, 0, swpFlags
End Sub

