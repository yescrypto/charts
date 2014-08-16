VERSION 5.00
Begin VB.UserControl ctrlCandle 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   555
   MousePointer    =   2  'Cross
   ScaleHeight     =   1695
   ScaleWidth      =   555
   ToolboxBitmap   =   "ctrlCandle.ctx":0000
   Begin VB.Timer tmrResize 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblCandle 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Enabled         =   0   'False
      ForeColor       =   &H0000C000&
      Height          =   1425
      Left            =   0
      MousePointer    =   2  'Cross
      TabIndex        =   0
      Top             =   0
      Width           =   555
   End
   Begin VB.Line lineWick 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      X1              =   277
      X2              =   277
      Y1              =   0
      Y2              =   1680
   End
End
Attribute VB_Name = "ctrlCandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private priceOpen As Double
Private priceClose As Double
Private priceHigh As Double
Private priceLow As Double

Private myVolume As Double
Private myTimePeriod As String
Private previousColor As ColorConstants

Public Sub setPrices(openingPrice As Double, closingPrice As Double, highestPrice As Double, lowestPrice As Double, Optional lastColor As ColorConstants = vbGreen, Optional timePeriod As String, Optional theVolume As Double)
 Dim priceRatio As Double, tempValue As Double, controlHeight As Long, controlWidth As Long, wickMiddle As Double, candleRange As Double, candleHeight As Long
    
    'set properties
    priceOpen = openingPrice
    openPrice = priceOpen
    priceClose = closingPrice
    closePrice = priceClose
    priceHigh = highestPrice
    highPrice = priceHigh
    priceLow = lowestPrice
    lowPrice = priceLow
    myVolume = theVolume
    myTimePeriod = timePeriod
    previousColor = lastColor
    
    'calculate and size everything
    controlHeight = UserControl.Height
    controlWidth = UserControl.Width
    
    If lowestPrice = highestPrice Then
        UserControl.BackStyle = 1: UserControl.BackColor = previousColor
    Else
        UserControl.BackStyle = 0
    End If
    
    'set line width
    If controlWidth < 500 Then
        lineWick.borderWidth = 1
    Else
        lineWick.borderWidth = Int(controlWidth / 500)
    End If
    
    tempValue = highestPrice - lowestPrice
    If tempValue <> 0 Then
        priceRatio = controlHeight / tempValue
    Else
        priceRatio = 0
    End If
    
    wickMiddle = Int((controlWidth / 2) - (lineWick.borderWidth / 2) + 0.5)
    lineWick.X1 = wickMiddle
    lineWick.X2 = wickMiddle
    
    lineWick.Y1 = 0
    lineWick.Y2 = controlHeight
    
    lblCandle.Width = controlWidth
    
    If lowestPrice = highestPrice Then
        UserControl.BackColor = lastColor
        UserControl.BackStyle = 1
        UserControl.Height = 15
        candleHeight = controlHeight
        candleRange = Abs(openingPrice - closingPrice)
    Else
        candleRange = Abs(openingPrice - closingPrice)
        candleHeight = candleRange * priceRatio
    End If
    
    lblCandle.Height = candleHeight
'    lineWick.ZOrder 0
'    lblCandle.ZOrder 0
    'MsgBox lblCandle.Height
    
    If openingPrice > closingPrice Then
        lblCandle.BackColor = RGB(255, 0, 0)
        lineWick.BorderColor = RGB(255, 0, 0)
        previousColor = RGB(255, 0, 0)
        lblCandle.Top = ((highestPrice - openingPrice) * priceRatio) ' - 10
    ElseIf closingPrice > openingPrice Then
        lblCandle.BackColor = RGB(0, 255, 0)
        lineWick.BorderColor = RGB(0, 255, 0)
        previousColor = RGB(0, 255, 0)
        lblCandle.Top = ((highestPrice - closingPrice) * priceRatio) ' - 10
    Else
        lblCandle.BackColor = previousColor
        lineWick.BorderColor = previousColor
        lblCandle.Top = ((highestPrice - openingPrice) * priceRatio) ' - 10
    End If
    
    If lblCandle.Top + lblCandle.Height > UserControl.Height Then
        lblCandle.Top = UserControl.Height - lblCandle.Height
    End If
    
End Sub

Public Function candleWidth() As Long
    candleWidth = lblCandle.Width
End Function
Public Property Get openPrice() As Variant
    openPrice = priceOpen
End Property

Public Property Let openPrice(ByVal priceValue As Variant)
    priceOpen = priceValue
    PropertyChanged "openPrice"
End Property

Public Property Get closePrice() As Variant
    closePrice = priceClose
End Property

Public Property Let closePrice(ByVal priceValue As Variant)
    priceClose = priceValue
    PropertyChanged "closePrice"
End Property

Public Property Get highPrice() As Variant
    highPrice = priceHigh
End Property

Public Property Let highPrice(ByVal priceValue As Variant)
    priceHigh = priceValue
    PropertyChanged "highPrice"
End Property

Public Property Get lowPrice() As Variant
    lowPrice = priceLow
End Property

Public Property Let lowPrice(ByVal priceValue As Variant)
    priceLow = priceValue
    PropertyChanged "lowPrice"
End Property

Public Property Get volume() As Variant
    volume = myVolume
End Property

Public Property Let volume(ByVal volumeValue As Variant)
    myVolume = volumeValue
    PropertyChanged "volume"
End Property
Public Property Get timePeriod() As Variant
    timePeriod = myTimePeriod
End Property

Public Property Let timePeriod(ByVal timePeriodValue As Variant)
    myTimePeriod = timePeriodValue
    PropertyChanged "timePeriod"
End Property

Public Property Get lastColor() As Variant
    lastColor = previousColor
End Property

Public Property Let lastColor(ByVal colorValue As Variant)
    previousColor = colorValue
    PropertyChanged "lastColor"
End Property

Private Sub tmrResize_Timer()
    tmrResize.Enabled = False
    hideMe True
    setPrices priceOpen, priceClose, priceHigh, priceLow, previousColor, myTimePeriod, myVolume
End Sub

Public Sub lightUp(upOrDown As Boolean)
    If upOrDown = True Or priceHigh = priceLow Then
        UserControl.BackStyle = 1
        'MsgBox "line:visible:" & CStr(lineWick.Visible) & "-color:" & CStr(lineWick.BorderColor) & vbCrLf & _
               "candle:visible:" & CStr(lblCandle.Visible) & "-color:" & CStr(lblCandle.BackColor)
        
    Else
        UserControl.BackStyle = 0
    End If
End Sub

Public Function isLit() As Boolean
    If UserControl.BackStyle = 1 Then
        isLit = True
    Else
        isLit = False
    End If
End Function


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    priceOpen = PropBag.ReadProperty("openPrice", False)
    priceClose = PropBag.ReadProperty("closePrice", False)
    priceHigh = PropBag.ReadProperty("highPrice", False)
    priceLow = PropBag.ReadProperty("lowPrice", False)
    myVolume = PropBag.ReadProperty("volume", False)
    myTimePeriod = PropBag.ReadProperty("timePeriod", False)
    previousColor = PropBag.ReadProperty("lastColor", False)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "openPrice", priceOpen
    PropBag.WriteProperty "closePrice", priceClose
    PropBag.WriteProperty "highPrice", priceHigh
    PropBag.WriteProperty "lowPrice", priceLow
    PropBag.WriteProperty "volume", myVolume
    PropBag.WriteProperty "timePeriod", myTimePeriod
    PropBag.WriteProperty "lastColor", previousColor
End Sub

Private Sub UserControl_Resize()
'    tmrResize.Enabled = False
'    tmrResize.Enabled = True
End Sub


