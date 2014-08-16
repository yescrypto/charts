VERSION 5.00
Begin VB.UserControl ctrlChart 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   MouseIcon       =   "ctrlChart.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   4065
   ScaleWidth      =   6150
   Begin prjChart.ctrlFloaterLabel lblMousePrice 
      Height          =   795
      Left            =   1740
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   1402
   End
   Begin prjChart.ctrlVolume ctrlVolume 
      Height          =   1275
      Left            =   -5220
      TabIndex        =   2
      Top             =   4680
      Width           =   5175
      _ExtentX        =   9340
      _ExtentY        =   2249
   End
   Begin VB.Timer tmrCross 
      Interval        =   33
      Left            =   5700
      Top             =   840
   End
   Begin VB.Timer tmrFloater 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5700
      Top             =   420
   End
   Begin VB.Timer tmrResize 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5700
      Top             =   0
   End
   Begin prjChart.ctrlCandle ctrlCandle 
      Height          =   15
      Index           =   0
      Left            =   3900
      TabIndex        =   0
      Top             =   540
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   26
      openPrice       =   0
      closePrice      =   0
      highPrice       =   0
      lowPrice        =   0
      volume          =   0
      timePeriod      =   ""
      lastColor       =   0
   End
   Begin VB.Label lblTimes 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
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
      Index           =   0
      Left            =   1200
      MousePointer    =   2  'Cross
      TabIndex        =   4
      Top             =   3300
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Line lineTimes 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      Index           =   0
      Visible         =   0   'False
      X1              =   1620
      X2              =   3060
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line lineHorizontal 
      BorderColor     =   &H00C0C0C0&
      X1              =   -1080
      X2              =   2220
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line lineVertical 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   0
      Y1              =   -720
      Y2              =   1080
   End
   Begin VB.Line linePrices 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      Index           =   0
      Visible         =   0   'False
      X1              =   0
      X2              =   1440
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblPrices 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0.000123"
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
      Index           =   0
      Left            =   0
      MousePointer    =   2  'Cross
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   780
   End
End
Attribute VB_Name = "ctrlChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lowestPrice As Double
Private highestPrice As Double

Private myVolume As Double

Private highRange As Double
Private lowRange As Double

Private graphSections As Long

Private usableWidth As Long
Private candleWidth As Long
Private widestLabel As Long

Private lastColor As ColorConstants

Private isActive As Boolean

Private currentCandle As Long

Private Const chartPercentage = 0.78
Private Const volumePercentage = 0.17

Private floaterMove As Long

Public Sub formatChart(lowPrice As Double, highPrice As Double, sections As Long)
 Dim upperRange As Double, lowerRange As Double, rangeDifference As Double, i As Long, highPriceString As String, _
     lowPriceString As String, sizeRatio As Double, tempValue As Double, timesTen As Double, tempLong As Long, _
     highDotPosition As Long, lowDotPosition As Long, tempHighVal As Double, tempLowVal As Double, tempLowLen As Long, _
     tempHighLen As Long, priceRatio As Double, usableHeight As Long, candleCount As Long, timeCount As Long, endOfCandle As Long, theDateTime As String
    
    isActive = True
    
    highPriceString = CStr(highPrice)
    lowPriceString = CStr(lowPrice)
    
    ctrlVolume.Height = (volumePercentage * UserControl.Height) - candleSpacing
    ctrlVolume.Width = UserControl.Width
    ctrlVolume.Left = 0
    ctrlVolume.Top = UserControl.Height * chartPercentage
    
    If highPrice = lowPrice Then
        graphSections = 2
    Else
        graphSections = sections
    End If
    
    candleCount = ctrlCandle.UBound
    timeCount = Int(candleCount / graphSections)
    
    If candleCount > 0 And timeCount > 0 Then
        If lineTimes.UBound > 0 Then
            For i = 1 To lineTimes.UBound
                Unload lineTimes(i)
                Unload lblTimes(i)
            Next i
        End If
        
        For i = 1 To graphSections
            Load lineTimes(i)
            Load lblTimes(i)
            
            With lineTimes(i)
                .Y1 = 0
                .Y2 = (chartPercentage + volumePercentage) * UserControl.Height + candleSpacing
                endOfCandle = ctrlCandle(i * timeCount).Width + ctrlCandle(i * timeCount).Left
                .X1 = endOfCandle
                .X2 = endOfCandle
                .Visible = True
            End With
            
            lblTimes(i).Left = endOfCandle - (lblTimes(i).Width / 2)
            lblTimes(i).Top = (chartPercentage + volumePercentage) * UserControl.Height
            theDateTime = ctrlCandle(i * timeCount).timePeriod
            If InStr(theDateTime, " ") <> 0 Then theDateTime = Split(theDateTime, " ", 2)(1)
            lblTimes(i).Caption = theDateTime
            lblTimes(i).Visible = True
        Next i
    End If
    
    If InStr(highPriceString, ".") = 0 Then highPriceString = highPriceString & ".0"
    If InStr(lowPriceString, ".") = 0 Then lowPriceString = lowPriceString & ".0"

    highDotPosition = InStr(highPriceString, ".")
    lowDotPosition = InStr(lowPriceString, ".")

    If lowDotPosition < highDotPosition Then
        lowPriceString = String(highDotPosition - lowDotPosition, "0") & lowPriceString
    End If
    
    tempLowLen = Len(lowPriceString): tempHighLen = Len(highPriceString)
    
    If tempLowLen < tempHighLen Then
        lowPriceString = lowPriceString & String(tempHighLen - tempLowLen, "0")
    ElseIf tempLowLen > tempHighLen Then
        highPriceString = highPriceString & String(tempLowLen - tempHighLen, "0")
    End If
    
    highPriceString = Replace(highPriceString, ".", "")
    lowPriceString = Replace(lowPriceString, ".", "")
    
    i = 1
    If highPrice > lowPrice Then
        Do
            If i > Len(highPriceString) Then
                highPriceString = highPriceString & "0"
                lowPriceString = lowPriceString & "0"
            End If
            tempHighVal = Val(Mid(highPriceString, 1, i) & "." & Right(highPriceString, Len(highPriceString) - i))
            tempLowVal = Val(Mid(lowPriceString, 1, i) & "." & Right(lowPriceString, Len(lowPriceString) - i))
            If tempHighVal - tempLowVal > graphSections Then Exit Do
            i = i + 1
        Loop
    Else
        Do
            If i > Len(highPriceString) Then
                highPriceString = highPriceString & "0"
                lowPriceString = lowPriceString & "0"
            End If
            tempHighVal = Val(Mid(highPriceString, 1, i) & "." & Right(highPriceString, Len(highPriceString) - i))
            tempLowVal = Val(Mid(lowPriceString, 1, i) & "." & Right(lowPriceString, Len(lowPriceString) - i))
            If tempHighVal > graphSections Then Exit Do
            i = i + 1
        Loop
    End If
    
    tempValue = tempHighVal / graphSections
    If tempValue <> Int(tempValue) Or highPrice = lowPrice Then tempValue = Int(tempValue) + 1
    upperRange = tempValue * graphSections
    lowerRange = Int(tempLowVal / graphSections) * graphSections
    
    tempValue = upperRange - lowerRange
    upperRange = upperRange / (10 ^ ((i + 1) - highDotPosition))
    lowerRange = lowerRange / (10 ^ ((i + 1) - highDotPosition))
    
    highestPrice = upperRange
    lowestPrice = lowerRange
    
    rangeDifference = tempValue / (10 ^ ((i + 1) - highDotPosition))
    usableHeight = UserControl.Height * chartPercentage
    sizeRatio = usableHeight / graphSections
    priceRatio = rangeDifference / graphSections
    
    If linePrices.UBound > 0 Then
        For i = 1 To linePrices.UBound
            Unload linePrices(i)
            Unload lblPrices(i)
        Next i
    End If
    lblPrices(0).Caption = CStr(upperRange) & " "
    lblPrices(0).Visible = True
'    lblPrices(0).ZOrder 0
    lblPrices(0).Top = 0
    lblPrices(0).Left = 0
    linePrices(0).X1 = 0
    linePrices(0).X2 = UserControl.Width
    linePrices(0).Y1 = 0
    linePrices(0).Y2 = 0
    
    For i = 1 To graphSections
        Load linePrices(i)
        linePrices(i).Visible = True
        Load lblPrices(i)
        lblPrices(i).Visible = True
'        lblPrices(i).ZOrder 0
        lblPrices(i).Caption = CStr(upperRange - (i * priceRatio)) & " "
        lblPrices(i).Top = (i * sizeRatio) - (lblPrices(i).Height / 2)
        lblPrices(i).Left = 0
        linePrices(i).X1 = lblPrices(i).Width
        linePrices(i).X2 = UserControl.Width
        linePrices(i).Y1 = i * sizeRatio
        linePrices(i).Y2 = i * sizeRatio
    Next i
    
    linePrices(i - 1).BorderStyle = 1
    linePrices(i - 1).X1 = 0
    usableHeight = UserControl.Height * chartPercentage
    lblPrices(graphSections).Top = usableHeight - lblPrices(graphSections).Height
    tempValue = usableHeight - candleSpacing
    linePrices(graphSections).Y1 = tempValue
    linePrices(graphSections).Y2 = tempValue
    tmrFloater.Enabled = True
    'ctrlVolume.formatVolume
End Sub

Public Sub addCandle(openPrice As Double, closePrice As Double, lowPrice As Double, highPrice As Double, Optional theVolume As Double, Optional theTime As String)
 Dim i As Long, candleCount As Long, chartRange As Double, priceRatio As Double, tempValue As Double
    
    candleCount = ctrlCandle.UBound + 1
    Load ctrlCandle(candleCount)
    
    If lastColor = 0 Then lastColor = vbGreen
    
    ctrlCandle(candleCount).setPrices openPrice, closePrice, highPrice, lowPrice, lastColor, theTime, theVolume
    
    If openPrice > closePrice Then
        lastColor = vbRed
    ElseIf openPrice < closePrice Then
        lastColor = vbGreen
    End If
    
    If candleCount = 1 Then
        highestPrice = highPrice
        lowestPrice = lowPrice
    Else
        If highPrice > highestPrice Then highestPrice = highPrice
        If lowPrice < lowestPrice Then lowestPrice = lowPrice
    End If
    
    
    For i = 0 To lblPrices.UBound
        If lblPrices(i).Width > widestLabel Then widestLabel = lblPrices(i).Width
    Next i
    
    usableWidth = UserControl.Width - widestLabel - candleSpacing - (candleCount * candleSpacing)
    candleWidth = usableWidth / candleCount
    
    myVolume = myVolume + theVolume
    
    ctrlCandle(candleCount).Visible = True
    'ctrlCandle(candleCount).ZOrder 1
End Sub

Public Sub sizeCandles()
 Dim i As Long, chartRange As Double, priceRatio As Double, usableHeight As Long
    
    chartRange = highestPrice - lowestPrice
    
    usableHeight = UserControl.Height * chartPercentage
    priceRatio = usableHeight / chartRange
    
    ctrlVolume.clearVolume
    ctrlVolume.Height = volumePercentage * UserControl.Height
    ctrlVolume.Width = UserControl.Width
    ctrlVolume.Left = 0
    ctrlVolume.Top = UserControl.Height * chartPercentage
    
    For i = 1 To ctrlCandle.UBound
        ctrlCandle(i).Height = (ctrlCandle(i).highPrice - ctrlCandle(i).lowPrice) * priceRatio
        ctrlCandle(i).Top = ((highestPrice - ctrlCandle(i).highPrice) * priceRatio)
'        If i > 1 And ctrlCandle(i).Top < ctrlCandle(i - 1).Top + ctrlCandle(i - 1).Height And ctrlCandle(i - 1).lowPrice = ctrlCandle(i).highPrice And ctrlCandle(i).highPrice = ctrlCandle(i).lowPrice Then
'            ctrlCandle(i).Top = ctrlCandle(i).Top + 30
'        End If
        ctrlCandle(i).Width = candleWidth
        ctrlCandle(i).Left = widestLabel + candleSpacing + ((candleWidth + candleSpacing) * (i - 1))
        ctrlCandle(i).setPrices ctrlCandle(i).openPrice, ctrlCandle(i).closePrice, ctrlCandle(i).highPrice, ctrlCandle(i).lowPrice, ctrlCandle(i).lastColor, ctrlCandle(i).timePeriod, ctrlCandle(i).volume
        ctrlVolume.addBar ctrlCandle(i).volume, ctrlCandle(i).lastColor, ctrlCandle(i).Left, ctrlCandle(i).candleWidth
    Next i

    If graphSections = 2 And lowestPrice < highestPrice Then graphSections = defaultSections
    formatChart lowestPrice, highestPrice, graphSections
    ctrlVolume.formatVolume
End Sub

Public Property Get chartSections() As Variant
    chartSections = graphSections
End Property

Public Property Let chartSections(ByVal sectionsValue As Variant)
    graphSections = sectionsValue
    PropertyChanged "chartSections"
End Property

Private Sub tmrCross_Timer()
 Dim mousePosition As POINTAPI, yAxis As Long, xAxis As Long, borderWidth As Long, scaleHeight As Long, untouchedY As Long, untouchedX As Long, parentTop As Long, _
     parentLeft As Long, i As Long, candleRange As Long, candleStart As Long, h As Long, candleCount As Long, priceRange As Double, priceRatio As Double, mousePrice As Double, _
     labelTop As Long, floaterTop As Long, labelLeft As Long, floaterLeft As Long
    
    On Error GoTo theEnd
    
    GetCursorPos mousePosition
    borderWidth = (UserControl.Parent.Width - UserControl.Parent.ScaleWidth) / 2
    scaleHeight = UserControl.Parent.Height - UserControl.Parent.scaleHeight - borderWidth
    untouchedY = ScaleY(mousePosition.Y, vbPixels, 1)
    parentTop = UserControl.Parent.Top
    yAxis = untouchedY - parentTop - scaleHeight
    untouchedX = ScaleX(mousePosition.X, vbPixels, 1)
    parentLeft = UserControl.Parent.Left
    xAxis = untouchedX - parentLeft - borderWidth
    
    If untouchedY > UserControl.Parent.Height + parentTop Or untouchedY < parentTop Or untouchedX > UserControl.Parent.Width + parentLeft Or untouchedX < parentLeft Then
        
        With lineHorizontal
            .X1 = 0
            .X2 = 0
            .Y1 = 0
            .Y2 = 0
        End With
        
        With lineVertical
            .Y1 = 0
            .Y2 = 0
            .X1 = 0
            .X2 = 0
        End With
        
        Exit Sub
    End If
    
    With lineHorizontal
        .X1 = 0
        .X2 = UserControl.Width
        .Y1 = yAxis
        .Y2 = yAxis
    End With
    
    With lineVertical
        .Y1 = 0
        .Y2 = UserControl.Height
        .X1 = xAxis
        .X2 = xAxis
    End With
    
    If tmrResize.Enabled = True Then Exit Sub
    
    candleCount = ctrlCandle.UBound
    If candleCount = 0 Then Exit Sub
    
    candleStart = ctrlCandle(1).Left
    If xAxis < ctrlCandle(1).Left Then
        If frmFloater.Visible = True Then frmFloater.Visible = False
        If lblMousePrice.Visible = True Then lblMousePrice.Visible = False
'        If ctrlCandle(1).isLit = True Then ctrlCandle(1).lightUp False
        Exit Sub
    End If
    
    candleRange = candleWidth + candleSpacing
    i = Int(((xAxis - candleStart) / candleRange) + 1)
    
    If i > ctrlCandle.UBound Then i = ctrlCandle.UBound
'    If ctrlCandle(i).isLit = False Then
'
'        'light it up
'        ctrlCandle(i).lightUp (True)
'
'
'    End If
'
'    'transparent the rest
'    If i > 1 Then
'        For h = 1 To i - 1
'            ctrlCandle(h).lightUp (False)
'        Next h
'    End If
'    If i < candleCount Then
'        For h = i + 1 To candleCount
'            ctrlCandle(h).lightUp (False)
'        Next h
'    End If
    
    floaterMove = floaterMove + 1
    If i <> currentCandle Or floaterMove > 1 Then
        'fill floater
        frmFloater.fillValues ctrlCandle(i).lowPrice, ctrlCandle(i).highPrice, ctrlCandle(i).openPrice, ctrlCandle(i).closePrice, ctrlCandle(i).volume, ctrlCandle(i).timePeriod
        currentCandle = i
        'position floater
        floaterLeft = untouchedX + candleSpacing
        floaterTop = untouchedY + candleSpacing
        labelLeft = xAxis + candleSpacing
        labelTop = yAxis - lblMousePrice.Height
        If floaterLeft + frmFloater.Width > UserControl.Parent.Left + UserControl.Parent.Width Then floaterLeft = untouchedX - candleSpacing - frmFloater.Width
        If floaterTop + frmFloater.Height > UserControl.Parent.Top + UserControl.Parent.Height Then floaterTop = untouchedY - candleSpacing - frmFloater.Height: labelTop = yAxis + candleSpacing
        frmFloater.Left = floaterLeft
        frmFloater.Top = floaterTop
        frmFloater.Visible = True
        floaterMove = 0
        'fill price label
        priceRange = highestPrice - lowestPrice
        priceRatio = (priceRange / (UserControl.Height * chartPercentage))
        mousePrice = Int(((highestPrice - (priceRatio * yAxis)) * 100000000) + 0.5) / 100000000
        
        lblMousePrice.setCaption CStr(mousePrice)
        
        'position price label
        If labelLeft + lblMousePrice.Width > UserControl.Width And yAxis < UserControl.Height * chartPercentage Then
            lblMousePrice.Left = xAxis - candleSpacing - lblMousePrice.Width
            lblMousePrice.Top = labelTop
            lblMousePrice.Visible = True
            lblMousePrice.ZOrder 0
        ElseIf yAxis > UserControl.Height * chartPercentage Then
            lblMousePrice.Visible = False
        Else
            lblMousePrice.Left = labelLeft
            lblMousePrice.Top = labelTop
            lblMousePrice.Visible = True
            lblMousePrice.ZOrder 0
        End If
    End If
    
    
theEnd:
End Sub

Private Sub tmrFloater_Timer()
 Dim mousePosition As POINTAPI, floaterTop As Long, floaterLeft As Long, i As Long
    
    On Error GoTo theEnd
    GetCursorPos mousePosition
    floaterTop = ScaleY(mousePosition.Y, vbPixels, 1) + candleSpacing
    floaterLeft = ScaleX(mousePosition.X, vbPixels, 1) + candleSpacing
    
    If floaterTop < UserControl.Parent.Top Or floaterTop > UserControl.Parent.Top + UserControl.Parent.Height Or _
       floaterLeft < UserControl.Parent.Left Or floaterLeft > UserControl.Parent.Left + UserControl.Parent.Width Then
        
        frmFloater.Visible = False
        lblMousePrice.Visible = False
'        If ctrlCandle.ubound > 0 Then
'            For i = 1 To ctrlCandle.ubound
'                ctrlCandle(i).lightUp False
'            Next i
'        End If
        'tmrFloater.Enabled = False
    
    End If
    
theEnd:
End Sub

Private Sub tmrResize_Timer()
 Dim chartRange As Double, priceRatio As Double, i As Long, candleCount As Long, usableHeight As Long, timeCount As Long, endOfCandle As Long, theDateTime As String
    
    tmrResize.Enabled = False
    
    hideMe True
    
    formatChart lowestPrice, highestPrice, graphSections
    ctrlVolume.Height = (volumePercentage * UserControl.Height) - candleSpacing
    ctrlVolume.Width = UserControl.Width
    ctrlVolume.Left = 0
    ctrlVolume.Top = UserControl.Height * chartPercentage
    
    candleCount = ctrlCandle.UBound
    
    If candleCount > 0 Then
        
        chartRange = highestPrice - lowestPrice
        
        usableHeight = UserControl.Height * chartPercentage
        priceRatio = usableHeight / chartRange
        
        For i = 0 To lblPrices.UBound
            If lblPrices(i).Width > widestLabel Then widestLabel = lblPrices(i).Width
        Next i
        
        usableWidth = UserControl.Width - widestLabel - candleSpacing - (candleCount * candleSpacing)
        candleWidth = usableWidth / candleCount
        
        For i = 1 To candleCount
            ctrlCandle(i).Height = (ctrlCandle(i).highPrice - ctrlCandle(i).lowPrice) * priceRatio
            ctrlCandle(i).Top = ((highestPrice - ctrlCandle(i).highPrice) * priceRatio)
            ctrlCandle(i).Width = candleWidth
            ctrlCandle(i).Left = widestLabel + candleSpacing + ((candleWidth + candleSpacing) * (i - 1))
            ctrlCandle(i).setPrices ctrlCandle(i).openPrice, ctrlCandle(i).closePrice, ctrlCandle(i).highPrice, ctrlCandle(i).lowPrice, ctrlCandle(i).lastColor, ctrlCandle(i).timePeriod, ctrlCandle(i).volume
            ctrlVolume.sizeBar i, ctrlCandle(i).candleWidth, ctrlCandle(i).Left
        Next i
        
        timeCount = Int(candleCount / graphSections)
        
        If timeCount > 0 Then
            If lineTimes.UBound > 0 Then
                For i = 1 To lineTimes.UBound
                    Unload lineTimes(i)
                    Unload lblTimes(i)
                Next i
            End If
            
            For i = 1 To graphSections
                Load lineTimes(i)
                Load lblTimes(i)
                
                With lineTimes(i)
                    .Y1 = 0
                    .Y2 = (chartPercentage + volumePercentage) * UserControl.Height + candleSpacing
                    endOfCandle = ctrlCandle(i * timeCount).Width + ctrlCandle(i * timeCount).Left
                    .X1 = endOfCandle
                    .X2 = endOfCandle
                    .Visible = True
                End With
                
                lblTimes(i).Left = endOfCandle - (lblTimes(i).Width / 2)
                lblTimes(i).Top = (chartPercentage + volumePercentage) * UserControl.Height
                theDateTime = ctrlCandle(i * timeCount).timePeriod
                If InStr(theDateTime, " ") <> 0 Then theDateTime = Split(theDateTime, " ", 2)(1)
                lblTimes(i).Caption = theDateTime
                lblTimes(i).Visible = True
            Next i
        End If
        
        ctrlVolume.formatVolume
    End If
    
End Sub

Public Function labelHeights() As Long
    labelHeights = lblPrices(0).Height * lblPrices.Count
End Function

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu frmChart.mnuMarkets
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    graphSections = PropBag.ReadProperty("chartSections", False)
End Sub

Private Sub UserControl_Resize()
    
    If isActive = False Then Exit Sub
    
    tmrResize.Enabled = False
    tmrResize.Enabled = True
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "chartSections", graphSections
End Sub

Public Sub clearChart()
 Dim i As Long
    
    tmrFloater.Enabled = False
    
    i = ctrlCandle.UBound
    If ctrlCandle.UBound > 0 Then
        For i = 1 To ctrlCandle.UBound
            Unload ctrlCandle(i)
        Next i
    End If
    
    lowestPrice = 99999999
    highestPrice = 0
    myVolume = 0
    ctrlVolume.clearVolume
    
End Sub

Public Sub newOrder(orderPrice As Double, theVolume As Double)
 Dim lastCandle As Long, candleHigh As Double, candleLow As Double, btcVolume As Double, chartRange As Double, priceRatio As Double, candleOpen As Double, usableHeight As Long
    lastCandle = ctrlCandle.UBound
    
    btcVolume = ctrlCandle(lastCandle).volume
    candleOpen = ctrlCandle(lastCandle).openPrice
    
    If btcVolume > 0 Then
        candleHigh = ctrlCandle(lastCandle).highPrice
        candleLow = ctrlCandle(lastCandle).lowPrice
    Else
        candleHigh = candleOpen
        candleLow = candleOpen
    End If
    
    If orderPrice > highestPrice Then
        formatChart lowestPrice, orderPrice, defaultSections
        sizeCandles
        UserControl.Parent.Caption = currentMarketSymbol & "  " & CStr(orderPrice) & "   H:" & CStr(highestPrice) & "   L:" & CStr(lowestPrice) & "   V:" & CStr(myVolume)
        Exit Sub
    ElseIf orderPrice < lowestPrice Then
        formatChart orderPrice, highestPrice, defaultSections
        sizeCandles
        UserControl.Parent.Caption = currentMarketSymbol & "  " & CStr(orderPrice) & "   H:" & CStr(highestPrice) & "   L:" & CStr(lowestPrice) & "   V:" & CStr(myVolume)
        Exit Sub
    End If
    
    chartRange = highestPrice - lowestPrice
    usableHeight = UserControl.Height * chartPercentage
    priceRatio = usableHeight / chartRange
    
    If orderPrice > candleHigh Then
        ctrlCandle(lastCandle).Height = (orderPrice - candleLow) * priceRatio
        ctrlCandle(lastCandle).Top = ((highestPrice - orderPrice) * priceRatio)
        
        ctrlCandle(lastCandle).setPrices candleOpen, orderPrice, orderPrice, candleLow, , "current", btcVolume + theVolume
        ctrlVolume.sizeBar lastCandle, ctrlCandle(lastCandle).candleWidth, ctrlCandle(lastCandle).Left, btcVolume + theVolume, ctrlCandle(lastCandle).lastColor
    ElseIf orderPrice < candleLow Then
        ctrlCandle(lastCandle).Height = (candleHigh - orderPrice) * priceRatio
        ctrlCandle(lastCandle).Top = ((highestPrice - candleHigh) * priceRatio)
        
        ctrlCandle(lastCandle).setPrices candleOpen, orderPrice, candleHigh, orderPrice, , "current", btcVolume + theVolume
        ctrlVolume.sizeBar lastCandle, ctrlCandle(lastCandle).candleWidth, ctrlCandle(lastCandle).Left, btcVolume + theVolume, ctrlCandle(lastCandle).lastColor
    Else
        ctrlCandle(lastCandle).Height = (candleHigh - candleLow) * priceRatio
        ctrlCandle(lastCandle).Top = ((highestPrice - candleHigh) * priceRatio)
        
        ctrlCandle(lastCandle).setPrices candleOpen, orderPrice, candleHigh, candleLow, , "current", btcVolume + theVolume
        ctrlVolume.sizeBar lastCandle, ctrlCandle(lastCandle).candleWidth, ctrlCandle(lastCandle).Left, btcVolume + theVolume, ctrlCandle(lastCandle).lastColor
    End If
    
    
    
    
    UserControl.Parent.Caption = currentMarketSymbol & "  " & CStr(orderPrice) & "   H:" & CStr(highestPrice) & "   L:" & CStr(lowestPrice) & "   V:" & CStr(myVolume)
    
End Sub


