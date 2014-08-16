VERSION 5.00
Object = "{FF0A3CE0-D4CD-11D3-9130-00105A17B608}#1.0#0"; "DartSecure2.dll"
Begin VB.Form frmChart 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   Caption         =   "Yes! CryptoCurrency!"
   ClientHeight    =   6915
   ClientLeft      =   10920
   ClientTop       =   1605
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmChart.frx":0000
   MousePointer    =   4  'Icon
   ScaleHeight     =   6915
   ScaleWidth      =   10050
   Begin VB.Timer tmrHideChart 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   9625
      Top             =   1680
   End
   Begin VB.Timer tmrMintpalRequestPause 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9625
      Top             =   1260
   End
   Begin VB.Timer tmrChartData 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   9625
      Top             =   840
   End
   Begin VB.Timer tmrRealtime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9625
      Top             =   420
   End
   Begin prjChart.ctrlChart priceChart 
      Height          =   6915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10035
      _extentx        =   17701
      _extenty        =   11986
      chartsections   =   10
   End
   Begin VB.Label lblLoading 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "loading chart..."
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   54
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   240
      MouseIcon       =   "frmChart.frx":0152
      MousePointer    =   4  'Icon
      TabIndex        =   1
      Top             =   60
      Width           =   9600
   End
   Begin DartSecureCtl.SecureTcp tcpChart 
      Left            =   9660
      OleObjectBlob   =   "frmChart.frx":02A4
      Top             =   0
   End
   Begin VB.Menu mnuMarkets 
      Caption         =   "markets"
      Begin VB.Menu mnuMintpal 
         Caption         =   "mintpal"
         Begin VB.Menu mnuBTCSymbol 
            Caption         =   "btc"
            Index           =   0
         End
         Begin VB.Menu mnuDash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLTCSymbol 
            Caption         =   "ltc"
            Index           =   0
         End
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMintpalTime 
         Caption         =   "time period"
         Begin VB.Menu mnuMintpal6hours 
            Caption         =   "6 hours"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuMintpal1day 
            Caption         =   "1 day"
         End
         Begin VB.Menu mnuMintpal3days 
            Caption         =   "3 days"
         End
         Begin VB.Menu mnuMintpal7days 
            Caption         =   "1 week"
         End
         Begin VB.Menu mnuMintpalMAX 
            Caption         =   "max"
         End
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "window"
         Begin VB.Menu mnuStayOnTop 
            Caption         =   "keep on top"
         End
         Begin VB.Menu mnuRemoveTitlebar 
            Caption         =   "remove titlebar"
         End
      End
   End
End
Attribute VB_Name = "frmChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintpalMarketID() As String
Private mintpalSymbol() As String

Private currentProcess As String

Private socketData As String
Private chunkedResponse As Boolean

Private chartPeriod As String
Private currentMarket As String
Private currentMarketSymbol As String
Private currentMarketID As String
Private lastCandleTime As String
Private currentTimestampPeriod As Long
Private currentTimeInterval As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1
Private Const WS_THICKFRAME = &H40000



Private Sub Form_Click()
    Me.BorderStyle = 0
    Me.Caption = Me.Caption
End Sub

Private Sub Form_Load()
    
    mnuMarkets.Visible = False
    
    chartPeriod = "6hh"
    
    currentProcess = "mintpal market summary"
    tcpChart.Connect "api.mintpal.com"
    Me.Visible = False
    
 End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then priceChart.clearChart: Exit Sub
    
    'priceChart.addCandle Val(txtOpen.Text), Val(txtClose.Text), Val(txtLow.Text), Val(txtHigh.Text)
End Sub

Private Sub Form_Resize()
 Dim labelHeights As Long
    
    labelHeights = priceChart.labelHeights
    
    If Me.WindowState <> 1 And Me.scaleHeight > priceChart.Top And Me.ScaleWidth > priceChart.Left And Me.Visible = True Then
        If Me.ScaleWidth < 2835 Then Me.Width = 3075
        If Me.scaleHeight < labelHeights + priceChart.Top Then Me.Height = labelHeights + priceChart.Top + (Me.Height - Me.scaleHeight) + 200
        priceChart.Height = Me.scaleHeight - priceChart.Top
        priceChart.Width = Me.ScaleWidth - priceChart.Left
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuBTCSymbol_Click(Index As Integer)
    mnuMintpalTime.Visible = True
    
    currentMarketID = mintpalMarketID(Index)
    currentMarketSymbol = mintpalSymbol(Index)
    currentMarket = "BTC"
    currentProcess = "mintpal timestamp"
    
    priceChart.clearChart
    
    tmrMintpalRequestPause.Enabled = True
End Sub

Private Sub mnuLTCSymbol_Click(Index As Integer)
    mnuMintpalTime.Visible = True
    
    currentMarketID = mintpalMarketID(Index)
    currentMarketSymbol = mintpalSymbol(Index)
    currentMarket = "LTC"
    currentProcess = "mintpal timestamp"
    
    priceChart.clearChart
    
    tmrMintpalRequestPause.Enabled = True
End Sub

Private Sub mnuMintpal6hours_Click()
 Dim checkedValue As Boolean
    
    checkedValue = mnuMintpal6hours.Checked
    mnuMintpal6hours.Checked = True
    mnuMintpal1day.Checked = False
    mnuMintpal3days.Checked = False
    mnuMintpal7days.Checked = False
    mnuMintpalMAX.Checked = False
    chartPeriod = "6hh"
    If checkedValue = False And Len(currentMarketSymbol) <> 0 Then
        currentProcess = "mintpal timestamp"
        tmrMintpalRequestPause.Enabled = True
    End If
End Sub

Private Sub mnuMintpal1day_Click()
 Dim checkedValue As Boolean
    
    checkedValue = mnuMintpal1day.Checked
    mnuMintpal6hours.Checked = False
    mnuMintpal1day.Checked = True
    mnuMintpal3days.Checked = False
    mnuMintpal7days.Checked = False
    mnuMintpalMAX.Checked = False
    chartPeriod = "1DD"
    If checkedValue = False And Len(currentMarketSymbol) <> 0 Then
        currentProcess = "mintpal timestamp"
        tmrMintpalRequestPause.Enabled = True
    End If
End Sub

Private Sub mnuMintpal3days_Click()
 Dim checkedValue As Boolean
    
    checkedValue = mnuMintpal3days.Checked
    mnuMintpal6hours.Checked = False
    mnuMintpal1day.Checked = False
    mnuMintpal3days.Checked = True
    mnuMintpal7days.Checked = False
    mnuMintpalMAX.Checked = False
    chartPeriod = "3DD"
    If checkedValue = False And Len(currentMarketSymbol) <> 0 Then
        currentProcess = "mintpal timestamp"
        tmrMintpalRequestPause.Enabled = True
    End If
End Sub

Private Sub mnuMintpal7days_Click()
 Dim checkedValue As Boolean
    
    checkedValue = mnuMintpal7days.Checked
    mnuMintpal6hours.Checked = False
    mnuMintpal6hours.Checked = False
    mnuMintpal1day.Checked = False
    mnuMintpal3days.Checked = False
    mnuMintpal7days.Checked = True
    mnuMintpalMAX.Checked = False
    chartPeriod = "7DD"
    If checkedValue = False And Len(currentMarketSymbol) <> 0 Then
        currentProcess = "mintpal timestamp"
        tmrMintpalRequestPause.Enabled = True
    End If
End Sub

Private Sub mnuMintpalMAX_Click()
 Dim checkedValue As Boolean
    
    checkedValue = mnuMintpalMAX.Checked
    mnuMintpal6hours.Checked = False
    mnuMintpal6hours.Checked = False
    mnuMintpal1day.Checked = False
    mnuMintpal3days.Checked = False
    mnuMintpal7days.Checked = False
    mnuMintpalMAX.Checked = True
    chartPeriod = "MAX"
    If checkedValue = False And Len(currentMarketSymbol) <> 0 Then
        currentProcess = "mintpal timestamp"
        tmrMintpalRequestPause.Enabled = True
    End If
End Sub
Private Sub mnuRemoveTitlebar_Click()
 Dim currentStyle As Long
    
    currentStyle = GetWindowLong(Me.hWnd, GWL_STYLE)
    
    If mnuRemoveTitlebar.Checked = True Then
        mnuRemoveTitlebar.Checked = False
        SetWindowLong Me.hWnd, GWL_STYLE, currentStyle Or WS_CAPTION
        
    Else
        mnuRemoveTitlebar.Checked = True
        currentStyle = currentStyle And Not WS_CAPTION
        SetWindowLong Me.hWnd, GWL_STYLE, currentStyle
        SetWindowLong Me.hWnd, GWL_STYLE, currentStyle Or WS_THICKFRAME
    End If
    
    Const swpFlags As Long = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE
    
    SetWindowPos Me.hWnd, 0, 0, 0, 0, 0, swpFlags

End Sub

Private Sub mnuStayOnTop_Click()
    If mnuStayOnTop.Checked = True Then
        mnuStayOnTop.Checked = False
        makeNormal Me.hWnd
    Else
        mnuStayOnTop.Checked = True
        stayOnTop Me.hWnd
    End If
End Sub

Private Sub tcpChart_Receive()
 Dim ByteString() As Byte, tempstring As String, ff As Long
    tcpChart.Receive tempstring
    socketData = socketData & tempstring
    
'    ff = FreeFile
'    Open App.Path & "\serverresponse.txt" For Binary As #ff
'        Put #ff, LOF(ff), tempstring
'    Close #ff
    
End Sub

Private Sub tcpChart_State()
 Dim tempstring As String, splitResponse() As String, ff As Long, processedData As String
    
    With tcpChart
        If .State = tcpConnected Then
            
            socketData = ""
            chunkedResponse = False
'            ff = FreeFile
'            Open App.Path & "\serverresponse.txt" For Output As #ff
'                Print #ff, " "
'            Close #ff
'            ff = FreeFile
'            Open App.Path & "\response.log" For Output As #ff
'                Print #ff, " "
'            Close #ff
            tmrRealtime.Enabled = False
            
            Select Case currentProcess
                
                Case "mintpal market summary"
                    
                    .Send marketSummaryPacket '"GET https://api.mintpal.com/v1/market/summary/" & vbCrLf 'marketSummaryPacket
                    
                Case "mintpal chart data"
                    
                    priceChart.clearChart
                    priceChart.Visible = False
                    .Send chartDataPacket(currentMarketID, chartPeriod)
                    
                Case "mintpal timestamp"
                    
                    .Send mintTimestampPacket
                    
                Case "mintpal recent trades"
                    
                    .Send recentTradesPacket
                    
            End Select
            
        ElseIf .State = tcpClosed Then
            
            If InStr(socketData, "Transfer-Encoding: chunked") <> 0 Then
                processedData = Split(socketData, vbCrLf & vbCrLf, 2)(1)
                processedData = processChunks(processedData)
            End If
            
            If InStr(socketData, "Content-Encoding: gzip") <> 0 Then
                processedData = GZip(StrConv(processedData, vbFromUnicode))
            End If
            
            
            Select Case currentProcess
                
                Case "mintpal market summary"
                    
                    processMintpalMarkets processedData
                    
                Case "mintpal chart data"
                    
                    processMintpalChart processedData
                    
                    
                Case "mintpal timestamp"
                    
                    processMintpalTimestamp processedData
                    currentProcess = "mintpal chart data"
                    tmrMintpalRequestPause.Enabled = True
                    
                Case "mintpal recent trades"
                    
                    processMintpalTrades processedData
                    tmrRealtime.Enabled = True
                
            End Select
            
        End If
    End With
End Sub

Private Sub processMintpalTrades(packetData As String)
 Dim tempstring As String, ff As Long
    
    tempstring = Mid(packetData, 3, Len(packetData) - 4)
    
'    ff = FreeFile
'    Open App.Path & "\response.log" For Output As #ff
'        Print #ff, tempstring
'    Close #ff
    
End Sub

Private Sub processMintpalTimestamp(packetData As String)
 Dim tempstring As String, ff As Long, theSplits() As String
    
    tempstring = Mid(packetData, 2, Len(packetData) - 2)
    
'    ff = FreeFile
'    Open App.Path & "\response.log" For Output As #ff
'        Print #ff, tempstring
'    Close #ff
    
    theSplits() = Split(packetData, ":")
    
    currentTimestampPeriod = Int(Val(theSplits(UBound(theSplits()))))
End Sub

Private Sub processMintpalChart(packetData As String)
 Dim tempstring As String, ff As Long, theTimePeriods() As String, theItems() As String, theValues() As String, openPrice As Double, closePrice As Double, highPrice As Double, _
     lowPrice As Double, btcVolume As Double, i As Long, h As Long, theUbound As Long, highestHigh As Double, lowestLow As Double, totalVolume As Double
    
    
    tempstring = Mid(packetData, 3, Len(packetData) - 4)
    
'    ff = FreeFile
'    Open App.Path & "\response.log" For Output As #ff
'        Print #ff, tempstring
'    Close #ff
    
    theTimePeriods() = Split(tempstring, "},{")
    hideMe True
    
    For i = 0 To UBound(theTimePeriods())
        
        tempstring = Mid(theTimePeriods(i), 2, Len(theTimePeriods(i)) - 2)
        theItems() = Split(tempstring, """,""")
        
        openPrice = 0: closePrice = 0: highPrice = 0: lowPrice = 999999999: lowestLow = 999999: btcVolume = 0
        
        For h = 0 To UBound(theItems())
            
            theValues() = Split(theItems(h), """:""", 2)
            
            Select Case theValues(0)
                
                Case "date"
                    tempstring = theValues(1)
                    'tempstring = Right(tempstring, Len(tempstring) - InStr(tempstring, " "))
                    lastCandleTime = tempstring
                    
                Case "open"
                    openPrice = Val(theValues(1))
                    
                Case "close"
                    closePrice = Val(theValues(1))
                    
                Case "high"
                    highPrice = Val(theValues(1))
                    
                Case "low"
                    lowPrice = Val(theValues(1))
                    
                Case "exchange_volume"
                    btcVolume = Val(theValues(1))
                
            End Select
            
        Next h
        
        If Len(theValues(1)) <> 0 Then priceChart.addCandle openPrice, closePrice, lowPrice, highPrice, btcVolume, lastCandleTime
        
        If highPrice > highestHigh Then highestHigh = highPrice
        If lowPrice < lowestLow Then lowestLow = lowPrice
        totalVolume = totalVolume + btcVolume
    Next i
    
    Me.Caption = currentMarketSymbol & "/" & currentMarket & "   last:" & CStr(closePrice) & "   H:" & CStr(highestHigh) & "   L:" & CStr(lowestLow) & "   V:" & CStr(totalVolume)
End Sub

Private Sub processMintpalMarkets(packetData As String)
 Dim tempstring As String, theMarkets() As String, theItems() As String, theValues() As String, marketID As String, theCode As String, theExchange As String, _
     i As Long, h As Long, ff As Integer, theUbound As Long
    
    If packetData = "" Then tmrMintpalRequestPause.Enabled = True: Exit Sub
    tempstring = Mid(packetData, 3, Len(packetData) - 4)
    theMarkets() = Split(tempstring, "},{")
    
'    ff = FreeFile
'    Open App.Path & "\response.log" For Output As #ff
'        Print #ff, tempstring
'    Close #ff
    
    ReDim mintpalMarketID(0): ReDim mintpalSymbol(0)
    
    For i = 0 To UBound(theMarkets())
        
        tempstring = Mid(theMarkets(i), 2, Len(theMarkets(i)) - 2)
        theItems() = Split(tempstring, """,""")
        marketID = "": theCode = "": theExchange = ""
        
        For h = 0 To UBound(theItems())
            theValues() = Split(theItems(h), """:""", 2)
            Select Case theValues(0)
                
                Case "market_id"
                    marketID = theValues(1)
                    
                Case "code"
                    theCode = theValues(1)
                    
                Case "exchange"
                    theExchange = theValues(1)
                
            End Select
            
            If Len(marketID) <> 0 And Len(theCode) <> 0 And Len(theExchange) <> 0 Then Exit For
        Next h
        
        theUbound = UBound(mintpalMarketID()) + 1
        
        ReDim Preserve mintpalMarketID(theUbound)
        ReDim Preserve mintpalSymbol(theUbound)
        
        mintpalMarketID(theUbound) = marketID
        mintpalSymbol(theUbound) = theCode
        
        If theExchange = "BTC" Then
            If mnuBTCSymbol.ubound < theUbound Then Load mnuBTCSymbol(theUbound)
            mnuBTCSymbol(theUbound).Caption = theCode & "/" & "BTC"
            mnuBTCSymbol(theUbound).Visible = True
        Else
            If mnuLTCSymbol.ubound < theUbound Then Load mnuLTCSymbol(theUbound)
            mnuLTCSymbol(theUbound).Caption = theCode & "/" & "LTC"
            mnuLTCSymbol(theUbound).Visible = True
        End If
        
    Next i
    
    If mnuBTCSymbol.ubound > 0 Then mnuBTCSymbol(0).Visible = False
    If mnuLTCSymbol.ubound > 0 Then mnuLTCSymbol(0).Visible = False
    Me.Visible = True
End Sub

Private Function marketSummaryPacket()
 Dim requestPacket As String
    
    requestPacket = "GET /v1/market/summary/ HTTP/1.1" & vbCrLf & _
                    "Host: api.mintpal.com" & vbCrLf & _
                    "Accept-Encoding: gzip,deflate" & vbCrLf & _
                    "Accept: application/json" & vbCrLf & _
                    "Connection: close" & vbCrLf & _
                    "Accept-Language: en-us" & vbCrLf & _
                    "User-Agent: Yes! CryptoCurrency!" & vbCrLf & vbCrLf
                    
    marketSummaryPacket = requestPacket
   
End Function

Private Function chartDataPacket(marketID As String, timePeriod As String)
 Dim requestPacket As String
    
    requestPacket = "GET /v1/market/chartdata/" & marketID & "/" & timePeriod & " HTTP/1.1" & vbCrLf & _
                    "Host: api.mintpal.com" & vbCrLf & _
                    "Accept-Encoding: gzip,deflate" & vbCrLf & _
                    "Accept: application/json" & vbCrLf & _
                    "Connection: close" & vbCrLf & _
                    "Accept-Language: en-us" & vbCrLf & _
                    "User-Agent: Yes! CryptoCurrency!" & vbCrLf & vbCrLf
                    
    chartDataPacket = requestPacket
    
End Function

Private Function mintTimestampPacket()
 Dim requestPacket As String
    
    requestPacket = "GET /timestamp HTTP/1.1" & vbCrLf & _
                    "Host: api.mintpal.com" & vbCrLf & _
                    "Accept-Encoding: gzip,deflate" & vbCrLf & _
                    "Accept: application/json" & vbCrLf & _
                    "Connection: close" & vbCrLf & _
                    "Accept-Language: en-us" & vbCrLf & _
                    "User-Agent: Yes! CryptoCurrency!" & vbCrLf & vbCrLf
                    
    mintTimestampPacket = requestPacket
    
End Function

Private Function recentTradesPacket()
 Dim requestPacket As String
    
    requestPacket = "GET /v1/market/trades/" & currentMarketSymbol & "/" & currentMarket & " HTTP/1.1" & vbCrLf & _
                    "Host: api.mintpal.com" & vbCrLf & _
                    "Accept-Encoding: gzip,deflate" & vbCrLf & _
                    "Accept: application/json" & vbCrLf & _
                    "Connection: close" & vbCrLf & _
                    "Accept-Language: en-us" & vbCrLf & _
                    "User-Agent: Yes! CryptoCurrency!" & vbCrLf & vbCrLf
                    
    recentTradesPacket = requestPacket
    
End Function

Private Function processChunks(theData As String) As String
 Dim crlfLong As Long, lengthValue As Long, tempData As String, processedData As String, hexValue As String
    
    tempData = theData
    crlfLong = InStr(tempData, vbCrLf)
    
    If crlfLong <> 0 Then lengthValue = Val("&H" & Left(tempData, crlfLong - 1))
    
    While crlfLong <> 0
        
        hexValue = Left(tempData, crlfLong - 1)
        lengthValue = Val("&H" & hexValue)
        If lengthValue > 0 Then processedData = processedData & Mid(tempData, crlfLong + 2, lengthValue)
        tempData = Mid(tempData, crlfLong + 2 + lengthValue)
        crlfLong = InStr(tempData, vbCrLf)
        Debug.Print CStr(crlfLong) & vbCrLf & CStr(lengthValue)
    Wend
    
    processChunks = processedData
    
End Function


Private Sub tmrChartData_Timer()
    tmrRealtime.Enabled = False
    
    currentProcess = "mintpal timestamp"
    
End Sub

Private Sub tmrHideChart_Timer()
    tmrHideChart.Enabled = False
    hideMe False
End Sub

Private Sub tmrRealtime_Timer()
    tmrRealtime.Enabled = False
    
'    currentProcess = "mintpal recent trades"
'    If tcpChart.State <> tcpClosed Then tcpChart.Close
'    tcpChart.Connect "api.mintpal.com"
End Sub

Private Sub tmrMintpalRequestPause_Timer()
    tmrMintpalRequestPause.Enabled = False
    If tcpChart.State <> tcpClosed Then tcpChart.Close
    tcpChart.Connect "api.mintpal.com"
End Sub







