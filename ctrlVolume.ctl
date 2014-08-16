VERSION 5.00
Begin VB.UserControl ctrlVolume 
   BackColor       =   &H00000000&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5265
   ScaleHeight     =   1275
   ScaleWidth      =   5265
   Begin VB.Line lineBottom 
      BorderColor     =   &H00FFFFFF&
      Visible         =   0   'False
      X1              =   1080
      X2              =   2760
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblVolume 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      ForeColor       =   &H0000C000&
      Height          =   765
      Index           =   0
      Left            =   0
      MousePointer    =   2  'Cross
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   435
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
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   780
   End
End
Attribute VB_Name = "ctrlVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const graphSections = 2

Private volumeValues() As Double
Private highestPrice As Double
Private lowestPrice As Double

Public Sub addBar(theVolume As Double, theColor As ColorConstants, leftPosition As Long, barWidth As Long)
 Dim i As Long, candleCount As Long, chartRange As Double, priceRatio As Double, tempValue As Double
    
    candleCount = lblVolume.UBound + 1
    Load lblVolume(candleCount)
    ReDim Preserve volumeValues(candleCount)
    volumeValues(candleCount) = theVolume
    lblVolume(candleCount).Left = leftPosition
    lblVolume(candleCount).Width = barWidth
    lblVolume(candleCount).BackColor = theColor
    
    lblVolume(candleCount).Visible = True
End Sub

Public Sub sizeBar(barIndex As Long, barWidth As Long, barLeft As Long, Optional theVolume As Double = 0, Optional theColor As ColorConstants = 0)
 Dim volumeRatio As Double, barHeight As Double
    
    volumeRatio = UserControl.Height / highestPrice
    If theVolume > 0 Then volumeValues(barIndex) = theVolume
    barHeight = volumeValues(barIndex) * volumeRatio
    lblVolume(barIndex).Height = barHeight
    lblVolume(barIndex).Top = UserControl.Height - barHeight
    lblVolume(barIndex).Width = barWidth
    lblVolume(barIndex).Left = barLeft
    If theColor <> 0 Then lblVolume(barIndex).BackColor = theColor
End Sub

Public Sub clearVolume()
 Dim i As Long
    
    If lblVolume.UBound > 0 Then
        For i = 1 To lblVolume.UBound
            Unload lblVolume(i)
        Next i
    End If
End Sub

Public Sub formatVolume()
 Dim upperRange As Double, lowerRange As Double, rangeDifference As Double, i As Long, highPriceString As String, _
     lowPriceString As String, sizeRatio As Double, tempValue As Double, timesTen As Double, tempLong As Long, _
     highDotPosition As Long, lowDotPosition As Long, tempHighVal As Double, tempLowVal As Double, tempLowLen As Long, _
     tempHighLen As Long, priceRatio As Double, highPrice As Double, lowPrice As Double
    
    highPrice = 0
    
    For i = 1 To UBound(volumeValues())
        If volumeValues(i) > highPrice Then highPrice = volumeValues(i)
    Next i
    
    highPriceString = CStr(highPrice)
    lowPrice = 0
    lowPriceString = "0"
    
    
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
    sizeRatio = UserControl.Height / graphSections
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
    linePrices(0).Y1 = candleSpacing
    linePrices(0).Y2 = candleSpacing
    
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
    
    lblPrices(graphSections).Top = UserControl.Height - lblPrices(graphSections).Height
    tempValue = UserControl.Height - candleSpacing
    lineBottom.X1 = 0
    lineBottom.X2 = UserControl.Width
    lineBottom.Y1 = tempValue
    lineBottom.Y2 = tempValue
    lineBottom.Visible = True
    
    For i = 1 To lblVolume.UBound
        sizeBar i, lblVolume(i).Width, lblVolume(i).Left
    Next i
    
End Sub
