Attribute VB_Name = "Publics"
' Publics.bas

Option Base 1

DefLng A-W
DefSng X-Z

Public red As Byte, green As Byte, blue As Byte

Public BlackThreshHold   ' default=128
Public FSADIV     ' Floyd-SteinbergA divisor, default = 8
Public FSBDIV     ' Floyd-SteinbergB divisor, default = 16
Public StuckiDiv  ' default = 42
Public BurkesDiv  ' default = 32
Public SierraDiv  ' default = 32
Public JarvisDiv  ' default = 48
Public HalftoneLimits()    ' 25,50,75,100,125,150,175,200,230
Public RandomLimits()      ' 25,50,75,100,125,150,175,200,230
Public LineLimits()        ' 60,120,180
Public LineLimits2()       ' 25,50,75,100,125,150,175,200,230

Public pic1mem() As Byte   ' For Loaded picture
Public pic2mem() As Byte   ' For Output dither picture
Public picHt, picWd        ' Picture Height & Width
Public IArr()        ' Error diffusion array
Public greysum       ' Average grey value
Public BScanLine     ' For saving file

Public NX, NY, NT    ' Pattern dims
Public TM() As Byte  ' Pattern array
Public TS$()         ' Pattern string array
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Ptrpic1mem '= VarPtr(pic1mem(1, 1, 1))
Public Ptrpic2mem '= VarPtr(pic2mem(1, 1, 1))
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' VB<->ASMVB switch
Public ASMVB As Boolean


Public Sub LngToRGB(LCul)
'Convert Long Colors() to RGB components
red = (LCul And &HFF&)
green = (LCul And &HFF00&) / &H100&
blue = (LCul And &HFF0000) / &H10000
End Sub

Public Sub SetDefaultFactors()
BlackThreshHold = 128
FSADIV = 8
FSBDIV = 16
StuckiDiv = 42
BurkesDiv = 32
SierraDiv = 42
JarvisDiv = 48

ReDim HalftoneLimits(9)
HalftoneLimits(1) = 25
HalftoneLimits(2) = 50
HalftoneLimits(3) = 75
HalftoneLimits(4) = 100
HalftoneLimits(5) = 125
HalftoneLimits(6) = 150
HalftoneLimits(7) = 175
HalftoneLimits(8) = 200
HalftoneLimits(9) = 230

ReDim RandomLimits(9)
RandomLimits(1) = 25
RandomLimits(2) = 50
RandomLimits(3) = 75
RandomLimits(4) = 100
RandomLimits(5) = 125
RandomLimits(6) = 150
RandomLimits(7) = 175
RandomLimits(8) = 200
RandomLimits(9) = 230

ReDim LineLimits(3)
LineLimits(1) = 60
LineLimits(2) = 120
LineLimits(3) = 180

ReDim LineLimits2(9)
LineLimits2(1) = 25
LineLimits2(2) = 50
LineLimits2(3) = 75
LineLimits2(4) = 100
LineLimits2(5) = 125
LineLimits2(6) = 150
LineLimits2(7) = 175
LineLimits2(8) = 200
LineLimits2(9) = 230
End Sub

Public Function ExtractFileName$(FSpec$)
ExtractFileName$ = " "
'In:  FSpec$ = Full FileSpec
'Out: FileName =ExtractFileName$
If FSpec$ = "" Then Exit Function

For i = Len(FSpec$) To 1 Step -1
   If Mid$(FSpec$, i, 1) = "\" Then
      ExtractFileName$ = Mid$(FSpec$, i + 1)
      Exit For
   End If
Next i
End Function

