Attribute VB_Name = "ScrollBar"
' ScrollBar.bas  by  Robert Rayment

Option Explicit

Public frm As Form1

Public BarIndex  As Integer
Public nTVal() As Long  ' Integerised value - Output
Public scrMax() As Long
Public scrMin() As Long
Public SmallChange() As Integer
Public LargeChange() As Integer

' Assumes Names:-

' frm.LabBack(0,1)
' frm.LabThumb(0,1)
' frm.fraScr(0,1)
' frm.cmdScr(0,1/2,3)

' Set up in Form1:--
'ReDim scrMax(0 To 1)
'ReDim scrMin(0 To 1)
'ReDim SmallChange(0 To 1)
'ReDim LargeChange(0 To 1)

''--- Vertical Scroll Bar Settings -------------
'scrMax(0) = 20: scrMin(0) = 1
'SmallChange(0) = 1: LargeChange(0) = 7 '2 ' NB > 0
'
''--- Horizontal Scroll Bar Settings -----------
'scrMax(1) = 6: scrMin(1) = 0
'SmallChange(1) = 1: LargeChange(1) = 2 ' NB > 0
''----------------------------------------------
'############################################################

Private scrRange() As Long     ' max-min

Private zScale() As Single     ' Converting numerical<->Twips
Private LabBackLimit() As Integer ' Thumb position limit
Private LabBackTL() As Integer ' = LabBack(Index).Top/Left

Private scrIndex As Integer    ' 0,1 (Vert) 2,3 (Horz) cmdScr buttons
Private ThumbSize() As Integer ' Thumb label height or width
'Private BarIndex As Integer    ' 0 Vert, 1 Horz
Private SStep As Long          ' Takes on Small or Large change

Private ThumbLoca() As Integer ' Y vert, X horz
Private XYLoca() As Single     ' Y vert, X horz
Private zTVal() As Single      ' Numeric value
Private TestLoca() As Integer  ' Test moved Y or X
Private ThumbDrag As Boolean   ' Thumb moving

' Call this after setting up max mins

Public Sub ModReDim()

ReDim scrRange(1)
ReDim zScale(1)
ReDim LabBackLimit(1)
ReDim LabBackTL(1)

ReDim ThumbSize(1)

ReDim ThumbLoca(1)
ReDim XYLoca(1)
ReDim zTVal(1)
ReDim nTVal(1)
ReDim TestLoca(1)

frm.Timer1.Enabled = False
frm.Timer1.Interval = 60


For SStep = 0 To 3
   frm.cmdScr(SStep).FontName = "WingDings"
   frm.cmdScr(SStep).FontSize = 8
Next SStep

' Basic sizing of vertical scrollbar
frm.cmdScr(0).Width = 15 * frm.fraScr(0).Width
frm.cmdScr(1).Width = frm.cmdScr(0).Width
frm.LabBack(0).Width = frm.cmdScr(0).Width - 4
frm.LabThumb(0).Width = frm.LabBack(0).Width - 4

frm.cmdScr(0).Height = frm.cmdScr(0).Width '/ 3
frm.cmdScr(1).Height = frm.cmdScr(0).Height
If 15 * frm.fraScr(0).Height <= 2 * frm.cmdScr(0).Height Then
   MsgBox ("Vert Frame size too small")
   End
End If
frm.LabBack(0).Height = 15 * frm.fraScr(0).Height - _
   2 * frm.cmdScr(0).Height

frm.cmdScr(0).Left = 0
frm.cmdScr(1).Left = 0
frm.LabBack(0).Left = 0
frm.LabThumb(0).Left = 0

frm.cmdScr(0).Top = 0
frm.LabBack(0).Top = frm.cmdScr(0).Height
frm.cmdScr(1).Top = 15 * frm.fraScr(0).Height - _
   frm.cmdScr(1).Height

' Basic sizing of horizontal scrollbar
frm.cmdScr(2).Height = 15 * frm.fraScr(1).Height
frm.cmdScr(3).Height = frm.cmdScr(2).Height
frm.LabBack(1).Height = frm.cmdScr(2).Height - 4
frm.LabThumb(1).Height = frm.LabBack(1).Height - 4

frm.cmdScr(2).Width = frm.cmdScr(2).Height '/ 3
frm.cmdScr(3).Width = frm.cmdScr(2).Width
If 15 * frm.fraScr(1).Width <= 2 * frm.cmdScr(2).Height Then
   MsgBox ("Horz Frame size too small")
   End
End If
frm.LabBack(1).Width = 15 * frm.fraScr(1).Width - _
   2 * frm.cmdScr(2).Width

frm.cmdScr(2).Top = 0
frm.cmdScr(3).Top = 0
frm.LabBack(1).Top = 0
frm.LabThumb(1).Top = 0

frm.cmdScr(2).Left = 0
frm.LabBack(1).Left = frm.cmdScr(2).Width
frm.cmdScr(3).Left = 15 * frm.fraScr(1).Width - _
   frm.cmdScr(3).Width


'  Set Ranges, Thumbsize, Limits & Windings

'----------------------------------------
'----------------------------------------
'--- Vertical Scroll bar ----------------
'----------------------------------------
'----------------------------------------

scrRange(0) = scrMax(0) - scrMin(0) ' NB <> 0
If Abs(scrRange(0)) < SmallChange(0) Or Abs(scrRange(0)) < LargeChange(0) Then
   MsgBox "Change(0) larger than Range(0)", , ""
   Unload frm
   End
End If

ThumbSize(0) = frm.LabBack(0).Height / Abs(scrRange(0))
If ThumbSize(0) < 150 Then ThumbSize(0) = 150   ' Limit min thumb size
frm.LabThumb(0).Height = ThumbSize(0)

zScale(0) = (frm.LabBack(0).Height - frm.LabThumb(0).Height) / (scrMax(0) - scrMin(0))
LabBackLimit(0) = frm.LabBack(0).Top + frm.LabBack(0).Height - frm.LabThumb(0).Height
LabBackTL(0) = frm.LabBack(0).Top

ThumbLoca(0) = frm.LabBack(0).Top
frm.LabThumb(0).Top = ThumbLoca(0)
TestLoca(0) = ThumbLoca(0)
ThumbDrag = False

nTVal(0) = scrMin(0)
' WingDings
frm.cmdScr(0).Caption = Chr$(217)
frm.cmdScr(1).Caption = Chr$(218)

'----------------------------------------
'----------------------------------------
'--- Horizontal Scroll Bar --------------
'----------------------------------------
'----------------------------------------

scrRange(1) = scrMax(1) - scrMin(1) ' NB <> 0
If Abs(scrRange(1)) < SmallChange(1) Or Abs(scrRange(1)) < LargeChange(1) Then
   MsgBox "Change(1) larger than Range(1)", , ""
   Unload frm
   End
End If

ThumbSize(1) = frm.LabBack(1).Width / Abs(scrRange(1))
If ThumbSize(1) < 150 Then ThumbSize(1) = 150   ' Limit min thumb size
frm.LabThumb(1).Width = ThumbSize(1)

zScale(1) = (frm.LabBack(1).Width - frm.LabThumb(1).Width) / (scrMax(1) - scrMin(1))
LabBackLimit(1) = frm.LabBack(1).Left + frm.LabBack(1).Width - frm.LabThumb(1).Width
LabBackTL(1) = frm.LabBack(1).Left

ThumbLoca(1) = frm.LabBack(1).Left
frm.LabThumb(1).Left = ThumbLoca(1)
TestLoca(1) = ThumbLoca(1)
ThumbDrag = False

nTVal(1) = scrMin(1)
' WingDings
frm.cmdScr(2).Caption = Chr$(215)
frm.cmdScr(3).Caption = Chr$(216)

'----------------------------------------
'----------------------------------------
'----------------------------------------
'----------------------------------------
End Sub

Public Sub LabThumbMouseDown(Index As Integer, x As Single, y As Single)

If Index = 0 Then XYLoca(Index) = y Else XYLoca(Index) = x
ThumbDrag = True

End Sub

Public Sub LabThumbMouseMove(Index As Integer, _
x As Single, y As Single)

If ThumbDrag = True Then

   Select Case Index
   Case 0   ' Vertical ScrollBar Thumb
      TestLoca(Index) = ThumbLoca(Index) + (y - XYLoca(Index))
   Case 1   ' Horizontal ScrollBar Thumb
      TestLoca(Index) = ThumbLoca(Index) + (x - XYLoca(Index))
   End Select

   If TestLoca(Index) >= LabBackTL(Index) Then
   
      If TestLoca(Index) <= LabBackLimit(Index) Then
                  
         If Index = 0 Then
            frm.LabThumb(Index).Top = TestLoca(Index)
         Else
            frm.LabThumb(Index).Left = TestLoca(Index)
         End If
         
         ThumbLoca(Index) = TestLoca(Index)
   
         zTVal(Index) = 1# * (TestLoca(Index) - LabBackTL(Index)) / zScale(Index)
   
         nTVal(Index) = Int(zTVal(Index) + 1.5) + scrMin(Index) - 1
         
      End If
   
   Else
      TestLoca(Index) = ThumbLoca(Index)
   End If

End If

End Sub

Public Sub LabThumbMouseUp(Index As Integer, _
x As Single, y As Single)

' Snap thumb position
ThumbDrag = False

TestLoca(Index) = 1# * (nTVal(Index) - scrMin(Index)) * zScale(Index) _
+ LabBackTL(Index) + 0.5

If TestLoca(Index) > LabBackLimit(Index) Then
   TestLoca(Index) = LabBackLimit(Index)
End If

If Index = 0 Then    ' Vert
   frm.LabThumb(Index).Top = TestLoca(Index)
Else  ' Horz
   frm.LabThumb(Index).Left = TestLoca(Index)
End If

End Sub

'###################################################################
'###################################################################

Public Sub LabBackMouseDown(Index As Integer, _
x As Single, y As Single)

If Index = 0 Then XYLoca(Index) = y Else XYLoca(Index) = x
   
BarIndex = Index

Select Case Index
Case 0   ' Vertical ScrollBar
   
   If XYLoca(Index) + frm.LabBack(Index).Top > _
   frm.LabThumb(Index).Top + frm.LabThumb(Index).Height Then
      scrIndex = 1   ' Down larger values
      SStep = LargeChange(Index)
      frm.Timer1.Enabled = True
   ElseIf XYLoca(Index) + frm.LabBack(Index).Top < frm.LabThumb(Index).Top Then
      scrIndex = 0   ' Up smaller values
      SStep = LargeChange(Index)
      frm.Timer1.Enabled = True
   Else
      frm.Timer1.Enabled = False
   End If

Case 1   ' Horizontal ScrollBar
   
   If XYLoca(Index) + frm.LabBack(Index).Left > _
   frm.LabThumb(Index).Left + frm.LabThumb(Index).Width Then
      scrIndex = 3   ' Right larger values
      SStep = LargeChange(Index)
      frm.Timer1.Enabled = True
   ElseIf XYLoca(Index) + frm.LabBack(Index).Left < frm.LabThumb(Index).Left Then
      scrIndex = 2   ' Left smaller values
      SStep = LargeChange(Index)
      frm.Timer1.Enabled = True
   Else
      frm.Timer1.Enabled = False
   End If

End Select

End Sub

Public Sub LabBackMouseUp()

frm.Timer1.Enabled = False

End Sub

'###################################################################
'###################################################################

Public Sub cmdScrMouseDown(Index As Integer, Button As Integer, _
x As Single, y As Single)

scrIndex = Index     ' 0,1 vert 2,3 horz
BarIndex = Index \ 2 ' 0 vert, 1 horz
SStep = SmallChange(BarIndex)
If Button = 2 Then
   If scrIndex = 0 Or scrIndex = 2 Then   ' Top & Left cmdScr
      
      nTVal(BarIndex) = scrMin(BarIndex) + Sgn(scrRange(BarIndex)) * SStep
      
   Else        ' Bottom & Right cmdScr
      
      nTVal(BarIndex) = scrMax(BarIndex) - Sgn(scrRange(BarIndex)) * SStep
      
   End If
End If
frm.Timer1.Enabled = True

End Sub

Public Sub cmdScrMouseUp()

frm.Timer1.Enabled = False

End Sub

'###################################################################
'###################################################################

Public Sub Timer1Timer()

'In:  nTVal, TestLoca, SStep, scrMin, scrMax, zScale, BarIndex, LabBackLimit
'Out: new nTVal, ThumbLoca

If scrRange(BarIndex) > 0 Then

   Select Case scrIndex
   Case 0, 2  ' Up/Left -> scrMin
      If nTVal(BarIndex) >= scrMin(BarIndex) + SStep Then
         nTVal(BarIndex) = nTVal(BarIndex) - SStep
      Else
         nTVal(BarIndex) = scrMin(BarIndex)
      End If
   Case 1, 3  ' Down/Right -> scrMax
      If nTVal(BarIndex) <= scrMax(BarIndex) - SStep Then
         nTVal(BarIndex) = nTVal(BarIndex) + SStep
      Else
         nTVal(BarIndex) = scrMax(BarIndex)
      End If
   End Select

Else    ' scrMax(BarIndex) < scrMin(BarIndex)
   
   Select Case scrIndex
   Case 0, 2  ' Up/Left -> scrMin
      If nTVal(BarIndex) <= scrMin(BarIndex) - SStep Then
         nTVal(BarIndex) = nTVal(BarIndex) + SStep
      Else
         nTVal(BarIndex) = scrMin(BarIndex)
      End If
   Case 1, 3  ' Down/Right -> scrMax
      If nTVal(BarIndex) >= scrMax(BarIndex) + SStep Then
         nTVal(BarIndex) = nTVal(BarIndex) - SStep
      Else
         nTVal(BarIndex) = scrMax(BarIndex)
      End If
   End Select

End If

' Fix Vert Or Horz Thumb position

TestLoca(BarIndex) = 1# * (nTVal(BarIndex) - scrMin(BarIndex)) * zScale(BarIndex) _
+ LabBackTL(BarIndex) + 0.5 '''''''

If TestLoca(BarIndex) > LabBackLimit(BarIndex) Then
   TestLoca(BarIndex) = LabBackLimit(BarIndex)
End If

If BarIndex = 0 Then ' Vert
   frm.LabThumb(BarIndex).Top = TestLoca(BarIndex)
Else  ' BarIndex = 1   Horz
   frm.LabThumb(BarIndex).Left = TestLoca(BarIndex)
End If

ThumbLoca(BarIndex) = TestLoca(BarIndex)

End Sub


