Attribute VB_Name = "ASM"
' ASM BAS


Option Base 1  ' Arrays base 1
DefLng A-W     ' All variable Long
DefSng X-Z     ' unless singles
               ' unless otherwise defined
'-----------------------------------------------------------------------------

' For calling machine code
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
(ByVal lpMCode As Long, _
ByVal Long1 As Long, ByVal Long2 As Long, _
ByVal Long3 As Long, ByVal Long4 As Long) As Long
'-----------------------------------------------------------------
Public ptrMC, ptrStruc        ' Ptrs to Machine Code & Structure

'MCode Structure
Public Type PStruc
   picWd As Long
   picHt As Long
   Ptrpic1mem As Long      ' pic1mem(1-4,1-picWd,1-picHt) bytes
   Ptrpic2mem As Long      ' pic2mem(1-4,1-picWd,1-picHt) bytes
   PtrIArr As Long         ' IArr(1-picWd,1-picHt)
   BlackThreshHold As Long  ' 0-255
   DithDIV As Long          ' FSADIV,FSBDIV,StuckiDiv, etc
   greysum As Long         ' Output test Also AvGrey
'---------------------------------
   NX As Long
   NY As Long
   NT As Long
   PtrLimits As Long       ' Limits(1-NT)
   PtrTM As Long           ' TM(1-NX, 1-NY, 1-NT) bytes
   RandSeed As Long        ' (0-.99999) * 255
   OpCode As Long
End Type
Public DStruc As PStruc

Public DitherMC() As Byte

Public Sub ASM_BlackWhite()
DStruc.OpCode = 0
res = CallWindowProc(ptrMC, ptrStruc, 2&, 3&, 4&)
'Stop
End Sub

Public Sub ASM_FloydSteinbergA()
DStruc.OpCode = 1
res = CallWindowProc(ptrMC, ptrStruc, 2&, 3&, 4&)
'Stop
End Sub

Public Sub ASM_FloydSteinbergB()
DStruc.OpCode = 2
res = CallWindowProc(ptrMC, ptrStruc, 2&, 3&, 4&)
'Stop
End Sub

Public Sub ASM_Stucki()
DStruc.OpCode = 3
res = CallWindowProc(ptrMC, ptrStruc, 2&, 3&, 4&)
'Stop
End Sub

Public Sub ASM_Burkes()
DStruc.OpCode = 4
res = CallWindowProc(ptrMC, ptrStruc, 2&, 3&, 4&)
'Stop
End Sub

Public Sub ASM_Sierra()
DStruc.OpCode = 5
res = CallWindowProc(ptrMC, ptrStruc, 2&, 3&, 4&)
'Stop
End Sub

Public Sub ASM_Jarvis()
DStruc.OpCode = 6
res = CallWindowProc(ptrMC, ptrStruc, 2&, 3&, 4&)
'Stop
End Sub

Public Sub ASM_DOPATTERN()
DStruc.OpCode = 7
res = CallWindowProc(ptrMC, ptrStruc, 2&, 3&, 4&)
'Stop
End Sub

Public Sub ASM_Random()
DStruc.OpCode = 8
res = CallWindowProc(ptrMC, ptrStruc, 2&, 3&, 4&)
'Stop
End Sub

' TEST
'Public Sub ASM_FillIntensityArray()
'DStruc.PtrIArr = VarPtr(IArr(1, 1))
'DStruc.OpCode = 1
'res = CallWindowProc(ptrMC, ptrStruc, 2&, 3&, 4&)
''Stop
'End Sub

Public Sub Loadmcode(InFile$, MCCode() As Byte)
'Load machine code into InCode() byte array
On Error GoTo InFileErr
If Dir$(InFile$) = "" Then
   MsgBox InFile$ & " missing", , "Dither"
   DoEvents
   Unload Form1
   End
End If
Open InFile$ For Binary As #1
MCSize& = LOF(1)
If MCSize& = 0 Then
InFileErr:
   MsgBox InFile$ & " missing", , "Dither"
   DoEvents
   Unload Form1
   End
End If
ReDim MCCode(MCSize&)
Get #1, , MCCode
Close #1
On Error GoTo 0
End Sub

