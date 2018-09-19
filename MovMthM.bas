Option Compare Database
Option Explicit
Sub Z_MovMth()
MovMth "Fct", "LNK", "LnkM", "LnkCCM"
End Sub
Function MdMthNy_zPfxEpt(A As CodeModule, MthTy$, Optional EptMthTy$) As String()
MdMthNy_zPfxEpt = AyWhPfxEpt(MdMthNy(A), MthTy, EptMthTy)
End Function

Function SrcyMthLno%(A, M)
Dim J%, L
For Each L In A
    J = J + 1
    If LinMthNm(L) = M Then SrcyMthLno = J: Exit Function
Next
Stop
End Function

Function MdMthLines$(A As CodeModule, M)
Dim L%(), O$(), J%
L = MdMthLnoAy(A, M)
For J = 0 To UB(L)
    Push O, MdMthLnoLines(A, L(J))
Next
MdMthLines = JnCrLf(O)
End Function

Function MdMthLnoLines$(A As CodeModule, MthLno%)
MdMthLnoLines = A.Lines(MthLno, MdMthLinCnt(A, MthLno))
End Function

Function LinIsPrp(A) As Boolean
LinIsPrp = HasPfx(LinRmvMdy(A), "Property ")
End Function

Function MdMthLnoAy(A As CodeModule, M) As Integer()
Dim L%, Lin$, O%()
L = MdMthLno(A, M)
If L = 0 Then Exit Function
If Not LinIsPrp(A.Lines(L, 1)) Then MdMthLnoAy = ApIntAy(L): Exit Function
Push O, L
For L = L + 1 To A.CountOfLines
    Lin = A.Lines(L, 1)
    If LinMthNm(Lin) = M Then
        If Sz(O) = 3 Then Exit For
    End If
Next
MdMthLnoAy = O
End Function

Function MdMthNy(A As CodeModule) As String()
If MdIsNoLin(A) Then Exit Function
Dim O$(), J%
For J = 1 To A.CountOfLines
    PushNonEmp O, LinMthNm(A.Lines(J, 1))
Next
MdMthNy = AyDist(O)
End Function

Function MdMthLy(A As CodeModule) As String()
If MdIsNoLin(A) Then Exit Function
Dim O$(), J%
For J = 1 To A.CountOfLines
    If LinIsMthLin(A.Lines(J, 1)) Then
        Push O, MdContLin(A, J)
    End If
Next
MdMthLy = O
End Function

Function LinIsMthLin(A) As Boolean
LinIsMthLin = LinMthKd(A) <> ""
End Function

Function PjMdAy(A As VBProject) As CodeModule()
Dim O() As CodeModule, I
For Each I In A.VBComponents
    PushObj O, CvCmp(I).CodeModule
Next
PjMdAy = O
End Function
Function CmpTy_Str$(A As vbext_ComponentType)
Dim O$
Select Case A
Case vbext_ComponentType.vbext_ct_StdModule: O = "Mod"
Case vbext_ComponentType.vbext_ct_ClassModule: O = "Cls"
Case vbext_ComponentType.vbext_ct_Document: O = "Doc"
Case Else: Stop
End Select
CmpTy_Str = O
End Function
Function CvCmp(A) As VBComponent
Set CvCmp = A
End Function
Function CvMd(A) As CodeModule
Set CvMd = A
End Function
Function MdTyStr$(A As CodeModule)
MdTyStr = CmpTy_Str(A.Parent.Type)
End Function
Function PjMthLy(A As VBProject) As String()
Dim I, O$(), N$, M As CodeModule
For Each I In PjMdAy(A)
    Set M = I
    N = MdTyStr(M) & "." & MdNm(M) & "."
    PushAy O, AyAddPfx(MdMthLy(M), N)
Next
PjMthLy = O
End Function
Function CPjMthLy() As String()
CPjMthLy = PjMthLy(CurPj)
End Function
Function CMdMthLy() As String()
CMdMthLy = MdMthLy(CurMd)
End Function
Function StrApp$(A, L)
If A = "" Then StrApp = L: Exit Function
StrApp = A & " " & L
End Function
Function LinesApp$(A, L)
If A = "" Then LinesApp = L: Exit Function
LinesApp = A & vbCrLf & L
End Function
Function MdContLin$(A As CodeModule, Lno%)
Dim O$, J%
J = Lno
Do
    O = RmvSfx(O, " _") & A.Lines(J, 1)
    J = J + 1
Loop Until LasChr(O) <> "_"
MdContLin = O
End Function
Sub MovMth(FmMd$, MthNmPfx$, ToMd$, Optional EptMthNmPfx$)
Dim Fm As CodeModule
Dim Ny$()
Set Fm = Md(FmMd)
Ny = MdMthNy_zPfxEpt(Fm, MthNmPfx, EptMthNmPfx)
AyDoAXB Ny, "MdMovMth", Fm, Md(ToMd)
End Sub

Sub MdRmvMth(A As CodeModule, M)
Dim L%(), J%, Cnt%
L = MdMthLnoAy(A, M)
For J = UB(L) To 0 Step -1
    Cnt = MdMthLinCnt(A, L(J))
    A.DeleteLines L(J), Cnt
Next
End Sub

Function MthKdAy()
Static O$(), Y As Boolean
If Not Y Then
    Y = True
    Push O, "Function"
    Push O, "Sub"
    Push O, "Property"
End If
MthKdAy = O
End Function
Function MthTyAy()
Static O$(), Y As Boolean
If Not Y Then
    Y = True
    Push O, "Function"
    Push O, "Sub"
    Push O, "Property Get"
    Push O, "Property Set"
    Push O, "Property Let"
End If
MthTyAy = O
End Function

Function AyFstEqV(A, V)
If AyHas(A, V) Then AyFstEqV = V
End Function

Function LinMthKd$(A)
LinMthKd = AyFstEqV(MthKdAy, LinT1(LinRmvMdy(A)))
End Function

Private Sub ZZ_MdMthLno()
Dim O$()
    Dim Lno, L%(), M, A As CodeModule, Ny$(), J%
    Set A = Md("Fct")
    Ny = MdMthNy(A)
    For Each M In Ny
        DoEvents
        J = J + 1
        Push L, MdMthLno(A, M)
        If J Mod 150 = 0 Then
            Debug.Print J, Sz(Ny), "Z_MdMthLno"
        End If
    Next

    For Each Lno In L
        Push O, Lno & " " & A.Lines(Lno, 1)
    Next
AyBrw O
End Sub

Private Sub Z_SrcyMthLno()
Dim O$()
    Dim Lno, L%(), M, A As CodeModule, Ny$(), J%, Srcy$()
    Set A = Md("Fct")
    Ny = MdMthNy(A)
    Srcy = MdLy(A)
    For Each M In Ny
        DoEvents
'        J = J + 1
'        If J > 500 Then
'            Debug.Print J, Sz(Ny), "Z_SrcyMthLno"
'        End If
        Push L, SrcyMthLno(Srcy, M)
    Next
   
    For Each Lno In L
        Push O, Lno & " " & A.Lines(Lno, 1)
    Next
AyBrw O
End Sub

Private Sub Z_MdMthLinCnt()
Dim O$()
    Dim J%, M, L%, E%, A As CodeModule, Ny$()
    Set A = Md("Fct")
    Ny = MdMthNy(A)
    For Each M In Ny
        DoEvents
        L = MdMthLno(A, M)
        E = MdMthLinCnt(A, L) + L - 1
        Push O, Format(L, "0000 ") & A.Lines(L, 1)
        Push O, Format(E, "0000 ") & A.Lines(E, 1)
    Next
AyBrw O
End Sub
Function MdMthLinCnt%(A As CodeModule, MthLno%)
Dim Kd$, Lin$, EndLin$, J%
Lin = A.Lines(MthLno, 1)
Kd = LinMthKd(Lin)
If Kd = "" Then Stop
EndLin = "End " & Kd
If HasSfx(Lin, EndLin) Then
    MdMthLinCnt = 1
    Exit Function
End If
For J = MthLno + 1 To A.CountOfLines
    If HasSfx(A.Lines(J, 1), EndLin) Then
        MdMthLinCnt = J - MthLno + 1
        Exit Function
    End If
Next
Stop
End Function
Function MdMthLno%(A As CodeModule, M)
Dim J%
For J = 1 To A.CountOfLines
    If LinMthNm(A.Lines(J, 1)) = M Then MdMthLno = J: Exit Function
Next
End Function
Sub MdMovMth(A As CodeModule, M, ToMd As CodeModule)
MdAppLines ToMd, MdMthLines(A, M)
MdRmvMth A, M
End Sub
Sub Z()
Z_MdMthLinCnt
Z_MovMth
Z_SrcyMthLno
End Sub