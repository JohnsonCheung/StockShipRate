Option Compare Database
Option Explicit

Function MdMthNy_zPfxEpt(A As CodeModule, MthPfx$, Optional EptMthPfx$) As String()
MdMthNy_zPfxEpt = AyWhPfxEpt(MdMthNy(A), MthPfx, EptMthPfx)
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
Dim B$
B = LinRmvMdy(A)
LinIsPrp = HasPfx(B, "Property")
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

Sub MovMth(FmMd$, MthPfx$, ToMd$, Optional EptMthPfx$)
Dim Fm As CodeModule
Dim Ny$()
Set Fm = Md(FmMd)
Ny = MdMthNy_zPfxEpt(Fm, MthPfx, EptMthPfx)
AyDoAXB Ny, "MdMthMov", Fm, Md(ToMd)
End Sub

Sub MdRmvMth(A As CodeModule, M)
Dim L%(), J%, Cnt%
L = MdMthLnoAy(A, M)
For J = UB(L) To 0 Step -1
    Cnt = MdMthLinCnt(A, L(J))
    A.DeleteLines L(J), Cnt
Next
End Sub
Function MthKdAy() As String()
Static O$()
If Sz(O) = 0 Then
    Push O, "Function"
    Push O, "Sub"
    Push O, "Property"
End If
MthKdAy = O
End Function
Function LinMthKd$(A)
Dim O
Const C1$ = "Function"
Const C2$ = "Sub"
Const C3$ = "Property"
O = LinT1(LinRmvMdy(A))
If AyHas(MthKdAy, O) Then LinMthKd = O
End Function

Sub AAAAA()
Z_MdMthLinCnt
End Sub
Sub A6()
Z_MdMthLno
End Sub
Sub A7()
Z_SrcyMthLno
End Sub
Private Sub Z_MdMthLno()
Dim O$()
    Dim Lno, L%(), M, A As CodeModule, Ny$(), J%
    Set A = Md("Fct")
    Ny = MdMthNy(A)
    For Each M In Ny
        J = J + 1
        Push L, MdMthLno(A, M)
        If J Mod 50 = 0 Then
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
        J = J + 1
        If J Mod 10 = 0 Then
            Debug.Print J, Sz(Ny), "Z_SrcyMthLno"
            Stop
        End If
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
        L = MdMthLno(A, M)
        E = MdMthLinCnt(A, L)
        Push O, Format("### ", L) & A.Lines(L, 1)
        Push O, Format("### ", E) & A.Lines(E, 1)
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