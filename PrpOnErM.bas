Option Compare Database
Option Explicit
Const Trc As Boolean = True

Sub EnsPrpOnEr()
MdEnsPrpOnEr CurMd
End Sub

Sub RmvPrpOnEr()
MdRmvPrpOnEr CurMd
End Sub

Private Sub Z_MdRmvPrpOnEr()
MdRmvPrpOnEr ZZMd
End Sub

Private Sub Z_MdEnsPrpOnEr()
MdEnsPrpOnEr ZZMd
End Sub

Sub MdRmvPrpOnEr(A As CodeModule)
If A.Parent.Type <> vbext_ct_ClassModule Then Exit Sub
Dim J%, L&()
L = MdPrpLnoAy(A)
If Not AyIsSrt(L) Then Stop
For J = UB(L) To 0 Step -1
    MdPrpRmvOnEr A, L(J)
Next
End Sub

Sub MdEnsPrpOnEr(A As CodeModule)
If A.Parent.Type <> vbext_ct_ClassModule Then Exit Sub
Dim J%, L&()
L = MdPrpLnoAy(A)
If Not AyIsSrt(L) Then Stop
For J = UB(L) To 0 Step -1
    MdPrpEnsOnEr A, L(J)
Next
End Sub

Function MdPrpLblXLin$(A As CodeModule, PrpLno)
Dim Nm$, Lin$
Lin = A.Lines(PrpLno, 1)
Nm = LinPrpNm(Lin)
If Nm = "" Then Stop
MdPrpLblXLin = FmtQQ("X: Debug.Print ""?.?.PrpEr...[""; Err.Description; ""]""", MdNm(A), Nm)
End Function

Private Sub MdPrpEnsOnEr(A As CodeModule, PrpLno&)
If HasSubStr(A.Lines(PrpLno, 1), "End Property") Then
    Exit Sub
End If
MdPrpEnsLblXLin A, PrpLno
MdPrpEnsExitPrpLin A, PrpLno
MdPrpEnsOnErLin A, PrpLno
End Sub

Private Sub MdPrpEnsLblXLin(A As CodeModule, PrpLno&)
Const CSub$ = "MdPrpEnsLblXLin"
Dim E$, L%, ActLblXLin$, EndPrpLno&
E = MdPrpLblXLin(A, PrpLno)
L = MdPrpLblXLno(A, PrpLno)
If L <> 0 Then
    ActLblXLin = A.Lines(L, 1)
End If
If E <> ActLblXLin Then
    If L = 0 Then
        EndPrpLno = MdPrpEndPrpLno(A, PrpLno)
        If EndPrpLno = 0 Then Stop
        A.InsertLines EndPrpLno, E
        If Trc Then FunMsgDmp CSub, "Inserted [at] with [line]", EndPrpLno, E
    Else
        A.ReplaceLine L, E
        If Trc Then FunMsgDmp CSub, "Replaced [at] with [line]", L, E
    End If
End If
End Sub

Private Sub MdPrpEnsExitPrpLin(A As CodeModule, PrpLno&)
Const CSub$ = "MdPrpEnsExitPrpLin"
Dim L&
L = MdPrpInsExitPrpLno(A, PrpLno)
If L = 0 Then Exit Sub
A.InsertLines L, "Exit Property"
If Trc Then FunMsgDmp CSub, "Exit Property is inserted [at]", L
End Sub

Private Sub MdPrpEnsOnErLin(A As CodeModule, PrpLno&)
Const CSub$ = "MdPrpEnsOnErLin"
Dim L&
L = MdPrpOnErLno(A, PrpLno)
If L <> 0 Then Exit Sub
A.InsertLines PrpLno + 1, "On Error Goto X"
If Trc Then FunMsgDmp CSub, "Exit Property is inserted [at]", L
End Sub

Private Sub MdPrpRmvOnEr(A As CodeModule, PrpLno&)
MdLno_Rmv A, MdPrpExitPrpLno(A, PrpLno)
MdLno_Rmv A, MdPrpOnErLno(A, PrpLno)
MdLno_Rmv A, MdPrpLblXLno(A, PrpLno)
End Sub

Function MdLno_Rmv(A As CodeModule, Lno)
If Lno = 0 Then Exit Function
MsgDmp "MdLno_Rmv: [Md]-[Lno]-[Lin] is removed", MdNm(A), Lno, A.Lines(Lno, 1)
A.DeleteLines Lno, 1
End Function

Function MdPrpInsExitPrpLno&(A As CodeModule, PrpLno)
If MdPrpExitPrpLno(A, PrpLno) <> 0 Then Exit Function
Dim L%
L = MdPrpLblXLno(A, PrpLno)
If L = 0 Then Stop
MdPrpInsExitPrpLno = L
End Function

Function MdPrpLblXLno&(A As CodeModule, PrpLno)
Dim J&, L$
For J = PrpLno + 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If HasPfx(L, "X: Debug.Print") Then MdPrpLblXLno = J: Exit Function
    If HasPfx(L, "End Property") Then Exit Function
Next
Stop
End Function

Function MdPrpLy(A As CodeModule) As String()
Dim O$(), Lno
For Lno = 0 To AyNz(MdPrpLnoAy(A))
    Push O, A.Lines(Lno, 1)
Next
MdPrpLy = O
End Function
Function MdPrpNy(A As CodeModule) As String()
Dim O$(), Lno
For Each Lno In AyNz(MdPrpLnoAy(A))
    PushNoDup O, LinPrpNm(A.Lines(Lno, 1))
Next
MdPrpNy = O
End Function
Function MdPrpLnoAy(A As CodeModule) As Long()
Dim O&(), Lno&
For Lno = 1 To A.CountOfLines
    If LinIsPrpLin(A.Lines(Lno, 1)) Then
        Push O, Lno
    End If
Next
MdPrpLnoAy = O
End Function
Private Function ZZMd() As CodeModule
Set ZZMd = CurVbe.VBProjects("StockShipRate").VBComponents("Schm").CodeModule
End Function

Sub Z()
Z_MdEnsPrpOnEr
Z_MdRmvPrpOnEr
End Sub