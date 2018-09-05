Option Compare Database
Option Explicit
Const Trc As Boolean = True
Sub Z_MdRmvPrpOnEr()
MdRmvPrpOnEr ZBMd
End Sub
Sub Z_MdEnsPrpOnEr()
MdEnsPrpOnEr ZBMd
End Sub
Sub MdRmvPrpOnEr(A As CodeModule)
If A.Parent.Type <> vbext_ct_ClassModule Then Exit Sub
Dim J%, L&(), Nm$
L = MdPrpLnoAy(A)
Nm = MdNm(A)
If Not AyIsSrt(L) Then Stop
For J = UB(L) To 0 Step -1
    MdPrpLno_RmvPrpOnEr A, L(J)
Next
End Sub

Sub MdEnsPrpOnEr(A As CodeModule)
If A.Parent.Type <> vbext_ct_ClassModule Then Exit Sub
Dim J%, L&(), Nm$
L = MdPrpLnoAy(A)
Nm = MdNm(A)
If Not AyIsSrt(L) Then Stop
For J = UB(L) To 0 Step -1
    MdPrpLno_EnsPrpOnEr A, L(J)
Next
End Sub

Function MdPrpLno_LblXLin$(A As CodeModule, PrpLno)
Dim Nm$
Nm = MdPrpLno_PrpNm(A, PrpLno)
If Nm = "" Then Stop
MdPrpLno_LblXLin = FmtQQ("X: Debug.Print ""?.?.PrpEr...", MdNm(A), Nm)
End Function

Private Sub MdPrpLno_EnsPrpOnEr(A As CodeModule, PrpLno&)
If HasSubStr(A.Lines(PrpLno, 1), "End Property") Then
    Exit Sub
End If
MdPrpLno_EnsLblXLin A, PrpLno
MdPrpLno_EnsExitPrpLin A, PrpLno
MdPrpLno_EnsOnErLin A, PrpLno
End Sub

Private Sub MdPrpLno_EnsLblXLin(A As CodeModule, PrpLno&)
Const CSub$ = "MdPrpLno_EnsLblXLin"
Dim E$, L%, ActLblXLin$, EndPrpLno&
E = MdPrpLno_LblXLin(A, PrpLno)
L = MdPrpLno_LblXLno(A, PrpLno)
If L <> 0 Then ActLblXLin = A.Lines(L, 1)
If E <> ActLblXLin Then
    If L = 0 Then
        EndPrpLno = MdPrpLno_EndPrpLno(A, PrpLno)
        If EndPrpLno = 0 Then Stop
        A.InsertLines EndPrpLno, E
        If Trc Then FunMsgDmp CSub, "Inserted [at] with [line]", EndPrpLno, E
    Else
        A.ReplaceLine L, E
        If Trc Then FunMsgDmp CSub, "Replaced [at] with [line]", L, E
    End If
End If
End Sub

Private Sub MdPrpLno_EnsExitPrpLin(A As CodeModule, PrpLno&)
Const CSub$ = "MdPrpLno_EnsExitPrpLin"
Dim L&
L = MdPrpLno_InsExitPrpLno(A, PrpLno)
If L = 0 Then Exit Sub
A.InsertLines L, "Exit Property"
If Trc Then FunMsgDmp CSub, "Exit Property is inserted [at]", L
End Sub

Private Sub MdPrpLno_EnsOnErLin(A As CodeModule, PrpLno&)
Const CSub$ = "MdPrpLno_EnsOnErLin"
Dim L&
L = MdPrpLno_OnErLno(A, PrpLno)
If L <> 0 Then Exit Sub
A.InsertLines PrpLno + 1, "On Error Goto X"
If Trc Then FunMsgDmp CSub, "Exit Property is inserted [at]", L
End Sub

Private Sub MdPrpLno_RmvPrpOnEr(A As CodeModule, PrpLno&)
MdLno_Rmv A, MdPrpLno_ExitPrpLno(A, PrpLno)
MdLno_Rmv A, MdPrpLno_OnErLno(A, PrpLno)
MdLno_Rmv A, MdPrpLno_LblXLno(A, PrpLno)
End Sub

Function MdLno_Rmv(A As CodeModule, Lno)
If Lno = 0 Then Exit Function
MsgDmp "MdLno_Rmv: [Md]-[Lno]-[Lin] is removed", MdNm(A), Lno, A.Lines(Lno, 1)
A.DeleteLines Lno, 1
End Function

Function MdPrpLno_InsExitPrpLno&(A As CodeModule, PrpLno)
If MdPrpLno_ExitPrpLno(A, PrpLno) <> 0 Then Exit Function
Dim L%
L = MdPrpLno_LblXLno(A, PrpLno)
If L = 0 Then Stop
MdPrpLno_InsExitPrpLno = L
End Function

Function MdPrpLno_LblXLno&(A As CodeModule, PrpLno)
Dim J&, L$
For J = PrpLno + 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If HasPfx(L, "X: Debug.Print") Then MdPrpLno_LblXLno = J: Exit Function
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
Private Function ZBMd() As CodeModule
Set ZBMd = CurVbe.VBProjects("StockShipRate").VBComponents("Schm").CodeModule
End Function