Option Compare Database
Option Explicit
Private Const ZZLines$ = _
         "E Mem | Mem Req AlwZLen" & _
vbCrLf & "E Txt | Txt Req" & _
vbCrLf & "E Crt | Dte Req Dft=Now" & _
vbCrLf & "E Dte | Dte" & _
vbCrLf & "F Amt * | *Amt" & _
vbCrLf & "F Crt * | CrtDte" & _
vbCrLf & "F Dte * | *Dte" & _
vbCrLf & "F Txt * | Fun * Txt" & _
vbCrLf & "F Mem * | Lines" & _
vbCrLf & "T Sess | * CrtDte" & _
vbCrLf & "T Msg  | * Fun *Txt | CrtDte" & _
vbCrLf & "T Lg   | * Sess Msg CrtDte" & _
vbCrLf & "T LgV  | * Lg Lines" & _
vbCrLf & "D . Fun | Function name that call the log" & _
vbCrLf & "D . Fun | Function name that call the log" & _
vbCrLf & "D . Msg | it will a new record when Lg-function is first time using the Fun+MsgTxt" & _
vbCrLf & "D . Msg | ..."

Sub DbCrtSchm(A As Database, SmLines$)
With Rslt(SmLines)
    AyBrwThw .Er
    AyDoPX .Td, "DbAddTd", A
    AyDoPX .PkSqy, "DbRun", A
    AyDoPX .SkSqy, "DbRun", A
    AyDoPX .FDes, "DbSetFDes", A
    AyDoPX .TDes, "DbSetTDes", A
End With
End Sub

Private Function BrkDLin(DLin) As DaoSmDRslt
Dim V$
Set BrkDLin = New DaoSmDRslt
Set BrkDLin.D = New DaoSmD
With BrkDLin.D
    AyAsg Lin3TAy(DLin), .T, .F, V, .Des
    If V <> "|" Then Stop
End With
End Function

Private Function BrkFLin(FLin) As DaoSmFRslt
Dim LikFldSsl$, A$, V$
Set BrkFLin = New DaoSmFRslt
Set BrkFLin.F = New DaoSmF
With BrkFLin.F
    AyAsg Lin3TAy(FLin), .E, .LikT, V, A
    .LikFny = SslSy(LikFldSsl)
End With
End Function

Private Function BrkELin(ELin As Ixl) As DaoSmERslt
Dim LikFldSsl$, A$(), V$, Ty$, Ay()
Set BrkELin = New DaoSmERslt
Set BrkELin.E = New DaoSmE
With BrkELin.E
    A = LinTermAy(ELin.Lin)
    .E = A(0)
    V = A(1)
    Ty = DaoShtTy_Ty(A(2))
    If Ty = 0 Then
        Push BrkELin.Er, ErMsgTyEr(ELin.Ix, A(2))
    End If
    A = AyMid(A, 2)
    If AyHas(A, "Req") Then
        .Req = True
        A = AyRmvEle(A, "Req", Cnt:=1)
    End If
    If AyHas(A, "AlwZ") Then
        .AlwZ = True
        A = AyRmvEle(A, "AlwZ", Cnt:=1)
    End If
    Ay = AyShiftItmEq(A, "Dft")
    .Dft = Ay(0)
    A = Ay(1)
    
    Ay = AyShiftItmEq(A, "VTxt")
    .VTxt = Ay(0)
    A = Ay(1)
    
    Ay = AyShiftItmEq(A, "VRul")
    .VRul = Ay(0)
    A = Ay(1)
    
    If .Ty = dbText Then
        Ay = AyShiftItmEq(A, "TxtSz")
        .TxtSz = Ay(0)
        A = Ay(1)
    End If
End With
End Function


Private Sub AAA()
Z_BrkTLin
End Sub

Private Sub Z_BrkTLin()
Dim Act As DaoSmTRslt
Dim Ept As DaoSmTRslt
Dim TLin$
TLin = "A"
Set Ept = New DaoSmTRslt
Push Ept.Er, "should have a |"
GoSub Tst
'
TLin = "A | B B"
Set Ept = New DaoSmTRslt
Push Ept.Er, "dup fields[B]"
GoSub Tst
'
TLin = "A | B B D C C"
Set Ept = New DaoSmTRslt
Push Ept.Er, "dup fields[B C]"
GoSub Tst
'
TLin = "A | * B D C"
Set Ept = New DaoSmTRslt
With Ept.T
    .T = "A"
    .Fny = SslSy("A B D C")
End With
GoSub Tst
'
TLin = "A | * B | D C"
Set Ept = New DaoSmTRslt
With Ept.T
    .T = "A"
    .Fny = SslSy("A B D C")
    .Sk = SslSy("B")
End With
GoSub Tst
'
TLin = "A |"
Set Ept = New DaoSmTRslt
With Ept
    Push .Er, "should have fields after |"
End With
GoSub Tst
Exit Sub
Tst:
    Act = BrkTLin(TLin)
    Ass IsTRsltEq(Act, Ept)
    Return
End Sub

Private Function IsTRsltEq(A As DaoSmTRslt, B As DaoSmTRslt) As Boolean
If Not AyIsEq(A.Er, B.Er) Then Exit Function
If Sz(A.Er) > 0 Then
    IsTRsltEq = True
    Exit Function
End If
IsTRsltEq = IsTItmEq(A.T, B.T)
End Function

Private Function IsTItmEq(A As DaoSmT, B As DaoSmT) As Boolean
If A.T <> B.T Then Exit Function
If Not AyIsEq(A.Fny, B.Fny) Then Exit Function
IsTItmEq = AyIsEq(A.Sk, B.Sk)
End Function

Private Function BrkTLin(TLin) As DaoSmTRslt
If Not HasSubStr(TLin, "|") Then
    Push BrkTLin.Er, "should have a |"
    Exit Function
End If
Dim A$, B$, C$, D$
BrkAsg TLin, "|", A, B
Set BrkTLin = New DaoSmTRslt
Set BrkTLin.T = New DaoSmT
With BrkTLin.T
    .T = A
    B = Replace(B, "*", A)
    BrkS1Asg B, "|", C, D
    If D = "" Then
        .Fny = SslSy(C)
    Else
        .Sk = SslSy(RmvPfx(C, A & " "))
        .Fny = SslSy(Replace(B, "|", " "))
    End If
    If Sz(.Fny) = 0 Then
        Push BrkTLin.Er, "should have fields after |"
        Exit Function
    End If
    Dim Dup$()
    Dup = AyWhDup(.Fny)
    If Sz(Dup) > 0 Then
        Push BrkTLin.Er, FmtQQ("dup fields[?]", JnSpc(Dup))
        Exit Function
    End If
End With
End Function

Private Function BrkD(DLy$()) As DaoSmDyRslt
Dim O As New DaoSmDyRslt, U%, J%, Er$(), D() As DaoSmD, R As DaoSmDRslt
U = UB(DLy)
ReDim D(U)
For J = 0 To U
    With BrkDLin(DLy(J))
        Set D(J) = .D
        PushAy O.Er, .Er
    End With
Next
O.D = D
Set BrkD = O
End Function

Private Function BrkT(TLy$()) As DaoSmTyRslt
If Sz(TLy) = 0 Then
    Push BrkT.Er, ErMsgNoTLin
    Exit Function
End If
Dim U%, J%, T() As DaoSmT
U = UB(TLy)
ReDim T(U)
Set BrkT = New DaoSmTyRslt
For J = 0 To U
    With BrkTLin(TLy(J))
        Set T(J) = .T
        Dim Pfx$
        Pfx = FmtQQ("T-Lin[?] ", TLy(J))
        PushAy BrkT.Er, AyAddPfx(.Er, Pfx)
    End With
Next
BrkT.T = T
End Function

Private Function BrkF(FLy$()) As DaoSmFyRslt
Dim U%, J%, Er$(), F() As DaoSmF
U = UB(FLy)
ReDim F(U)
Set BrkF = New DaoSmFyRslt
For J = 0 To U
    With BrkFLin(FLy(J))
        Set F(J) = .F
        PushAy BrkF.Er, .Er
    End With
Next
BrkF.F = F
End Function

Private Function BrkE(A() As Ixl) As DaoSmEyRslt
Dim U%, J%, Er$(), E() As DaoSmE
U = UB(A)
ReDim E(U)
Set BrkE = New DaoSmEyRslt
For J = 0 To U
    With BrkELin(A(J))
        Set E(J) = .E
        PushAy BrkE.Er, .Er
    End With
Next
BrkE.E = E
End Function
Private Function Brk(SmLines$) As DaoSmBrk
Dim Cln$(), E As DaoSmEyRslt, D As DaoSmDyRslt, F As DaoSmFyRslt, T As DaoSmTyRslt
Dim Er$(), TLy$(), EIxly() As Ixl, ELnoAy%(), DLy$()
Dim ClnIxly() As Ixl
ClnIxly = LyClnIxly(SplitCrLf(SmLines))
Cln = IxlyLy(ClnIxly)
TLy = AyWhRmvT1(Cln, "T")
DLy = AyWhRmvT1(Cln, "D")
EIxly = IxlyWhRmvT1(ClnIxly, "E")
Set D = BrkD(DLy)
Set E = BrkE(EIxly)
Set F = BrkF(AyWhRmvT1(Cln, "F"))
Set T = BrkT(TLy)
Er = ClnChk(Cln, "D E F T")
Brk.Er = CvSy(AyAddAp(Er, D.Er, E.Er, F.Er, , T.Er))
If Sz(Brk.Er) > 0 Then Exit Function
With Brk.Dta
    .E = E.E
    .F = F.F
    .T = T.T
    .D = D.D
    .Eny = Eny(IxlyLy(EIxly))
    .Tny = Tny(TLy)
End With
End Function

Private Function Tny(TLy$()) As String()
Tny = AyT1Ay(TLy)
End Function

Private Function Er_DupT(Tny$()) As String()
Er_DupT = AyDupChk(Tny, "These T[?] in T-Lin is duplicated")
End Function

Private Function Er_DupE(Eny$()) As String()
Er_DupE = AyDupChk(Eny, "These E[?] in E-Lin is duplicated")
End Function

Private Function Eny(ELy$()) As String()
Eny = AyT1Ay(ELy)
End Function

Private Function Er(A As DaoSmBrk) As String()
Dim D As DaoSmDta
D = A.Dta
Er = AyAddAp _
    (A.Er _
   , Er_DupT(D.Tny) _
   , Er_DupE(D.Eny) _
   , Er_TzDLy_NotIn_Tny(D.D, D.Tny) _
   , Er_FzDLy_NotIn_TblFny(D.D, D.Tny, D.T) _
   , Er_EzFLy_NotIn_Eny(D.F, D.Eny) _
    )
End Function

Private Function PkSqy(A As DaoSmDta) As String()
Dim J%, O$(), T() As DaoSmT
T = A.T
With A
    For J = 0 To UBound(T)
        PushNonEmp O, PkSql(T(J))
    Next
End With
PkSqy = O
End Function

Private Function PkSql$(A As DaoSmT)
With A
    If AyHas(.Fny, .T) Then PkSql = SqlzCrtPk(.T)
End With
End Function

Private Function ItmT(T, A() As DaoSmT) As DaoSmT
Dim J%
For J = 0 To UBound(A)
    With A(J)
        If .T = T Then Set ItmT = A(J): Exit Function
    End With
Next
End Function

Private Function SkSql$(T, TBrk() As DaoSmT)
Dim M As DaoSmT
Set M = ItmT(T, TBrk)
If Sz(M.Sk) = 0 Then Exit Function
SkSql = SqlzCrtSk(T, M.Sk)
End Function

Private Function SkSqy(A As DaoSmDta) As String()
Dim T, O$()
For Each T In A.Tny
    PushNonEmp O, SkSql(T, A.T)
Next
SkSqy = O
End Function

Private Function Td(A As DaoSmDta) As DAO.TableDef()
Dim O() As DAO.TableDef, I, T$, FdAy1() As DAO.Field2
For Each I In A.Tny
    T = I
    FdAy1 = FdAy(T, A)
    PushObj O, NewTd(T, FdAy1)
Next
Td = O
End Function

Private Function Fny(T, TBrk() As DaoSmT) As String()
Dim J%
With ItmT(T, TBrk)
    Fny = .Fny
    If .T <> T Then Stop
End With
End Function

Private Function ItmE(T$, F$, FBrk() As DaoSmF, EBrk() As DaoSmE) As DaoSmE
Dim J%, O As DaoSmF, M As DaoSmF
For J = 0 To UBound(FBrk)
    M = FBrk(J)
    If T Like M.LikT Then
        If LikAyHas(M.LikFny, F) Then
            ItmE = ItmE__1(M.E, EBrk)
            If ItmE.E <> M.E Then Stop
            Exit Function
        End If
    End If
Next
End Function

Private Function ItmE__1(E$, EBrk() As DaoSmE) As DaoSmE
Dim J%
For J = 0 To UBound(EBrk)
    If EBrk(J).E = E Then
        ItmE__1 = EBrk(J)
        Exit Function
    End If
Next
End Function

Private Function Fd(T$, F$, Tny$(), FBrk() As DaoSmF, EBrk() As DaoSmE) As DAO.Field2
Dim E As DaoSmE
Select Case True
Case T = F: Set Fd = NewFd_zId(F)
Case AyHas(Tny, T): Set Fd = NewFd_zFk(F)
Case Else
E = ItmE(T, F, FBrk, EBrk)
With E
    Set Fd = NewFd(F, .Ty, .TxtSz, .Expr, .Dft, .Req, .VRul, .VTxt)
End With
End Select
End Function

Private Function FdAy(T$, A As DaoSmDta) As DAO.Field2()
Dim I, F$, O() As DAO.Field2
For Each I In Fny(T, A.T)
    F = I
    PushObj O, Fd(T, F, A.Tny, A.F, A.E)
Next
FdAy = O
End Function
Private Function Er_EzFLy_NotIn_Eny(F() As DaoSmF, Eny$()) As String()
End Function

Private Function Er_TzDLy_NotIn_Tny(D() As DaoSmD, Tny$()) As String()
Dim O$(), M As DaoSmD, Tssl$, J%
Tssl = JnSpc(Tny)
For J = 0 To UB(D)
    With D(J)
        If Not AyHas(Tny, .T) Then Push O, ErMsg_TzDLy_NotIn_Tny(.Lno, .T, Tssl)
    End With
Next
Er_TzDLy_NotIn_Tny = O
End Function

Private Function Er_FzDLy_NotIn_TblFny(D() As DaoSmD, Tny$(), TBrk() As DaoSmT) As String()
Dim J%, O$(), M As DaoSmD, F, Fny1$()
For J = 0 To UB(D)
    With D(J)
        If Not AyHas(Tny, .T) Then GoTo Nxt
        Fny1 = Fny(.T, TBrk)
        For Each F In AyNz(Fny1)
            If Not AyHas(Fny1, .F) Then
                Push O, ErMsg_FzDLy_NotIn_TblFny(.Lno, .T, CStr(F), JnSpc(Fny1))
            End If
        Next
    End With
Nxt:
Next
Er_FzDLy_NotIn_TblFny = O
End Function

Private Function ErMsg_TzDLy_NotIn_Tny$(Lno%, T$, Tssl$)
ErMsg_TzDLy_NotIn_Tny = ErMsg(Lno, FmtQQ("T[?] is invalid.  Valid T[?]", T, Tssl))
End Function

Private Function ErMsg_FzDLy_NotIn_TblFny$(Lno%, T$, F$, Fssl$)
ErMsg_FzDLy_NotIn_TblFny = ErMsg(Lno, FmtQQ("F[?] is invalid in T[?].  Valid F[?]", F, T, Fssl))
End Function

Private Function TDesAy(A As DaoSmDta) As TDes()
Stop '
End Function

Private Function FDesAy(A As DaoSmDta) As FDes()
Stop '
End Function
Private Function ErMsgTblFldEr$(Lno%, T$, F$)
ErMsgTblFldEr = ErMsg(Lno, FmtQQ("T[?] has invalid F[?], which cannot be found in any F-Lines"))
End Function
Private Function ErMsgFldEleEr$(Lno%, E$, Essl$)
ErMsgFldEleEr = ErMsg(Lno, FmtQQ("E[?] is invalid.  Valid E is [?]", E, Essl))
End Function
Private Function ErMsgDupF$(Lno%, T$, Fny$())
ErMsgDupF = ErMsg(Lno, FmtQQ("F[?] is dup in T[?]", JnSpc(Fny), T))
End Function
Private Function ErMsgExcessTxtSz$(Lno%, Ty$)
ErMsgExcessTxtSz = ErMsg(Lno, FmtQQ("Ty[?] is not Txt, it should not have TxtSz", Ty))
End Function
Private Function ErMsgDupT$(LnoAy%(), T$)
ErMsgDupT = ErMsg1(LnoAy, FmtQQ("This T[?] is dup", T))
End Function
Private Function ErMsgTyEr$(Lno%, Ty$)
ErMsgTyEr = ErMsg(Lno, FmtQQ("Invalid DaoShtTy[?].  Valid ShtTy[?]", Ty, DaoShtTySsl))
End Function
Private Function ErMsgDupE$(LnoAy%(), E$)
ErMsgDupE = ErMsg1(LnoAy, FmtQQ("This E[?] is dup", E))
End Function
Private Function ErMsgNoTLin$()
ErMsgNoTLin = "No T-Line"
End Function
Private Function ErMsgNoELin$()
ErMsgNoELin = "No E-Line"
End Function
Private Function ErMsgNoFLin$()
ErMsgNoFLin = "No F-Line"
End Function
Private Function ErMsg1(LnoAy%(), M$)
ErMsg1 = "--" & Join(AyAddPfx(LnoAy, "Lno"), ".") & "  " & M
End Function

Private Function ErMsg$(Lno%, M$)
ErMsg = "--Lno" & Lno & ".  " & M
End Function

Private Function Rslt(SmLines$) As DaoSmRslt
Dim B As DaoSmBrk, Er1$(), D As DaoSmDta

Set B = Brk(SmLines)
Er1 = Er(B):  If Sz(Er1) > 0 Then Rslt.Er = Er1: Exit Function
Set D = B.Dta
Set Rslt = New DaoSmRslt
With Rslt
    .Td = Td(D)
    .PkSqy = PkSqy(D)
    .SkSqy = SkSqy(D)
    .FDes = FDesAy(D)
    .TDes = TDesAy(D)
End With
End Function

Private Sub Z_DbCrtSchm()
DbCrtSchm W, ZZLines
End Sub

Sub AA()
Z
End Sub

Sub Z()
Z_DbCrtSchm
End Sub