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
Private Type T: Lno As Integer: T As String: Fny() As String: Sk() As String:     End Type
Private Type F: Lno As Integer: E As String: LikT As String:  LikFny() As String: End Type
Private Type D: Lno As Integer: T As String: F As String:     Des As String:      End Type
Private Type E
    Lno As Integer
    E As String
    Ty As DAO.DataTypeEnum
    Req As Boolean
    AlwZ As Boolean
    TxtSz As Byte
    Expr As String
    VRul As String
    Dft As String
    VTxt As String
End Type
Private Type ERslt: E As E: Er() As String: End Type
Private Type FRslt: F As F: Er() As String: End Type
Private Type TRslt: T As T: Er() As String: End Type
Private Type DRslt: D As D: Er() As String: End Type
Private Type EyRslt: E() As E: Er() As String: End Type
Private Type FyRslt: F() As F: Er() As String: End Type
Private Type TyRslt: T() As T: Er() As String: End Type
Private Type DyRslt: D() As D: Er() As String: End Type
Private Type Dta
    E() As E
    F() As F
    T() As T
    D() As D
    Tny() As String
    Eny() As String
End Type
Private Type Brk: Dta As Dta: Er() As String: End Type
Private Type Rslt
    Er() As String
    Td() As DAO.TableDef
    SkSqy() As String
    PkSqy() As String
    TDes() As TDes
    FDes() As FDes
End Type

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

Private Function BrkDLin(D As Ixl) As DRslt
Dim V$
With BrkDLin.D
    AyAsg Lin3TAy(D.Lin), .T, .F, V, .Des
    If V <> "|" Then Stop
End With
End Function

Private Function BrkFLin(F As Ixl) As FRslt
Dim LikFldSsl$, A$, V$
With BrkFLin.F
    AyAsg Lin3TAy(F.Lin), .E, .LikT, V, A
    .LikFny = SslSy(LikFldSsl)
End With
End Function

Private Function BrkELin(ELin As Ixl) As ERslt
Dim LikFldSsl$, A$(), V$, Ty$, Ay()
With BrkELin.E
    A = LinTermAy(ELin.Lin)
    .E = A(0)
    V = A(1)
    Ty = DaoShtTy_Ty(A(2))
    If Ty = 0 Then
        Push BrkELin.Er, ErMsg_TyEr(ELin.Ix, A(2))
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
Dim Act As TRslt
Dim Ept As TRslt
Dim Emp As TRslt
Dim TLin As New Ixl
TLin.Ix = 999
TLin.Lin = "A"
Ept = Emp
Push Ept.Er, "should have a |"
GoSub Tst
'
TLin.Lin = "A | B B"
Ept = Emp
Push Ept.Er, "dup fields[B]"
GoSub Tst
'
TLin = "A | B B D C C"
Ept = Emp
Push Ept.Er, "dup fields[B C]"
GoSub Tst
'
TLin = "A | * B D C"
Ept = Emp
With Ept.T
    .T = "A"
    .Fny = SslSy("A B D C")
End With
GoSub Tst
'
TLin = "A | * B | D C"
Ept = Emp
With Ept.T
    .T = "A"
    .Fny = SslSy("A B D C")
    .Sk = SslSy("B")
End With
GoSub Tst
'
TLin = "A |"
Ept = Emp
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

Private Function IsTRsltEq(A As TRslt, B As TRslt) As Boolean
If Not AyIsEq(A.Er, B.Er) Then Exit Function
If Sz(A.Er) > 0 Then
    IsTRsltEq = True
    Exit Function
End If
IsTRsltEq = IsTItmEq(A.T, B.T)
End Function

Private Function IsTItmEq(A As T, B As T) As Boolean
If A.T <> B.T Then Exit Function
If Not AyIsEq(A.Fny, B.Fny) Then Exit Function
IsTItmEq = AyIsEq(A.Sk, B.Sk)
End Function

Private Function BrkTLin(T As Ixl) As TRslt
If Not HasSubStr(T.Lin, "|") Then
    Push BrkTLin.Er, "should have a |"
    Exit Function
End If
Dim A$, B$, C$, D$
BrkAsg T.Lin, "|", A, B
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
        Stop '
'       Push BrkTLin.Er, ErMsg_DupF(T.Ix + 1)
        Exit Function
    End If
End With
End Function

Private Function BrkD(D() As Ixl) As DyRslt
Dim U%, J%, Er$()
U = UB(D)
ReDim BrkD.D(U)
For J = 0 To U
    With BrkDLin(D(J))
        BrkD.D(J) = .D
        PushAy BrkD.Er, .Er
    End With
Next
End Function

Private Function BrkT(A() As Ixl) As TyRslt
Dim U%: U = UB(A)
If Sz(A) = 0 Then
    Push BrkT.Er, ErMsg_NoTLin
    Exit Function
End If
ReDim BrkT.T(U)
Dim J%
For J = 0 To U
    With BrkTLin(A(J))
        BrkT.T(J) = .T
        PushAy BrkT.Er, .Er
    End With
Next
End Function

Private Function BrkF(A() As Ixl) As FyRslt
Dim U%
U = UB(A)
If U = -1 Then
    Push BrkF.Er, ErMsg_NoFLin
    Exit Function
End If
ReDim BrkF.F(U)
Dim J%
For J = 0 To U
    With BrkFLin(A(J))
        BrkF.F(J) = .F
        PushAy BrkF.Er, .Er
    End With
Next
End Function

Private Function BrkE(A() As Ixl) As EyRslt
Dim U%
U = UB(A)
If U = -1 Then
    Push BrkE.Er, ErMsg_NoELin
    Exit Function
End If
ReDim BrkE.E(U)
Dim J%
For J = 0 To U
    With BrkELin(A(J))
        BrkE.E(J) = .E
        PushAy BrkE.Er, .Er
    End With
Next
End Function
Private Function Brk(SmLines$) As Brk
Dim Cln$(), Re As EyRslt, RD As DyRslt, RF As FyRslt, RT As TyRslt
Dim Er$()
Dim ClnIxly() As Ixl
Dim E() As Ixl
Dim F() As Ixl
Dim D() As Ixl
Dim T() As Ixl

ClnIxly = LyClnIxly(SplitCrLf(SmLines))
T = IxlyWhRmvT1(ClnIxly, "T")
D = IxlyWhRmvT1(ClnIxly, "D")
E = IxlyWhRmvT1(ClnIxly, "E")
F = IxlyWhRmvT1(ClnIxly, "F")
Re = BrkE(E)
RF = BrkF(F)
RD = BrkD(D)
RT = BrkT(T)
Er = IxlyT1Chk(ClnIxly, "D E F T")
Brk.Er = CvSy(AyAddAp(Er, RD.Er, Re.Er, RF.Er, , RT.Er))
If Sz(Brk.Er) > 0 Then Exit Function

With Brk.Dta
    .E = Re.E
    .F = RF.F
    .T = RT.T
    .D = RD.D
    .Eny = AyT1Ay(IxlyLy(E))
    .Tny = AyT1Ay(IxlyLy(T))
End With
End Function
Private Function TLnoAy(T$, TBrk() As T) As Integer()
Dim J%
For J = 0 To UBound(TBrk)
    Push TLnoAy, TBrk(J).Lno
Next
End Function
Private Function ELnoAy(E$, EBrk() As E) As Integer()
Dim J%
For J = 0 To UBound(EBrk)
    Push ELnoAy, EBrk(J).Lno
Next
End Function
Private Function Er_DupT(A() As T, Tny$()) As String()
Dim Dup$(), IT, T$, LnoAy%()
Dup = AyWhDup(Tny)
If Sz(Dup) = 0 Then Exit Function
For Each IT In Dup
    T = IT
    LnoAy = TLnoAy(T, A)
    Push Er_DupT, ErMsg_DupT(LnoAy, T)
Next
End Function

Private Function Er_DupE(A() As E, Eny$()) As String()
Dim Dup$(), IE, E$, LnoAy%()
Dup = AyWhDup(Eny)
If Sz(Dup) = 0 Then Exit Function
For Each IE In Dup
    E = IE
    LnoAy = ELnoAy(E, A)
    Push Er_DupE, ErMsg_DupE(LnoAy, E)
Next
End Function

Private Function Er(A As Brk) As String()
With A.Dta
    Er = AyAddAp _
        (A.Er _
       , Er_DupT(.T, .Tny) _
       , Er_DupE(.E, .Eny) _
       , Er_TzDLy_NotIn_Tny(.D, .Tny) _
       , Er_FzDLy_NotIn_TblFny(.D, .Tny, .T) _
       , Er_EzFLy_NotIn_Eny(.F, .Eny) _
        )
End With
End Function

Private Function PkSqy(A As Dta) As String()
Dim J%, O$(), T() As T
T = A.T
With A
    For J = 0 To UBound(T)
        PushNonEmp O, PkSql(T(J))
    Next
End With
PkSqy = O
End Function

Private Function PkSql$(A As T)
With A
    If AyHas(.Fny, .T) Then PkSql = SqlzCrtPk(.T)
End With
End Function

Private Function ItmT(T, A() As T) As T
Dim J%
For J = 0 To UBound(A)
    With A(J)
        If .T = T Then ItmT = A(J): Exit Function
    End With
Next
End Function

Private Function SkSql$(T, TBrk() As T)
Dim M As T
M = ItmT(T, TBrk)
If Sz(M.Sk) = 0 Then Exit Function
SkSql = SqlzCrtSk(T, M.Sk)
End Function

Private Function SkSqy(A As Dta) As String()
Dim T, O$()
For Each T In A.Tny
    PushNonEmp O, SkSql(T, A.T)
Next
SkSqy = O
End Function

Private Function Td(A As Dta) As DAO.TableDef()
Dim O() As DAO.TableDef, I, T$, FdAy1() As DAO.Field2
For Each I In A.Tny
    T = I
    FdAy1 = FdAy(T, A)
    PushObj O, NewTd(T, FdAy1)
Next
Td = O
End Function

Private Function Fny(T, TBrk() As T) As String()
Dim J%
With ItmT(T, TBrk)
    Fny = .Fny
    If .T <> T Then Stop
End With
End Function

Private Function ItmE(T$, F$, FBrk() As F, EBrk() As E) As E
Dim J%, O As F, M As F
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

Private Function ItmE__1(E$, EBrk() As E) As E
Dim J%
For J = 0 To UBound(EBrk)
    If EBrk(J).E = E Then
        ItmE__1 = EBrk(J)
        Exit Function
    End If
Next
End Function

Private Function Fd(T$, F$, Tny$(), FBrk() As F, EBrk() As E) As DAO.Field2
Dim E As E
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

Private Function FdAy(T$, A As Dta) As DAO.Field2()
Dim I, F$, O() As DAO.Field2
For Each I In Fny(T, A.T)
    F = I
    PushObj O, Fd(T, F, A.Tny, A.F, A.E)
Next
FdAy = O
End Function
Private Function FUB%(A() As F)
FUB = FSz(A) - 1
End Function
Private Function FSz%(A() As F)
On Error Resume Next
FSz = UBound(A)
End Function

Private Function Er_EzFLy_NotIn_Eny(F() As F, Eny$()) As String()
Dim J%, O$()
For J = 0 To FUB(F)
    With F(J)
        Stop '
        'If Not AyHas(Eny, .E) Then Push O, ErMsg_EzFLy_NotIn_Eny(.Lno, .E)
    End With
Next
Er_EzFLy_NotIn_Eny = O
End Function

Private Function Er_TzDLy_NotIn_Tny(D() As D, Tny$()) As String()
Dim Tssl$, J%
Tssl = JnSpc(Tny)
For J = 0 To DUB(D)
    With D(J)
        If Not AyHas(Tny, .T) Then
            Push Er_TzDLy_NotIn_Tny, ErMsg_TzDLy_NotIn_Tny(.Lno, .T, Tssl)
        End If
    End With
Next
End Function

Private Function Er_FzDLy_NotIn_TblFny(D() As D, Tny$(), T() As T) As String()
Dim J%, Fny1$()
For J = 0 To DUB(D)
    With D(J)
        If Not AyHas(Tny, .T) Then GoTo Nxt
        Fny1 = Fny(.T, T)
        If Not AyHas(Fny1, .F) Then
            Push Er_FzDLy_NotIn_TblFny, ErMsg_FzDLy_NotIn_TblFny(.Lno, .T, .F, JnSpc(Fny1))
        End If
    End With
Nxt:
Next
End Function

Private Function ErMsg_TzDLy_NotIn_Tny$(Lno%, T$, Tssl$)
ErMsg_TzDLy_NotIn_Tny = ErMsg(Lno, FmtQQ("T[?] is invalid.  Valid T[?]", T, Tssl))
End Function

Private Function ErMsg_FzDLy_NotIn_TblFny$(Lno%, T$, F$, Fssl$)
ErMsg_FzDLy_NotIn_TblFny = ErMsg(Lno, FmtQQ("F[?] is invalid in T[?].  Valid F[?]", F, T, Fssl))
End Function

Private Function ErMsg_EzFLy_NotIn_Eny$(Lno%, E$, Essl$)
ErMsg_EzFLy_NotIn_Eny = ErMsg(Lno, FmtQQ("E[?] of is not in E-Lin[?]", E, Essl))
End Function

Private Function TDesAy(A As Dta) As TDes()
Stop '
End Function

Private Function FDesAy(A As Dta) As FDes()
Stop '
End Function

Private Function ErMsg_TblFldEr$(Lno%, T$, F$)
ErMsg_TblFldEr = ErMsg(Lno, FmtQQ("T[?] has invalid F[?], which cannot be found in any F-Lines"))
End Function

Private Function ErMsg_FldEleEr$(Lno%, E$, Essl$)
ErMsg_FldEleEr = ErMsg(Lno, FmtQQ("E[?] is invalid.  Valid E is [?]", E, Essl))
End Function

Private Function ErMsg_DupF$(Lno%, T$, Fny$())
ErMsg_DupF = ErMsg(Lno, FmtQQ("F[?] is dup in T[?]", JnSpc(Fny), T))
End Function

Private Function ErMsg_ExcessTxtSz$(Lno%, Ty$)
ErMsg_ExcessTxtSz = ErMsg(Lno, FmtQQ("Ty[?] is not Txt, it should not have TxtSz", Ty))
End Function

Private Function ErMsg_DupT$(LnoAy%(), T$)
ErMsg_DupT = ErMsg1(LnoAy, FmtQQ("This T[?] is dup", T))
End Function

Private Function ErMsg_TyEr$(Lno%, Ty$)
ErMsg_TyEr = ErMsg(Lno, FmtQQ("Invalid DaoShtTy[?].  Valid ShtTy[?]", Ty, DaoShtTySsl))
End Function

Private Function ErMsg_DupE$(LnoAy%(), E$)
ErMsg_DupE = ErMsg1(LnoAy, FmtQQ("This E[?] is dup", E))
End Function

Private Function ErMsg_NoTLin$()
ErMsg_NoTLin = "No T-Line"
End Function

Private Function ErMsg_NoELin$()
ErMsg_NoELin = "No E-Line"
End Function

Private Function ErMsg_NoFLin$()
ErMsg_NoFLin = "No F-Line"
End Function

Private Function ErMsg1(LnoAy%(), M$)
ErMsg1 = "--" & Join(AyAddPfx(LnoAy, "Lno"), ".") & "  " & M
End Function

Private Function ErMsg$(Lno%, M$)
ErMsg = "--Lno" & Lno & ".  " & M
End Function

Private Function Rslt(SmLines$) As Rslt
Dim B As Brk, Er1$(), D As Dta
B = Brk(SmLines)
Er1 = Er(B):  If Sz(Er1) > 0 Then Rslt.Er = Er1: Exit Function
D = B.Dta
With Rslt
    .Td = Td(D)
    .PkSqy = PkSqy(D)
    .SkSqy = SkSqy(D)
    .FDes = FDesAy(D)
    .TDes = TDesAy(D)
End With
End Function

Private Sub Z_DbCrtSchm()
WKill
DbCrtSchm W, ZZLines
End Sub

Sub AA()
Z
End Sub

Sub Z()
Z_DbCrtSchm
End Sub

Function TUB%(A() As T)
TUB = TSz(A) - 1
End Function

Function TSz%(A() As T)
On Error Resume Next
TSz = UBound(A) + 1
End Function
Function DUB%(A() As D)
DUB = DSz(A) - 1
End Function
Function DSz%(A() As D)
On Error Resume Next
DSz = UBound(A) + 1
End Function