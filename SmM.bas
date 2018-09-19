Option Compare Database
Option Explicit
Private Const ZZSmLines$ = _
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
Private Type E
    E As String
    Req As Boolean
    Ty As DAO.DataTypeEnum
    AlwZ As Boolean
    TxtSz As Byte
    VRul As String
    VTxt As String
    Dft As String
End Type
Private Type F
    E As String
    LikT As String
    LikFny() As String
End Type
Private Type D
    T As String
    F As String
    Des As String
End Type
Private Type T
    Sk() As String
    T As String
    Fny() As String
End Type
Private Type Rslt
    Er() As String
    SkSqy() As String
    PkSqy() As String
    Td() As DAO.TableDef
    FDes() As FDes
    TDes() As TDes
End Type
Private Type Dta
    E() As E
    F() As F
    T() As T
    D() As D
    Eny() As String
    Tny() As String
End Type
Private Type Brk
    Er() As String
    Dta As Dta
End Type
Private Type ERslt
    E As E
    Er() As String
End Type
Private Type FRslt
    F As F
    Er() As String
End Type
Private Type TRslt
    T As T
    Er() As String
End Type
Private Type DRslt
    D As D
    Er() As String
End Type
Private Type EyRslt
    E() As E
    Er() As String
End Type
Private Type FyRslt
    F() As F
    Er() As String
End Type
Private Type TyRslt
    T() As T
    Er() As String
End Type
Private Type DyRslt
    D() As D
    Er() As String
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

Private Function BrkDItm(DLin) As DRslt
Dim V$
With BrkDItm.D
    AyAsg Lin3TAy(DLin), .T, .F, V, .Des
    If V <> "|" Then Stop
End With
End Function

Private Function BrkFItm(FLin) As FRslt
Dim LikFldSsl$, A$, V$
With BrkFItm.F
    AyAsg Lin3TAy(FLin), .E, .LikT, V, A
    .LikFny = SslSy(LikFldSsl)
End With
End Function

Private Function BrkEItm(ELin, Lno%) As ERslt
Dim LikFldSsl$, A$(), V$, Ty$, Ay()
With BrkEItm.E
    A = LinTermAy(ELin)
    .E = A(0)
    V = A(1)
    Ty = DaoShtTy_Ty(A(2))
    If Ty = 0 Then
        Push BrkEItm.Er, ErMsgTyEr(Lno, A(2))
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

Private Function TMsg$(TLin, Msg$)
TMsg = FmtQQ("T-Lin[?] ?", TLin, Msg)
End Function

Private Sub AAA()
Z_BrkTItm
End Sub

Private Sub Z_BrkTItm()
Dim Act As TRslt
Dim Ept As TRslt
Dim Emp As TRslt
Dim TLin$
TLin = "A"
Ept = Emp
Push Ept.Er, "should have a |"
GoSub Tst
'
TLin = "A | B B"
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
    Act = BrkTItm(TLin)
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

Private Function BrkTItm(TLin) As TRslt
If Not HasSubStr(TLin, "|") Then
    Push BrkTItm.Er, "should have a |"
    Exit Function
End If
Dim A$, B$, C$, D$
BrkAsg TLin, "|", A, B
With BrkTItm.T
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
        Push BrkTItm.Er, "should have fields after |"
        Exit Function
    End If
    Dim Dup$()
    Dup = AyWhDup(.Fny)
    If Sz(Dup) > 0 Then
        Push BrkTItm.Er, FmtQQ("dup fields[?]", JnSpc(Dup))
        Exit Function
    End If
End With
End Function

Private Function BrkD(DLy$()) As DyRslt
Dim O As DyRslt, U%, J%, Er$()
U = UB(DLy)
ReDim O.D(U)
For J = 0 To U
    With BrkDItm(DLy(J))
        O.D(J) = .D
        PushAy Er, .Er
    End With
Next
O.Er = Er
BrkD = O
End Function

Private Function BrkT(TLy$()) As TyRslt
If Sz(TLy) = 0 Then
    Push BrkT.Er, ErMsgNoTLin
    Exit Function
End If
Dim U%, J%
U = UB(TLy)
ReDim BrkT.T(U)
For J = 0 To U
    With BrkTItm(TLy(J))
        BrkT.T(J) = .T
        Dim Pfx$
        Pfx = FmtQQ("T-Lin[?] ", TLy(J))
        PushAy BrkT.Er, AyAddPfx(.Er, Pfx)
    End With
Next
End Function

Private Function BrkF(FLy$()) As FyRslt
Dim O As FyRslt, U%, J%, Er$()
U = UB(FLy)
ReDim O.F(U)
For J = 0 To U
    With BrkFItm(FLy(J))
        O.F(J) = .F
        PushAy Er, .Er
    End With
Next
O.Er = Er
BrkF = O
End Function

Private Function BrkE(ELy$(), LnoAy%()) As EyRslt
Dim O As EyRslt, U%, J%, Er$()
U = UB(ELy)
ReDim O.E(U)
For J = 0 To U
    With BrkEItm(ELy(J), LnoAy(J))
        O.E(J) = .E
        PushAy Er, .Er
    End With
Next
O.Er = Er
BrkE = O
End Function
Private Function OyStrPy(A, P) As String()
OyStrPy = OyPyInto(A, P, EmpSy)
End Function
Private Function OyPyInto(A, P, OInto)
Dim O, X
O = OInto
Erase O
If Sz(A) = 0 Then OyPyInto = O: Exit Function
For Each X In A
    Push O, ObjPrp(X, P)
Next
OyPyInto = O
End Function
Private Function IxlyLy(A() As Ixl) As String()
IxlyLy = OyStrPy(A, "Lin")
End Function
Private Function Brk(SmLines$) As Brk
Dim Cln$(), E As EyRslt, D As DyRslt, F As FyRslt, T As TyRslt, Er$(), TLy$(), ELy$(), ELnoAy%()
Cln = IxlyLy(LyClnIxly(SplitCrLf(SmLines)))
TLy = AyWhRmvT1(Cln, "T")
ELy = AyWhRmvT1(Cln, "E")
D = BrkD(AyWhRmvT1(Cln, "D"))
E = BrkE(ELy, ELnoAy)
F = BrkF(AyWhRmvT1(Cln, "F"))
T = BrkT(TLy)
Er = ClnChk(Cln, "D E F T")
Brk.Er = AyAddAp(Er, D.Er, E.Er, F.Er, , T.Er)
If Sz(Brk.Er) > 0 Then Exit Function
With Brk.Dta
    .E = E.E
    .F = F.F
    .T = T.T
    .D = D.D
    .Eny = Eny(ELy)
    .Tny = Tny(TLy)
End With
End Function

Private Function Tny(TLy$()) As String()
Tny = AyT1Ay(TLy)
End Function

Private Function ErDupT(Tny$()) As String()
ErDupT = AyDupChk(Tny, "These T[?] in T-Lin is duplicated")
End Function

Private Function ErDupE(Eny$()) As String()
ErDupE = AyDupChk(Eny, "These E[?] in E-Lin is duplicated")
End Function

Private Function ErFldEleIsInVdt(T() As T, EAy$()) As String()
Stop '
End Function

Private Function Eny(ELy$()) As String()
Eny = AyT1Ay(ELy)
End Function
Private Function Er(A As Brk) As String()
Dim D As Dta
D = A.Dta
Er = AyAddAp _
    (A.Er _
   , ErDupT(D.Tny) _
   , ErDupE(D.Eny) _
   , ErFldEleIsInVdt(D.T, D.Eny) _
    )
End Function

Private Function PkSqy(A As Dta) As String()
Dim J%, O$()
With A
    For J = 0 To UBound(.T)
        PushNonEmp O, PkSql(.T(J))
    Next
End With
PkSqy = O
End Function
Private Function PkSql$(A As T)
With A
    If AyHas(.Fny, .T) Then PkSql = SqlzCrtPk(.T)
End With
End Function

Private Function TBrkItm(T, A() As T) As T
Dim J%
For J = 0 To UBound(A)
    With A(J)
        If .T = T Then TBrkItm = A(J): Exit Function
    End With
Next
End Function

Private Function SkSql$(T, TBrk() As T)
Dim M As T
M = TBrkItm(T, TBrk)
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
Dim O() As DAO.TableDef, I
For Each I In A.Tny
    PushObj O, NewTd(I, FdAy(I, A))
Next
Td = O
End Function

Private Function FdAy(T, A As Dta) As DAO.Field()
Stop '
End Function

Private Function TDes(A As Dta) As TDes()
Stop '
End Function

Private Function FDes(A As Dta) As FDes()
Stop '
End Function
Private Function ErMsgDupT$(LnoAy%(), T$)

End Function
Private Function ErMsgTyEr$(Lno, Ty$)

End Function
Private Function ErMsgDupE$(LnoAy%(), E$)

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
Private Function ErMsgDupF$(Lno%, T$, Fny$())

End Function
Private Function Rslt(SmLines$) As Rslt
Dim B As Brk
    B = Brk(SmLines)
Dim E$()
    E = Er(B)
    If Sz(E) > 0 Then Rslt.Er = E: Exit Function
With Rslt
    Dim D As Dta
    D = B.Dta
    .Td = Td(D)
    .PkSqy = PkSqy(D)
    .SkSqy = SkSqy(D)
    .FDes = FDes(D)
    .TDes = TDes(D)
End With
End Function

Private Sub Z_DbCrtSchm()
DbCrtSchm W, ZZSmLines
End Sub

Sub AA()
Z
End Sub

Sub Z()
Z_DbCrtSchm
End Sub