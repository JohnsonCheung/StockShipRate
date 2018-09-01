Option Compare Database
Option Explicit
Const C_FEle$ = "FEle"
Const C_Ele$ = "Ele"
Const C_TFld$ = "TFld"
Const C_TDes$ = "TDes"
Const C_FDes$ = "FDes"
Function E$(F, T, Tny, LyFEle)
Select Case True
Case IsId(T, F): E = "*Id"
Case IsFk(F, Tny): E = "*Fk"
Case Else: E = LinT1(LinFEle(LyFEle, F))
End Select
End Function

Function LinFEle$(LyFEle, F)
Dim L
If Sz(LyFEle) = 0 Then Exit Function
For Each L In LyFEle
    If StrInLikSsl(F, LinRmvT1(L)) Then
        LinFEle = L
        Exit Function
    End If
Next
End Function

Public Function Z_Schmy() As String() ' Z_Ly
Dim O$()
Push O, "dfd"
Push O, "Ele Mem   Mem"
Push O, "Ele Txt   Txt;Req;AlwZLen;Dft=Johnson;VRul=VRul;VTxt=VTxt"
Push O, "Ele Crt   Dte;Req;Dft=Now;"
Push O, "Ele Dte   Dte"
Push O, "FEle Amt *Amt"
Push O, "FEle Crt CrtDte"
Push O, "FEle Dte *Dte"
Push O, "FEle Txt Fun *Txt"
Push O, "FEle Mem Lines"
Push O, "TFld Sess * CrtDte"
Push O, "TFld Msg  * Fun *Txt | CrtDte"
Push O, "TFld Lg   * Sess Msg CrtDte"
Push O, "TFld LgV  * Lg Lines"
Push O, "FDes Fun Function name that call the log"
Push O, "FDes Fun Function name that call the log"
Push O, "TDes Msg it will a new record when Lg-function is first time using the Fun+MsgTxt"
Push O, "TDes Msg ..."
Z_Schmy = O
End Function

Function TFELy(LyTFld, LyFEle) As String()
Dim O$(), T, F, Tny1$(), E1$
Tny1 = Tny(LyTFld)
For Each T In Tny1
    For Each F In Fny(T, LyTFld)
        E1 = E(F, T, Tny1, LyFEle)
        Push O, ApLin(T, F, E1)
    Next
Next
TFELy = O
End Function

Function TFEFdLy(LyTFld, LyFEle, EleLy$()) As String()
Dim O$(), T, F, E1, Tny1$()
Tny1 = Tny(LyTFld)
For Each T In Tny1
    For Each F In Fny(T, LyTFld)
        E1 = E(F, T, Tny1, LyFEle)
        Push O, ApLin(T, F, E1, FdStr(T, F, E1, Tny1, EleLy))
    Next
Next
TFEFdLy = O
End Function

Function AyT1Chd(A, T1) As String()
AyT1Chd = AyRmvT1(AyWhT1EqV(A, T1))
End Function

Function Ly_Er() As String()
Ly_Er = AyWhPredXPNot(Ly, "LinInT1Ay", Sy(C_Ele, C_FDes, C_TDes, C_TFld))
End Function

Function EleLy(Schmy$()) As String():  EleLy = AyT1Chd(Schmy, C_Ele):   End Function
Function LyFEle(Schmy$()) As String(): LyFEle = AyT1Chd(Schmy, C_FEle): End Function
Function LyTFld(Schmy$()) As String(): LyTFld = AyT1Chd(Schmy, C_TFld): End Function
Function LyFDes(Schmy$()) As String(): LyFDes = AyT1Chd(Schmy, C_FDes): End Function
Function LyTDes(Schmy$()) As String(): LyTDes = AyT1Chd(Schmy, C_TDes): End Function
Function PkTny(LyTFld$()) As String()
PkTny = AyT1Ay(PkTFLy(LyTFld))
End Function

Sub Z()
Z_Tny
Z_DbCrtSchm
End Sub

Sub Z_Tny()
Expect = SslSy("Sess Msg Lg LgV")
Actual = Tny(LyTFld(Z_Ly))
C
End Sub

Sub ZZ_Tny()
Dim T, Tny1, LyTFld1$()
LyTFld1 = LyTFld(Z_Schmy)
GoSub Sep
D "Tny"
D "---"
Tny1 = Tny(LyTFld(Z_Schmy))
D Tny1
GoSub Sep
For Each T In Tny1
    GoSub Prt
Next
D SkSqy(LyTFld1)
D PkSqy(LyTFld1)
Exit Sub
Prt:
    D T
    D UnderLin(T)
    D Fny(T, LyTFld1)
    GoSub Sep
    Return
Sep:
    D "--------------------"
    Return
End Sub

Function EleLin$(EleLy$(), E)
EleLin = AyFstT1(EleLy, E)
End Function

Function EleSpecStr$(E, EleLy$())
EleSpecStr = LinRmvT1(EleLin(EleLy, E))
End Function

Function FdDr(F, E, EleLy$()) As Variant()
With FdSpec(F, E, EleLy)
FdDr = Array(.F, .Ty, .Sz, .Req, .AlwZLen, .Dft, .VRul, .VTxt)
End With
End Function

Function FdStr$(F, E, T, Tny, EleLy$())
FdStr = FdM.FdStr(Fd(F, E, T, Tny, EleLy$()))
End Function

Function Fd(F, E, T, Tny, EleLy$()) As dao.Field
Select Case True
Case IsId(T, F): Set Fd = NewFd_zId(F)
Case IsFk(F, Tny): Set Fd = NewFd_zFk(F)
Case Else: Set Fd = NewFd_zSpec(FdSpec(F, E, EleLy))
End Select
End Function

Function FdSpec(F, E, EleLy$()) As FdSpec
FdSpec = EleSpecStr_FdSpec(EleSpecStr(E, EleLy), F)
End Function

Function Td(T, LyTFld, LyFEle, EleLy$()) As dao.TableDef
Set Td = NewTd(T, FdAy(T, LyTFld, LyFEle, EleLy))
End Function

Function Tny(LyTFld) As String()
Tny = AyMapSy(LyTFld, "LinT1")
End Function

Function TdAy(LyTFld, LyFEle, EleLy$()) As dao.TableDef()
Dim O() As dao.TableDef, T
For Each T In Tny(LyTFld)
    PushObj O, Td(T, LyTFld, LyFEle, EleLy)
Next
TdAy = O
End Function

Function PkSqy(LyTFld$()) As String()
PkSqy = AyMapSy(PkTny(LyTFld), "TnPkSql")
End Function

Function SkSslAy(LyTFld) As String()
Dim A$(), O$(), L
A = LyTFld
If Sz(A) = 0 Then Exit Function
For Each L In A
    PushNonEmpty O, SkSsl(L)
Next
SkSslAy = O
End Function

Function SkSsl$(L)
Dim A$, B$
A = SkP1(L): If A = "" Then Exit Function
B = Replace(A, " * ", "")
SkSsl = Replace(B, "*", LinT1(B))
End Function

Function SkP1$(L)
SkP1 = Trim(TakBef(L, "|"))
End Function
Function PkTFLy(LyTFld$()) As String()
PkTFLy = AyWhPred(LyTFld, "TFLinHasPk")
End Function

Function SkSqy(LyTFld) As String()
Dim O$(), A$(), B$(), J%, U%, T
A = SkSslAy(LyTFld)
U = UB(A)
If UB(A) = -1 Then Exit Function
ReDim O(U)
For J = 0 To U
    T = LinShiftT1(A(J))
    O(J) = TnSkSql(T, A(J))
Next
SkSqy = O
End Function

Sub Z_DbCrtSchm()
Dim Fb$
Fb = TmpFb
FbCrt Fb
DbCrtSchm1 FbDb(Fb), Z_Ly
FbBrw Fb
End Sub

Sub DbCrtSchm1(A As Database, SchmLy$())
Dim LyTFld1$(), LyFEle1$(), EleLy1$()
EleLy1 = EleLy(SchmLy)
LyTFld1 = LyTFld(SchmLy)
LyFEle1 = LyFEle(SchmLy)
AyDoPX TdAy(LyTFld1, LyFEle1, EleLy1), "DbAppTd", A
AyDoPX PkSqy(LyTFld1), "DbRun", A
AyDoPX SkSqy(LyTFld1), "DbRun", A
End Sub

Function TFLin$(T, LyTFld)
TFLin = AySng(AyWhT1EqV(LyTFld, T), "Schm.TFLin.PrpEr")
End Function

Function Fny(T, LyTFld) As String()
Dim A$, B$
A = TFLin(T, LyTFld)
If LinShiftT1(A) <> T Then Debug.Print "Schm.Fny PrpEr": Exit Function
B = Replace(A, "*", T)
Fny = AyRmvEle(SslSy(B), "|")
End Function

Function FdAy(T, LyTFld, LyFEle, EleLy$()) As dao.Field()
Dim O() As dao.Field, F, E1$, Tny1$()
Tny1 = Tny(LyTFld)
For Each F In Fny(T, LyTFld)
    E1 = E(F, T, Tny1, LyFEle)
    PushObj O, Fd(F, E1, T, Tny1, EleLy)
Next
FdAy = O
End Function

Function IsFk(F, Tny) As Boolean
IsFk = AyHas(Tny, F)
End Function

Function IsId(T, F) As Boolean
IsId = T = F
End Function