Option Compare Database
Option Explicit
Const C_EF$ = "EF"
Const C_E$ = "E"
Const C_TF$ = "TF"
Const C_D$ = "D"
Function E$(F, EFLy$())
E = LinT1(EFLin(F, EFLy))
End Function

Function EFLin$(F, EFLy$())
EFLin = T1LikSslAy_T1(EFLy, F)
End Function

Function Z_ELy() As String()
Z_ELy = ELy(Z_Schmy)
End Function
Property Get Z_Schmy() As String()
Dim O$()
Push O, "dfd"
Push O, "E Mem   Mem"
Push O, "E Txt   Txt;Req;AlwZLen;Dft=Johnson;VRul=VRul;VTxt=VTxt"
Push O, "E Crt   Dte;Req;Dft=Now;"
Push O, "E Dte   Dte"
Push O, "EF . Amt *Amt"
Push O, "EF . Crt CrtDte"
Push O, "EF . Dte *Dte"
Push O, "EF . Txt Fun *Txt"
Push O, "EF . Mem Lines"
Push O, "TF Sess * CrtDte"
Push O, "TF Msg  * Fun *Txt | CrtDte"
Push O, "TF Lg   * Sess Msg CrtDte"
Push O, "TF LgV  * Lg Lines"
Push O, "D . Fun Function name that call the log"
Push O, "D . Fun Function name that call the log"
Push O, "D . Msg it will a new record when Lg-function is first time using the Fun+MsgTxt"
Push O, "D . Msg ..."
Z_Schmy = O
End Property

Function QTFELy_z() As String()
QTFELy_z = QTFELy_by_Schmy(Z_Schmy)
End Function
Function QTFELy_by_Schmy(A$()) As String()
Dim B$(), C$()
B = TFLy(A)
C = EFLy(A)
QTFELy_by_Schmy = QTFELy(B, C)
End Function
Function QTFELy(TFLy$(), EFLy$()) As String()
Dim O$(), T, F, Tny1$(), E1$
Tny1 = Tny(TFLy)
For Each T In Tny1
    For Each F In Fny(T, TFLy)
        E1 = E(F, EFLy)
        Push O, ApLin(T, F, E1)
    Next
Next
QTFELy = O
End Function

Function QTFEFdLy(TFLy$(), EFLy$(), ELy$()) As String()
Dim O$(), T, F, E1, Tny1$()
Tny1 = Tny(TFLy)
For Each T In Tny1
    For Each F In Fny(T, TFLy)
        E1 = E(F, EFLy)
        Push O, ApLin(T, F, E1, FdScl(E1, ELy))
    Next
Next
QTFEFdLy = O
End Function

Function ELy(Schmy$()) As String():  ELy = AyT1Chd(Schmy, C_E):   End Function
Function EFLy(Schmy$()) As String(): EFLy = AyT1Chd(Schmy, C_EF): End Function
Function TFLy(Schmy$()) As String(): TFLy = AyT1Chd(Schmy, C_TF): End Function
Function DLy(Schmy$()) As String():  DLy = AyT1Chd(Schmy, C_D):   End Function

Function PkTny(TFLy$()) As String()
PkTny = AyT1Ay(PkTFLy(TFLy))
End Function

Sub Z()
Z_Tny
Z_DbCrtSchm
End Sub

Sub Z_Tny()
Expect = SslSy("Sess Msg Lg LgV")
Actual = Tny(TFLy(Z_Schmy))
C
End Sub

Sub ZZ_Tny()
Dim T, Tny1, TFLy1$()
TFLy1 = TFLy(Z_Schmy)
GoSub Sep
D "Tny"
D "---"
Tny1 = Tny(TFLy(Z_Schmy))
D Tny1
GoSub Sep
For Each T In Tny1
    GoSub Prt
Next
D SkSqy(TFLy1)
D PkSqy(TFLy1)
Exit Sub
Prt:
    D T
    D UnderLin(T)
    D Fny(T, TFLy1)
    GoSub Sep
    Return
Sep:
    D "--------------------"
    Return
End Sub

Function ELin$(E, ELy$())
ELin = AyFstT1(ELy, E)
End Function
Function EScl$(E, ELy$())
EScl = LinRmvT1(ELin(E, ELy))
End Function

Function FdScl$(E, ELy$())
FdScl = EScl(E, ELy)
End Function

Function Fd(F, T, Tny$(), EFLy$(), ELy$()) As DAO.Field
Select Case True
Case IsId(T, F):   Set Fd = NewFd_zId(F)
Case IsFk(F, Tny): Set Fd = NewFd_zFk(F)
Case Else:
    Dim E1$, FdScl1$
    E1 = E(F, EFLy)
    FdScl1 = FdScl(E1, ELy)
    Set Fd = NewFd_zFdScl(FdScl1)
End Select
End Function

Function Td(T, TFLy$(), EFLy$(), ELy$()) As DAO.TableDef
Set Td = NewTd(T, FdAy(T, TFLy, EFLy, ELy))
End Function

Function Tny(TFLy$()) As String()
Tny = AyMapSy(TFLy, "LinT1")
End Function

Function TdAy(TFLy$(), EFLy$(), ELy$()) As DAO.TableDef()
TdAy = AyMapXABCInto(Tny(TFLy), "Td", TFLy, EFLy, ELy, TdAy)
End Function

Function PkSqy(TFLy$()) As String()
PkSqy = AyMapSy(PkTny(TFLy), "TnPkSql")
End Function

Function SkSslAy(TFLy$()) As String()
Dim A$(), O$(), L
A = TFLy
If Sz(A) = 0 Then Exit Function
For Each L In A
    PushNonEmp O, SkSsl(L)
Next
SkSslAy = O
End Function

Function SkSsl$(TFLin)
Dim A$, B$
A = SkP1(TFLin): If A = "" Then Exit Function
B = Replace(A, " * ", "")
SkSsl = Replace(B, "*", LinT1(B))
End Function

Function SkP1$(TFLin)
SkP1 = Trim(TakBef(TFLin, "|"))
End Function
Function PkTFLy(TFLy$()) As String()
PkTFLy = AyWhPred(TFLy, "TFLinHasPk")
End Function

Function Z_SkSqy() As String()
Z_SkSqy = SkSqy(Z_TFLy)
End Function

Function Z_TFLy() As String()
Z_TFLy = TFLy(Z_Schmy)
End Function

Function SkSqy(TFLy$()) As String()
Dim O$(), A$(), B$(), J%, U%, T
A = SkSslAy(TFLy)
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
DbCrtSchm FbDb(Fb), Z_Schmy
FbBrw Fb
End Sub

Sub DbCrtSchm(A As Database, Schmy$())
Dim TF$(), EF$(), E$(), D$()
E = ELy(Schmy)
TF = TFLy(Schmy)
EF = EFLy(Schmy)
D = DLy(Schmy)
AyDoPX TdAy(TF, EF, E), "DbAppTd", A
AyDoPX PkSqy(TF), "DbRun", A
AyDoPX SkSqy(TF), "DbRun", A
End Sub

Function TFLin$(T, TFLy$())
TFLin = AySng(AyWhT1EqV(TFLy, T), "Schm.TFLin.PrpEr")
End Function

Function Fny(T, TFLy$()) As String()
Dim A$, B$
A = TFLin(T, TFLy)
If LinShiftT1(A) <> T Then Debug.Print "Schm.Fny PrpEr": Exit Function
B = Replace(A, "*", T)
Fny = AyRmvEle(SslSy(B), "|")
End Function

Function FdAy(T, TFLy$(), EFLy$(), ELy$()) As DAO.Field()
Dim Fny1$(), Tny1$()
Tny1 = Tny(TFLy)
Fny1 = Fny(T, TFLy)
FdAy = AyMapXABCDInto(Fny1, "Fd", T, Tny1, EFLy, ELy, FdAy)
End Function

Function IsFk(F, Tny$()) As Boolean
IsFk = AyHas(Tny, F)
End Function

Function IsId(T, F) As Boolean
IsId = T = F
End Function