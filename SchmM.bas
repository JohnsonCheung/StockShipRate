Option Compare Database
Option Explicit
Const C_EF$ = "EF"
Const C_E$ = "E"
Const C_TF$ = "TF"
Const C_D$ = "D"
Private ZAF$
Private ZAT$
Function EFLyE$(A$(), T, F)
EFLyE = LinT1(EFLin(F, T, A))
End Function

Function ELyEAy(ELy$()) As String()
ELyEAy = AyT1Ay(ELy)
End Function

Private Sub Z_EFLin()
ZAT = "Msg"
ZAF = "MsgTxt"
ChkEq "Txt * Fun *Txt", ZBEFLin
End Sub
Private Property Get ZBEFLin$()
ZBEFLin = EFLin(ZAF, ZAT, ZBEFLy)
End Property
Function EFLin$(F, T, EFLy$())
EFLin = T1LikLikSslAy_FstT1T2Eq(EFLy, T, F)
End Function

Private Property Get ZBELy() As String()
ZBELy = ELy(ZASchmy)
End Property

Private Property Get ZASchmy() As String()
ZASchmy = LgSchmy
End Property

Property Get ZBQTFELy() As String()
ZBQTFELy = SchmyQTFELy(ZASchmy)
End Property

Function SchmyQTFELy(A$()) As String()
Dim B$(), C$()
B = TFLy(A)
C = EFLy(A)
SchmyQTFELy = QTFELy(B, C)
End Function

Function QTFELy(TFLy$(), EFLy$()) As String()
Dim O$(), T, F, Tny1$(), E1$
Tny1 = TFLyTny(TFLy)
For Each T In Tny1
    For Each F In TFLyFny(TFLy, T)
        E1 = EFLyE(EFLy, T, F)
        Push O, ApLin(T, F, E1)
    Next
Next
QTFELy = O
End Function

Function QTFEFdLy(TFLy$(), EFLy$(), ELy$()) As String()
Dim O$(), T, F, E1, Tny1$()
Tny1 = TFLyTny(TFLy)
For Each T In Tny1
    For Each F In TFLyFny(TFLy, T)
        E1 = EFLyE(EFLy, T, F)
        Push O, ApLin(T, F, E1, FdScl(F, E1, ELy))
    Next
Next
QTFEFdLy = O
End Function

Function EFLy(Schmy$()) As String(): EFLy = AyT1Chd(Schmy, C_EF): End Function
Function TFLy(Schmy$()) As String(): TFLy = AyT1Chd(Schmy, C_TF): End Function
Function ELy(Schmy$()) As String():  ELy = AyT1Chd(Schmy, C_E):   End Function
Function DLy(Schmy$()) As String():  DLy = AyT1Chd(Schmy, C_D):   End Function

Function PkTny(TFLy$()) As String()
PkTny = AyT1Ay(PkTFLy(TFLy))
End Function

Sub Z()
Z_Tny
Z_EFLin
Z_DbCrtSchm
End Sub

Sub ZZ_FdAy()
ZAT = "Sess"
Actual = ZBFdAy
Stop
End Sub

Private Property Get ZBEFLy() As String()
ZBEFLy = EFLy(ZASchmy)
End Property

Private Function ZBFdAy() As dao.Field()
ZBFdAy = FdAy(ZAT, ZBTFLy, ZBEFLy, ZBELy)
End Function

Private Property Get ZBTny() As String()
ZBTny = TFLyTny(ZBTFLy)
End Property

Private Sub Z_Tny()
ChkEq ZBTny, SslSy("Sess Msg Lg LgV")
End Sub

Private Sub ZZ_Tny()
Dim T, Tny$(), TFLy$()
TFLy = ZBTFLy
Tny = ZBTny
GoSub Sep
D "Tny"
D "---"
D ZBTny
GoSub Sep
For Each T In Tny
    GoSub Prt
Next
D SkSqy(TFLy)
D PkSqy(TFLy)
Exit Sub
Prt:
    D T
    D UnderLin(T)
    D TFLyFny(TFLy, T)
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

Function FdScl$(F, E, ELy$())
FdScl = F & ";" & EScl(E, ELy)
End Function

Function Fd(F, T, Tny$(), EFLy$(), ELy$()) As dao.Field
Select Case True
Case IsId(T, F):   Set Fd = NewFd_zId(F)
Case IsFk(F, Tny): Set Fd = NewFd_zFk(F)
Case Else:
    Dim E1$, FdScl1$
    E1 = EFLyE(EFLy, T, F)
    FdScl1 = FdScl(F, E1, ELy)
    Set Fd = NewFd_zFdScl(FdScl1)
End Select
End Function

Function Td(T, TFLy$(), EFLy$(), ELy$()) As dao.TableDef
Set Td = NewTd(T, FdAy(T, TFLy, EFLy, ELy))
End Function

Function TFLyTny(TFLy$()) As String()
TFLyTny = AyMapSy(TFLy, "LinT1")
End Function

Function TdAy(TFLy$(), EFLy$(), ELy$()) As dao.TableDef()
TdAy = AyMapXABCInto(TFLyTny(TFLy), "Td", TFLy, EFLy, ELy, TdAy)
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

Private Property Get ZBSkSqy() As String()
ZBSkSqy = SkSqy(ZBTFLy)
End Property

Private Property Get ZBTFLy() As String()
ZBTFLy = TFLy(ZASchmy)
End Property

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
Sub AAA()
Z_DbCrtSchm
End Sub
Private Sub Z_DbCrtSchm()
Dim Fb$
Fb = TmpFb
DbCrtSchm FbCrt(Fb), ZASchmy
Kill Fb
End Sub

Function TfEfED_Er(TF$(), EF$(), ELy$(), D$()) As String()
Dim Tny1$(), EAy1$(), A1$(), A2$(), A3$(), A4$(), A5$()
EAy1 = ELyEAy(ELy)
Tny1 = TFLyTny(TF)
A1 = ErDupE(EAy1)
A2 = ErDupF(Tny1, TF)
A3 = ErDupT(Tny1)
A4 = ErE(EAy1)
A5 = ErFldHasNoEle(Tny1, TF, EF)
TfEfED_Er = AyAddAp(A1, A2, A3, A4, A5)
End Function
Function ErDupT(Tny$()) As String()
ErDupT = AyDupChk(Tny, "These T[?] is duplicated in TFld-lines")
End Function

Function ErDupE(EAy$()) As String()
ErDupE = AyDupChk(EAy, "These Ele[?] are duplicated in Ele-lines")
End Function

Function ErDupF(Tny$(), TF$()) As String()
Dim T
For Each T In AyNz(Tny)
    PushAy ErDupF, AyDupChk(TFLyFny(TF, T), FmtQQ("These F[?] are duplicated in T[?]", "?", T))
Next
End Function
Function ELinEr(A$) As String()
Dim Ay$()
Ay = SplitSC(A)

End Function
Function ErE(ELy$()) As String()
ErE = AyOfAy_Ay(AyMap(ELy, "ELinEr"))
End Function

Function ErFldHasNoEle(Tny$(), TFLy$(), EFLy$()) As String()
Dim T, F, E1$
For Each T In AyNz(Tny)
    For Each F In AyNz(TFLyFny(TFLy, T))
        If T = F Then GoTo Nxt
        If AyHas(Tny, F) Then GoTo Nxt
        E1 = EFLyE(EFLy, T, F)
        PushNonEmp ErFldHasNoEle, StrEmpChkMsg(E1, FmtQQ("T[?] F[?] has no TEle", T, F))
Nxt:
    Next
Next
End Function

Function ErNoTF(TF$()) As String()
ErNoTF = AyEmpChk(TF, "No TFld lines")
End Function

Sub SchmyAsg(A, OEr$(), OTF$(), OEF$(), OE$(), OD$())
Dim Ny$(), Er$()
    Ny = Sy(C_TF, C_EF, C_E, C_D)
    AyAsg LyBrk1(A, Ny), OTF, OEF, OE, OD, OEr
End Sub

Sub DbCrtSchm(A As Database, Schmy$())
Dim Er1$(), TF$(), EF$(), E$(), D$()
    SchmyAsg Schmy, Er1, TF, EF, E, D
Dim Er$()
    Er = AyAdd(TfEfED_Er(TF, EF, E, D), Er1)
    Er = AyAdd_zIf_B_IsNonEmp(Schmy, Er)
AyBrwThw Er
Dim Tny1$()
    Tny1 = TFLyTny(TF)
AyDoPX TdAy(TF, EF, E), "DbAppTd", A
AyDoPX PkSqy(TF), "DbRun", A
AyDoPX SkSqy(TF), "DbRun", A
AyDoPX DLyTDesLy(D), "DbSetTblDes", A
AyDoPX DLyTFDesLy(D, Tny1), "DbSetTFDes", A
End Sub

Function SchmyEr(A) As String()
Dim Er1$(), TF$(), EF$(), E$(), D$(), Tny$()
SchmyAsg A, Er1, TF, EF, E, D
SchmyEr = AyAdd(TfEfED_Er(TF, EF, E, D), Er1)
End Function

Function DLyTDesLy(DLy$()) As String()

End Function

Function DLyTFDesLy(DLy$(), Tny$()) As String()

End Function

Function TFLin$(T, TFLy$())
TFLin = AySng(AyWhT1EqV(TFLy, T), "Schm.TFLin.PrpEr")
End Function

Function TFLyFny(TFLy$(), T) As String()
Dim A$, B$
A = TFLin(T, TFLy)
If LinShiftT1(A) <> T Then Debug.Print "Schm.Fny PrpEr": Exit Function
B = Replace(A, "*", T)
TFLyFny = AyRmvEle(SslSy(B), "|")
End Function

Function FdAy(T, TFLy$(), EFLy$(), ELy$()) As dao.Field()
Dim Fny$(), Tny$()
Tny = TFLyTny(TFLy)
Fny = TFLyFny(TFLy, T)
FdAy = AyMapXABCDInto(Fny, "Fd", T, Tny, EFLy, ELy, FdAy)
End Function

Function IsFk(F, Tny$()) As Boolean
IsFk = AyHas(Tny, F)
End Function

Function IsId(T, F) As Boolean
IsId = T = F
End Function