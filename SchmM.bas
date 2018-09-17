Option Compare Database
Option Explicit
Const C_F$ = "F"
Const C_E$ = "E"
Const C_T$ = "T"
Const C_D$ = "D"
Private ZZF$
Private ZZT$

Function TF_E$(T, F, FLy$())
TF_E = LinT1(TF_FLin(T, F, FLy))
End Function

Function ELy_EAy(ELy$()) As String()
ELy_EAy = AyT1Ay(ELy)
End Function

Private Sub Z_FLin()
ZZT = "Msg"
ZZF = "MsgTxt"
ChkEq "Txt * Fun *Txt", ZZFLin
End Sub

Private Property Get ZZFLin$()
ZZFLin = TF_FLin(ZZF, ZZT, ZZFLy)
End Property

Function TF_FLin$(T, F, FLy$())
TF_FLin = T1LikLikSslAy_FstT2T3Eq(FLy, T, F)
End Function

Private Property Get ZZELy() As String()
ZZELy = ToELy(ZZSchmy)
End Property

Private Property Get ZZSchmy() As String()
ZZSchmy = LgIniSchmy
End Property

Private Property Get ZZQTFELy() As String()
ZZQTFELy = SchmyQTFELy(ZZSchmy)
End Property

Function SchmyQTFELy(A$()) As String()
SchmyQTFELy = QTFELy(ToTLy(A), ToFLy(A))
End Function

Function QTFELy(TLy$(), FLy$()) As String()
Dim O$(), T, F, Tny1$(), E1$
Tny1 = TLyTny(TLy)
For Each T In Tny1
    For Each F In TLyFny(TLy, T)
        E1 = TF_E(T, F, FLy)
        Push O, ApLin(T, F, E1)
    Next
Next
QTFELy = O
End Function

Function ToTFEFdLy(TLy$(), FLy$(), ELy$()) As String()
Dim O$(), T, F, E$, Tny$()
Tny = TLyTny(TLy)
For Each T In Tny
    For Each F In TLyFny(FLy, T)
        E = TF_E(T, F, ELy)
        Push O, ApLin(T, F, E, FE_FdScl(F, E, ELy))
    Next
Next
ToTFEFdLy = O
End Function

Function ToFLy(Schmy$()) As String(): ToFLy = AyT1Chd(Schmy, C_F): End Function
Function ToTLy(Schmy$()) As String(): ToTLy = AyT1Chd(Schmy, C_T): End Function
Function ToELy(Schmy$()) As String(): ToELy = AyT1Chd(Schmy, C_E):   End Function
Function ToDLy(Schmy$()) As String(): ToDLy = AyT1Chd(Schmy, C_D):   End Function

Function PkTny(TLy$()) As String()
PkTny = AyT1Ay(PkTLy(TLy))
End Function

Sub Z()
Z_Tny
Z_FLin
Z_DbCrtSchm
End Sub

Private Sub ZZ_FdAy()
ZZT = "Sess"
Actual = ZZFdAy
Stop
End Sub

Private Property Get ZZFLy() As String()
ZZFLy = ToFLy(ZZSchmy)
End Property

Private Function ZZFdAy() As DAO.Field()
ZZFdAy = FdAy(ZZT, ZZTny, ZZTLy, ZZFLy, ZZELy)
End Function

Private Property Get ZZTny() As String()
ZZTny = TLyTny(ZZTLy)
End Property

Private Sub Z_Tny()
ChkEq ZZTny, SslSy("Sess Msg Lg LgV")
End Sub

Private Sub ZZ_Tny()
Dim T, Tny$(), TLy$()
TLy = ZZTLy
Tny = ZZTny
GoSub Sep
D "Tny"
D "---"
D ZZTny
GoSub Sep
For Each T In Tny
    GoSub Prt
Next
D TLy_SkSqy(TLy)
D TLy_PkSqy(TLy)
Exit Sub
Prt:
    D T
    D UnderLin(T)
    D TLyFny(TLy, T)
    GoSub Sep
    Return
Sep:
    D "--------------------"
    Return
End Sub

Function E_ELin$(E, ELy$())
E_ELin = AyFstT1(ELy, E)
End Function

Function E_EScl$(E, ELy$())
E_EScl = LinRmvT1(E_ELin(E, ELy))
End Function

Function FE_FdScl$(F, E, ELy$())
FE_FdScl = F & ";" & E_EScl(E, ELy)
End Function

Function Fd(F, T, Tny$(), FLy$(), ELy$()) As DAO.Field
Select Case True
Case IsId(T, F):   Set Fd = NewFd_zId(F)
Case IsFk(F, Tny): Set Fd = NewFd_zFk(F)
Case Else:
    Dim E$, FdScl$
    E = TF_E(T, F, FLy)
    FdScl = FE_FdScl(F, E, ELy)
    Set Fd = NewFd_zFdScl(FdScl)
End Select
End Function

Function Td(T, Tny$(), TLy$(), FLy$(), ELy$()) As DAO.TableDef
Set Td = NewTd(T, FdAy(T, Tny, TLy, FLy, ELy))
End Function

Function TLyTny(FLy$()) As String()
TLyTny = AyT1Ay(FLy)
End Function


Function TdAy(Tny$(), TLy$(), FLy$(), ELy$()) As DAO.TableDef()
TdAy = AyMapXABCDInto(TLyTny(TLy), "Td", Tny, TLy, FLy, ELy, TdAy)
End Function

Function TLy_PkSqy(TLy$()) As String()
TLy_PkSqy = AyMapSy(PkTny(TLy), "TnPkSql")
End Function
Function SkSslAy(TLy$()) As String()
Dim A$(), O$(), L
A = TLy
If Sz(A) = 0 Then Exit Function
For Each L In A
    PushNonEmp O, TFLin_SkSsl(L)
Next
SkSslAy = O
End Function

Function TFLin_SkSsl$(A)
Dim B$, C$
B = Trim(TakBef(A, "|")): If B = "" Then Exit Function
C = Replace(B, " * ", " ")
TFLin_SkSsl = Replace(C, "*", LinT1(A))
End Function

Function PkTLy(TLy$()) As String()
PkTLy = AyWhPred(TLy, "TFLinHasPk")
End Function

Private Property Get ZZSkSqy() As String()
ZZSkSqy = TLy_SkSqy(ZZTLy)
End Property

Private Property Get ZZTLy() As String()
ZZTLy = ToTLy(ZZSchmy)
End Property

Function TLy_SkSqy(TLy$()) As String()
TLy_SkSqy = AyMapSy(SkSslAy(TLy), "TnSkSsl_SkSql")
End Function

Private Sub Z_DbCrtSchm()
Dim Fb$
Fb = TmpFb
DbCrtSchm FbCrt(Fb), ZZSchmy
Kill Fb
End Sub

Function XEr(T$(), F$(), E$(), D$()) As String()
Dim Tny1$(), EAy1$(), A1$(), A2$(), A3$(), A4$(), A5$(), A6$()
EAy1 = ELy_EAy(E)
Tny1 = TLyTny(T)
A1 = ErDupE(EAy1)
A2 = ErDupF(Tny1, T)
A3 = ErDupT(Tny1)
A4 = ErE(EAy1)
A5 = ErFldHasNoEle(Tny1, T, F)
A6 = ErFEle_NotIn_EAy(F, EAy1)
XEr = AyAddAp(A1, A2, A3, A4, A5, A6)
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
    PushAy ErDupF, AyDupChk(TLyFny(TF, T), FmtQQ("These F[?] are duplicated in T[?]", "?", T))
Next
End Function

Function ELinChk(ByVal A$) As String()
ELinChk = SclChk(TakAft(A, ";"), VdtEleSclNmSsl)
End Function

Function ErE(ELy$()) As String()
ErE = AyOfAy_Ay(AyMap(ELy, "ELinChk"))
End Function

Function ErFldHasNoEle(Tny$(), TLy$(), FLy$()) As String()
Dim T, F, E$
For Each T In AyNz(Tny)
    For Each F In AyNz(TLyFny(TLy, T))
        If T = F Then GoTo Nxt
        If AyHas(Tny, F) Then GoTo Nxt
        E = TF_E(T, F, FLy)
        If E = "" Then
            Push ErFldHasNoEle, FmtQQ("T[?] F[?] cannot be found in any EF-lines", T, F)
        End If
Nxt:
    Next
Next
End Function

Function FLy_EAy(FLy$()) As String()
FLy_EAy = AyT1Ay(FLy)
End Function

Function ErFEle_NotIn_EAy(FLy$(), EAy$()) As String()
Dim E$(), Er$()
E = FLy_EAy(FLy)
Er = AyMinus(E, EAy)
If Sz(Er) = 0 Then Exit Function
ErFEle_NotIn_EAy = MsgLy("These [Ele] in F-lines are not found in the elements of E-lines", JnSpc(Er))
End Function

Function ErNoTF(T$()) As String()
ErNoTF = AyEmpChk(T, "No TFld lines")
End Function

Sub SchmyAsg(A, OEr$(), OT$(), OF$(), OE$(), OD$())
Dim Ny$(), Er$()
    Ny = Sy(C_T, C_F, C_E, C_D)
    AyAsg LyBrk1(A, Ny), OT, OF, OE, OD, OEr
End Sub
Sub DbCrtSchm_z1(A As Database, Er$(), TdAyFun$, PkSqyFun$, SkSqyFun$, DLyTDesLyFun$, DLyTFDesLyFun$)

End Sub
Sub DbCrtSchm(A As Database, Schmy$())
Dim Er1$(), TF$(), EF$(), E$(), D$()
    SchmyAsg Schmy, Er1, TF, EF, E, D
Dim Er$()
    Er = AyAdd(XEr(TF, EF, E, D), Er1)
    If Sz(Er) > 0 Then Er = AyAdd(Schmy, Er)
AyBrwThw Er
Dim Tny$()
    Tny = TLyTny(TF)
AyDoPX TdAy(Tny, TF, EF, E), "DbAppTd", A
AyDoPX TLy_PkSqy(TF), "DbRun", A
AyDoPX TLy_SkSqy(TF), "DbRun", A
AyDoPX DLyTDesLy(D), "DbSetTblDes", A
AyDoPX DLyFDesLy(D, Tny), "DbSetFDes", A
End Sub

Function SchmyEr(A) As String()
Dim Er1$(), T$(), F$(), E$(), D$(), Tny$()
SchmyAsg A, Er1, T, F, E, D
SchmyEr = AyAdd(XEr(T, F, E, D), Er1)
End Function

Function DLyTDesLy(DLy$()) As String()

End Function

Function DLyFDesLy(DLy$(), Tny$()) As String()

End Function

Function ToTLin$(T, TLy$())
ToTLin = AySng(AyWhT1(TLy, T), "Schm.TFLin.PrpEr")
End Function

Function TLyFny(TLy$(), T) As String()
Dim A$, B$
A = ToTLin(T, TLy)
If LinShiftT1(A) <> T Then Debug.Print "Schm.Fny PrpEr": Exit Function
B = Replace(A, "*", T)
TLyFny = AyRmvEle(SslSy(B), "|")
End Function

Function FdAy(T, Tny$(), TLy$(), FLy$(), ELy$()) As DAO.Field()
Dim Fny$()
Fny = TLyFny(TLy, T)
FdAy = AyMapXABCDInto(Fny, "Fd", T, Tny, FLy, ELy, FdAy)
End Function

Function IsFk(F, Tny$()) As Boolean
IsFk = AyHas(Tny, F)
End Function

Function IsId(T, F) As Boolean
IsId = T = F
End Function