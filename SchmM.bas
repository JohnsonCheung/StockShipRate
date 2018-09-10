Option Compare Database
Option Explicit
Const C_ETF$ = "ETF"
Const C_E$ = "E"
Const C_TF$ = "TF"
Const C_D$ = "D"
Private ZAF$
Private ZAT$
Function TF_E$(T, F, ETFLy$())
TF_E = LinT1(TF_ETFLin(T, F, ETFLy))
End Function

Function ELyEAy(ELy$()) As String()
ELyEAy = AyT1Ay(ELy)
End Function

Private Sub Z_ETFLin()
ZAT = "Msg"
ZAF = "MsgTxt"
ChkEq "Txt * Fun *Txt", ZBETFLin
End Sub
Private Property Get ZBETFLin$()
ZBETFLin = TF_ETFLin(ZAF, ZAT, ZBETFLy)
End Property
Function TF_ETFLin$(T, F, ETFLy$())
TF_ETFLin = T1LikLikSslAy_FstT2T3Eq(ETFLy, T, F)
End Function

Private Property Get ZBELy() As String()
ZBELy = ELy(ZASchmy)
End Property

Private Property Get ZASchmy() As String()
ZASchmy = LgIniSchmy
End Property

Property Get ZBQTFELy() As String()
ZBQTFELy = SchmyQTFELy(ZASchmy)
End Property

Function SchmyQTFELy(A$()) As String()
Dim B$(), C$()
B = TFLy(A)
C = ETFLy(A)
SchmyQTFELy = QTFELy(B, C)
End Function

Function QTFELy(TFLy$(), ETFLy$()) As String()
Dim O$(), T, F, Tny1$(), E1$
Tny1 = TFLyTny(TFLy)
For Each T In Tny1
    For Each F In TFLyFny(TFLy, T)
        E1 = TF_E(T, F, ETFLy)
        Push O, ApLin(T, F, E1)
    Next
Next
QTFELy = O
End Function

Function QTFEFdLy(TFLy$(), ETFLy$(), ELy$()) As String()
Dim O$(), T, F, E$, Tny$()
Tny = TFLyTny(TFLy)
For Each T In Tny
    For Each F In TFLyFny(TFLy, T)
        E = TF_E(T, F, ETFLy)
        Push O, ApLin(T, F, E, FE_FdScl(F, E, ELy))
    Next
Next
QTFEFdLy = O
End Function

Function ETFLy(Schmy$()) As String(): ETFLy = AyT1Chd(Schmy, C_ETF): End Function
Function TFLy(Schmy$()) As String(): TFLy = AyT1Chd(Schmy, C_TF): End Function
Function ELy(Schmy$()) As String():  ELy = AyT1Chd(Schmy, C_E):   End Function
Function DLy(Schmy$()) As String():  DLy = AyT1Chd(Schmy, C_D):   End Function

Function PkTny(TFLy$()) As String()
PkTny = AyT1Ay(PkTFLy(TFLy))
End Function

Sub Z()
Z_Tny
Z_ETFLin
Z_DbCrtSchm
End Sub

Sub ZZ_FdAy()
ZAT = "Sess"
Actual = ZBFdAy
Stop
End Sub

Private Property Get ZBETFLy() As String()
ZBETFLy = ETFLy(ZASchmy)
End Property

Private Function ZBFdAy() As dao.Field()
ZBFdAy = FdAy(ZAT, ZBTFLy, ZBETFLy, ZBELy)
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
D TFLy_SkSqy(TFLy)
D TFLy_PkSqy(TFLy)
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

Function E_ELin$(E, ELy$())
E_ELin = AyFstT1(ELy, E)
End Function

Function E_EScl$(E, ELy$())
E_EScl = LinRmvT1(E_ELin(E, ELy))
End Function

Function FE_FdScl$(F, E, ELy$())
FE_FdScl = F & ";" & E_EScl(E, ELy)
End Function

Function Fd(F, T, Tny$(), ETFLy$(), ELy$()) As dao.Field
Select Case True
Case IsId(T, F):   Set Fd = NewFd_zId(F)
Case IsFk(F, Tny): Set Fd = NewFd_zFk(F)
Case Else:
    Dim E$, FdScl$
    E = TF_E(T, F, ETFLy)
    FdScl = FE_FdScl(F, E, ELy)
    Set Fd = NewFd_zFdScl(FdScl)
End Select
End Function

Function Td(T, TFLy$(), ETFLy$(), ELy$()) As dao.TableDef
Set Td = NewTd(T, FdAy(T, TFLy, ETFLy, ELy))
End Function

Function TFLyTny(TFLy$()) As String()
TFLyTny = AyMapSy(TFLy, "LinT1")
End Function

Function TdAy(TFLy$(), ETFLy$(), ELy$()) As dao.TableDef()
TdAy = AyMapXABCInto(TFLyTny(TFLy), "Td", TFLy, ETFLy, ELy, TdAy)
End Function

Function TFLy_PkSqy(TFLy$()) As String()
TFLy_PkSqy = AyMapSy(PkTny(TFLy), "TnPkSql")
End Function

Function SkSslAy(TFLy$()) As String()
Dim A$(), O$(), L
A = TFLy
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

Function PkTFLy(TFLy$()) As String()
PkTFLy = AyWhPred(TFLy, "TFLinHasPk")
End Function

Private Property Get ZBSkSqy() As String()
ZBSkSqy = TFLy_SkSqy(ZBTFLy)
End Property

Private Property Get ZBTFLy() As String()
ZBTFLy = TFLy(ZASchmy)
End Property

Function TFLy_SkSqy(TFLy$()) As String()
TFLy_SkSqy = AyMapSy(SkSslAy(TFLy), "TnSkSsl_SkSql")
End Function

Private Sub Z_DbCrtSchm()
Dim Fb$
Fb = TmpFb
DbCrtSchm FbCrt(Fb), ZASchmy
Kill Fb
End Sub

Function TfEtfED_Er(TF$(), ETF$(), ELy$(), D$()) As String()
Dim Tny1$(), EAy1$(), A1$(), A2$(), A3$(), A4$(), A5$(), A6$()
EAy1 = ELyEAy(ELy)
Tny1 = TFLyTny(TF)
A1 = ErDupE(EAy1)
A2 = ErDupF(Tny1, TF)
A3 = ErDupT(Tny1)
A4 = ErE(EAy1)
A5 = ErFldHasNoEle(Tny1, TF, ETF)
A6 = ErETFEle_NotIn_EAy(ETF, EAy1)
TfEtfED_Er = AyAddAp(A1, A2, A3, A4, A5, A6)
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

Function ELinChk(ByVal A$) As String()
ELinChk = SclChk(TakAft(A, ";"), VdtEleSclNmSsl)
End Function

Function ErE(ELy$()) As String()
ErE = AyOfAy_Ay(AyMap(ELy, "ELinChk"))
End Function

Function ErFldHasNoEle(Tny$(), TFLy$(), ETFLy$()) As String()
Dim T, F, E$
For Each T In AyNz(Tny)
    For Each F In AyNz(TFLyFny(TFLy, T))
        If T = F Then GoTo Nxt
        If AyHas(Tny, F) Then GoTo Nxt
        E = TF_E(T, F, ETFLy)
        If E = "" Then
            Push ErFldHasNoEle, FmtQQ("T[?] F[?] cannot be found in any EF-lines", T, F)
        End If
Nxt:
    Next
Next
End Function

Function ETF_EAy(ETF$()) As String()
ETF_EAy = AyT1Ay(ETF)
End Function

Function ErETFEle_NotIn_EAy(ETFLy$(), EAy$()) As String()
Dim E$(), Er$()
E = ETF_EAy(ETFLy)
Er = AyMinus(E, EAy)
If Sz(Er) = 0 Then Exit Function
ErETFEle_NotIn_EAy = MsgLy("These [Ele] in ETF-lines are not found in the elements of E-lines", JnSpc(Er))
End Function

Function ErNoTF(TF$()) As String()
ErNoTF = AyEmpChk(TF, "No TFld lines")
End Function

Sub SchmyAsg(A, OEr$(), OTF$(), OETF$(), OE$(), OD$())
Dim Ny$(), Er$()
    Ny = Sy(C_TF, C_ETF, C_E, C_D)
    AyAsg LyBrk1(A, Ny), OTF, OETF, OE, OD, OEr
End Sub

Sub DbCrtSchm(A As Database, Schmy$())
Dim Er1$(), TF$(), EF$(), E$(), D$()
    SchmyAsg Schmy, Er1, TF, EF, E, D
Dim Er$()
    Er = AyAdd(TfEtfED_Er(TF, EF, E, D), Er1)
    Er = AyAdd_zIf_B_IsNonEmp(Schmy, Er)
AyBrwThw Er
Dim Tny1$()
    Tny1 = TFLyTny(TF)
AyDoPX TdAy(TF, EF, E), "DbAppTd", A
AyDoPX TFLy_PkSqy(TF), "DbRun", A
AyDoPX TFLy_SkSqy(TF), "DbRun", A
AyDoPX DLyTDesLy(D), "DbSetTblDes", A
AyDoPX DLyTFDesLy(D, Tny1), "DbSetTFDes", A
End Sub

Function SchmyEr(A) As String()
Dim Er1$(), TF$(), ETF$(), E$(), D$(), Tny$()
SchmyAsg A, Er1, TF, ETF, E, D
SchmyEr = AyAdd(TfEtfED_Er(TF, ETF, E, D), Er1)
End Function

Function DLyTDesLy(DLy$()) As String()

End Function

Function DLyTFDesLy(DLy$(), Tny$()) As String()

End Function

Function TFLin$(T, TFLy$())
TFLin = AySng(AyWhT1(TFLy, T), "Schm.TFLin.PrpEr")
End Function

Function TFLyFny(TFLy$(), T) As String()
Dim A$, B$
A = TFLin(T, TFLy)
If LinShiftT1(A) <> T Then Debug.Print "Schm.Fny PrpEr": Exit Function
B = Replace(A, "*", T)
TFLyFny = AyRmvEle(SslSy(B), "|")
End Function

Function FdAy(T, TFLy$(), ETFLy$(), ELy$()) As dao.Field()
Dim Fny$(), Tny$()
Tny = TFLyTny(TFLy)
Fny = TFLyFny(TFLy, T)
FdAy = AyMapXABCDInto(Fny, "Fd", T, Tny, ETFLy, ELy, FdAy)
End Function

Function IsFk(F, Tny$()) As Boolean
IsFk = AyHas(Tny, F)
End Function

Function IsId(T, F) As Boolean
IsId = T = F
End Function