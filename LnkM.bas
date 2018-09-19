Option Compare Database
Option Explicit
Private Type A1
Er() As String
Dryz_Db_T_Fb_Fbt() As Variant
Dryz_Db_T_Fx_Ws() As Variant
Sqy() As String
End Type
Type LnkSpec
    AFb() As LnkAFil
    AFx() As LnkAFil
    ASw() As LnkASw
    FmFb() As LnkFmFil
    FmFx() As LnkFmFil
    FmIp() As String
    FmStu() As LnkFmStu
    FmSw() As LnkFmSw
    FmWh() As LnkFmWh
    IpFb() As LnkIpFil
    IpFx() As LnkIpFil
    IpS1() As String
    IpWs() As LnkIpWs
    StEle() As LnkStEle
    StExt() As LnkStExt
    StFld() As LnkStFld
End Type
Type FR: Er() As String: OkFilKind() As String: End Type ' FilRslt
Type Wr: Er() As String: OkWny() As String:  End Type ' WnyRslt
Type TR: Er() As String: OkTny() As String:  End Type ' TnyRslt
Type CR: Er() As String:                     End Type ' ColRslt
Sub AA()
Z_LNKDbImp
End Sub

Private Sub Z_LNKDbImp()
LNKDbImp WDb, LNKAllLines
End Sub

Sub LNKDbImp(Db As Database, LnkSpec$)
With A1(Db, LnkSpec)
AyBrwThw AyAddAp(.Er)
C4DryDo .Dryz_Db_T_Fx_Ws, "DbtLnkFx"
C4DryDo .Dryz_Db_T_Fb_Fbt, "DbtLnkFb"
DbRunSqy Db, .Sqy
End With
End Sub
Private Function A1(Db As Database, LnkSpec$) As A1
Dim A As LnkSpec
    A = NewLnkSpec(LnkSpec)
A1.Er = A2Er(A)
If Sz(A1.Er) > 0 Then Exit Function
A1.Dryz_Db_T_Fx_Ws = Dry_Db_T_Fx_Ws(Db)
A1.Dryz_Db_T_Fb_Fbt = Dry_Db_T_Fb_Fbt(Db)

Dim InAct1$():     InAct1 = InActInpy(A.ASw, A.FmSw)
Dim ActInpy1$(): ActInpy1 = ActInpy(A.FmIp, InAct1)
Dim B() As SqlSelInto:  B = SelIntoAy(ActInpy1, A)
                   A1.Sqy = AyMapSy(B, "ToSql_by_SqlSelInto")
End Function
Function ActInpy(FmIp$(), InAct$()) As String()
Dim Inpy$():   Inpy = SslSy(AyWhRmvTT(NoT1, "Inp", "|")(0))
ActInpy = AyMinus(Inpy, InAct)
End Function
Private Function A2Er(A As LnkSpec) As String()
Dim F() As ChkFil
Dim T() As ChkTbl
Dim C() As ChkCol
    F = NewChkFil(A)
    T = NewChkTbl(A)
    C = NewChkCol(A)
    Stop '
Dim FR As FR, TR As TR, CR As CR
    Dim OkFfny$(), OkWny$(), OkTny$()
    FR = ChkFil(F)
    TR = ChkTbl(T, FR.OkFilKind)
    CR = ChkCol(C, TR.OkTny)

End Function

Private Function SelIntoAy(ActInpy$(), A As LnkSpec) As SqlSelInto()
Dim Inp$, I, J%, O() As SqlSelInto
ReDim O(UB(ActInpy))
For Each I In ActInpy
    Set O(J) = New SqlSelInto
    With O(J)
        Inp = I
'        .Ny = InpNy(Inp, A.StInp, A.StFld)
        .Ey = NyEy(.Ny, A.StEle)
        .Fm = ">" & Inp
        .Into = "#I" & Inp
        .Wh = InpWhBExpr(Inp, A.FmWh)
    End With
    J = J + 1
Next
SelIntoAy = O
End Function

Private Function InpWhBExpr$(Inp$, FmWh() As LnkFmWh)

End Function

Private Function ZZCln() As String()
ZZCln = LyCln(SplitCrLf(LNKAllLines))
End Function

Private Function NewLnkSpec(LnkSpec$) As LnkSpec
Dim Cln$():   Cln = LyCln(SplitCrLf(LnkSpec))
Dim AFx() As LnkAFil
Dim AFb() As LnkAFil
Dim ASw() As LnkASw

Dim FmFx() As LnkFmFil
Dim FmFb() As LnkFmFil
Dim FmIp() As String
Dim FmSw() As LnkFmSw
Dim FmWh() As LnkFmWh
Dim FmStu() As LnkFmStu

Dim IpFx() As LnkIpFil
Dim IpFb() As LnkIpFil
Dim IpS1() As String
Dim IpWs() As LnkIpWs

Dim StEle() As LnkStEle
Dim StExt() As LnkStExt
Dim StFld() As LnkStFld
    
    FmIp = SslSy(AyWhRmvTT(Cln, "FmIp", "|")(0))
    FmSw = NewFmSw(AyWhRmvT1(Cln, "IpSw"))
    FmFx = NewFmFil(AyWhRmvTT(Cln, "IpFx", "|"))
    FmFb = NewFmFil(AyWhRmvTT(Cln, "IpFb", "|"))
    FmWh = NewFmWh(AyWhRmvT1(Cln, "FmWh"))
    IpS1 = AyWhRmvTT(Cln, "IpS1", "|")
    IpWs = NewIpWs(AyWhRmvTT(Cln, "IpWs", "|"))
Stop

With NewLnkSpec
    .AFx = AFx
    .AFb = AFb
    .ASw = ASw
    .FmFx = FmFx
    .FmFb = FmFb
    .FmIp = FmIp
    .FmSw = FmSw
    .FmStu = FmStu
    .FmWh = FmWh
    .IpFx = IpFx
    .IpFb = IpFb
    .IpS1 = IpS1
    .IpWs = IpWs
    .StEle = StEle
    .StExt = StExt
    .StFld = StFld
End With
End Function
Function NewFmWh(Ly$())

End Function

Function NewFmSw(Ly$())

End Function
Function NewFmFil(Ly$())

End Function
Sub SrtDcl()
MdSrtDclDim Md("LnkM")
End Sub

Function NewIpWs(Ly$()) As LnkIpWs()

End Function

Function NewIpFil(Ly$()) As LnkIpFil()
If Sz(Ly) = 0 Then Exit Function
Dim O() As LnkIpFil, J%, L, Ay
ReDim O(UB(Ly))
For Each L In Ly
    Ay = AyT1Rst(SslSy(L))
    Set O(J) = New LnkIpFil
    O(J).Fil = Ay(0)
    O(J).Inpy = CvSy(Ay(1))
    J = J + 1
Next
NewIpFil = O
End Function

Function NewStExt(Lin) As LnkStExt
Dim O As New LnkStExt
With O
    AyAsg Lin3TAy(Lin), .LikInp, .F, , .Ext
End With
Set NewStExt = O
End Function

Private Function NyEy(Ny$(), A() As LnkStEle) As String()

End Function

Private Function NewStFld(Lin) As LnkStFld
Dim O As New LnkStFld, A$
With O
    AyAsg Lin2TAy(Lin), .Stu, , A
    .Fny = SslSy(A)
End With
Set NewStFld = O
End Function
Private Function InpExt(Inp$, A As LnkSpec) As String()

End Function
Private Function NewChkFil(A As LnkSpec) As ChkFil()
Dim O() As ChkFil

NewChkFil = O
End Function
Private Function NewChkTbl(A As LnkSpec) As ChkTbl()
Stop '
End Function
Private Function NewChkCol(A As LnkSpec) As ChkCol()

End Function
Function ChkFil(A() As ChkFil) As FR
Dim M() As ChkFil, Er$()
Dim J%
'For J = 0 To UBound(A)
'    If FilChk(A(J)) = "" Then
'    End If
'Next
With ChkFil
    .Er = Er
End With
End Function

Private Function ChkTbl(A() As ChkTbl, OkFilKind$()) As TR

End Function

Private Function ChkCol(A() As ChkCol, OkFilKind$()) As CR

End Function
Private Function Dry_Db_T_Fx_Ws(A As Database) As Variant()
Stop '
End Function
Function Dry_Db_T_Fb_Fbt(A As Database) As Variant()
Stop '
End Function
Sub A()
Stop
End Sub

Private Function TblColLy(T) As String()
TblColLy = AyRmvT1(AyWhT1(Cln, T))
End Function

Private Function ClnSrt() As String()
ClnSrt = AySrt(Cln)
End Function

Private Function Cln() As String()
Cln = LyCln(SplitCrLf(LNKAllLines))
End Function

Private Function Tny() As String()
Tny = AySrt(AyDistT1Ay(Cln))
End Function

Private Function LyStuInp(NoT1$()) As String()
LyStuInp = LyXXX(NoT1, "StuInp")
End Function

Private Function FldInpy(NoT1$()) As String()
FldInpy = AyT1Ay(LyFld(NoT1))
End Function

Private Function LnkFt$()
LnkFt = SpnmFt("Lnk")
End Function

Private Function FbNy() As String()
FbNy = AyT1Ay(LyFb)
End Function
Private Function LyFb() As String()
LyFb = AyRmvT1(AyWhT1(Cln, "0-Fb"))
End Function
Private Sub AImp()
SpnmImp "A"
End Sub
Private Sub AEdt()
SpnmEdt "A"
End Sub
Private Function FxNy() As String()
FxNy = AyT1Ay(FxLy)
End Function
Private Function Fxy() As String()
Fxy = AyRmvT1(FxLy)
End Function
Private Function FxLy() As String()
FxLy = AyWhRmv3T(Cln, "0", "A", "Fx")
End Function
Private Function NoT1() As String()
NoT1 = AyWhNotPfx(AyRmvT1(Cln), "/")
End Function
Private Function ActFldLy(ActInpy$(), LyFld$()) As String()
ActFldLy = AyWhT1InAy(LyFld, ActInpy)
End Function
Private Function LyFld(NoT1$()) As String()
LyFld = LyXXX(NoT1, "Fld")
End Function
Private Function LyXXX(NoT1$(), XXX$) As String()
LyXXX = AyWhRmvT1(NoT1, XXX)
End Function
Private Function LyExt(NoT1$()) As String()
LyExt = LyXXX(NoT1, "Ext")
End Function

Private Function ImpSql$(FldLin$)
Dim T$, Fny$()
T = LinT1(FldLin)
Fny = SslSy(LinRmvTT(FldLin))

End Function
Private Sub AssTbl()
Stop '
End Sub
Private Sub AssCol()
Stop '
End Sub
Private Function InActInpy__zSel(SwNm$, TF As Boolean, ASw() As LnkASw) As Boolean
Dim IA As LnkASw, I
For Each I In ASw
    Set IA = I
    If SwNm = IA.SwNm Then
        InActInpy__zSel = IA.TF = TF
        Exit Function
    End If
Next
Stop
End Function
Private Function InActInpy(ASw() As LnkASw, FmSw() As LnkFmSw) As String()
Dim O$(), I, IFm As LnkFmSw, SwNm$, TF As Boolean
If Sz(ASw) = 0 Then Exit Function

For Each I In FmSw
    Set IFm = I
    SwNm = IFm.SwNm
    TF = IFm.TF
    If Not InActInpy__zSel(SwNm, TF, ASw) Then
        PushAy O, SslSy(LinRmv3T(IFm.Inpy))
    End If
Next
InActInpy = O
End Function
Private Function Dry_zT_Fx_Ws() As Variant()

End Function
Private Function Dry_zDb_T_Fx_Ws(A As Database) As Variant()
Dry_zDb_T_Fx_Ws = DryInsConst(Dry_zT_Fx_Ws, A)
End Function
Private Function ActFby() As String()

End Function
Private Function FbTTy(A) As String()

End Function
Private Function FbFbtty(A$, TTy$) As String()

End Function
Private Function Dry_zTT_Fb_Fbtt(Db As Database) As Variant()
'Dry_zTT_Fb_Fbtt = AyZip3(FbTTy(Db), ActFby, XFb_Fbtty)
End Function
Private Function Dry_zDb_TT_Fb_Fbtt(Db As Database) As Variant()
Dry_zDb_TT_Fb_Fbtt = DryInsConst(Dry_zTT_Fb_Fbtt(Db), Db)
End Function
Private Sub AssWs()
Stop '
End Sub
Private Sub AssFil()
'FilLinAyAss ActFilLinAy
End Sub

Private Function ActFilLinAy() As String()
Stop '
End Function
Private Function Ky() As String()
Ky = AyDistT1Ay(Cln)
End Function
Private Function WhT1(T1$) As String()
WhT1 = AyWhRmvT1(Cln, T1)
End Function
Sub Z()
Z_LNKDbImp
End Sub