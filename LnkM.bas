Option Compare Database
Option Explicit
Private Type A1
Er() As String
Dryz_Db_T_Fb_Fbt() As Variant
Dryz_Db_T_Fx_Ws() As Variant
Sqy() As String
End Type
Type FR: Er() As String: OkFilKind() As String: End Type ' FilRslt
Type Wr: Er() As String: OkWny() As String:  End Type ' WnyRslt
Type TR: Er() As String: OkTny() As String:  End Type ' TnyRslt
Type CR: Er() As String:                     End Type ' ColRslt
Sub AA()
Z_LNKDbImp
End Sub
Private Sub Z_LNKDbImp()
LNKDbImp WDb, SampleLnkSpec
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
    Set A = NewLnkSpec(LnkSpec)
A1.Er = A2Er(A)
If Sz(A1.Er) > 0 Then Exit Function
A1.Dryz_Db_T_Fx_Ws = Dry_Db_T_Fx_Ws(Db)
A1.Dryz_Db_T_Fb_Fbt = Dry_Db_T_Fb_Fbt(Db)
Dim ActInpy$(): Stop '
Dim B() As SqlSelInto:    B = SelIntoAy(ActInpy, A)
A1.Sqy = AyMapSy(B, "ToSql_by_SqlSelInto")
End Function
Private Function A2Er(A As LnkSpec) As String()
Dim Inpy$():   Inpy = SslSy(AyWhRmvTT(NoT1, "Inp", "|")(0))
Dim InAct$(): InAct = InActInpy(NoT1)
Dim Act$():     Act = AyMinus(Inpy, InAct)       'ActInpy
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
        .Ny = InpNy(Inp, A.StInp, A.StFld)
        .Ey = NyEy(.Ny, A.StEle)
        .Fm = ">" & Inp
        .Into = "#I" & Inp
        .Wh = InpWhBExpr(Inp, A.IpWh)
    End With
    J = J + 1
Next
SelIntoAy = O
End Function
Function ZZCln() As String()
ZZCln = LyCln(SplitCrLf(SampleLnkSpec))
End Function
Function ZZNoT1() As String()
ZZNoT1 = AyWhNotPfx(AyRmvT1(ZZCln), "/")
End Function
Private Function NewLnkSpec(LnkSpec$) As LnkSpec
Dim NoT1$():   NoT1 = AyWhNotPfx(AyRmvT1(LyCln(SplitCrLf(LnkSpec))), "/")
Dim PmFx() As PmFil
Dim PmFb() As PmFil
Dim PmSw() As PmSw

Dim Inp() As String
Dim IpSw() As IpSw
Dim IpFx() As IpFil
Dim IpFb() As IpFil
Dim IpS1() As String
Dim IpWs() As IpWs
Dim IpWh() As IpWh

Dim StInp() As StInp
Dim StEle() As StEle
Dim StExt() As StExt
Dim StFld() As StFld
    
    Inp = SslSy(AyWhRmvTT(NoT1, "Inp", "|")(0))
    IpSw = NewIpSw(AyWhRmvT1(NoT1, "IpSw"))
    IpFx = NewIpFil(AyWhRmvTT(NoT1, "IpFx", "|"))
    IpFb = NewIpFil(AyWhRmvTT(NoT1, "IpFb", "|"))
    IpS1 = AyWhRmvTT(NoT1, "IpS1", "|")
    IpWs = NewIpWs(AyWhRmvTT(NoT1, "IpWs", "|"))
Stop

Dim O As New LnkSpec
Set NewLnkSpec = O.Init(PmFx, PmFb, PmSw, Inp, IpSw, IpFx, IpFb, IpS1, IpWs, IpWh, StInp, StEle, StExt, StFld)
End Function
Function NewIpSw(Ly$()) As IpSw()

End Function
Function NewIpWs(Ly$()) As IpWs()

End Function
Function NewIpFil(Ly$()) As IpFil()
If Sz(Ly) = 0 Then Exit Function
Dim O() As IpFil, J%, L, Ay
ReDim O(UB(Ly))
For Each L In Ly
    Ay = AyT1Rst(SslSy(L))
    Set O(J) = New IpFil
    O(J).Fil = Ay(0)
    O(J).Inpy = CvSy(Ay(1))
    J = J + 1
Next
NewIpFil = O
End Function
Function NewStExt(Lin) As StExt
Dim O As New StExt
With O
    AyAsg Lin3TAy(Lin), .LikInp, .F, , .Ext
End With
Set NewStExt = O
End Function

Private Function NyEy(Ny$(), A() As StEle) As String()

End Function
Private Function InpFldAy(ActInp$(), NoT1$()) As InpFld()
Dim StuInpLy$(), FldLy$()


End Function

Private Function NewStFld(Lin) As StFld
Dim O As New StFld, A$
With O
    AyAsg Lin2TAy(Lin), .Stu, , A
    .Fny = SslSy(A)
End With
Set NewStFld = O
End Function
Private Function InpExt(Inp$, A As LnkSpec) As String()

End Function
Private Function InpNy(Inp$, StInp() As StInp, StFld() As StFld) As String()
'InpNy = ObjPrp(OyWhPrpEqV(B, "Inp", Inp), "Fny")
End Function

Private Function InpWhBExpr$(Inp$, IpWh() As IpWh)
Stop '
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
Private Sub PmImp()
SpnmImp "Pm"
End Sub
Private Sub PmEdt()
SpnmEdt "Pm"
End Sub
Private Function FxNy() As String()
FxNy = AyT1Ay(FxLy)
End Function
Private Function Fxy() As String()
Fxy = AyRmvT1(FxLy)
End Function
Private Function FxLy() As String()
FxLy = AyWhRmv3T(Cln, "0", "Pm", "Fx")
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
Private Function TWhLy() As String()
'TWhLy = T1Ly("A-Wh")
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
Private Function InActInpy__zSel(L, SwNy$(), T_or_F_Ay$()) As Boolean
Dim Sw$, T_or_F$, Ix%
LinTTAsg L, Sw, T_or_F
If T_or_F <> "T" And T_or_F <> "F" Then Stop
Ix = AyIx(SwNy, Sw)
InActInpy__zSel = T_or_F_Ay(Ix) <> T_or_F
End Function
Private Function InActInpy(NoT1$()) As String()
Dim PmSwLy$(): PmSwLy = AyWhRmvT1(NoT1, "PmSw") '
Dim SwLy$():     SwLy = AyWhRmvT1(NoT1, "Sw")
Dim L, O$()
If Sz(PmSwLy) = 0 Then Exit Function

Dim SwNy$(), T_or_F_Ay$()
    Dim T_or_F
    AyTAyRstAyAsg PmSwLy, SwNy, T_or_F_Ay
    For Each T_or_F In T_or_F_Ay
        Select Case T_or_F
        Case "T", "F"
        Case Else: Stop
        End Select
    Next
For Each L In SwLy
    If InActInpy__zSel(L, SwNy, T_or_F_Ay) Then
        PushAy O, SslSy(LinRmv3T(L))
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