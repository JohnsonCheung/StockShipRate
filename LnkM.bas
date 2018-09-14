Option Compare Database
Option Explicit
Type ChkFilDta: Kind As String: Ffn As String: End Type
Type ChkTnyDta: FbKind As String: Fb As String: Tny() As String: End Type
Type ChkWnyDta: FxKind As String: Fx As String: Wny() As String: End Type
Type ChkColDta: Kind As String: Ffn As String: Fld As String: TyAy() As ADODB.DataTypeEnum: End Type
Type FR: Er() As String: OkFilKind() As String: End Type ' FilRslt
Type WR: Er() As String: OkWny() As String:  End Type ' WnyRslt
Type TR: Er() As String: OkTny() As String:  End Type ' TnyRslt
Type CR: Er() As String:                     End Type ' ColRslt
Type SqlSelIntoDta: T As String: Ny() As String: Ey() As String: End Type
Sub LNKDbImp(Db As Database)
Dim F() As ChkFilDta
Dim W() As ChkWnyDta
Dim T() As ChkTnyDta
Dim C() As ChkColDta
    F = ChkFilDta
    W = ChkWnyDta
    T = ChkTnyDta
    C = ChkColDta
    Stop '
Dim FR As FR, WR As WR, TR As TR, CR As CR
    Dim OkFfny$(), OkWny$(), OkTny$()
    FR = ChkFil(F)
    WR = ChkWny(W, FR.OkFilKind)
    TR = ChkTny(T, FR.OkFilKind)
    CR = ChkCol(C, CvSy(AyAdd(WR.OkWny, TR.OkTny)))
AyBrwThw AyAddAp(FR.Er, WR.Er, TR.Er, CR.Er)
Dim FxDry(), FbDry()
    FxDry = Dry_Db_T_Fx_Ws(Db)
    FbDry = Dry_Db_T_Fb_Fbt(Db)
C4DryDo FxDry, "DbtLnkFx"
C4DryDo FbDry, "DbtLnkFx"
DbRunSqy Db, SqlSelIntoDtaAy_Sqy(SqlSelIntoDta)
End Sub

Function SqlSelIntoDtaAy_Sqy(A() As SqlSelIntoDta) As String()
Dim J%, O$()
For J = 0 To UBound(A)
    Push O, SqlSelIntoDta_Sql(A(J))
Next
SqlSelIntoDtaAy_Sqy = O
End Function

Private Function SqlSelIntoDta() As SqlSelIntoDta()

End Function
Private Function ChkFilDta() As ChkFilDta()

End Function
Private Function ChkWnyDta() As ChkWnyDta()

End Function
Private Function ChkTnyDta() As ChkTnyDta()

End Function
Private Function ChkColDta() As ChkColDta()

End Function
Function ChkFil(A() As ChkFilDta) As FR
Dim J%
For J = 0 To UBound(A)
'    If FilChk(A(J)) = "" Then
'    End If
Next
End Function

Private Function ChkTny(A() As ChkTnyDta, OkFilKind$()) As TR

End Function

Private Function ChkWny(A() As ChkWnyDta, OkFilKind$()) As WR

End Function

Private Function ChkCol(A() As ChkColDta, OkFilKind$()) As CR

End Function

Private Function Sql$(A As SqlSelIntoDta)

End Function
Function SqlSelIntoDta_Sql$(A As SqlSelIntoDta)

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
Private Function LyStuInp() As String()
LyStuInp = LyXXX("StuInp")
End Function
Private Function FldInpy() As String()
FldInpy = AyT1Ay(LyFld)
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
Private Function IniActInpy() As String()
IniActInpy = AyMinus(Inpy, InActInpy)
End Function
Private Function Inpy() As String()

End Function
Private Function ActFldLy() As String()
ActFldLy = AyWhT1InAy(LyFld, ActInpy)
End Function
Private Function LyFld() As String()
LyFld = LyXXX("Fld")
End Function
Private Function LyXXX(XXX$) As String()
LyXXX = AyWhRmvT1(NoT1, XXX)
End Function
Private Function LyExt() As String()
LyExt = LyXXX("Ext")
End Function
Private Function LyPmSw() As String()
LyPmSw = LyXXX("PmSw")
End Function
Private Function LySw() As String()
LySw = LyXXX("Sw")
End Function
Private Function TWhLy() As String()
'TWhLy = T1Ly("A-Wh")
End Function
Private Function ImpSql$(FldLin$)
Dim T$, Fny$()
T = LinT1(FldLin)
Fny = SslSy(LinRmvTT(FldLin))

End Function
Private Function IniSqy() As String()
IniSqy = AyMapSy(ActFldLy, "ImpSql")
End Function
Private Function IniInpy() As String()
Inpy = SslSy(LinRmv3T(AyFstTT(Cln, "a", "Inp")))
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
Private Function InActInpy() As String()
Dim Sw$(), PmSw$(), L, O$()
PmSw = LyPmSw
If Sz(PmSw) = 0 Then Exit Function
Sw = LySw

Dim SwNy$(), T_or_F_Ay$()
    Dim T_or_F
    AyTAyRstAyAsg PmSw, SwNy, T_or_F_Ay
    For Each T_or_F In T_or_F_Ay
        Select Case T_or_F
        Case "T", "F"
        Case Else: Stop
        End Select
    Next
For Each L In Sw
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
Private Function Dry_zTT_Fb_Fbtt() As Variant()
Dry_zTT_Fb_Fbtt = AyZip3(FbTTy(A), ActFby, XFb_Fbtty)
End Function
Private Function Dry_zDb_TT_Fb_Fbtt() As Variant()
Dry_zDb_TT_Fb_Fbtt = DryInsConst(Dry_zTT_Fb_Fbtt, Db)
End Function
Sub LnkFb()
C4DryDo Dry_zDb_TT_Fb_Fbtt(Db), "DbttFb"
End Sub
Private Sub AssWs()
Stop '
End Sub
Private Sub AssFil()
FilLinAyAss ActFilLinAy
End Sub
Private Function ActFilLinAy() As String()
Stop '
End Function
Private Function T1() As String()
T1 = AyDistT1Ay(Cln)
End Function
Private Function WhT1(T1$) As String()
WhT1 = AyWhRmvT1(Cln, T1)
End Function