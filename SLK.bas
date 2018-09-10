Option Compare Database
Option Explicit
'SLK = SpecLnk see StkShpRateSLKLines
Function SLKClnXFilChk(A$()) As String()

End Function

Function SLKClnXWsChk(A$()) As String()

End Function

Function SLKClnXTblChk(A$()) As String()

End Function

Function SLKClnXColChk(A$()) As String()

End Function
Function SLKClnXFstChk(A$()) As String()
SLKClnXFstChk = AyAddAp(SLKClnXFilChk(A), SLKClnXWsChk(A), SLKClnXTblChk(A))
End Function
Sub DbImp(A As Database, SLKLy$(), SwFxFbLy$())
Dim Cln$(), W1(), W2(), Sqy$(), FxLy$(), FbLy$()
Cln = LyCln(A)
FxLy = AyWhRmvT1(SwFxFbLy, "Fx")
FbLy = AyWhRmvT1(SwFxFbLy, "Fb")
AyBrwThw SLKClnXFstChk(Cln)
W1 = DryInsConst(SLKClnDry_zT_Fx_Ws(Cln, FxLy), A)
W2 = DryInsConst(SLKClnDry_zTT_Fb_Fbtt(Cln, FbLy), A)
Sqy = SLKClnImpSqy(Cln)
C4DryDo W1, "DbtLnkFx"
C3DryDo W2, "DbttLnkFb"
AyBrwThw SLKClnXColChk(LSLy)
DbRunSqy A, Sqy
End Sub
Function SLKClnDry_zTT_Fb_Fbtt(A$(), FbLy$()) As Variant()

End Function

Function SLKClnDry_zT_Fx_Ws(A$(), FxLy$()) As Variant()

End Function

Function SLKClnLFldLy(A$()) As String()
SLKClnLFldLy = AyWhT1(A, "D-Fld")
End Function

Function SLKClnLExtLy(A$()) As String()
SLKClnLExtLy = AyWhT1(A, "D-Fld")
End Function

Function SLKClnLWhLy(A$()) As String()
SLKClnLWhLy = AyWhT1(A, "D-Fld")
End Function

Function SLKClnImpSqy(A$()) As String()
SLKClnImpSqy = AyMapXABSy(SLKClnLFldLy(A), "SLKFldLin_ImpSql", SLKClnLExtLy(A), SLKClnLWhLy(A))
End Function

Function SLKFldLin_ImpSql$(ByVal A$, ExtLy$(), WhLy$())
Dim T$
Dim Fm$, Into$, Ny$(), Ey$(), Wh$
T = LinShiftTerm(A)
Fm = ">" & T
Into = "#I" & T
Ny = SslSy(A)
'Ey = TTXAyXy(ExtLy, T, Ny)
Ey = InpEy(A)
'Wh = SqpWh(InpSqWhBExpr(A))
SLKFldLin_ImpSql = SqlSelInto(Fm, Into, Ny, Ey, Wh)
End Function