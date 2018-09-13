Option Compare Database
Option Explicit
'LS = LnkSpec see LSLines
Sub LSDbImp(A As Database, LSAll$())
Dim W1(), W2()
AyBrwThw LSxChkFst(LSAll)
W1 = DryInsConst(LSClnDry_zT_Fx_Ws(LSAll), A)
W2 = DryInsConst(LSClnDry_zTT_Fb_Fbtt(LSAll), A)
C4DryDo W1, "DbtLnkFx"
C3DryDo W2, "DbttLnkFb"
AyBrwThw LSxChkCol(LSAll)
DbRunSqy A, LSxImpSqy(LSAll)
End Sub
Function LSxActInpy(LSAll$()) As String()
Dim Inpy$(), InAct$()
Inpy = LSxInpy(LSAll)
InAct = LSxInActInpy(LSAll)
LSxActInpy = AyMinus(Inpy, InAct)
End Function

Function LSxInActInpy(LSAll$()) As String()
Dim Wny$(), Tny$()
Wny = LSxInActWny(LSAll)
Tny = LSxInActTny(LSAll)
LSxInActInpy = AyAdd(Wny, Tny)
End Function

Function LSxInActWny(LSAll$()) As String()

End Function

Function LSxInActTny(LSAll$()) As String()

End Function

Function LSxInActFx(LSAll$()) As String()

End Function

Function LSxInActFb(LSAll$()) As String()

End Function

Function LSxInpy(LSAll$()) As String()
LSxInpy = AyWhRmvT1(LSAll, "Inp")
End Function

Function LSxImpSqy(LSAll$()) As String()
Dim Inpy$()
Inpy = LSxActInpy(LSAll)
LSxImpSqy = AyMapXPSy(Inpy, "LSxImpSql", LSAll)
End Function

Function LSxImpSql(Inp$, LSAll$()) As String()

End Function

Function LSxChkCol(LSAll$()) As String()

End Function
Function LSxActFxWnyLy(LSAll$()) As String()

End Function
Function LSxActFbTnyLy(LSAll$()) As String()

End Function
Function LSxChkFst(LSAll$()) As String()
Dim Fx$(), Fb$()
Fx = LSxActFxWnyLy(LSAll)
Fb = LSxActFbTnyLy(LSAll)
LSxChkFst = AyAlign1T(AyAdd(FxWnyLy_Chk(Fx), FbTnyLy_Chk(Fb)))
End Function
Function FxWnyLin_Chk(A$) As String()

End Function
Function FbTnyLin_Chk(A$) As String()

End Function
Function FxWnyLy_Chk(A$()) As String()

End Function
Function FbTnyLy_Chk(A$()) As String()

End Function

Function LSClnDry_zTT_Fb_Fbtt(LSAll$()) As Variant()

End Function

Function LSClnDry_zT_Fx_Ws(LSAll$()) As Variant()

End Function

Function LSClnLyStruFld(A$()) As String()
LSClnLyStruFld = AyWhT1(A, "StruFld")
End Function

Function LSClnLyStruExt(A$()) As String()
LSClnLyStruExt = AyWhT1(A, "StruExt")
End Function

Function LSClnLyInpWh(A$()) As String()
LSClnLyInpWh = AyWhT1(A, "InpWh")
End Function

Function LSClnImpSqy(A$()) As String()
LSClnImpSqy = AyMapXABSy(LSClnLyStruFld(A), "LSFldLin_ImpSql", LSClnLyStruExt(A), LSClnLyInpWh(A))
End Function

Function LSFldLin_ImpSql$(A$, ExtLy$(), WhLy$())
Dim T$, FldSsl$
Dim Fm$, Into$, Ny$(), Ey$(), Wh$
LinShiftTermAsg A, T, FldSsl
Fm = ">" & T
Into = "#I" & T
Ny = SslSy(FldSsl)
'Ey = TTXAyXy(ExtLy, T, Ny)
'Ey = InpEy(A)
'Wh = SqpWh(InpSqWhBExpr(A))
LSFldLin_ImpSql = SqlSelInto(Fm, Into, Ny, Ey, Wh)
End Function

Function LSInpWhBExpr$(A)
'LSInpWhBExpr = AyFstT1(LST1Ly("A-Wh"), A)
End Function

Function LSTfExtNm$(T, F, ExtLy$())
LSTfExtNm = TTXAyFst(ExtLy, T, F)
End Function