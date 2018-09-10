Option Compare Database
Option Explicit
Function SPLStarLy() As String()
SPLStarLy = AyWhHasPfx(SPLLy, "*")
End Function

Function SPLLFldLy() As String()
SPLLFldLy = AyWhRmvT1(SPLCln, "C-Fld")
End Function
Function SPLInpFny(A$) As String()
SPLInpFny = SslSy(SPLInpFldSsl(A))
End Function

Function SPLInpFldSsl$(A$)
SPLInpFldSsl = LinRmvT1(AyFstT1(SPLFldLy, A))
End Function
Function SPLLines$()
SPLLines = SpnmLines("Lnk")
End Function

Function SPLTfExt$(T, F, ExtLy$())
SPLTfExt = TTXAyFst(ExtLy, T, F)
End Function

Function SPLLExtLy() As String()
SPLLExtLy = SPLLxxxLy("C-Ext")
End Function

Function SPLLxxxLy(XXX$) As String()
SPLLxxxLy = AyWhRmvT1(SPLCln, XXX)
End Function

Function SPLTFnyEy(ByVal T, Fny$()) As String()
SPLTFnyEy = AyMapAXBSy(Fny, "SplTfExt", T, SPLExtLy)
End Function

Function SPLFldLin_ImpSql$(ByVal A$)
Dim Fm$, Into$, Ny$(), Ey$(), Wh$, T$
T = LinShiftTerm(A)
Fm = ">" & T
Into = "#I" & T
Ny = SslSy(A)
Ey = SPLTFnyEy(T, Ny)
Wh = SPLInpWhBExpr(A)
SPLFldLin_ImpSql = SqlSelInto(Fm, Into, Ny, Ey, Wh)
End Function

Function SPLInpWhBExpr$(A)
SPLInpWhBExpr = AyFstT1(SPLT1Ly("A-Wh"), A)
End Function

Function SPLT1Ly(T1$)
SPLT1Ly = AyWhRmvT1(SPLCln, T1)
End Function

Function SPLTWhLy() As String()
SPLTWhLy = SPLT1Ly("A-Wh")
End Function

Function SPLCln() As String()
SPLCln = LyCln(SPLLy)
End Function

Function SPLLy() As String()
SPLLy = SplitCrLf(SPLLines)
End Function
Function SPLSqy() As String()
SPLSqy = AyMapSy(SPLFldLy, "SPLFldLin_ImpSql")
End Function

Function SPLInpy() As String()
SPLInpy = SLKClnInpy(SPLCln)
End Function