Option Compare Database
Option Explicit
'-- LdDTim --------------------------------------------------
Property Get LdDTim_xInv$()
LdDTim_xInv = QQDTim("Select IR_LoadDte from YM where Y=? and M=? and IR_Fx='?'", Y, M, IFxInv)
End Property
Property Get LdDTim_xRate$()
LdDTim_xRate = QQDTim("Select RateSc_LoadDte from YM where Y=? and M=?", Y, M)
End Property
Property Get LdDTim_xMB52$()
LdDTim_xMB52 = QQDTim("Select EndOH_LoadDte from YM where Y=? and M=? and EndOH_Fx='?'", Y, M, IFxMB52)
End Property
'-- LdTSz --------------------------------------------------
Function FxLdTSz$(A, FldPfx$)
Dim P$
P = FldPfx
FxLdTSz = RsTSz(QQRs("Select ?_FxTim,?_FxSz from YM where ?_Fx='?' and Y=? and M=?", P, P, P, A, Y, M))
End Function
Function IniFx_LdTSz$(A, FldPfx$)
Dim P$
P = FldPfx
IniFx_LdTSz = RsTSz(QQRs("Select ?_FxSz, ?_FxTim from IniYM where ?_Fx='?'", P, P, P, A))
End Function
'---
Property Get LdTSz_xInv$()
LdTSz_xInv = FxLdTSz(IFxInv, "IR")
End Property

Property Get LdTSz_xMB52$()
LdTSz_xMB52 = FxLdTSz(IFxMB52, "EndOH")
End Property
'-- IsLd --------------------------------------------------
Property Get IsLd_xMB52() As Boolean
IsLd_xMB52 = FfnTSz(IFxMB52) = LdTSz_xMB52
End Property
Property Get IsLd_xRate() As Boolean
Dim A$, S$
'A = LdTSz_xRate
If S = "" Then Exit Property
If LdDTim_xInv > A Then Exit Property
If LdTSz_xMB52 > A Then Exit Property
If IsFstYM Then
'    If LdTSz_xIniMB52 > A Then Exit Property
'    If LdTSz_xIniRate > A Then Exit Property
    IsLd_xRate = True
    Exit Property
End If
'If LdDTim_xLasMB52 > A Then Exit Property
'If LdDTim_xLasRate > A Then Exit Property
IsLd_xRate = True
End Property
Property Get IsLd_xMB521() As Boolean
IsLd_xMB521 = LdTSz_xMB52 = FfnTSz(IFxMB52)
End Property
Property Get IsLd_xInv() As Boolean
IsLd_xInv = LdTSz_xInv = FfnTSz(IFxInv)
End Property