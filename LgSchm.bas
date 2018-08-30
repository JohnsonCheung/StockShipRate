Option Compare Database
Option Explicit

Const LgSNm$ = "LgSchm" ' The LgSchm-Spnm

Property Get LgSchm_Lines$()
LgSchm_Lines = SpnmLines(LgSNm)
End Property

Sub LgSchm_Imp()
SpnmImp "LgSchm"
End Sub

Property Get LgSchm_Ft$()
LgSchm_Ft = SpnmFt(LgSNm)
End Property

Sub LgSchm_Brw()
SpnmBrw LgSNm
End Sub

Sub LgSchm_Ini()
SpnmIni LgSNm
End Sub