Option Compare Database
Option Explicit
Property Get IsFstYM() As Boolean
IsFstYM = FstY = Y And FstM = M
End Property
Property Get IsSndYM() As Boolean
IsSndYM = SndY = Y And SndM = M
End Property
Property Get SndY() As Byte
SndY = YM_YofNxtM(FstY, FstM)
End Property
Property Get SndM() As Byte
SndM = M_NxtM(FstM)
End Property
Property Get FstY() As Byte
FstY = SqlV("Select Min(Y) from YM")
End Property
Property Get FstM() As Byte
FstM = QQV("Select Min(M) from YM where Y=?", FstY)
End Property
Property Get Y() As Byte
Y = SqlV("Select Y from CurYM")
End Property
Property Get M() As Byte
M = SqlV("Select M from CurYM")
End Property
Property Let M(V As Byte)
RsF(TblRs("CurYM"), "M") = V
End Property
Property Let Y(V As Byte)
RsF(TblRs("CurYM"), "Y") = V
End Property

Function FmDte() As Date
FmDte = DateSerial(Y, M, 1)
End Function
Function ToDte() As Date
ToDte = DteLasDayOfMth(FmDte)
End Function
Function FmYYYYxMMxDD$()
FmYYYYxMMxDD = Format(FmDte, "YYYY-MM-DD")
End Function
Function ToYYYYxMMxDD$()
ToYYYYxMMxDD = Format(ToDte, "YYYY-MM-DD")
End Function
Function BegY() As Byte
BegY = IIf(M = 1, Y - 1, Y)
End Function
Function BegM() As Byte
BegM = IIf(M = 1, 12, M - 1)
End Function
Function NxtY() As Byte
NxtY = IIf(M = 12, Y + 1, Y)
End Function

Function NxtM() As Byte
NxtM = IIf(M = 12, 1, M + 1)
End Function

Function YYYYxMM$()
YYYYxMM = YYYY & "-" & MM
End Function

Function YYYYxMMxLasDD$()
YYYYxMMxLasDD = Format(YM_LasDte(Y, M), "YYYY-MM-DD")
End Function

Function PrvYYYYxMM$()
Dim YYYY$, MM$, Y As Byte, M As Byte
M = M_PrvM(FstM)
Y = YM_YofPrvM(FstY, FstM)
PrvYYYYxMM = Y + 2000 & "-" & Format(M, "00")
End Function

Function MM$()
MM = Format(M, "00")
End Function

Function YYYY$()
YYYY = Format(2000 + Y)
End Function