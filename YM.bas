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
RsV(TblRs("CurYM"), "M") = V
End Property
Property Let Y(V As Byte)
RsV(TblRs("CurYM"), "Y") = V
End Property

Property Get FmDte() As Date
FmDte = DateSerial(Y, M, 1)
End Property
Property Get ToDte() As Date
ToDte = DteLasDayOfMth(FmDte)
End Property
Property Get FmYYYYxMMxDD$()
FmYYYYxMMxDD = Format(FmDte, "YYYY-MM-DD")
End Property
Property Get ToYYYYxMMxDD$()
ToYYYYxMMxDD = Format(ToDte, "YYYY-MM-DD")
End Property
Property Get BegY() As Byte
BegY = IIf(M = 1, Y - 1, Y)
End Property
Property Get BegM() As Byte
BegM = IIf(M = 1, 12, M - 1)
End Property
Property Get NxtY() As Byte
NxtY = IIf(M = 12, Y + 1, Y)
End Property

Property Get NxtM() As Byte
NxtM = IIf(M = 12, 1, M + 1)
End Property

Property Get YYYYxMM$()
YYYYxMM = YYYY & "-" & MM
End Property

Property Get YYYYxMMxLasDD$()
YYYYxMMxLasDD = Format(YM_LasDte(Y, M), "YYYY-MM-DD")
End Property

Property Get PrvYYYYxMM$()
Dim YYYY$, MM$, Y As Byte, M As Byte
M = M_PrvM(FstM)
Y = YM_YofPrvM(FstY, FstM)
PrvYYYYxMM = Y + 2000 & "-" & Format(M, "00")
End Property

Property Get MM$()
MM = Format(M, "00")
End Property

Property Get YYYY$()
YYYY = Format(2000 + Y)
End Property