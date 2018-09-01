Option Compare Database
Option Explicit
Private X_Spnm$

Sub SpnmImp(A)
DbImpSpec CurrentDb, A
End Sub

Function SpnmFt$(A)
SpnmFt = PgmObjPth & SpnmFn(A)
End Function

Sub DbImpSpec(A As Database, Spnm)
X_Spnm = Spnm
Const CSub$ = "DbImpSpec"
Dim Ft$
    Ft = SpnmFt(Spnm)
    
Dim CurOld As Boolean
Dim CurNew As Boolean
Dim SamTim As Boolean
Dim DifSz As Boolean
Dim SamSz As Boolean
Dim DifFt As Boolean
Dim Rs As dao.Recordset
    Q = FmtQQ("Select Ft,Lines,Tim,Sz,LdTim from Spec where SpecNm = '?'", Spnm)
    Set Rs = CurrentDb.OpenRecordset(Q)
    
    Dim CurT As Date, LasT As Date 'CurTim and LasTim
    Dim CurS&, LasS&
    Dim LasFt$, LdDTim$
    CurS = FfnSz(Ft)
    CurT = FfnTim(Ft)
    With Rs
        LasS = Nz(Rs!Sz, -1)
        LasT = Nz(!Tim, 0)
        LasFt = Nz(!Ft, "")
        LdDTim = DteDTim(!LdTim)
    End With
    SamTim = CurT = LasT
    CurOld = CurT < LasT
    CurNew = CurT > LasT
    SamSz = CurS = LasS
    DifSz = Not SamSz
    DifFt = Ft <> LasFt
    

Const Imported$ = "***** IMPORTED *****"
Const NoImport$ = "----- no import -----"
Const FtDif______$ = "Ft is dif."
Const SamTimSz___$ = "Sam tim & sz."
Const SamTimDifSz$ = "Sam tim & sz. (Odd!)"
Const CurIsOld___$ = "Cur is old."
Const CurIsNew___$ = "Cur is new."
Const C$ = "|[SpecNm] [Db] [Cur-Ft] [Las-Ft] [Cur-Tim] [Las-Tim] [Cur-Sz] [Las-Sz] [Imported-Time]."

Select Case True
Case DifFt:  RsUpd Rs, Ft, FtLines(Ft), CurT, CurS, Now
Case CurNew: RsUpd Rs, Ft, FtLines(Ft), CurT, CurS, Now
End Select

Dim Av(): Av = Array(Spnm, DbNm(A), Ft, LasFt, CurT, LasT, CurS, LasS, LdDTim)
Select Case True
Case DifFt:            FunMsgDmpAv CSub, Imported & FtDif______ & C, Av
Case SamTim And SamSz: FunMsgDmpAv CSub, NoImport & SamTimSz___ & C, Av
Case SamTim And DifSz: FunMsgDmpAv CSub, NoImport & SamTimDifSz & C, Av
Case CurOld:           FunMsgDmpAv CSub, NoImport & CurIsOld___ & C, Av
Case CurNew:           FunMsgDmpAv CSub, Imported & CurIsNew___ & C, Av
Case Else: Stop
End Select
End Sub

Sub SpnmExp(A, Optional OvrWrt As Boolean)
StrWrt SpnmLines(A), SpnmFt(A), Not OvrWrt
End Sub
Function SpnmLines$(A)
SpnmLines = SpnmV(A, "Lines")
End Function
Function SpnmLy(A) As String()
SpnmLy = SplitCrLf(SpnmLines(A))
End Function

Function SpnmV(A, ValNm$)
SpnmV = TfkV("Spec", ValNm, A)
End Function

Sub SpnmBrw(A)
FtBrw SpnmFt(A)
End Sub

Sub SpnmIni(A)
Dim Ft$: Ft = SpnmFt(A)
If FfnIsExist(Ft) Then Debug.Print "SpecNm-[" & A & "] of Ft-[" & Ft & "] existed.": Exit Sub
StrWrt "", Ft
Debug.Print "SpecNm-[" & A & "] of Ft-[" & Ft & "] is created."
End Sub
Property Get SpnmFn$(A)
SpnmFn = A & "(Spec).txt"
End Property

Property Get SpnmFny() As String()
SpnmFny = DbtFny(CurrentDb, "Spec")
End Property


Function SpnmTim(A) As Date
SpnmTim = Nz(TfkV("Spec", "Tim", A), 0)
End Function