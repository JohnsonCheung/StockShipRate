Option Compare Database
Option Explicit
Sub SpnmImp(A)
DbImpSpec CurrentDb, A
End Sub

Function SpnmFt$(A)
SpnmFt = PgmObjPth & SpnmFn(A)
End Function

Sub DbImpSpec(A As Database, Spnm)
Const CSub$ = "DbImpSpec"
Dim T As Date, FtSz&
Dim K$, Ft$, Sk$()
If False Then
    'Assume Db has T = Spec * SpecNm Lines Ft FtTim FtSz. It has Sk of field = SpecNm
    'Also, assume T has Ft, XXX_Tim, XXX_Sz, XXX_LdDte field, where XXX is
    Sk = DbtSk(A, "Spec")
    If Sz(Sk) <> 1 Then Stop
    With A.TableDefs("Spec").OpenRecordset
        .Index = "Spec"
        .Seek "=", K
        If .NoMatch Then
            .AddNew
            .Fields(Sk(0)).Value = K
            .Update
            .Seek "=", K
        End If
        .Edit
        !Lines = FtLines(Ft)
        !Ft = Ft
        !FtTim = FfnTim(Ft)
        !FtSz = FfnSz(Ft)
        !LdDte = Now
        .Update
    End With
End If

If False Then
    Ft = SpnmFt(A)
    If Not FfnIsExist(A) Then
        FunMsgDmp "LgSchm_Imp", "[LgSchmFt] not exist, no LgSchm_Imp.", A
        Exit Sub
    End If
    FfnAsgTSz A, T, FtSz
    If SpnmFt_IsNew(A, Ft) Then
        'SpnmImpFt "LgSchmp", A
        FunMsgDmp CSub, "[LgSchmFt] of [time] and [size] is imported", A, T, FtSz
    Else
        FunMsgDmp "LgSchm_Imp", "[LgSchmFt-Tim] is same or older than [Imported-Schm-Tim], no import.  They have [Sz1] and [Sz2]", _
            T, LgSchm_Tim, FtSz, LgSchm_Sz
    End If
End If

Q = FmtQQ("Select * from Spec where SpecNm = '?'", Spnm)
Dim Rs As DAO.Recordset
Set Rs = A.OpenRecordset(Q)
With Rs
    If .EOF Then
        .AddNew
        !SpecNm = A
        .Update
        .Requery
    End If
End With

Dim CurOld As Boolean, CurNew As Boolean, SamTim As Boolean 'compare the Rs's tim
    Dim CurT As Date, LasT As Date 'CurTim and LasTim
    CurT = FfnTim(Ft)
    LasT = Rs!Tim
    SamTim = CurT = LasT
    CurOld = CurT < LasT
    CurNew = CurT > LasT
Dim SamSz As Boolean, DifSz As Boolean
    Dim CurS&, LasS&
    CurS = FfnSz(Ft)
    LasS = Rs!Sz
    SamSz = CurS = LasS
    DifSz = Not SamSz
    
Select Case True
Case SamTim And SamSz
    FunMsgDmp CSub, "[Ft] of [SpecNm] with [Tim] & [Sz] is same as [last-time] import.  No import.", _
        Ft, A, DteDTim(CurT), CurS, SpnmLdDTim(A)
Case SamTim And DifSz
    FunMsgDmp CSub, "[Ft] of [SpecNm] with same [Tim] & [Sz] is same as [last-time] import.  No import.  They have [Sz1] and [Sz2]", _
        Ft, LgSchm_Tim, LasS, CurS
Case CurOld
'    FunMsgDmp CSub, "[Ft] of [SpecNm] with same [Tim] & [Sz] is same as [last-time] import.  No import.  They have [Sz1] and [Sz2]", _
'        Ft, LgSchmFtTim, LasS, CurS
Case CurNew
    With Rs
        .Edit
        !Ft = FtLines(Ft)
        !Sz = CurS
        !Tim = CurT
        !LdDte = Now
        .Update
    End With
'    FunMsgDmp CSub, "**** IMPORTED ****|[SpecNm] [file] with [time] & [size] is newer than [last-time] with [last-time-Ft-time] and [last-time-Ft-size].", _
'        Ft, LgSchmFtTim, LasS, CurS
Case Else: Stop
End Select
End Sub

Function SpnmLdDTim$(Spnm)
End Function

Function SpnmLines$(A)
SpnmLines = TfkV("Spec", "Lines", A)
End Function

Function SpnmDTim$(A)
SpnmDTim = DteDTim(SpnmTim(A))
End Function

Function SpnmFt_IsNew(A, Ft) As Boolean
SpnmFt_IsNew = FfnTim(A) > SpnmTim(A)
End Function

Sub SpnmExp(A, Optional OvrWrt As Boolean)
StrWrt SpnmLines(A), SpnmFt(A), Not OvrWrt
End Sub

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
SpnmFn = A & ".txt"
End Property

Property Get SpnmFny() As String()
SpnmFny = DbtFny(CurrentDb, "Spec")
End Property


Function SpnmTim(A) As Date
SpnmTim = Nz(TfkV("Spec", "Tim", A), 0)
End Function