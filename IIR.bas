Option Compare Database
Function Load() As Boolean
'Return true if Er
AssYM
Dim Fx$, FxDone$, LDte
Fx = IFx
FxDone = IFxDone
If Not FfnIsExist(Fx) Then
    LDte = LoadDte
    If IsDte(LDte) Then
        PushSts "Invoice is already loaded [At]", LDte
        Exit Function
    End If
    AyBrw MsgAp_Ly("[Invoices File] not found in [import Folder]", FfnFn(Fx), FfnPth(Fx))
    GoTo Er
End If
If FfnIsExist(FxDone) Then
    AyBrw MsgAp_Ly("[Invoices File] is found in both [import filer] and [done folder].|Please remove one of them", FfnFn(FxDone), IPthDone, FfnPth(FxDone))
    GoTo Er
End If
If FxLoad(Fx) Then GoTo Er
Exit Function
Er:
    Load = True
End Function
Function LoadDte()
LoadDte = QQSqlV("Select IR_LoadDte from YM where Y=? and M=?", Y, M)
End Function
Function IFxDone$()
IFxDone = IPthDone & IFxFn
End Function
Function IFxFn$()
IFxFn = FmtQQ("Invoices ?.xlsx", YYYYxMM)
End Function
Function IFx$()
IFx = IPth & IFxFn
End Function
Function IPth$()
Dim O$
O = PnmVal("InvPth") & YYYY & "\"
PthEns O
IPth = O
End Function
Function PthFnIr(A, Optional Spec$ = "*.*") As VBA.Collection
Dim O As New Collection
Dim B$, P$
P = PthEnsSfx(A)
B = Dir(P & Spec)
Dim J%
While B <> ""
    J = J + 1
    If J > 10000 Then Stop
    O.Add B
    B = Dir
Wend
Set PthFnIr = O
End Function
Function PthUp$(A)
PthUp = TakBefOrAllRev(RmvSfx(A, "\"), "\")
End Function

Sub PthMovFilUp(A)
Dim I
Tar$ = PthUp(A)
For Each I In PthFnIr(A)
    FfnMov I, Tar
Next
End Sub
Function IPthDone$()
Dim O$
O = IPth & "Done\"
PthEns O
IPthDone = O
End Function

Sub ZZ_FxIsInvFxChk()
Y = 18
M = 2
Dim A$()
A = IFxAy
D FxIsInvFxChk(A(0))
End Sub

Function FxIsInvFxChk(A) As String()
Dim O$()
O = FfnNotFndChk(A)
If Sz(O) > 0 Then
    FxIsInvFxChk = O
    Exit Function
End If
Dim WsNy$()
WsNy = FxWsNy(A)
If AyHasAy(WsNy, SslSy("Invoices Detail")) Then Exit Function
FxIsInvFxChk = MsgAp_Ly("[Excel file] does not have worksheet 'Invoices' and 'Detail'.  It has [these worksheets].", A, WsNy)
End Function

Sub IPthBrw()
PthBrw IPth
End Sub

Sub Z()
ZZ_FxLoad
End Sub
Sub ZZ_FxLoad()
Y = 18
M = 2
PthMovFilUp IPthDone
If FxLoad(IFx) Then Stop
BrwSts
End Sub

Function FxLoad(A) As Boolean
Dim O$(), B$(), C$()
If AyBrwEr(FxIsInvFxChk(A)) Then GoTo X
WIni
WtLnkFx ">InvH", A, "Invoices"
WtLnkFx ">InvD", A, "Detail"

B = WtChkCol(">InvD", LnkColStr.InvD)
C = WtChkCol(">InvH", LnkColStr.InvH)
If AyBrwEr(AyAdd(C, B)) Then GoTo X

WttLnkFb "YM InvH InvD", IFbStkShpRate

WImp ">InvH", LnkColStr.InvH
WImp ">InvD", LnkColStr.InvD

FxUpd_zInvH_and_InvD_and_YM A
FxMov A
Exit Function
X:
    FxLoad = True
End Function
Sub RestoreTstInvFile()

End Sub
Sub FxMov(A)
Dim P$
P = IPth & "Done\":  PthEns P
P = P & TmpNm & "\": PthEns P
FfnMov A, P
End Sub
Sub FxUpd_zInvH_and_InvD_and_YM(A)
'#IInvH & #IInvD are imported
'Replace InvH and InvD after validation
'
'#IInvD: VndShtNm InvNo Sku Sc Amt
'#IInvH: VndShtNm InvNo Dte Whs Sc Amt
'InvD: VndShtNm InvNo Sku Sc Amt
'InvH: VndShtNm InvNo Whs Dte Sc Amt DteCrt

WQQ "Delete x.* from [InvD] x inner join [InvH] a on a.VndShtNm=x.VndShtNm and a.InvNo=x.InvNo where Year(Dte)=? and Month(Dte)=?", Y, M
WQQ "Delete * from [InvH] where Year(Dte)=? and Month(Dte)=?", Y, M

WRun "insert into [InvD] (VndShtNm,InvNo,Sku,Sc,Amt)" & _
                 " select VndShtNm,InvNo,Sku,Sc,Amt from [#IInvD]'"
WRun "insert into [InvH] (VndShtNm,InvNo,Whs,Dte,Sc,Amt)" & _
                 " select VndShtNm,InvNo,Whs,Dte,Sc,Amt from [#IInvH]'"

Dim NInv%
Dim NSku
Dim NInvLin
Dim Amt@
Dim Sc#
With WqRs("Select Count(*), Sum(Amt), Sum(Sc) from [#IInvH]")
    NInv = .Fields(0).Value
    Amt = .Fields(1).Value
    Sc = .Fields(2).Value
    .Close
End With
NSku = WqV("Select Count(*) from (Select Distinct Sku from [#IInvD])")
NInvLin = WqV("Select Count(*) from [#IInvD]")

With WqRs(FmtQQ("Select IR_LoadDte, IR_Sc, IR_Amt, IR_NInv, IR_NSku, IR_NInvLin from YM where Y=? and M=?", Y, M))
    .Edit
    !IR_LoadDte = Now
    !IR_Sc = Sc
    !IR_Amt = Amt
    !IR_NInv = NInv
    !IR_NInvLin = NInvLin
    .Update
    .Close
End With
PushSts "[Invoice file] with [n-invoices], [n-lines], [total-Sc] and [total-amt] are loaded in [year] and [month]", A, NInv, NInvLin, Sc, Amt, Y + 2000, M
End Sub