Option Compare Database
Option Explicit
Function MB52FnSpec$()
MB52FnSpec = "MB52 " & YYYYxMM & "-??.xls"
End Function
Function Load() As Boolean
'Return True if error
Dim Fx$
Fx = IFx
Stop
If Fx = "" Then
    MsgAp_Brw "No MB52 file [like-this] is found in [import folder]", MB52FnSpec, IPth
    GoTo Er
End If
If IFxIsAlreadyLoaded(Fx) Then
    Exit Function
End If
If FxLoad(Fx) Then GoTo Er
Exit Function
Er:
    Load = True
End Function


Function IFxFnAy_WhYM(A$()) As String()
IFxFnAy_WhYM = AyWhLik(A, MB52FnSpec)
End Function

Function IFxAy() As String()
Dim P$
P = IPth
IFxAy = AyAddPfx(IFxFnAy_WhYM(PthFnAy(P, MB52FnSpec)), P)
End Function

Function IFxIsAlreadyLoaded(A$) As Boolean
Dim Tim As Date, Sz&
With QQSqlRs("Select BegOH_FxSz, BegOH_FxTim from YM where BegOH_Fx='?'", A)
    If .EOF Then Exit Function
    Sz = .Fields(0).Value
    Tim = .Fields(1).Value
End With
If FfnTim(A) <> Tim Then Exit Function
If FfnSz(A) <> Sz Then Exit Function
IFxIsAlreadyLoaded = True
End Function

Sub ZZ_FxLoad()
Y = 18
M = 7
Debug.Print FxLoad(IFx)
End Sub
Function FxLoad(A) As Boolean
'return true if er
WIni
WtLnkFx ">MB52", A
WtLnkFx ">Uom", IFxUOM
WtChkCol ">MB52", LnkColStr.MB52
WtChkCol ">Uom", LnkColStr.Uom
WttLnkFb "YM YMOH", IFbStkShpRate
WImp ">MB52", LnkColStr.MB52
WImp ">Uom", LnkColStr.Uom
Fx_Upd_YM_and_YMOH A
Exit Function
X:
    FxLoad = True
End Function

Sub Fx_Upd_YM_and_YMOH(A)
'#IMB52 is imported
'Import into YMOH & Update YM
WDrp "#OH"
WRun "Select Distinct Sku,Whs,Sum(x.QUnRes+x.QInsp+x.QBlk) as OH into [#OH] from [#IMB52] x group by Sku,Whs"
WRun "Alter Table [#OH] add column Sc_U double, Sc double"
WRun "Update [#OH] x inner join [#IUom] a on x.Sku=a.Sku and x.Whs=a.Whs set x.Sc_U=a.Sc_U"
WRun "Update [#OH] set Sc = OH / Sc_U where Sc_U is not null and Sc_U<>0"
'
WQQ "Delete from [YMOH] where Y=? and M=?", Y, M
WQQ "Insert into [YMOH] (Y,M,Sku,Whs,OH,Sc_U,Sc) select ?,?,Sku,Whs,OH,Sc_U,Sc from [#OH]", Y, M

'Update YM: Y M *Fx *FxTim *FxSz *NRec *NSku *Sc *DteLoad
Dim Tim As Date
Dim Sz&
Dim NRec&
Dim NSku%
Dim OH&, Sc#
    Tim = FfnTim(A)
    Sz = FfnSz(A)
    NRec = DbqV(W, "Select Count(*) from [#IMB52]")
    Sc = DbqV(W, "Select Sum(Sc) from [#OH]")
    OH = DbqV(W, "Select Sum(OH) from [#OH]")
    NSku = DbqV(W, "Select Count(*) from (Select Distinct Sku From [#OH])")
WQQ "Update [YM]" & _
" set" & _
" BegOH_Fx='?'," & _
" BegOH_FxTim=#?#," & _
" BegOH_FxSz=?," & _
" BegOH_NRec=?," & _
" BegOH_NSku=?," & _
" BegOH_Sc=?," & _
" BegOH_LoadDte=#?#" & _
" where Y=? and M=?", _
A, Tim, Sz, NRec, NSku, Round(Sc, 1), NowStr, Y, M
WDrp "#OH"
PushSts "[MB52] of [Size] and [time] with [n-records], [n-Sku], [total-Sc] and [total-amt] are loaded in [year] and [month]", A, Sz, Tim, NRec, NSku, Sc, Y, M
End Sub