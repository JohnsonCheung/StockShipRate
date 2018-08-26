Option Compare Database
Option Explicit
Public LnkColStr As New LnkColStr
Function OSubSsl_Run(A)
Dim IQ, Q$
For Each IQ In CvNy(A)
    Q = IQ
    Select Case True
    Case HasPfx(Q, "O"): Q = "@" & Mid(IQ, 2)
    Case HasPfx(Q, "Tmp")
    Case Else: Stop
    End Select
    
    MsgRunQry Q
    Run IQ
Next
End Function
Sub MsgRunQry(A$)
MsgSet "Running query (" & A & ") ..."
End Sub
Function ChkFil() As Boolean
Dim O$()
If IsFstYM Then
    If Not FfnIsExist(IFxRate) Then
        PushAy O, MsgAp_Ly("It is first record of Year/Month, [rate file] (ZHT1) is needed.|It is not found in [folder]", FfnFn(IFxRate), FfnPth(IFxRate))
    End If
End If
If IFxMB52 = "" Then
    PushAy O, MsgAp_Ly("No [MB52] in [folder].", MB52FnSpec, MB52Pth)
Else
    If Not FxHasSheet1(IFxMB52) Then
        PushAy O, MsgAp_Ly("[MB52] in [folder] does not have Sheet1, but [these].", FfnFn(IFxMB52), MB52Pth, FxWsNy(IFxMB52))
        
    End If
End If

If FfnIsExist(IFxInv) Then
    PushAy O, InvChk(IFxInv)
Else
    PushAy O, MsgAp_Ly("[Invoices file] not found in [folder]", InvFn, InvPth)
End If
If Not FfnIsExist(IFxUOM) Then
    PushAy O, MsgAp_Ly("[Sales text file] not found in [folder]", FfnFn(IFxUOM), FfnPth(IFxUOM))
End If
ChkFil = AyBrwEr(O)
End Function
Sub Lnk()
WtLnkFx ">InvH", IFxInv, "Invoices"
WtLnkFx ">InvD", IFxInv, "Detail"
WtLnkFx ">MB52", IFxMB52
WtLnkFx ">Uom", IFxUOM
If IsFstYM Then
WtLnkFx ">ZHT18601", IFxRate, "8601"
WtLnkFx ">ZHT18701", IFxRate, "8701"
End If
If IsDev Then
    WttLnkFb "IniYM IniRate IniOH YMInvH YMInvD YM YMOH", IFbStkShpRate, "^IniYM ^IniRate ^IniOH ^YMInvH ^YMInvD ^YM ^YMOH"
Else
    WttLnkFb "IniYM IniRate IniOH YMInvH YMInvD YM YMOH", IFbStkShpRate
End If
End Sub

Function ChkCol() As Boolean
Dim A$(), B$(), C$(), D$(), E$(), F$()
A = WtChkCol(">MB52", LnkColStr.MB52)
B = WtChkCol(">Uom", LnkColStr.Uom)
If IsFstYM Then
C = WtChkCol(">ZHT18601", LnkColStr.ZHT1)
D = WtChkCol(">ZHT18701", LnkColStr.ZHT1)
End If
E = WtChkCol(">InvD", LnkColStr.InvD)
F = WtChkCol(">InvH", LnkColStr.InvH)
ChkCol = AyBrwEr(AyAddAp(A, C, B, D, E, F))
End Function
Sub Import()
If IsFstYM Then
    WImp ">ZHT18601", LnkColStr.ZHT1
    WImp ">ZHT18701", LnkColStr.ZHT1
End If
WImp ">MB52", LnkColStr.MB52
WImp ">InvH", LnkColStr.InvH
WImp ">InvD", LnkColStr.InvD
WImp ">Uom", LnkColStr.Uom
End Sub
Function OupPth$()
Dim A$
A = CurDbPth & "Output\"
PthEns A
OupPth = A
End Function
Function IFbStkShpRate$()
If IsDev Then
    IFbStkShpRate = CurrentDb.Name
Else
    IFbStkShpRate = "N:\SAPAcessReports\StockShipRate\StockShipRate_Data.accdb"
End If
End Function
Function OupFx$()
Dim O$
O = OupPth & FmtQQ("? ?.xlsx", Apn, YYYYxMM)

End Function
Private Sub MsgSet(A$)
Form_Main.MsgSet A
End Sub
Private Sub MsgClr()
Form_Main.MsgClr
End Sub

Function IFxUOM$()
IFxUOM = PnmFfn("UOM")
End Function

Sub DocUOM()
'InpX: [>UOM]     Material [Base Unit of Measure] [Material Description] [Unit per case]
'Oup : UOM        Sku      SkuUOM                 Des                    Sc_U

'Note on [Sales text.xls]
'Col  Xls Title            FldName     Means
'F    Base Unit of Measure SkuUOM      either COL (bottle) or PCE (set)
'J    Unit per case        Sc_U        how many unit per AC
'K    SC                   SC_U        how many unit per SC   ('no need)
'L    COL per case         AC_B        how many bottle per AC
'-----
'Letter meaning
'B = Bottle
'AC = act case
'SC = standard case
'U = Unit  (Bottle(COL) or Set (PCE))

' "SC              as SC_U," & _  no need
' "[COL per case]  as AC_B," & _ no need
End Sub
Function NxtYMStr$()
Dim Y As Byte, M As Byte
Y = SqlV("Select Max(Y) from YM")
M = QQV("Select Max(M) from YM where Y=?", Y)
If M = 12 Then
    M = 1
    Y = Y + 1
Else
    M = M + 1
End If
NxtYMStr = Y & "." & M
End Function
Sub TblYM_Rfh()
Dim Y As Byte, M As Byte
BrkAsg NxtYMStr, ".", Y, M
Dim J%, I%
For J = Y To CurY - 1
    For I = 1 To 12
        If Not SqlAny(FmtQQ("Select Y from [YM] where Y=? and M=?", J, I)) Then
            DoCmd.RunSQL FmtQQ("Insert into [YM] (Y,M) values (?,?)", J, I)
        End If
    Next
Next
For I = 1 To CurM
    If Not SqlAny(FmtQQ("Select Y from [YM] where Y=? and M=?", CurY, I)) Then
        DoCmd.RunSQL FmtQQ("Insert into [YM] (Y,M) values (?,?)", CurY, I)
    End If
Next
End Sub
Sub TblYM_Dlt_BelowIniYM()
QQRun "Delete * from YM where Y<?", IniY
QQRun "Delete * from YM where Y=? and M<?", IniY, IniM
End Sub
Sub TblYM_Ins_UpToCurYM()
Dim Y As Byte, M As Byte, FmM As Byte, ToM As Byte
For Y = IniY To CurY
    FmM = IIf(IniY = Y, IniM, 1)
    ToM = IIf(CurY = Y, CurM, 12)
    For M = FmM To ToM
        If Not QQAny("Select Y from YM where Y=? and M=?", Y, M) Then
            QQRun "Insert into YM (Y,M) values(?,?)", Y, M
        End If
    Next
Next
End Sub

Sub ZZ_YM_Ini()
Y = 18
M = 1
'TblYM_Ini
End Sub

Function InvLoadDte()
InvLoadDte = QQSqlV("Select IR_LoadDte from YM where Y=? and M=?", Y, M)
End Function

Function InvFn$()
InvFn = FmtQQ("Invoices ?.xlsx", YYYYxMM)
End Function
Function IFxInv$()
IFxInv = InvPth & InvFn
End Function
Function InvHom$()
If IsDev Then
    InvHom = CurDbPth & "Sample\"
Else
    InvHom = PthEns(AppHom & "Import Invoices\")
End If
End Function
Function InvPth$()
InvPth = PthEns(InvHom & YYYY & "\")
End Function

Sub ZZ_InvChk()
Y = 18
M = 2
D InvChk(IFxInv)
End Sub

Function InvChk(A) As String()
Dim O$()
Dim WsNy$()
WsNy = FxWsNy(A)
If AyHasAy(WsNy, SslSy("Invoices Detail")) Then Exit Function
InvChk = MsgAp_Ly("[Invoices file] in [folder] does not have worksheet 'Invoices' and 'Detail', but [these].", FfnFn(A), FfnPth(A), WsNy)
End Function

Sub InvPthBrw()
PthBrw InvPth
End Sub
Sub LoadRate()
End Sub

Sub LoadInv(Optional IsForceLoad As Boolean)
'#IInvH & #IInvD are imported
'Replace InvH and InvD after validation
'
'#IInvD: VndShtNm InvNo Sku Sc Amt
'#IInvH: VndShtNm InvNo Dte Whs Sc Amt
'InvD: VndShtNm InvNo Sku Sc Amt
'InvH: VndShtNm InvNo Whs Dte Sc Amt DteCrt
Dim A$
A = IFxInv
If Not IsForceLoad Then
    If IsLd_xInv Then Exit Sub
End If
WQQ "Delete x.* from [YMInvD] where InvH in (Select InvH from InvH where Year(Dte)=? and Month(Dte)=?)", Y, M
WQQ "Delete * from [YMInvH] where Year(Dte)=? and Month(Dte)=?", Y, M

WRun "insert into [YMInvH] (VndShtNm,InvNo,Whs,Dte,Sc,Amt)" & _
                   " select VndShtNm,InvNo,Whs,Dte,Sc,Amt from [#IInvH]'"
WRun "Alter Table [#IInvD] add column InvH Long"
WRun "Update [#IInvD] x inner join [YMInvH] a on x.VndShtNm=a.VndShtNm and x.InvNo=a.InvNo set x.InvH=a.InvH"
WRun "insert into [YMInvD] (InvH,Sku,Sc,Amt)" & _
                   " select InvH,Sku,Sc,Amt from [#IInvD]'"
Dim Q$
    Q = FmtQQ("Select IR_Fx, IR_FxSz, IR_FxTim, IR_Sc, IR_Amt, IR_NInv, IR_NInvLin, IR_NSku, IR_LoadDte from YM where Y=? and M=?", Y, M)

RsUpdDr SqlRs(Q), TmpInvHD_IRDr(A)
End Sub

Function IRDrLy(A()) As String()
Dim Fx$, Sz&, Tim As Date, Sc#, Amt@, NInv%, NInvLin%
AyAsg A, Fx, Sz, Tim, Sc, Amt, NInv, NInvLin
PushSts "[Invoice file] of [time] and [size] with [n-invoices], [n-lines], [total-Sc] and [total-amt] are loaded in [year] and [month]", _
    Fx, Tim, FfnSz(A), NInv, NInvLin, Round(Sc, 1), Round(Amt, 2), Y + 2000, M

End Function

Function TmpInvHD_IRDr(Fx) As Variant()
Dim Sz&, Tim As Date, Sc#, Amt@, NInv%, NInvLin%, NSku%
With WQQRs("Select Count(*), Sum(Amt), Sum(Sc) from [#IInvH]")
    NInv = .Fields(0).Value
    Amt = .Fields(1).Value
    Sc = .Fields(2).Value
    .Close
End With
NSku = WQV("Select Count(*) from (Select Distinct Sku from [#IInvD])")
NInvLin = WQV("Select Count(*) from [#IInvD]")
TmpInvHD_IRDr = Array(Fx, Sz, Tim, Sc, Amt, NInv, NInvLin, NSku, Now)
End Function
Property Get MB52FnSpec$()
MB52FnSpec = "MB52 " & YYYYxMM & "-??.xls"
End Property
Property Get IniMB52FnSpec$()
IniMB52FnSpec = "MB52 " & IniPrvYYYYxMM & "-??.xls"
End Property

Sub LoadMB52(Optional IsForceLoad As Boolean)
If Not IsForceLoad Then
    If IsLd_xMB52 Then Exit Sub
End If
'#IMB52 is imported into YMTbl and OHTbl
'Import into YMOH & Update YM
WDrp "#OH"
WRun "Select Distinct Sku,Whs,Sum(x.QUnRes+x.QInsp+x.QBlk) as OH into [#OH] from [#IMB52] x group by Sku,Whs"
'
WQQ "Delete from [YMOH] where Y=? and M=?", Y, M
WQQ "Insert into [YMOH] (Y,M,Sku,Whs,OH) select ?,?,Sku,Whs,OH from [#OH]", Y, M

'Update YM: Y M *Fx *FxTim *FxSz *NRec *NSku *Sc *DteLoad
Dim NRec&
Dim NSku%
Dim Sz&
Dim Tim As Date
Dim OH&, Sc#
Dim A$
    A = IFxMB52
    Tim = FfnTim(A)
    Sz = FfnSz(A)
    NRec = WQQV("Select Count(*) from [#IMB52]")
    Sc = DbqV(W, "Select Sum(Sc) from [#OH]")
    OH = DbqV(W, "Select Sum(OH) from [#OH]")
    NSku = DbqV(W, "Select Count(*) from (Select Distinct Sku From [#OH])")

With WQQRs("Select EndOH_Fx, EndOH_FxSz, EndOH_FxTim, EndOH_NRec, EndOH_LoadDte, EndOH_Sc, EndOH_Amt, EndOH_NSku from YM where Y=? and M=?", Y, M)
    .Edit
    !EndOH_Fx = A
    !EndOH_FxSz = FfnSz(A)
    !EndOH_FxTim = FfnTim(A)
    !EndOH_LoadDte = Now
    .Update
    .Close
End With
WDrp "#OH"
PushSts "[MB52] of [time] and [size] with [n-records], [n-Sku], [total-Sc] are loaded in [year] and [month]", _
    A, Tim, Sz, NRec, NSku, Round(Sc, 2), Y + 2000, M
End Sub
Function MB52yWhYM(A$()) As String()
MB52yWhYM = AyWhLik(A, MB52FnSpec)
End Function
Property Get IFxMB52$()
IFxMB52 = AyMax(MB52y)
End Property
Property Get IFxIniMB52$()
IFxIniMB52 = AyMax(IniMB52y)
End Property
Property Get IFxIniRate$()
IFxIniRate = PnmFfn("ZHT1")
End Property
Property Get MB52y() As String()
Dim P$
P = MB52Pth
MB52y = AyAddPfx(MB52yWhYM(PthFnAy(P, MB52FnSpec)), P)
End Property
Property Get IniMB52y() As String()
Dim P$
P = MB52Pth
IniMB52y = AyAddPfx(MB52yWhYM(PthFnAy(P, IniMB52FnSpec)), P)
End Property
Property Get MB52Pth$()
MB52Pth = PthEnsSfx(PnmVal("MB52Pth")) & 2000 + Y & "\"
End Property


Sub Rpt()
'The @Main is the detail of showing how NxtMth YMRate is calculate
If Not Cfm Then Exit Sub
WIni
If ChkFil Then Exit Sub
Lnk
If ChkCol Then Exit Sub
Dim IsForceLoad As Boolean
IsForceLoad = True
Import
LoadIniMB52 IsForceLoad
LoadIniRate IsForceLoad
LoadMB52 IsForceLoad
LoadInv IsForceLoad
LoadRate
Stop
Tmp
Oup
Upd
Gen
WCls
End Sub
Sub LoadIniMB52(Optional IsForceLoad As Boolean)
If Not IsForceLoad Then
    If IsLd_xIniMB52 Then Exit Sub
End If
If Not IsFstYM Then Exit Sub
Stop '
End Sub
Function Cfm() As Boolean
'Reset_YM           this should be put in front of starting the whole process
'Reset_YMRate
Cfm = True
End Function
Sub Gen()
OupFx_Gen OupFx, WFb
End Sub
Sub Oup()
OMain
End Sub
Sub Upd()
'YMRate: Y M ZHT1 Whs RateSc FmDte ToDte DteCrt
WQQ "Delete * from [YMRate] where Y=? and M=?", Y, M
'WQQ "Insert into YMRate (Y,M,ZHT1,Whs,RateSc) Select ?,?,ZHT1,Whs,RateSc from [$ZHT1]", Y, M
'TmpRate_Upd_YM "$ZHT1"
Stop
WDrp "$ZHT1"
End Sub

Sub Tmp()
TmpRate
End Sub
Sub TmpRate()
WDrp "$Rate"
WQQ "Select ZHT1,Whs,RateSc into [$Rate] from [YMRate] where Y=? and M=?", Y, M
WQQ "Create Index Pk on [$Rate] (ZHT1,Whs) with primary"
End Sub
Sub OMain()
'YMRate: Y M ZHT1 Whs RateSc FmDte ToDte DteCrt
'YMOH:   Y M Sku Whs OH Sc Sc_U
'YMInvD: VndShtNm InvNo Sku Sc Amt
'YMInvH: VndShtNm InvNo Whs Dte Sc Amt DteCrt
'#IUom   Sku Whs Des StkUom Sc_U ProdH
'Given: BegOHSc   =  100(Sc)
'       BegRateSc =  $0.5/Sc  => BegAmt = $50
'       IRSc      =  30(Sc)
'       IRAmt     =  $21      => IRRateSc = $0.7/Sc
'       EndOHSc   =  40(Sc)
'To Find: EndRateSc
'Work:
'      SellSc    = BegOHSc + IRSc - EndOHSc    = 100(Sc) + 30(Sc) - 40(Sc) = 90(Sc)
'      SellAmt   = SellSc * BegRateSc          = 90(Sc) * $0.5/Sc          = $45
'      EndAmt    = BegAmt + IRAmt - SellAmt    = $50 + $21 - $45           = $26
'      EndRateSc = EndAmt / EndOHSc            = $26 / 40(Sc)              = $0.65/Sc (**)
    If IsDte(InvLoadDte) Then
        PushSts "[Invoice] is already loaded [At]", IFxInv, InvLoadDte
        Exit Sub
    End If


WDrp "#BegOH #EndOH #IR #SkuWhs #SkuWhs1 @Main"

'#EndOH Sku Whs EndOH EndOHSc Sc_U
WQQ "Select Sku,Whs,OH as EndOH into [#EndOH] from [YMOH] where Y=? and M=?", Y, M
WRun "Alter Table [#EndOH] Add Column Sc_U Single, EndOHSc Single"
WRun "Update [#EndOH] x inner join [#IUom] a on x.Sku=a.Sku and x.Whs=a.Sku set x.Sc_U=a.Sc_U"
WRun "Update [#EndOH] set EndOHSc=EndOH/Sc_U where Sc_U is not null"

'#BegOH Sku Whs BegOH BegOHSc Sc_U
WQQ "Select Sku,Whs,OH as BegOH into [#BegOH] from [YMOH] where Y=? and M=?", BegY, BegM
WRun "Alter Table [#BegOH] Add Column Sc_U Single, BegOHSc Single"
WRun "Update [#BegOH] x inner join [#IUom] a on x.Sku=a.Sku and x.Whs=a.Whs set x.Sc_U=a.Sc_U"
WRun "Update [#BegOH] set BegOHSc=BegOH/Sc_U where Sc_U is not null"

'#IR Sku Whs IRSc IRAmt
WQQ "Select Distinct Sku,Whs,Sum(x.Sc) as IRSc, Sum(a.Amt) as IRAmt into [#IR]" & _
" from [YMInvD] x inner join [YMInvH] a on x.InvNo=a.InvNo and x.VndShtNm=a.VndShtNm" & _
" where a.Dte between #?# and #?#" & _
" group by Sku,Whs", _
FmYYYYxMMxDD, ToYYYYxMMxDD

'#SkuWhs
WRun "Select Sku,Whs into [#SkuWhs1] from [#BegOH]"
WRun "Insert into [#SkuWhs1] Select Sku,Whs from [#EndOH]"
WRun "Insert into [#SkuWhs1] Select Sku,Whs from [#IR]"
WRun "Select Distinct Sku,Whs into [#SkuWhs] from [#SkuWhs1]"

'@Main
WRun "Select x.Sku,x.Whs, BegOHSc, EndOHSc, IRSc,IRAmt" & _
" into [@Main]" & _
" from (([#SkuWhs] x" & _
" left join [#BegOH] a on x.Sku=a.Sku and x.Whs=a.Whs)" & _
" left join [#EndOH] b on x.Sku=b.Sku and x.Whs=b.Whs)" & _
" left join [#IR]    c on x.Sku=c.Sku and x.Whs=c.Whs"

'AddCol ProdH M32 M35 M38
WRun "Alter Table [@Main] add column Sc_U double, ProdH text(20),M32 Text(2), M35 Text(5), M38 Text(8),ZHT1 Text(8), RateSc Double"
WRun "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku and x.Whs=a.Whs set x.ProdH=a.ProdH, x.Sc_U=a.Sc_U"
WRun "Update [@Main] set M32=Mid(ProdH,3,2),M35=Mid(ProdH,3,5),M38=Mid(ProdH,3,8)"

'AddCol ZHT1 RateSc
WRun "Update [@Main] x inner join [$Rate] a on x.Whs=a.Whs and x.M38=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"
WRun "Update [@Main] x inner join [$Rate] a on x.Whs=a.Whs and x.M35=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"
WRun "Update [@Main] x inner join [$Rate] a on x.Whs=a.Whs and x.M32=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"

'RenCol RateSc -> BegRateSc
WReOpn
WtRenCol "@Main", "RateSc", "BegRateSc"

'@Main Sku ZHT1 BegOHSc BegRateSc BegAmt IRSc IRAmt OldRateSc EndAmt EndOHSc EndRateSc
             'Sku Whs
             'ZHT1 ZBrdNm ZBrd ZQlyNm ZQly Z8Nm Z8
             'EndAmt = EndOHSc
'AddCol Des StkUom
WRun "Alter Table [@Main] Add Column Des Text(255), StkUom Text(10)"
WRun "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku set x.Sc_U = a.Sc_U,x.Des=a.Des,x.StkUom=a.StkUom"

'AddCol BegAmt SellSc SellAmt EndAmt EndRateSc
WRun "Alter Table [@Main]" & _
" Add Column " & _
"BegAmt    Currency," & _
"SellSc    Double," & _
"SellAmt   Currency," & _
"EndAmt    Currency," & _
"EndRateSc Double"

Const LoFmlVbl1$ = _
" SellSc = [BegOHSc] +  [IRSc] - [EndOHSc] |" & _
" SellAmt = [SellSc] * [BegRateSc] |" & _
" BegAmt = [BegOHSc] * [BegRateSc] |" & _
" EndAmt = [BegAmt] + [IRAmt] - [SellAmt] |" & _
" EndRateSc = If([EndOHSc]=0,[BegRateSc],[EndAmt]/[EndOHSc])"

'Stream ProdH F2 M32 M35 M38 Topaz ZHT1 RateSc Z2 Z5 Z8
'ProdH Topaz
'Stream
'Z2 Z5 Z8
'Amt
'WRun "Alter Table [@Main] add column Stream Text(10), Topaz Text(20), ProdH Text(8), F2 Text(2), M32 text(2), M35 text(5), M38 Text(8), ZHT1 Text(8), Z2 text(2), Z5 text(5), Z8 Text(8), RateSc Currency, Amt Currency"
'WRun "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku set x.ProdH=a.ProdH,x.Topaz=a.Topaz"
'WRun "Update [@Main] set Stream=IIf(Left(Topaz,3)='UDV','Diageo','MH')"
'WRun "Update [@Main] Set Z2=Left(ZHT1,2), Z5=Left(ZHT1,5), Z8=Left(ZHT1,8) where not ZHT1 is null"
'WRun "Update [@Main] Set Amt = RateSc * OH_Sc where RateSc is not null"
WDrp "#SkuWhs #SkuWhs1 #BegOH #EndOH #IR"
End Sub

Sub Reset_YM()
'YM: Y M
'    BegOH_LoadDte BegOH_Sc BegOH_Amt BegOH_Fx BegOH_FxSz BegOH_FxTim BegOH_NSku BegOH_NRec
'    RateSc_NRec RateSc_Max RateSc_Min RateSc_Avg RateSc_LoadDte
'    IR_LoadDte IR_Sc IR_Amt IR_NInv IR_NSku IR_NInvLin
DoCmd.RunSQL FmtQQ("Update [YM] set RateSc_NRec=null, RateSc_Max=null, RateSc_Min=null, RateSc_Avg = Null,RateSc_LoadDte=null" & _
" where Y>? or (Y=? and M>?)", Y, Y, M)
End Sub
Sub Reset_YMRate()
DoCmd.RunSQL FmtQQ("Delete * from [YMRate] where Y>? or (Y=? and M>?)", Y, Y, M)
End Sub

Sub Tmp_AddRateColumns(A$)
'TmpTbl-A should have <Sku> and no X:<ProdH M32 M35 M38 ZHT1 RateSc Z2 Z5 Z8>
'TmpRate  should have <ZHT1 Whs RateSc>
'TmpTbl-A will have X:<> added
'#IUom: SKu Whs ProdH

'ProdH M32 M35 M38 ZHT1 RateSc Z2 Z5 Z8
WQQ "Alter Table [?] add column ProdH text(15), M32 text(2), M35 text(5), M38 Text(8), ZHT1 Text(8), Z2 Text(2), Z5 Text(5), Z8 Text(8), RateSc Currency", A

'ProdH
WQQ "Update [?] x inner join [#IUom] a on x.Sku=a.Sku and x.Whs=a.Whs set x.ProdH=a.ProdH", A

'M32 M35 M38
WQQ "Update [?] set M32=Mid(ProdH,3,2),M35=Mid(ProdH,3,5),M38=Mid(ProdH,3,8)", A

'ZHT1 RateSc
WQQ "Update [?] x inner join [$Rate] a on x.Whs=a.Whs and x.M38=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null", A
WQQ "Update [?] x inner join [$Rate] a on x.Whs=a.Whs and x.M35=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null", A
WQQ "Update [?] x inner join [$Rate] a on x.Whs=a.Whs and x.M32=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null", A

'Z2 Z5 Z8
WQQ "Update [?] Set Z2=Left(ZHT1,2), Z5=Left(ZHT1,5), Z8=Left(ZHT1,8) where not ZHT1 is null", A
End Sub

Sub LoadIniRate(Optional IsForceLoad As Boolean)
If Not IsFstYM Then Exit Sub
If Not IsForceLoad Then
    If IsLd_xIniRate Then
        Dim Kind$
        AyDmp FfnAlreadyLoadedMsgLy(IFxIniRate, Kind, LdDTim_xIniRate)
        Exit Sub
    End If
End If
WDrp "#Cpy1 #Cpy2 #Cpy"
WRun "Select '8701' as Whs,x.* into [#Cpy1] from [#IZHT18701] x"
WRun "Select '8601' as Whs,x.* into [#Cpy2] from [#IZHT18601] x"

WRun "Select * into [#Cpy] from [#Cpy1] where False"
WRun "Insert into [#Cpy] select * from [#Cpy1]"
WRun "Insert into [#Cpy] select * from [#Cpy2]"

WRun "Alter Table [#Cpy] Add Column FmDte Date,ToDte Date"
WRun "Update [#Cpy] Set" & _
" FmDte = DateSerial(RIGHT(VdtFm,4),MID(VdtFm,4,2),LEFT(VdtFm,2))," & _
" ToDte = DateSerial(RIGHT(VdtTo,4),MID(VdtTo,4,2),LEFT(VdtTo,2))"
WQQ "Delete * from [#Cpy]" & _
" Where not #?# between FmDte and ToDte", YYYYxMM & "-01"

WRun "Delete * from [IniRate]"
WRun "Insert into [IniRate] (ZHT1,Whs,RateSc,FmDte,ToDte) select ZHT1,Whs,RateSc,FmDte,ToDte from [#Cpy]"

Dim A$
A = IFxIniRate
With W.OpenRecordset("Select IniRate_Fx, IniRate_FxSz, IniRate_FxTim, IniRate_LoadDte from [IniYM]")
    .Edit
    !IniRate_Fx = A
    !IniRate_FxSz = FfnSz(A)
    !IniRate_FxTim = FfnTim(A)
    !IniRate_LoadDte = Now
    .Update
    .Close
End With
WDrp "#Cpy #Cpy1 #Cpy2"
End Sub

Function IFxRate$()
IFxRate = PnmFfn("ZHT1")
End Function

Sub IFxRateOpn()
FxOpn IFxRate
End Sub

Sub DbtAddPfx(A As Database, T, Pfx)
DbtRen A, T, Pfx & T
End Sub
Sub LnkCcm()
'Ccm is stand for Space-[C]ir[c]umflex-accent
'Develop in local, some N:\ table is needed to be in currentdb.
'This N:\ table is dup in currentdb as ^xxx CcmTny
'When in development, each currentdb ^xxx is require to create a xxx table as linking to ^xxx
'When in N:\SAPAccessReports\ is avaiable, ^xxx is require to link to data-db as in Des
DrpLnkTbl
If IsDev Then
    LnkCcmLcl
Else
    LnkCcmNDrive
End If
End Sub
Sub LnkCcmLcl()
AyDo CcmTny, "CcmTbl_LnkLcl"
End Sub
Property Get ErCcmTny() As String()
ErCcmTny = AyWhPredFalse(CcmTny, "CcmTbl_IsVdt")
End Property
Property Get VdtCcmTny() As String()
VdtCcmTny = AyWhPred(CcmTny, "CcmTbl_IsVdt")
End Property
Property Get TblDes$(T)
TblDes = DbtDes(CurrentDb, T)
End Property
Property Let TblDes(T, Des$)
DbtDes(CurrentDb, T) = Des
End Property
Property Let DbtDes(A As Database, T, Des$)
TblSetPrp T, "Description", Des
End Property
Sub TblSetPrp(T, P, V)
DbtSetPrp CurrentDb, T, P, V
End Sub
Sub DbtSetPrp(A As Database, T, P, V)
If PrpsHasPrp(A.TableDefs(T).Properties, P) Then
    A.TableDefs(T).Properties(P).Value = V
Else
    A.TableDefs(T).Properties.Append DbtCrtPrp(A, T, P, V)
End If
End Sub
Function DbtCrtPrp(A As Database, T, P, V) As DAO.Property
Set DbtCrtPrp = A.TableDefs(T).CreateProperty(P, VarDaoTy(V), V)
End Function
Property Get DbtDes$(A As Database, T)
DbtDes = DbtPrp(A, T, "Description")
End Property
Function DbtHasPrp(A As Database, T, P) As Boolean
DbtHasPrp = PrpsHasPrp(A.TableDefs(T).Properties, P)
End Function
Function PrpsHasPrp(A As DAO.Properties, P) As Boolean
PrpsHasPrp = ItrHasNm(A, P)
End Function
Function DbtPrp(A As Database, T, P)
If Not DbtHasPrp(A, T, P) Then Exit Function
DbtPrp = A.TableDefs(T).Properties(P).Value
End Function
Function CcmTbl_IsVdt(A$) As Boolean
Dim D$, App As EApp, DtaFb$
D = TblDes(A)
If Not IsEAppStr(D) Then Exit Function
App = EAppStr_EApp(D)
DtaFb = EAppDtaFb(App)
If Not FfnIsExist(DtaFb) Then Exit Function
CcmTbl_IsVdt = True
End Function

Sub CcmTbl_LnkNDrive(A)
Dim EAppStr$, DtaFb$
EAppStr = TblDes(A)
DtaFb = EAppStr_DtaFb(EAppStr)
DbttLnkFb CurrentDb, Mid(A, 2), DtaFb, Mid(A, 2)
End Sub

Sub CcmTbl_LnkLcl(A)
If FstChr(A) <> "^" Then Stop
Dim T$
T = Mid(A, 2)
DbttLnkFb CurrentDb, T, CurrentDb.Name, A
End Sub

Sub LnkCcmNDrive()
Dim Vdt$(), Er$(), Av()
Av = AyPredSplit(CcmTny, "CcmTbl_IsVdt")
Vdt = Av(0)
Er = Av(1)
If Sz(Er) > 0 Then
    MsgBrw "These [table-des] are not pointing to a data fb", AyAlignT1(AyMap(Er, "TblTblDes"))
End If
AyDo Vdt, "CcmTbl_LnkNDrive"
MsgDmp "These [tables] are linked to data fb", AyMap(Vdt, "TblTblDes")
End Sub
Sub AySplit_xInto_T1Ay_RestAy_Asg(A, OT1Ay$(), ORestAy$())
Dim U&, J&
U = UB(A)
If U = -1 Then
    Erase OT1Ay, ORestAy
    Exit Sub
End If
ReDim OT1Ay(U)
ReDim ORestAy(U)
For J = 0 To U
    BrkAsg A(J), " ", OT1Ay(J), ORestAy(J)
Next
End Sub
Function AyabAdd(A, B, Optional Sep$) As String()
Dim O$(), J&, U&
U = UB(A): If U <> UB(B) Then Stop
If U = -1 Then Exit Function
ReDim O(U)
For J = 0 To U
    O(J) = A(J) & Sep & B(J)
Next
AyabAdd = O
End Function
Function AyabAddWSpc(A, B) As String()
AyabAddWSpc = AyabAdd(A, B, " ")
End Function
Function AyAlignT1(A) As String()
Dim T1$(), Rest$()
    AySplit_xInto_T1Ay_RestAy_Asg A, T1, Rest
T1 = AyAlignL(T1)
AyAlignT1 = AyabAddWSpc(T1, Rest)
End Function
Sub MsgDmp(A$, ParamArray Ap())
Dim Av(): Av = Ap
AyDmp MsgAv_Ly(A, Av)
End Sub
Function TblTblDes$(T)
TblTblDes = T & " " & TblDes(T)
End Function

Sub TblAddPfx(T, Pfx$)
DbtAddPfx CurrentDb, T, Pfx
End Sub

Sub DbttAddPfx(A As Database, TT, Pfx)
AyDoAXB CvTT(TT), "DbtAddPfx", A, Pfx
End Sub
Sub AyDoAXB(Ay, AXB$, A, B)
If Sz(Ay) = 0 Then Exit Sub
Dim X
For Each X In Ay
    Run AXB, A, X, B
Next
End Sub
Sub TTAddPfx(TT, Pfx$)
DbttAddPfx CurrentDb, TT, Pfx
End Sub

Function AyWhPredFalse(A, Pred$)
Dim O, X
O = AyCln(A)
If Sz(A) > 0 Then
    For Each X In A
        If Not Run(Pred, X) Then
            Push O, X
        End If
    Next
End If
AyWhPredFalse = O
End Function
Function AyWhPred(A, Pred$)
Dim O, X
O = AyCln(A)
If Sz(A) > 0 Then
    For Each X In A
        If Run(Pred, X) Then
            Push O, X
        End If
    Next
End If
AyWhPred = O
End Function

Function AyMap(A, Map$)
AyMap = AyMapInto(A, Map, EmpAy)
End Function

Sub MsgBrw(Msg$, ParamArray Ap())
Dim Av(): Av = Ap
MsgAv_Brw Msg, Av
End Sub