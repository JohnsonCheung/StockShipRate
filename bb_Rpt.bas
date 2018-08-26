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
Const TT$ = "CurYM IniRate IniRateH InvH InvD YM YMRate YMOH"
Dim Fbtt$()
If IsDev Then
    Fbtt = AyAddPfx(CvTT(TT), "^")
End If
WttLnkFb TT, IFbStkShpRate, Fbtt
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

Sub LoadInv(Optional IsForceLoad As Boolean)
'#IInvH & #IInvD are imported
'Replace InvH and InvD after validation
'
'#IInvD: VndShtNm InvNo Sku Sc Amt
'#IInvH: VndShtNm InvNo Dte Whs Sc Amt
'InvD: VndShtNm InvNo Sku Sc Amt
'InvH: VndShtNm InvNo Whs Dte Sc Amt DteCrt
If Not IsForceLoad Then
    If IsLd_xInv Then Exit Sub
End If
Dim A$, Q$
A = IFxInv
Q = FmtQQ("Delete x.* from [InvD] where InvH in (Select InvH from InvH where Year(Dte)=? and Month(Dte)=?)", Y, M): W.Execute Q
Q = FmtQQ("Delete * from [YMInvH] where Year(Dte)=? and Month(Dte)=?", Y, M): W.Execute Q
W.Execute "insert into [InvH] (VndShtNm,InvNo,Whs,Dte,Sc,Amt)" & _
                   " select VndShtNm,InvNo,Whs,Dte,Sc,Amt from [#IInvH]'"
W.Execute "Alter Table [#IInvD] add column InvH Long"
W.Execute "Update [#IInvD] x inner join [YMInvH] a on x.VndShtNm=a.VndShtNm and x.InvNo=a.InvNo set x.InvH=a.InvH"
W.Execute "insert into [YMInvD] (InvH,Sku,Sc,Amt)" & _
                   " select InvH,Sku,Sc,Amt from [#IInvD]'"
Q = FmtQQ("Select IR_Fx, IR_FxSz, IR_FxTim, IR_LoadDte from YM where Y=? and M=?", Y, M)

RsUpdDr WQRs(Q), FfnStamp(IFxInv)
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
Q = "Select Distinct Sku,Whs,Sum(x.QUnRes+x.QInsp+x.QBlk) as OH into [#OH] from [#IMB52] x group by Sku,Whs": W.Execute Q
Q = FmtQQ("Delete from [YMOH] where Y=? and M=?", Y, M): W.Execute Q
Q = FmtQQ("Insert into [YMOH] (Y,M,Sku,Whs,OH) select ?,?,Sku,Whs,OH from [#OH]", Y, M): W.Execute Q

'Update YM: Y M *Fx *FxTim *FxSz *NRec *NSku *Sc *DteLoad
Q = FmtQQ("Select EndOH_Fx, EndOH_FxSz, EndOH_FxTim, EndOH_LoadDte from YM where Y=? and M=?", Y, M)
RsUpdDr WQRs(Q), FfnStamp(IFxMB52)
WDrp "#OH"
End Sub
Property Get IFxMB52$()
IFxMB52 = AyMax(MB52y)
End Property
Property Get MB52y() As String()
MB52y = PthFfnAy(MB52Pth, MB52FnSpec)
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
W.Execute "Alter Table [#EndOH] Add Column Sc_U Single, EndOHSc Single"
W.Execute "Update [#EndOH] x inner join [#IUom] a on x.Sku=a.Sku and x.Whs=a.Sku set x.Sc_U=a.Sc_U"
W.Execute "Update [#EndOH] set EndOHSc=EndOH/Sc_U where Sc_U is not null"

'#BegOH Sku Whs BegOH BegOHSc Sc_U
WQQ "Select Sku,Whs,OH as BegOH into [#BegOH] from [YMOH] where Y=? and M=?", BegY, BegM
W.Execute "Alter Table [#BegOH] Add Column Sc_U Single, BegOHSc Single"
W.Execute "Update [#BegOH] x inner join [#IUom] a on x.Sku=a.Sku and x.Whs=a.Whs set x.Sc_U=a.Sc_U"
W.Execute "Update [#BegOH] set BegOHSc=BegOH/Sc_U where Sc_U is not null"

'#IR Sku Whs IRSc IRAmt
WQQ "Select Distinct Sku,Whs,Sum(x.Sc) as IRSc, Sum(a.Amt) as IRAmt into [#IR]" & _
" from [YMInvD] x inner join [YMInvH] a on x.InvNo=a.InvNo and x.VndShtNm=a.VndShtNm" & _
" where a.Dte between #?# and #?#" & _
" group by Sku,Whs", _
FmYYYYxMMxDD, ToYYYYxMMxDD

'#SkuWhs
W.Execute "Select Sku,Whs into [#SkuWhs1] from [#BegOH]"
W.Execute "Insert into [#SkuWhs1] Select Sku,Whs from [#EndOH]"
W.Execute "Insert into [#SkuWhs1] Select Sku,Whs from [#IR]"
W.Execute "Select Distinct Sku,Whs into [#SkuWhs] from [#SkuWhs1]"

'@Main
W.Execute "Select x.Sku,x.Whs, BegOHSc, EndOHSc, IRSc,IRAmt" & _
" into [@Main]" & _
" from (([#SkuWhs] x" & _
" left join [#BegOH] a on x.Sku=a.Sku and x.Whs=a.Whs)" & _
" left join [#EndOH] b on x.Sku=b.Sku and x.Whs=b.Whs)" & _
" left join [#IR]    c on x.Sku=c.Sku and x.Whs=c.Whs"

'AddCol ProdH M32 M35 M38
W.Execute "Alter Table [@Main] add column Sc_U double, ProdH text(20),M32 Text(2), M35 Text(5), M38 Text(8),ZHT1 Text(8), RateSc Double"
W.Execute "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku and x.Whs=a.Whs set x.ProdH=a.ProdH, x.Sc_U=a.Sc_U"
W.Execute "Update [@Main] set M32=Mid(ProdH,3,2),M35=Mid(ProdH,3,5),M38=Mid(ProdH,3,8)"

'AddCol ZHT1 RateSc
W.Execute "Update [@Main] x inner join [$Rate] a on x.Whs=a.Whs and x.M38=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"
W.Execute "Update [@Main] x inner join [$Rate] a on x.Whs=a.Whs and x.M35=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"
W.Execute "Update [@Main] x inner join [$Rate] a on x.Whs=a.Whs and x.M32=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"

'RenCol RateSc -> BegRateSc
WReOpn
WtRenCol "@Main", "RateSc", "BegRateSc"

'@Main Sku ZHT1 BegOHSc BegRateSc BegAmt IRSc IRAmt OldRateSc EndAmt EndOHSc EndRateSc
             'Sku Whs
             'ZHT1 ZBrdNm ZBrd ZQlyNm ZQly Z8Nm Z8
             'EndAmt = EndOHSc
'AddCol Des StkUom
W.Execute "Alter Table [@Main] Add Column Des Text(255), StkUom Text(10)"
W.Execute "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku set x.Sc_U = a.Sc_U,x.Des=a.Des,x.StkUom=a.StkUom"

'AddCol BegAmt SellSc SellAmt EndAmt EndRateSc
W.Execute "Alter Table [@Main]" & _
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
'W.Execute "Alter Table [@Main] add column Stream Text(10), Topaz Text(20), ProdH Text(8), F2 Text(2), M32 text(2), M35 text(5), M38 Text(8), ZHT1 Text(8), Z2 text(2), Z5 text(5), Z8 Text(8), RateSc Currency, Amt Currency"
'W.Execute "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku set x.ProdH=a.ProdH,x.Topaz=a.Topaz"
'W.Execute "Update [@Main] set Stream=IIf(Left(Topaz,3)='UDV','Diageo','MH')"
'W.Execute "Update [@Main] Set Z2=Left(ZHT1,2), Z5=Left(ZHT1,5), Z8=Left(ZHT1,8) where not ZHT1 is null"
'W.Execute "Update [@Main] Set Amt = RateSc * OH_Sc where RateSc is not null"
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

Sub LoadRate(Optional IsForceLoad As Boolean)
If Not IsForceLoad Then
    If IsLd_xRate Then
        Dim Kind$
        AyDmp FfnAlreadyLoadedMsgLy(IFxRate, Kind, LdDTim_xRate)
        Exit Sub
    End If
End If
If IsFstYM Then
    LoadRate_xFmZHT1
Else
    LoadRate_xFmCalc
End If
End Sub

Sub LoadRate_xFmCalc()

End Sub
Sub LoadRate_xFmZHT1()
If Not IsFstYM Then Stop
WDrp "#Cpy1 #Cpy2 #Cpy"
W.Execute "Select '8701' as Whs,x.* into [#Cpy1] from [#IZHT18701] x"
W.Execute "Select '8601' as Whs,x.* into [#Cpy2] from [#IZHT18601] x"

W.Execute "Select * into [#Cpy] from [#Cpy1] where False"
W.Execute "Insert into [#Cpy] select * from [#Cpy1]"
W.Execute "Insert into [#Cpy] select * from [#Cpy2]"

W.Execute "Alter Table [#Cpy] Add Column FmDte Date,ToDte Date"
W.Execute "Update [#Cpy] Set" & _
" FmDte = DateSerial(RIGHT(VdtFm,4),MID(VdtFm,4,2),LEFT(VdtFm,2))," & _
" ToDte = DateSerial(RIGHT(VdtTo,4),MID(VdtTo,4,2),LEFT(VdtTo,2))"
WQQ "Delete * from [#Cpy]" & _
" Where not #?# between FmDte and ToDte", YYYYxMM & "-01"

Q = "Delete * from [IniRate]": W.Execute Q
Q = "Insert into [IniRate] (ZHT1,Whs,RateSc,FmDte,ToDte) select ZHT1,Whs,RateSc,FmDte,ToDte from [#Cpy]": W.Execute Q
Q = "Select Fx, FxSz, FxTim, LoadDte from [IniRateH]"
RsUpdDr WQRs(Q), FfnStamp(IFxRate)
WDrp "#Cpy #Cpy1 #Cpy2"
End Sub
Function IFxRate$()
IFxRate = PnmFfn("ZHT1")
End Function

Sub IFxRateOpn()
FxOpn IFxRate
End Sub