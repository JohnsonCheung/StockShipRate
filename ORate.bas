Option Compare Database
Sub Calc()
'The @Main is the detail of showing how NxtMth YMZHT1 is calculate
Y = 18
M = 2
AssYM
'If IsNoDta Then Exit Sub

WIni
WttLnkFb "InvH InvD YM YMOH YMZHT1", IFbStkShpRate
WtLnkFx ">Uom", IFxUOM
WImp ">Uom", LnkColStr.Uom
'Reset_YM           this should be put in front of starting the whole process
'Reset_YMZHT1
Tmp
Oup
Gen
WCls
End Sub

Sub Gen()
End Sub
Sub Oup()
OMain
End Sub
Sub OMain()
'YMZHT1: Y M ZHT1 Whs RateSc FmDte ToDte DteCrt
QQ "Delete * from [YMZHT1] where Y=? and M=?", Y, M
QQ "Insert into YMZHT1 (Y,M,ZHT1,Whs,RateSc) Select ?,?,ZHT1,Whs,RateSc from [$ZHT1]", Y, M
TmpRate_Upd_YM "$ZHT1"
WDrp "$ZHT1"
End Sub

Sub Tmp()
TmpRate
TmpZHT1
End Sub
Sub TmpRate()
WDrp "$Rate"
WQQ "Select ZHT1,Whs,RateSc into [$Rate] from [YMZHT1] where Y=? and M=?", Y, M
WQQ "Create Index Pk on [$Rate] (ZHT1,Whs) with primary"
End Sub
Sub TmpZHT1()
'YMZHT1: Y M ZHT1 Whs RateSc FmDte ToDte DteCrt
'YMOH: Y M Sku Whs OH Sc Sc_U
'InvD: VndShtNm InvNo Sku Sc Amt
'InvH: VndShtNm InvNo Whs Dte Sc Amt DteCrt
'#IUom     Sku Whs Des StkUom Sc_U ProdH

'Given: BegOHSc   =  100(Sc)
'       BegRateSc =  $0.5/Sc  => BegAmt = $50
'       IRSc      =  30(Sc)
'       IRAmt     =  $21      => IRRateSc = $0.7/Sc
'       EndOHSc   =  40(Sc)
'To Find: EndRateSc
'Work:
'      SellSc    = BegOHSc + IRSc - EndOHSc    = 100(Sc) + 30(Sc) - 40(Sc) = 90(Sc)
'      SellAmt   = SellSc * OldRateSc          = 90(Sc) * $0.5/Sc          = $45
'      EndAmt    = BegAmt + IRAmt - SellAmt    = $50 + $21 - $45           = $26
'      EndRateSc = EndAmt / EndOHSc            = $26 / 40(Sc)              = $0.65/Sc (**)
WDrp "#BegOH #EndOH #IR #SkuWhs #SkuWhs1 @Main"

'#BegOH
'WQQ "Select Sku,Whs,Sc as BegOHSc into [#BegOH] from [YMOH] where Y=? and M=?", Y, M
'Tmp_AddRateColumns "#BegOH"

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
" from [InvD] x inner join [InvH] a on x.InvNo=a.InvNo and x.VndShtNm=a.VndShtNm" & _
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

'Add ZHT1 RateSc
WRun "Update [@Main] x inner join [$Rate] a on x.Whs=a.Whs and x.M38=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"
WRun "Update [@Main] x inner join [$Rate] a on x.Whs=a.Whs and x.M35=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"
WRun "Update [@Main] x inner join [$Rate] a on x.Whs=a.Whs and x.M32=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"

'@Main Sku ZHT1 BegOHSc BegRateSc BegAmt IRSc IRAmt OldRateSc EndAmt EndOHSc EndRateSc
             'Sku Whs
             'ZHT1 ZBrdNm ZBrd ZQlyNm ZQly Z8Nm Z8
             'EndAmt = EndOHSc
'Des StkUom
WRun "Alter Table [@Main] Add Column Des Text(255), StkUom Text(10)"
WRun "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku set x.Sc_U = a.Sc_U,x.Des=a.Des,x.StkUom=a.StkUom"

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
Stop
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
Sub Reset_YMZHT1()
DoCmd.RunSQL FmtQQ("Delete * from [YMZHT1] where Y>? or (Y=? and M>?)", Y, Y, M)
End Sub

Function IsNoDta() As Boolean
IsNoDta = True
With QQSqlRs("Select RateSc_NRec,BegOH_NRec,IR_NInvLin from YM where Y=? and M=?", Y, M)
    If !RateSc_NRec = 0 Then
        MsgBox "No rate yet", vbCritical
        Exit Function
    End If
    If !IR_NInvLin = 0 Then
        MsgBox "No invoices yet", vbCritical
        Exit Function
    End If
    If !BegOH_NRec = 0 Then
        MsgBox "No begin OH yet", vbCritical
        Exit Function
    End If
    .Close
End With
IsNoDta = False
End Function


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