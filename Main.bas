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

Function ChkFil() As String()
ChkFil = AyAddAp(ChkF_MB52, ChkF_Uom, Either(IsFstYM, "ChkF_Rate ChkF_Inv"))
End Function

Function ChkF_Rate() As String()
If Not FfnIsExist(IFxRate) Then
    ChkF_Rate = MsgAp_Ly("It is first record of Year/Month, [rate file (ZHT1)] is needed.|It is not found in [folder]", _
    FfnFn(IFxRate), FfnPth(IFxRate))
    Exit Function
End If
ChkF_Rate = FxChkWs(IFxRate, "Rate file (ZHT1)", "8701 8601")
End Function

Function ChkF_MB52() As String()
ChkF_MB52 = FxChkWs(IFxMB52, MB52FnSpec)
End Function

Function ChkF_Inv() As String()
ChkF_Inv = FxChkWs(IFxInv, "Invoice file", "Invoices Detail")
End Function

Function ChkF_Uom() As String()
ChkF_Uom = FxChkWs(IFxUom, "Sales text file")
End Function

Function Lnk() As String()
Lnk = ChkFil: If Sz(Lnk) > 0 Then Exit Function
WtLnkFx ">InvH", IFxInv, "Invoices"
WtLnkFx ">InvD", IFxInv, "Detail"
WtLnkFx ">MB52", IFxMB52
WtLnkFx ">Uom", IFxUom
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
Lnk = ChkCol
End Function

Function ChkCol() As String()
Dim A$(), B$(), C$(), D$(), E$(), F$()
A = WtChkCol(">MB52", LnkColStr.MB52)
B = WtChkCol(">Uom", LnkColStr.Uom)
If IsFstYM Then
    C = WtChkCol(">ZHT18601", LnkColStr.ZHT1)
    D = WtChkCol(">ZHT18701", LnkColStr.ZHT1)
Else
    E = WtChkCol(">InvD", LnkColStr.InvD)
    F = WtChkCol(">InvH", LnkColStr.InvH)
End If
ChkCol = AyAddAp(A, C, B, D, E, F)
End Function

Sub Import()
WImp ">MB52", LnkColStr.MB52
WImp ">Uom", LnkColStr.Uom
If IsFstYM Then
    WImp ">ZHT18601", LnkColStr.ZHT1
    WImp ">ZHT18701", LnkColStr.ZHT1
Else
    WImp ">InvH", LnkColStr.InvH
    WImp ">InvD", LnkColStr.InvD
End If
End Sub

Function OupPth$()
OupPth = PthEns(CDbPth & "Output\")
End Function

Function IFbStkShpRate$()
If IsDev Then
    IFbStkShpRate = CurrentDb.Name
Else
    IFbStkShpRate = "N:\SAPAcessReports\StockShipRate\StockShipRate_Data.accdb"
End If
End Function

Property Get OupFx$()
OupFx = FfnNxt(OupPth & FmtQQ("? ?.xlsx", Apn, YYYYxMM))
End Property

Private Sub MsgSet(A$)
Form_Main.MsgSet A
End Sub
Private Sub MsgClr()
Form_Main.MsgClr
End Sub

Property Get IFxUom$()
IFxUom = PnmFfn("Uom")
End Property

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

Sub TblYM_Ins_UpToCurYM()
Dim Y As Byte, M As Byte, FmM As Byte, ToM As Byte
For Y = FstY To CurY
    FmM = IIf(FstY = Y, FstM, 1)
    ToM = IIf(CurY = Y, CurM, 12)
    For M = FmM To ToM
        If Not QQAny("Select Y from YM where Y=? and M=?", Y, M) Then
            QQRun "Insert into YM (Y,M) values(?,?)", Y, M
        End If
    Next
Next
End Sub

Property Get InvLdDte()
Q = FmtQQ("Select IR_LdDte from YM where Y=? and M=?", Y, M)
InvLdDte = RsV(CurrentDb.OpenRecordset(Q))
End Property

Property Get InvFn$()
InvFn = FmtQQ("Invoices ?.xlsx", YYYYxMM)
End Property

Property Get IFxInv$()
IFxInv = InvPth & InvFn
End Property

Property Get InvHom$()
If IsDev Then
    InvHom = CDbPth & "Sample\"
Else
    InvHom = PthEns(AppHom & "Import Invoices\")
End If
End Property

Property Get InvPth$()
InvPth = PthEns(InvHom & YYYY & "\")
End Property

Sub InvPthBrw()
PthBrw InvPth
End Sub

Sub LdInv(Optional IsForceLd As Boolean)
If IsFstYM Then Stop

'#IInvH & #IInvD are imported
'Replace InvH and InvD after validation
'
'#IInvD: VndShtNm InvNo Sku Sc Amt
'#IInvH: VndShtNm InvNo Dte Whs Sc Amt
'InvD: VndShtNm InvNo Sku Sc Amt
'InvH: VndShtNm InvNo Whs Dte Sc Amt DteCrt
If Not IsForceLd Then
    If InvIsLd Then Exit Sub
End If
Dim A$, Q$
A = IFxInv
Q = FmtQQ("Delete x.* from [InvD] where InvH in (Select InvH from InvH where Year(Dte)=? and Month(Dte)=?)", Y, M): W.Execute Q
Q = FmtQQ("Delete * from [InvH] where Year(Dte)=? and Month(Dte)=?", Y, M): W.Execute Q
W.Execute "insert into [InvH] (VndShtNm,InvNo,Whs,Dte,Sc,Amt)" & _
                      " select VndShtNm,InvNo,Whs,Dte,Sc,Amt from [#IInvH]'"
W.Execute "Alter Table [#IInvD] add column InvH Long"
W.Execute "Update [#IInvD] x inner join [InvH] a on x.VndShtNm=a.VndShtNm and x.InvNo=a.InvNo set x.InvH=a.InvH"
W.Execute "insert into [InvD] (InvH,Sku,Sc,Amt)" & _
                      " select InvH,Sku,Sc,Amt from [#IInvD]'"
Q = FmtQQ("Select IR_Fx, IR_FxSz, IR_FxTim, IR_LdDte from YM where Y=? and M=?", Y, M)

RsUpdDr W.OpenRecordset(Q), FfnStamp(A)
End Sub

Function IRDrLy(A()) As String()
Dim Fx$, Sz&, Tim As Date, Sc#, Amt@, NInv%, NInvLin%
AyAsg A, Fx, Sz, Tim, Sc, Amt, NInv, NInvLin
IRDrLy = MsgLy("[Invoice file] of [time] and [size] with [n-invoices], [n-lines], [total-Sc] and [total-amt] are loaded in [year] and [month]", _
    Fx, Tim, FfnSz(A), NInv, NInvLin, Round(Sc, 1), Round(Amt, 2), Y + 2000, M)
End Function

Function TmpInvHD_IRDr(Fx) As Variant()
Dim Sz&, Tim As Date, Sc#, Amt@, NInv%, NInvLin%, NSku%
With W.OpenRecordset("Select Count(*), Sum(Amt), Sum(Sc) from [#IInvH]")
    NInv = .Fields(0).Value
    Amt = .Fields(1).Value
    Sc = .Fields(2).Value
    .Close
End With
NSku = W.OpenRecordset("Select Count(*) from (Select Distinct Sku from [#IInvD])").Fields(0).Value
NInvLin = W.OpenRecordset("Select Count(*) from [#IInvD]").Fields(0).Value
TmpInvHD_IRDr = Array(Fx, Sz, Tim, Sc, Amt, NInv, NInvLin, NSku, Now)
End Function
Property Get MB52FnSpec$()
MB52FnSpec = "MB52 " & YYYYxMM & "-??.xls"
End Property
Property Get IniMB52FnSpec$()
IniMB52FnSpec = "MB52 " & YYYYxMM & "-??.xls"
End Property

Property Get MB52IsLd() As Boolean
MB52IsLd = FfnTSz(IFxMB52) = MB52LdTSz
End Property

Sub LdMB52(Optional IsForceLd As Boolean)
If Not IsForceLd Then
    Lg "LdMB52", "[MB52IsLd] with [IFxMB52] [TSz] <> [MB52LdTSz]", MB52IsLd, IFxMB52, FfnTSz(IFxMB52), MB52LdTSz
    If MB52IsLd Then
        Exit Sub
    End If
End If
'#IMB52 is imported into YMTbl and OHTbl
'Import into YMOH & Update YM
WDrp "#OH"
Q = "Select Distinct Sku,Whs,Sum(x.QUnRes+x.QInsp+x.QBlk) as EndOH into [#OH] from [#IMB52] x group by Sku,Whs": W.Execute Q
Q = FmtQQ("Delete from [YMOH] where Y=? and M=?", Y, M): W.Execute Q
Q = FmtQQ("Insert into [YMOH] (Y,M,Sku,Whs,EndOH) select ?,?,Sku,Whs,EndOH from [#OH]", Y, M): W.Execute Q

'Update YM: Y M *Fx *FxTim *FxSz *NRec *NSku *Sc *DteLd
Q = FmtQQ("Select EndOH_Fx, EndOH_FxSz, EndOH_FxTim, EndOH_LdDte from YM where Y=? and M=?", Y, M)
RsUpdDr W.OpenRecordset(Q), FfnStamp(IFxMB52)
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
Const IsForceLd As Boolean = True
If Not Cfm Then Exit Sub
WIni
If AyBrwEr(Lnk) Then Exit Sub
Import
LdDta True
If Not IsFstYM Then
    Oup
    Upd
    Gen
End If
WCls
End Sub

Sub LdDta(Optional IsForceLd As Boolean)
Lg "LdDta", "Start with [IsForceLd]", IsForceLd
LdMB52 IsForceLd
If IsFstYM Then
    LdRate IsForceLd
Else
    LdInv IsForceLd
End If
Lg "LdDta", "End"
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
'Update YMRate: Y M Sku Whs RateSc
'By @Main
Q = FmtQQ("Delete * from [YMRate] where Y=? and M=?", Y, M): W.Execute Q
Q = FmtQQ("Insert into YMRate (Y,M,Sku,Whs,RateSc) Select ?,?,Sku,Whs,EndRateSc from [@Main]", Y, M): W.Execute Q
End Sub

Sub OMain()
'YMRate: Y M ZHT1 Whs RateSc FmDte ToDte DteCrt
'YMOH:   Y M Sku Whs OH Sc Sc_U
'InvD: VndShtNm InvNo Sku Sc Amt
'InvH: VndShtNm InvNo Whs Dte Sc Amt DteCrt
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

If IsFstYM Then Stop
WDrp "#BegOH #EndOH #IR #SkuWhs #SkuWhs1 #Rate @Main"

'#EndOH Sku Whs EndOH EndOHSc Sc_U
Q = FmtQQ("Select Sku,Whs,EndOH into [#EndOH] from [YMOH] where Y=? and M=?", Y, M): W.Execute Q
W.Execute "Alter Table [#EndOH] Add Column Sc_U Single, EndOHSc Single"
W.Execute "Update [#EndOH] x inner join [#IUom] a on x.Sku=a.Sku and x.Whs=a.Whs set x.Sc_U=a.Sc_U"
W.Execute "Update [#EndOH] set EndOHSc=EndOH/Sc_U where Nz(Sc_U,0)<>0"

'#BegOH Sku Whs BegOH BegOHSc Sc_U
Q = FmtQQ("Select Sku,Whs,EndOH as BegOH into [#BegOH] from [YMOH] where Y=? and M=?", BegY, BegM): W.Execute Q
W.Execute "Alter Table [#BegOH] Add Column Sc_U Single, BegOHSc Single"
W.Execute "Update [#BegOH] x inner join [#IUom] a on x.Sku=a.Sku and x.Whs=a.Whs set x.Sc_U=a.Sc_U"
W.Execute "Update [#BegOH] set BegOHSc=BegOH/Sc_U where Nz(Sc_U,0)<>0"
'#IR Sku Whs IRSc IRAmt
Q = FmtQQ("Select Distinct Sku,Whs,Sum(x.Sc) as IRSc, Sum(a.Amt) as IRAmt into [#IR]" & _
" from [InvD] x inner join [InvH] a on x.InvH=a.InvH" & _
" where a.Dte between #?# and #?#" & _
" group by Sku,Whs", _
FmYYYYxMMxDD, ToYYYYxMMxDD): W.Execute Q

'#SkuWhs
W.Execute "Select Sku,Whs into [#SkuWhs1] from [#BegOH]"
W.Execute "Insert into [#SkuWhs1] Select Sku,Whs from [#EndOH]"
W.Execute "Insert into [#SkuWhs1] Select Sku,Whs from [#IR]"
W.Execute "Select Distinct Sku,Whs into [#SkuWhs] from [#SkuWhs1]"

'@Main
W.Execute "Select x.Sku,x.Whs, BegOHSc, EndOHSc, IRSc, IRAmt" & _
" into [@Main]" & _
" from (([#SkuWhs] x" & _
" left join [#BegOH] a on x.Sku=a.Sku and x.Whs=a.Whs)" & _
" left join [#EndOH] b on x.Sku=b.Sku and x.Whs=b.Whs)" & _
" left join [#IR]    c on x.Sku=c.Sku and x.Whs=c.Whs"

'AddCol ProdH M32 M35 M38
W.Execute "Alter Table [@Main] add column Sc_U double, ProdH text(20),M32 Text(2), M35 Text(5), M38 Text(8)"
W.Execute "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku and x.Whs=a.Whs set x.ProdH=a.ProdH, x.Sc_U=a.Sc_U"
W.Execute "Update [@Main] set M32=Mid(ProdH,3,2),M35=Mid(ProdH,3,5),M38=Mid(ProdH,3,8)"

'AddCol ZHT1Sku BegRateSc
W.Execute "Alter Table [@Main] add column ZHT1 Text(8), BegRateSc Double"
If IsSndYM Then
    W.Execute "Select ZHT1,Whs,RateSc as BegRateSc into [#Rate] from [IniRate]"
    W.Execute "Create Index Pk on [#Rate] (ZHT1,Whs) with primary"
    W.Execute "Update [@Main] x inner join [#Rate] a on x.Whs=a.Whs and x.M38=a.ZHT1 set x.BegRateSc=a.BegRateSc,x.ZHT1=a.ZHT1 where x.BegRateSc Is Null"
    W.Execute "Update [@Main] x inner join [#Rate] a on x.Whs=a.Whs and x.M35=a.ZHT1 set x.BegRateSc=a.BegRateSc,x.ZHT1=a.ZHT1 where x.BegRateSc Is Null"
    W.Execute "Update [@Main] x inner join [#Rate] a on x.Whs=a.Whs and x.M32=a.ZHT1 set x.BegRateSc=a.BegRateSc,x.ZHT1=a.ZHT1 where x.BegRateSc Is Null"
Else
    Q = FmtQQ("Select Sku,Whs,RateSc as BegRateSc into [#Rate] from [YMRate] where Y=? and M=?", Y, M): W.Execute Q
    W.Execute "Create Index Pk on [#Rate] (Sku,Whs) with primary"
    W.Execute "Update [@Main] x inner join [#Rate] a on x.Whs=a.Whs and x.Sku=a.Sku set x.BegRateSc=a.BegRateSc"
End If

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

W.Execute "Update [@Main] set SellSc = Nz(BegOHSc,0) + Nz(IRSc,0) - Nz(EndOHSc)"
W.Execute "Update [@Main] set SellAmt = Nz(SellSc,0) * Nz(BegRateSc,0)"
W.Execute "Update [@Main] set BegAmt = Nz(BegOHSc,0) * Nz(BegRateSc,0)"
W.Execute "Update [@Main] set EndAmt = Nz(BegAmt,0) + Nz(IRAmt,0) - Nz(SellAmt,0)"
W.Execute "Update [@Main] set EndRateSc = IIf(Nz(EndOHSc,0)=0,Nz(BegRateSc,0),Nz(EndAMt,0)/EndOHSc)"

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
'WDrp "#SkuWhs #SkuWhs1 #BegOH #EndOH #IR #Rate"
End Sub

Sub LdRate(Optional IsForceLd As Boolean)
If Not IsForceLd Then
    If RateIsLd Then
        Dim Kind$
        AyDmp FfnAlreadyLdMsgLy(IFxRate, Kind, RateLdDTim)
        Exit Sub
    End If
End If
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
Q = FmtQQ("Delete * from [#Cpy]" & _
" Where not #?# between FmDte and ToDte", YYYYxMMxLasDD): W.Execute Q

Q = "Delete * from [IniRate]": W.Execute Q
Q = "Insert into [IniRate] (ZHT1,Whs,RateSc,FmDte,ToDte) select ZHT1,Whs,RateSc,FmDte,ToDte from [#Cpy]": W.Execute Q
Q = "Select Fx, FxSz, FxTim, LdDte from [IniRateH]"
If DbtNRec(W, "IniRateH") = 0 Then
    W.Execute "Insert into [IniRateH] (Fx) values ('x')"
End If
RsUpdDr W.OpenRecordset(Q), FfnStamp(IFxRate)
WDrp "#Cpy #Cpy1 #Cpy2"
End Sub
Property Get IFxRate$()
IFxRate = PnmFfn("ZHT1")
End Property

Sub IFxRateOpn()
FxOpn IFxRate
End Sub
Property Get InvLdDTim$()
InvLdDTim = QQDTim("Select IR_LdDte from YM where Y=? and M=? and IR_Fx='?'", Y, M, IFxInv)
End Property
Property Get RateLdDTim$()
RateLdDTim = QQDTim("Select LdDte from IniRate")
End Property
Property Get MB52LdDTim$()
MB52LdDTim = QQDTim("Select EndOH_LdDte from YM where Y=? and M=? and EndOH_Fx='?'", Y, M, IFxMB52)
End Property
Function FxLdTSz$(A, Optional ByVal FldPfx$)
Dim P$
P = FldPfx
Q = FmtQQ("Select ?_FxTim,?_FxSz from YM where ?_Fx='?' and Y=? and M=?", P, P, P, A, Y, M)
FxLdTSz = RsTSz(W.OpenRecordset(Q))
End Function
Property Get InvLdTSz$()
InvLdTSz = FxLdTSz(IFxInv, "IR")
End Property
Property Get MB52LdTSz$()
MB52LdTSz = FxLdTSz(IFxInv, "EndOH")
End Property
Property Get RateLdTSz$()
Q = FmtQQ("Select FxTim,FxSz from IniRate")
RateLdTSz = RsTSz(W.OpenRecordset(Q))
End Property

Property Get MB52TSz$()
MB52TSz = FxLdTSz(IFxMB52, "EndOH")
End Property
Property Get RateIsLd() As Boolean
RateIsLd = FfnTSz(IFxRate) = RateLdTSz
End Property
Property Get RateIsCalced() As Boolean

End Property
Property Get InvIsLd() As Boolean
InvIsLd = FfnTSz(IFxInv) = InvLdTSz
End Property