Option Compare Database
Option Explicit
Type Pm
    BegYYMM As String
    EndYYMM As String
    BegDte As Date
    EndDte As Date
End Type
Const LnkColStrzZHT1$ = _
" ZHT1   Txt Brand  |" & _
" RateSc Dbl Amount |" & _
" VdtFm  Txt [Valid From]  |" & _
" VdtTo  Txt [Valid to]"

Const LnkColStrzUom$ = _
 "Sku    Txt Material |" & _
 "Des    Txt [Material Description] |" & _
 "Sc_U   Txt SC |" & _
 "StkUom Txt [Base Unit of Measure] |" & _
 "Topaz  Txt [Topaz Code] |" & _
 "ProdH  Txt [Product hierarchy]"
 
Const LnkColStrzMB52$ = _
    " Sku    Txt Material |" & _
    " Whs    Txt Plant    |" & _
    " QInsp  Dbl [In Quality Insp#]|" & _
    " QUnRes Dbl UnRestricted|" & _
    " QBlk   Dbl Blocked"
'    " Loc    Txt [Storage Location] |" & _
'    " BchNo  Txt Batch |"

Sub Import()
'Create 5-Imp-Table [#I*] from 5-lnk-table [>*]
WImp ">UOM", LnkColStrzUom
WImp ">ZHT18601", LnkColStrzZHT1
WImp ">ZHT18701", LnkColStrzZHT1
WImp ">MB52", LnkColStrzMB52
End Sub
Sub MsgRunQry(A$)
MsgSet "Running query (" & A & ") ..."
End Sub
Function Lnk() As String()
Dim A$(), B$(), C$(), D$(), E$(), F$(), O$()
A = WtLnkFx(">UOM", IFxUOM)
B = WtLnkFx(">MB52", IFxMB52)
C = WtLnkFx(">ZHT18601", IFxZHT1, "8601")
D = WtLnkFx(">ZHT18701", IFxZHT1, "8701")
O = AyAddAp(A, B, C)
If Sz(O) > 0 Then Lnk = O: Exit Function
A = WtChkCol(">UOM", LnkColStrzUom)
B = WtChkCol(">MB52", LnkColStrzMB52)
C = WtChkCol(">ZHT18601", LnkColStrzZHT1)
C = WtChkCol(">ZHT18701", LnkColStrzZHT1)
Lnk = AyAddAp(A, B, C, D)
End Function

Function Rpt()
WIni
If AyBrwEr(Lnk) Then Exit Function
If AyBrwEr(Er) Then Exit Function
Import
Oup
Gen
WQuit
End Function
Function Er() As String()
End Function
Function ErzMB52_8601_8701_Missing() As String()
Dim N&, O$()
DbtLnkFx W, "#A", IFxMB52
N = DbtNRow(W, "#A", "[Plant]='8601' or [Plant='8701'")
WDrp "#A"
If N = 0 Then
    Push O, "MB52 Excel: " & IFxMB52
    Push O, "Worksheet : Sheet1"
    Push O, "Above MB52 file has no data for [Plant]=8601 or [Plant]=8701"
    Push O, "------------------------------------------------------------"
    'ErzMB52_8601_0002_Missing = O
End If
End Function
Function IFxZHT1$()
IFxZHT1 = PnmFfn("ZHT1")
End Function
Function TpFx$()
TpFx = TpPth & Apn & "(Template).xlsx"
End Function
Function TpFxm$()
TpFxm = TpPth & Apn & "(Template).xlsm"
End Function
Function OupPth$()
Dim A$
A = CurDbPth & "Output\"
PthEns A
OupPth = A
End Function
Function OupFx$()
Dim A$, B$
A = OupPth & FmtQQ("? ?.xlsx", Apn, Mid(PnmVal("MB52Fn"), 6, 10))
B = FfnNxt(A)
OupFx = B
End Function
Sub TpOpn()
FxOpn TpFx
End Sub
Function TpWb() As Workbook
Set TpWb = FxWb(TpFx)
End Function
Function TpWsCdNy() As String()
TpWsCdNy = FxWsCdNy(TpFx)
End Function
Private Sub MsgSet(A$)
Form_Main.MsgSet A
End Sub
Private Sub MsgClr()
Form_Main.MsgClr
End Sub
Sub Gen()
MsgSet "Export to Excel ....."
OupFx_Gen OupFx, WFb
End Sub
Function IFxMB52$()
IFxMB52 = PnmFfn("MB52")
End Function
Function IFxUOM$()
IFxUOM = PnmFfn("UOM")
End Function

Property Get Pm() As Pm
Static X As Boolean, Y As Pm
If Not X Then
    X = True
    With Y
        .BegYYMM = PnmVal("MEYYYYMM")
        .BegDte = YYMM_FstDte(.BegYYMM)
        .EndDte = Pm__EndDte(.BegDte, CByte(PnmVal("NMthGR")))
        .EndYYMM = DteYYMM(.EndDte)
    End With
End If
Pm = Y
End Property
Private Function Pm__EndDte(BegDte As Date, NMthGR%) As Date
Dim D As Date
D = DateTime.DateAdd("M", -NMthGR, BegDte)
End Function
Sub ORate()
'VdtFm & VdtTo format DD.MM.YYYY
'1: #IZHT1 VdtFm VdtTo L3 RateSc
'2: #IUom     SKu Sc_U
'O: @Rate  ZHT1 RateSc
WDrp "#Cpy1 #Cpy2 #Cpy @Rate"
WRun "Select '8701' as Whs,x.* into [#Cpy1] from [#IZHT18701] x"
WRun "Select '8601' as Whs,x.* into [#Cpy2] from [#IZHT18601] x"

WRun "Select * into [#Cpy] from [#Cpy1] where False"
WRun "Insert into [#Cpy] select * from [#Cpy1]"
WRun "Insert into [#Cpy] select * from [#Cpy2]"

WRun "Alter Table [#Cpy] Add Column VdtFmDte Date,VdtToDte Date,IsCur YesNo"
WRun "Update [#Cpy] Set" & _
" VdtFmDte = DateSerial(RIGHT(VdtFm,4),MID(VdtFm,4,2),LEFT(VdtFm,2))," & _
" VdtToDte = DateSerial(RIGHT(VdtTo,4),MID(VdtTo,4,2),LEFT(VdtTo,2))"
WRun "Update [#Cpy] set IsCur = true where Now between VdtFmDte and VdtToDte"

WRun "Select Whs,ZHT1,RateSc into [@Rate] from [#Cpy]"
WDrp "#Cpy #Cpy1 #Cpy2"
End Sub

Sub OMain()
'Pm BegYYMM BegDte EndYYMM EndDte
'MEMB52  YYMM Sku BchNo Whs Loc QUnRes QInsp QBlk
'ZHT1  Sku ZHT1 Whs FmDte ToDte RateSc
'#IUom     Sku Whs Des StkUom Sc_U ProdH
'#IProdH   ProdH Nm
'#IInvH
'#IInvD
'Given: BegOHSc   =  100(Sc)
'       BegRateSc =  $0.5/Sc  => BegAmt = $50
'       GRSc      =  30(Sc)
'       GRAmt     =  $21      => GRRateSc = $0.7/Sc
'       EndOHSc   =  40(Sc)
'To Find: EndRateSc
'Work:
'      SellSc    = BegOHSc + GRSc - EndOHSc    = 100(Sc) + 30(Sc) - 40(Sc) = 90(Sc)
'      SellAmt   = SellSc * OldRateSc          = 90(Sc) * $0.5/Sc          = $45
'      EndAmt    = BegAmt + GRAmt - SellAmt    = $50 + $21 - $45           = $26
'      EndRateSc = EndAmt / EndOHSc            = $26 / 40(Sc)              = $0.65/Sc (**)
WDrp "#BegOH #EndOH @Main"
WRun FmtQQ("Select Sku,Wh,OH as BegOH into [#BegOH] where YYMM='?'", Pm.BegYYMM)
WRun "Alter Table [#BegOH] Add Column Sc_U Single, BegOHSc Single"
WRun "Update [#BegOH] x inner join [#IUom] a on x.Sku=a.Sku and x.Whs=a.Sku set x.Sc_U=a.Sc_U"
WRun "Update [#BegOH] set BegOHSc=BegOH/Sc_U where Sc_U is not null"

WRun FmtQQ("Select Sku,Wh,OH as EndOH into [#EndOH] where YYMM='?'", Pm.EndYYMM)
WRun "Alter Table [#EndOH] Add Column Sc_U Single, EndOHSc Single"
WRun "Update [#EndOH] x inner join [#IUom] a on x.Sku=a.Sku and x.Whs=a.Sku set x.Sc_U=a.Sc_U"
WRun "Update [#EndOH] set EndOHSc=EndOH/Sc_U where Sc_U is not null"


WDrp "@Main" 'Sku ZHT1 BegOHSc BegRateSc BegAmt GRSc GRAmt OldRateSc EndAmt EndOHSc EndRateSc
             'Sku Whs
             'ZHT1 ZBrdNm ZBrd ZQlyNm ZQly Z8Nm Z8
             'EndAmt = EndOHSc
WRun "Select Distinct Whs,Sku,Sum(QUnRes+QBlk+QInsp) As OH into [@Main] from [#IMB52] Group by Whs,Sku"

'Des StkUom Sc_U OH_Sc
WRun "Alter Table [@Main] Add Column Des Text(255), StkUom Text(10),Sc_U Int, OH_Sc Double"
WRun "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku set x.Sc_U = a.Sc_U,x.Des=a.Des,x.StkUom=a.StkUom"
WRun "Update [@Main] set OH_Sc=OH/Sc_U where Sc_U>0"

'Stream ProdH F2 M32 M35 M37 Topaz ZHT1 RateSc Z2 Z5 Z7
WRun "Alter Table [@Main] add column Stream Text(10), Topaz Text(20), ProdH text(7), F2 Text(2), M32 text(2), M35 text(5), M37 text(7), ZHT1 Text(7), Z2 text(2), Z5 text(5), Z7 text(7), RateSc Currency, Amt Currency"

'ProdH Topaz
WRun "Update [@Main] x inner join [#IUom] a on x.Sku=a.Sku set x.ProdH=a.ProdH,x.Topaz=a.Topaz"

'F2 M32 M35 M37
WRun "Update [@Main] set F2=Left(ProdH,2),M32=Mid(ProdH,3,2),M35=Mid(ProdH,3,5),M37=Mid(ProdH,3,7)"

'ZHT1 RateSc
WRun "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M37=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"
WRun "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M35=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"
WRun "Update [@Main] x inner join [@Rate] a on x.Whs=a.Whs and x.M32=a.ZHT1 set x.RateSc=a.RateSc,x.ZHT1=a.ZHT1 where x.RateSc Is Null"

'Stream
WRun "Update [@Main] set Stream=IIf(Left(Topaz,3)='UDV','Diageo','MH')"

'Z2 Z5 Z7
WRun "Update [@Main] Set Z2=Left(ZHT1,2), Z5=Left(ZHT1,5), Z7=Left(ZHT1,7) where not ZHT1 is null"

'Amt
WRun "Update [@Main] Set Amt = RateSc * OH_Sc where RateSc is not null"
End Sub

Sub Oup()
MsgRunQry "@Rate": ORate
MsgRunQry "@Main": OMain
End Sub

Sub IMB52Opn()
FxOpn IFxMB52
End Sub

Sub IZHT1Opn()
FxOpn IFxZHT1
End Sub
Function IZHT1Fny() As String()
AyDmp DbtFny(W, ">ZHT1")
End Function
Function PmEndYYMM_Chk$(A$)

End Function
Function PmNMthGR_Chk$(A As Byte, EndYYMM$)

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