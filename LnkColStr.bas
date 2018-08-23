Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Property Get ZHT1$()
ZHT1 = _
" ZHT1   Txt Brand  |" & _
" RateSc Dbl Amount |" & _
" VdtFm  Txt [Valid From]  |" & _
" VdtTo  Txt [Valid to]"
End Property
Property Get Uom$()
Uom = _
 "Sku    Txt Material |" & _
 "Whs    Txt Plant |" & _
 "Des    Txt [Material Description] |" & _
 "Sc_U   Txt SC |" & _
 "StkUom Txt [Base Unit of Measure] |" & _
 "ProdH  Txt [Product hierarchy]"
End Property
Property Get MB52$()
MB52 = _
    " Sku    Txt Material |" & _
    " Whs    Txt Plant    |" & _
    " QInsp  Dbl [In Quality Insp#]|" & _
    " QUnRes Dbl UnRestricted|" & _
    " QBlk   Dbl Blocked"
'    " Loc    Txt [Storage Location] |" & _
'    " BchNo  Txt Batch |"
End Property
'>InvH: [Vendor] [InvNo] [Date] [Amt] [Sc]
'>InvD: [InvNo] [Sku] [Sc] [Amt]
'InvH: VndShtNm InvNo Whs Dte Sc Amt DteCrt
'InvD: VndShtNm InvNo Sku Sc Amt
Property Get InvH$()
InvH = _
    " VndShtNm Txt |" & _
    " InvNo    Txt |" & _
    " Dte      Dte InvDte|" & _
    " Whs      Txt Plant  |" & _
    " Sc       Dbl | " & _
    " Amt      Cur"
End Property
Property Get InvD$()
InvD = _
    " VndShtNm Txt |" & _
    " InvNo    Txt |" & _
    " Sku      Txt |" & _
    " Sc       Dbl |" & _
    " Amt      Cur "
End Property