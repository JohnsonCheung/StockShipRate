Option Compare Database
Option Explicit
Sub ZZ_IniLoad()
Y = 18
M = 2
IniLoad
End Sub
Sub IniLoad()
AssYM
WIni
WttLnkFb "YM YMZHT1", IFbStkShpRate
If Not SqlAny(FmtQQ("Select Y from YM where Y=? and M=?", Y, M)) Then
    MsgAp_Brw "Program Error: No record of [Y] and [M] in table YM", Y, M
    Exit Sub
End If
Dim A$(), B$()
WtLnkFx ">ZHT18601", IFxZHT1, "8601"
WtLnkFx ">ZHT18701", IFxZHT1, "8701"
A = WtChkCol(">ZHT18601", LnkColStr.ZHT1)
B = WtChkCol(">ZHT18701", LnkColStr.ZHT1)
If AyBrwEr(AyAdd(A, B)) Then Exit Sub
WImp ">ZHT18601", LnkColStr.ZHT1
WImp ">ZHT18701", LnkColStr.ZHT1

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

WRun "Delete from [YMZHT1]"
WQQ "Insert into [YMZHT1] (Y,M,ZHT1,Whs,RateSc,FmDte,ToDte) select ?,?,ZHT1,Whs,RateSc,FmDte,ToDte from [#Cpy]", Y, M

TmpRate_Upd_YM "#Cpy"
WDrp "#Cpy #Cpy1 #Cpy2"
Done
End Sub
Function IFxZHT1$()
IFxZHT1 = PnmFfn("ZHT1")
End Function

Sub IZHT1Opn()
FxOpn IFxZHT1
End Sub