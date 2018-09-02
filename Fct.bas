Option Compare Database
Option Explicit
Const C_Des$ = "Description"
Public Fso As New Scripting.FileSystemObject
Public AAA$()
Public Fcmd As New Fcmd
Public Const Z_ReSeqSpec$ = _
"Flg RecTy Amt Key Uom MovTy Qty BchRateUX RateTy Bch Las GL |" & _
" Flg IsAlert IsWithSku |" & _
" Key Sku PstMth PstDte |" & _
" Bch BchNo BchPermitDate BchPermit |" & _
" Las LasBchNo LasPermitDate LasPermit |" & _
" GL GLDocNo GLDocDte GLAsg GLDocTy GLLin GLPstKy GLPc GLAc GLBusA GLRef |" & _
" Uom Des StkUom Ac_U"
Type AttRs
    TblRs As DAO.Recordset
    AttRs As DAO.Recordset
End Type
Type XlsLnkInf
    IsXlsLnk As Boolean
    Fx As String
    WsNm As String
End Type
Enum EApp
    EDuty = 1
    EStkHld = 2
    EShpRate = 3
    EShpCst = 4
    ETaxCmp = 5
    ETaxAlert = 6
End Enum
Public Q$
Const PSep$ = " "
Const PSep1$ = " "
Public Actual, Expect
Public Schm As New Schm
Public EnsPrpOnEr As New EnsPrpOnEr
Private X_W As Database
Function FdScl_Fd(A$) As DAO.Field2
Dim J%, F$, L$, T$, Ay$(), Sz%, Des$, Rq As Boolean, Ty As DAO.DataTypeEnum, AlwZLen As Boolean, Dft$, VRul$, VTxt$
If A = "" Then Exit Function
Ay = AyRmvEmp(AyTrim(SplitSC(A)))
F = Ay(0)
T = Ay(1)
Ty = DaoShtTy_Ty(T)
For J = 2 To UB(Ay)
    L = Ay(J)
    Select Case True
    Case L = "Req": Rq = True
    Case L = "AlwZLen": AlwZLen = True
    Case HasPfx(L, "Sz="): Sz = RmvPfx(L, "Sz=")
    Case HasPfx(L, "Dft="): Dft = RmvPfx(L, "Dft=")
    Case HasPfx(L, "VRul="): VRul = RmvPfx(L, "VRul=")
    Case HasPfx(L, "VTxt="): VTxt = RmvPfx(L, "VTxt=")
    Case HasPfx(L, "Des="): Des = RmvPfx(L, "Des=")
    Case Else: Debug.Print "FdScl_Fd: there is itm[" & L & "] in FdScl[" & A & "] unexpected."
    End Select
Next
Dim O As New DAO.Field
With O
    .Name = F
    .DefaultValue = Dft
    .Required = Rq
    .Size = Sz
    .Type = Ty
    If Ty = DAO.DataTypeEnum.dbText Then
        .AllowZeroLength = AlwZLen
    End If
    .ValidationRule = VRul
    .ValidationText = VTxt
End With
Set O = FdScl_Fd
End Function
Function TdScly_AddPfx(A) As String()
Dim O$(), U&, J&, X
U = UB(A)
If U = -1 Then Exit Function
ReDim O(U)
For Each X In A
    O(J) = IIf(J = 0, "Td;", "Fd;") & X
    J = J + 1
Next
TdScly_AddPfx = O
End Function
Function DbScly(A As Database) As String()
DbScly = AySy(AyOfAy_Ay(AyMap(ItrMap(A.TableDefs, "TdScly"), "TdScly_AddPfx")))
End Function
Function TdScly(A As DAO.TableDef) As String()
TdScly = AyAdd(Sy(TdScl(A)), TdFdScly(A))
End Function
Function TdScl$(A As DAO.TableDef)
TdScl = A.Name
End Function
Function TdFdScly(A As DAO.TableDef) As String()
TdFdScly = ItrMapSy(A.Fields, "FdScl1")
End Function
Function ItrMapInto(A, Map$, OInto)
Dim O: O = OInto
Erase O
Dim X
For Each X In A
    Push O, Run(Map, X)
Next
ItrMapInto = O
End Function
Function ItrMapSy(A, Map$) As String()
ItrMapSy = ItrMapInto(A, Map, EmpSy)
End Function
Function NewFd_zFdScl(FdScl$) As DAO.Field2
Set NewFd_zFdScl = FdScl_Fd(FdScl)
End Function

Function BoolTxt$(A As Boolean, T$)
If A Then BoolTxt = T
End Function

Function AddLbl$(A, Lbl$)
If A <> "" Then AddLbl = Lbl & "=" & Replace(A, ";", "%3B")
End Function

Function FdScl1$(A As DAO.Field2)
Dim Rq$, Ty$, Sz$, ZLen$, Rul$, Dft$, VTxt$, Expr$, Des$
Des = AddLbl(FdDes(A), "Des")
Rq = BoolTxt(A.Required, "Req")
ZLen = BoolTxt(A.AllowZeroLength, "AlwZLen")
Ty = DaoTy_ShtTy(A.Type)
Sz = BoolTxt(A.Type = dbText, "Sz=" & A.Size)
Rul = AddLbl(A.ValidationText, "VTxt")
VTxt = AddLbl(A.ValidationRule, "VRul")
Expr = AddLbl(A.Expression, "Expr")
Dft = AddLbl(A.DefaultValue, "Dft")
FdScl1 = ApScl(A.Name, Ty, Sz, Rq, ZLen, Rul, VTxt, Dft, Expr)
End Function

Function MdNm$(A As CodeModule)
If IsNothing(A) Then Exit Function
MdNm = A.Parent.Name
End Function

Function MdLno_zExitPrp%(A As CodeModule, PrpLno)
If HasSfx(A.Lines(PrpLno, 1), "End Property") Then Exit Function
Dim J%, L$
For J = PrpLno + 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If HasPfx(L, "Exit Property") Then MdLno_zExitPrp = J: Exit Function
    If HasPfx(L, "End Property") Then Exit Function
Next
Stop
End Function

Function MdLno_zEndPrp%(A As CodeModule, PrpLno)
If HasSfx(A.Lines(PrpLno, 1), "End Property") Then MdLno_zEndPrp = PrpLno: Exit Function
Dim J%
For J = PrpLno + 1 To A.CountOfLines
    If HasPfx(A.Lines(J, 1), "End Property") Then MdLno_zEndPrp = J: Exit Function
Next
Stop
End Function
Function StrInLikSsl(A, LikSsl) As Boolean
StrInLikSsl = StrInLikAy(A, SslSy(LikSsl))
End Function
Function AyItr(A) As Collection
Dim O As New Collection
If Sz(A) > 0 Then
    Dim X
    For Each X In A
        O.Add X
    Next
End If
Set AyItr = O
End Function
Property Get W() As Database
If IsNothing(X_W) Then WEns: WOpn
Set W = X_W
End Property
Function ApLin$(ParamArray Ap())
Dim Av(): Av = Ap
ApLin = JnSpc(AyRmvEmp(Av))
End Function
Function ApScl$(ParamArray Ap())
Dim Av(): Av = Ap
ApScl = JnSC(AyRmvEmp(Av))
End Function

Function AyHasPredXPTrue(A, XP$, P) As Boolean
If Sz(A) = 0 Then Exit Function
Dim X
For Each X In A
    If Run(XP, X, P) Then
        AyHasPredXPTrue = True
        Exit Function
    End If
Next
End Function

Function LinHasT1(A, T1) As Boolean
LinHasT1 = LinT1(A) = T1
End Function

Function LinHasT2(A, T2) As Boolean
LinHasT2 = LinT2(A) = T2
End Function

Function LinT2$(A)
LinT2 = LinT1(LinRmvT1(A))
End Function

Function LinInT1Ay(A, T1Ay$())
LinInT1Ay = AyHas(T1Ay, LinT1(A))
End Function

Function NewFd_zId(F) As DAO.Field2
Dim O As New DAO.Field
With O
    .Name = F
    .Type = dbLong
    .Attributes = DAO.FieldAttributeEnum.dbAutoIncrField
    .Required = True
End With
Set NewFd_zId = O
End Function
Function NewFd_zFk(F) As DAO.Field2
Dim O As New DAO.Field
With O
    .Name = F
    .Type = dbLong
End With
Set NewFd_zFk = O
End Function

Function LinRmvT1$(A)
Dim O$: O = A
LinShiftTerm O
LinRmvT1 = O
End Function

Function LonmXReSeq$(A)
LonmXReSeq = TfkV("LoReSeq", "ReSeq", A)
End Function

Function AyWhT1EqV(A, V) As String()
AyWhT1EqV = AyWhPredXP(A, "LinHasT1", V)
End Function

Function AyWhT2EqV(A$(), V) As String()
AyWhT2EqV = AyWhPredXP(A, "LinHasT2", V)
End Function

Function Sy(ParamArray Ap())
Dim Av(): Av = Ap
Sy = AySy(Av)
End Function

Sub WtReSeq(T, ReSeqSpec$)
DbtReSeq W, T, ReSeqSpec
End Sub

Sub ZZ_TFDes()
TFDes("Att", "AttNm") = "AttNm"
End Sub
Property Get DbtfDes$(A As Database, T, F)
DbtfDes = DbtfPrp(A, T, F, C_Des)
End Property
Property Let DbtfDes(A As Database, T, F, Des$)
DbtfPrp(A, T, F, C_Des) = Des
End Property
Property Get DbtfPrp(A As Database, T, F, P)
If Not DbtfHasPrp(A, T, F, P) Then Exit Property
DbtfPrp = A.TableDefs(T).Fields(F).Properties(P).Value
End Property
Function PrpHas(A As DAO.Properties, P) As Boolean
PrpHas = ItrHasNm(A, P)
End Function
Function FdDes$(A As DAO.Field)
If PrpHas(A.Properties, C_Des) Then FdDes = A.Properties(C_Des)
End Function
Property Let DbtfPrp(A As Database, T, F, P, V)
If DbtfHasPrp(A, T, F, P) Then
    A.TableDefs(T).Fields(F).Properties(P).Value = V
Else
    With A.TableDefs(T)
        .Fields(F).Properties.Append .CreateProperty(P, VarDaoTy(V), V)
    End With
End If
End Property

Property Get TFDes$(T, F)
TFDes = TFldPrp(T, F, C_Des)
End Property
Property Let TFDes(T, F, Des$)
TFldPrp(T, F, C_Des) = Des
End Property
Property Get TFldPrp(T, F, P)
TFldPrp = DbtfPrp(CurrentDb, T, F, P)
End Property
Property Let TFldPrp(T, F, P, V)
DbtfPrp(CurrentDb, T, F, P) = V
End Property
Function DbtfHasPrp(A As Database, T, F, P) As Boolean
DbtfHasPrp = ItrHasNm(A.TableDefs(T).Fields(F).Properties, P)
End Function
Function SslAy_Sy(A$()) As String()
Dim O$(), L
If Sz(A) = 0 Then Exit Function
For Each L In A
    PushAy O, SslSy(L)
Next
SslAy_Sy = O
End Function

Function HasSubStr(A, SubStr) As Boolean
HasSubStr = InStr(A, SubStr) > 0
End Function
Function TFLinHasPk(A) As Boolean
TFLinHasPk = HasSubStr(A, " * ")
End Function
Function TFLinHasSk(A) As Boolean
TFLinHasSk = HasSubStr(A, " | ")
End Function
Function LyT1Ay(A) As String()
Dim O$(), L, J&
If Sz(A) = 0 Then Exit Function
ReDim O(UB(A))
For Each L In A
    BrkAsg L, " ", O(J)
    J = J + 1
Next
End Function
Function AyabMapInto(A, B, MapAB$, OInto)
Dim J&, U&, O
O = OInto: Erase O
U = UB(A)
If U <> UB(B) Then Stop
For J = 0 To U
    Push O, Run(MapAB, A(J), B(J))
Next
AyabMapInto = O
End Function
Function AyabMapSy(A, B, MapAB$) As String()
AyabMapSy = AyabMapInto(A, B, MapAB, EmpSy)
End Function
Function TnPkSql$(A)
TnPkSql = FmtQQ("Create Index PrimaryKey on [?] (?) with Primary", A, A)
End Function
Function TnSkSql$(A, SkNy0)
TnSkSql = FmtQQ("Create Unique Index [?] on [?] (?)", A, A, JnComma(CvNy(SkNy0)))
End Function
Sub LinesBrkAsg1(A, OErLy$(), ORmkDic As Dictionary, Ny0, ParamArray OLyAp())
Dim Ny$(), L, L1$, T1$, T2$, NmDic As Dictionary, Rmk$(), Ix%
Ny = CvNy(Ny0)
Set NmDic = AyIxDic(Ny)
For Each L In SplitCrLf(A)
    L1 = LTrim(L)
    Select Case FstChr(L)
    Case "'"
        Push Rmk, L
    Case ""
    Case Else
        BrkS1Asg L1, " ", T1, T2
        If NmDic.Exists(T1) Then
            Ix = NmDic(T1)
            Push OLyAp(Ix), T2 '<----
            If Sz(Rmk) > 0 Then
                ORmkDic.Add T1 & "." & UB(OLyAp(Ix)), Rmk
                Erase Rmk
            End If
        Else
            Push OErLy, L
        End If
    End Select
Next
End Sub

Sub LinesBrkAsg(A, Ny0, ParamArray OLyAp())
Dim Ny$(), L, T1$, T2$, NmDic As Dictionary
Ny = CvNy(Ny0)
Set NmDic = AyIxDic(Ny)
For Each L In SplitCrLf(A)
    Select Case FstChr(L)
    Case "'", " "
    Case Else
        BrkAsg L, " ", T1, T2
        If NmDic.Exists(T1) Then
            Push OLyAp(NmDic(T1)), T2 '<----
        End If
    End Select
Next
End Sub
Function AyIxDic(A) As Dictionary
Dim O As New Dictionary, J&
For J = 0 To UB(A)
    O.Add A(J), J
Next
Set AyIxDic = O
End Function
Sub NDriveMap()
NDriveRmv
Shell "Subst N: c:\users\user\desktop\MHD"
End Sub
Sub NDriveRmv()
Shell "Subst /d N:"
End Sub
Function EAppStr_DtaFb$(A)
Dim App As EApp
If Not IsEAppStr(A) Then Exit Function
EAppStr_DtaFb = EAppDtaFb(EAppStr_EApp(A))
End Function
Function IsEAppStr(A) As Boolean
Select Case A
Case _
"Duty", _
"SkHld", _
"ShpRate", _
"ShpCst", _
"TaxCmp", _
"TaxAlert"
IsEAppStr = True
End Select
End Function
Property Get AppRoot$()
AppRoot = "N:\SAPAccessReports\"
End Property
Property Get AppHom$()
AppHom = AppRoot & Apn & "\"
End Property
Function EAppFdr$(A As EApp)
Dim O$
Select Case A
Case EApp.EDuty: O = "DutyPrepay"
Case EApp.EStkHld: O = "StockHolding"
Case EApp.EShpRate: O = "StockShipRate"
Case EApp.EShpCst: O = "StockShipCost"
Case EApp.ETaxCmp: O = "TaxExpCmp"
Case EApp.ETaxAlert: O = "TaxRateAlert"
Case Else: Stop
End Select
EAppFdr = O
End Function
Function EAppDtaFn$(A As EApp)
Dim O$
Select Case A
Case EApp.EShpRate: O = "StockShipRate_Data.accdb"
Case Else: Stop
End Select
EAppDtaFn = O
End Function
Function EAppPth$(A As EApp)
EAppPth = AppRoot & EAppFdr(A) & "\"
End Function
Function EAppDtaFb$(A As EApp)
EAppDtaFb = EAppPth(A) & EAppDtaFn(A)
End Function
Function EAppStr_EApp(A) As EApp
Dim O As EApp
Select Case A
Case "Duty": O = EApp.EDuty
Case "StkHld": O = EApp.EStkHld
Case "ShpRate": O = EApp.EShpRate
Case "ShpCst": O = EApp.EShpCst
Case "TaxCmp": O = EApp.ETaxCmp
Case "TaxAlert": O = EApp.ETaxAlert
Case Else: Stop
End Select
EAppStr_EApp = O
End Function
Function EAppStr$(A As EApp)
Dim O$
Select Case A
Case EApp.EDuty: O = "Duty"
Case EApp.EStkHld: O = "StkHld"
Case EApp.EShpRate: O = "ShpRate"
Case EApp.EShpCst: O = "ShpCst"
Case EApp.ETaxCmp: O = "TaxCmp"
Case EApp.ETaxAlert: O = "TaxAlert"
Case Else: Stop
End Select
EAppStr = O
End Function
Property Get CcmTny() As String()
CcmTny = DbCcmTny(CurrentDb)
End Property
Sub SetCcmTblDes(Des$)
Dim T
For Each T In CcmTny
    TblDes(T) = Des
Next
End Sub

Function DbCcmTny(A As Database) As String()
DbCcmTny = AyWhHasPfx(DbTny(A), "^")
End Function

Property Get CnSy() As String()
CnSy = CDbCnSy
End Property
Function AyabNonEmpBLy(A, B, Optional Sep$ = " ") As String()
Dim J&, O$()
For J = 0 To UB(A)
    If Not IsEmp(B(J)) Then
        Push O, A(J) & Sep & B(J)
    End If
Next
AyabNonEmpBLy = O
End Function
Function DbCnSy(A As Database) As String()
Dim T$(), S()
T = AyQuoteSqBkt(DbTny(A))
S = AyMapPX(T, "DbtCnStr", A)
DbCnSy = AyabNonEmpBLy(T, S)
End Function

Function AyMapPX(A, PX$, P)
AyMapPX = AyMapPXInto(A, PX, P, EmpAy)
End Function
Function AyMapPXSy(A, PX$, P) As String()
AyMapPXSy = AyMapPXInto(A, PX, P, EmpSy)
End Function

Function AyMapPXInto(A, PX$, P, OInto)
Dim O: O = OInto
Erase O
If Sz(A) > 0 Then
    Dim J&, X
    ReDim O(UB(A))
    For Each X In A
        Asg Run(PX, P, X), O(J)
        J = J + 1
    Next
End If
AyMapPXInto = O
End Function
Function DbtSrc$(A As Database, T)
DbtSrc = A.TableDefs(T).SourceTableName
End Function
Function RsTSz$(A As DAO.Recordset)
If A.Fields(0).Type <> DAO.dbDate Then Stop
If A.Fields(1).Type <> DAO.dbLong Then Stop
RsTSz = DteDTim(A.Fields(0).Value) & "." & A.Fields(1).Value
End Function
Function JnVBar$(A)
JnVBar = Join(A, "|")
End Function

Function LonmXWdt(A) As String()
'From Table LoColWdt with fields LoNm Wdt FldLikSsl Seq
'Return {Wdt} {FldLikSsl...}
'       ..
Q = FmtQQ("Select Wdt & ' ' & FldLikSsl from [LoWdt] where LoNm='?' or LoNm='*' order by LoNm,Seq", A)
LonmXWdt = RsSy(CurrentDb.OpenRecordset(Q))
End Function

Function LonmXAlignC$(A)
'From Table LoAlignC with fields FldLikSsl
'Return {FldLikSsl..}
Q = FmtQQ("Select FldLikSsl from [LoAlignC] where LoNm in ('?','*')", A)
LonmXAlignC = JnSpc(RsSy(CurrentDb.OpenRecordset(Q)))
End Function

Function LonmXTSum$(A)
'From Table LoAlignC with fields FldLikSsl
'Return {FldLikSsl..}
Q = FmtQQ("Select FldLikSsl from [LoTSum] where LoNm in ('?','*')", A)
LonmXTSum = JnSpc(RsSy(CurrentDb.OpenRecordset(Q)))
End Function

Function LonmXTCnt$(A)
'From Table LoAlignC with fields FldLikSsl
'Return {FldLikSsl..}
Q = FmtQQ("Select FldLikSsl from [LoTCnt] where LoNm in ('?','*')", A)
LonmXTCnt = JnSpc(RsSy(CurrentDb.OpenRecordset(Q)))
End Function

Function LonmXTAvg$(A)
'From Table LoAlignC with fields FldLikSsl
'Return {FldLikSsl..}
Q = FmtQQ("Select FldLikSsl from [LoTAvg] where LoNm in ('?','*')", A)
LonmXTAvg = JnSpc(RsSy(CurrentDb.OpenRecordset(Q)))
End Function

Function LonmXFmt(A) As String()
'From Table LoFmt with fields LoNm Fmt FldLikSsl Seq
'Return {Fmt} {FldLikSsl..}
'       ..
Q = FmtQQ("Select Fmt & ' ' & FldLikSsl from [LoFmt] where LoNm='?' or LoNm='*' order by LoNm,Seq", A)
LonmXFmt = RsSy(CurrentDb.OpenRecordset(Q))
End Function

Function QQRs(QQSql, ParamArray Ap()) As DAO.Recordset
Dim Av(): Av = Ap
Set QQRs = DbqRs(CurrentDb, FmtQQAv(QQSql, Av))
End Function

Function QQV(QQSql, ParamArray Ap())
Dim Av(): Av = Ap
QQV = DbqV(CurrentDb, FmtQQAv(QQSql, Av))
End Function

Function LonmTblNm$(A)
If Not HasPfx(A, "T_") Then Stop
LonmTblNm = "@" & Mid(A, 3)
End Function
Sub LoRfhAllFmt(A As ListObject)
LoRfhFml A
LoRfhFmt A
LoRfhAlignC A
LoRfhWdt A
LoRfhTot A
End Sub

Sub LoRfhAlignC(A As ListObject)
LoSetAlignC_zX A, LonmXAlignC(A.Name)
End Sub

Sub LoRfhWdt(A As ListObject)
LoSetWdt_zX A, LonmXWdt(A.Name)
End Sub

Sub LoRfhTot(A As ListObject)
LoSetTot_zX A, LonmXTSum(A.Name), xlTotalsCalculationSum
LoSetTot_zX A, LonmXTAvg(A.Name), xlTotalsCalculationAverage
LoSetTot_zX A, LonmXTCnt(A.Name), xlTotalsCalculationCount
End Sub

Sub LoRfhFmt(A As ListObject)
LoSetFmt_zX A, LonmXFmt(A.Name)
End Sub

Sub LoRfhFml(A As ListObject)
LoSetFml_zX A, LonmXFml(A.Name)
End Sub
Function LcnmFmt$(A, XFmt$())
If Sz(XFmt) = 0 Then Exit Function
Dim F, Fmt$, FldNmLikSsl$, LikAy$()
For Each F In XFmt
    LinAsgT1Rest F, Fmt, FldNmLikSsl
    LikAy = SslSy(FldNmLikSsl)
    If StrInLikAy(A, LikAy) Then
        LcnmFmt = Fmt
        Exit Function
    End If
Next
End Function
Function LcnmWdt%(A, XWdt$())
If Sz(XWdt) = 0 Then Exit Function
Dim W, Wdt%, FldNmLikSsl$, LikAy$()
For Each W In XWdt
    LinAsgT1Rest W, Wdt, FldNmLikSsl
    LikAy = SslSy(FldNmLikSsl)
    If StrInLikAy(A, LikAy) Then
        LcnmWdt = Wdt
        Exit Function
    End If
Next
End Function

Function StrInLikAy(A, LikAy$()) As Boolean
StrInLikAy = LikAy_HasNm(LikAy, A)
End Function
Sub LcSetFmt(A As ListColumn, F$)
If F = "" Then Exit Sub
A.DataBodyRange.NumberFormat = F
End Sub
Sub LcSetWdt(A As ListColumn, W%)
If W <= 0 Then Exit Sub
A.DataBodyRange.ColumnWidth = W
End Sub
Sub LoSetFmt_zX(A As ListObject, XFmt$())
'XFmt is {Fmt} {FldLikSsl..}
'        ..
If Sz(XFmt) = 0 Then Exit Sub
Dim C As ListColumn
For Each C In A.ListColumns
    LcSetFmt C, LcnmFmt(C.Name, XFmt)
Next
End Sub

Function LonmXFml(A) As String()
'From Table LoFml with fields LoNm FldNm Fml
'Return {FldNm} {Fml}
'       ..
Q = FmtQQ("Select FldNm & ' ' & Fml from [LoFml] where LoNm='?'", A)
LonmXFml = RsSy(CurrentDb.OpenRecordset(Q))
End Function

Sub ZZ_DbtPrp()
TblDrp "Tmp"
DoCmd.RunSQL "Create Table Tmp (F1 Text)"
DbtPrp(CurrentDb, "Tmp", "XX") = "AFdf"
Debug.Assert DbtPrp(CurrentDb, "Tmp", "XX") = "AFdf"
End Sub
Property Get DbtPrpLoFmlVbl$(A As Database, T)
DbtPrpLoFmlVbl = DbtPrp(A, T, "LoFmlVbl")
End Property
Property Get TblPrpLoFmlVbl$(T)
TblPrpLoFmlVbl = DbtPrpLoFmlVbl(CurrentDb, T)
End Property
Property Let TblPrpLoFmlVbl(T, LoFmlVbl$)
DbtPrpLoFmlVbl(CurrentDb, T) = LoFmlVbl
End Property
Property Let DbtPrpLoFmlVbl(A As Database, T, LoFmlVbl$)
DbtPrp(A, T, "LoFmlVbl") = LoFmlVbl
End Property
Property Get DbtPrp(A As Database, T, P)
If Not DbtHasPrp(A, T, P) Then Exit Property
DbtPrp = A.TableDefs(T).Properties(P).Value
End Property
Function VarDaoTy(A) As DAO.DataTypeEnum
Dim O As DAO.DataTypeEnum
Select Case VarType(A)
Case VbVarType.vbInteger: O = dbInteger
Case VbVarType.vbLong: O = dbLong
Case VbVarType.vbString: O = dbText
Case VbVarType.vbDate: O = dbDate
Case Else: Stop
End Select
VarDaoTy = O

End Function
Function CvFld2(A As DAO.Field) As DAO.Field2
Set CvFld2 = A
End Function
Function DbtfVal(A As Database, T, F)
DbtfVal = A.TableDefs(T).OpenRecordset.Fields(F).Value
End Function
Function DbtfkV(A As Database, T, F, K())
Dim W$, Sk$(), Rs As DAO.Recordset
Sk = DbtSk(A, T)
W = KyVy_BExpr(Sk, K)
Q = FmtQQ("Select ? from [?] where ?", F, T, W)
Set Rs = A.OpenRecordset(Q)
DbtfkV = RsV(Rs, F)
End Function
Function KyVy_BExpr$(Ky$(), Vy())
Dim U%, S$
U = UB(Ky)
If U <> UB(Vy) Then Stop
Dim O$(), J%, V
For J = 0 To U
    If IsNull(Vy(J)) Then
        Push O, Ky(J) & " is null"
    Else
        V = Vy(J): GoSub X
        Push O, Ky(J) & "=" & S
    End If
Next
KyVy_BExpr = Join(O, " and ")
Exit Function
X:
Select Case True
Case IsStr(V): S = "'" & V & "'"
Case IsDte(V): S = "#" & V & "#"
Case IsBool(V): S = IIf(V, "TRUE", "FALSE")
Case IsNull(V): Stop
Case IsNumeric(V): S = V
Case Else: Stop
End Select
Return
End Function
Function TfidV(T, F, Id&)
TfidV = DbtfidV(CurrentDb, T, F, Id)
End Function
Function DbtidRs(A As Database, T, Id&) As DAO.Recordset
Q = FmtQQ("Select * From ? where ?=?", T, T, Id)
Set DbtidRs = A.OpenRecordset(Q)
End Function
Function DbtfidV(A As Database, T, F, Id&)
DbtfidV = DbtidRs(A, T, Id).Fields(F).Value
End Function
Sub DbtfAddExpr(A As Database, T, F, Expr$, Optional Ty As DAO.DataTypeEnum = dbText, Optional TxtSz% = 255)
A.TableDefs(T).Fields.Append NewFd(F, Ty, TxtSz, Expr)
End Sub
Function DicKey_Asg(A As Dictionary, K, O) As Boolean
If A.Exists(K) Then
    O = A(K)
    DicKey_Asg = True
End If
End Function
Sub TdAddFd(A As DAO.TableDef, F As DAO.Field)
A.Fields.Append F
End Sub

Function NewFd(F, Optional Ty As DAO.DataTypeEnum = dbText, Optional TxtSz% = 255, Optional Expr$, Optional Dft$, Optional Req As Boolean, Optional VRul$, Optional VTxt$) As DAO.Field2
Dim O As New DAO.Field
With O
    .Name = F
    .Required = Req
    If Ty <> 0 Then .Type = Ty
    If Ty = dbText Then
        .Size = TxtSz
    End If
    If Expr <> "" Then
        CvFld2(O).Expression = Expr
    End If
    O.DefaultValue = Dft
End With
Set NewFd = O
End Function
Function DaoTbl(T) As DAO.TableDef
Dim O As New DAO.TableDef
With O
    .Name = T
    .Fields.Append NewFd("F1")
End With
Set DaoTbl = O
End Function
Sub ZZ_DbtfAddExpr()
TblDrp "Tmp"
Dim A As DAO.TableDef
Set A = DbAddTbl(CurrentDb, "Tmp")
DbtfAddExpr CurrentDb, "Tmp", "F2", "[F1]+"" hello!"""
TblDrp "Tmp"
End Sub
Sub ZZ_DbAddTbl()
Dim A As DAO.TableDef
TblDrp "Tmp"
Set A = DbAddTbl(CurrentDb, "Tmp")
TblDrp "Tmp"
End Sub
Function DbAddTbl(A As Database, T) As DAO.TableDef
Dim O As DAO.TableDef
Set O = DaoTbl(T)
A.TableDefs.Append O
Set DbAddTbl = O
End Function
Function DbtAddFld(A As Database, T, F, Optional Ty As DAO.DataTypeEnum = dbText, Optional TxtSz% = 255) As DAO.Field2
Dim O As DAO.Field2
Set O = NewFd(F, Ty, TxtSz)
A.TableDefs(T).Fields.Append O
Set DbtAddFld = O
End Function
Function FldPrpNy(A As DAO.Field) As String()
FldPrpNy = ItrNy(A.Properties)
End Function
Sub AyPushMsgAv(A, Msg$, Av())
PushAy A, MsgAv_Ly(Msg, Av)
End Sub
Function YYMM_FstDte(A) As Date
YYMM_FstDte = DateSerial(Left(A, 2), Mid(A, 3, 2), 1)
End Function
Function YM_LasDte(Y As Byte, M As Byte) As Date
YM_LasDte = DteNxtMth(YM_FstDte(Y, M))
End Function
Function NowStr$()
NowStr = Format(Now, "YYYY-MM-DD HH:MM:SS")
End Function
Function DteFstDayOfMth(A As Date) As Date
DteFstDayOfMth = DateSerial(Year(A), Month(A), 1)
End Function
Function DteNxtMth(A As Date) As Date
DteNxtMth = DateTime.DateAdd("M", 1, A)
End Function
Function YM_FstDte(Y As Byte, M As Byte) As Date
YM_FstDte = DateSerial(2000 + Y, M, 1)
End Function
Function DteYYMM$(A As Date)
DteYYMM = Right(Year(A), 2) & Format(Month(A), "00")
End Function
Property Get Apn$()
Static X As Boolean, Y$
If Not X Then
    X = True
    Y = SqlV("Select Apn from [Apn]")
End If
Apn = Y
End Property
Sub TmpHomBrw()
PthBrw TmpHom
End Sub
Sub WBrw()
AcsVis WAcs
End Sub
Sub WCls()
On Error Resume Next
X_W.Close
Set X_W = Nothing
End Sub
Sub WDrp(TT)
DbDrpTT W, TT
End Sub
Function PnmFfn$(A)
PnmFfn = PnmPth(A) & PnmFn(A)
End Function
Function PnmPth$(A)
PnmPth = PthEnsSfx(PnmVal(A & "Pth"))
End Function
Sub QQRun(QQSql, ParamArray Ap())
Dim Av(): Av = Ap
DoCmd.RunSQL FmtQQAv(QQSql, Av)
End Sub
Function QQAny(QQSql, ParamArray Ap()) As Boolean
Dim Av(): Av = Ap
QQAny = SqlAny(FmtQQAv(QQSql, Av))
End Function
Sub WtfAddExpr(T, F, Expr$)
DbtfAddExpr W, T, F, Expr
End Sub
Sub WReOpn()
WCls
WOpn
End Sub
Function RgRCC(A As Range, R, C1, C2) As Range
Set RgRCC = RgRCRC(A, R, C1, R, C2)
End Function

Sub ZZ_WtLnkFx()
WtLnkFx ">UOM", IFxUom
End Sub
Sub LcSetFml(A As ListColumn, Fml$)
A.DataBodyRange.Formula = Fml
End Sub
Sub LoSetFml_zX(A As ListObject, XFml$())
If Sz(XFml) = 0 Then Exit Sub
Dim ColNm$, Fml$, I
For Each I In XFml
    LinAsgT1Rest I, ColNm, Fml
    LcSetFml A.ListColumns(ColNm), "=" & Fml
Next
End Sub
Function StrInSfxAy(A, SfxAy$()) As Boolean
Dim Sfx
For Each Sfx In SfxAy
    If HasSfx(A, Sfx) Then StrInSfxAy = True: Exit Function
Next
End Function
Sub ZZ_LinShiftT1()
Dim L$, A$
L = " S   DFKDF SLDF  "
A = LinShiftT1(L)
Debug.Assert A = "S"
Debug.Assert L = "DFKDF SLDF"
End Sub
Function LinShiftT1$(OLin)
Dim T$, R$
BrkS1Asg LTrim(OLin), " ", T, R
LinShiftT1 = T
OLin = R
End Function
Sub LinAsgT1Rest(A, OT1, ORest)
BrkAsg Trim(A), " ", OT1, ORest
End Sub
Sub BrkS1Asg(A, Sep$, Optional O1, Optional O2, Optional NoTrim As Boolean)
BrkS1AtAsg A, InStr(A, Sep), Sep, O1, O2, NoTrim
End Sub
Sub BrkAsg(A, Sep$, Optional O1, Optional O2, Optional NoTrim As Boolean)
BrkAtAsg A, InStr(A, Sep), Sep, O1, O2, NoTrim
End Sub
Sub BrkAtAsg(A, At&, Sep$, O1, O2, Optional NoTrim As Boolean)
If At = 0 Then
    MsgBrw "[Str] does not have [Sep].  @BrkAtAsg.", A, Sep
    Stop
    Exit Sub
End If
O1 = Left(A, At - 1)
O2 = Mid(A, At + Len(Sep))
If Not NoTrim Then
    O1 = Trim(O1)
    O2 = Trim(O2)
End If
End Sub
Sub BrkS1AtAsg(A, At&, Sep$, O1, O2, Optional NoTrim As Boolean)
If At = 0 Then
    O1 = A
    O2 = ""
    Exit Sub
End If
O1 = Left(A, At - 1)
O2 = Mid(A, At + Len(Sep))
If Not NoTrim Then
    O1 = Trim(O1)
    O2 = Trim(O2)
End If
End Sub
Sub WtRen(Fmt$, ToT$, Optional ReOpnFst As Boolean)
DbtRen W, Fmt, ToT, ReOpnFst
End Sub


Sub WClr()
Exit Sub
Dim T, Tny$()
Tny = WTny: If Sz(Tny) = 0 Then Exit Sub
For Each T In Tny
    WDrp T
Next
End Sub
Function WTny() As String()
WTny = DbTny(W)
End Function
Function WStru$(Optional TT$)
If TT = "" Then
    WStru = DbStru(W)
Else
    WStru = DbttStru(W, TT)
End If
End Function
Function WAcs() As Access.Application
Set WAcs = ApnAcs(Apn)
End Function
Function WtFny(T$) As String()
WtFny = DbtFny(W, T)
End Function
Function WtStru$(T$)
WtStru = DbtStru(W, T)
End Function

Function WttStru$(TT)
WttStru = DbttStru(W, TT)
End Function
Function WFb$()
WFb = ApnWFb(Apn)
End Function
Sub WImp(T$, LnkColStr$, Optional WhBExpr$)
If FstChr(T) <> ">" Then Stop
DbtImpMap W, T, LnkColStr, WhBExpr
End Sub

Sub FfnMov(Fm, ToFfn)
Fso.MoveFile Fm, ToFfn
End Sub
Sub RsDmp(A As Recordset)
AyDmp RsCsvLy(A)
A.MoveFirst
End Sub
Sub RsDmpByFny0(A As Recordset, Fny0)
AyDmp RsCsvLyByFny0(A, Fny0)
A.MoveFirst
End Sub
Function AttRs(A) As AttRs
AttRs = DbAtt_AttRs(CurrentDb, A)
End Function
Function AttFny() As String()
AttFny = ItrNy(DbFstAttRs(CurrentDb).AttRs.Fields)
End Function
Function RsV(A As DAO.Recordset, Optional F = 0)
If A.EOF Then Exit Function
RsV = A.Fields(F).Value
End Function
Function AttRs_Exp$(A As AttRs, ToFfn)
'Export the only File in {AttRs} {ToFfn}
Dim Fn$, Ext$, T$, F2 As DAO.Field2
With A.AttRs
    If FfnExt(CStr(!FileName)) <> FfnExt(ToFfn) Then Stop
    Set F2 = .Fields("FileData")
End With
F2.SaveToFile ToFfn
AttRs_Exp = ToFfn
End Function
Function DbAtt_Exp$(A As Database, Att, ToFfn)
'Exporting the first File in Att.
'If no file in att, error
'If any, export and return the
Dim N%
N = DbAtt_FilCnt(A, Att)
If N <> 1 Then
    Er "[Att] in [Db] has [FilCnt] which should be one.|Not export to [ToFfn].  (@DbAtt_Exp)", _
        Att, A.Name, N, ToFfn
End If
DbAtt_Exp = AttRs_Exp(DbAtt_AttRs(A, Att), ToFfn)
FunMsgDmp "DbAtt_Exp", "[Att] is exported [ToFfn] from [Db]", Att, ToFfn, DbNm(A)
End Function
Function RsLy(A As DAO.Recordset, Optional Sep$ = " ") As String()
Dim O$()
With A
    Push O, Join(RsFny(A), Sep)
    While Not .EOF
        Push O, RsLin(A, Sep)
        .MoveNext
    Wend
End With
RsLy = O
End Function

Sub RaiseErr()
Err.Raise -1, , "Please check messages opened in notepad"
End Sub

Sub Er(Msg$, ParamArray Ap())
Dim Av(): Av = Ap
Dim O$()
O = MsgAv_Ly(Msg, Av)
AyBrw O
RaiseErr
End Sub

Function MsgNy(A) As String()
Dim O$(), P%, J%
O = Split(A, "[")
AyShift O
For J = 0 To UB(O)
    P = InStr(O(J), "]")
    O(J) = "[" & Left(O(J), P)
Next
MsgNy = O
End Function
Sub Ny0LyDmp(A, ParamArray Ap())
Dim Av(): Av = Ap
D NyLy(CvNy(A), Av, 0)
End Sub
Function NyLy(A$(), Av(), Optional Indent% = 4) As String()
NyLy = NyAv_Ly(A, Av, Indent)
End Function

Function NyLin$(A$(), Av())
NyLin = NyAv_Lin(A, Av)
End Function

Sub AyabSetSamMax(A, B)
Dim U1&, U2&
U1 = UB(A)
U2 = UB(B)
Select Case True
Case U1 > U2: ReDim Preserve B(U1)
Case U1 < U2: ReDim Preserve A(U2)
End Select
End Sub

Function NyAv_Ly(A$(), Av(), Optional Indent% = 4) As String()
Dim W%, O$(), J%, A1$(), A2$()
W = AyWdt(A)
A1 = AyAlignL(A)
A2 = AyAddSfx(A1, " : ")
AyabSetSamMax A2, Av
For J = 0 To UB(A)
    PushAy O, NmV_Ly(A2(J), Av(J))
Next
NyAv_Ly = AyAddPfx(O, Space(Indent))
End Function
Function NyAv_Lin$(A$(), Av())
Dim U&
U = UB(A)
If U = -1 Then Exit Function
Dim O$(), J%
For J = 0 To U
    Push O, NmV_Lin(A(J), Av(J))
Next
NyAv_Lin = Join(AyAddPfx(O, " | "))
End Function
Function EnsSfxDot$(A)
EnsSfxDot = EnsSfx(A, ".")
End Function
Function EnsSfxSC$(A)
EnsSfxSC = EnsSfx(A, ";")
End Function
Function EnsSfx$(A, Sfx$)
If HasSfx(A, Sfx) Then EnsSfx = A: Exit Function
EnsSfx = A & Sfx
End Function
Function NmV_Ly(Nm$, V) As String()
Dim O$(), S$, J%
O = VarLy(V)
If Sz(O) = 0 Then
    NmV_Ly = ApSy(Nm)
    Exit Function
End If
O(0) = Nm & O(0)
S = Space(Len(Nm))
For J = 1 To UB(O)
    O(J) = S & O(J)
Next
NmV_Ly = O
End Function
Function NmV_Lin$(Nm$, V)
NmV_Lin = Nm & "=[" & VarLin(V) & "]"
End Function
Function VarLines$(V)
VarLines = JnCrLf(VarLy(V))
End Function
Function IsItr(A) As Boolean
IsItr = TypeName(A) = "Collection"
End Function
Function ItrSy(A) As String()
Dim O$(), X
For Each X In A
    Push O, CStr(X)
Next
ItrSy = O
End Function
Function VarLy(V) As String()
Select Case True
Case IsItr(V):     VarLy = ItrSy(V)
Case IsStr(V):     VarLy = SplitCrLf(V)
Case IsPrim(V):    VarLy = ApSy(V)
Case IsArray(V):   VarLy = AySy(V)
Case IsObject(V):  VarLy = ApSy("*Type: " & TypeName(V))
Case IsEmpty(V):   VarLy = ApSy("*Empty")
Case IsMissing(V): VarLy = ApSy("*Missing")
Case Else: Stop
End Select
End Function

Function AySampleLin$(A)
Dim S$, U&
U = UB(A)
If U >= 0 Then
    Select Case True
    Case IsPrim(A(0)): S = "[" & A(0) & "]"
    Case IsObject(A(0)), IsArray(A(0)): S = "[*Ty:" & TypeName(A(0)) & "]"
    Case Else: Stop
    End Select
End If
AySampleLin = "*Ay:[" & U & "]" & S
End Function
Function VarLin$(V)
Select Case True
Case IsPrim(V):   VarLin = V
Case IsArray(V):  VarLin = AySampleLin(V)
Case IsObject(V): VarLin = "*Type:" & TypeName(V)
Case Else: Stop
End Select
End Function

Function IsPrim(A) As Boolean
Select Case VarType(A)
Case _
   VbVarType.vbBoolean, _
   VbVarType.vbByte, _
   VbVarType.vbCurrency, _
   VbVarType.vbDate, _
   VbVarType.vbDecimal, _
   VbVarType.vbDouble, _
   VbVarType.vbInteger, _
   VbVarType.vbLong, _
   VbVarType.vbSingle, _
   VbVarType.vbString
   IsPrim = True
End Select
End Function

Sub Trc(Fun$, Msg$, ParamArray Ap())
Dim Av(): Av = Ap
FunMsgAv_Dmp Fun, Msg, Av
End Sub
Sub FunMsgAv_Dmp(Fun$, Msg$, Av())
AyDmp FunMsgAv_Ly(Fun, Msg, Av)
End Sub
Sub MsgAp_Dmp(A$, ParamArray Ap())
Dim Av(): Av = Ap
AyDmp MsgAv_Ly(A, Av)
End Sub
Sub MsgBrw(A$, ParamArray Ap())
Dim Av(): Av = Ap
MsgAv_Brw A, Av
End Sub
Sub MsgAv_Brw(A$, Av())
AyBrw MsgAv_Ly(A, Av)
End Sub
Sub FunMsgAv_Brw(A, Msg$, Av())
AyBrw FunMsgAv_Ly(A, Msg, Av)
End Sub
Sub MsgBrwStop(A$, ParamArray Ap())
Dim Av(): Av = Ap
MsgAv_Brw A, Av
Stop
End Sub
Function MsgLy(A$, ParamArray Ap()) As String()
Dim Av(): Av = Ap
MsgLy = MsgAv_Ly(A, Av)
End Function
Function MsgAp_Ly(A$, ParamArray Ap()) As String()
Dim Av(): Av = Ap
MsgAp_Ly = MsgAv_Ly(A, Av)
End Function
Function MsgAp_Lin$(A$, ParamArray Ap())
Dim Av(): Av = Ap
MsgAp_Lin = MsgAv_Lin(A, Av)
End Function
Function MsgAv_Ly(A$, Av()) As String()
Dim B$(), C$()
B = SplitVBar(A)
C = NyAv_Ly(MsgNy(A), Av)
MsgAv_Ly = AyAdd(B, C)
End Function
Function MsgAv_Lin$(A$, Av())
Dim B$(), C$
B = SplitVBar(A)
C = NyLin(MsgNy(A), Av)
MsgAv_Lin = EnsSfxDot(A) & C
End Function
Function CvAv(A) As Variant()
CvAv = A
End Function
Sub Commit(Optional Msg$ = "Commit")
AppCommit Msg
End Sub
Sub FcommitBrw()
FtBrw BldCommitFcmd
End Sub
Function BldCommitFcmd$()
Dim O$(), Cd$, GitAdd$, GitCommit$, GitPush
Cd = FmtQQ("Cd ""?""", SrcPth)
GitAdd = "git add -A"
GitCommit = "git commit --message=%1%"
Push O, Cd
Push O, GitAdd
Push O, GitCommit
Push O, "Pause"
BldCommitFcmd = AyWrt(O, Fcmd.Commit)
End Function

Function BldPushAppFcmd$()
Dim O$(), Cd$, GitPush
Cd = FmtQQ("Cd ""?""", SrcPth)
GitPush = "git push -u https://johnsoncheung@github.com/johnsoncheung/StockShipRate.git master"
Push O, Cd
Push O, GitPush
Push O, "Pause"
BldPushAppFcmd = AyWrt(O, Fcmd.PushApp)
End Function
Sub AppCommit(Optional Msg$ = "Commit")
AppExp
FcmdRunMax BldCommitFcmd, Msg
End Sub
Sub AppPush()
FcmdRunMax BldPushAppFcmd
End Sub
Sub ZZ_FcmdRunMax()
FcmdRunMax "Cmd"
MsgBox "AA"
End Sub
Function AyQuoteDbl(A) As String()
If Sz(A) = 0 Then Exit Function
Dim I, O$()
For Each I In A
    Push O, QuoteDbl(I)
Next
AyQuoteDbl = O
End Function
Function QuoteDbl$(A)
QuoteDbl = """" & A & """"
End Function

Function FcmdRunMax$(A$, ParamArray Ap())
' WinSty As VbAppWinStyle = vbMaximizedFocus)
Dim Av(): Av = Ap
Dim Cmd$
    Cmd = JnSpc(AyQuoteDbl(AyAdd(Array(A), Av)))
Shell Cmd, vbMaximizedFocus
FcmdRunMax = A
End Function
Function FunMsgLy(A, Msg$, Av()) As String()
FunMsgLy = FunMsgAv_Ly(A, Msg, Av)
End Function
Function FunMsgAv_Ly(A, Msg$, Av()) As String()
Dim B$(), C$()
B = SplitVBar(Msg)
C = NyAv_Ly(CvSy(AyAdd(ApSy("Fun"), MsgNy(Msg))), CvAv(AyAdd(Array(A), Av)))
FunMsgAv_Ly = AyAdd(B, C)
End Function
Function FldVal(A As DAO.Field)
Asg A.Value, FldVal
End Function
Function DbAtt_AttRs(A As Database, Att) As AttRs
Q = FmtQQ("Select Att,FilTim,FilSz,AttNm from Att where AttNm='?'", Att)
With DbAtt_AttRs
    Set .TblRs = A.OpenRecordset(Q)
    With .TblRs
        If .EOF Then
            .AddNew
            !AttNm = Att
            .Update
            .MoveFirst
        End If
    End With
    Set .AttRs = FldVal(.TblRs!Att) '.Fields(0).Value
End With
End Function
Function DbFstAttRs(A As Database) As AttRs
With DbFstAttRs
    Set .TblRs = A.TableDefs("Att").OpenRecordset
    Set .AttRs = .TblRs.Fields("Att").Value
End With
End Function
Sub ZZ_DbAttExpFfn()
Dim T$
T = TmpFx
DbAttExpFfn CurrentDb, "Tp", "TaxRateAlert(Template).xlsm", T
Debug.Assert FfnIsExist(T)
Kill T
End Sub
Function DbAttExpFfn$(A As Database, Att$, AttFn$, ToFfn$)
Dim F2 As Field2, O$(), AttRs As AttRs
If FfnExt(AttFn) <> FfnExt(ToFfn) Then
    Stop
End If
If FfnIsExist(ToFfn) Then Stop
AttRs = DbAtt_AttRs(A, Att)
With AttRs
    With .AttRs
        .MoveFirst
        While Not .EOF
            If !FileName = AttFn Then
                Set F2 = !FileData
                F2.SaveToFile ToFfn
                DbAttExpFfn = ToFfn
                Exit Function
            End If
            .MoveNext
        Wend
        Push O, "Database          : " & A.Name
        Push O, "AttKey            : " & Att
        Push O, "Missing-AttFn     : " & AttFn
        Push O, "AttKey-File-Count : " & AttRs.AttRs.RecordCount
        PushAy O, AyAddPfx(RsSy(AttRs.AttRs, "FileName"), "Fn in AttKey      : ")
        Push O, "Att-Table in Database has AttKey, but no Fn-of-Ffn"
        AyBrw O
        Stop
        Exit Function
    End With
End With
If IsNothing(F2) Then Stop
F2.SaveToFile ToFfn
DbAttExpFfn = ToFfn
End Function
Sub DbDrpAtt(A As Database, Att)
A.Execute FmtQQ("Delete * from Att where AttNm='?'", Att)
End Sub
Sub AttDrp(Att)
DbDrpAtt CurrentDb, Att
End Sub
Sub AttyDrp(Atty0)
DbDrpAtty CurrentDb, Atty0
End Sub
Sub DbDrpAtty(A As Database, Atty0)
AyDoPX CvNy(Atty0), "DbDrpAtt", A
End Sub
Sub AttClr(A)
DbAtt_Clr CurrentDb, A
End Sub
Sub DbAtt_Clr(A As Database, Att)
RsClr DbAtt_AttRs(A, Att).AttRs
End Sub
Sub RsClr(A As DAO.Recordset)
With A
    While Not .EOF
        .Delete
        .MoveNext
    Wend
End With
End Sub
Function AttExpFfn$(A$, AttFn$, ToFfn$)
AttExpFfn = DbAttExpFfn(CurrentDb, A, AttFn, ToFfn)
End Function

Function DbAttTblRs(A As Database, AttNm$) As DAO.Recordset
Set DbAttTblRs = A.OpenRecordset(FmtQQ("Select * from Att where AttNm='?'", AttNm))
End Function

Function DbAttFnAy(A As Database, Att$) As String()
Dim T As DAO.Recordset ' AttTblRs
Dim F As DAO.Recordset ' AttFldRs
Set T = DbAttTblRs(A, Att)
Set F = T.Fields("Att").Value
DbAttFnAy = RsSy(F, "FileName")
End Function

Function AttFnAy(A) As String()
AttFnAy = DbAttFnAy(CurrentDb, "AA")
End Function

Function ZZ_AttFnAy()
D AttFnAy("AA")
End Function
Sub ZZ_AttImp()
Dim T$
T = TmpFt
StrWrt "sdfdf", T
AttImp "AA", T
Kill T
'T = TmpFt
'AttExpFfn "AA", T
'FtBrw T
End Sub

Function RsMovFst(A As DAO.Recordset) As DAO.Recordset
A.MoveFirst
Set RsMovFst = A
End Function
Function AttFfn$(A)
'Return Fst-Ffn-of-Att-A
AttFfn = RsMovFst(AttRs(A).AttRs)!FileName
End Function
Function RsNRec&(A As DAO.Recordset)
Dim O&
With A
    .MoveFirst
    While Not .EOF
        O = O + 1
        .MoveNext
    Wend
    .MoveFirst
End With
RsNRec = O
End Function
Function AttRs_FilCnt%(A As AttRs)
AttRs_FilCnt = RsNRec(A.AttRs)
End Function
Function DbAtt_FilCnt%(A As Database, Att)
'DbAtt_FilCnt = DbAtt_AttRs(A, Att).AttRs.RecordCount
DbAtt_FilCnt = AttRs_FilCnt(DbAtt_AttRs(A, Att))
End Function
Function AttFilCnt%(A)
AttFilCnt = DbAtt_FilCnt(CurrentDb, A)
End Function
Function AttExp$(A, ToFfn)
'Exporting the only file in Att
AttExp = DbAtt_Exp(CurrentDb, A, ToFfn)
End Function

Sub AttImp(A$, FmFfn$)
DbAtt_Imp CurrentDb, A, FmFfn
End Sub

Sub DbAtt_Imp(A As Database, Att$, FmFfn$)
AttRs_Imp DbAtt_AttRs(A, Att), FmFfn
End Sub

Function AttFstFn$(A)
AttFstFn = DbAtt_FstFn(CurrentDb, A)
End Function
Function AttRs_FstFn$(A As AttRs)
With A.AttRs
    If .EOF Then
        If .BOF Then
            FunMsgDmp "AttRs_FstFn", "[AttNm] has no attachment files", AttRs_AttNm(A)
            Exit Function
        End If
    End If
    .MoveFirst
    AttRs_FstFn = !FileName
End With
End Function
Function DbAtt_FstFn(A As Database, Att)
DbAtt_FstFn = AttRs_FstFn(DbAtt_AttRs(A, Att))
End Function

Function RsHasFldV(A As DAO.Recordset, F$, V) As Boolean
With A
    If .BOF Then
        If .EOF Then Exit Function
    End If
    .MoveFirst
    While Not .EOF
        If .Fields(F) = V Then RsHasFldV = True: Exit Function
        .MoveNext
    Wend
End With
End Function
Function DbAttNy(A As Database) As String()
Q = "Select AttNm from Att order by AttNm": DbAttNy = RsSy(A.OpenRecordset(Q))
End Function

Property Get AttNy() As String()
AttNy = CDbAttNy
End Property
Function AttRs_AttNm$(A As AttRs)
AttRs_AttNm = A.TblRs!AttNm
End Function

Sub AttRs_Imp(A As AttRs, Ffn$)
Const CSub$ = "AttRs_Imp"
Dim F2 As Field2
Dim S&, T$
S = FfnSz(Ffn)
T = FfnDTim(Ffn)
FunMsgDmp CSub, "[Att] is going to import [Ffn] with [Sz] and [Tim]", FldVal(A.TblRs!AttNm), Ffn, S, T
With A
    .TblRs.Edit
    With .AttRs
        If RsHasFldV(A.AttRs, "FileName", FfnFn(Ffn)) Then
            MsgDmp "Ffn is found in Att and it is replaced"
            .Edit
        Else
            MsgDmp "Ffn is not found in Att and it is imported"
            .AddNew
        End If
        Set F2 = !FileData
        F2.LoadFromFile Ffn
        .Update
    End With
    .TblRs.Fields!FilTim = FfnTim(Ffn)
    .TblRs.Fields!FilSz = FfnSz(Ffn)
    .TblRs.Update
End With
End Sub
Function AttLines$(A)
AttLines = DbAtt_Lines(CurrentDb, A)
End Function
Function DbAtt_Lines$(A As Database, Att)
DbAtt_Lines = AttRs_Lines(DbAtt_AttRs(A, Att))
End Function

Function AttRs_Lines$(A As AttRs)
Dim F As DAO.Field2, N%, Fn$
N = AttRs_FilCnt(A)
If N <> 1 Then
    FunMsgDmp "AttRs_Lines", "The [AttNm] should have one 1 attachment, but now [n-attachments]", AttRs_AttNm(A), N
    Exit Function
End If
Fn = FfnExt(AttRs_FstFn(A))
If Fn <> ".txt" Then
    FunMsgDmp "AttRs_Lines", "The [AttNm] has [Att-Fn] not being [.txt].  Cannot return Lines", AttRs_AttNm(A), Fn
    Exit Function
End If
AttRs_Lines = Fld2Lines(A.AttRs!FileData)
End Function

Function Fld2Lines$(A As DAO.Field2)
Dim O$, M$, Off&
X:
M = A.GetChunk(Off, 1024)
O = O & M
If Len(M) = 1024 Then
    Off = Off + 1024
    GoTo X
End If
Fld2Lines = O
End Function

Function TfkV(T, F, ParamArray K())
Dim Av(): Av = K
TfkV = DbtfkV(CurrentDb, T, F, Av)
End Function

Property Let RsF(A As DAO.Recordset, Optional F = 0, V)
With A
    .Edit
    .Fields(F).Value = V
    .Update
End With
End Property
Property Get RsF(A As DAO.Recordset, Optional F = 0)
RsF = A.Fields(F).Value
End Property

Function FxHasWs(A, Optional WsNy0 = "Sheet1") As Boolean
FxHasWs = AyHasAy(FxWsNy(A), CvNy(WsNy0))
End Function

Function FxDaoCnStr$(A)
'Excel 8.0;HDR=YES;IMEX=2;DATABASE=D:\Data\MyDoc\Development\ISS\Imports\PO\PUR904 (On-Line).xls;TABLE='PUR904 (On-Line)'
'INTO [Excel 8.0;HDR=YES;IMEX=2;DATABASE={0}].{1} FROM {2}"
'Excel 12.0 Xml;HDR=YES;IMEX=2;ACCDB=YES;DATABASE=C:\Users\sium\Desktop\TaxRate\sales text.xlsx;TABLE=Sheet1$
Dim O$
Select Case FfnExt(A)
Case ".xlsx":: O = "Excel 12.0 Xml;HDR=YES;IMEX=2;ACCDB=YES;DATABASE=" & A & ";"
Case ".xls": O = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=" & A & ";"
Case Else: Stop
End Select
FxDaoCnStr = O
End Function

Sub ZZ_FbWb_zExpOupTbl()
Dim W As Workbook
Set W = FbWb_zExpOupTbl(WFb)
WbVis W
Stop
W.Close False
Set W = Nothing
End Sub
Function AyPredSplit(A, Pred$) As Variant()
Dim O1, O2
O1 = AyCln(A)
O2 = O1
Dim X
For Each X In A
    If Run(Pred, X) Then
        Push O1, X
    Else
        Push O2, X
    End If
Next
AyPredSplit = Array(O1, O2)
End Function
Function AyWhHasPfx(A, Pfx$) As String()
AyWhHasPfx = AyWhPredXP(A, "HasPfx", Pfx)
End Function
Sub ZZ_FbOupTny()
D FbOupTny(WFb)
End Sub
Function FbOupTny(A) As String()
FbOupTny = AyWhHasPfx(FbTny(A), "@")
End Function

Sub AyRunABX(Ay, ABX$, A, B)
If Sz(Ay) = 0 Then Exit Sub
Dim X
For Each X In Ay
    Run ABX, A, B, X
Next
End Sub

Sub FbWrtFx_zForExpOupTb(A$, Fx$)
FbWb_zExpOupTbl(A).SaveAs Fx
End Sub
Sub WcAt(A As WorkbookConnection, At As Range)
Dim Lo As ListObject
Set Lo = RgWs(At).ListObjects.Add(SourceType:=0, Source:=A.OLEDBConnection.Connection, Destination:=At)
With Lo.QueryTable
    .CommandType = xlCmdTable
    .CommandText = A.Name
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .BackgroundQuery = True
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .PreserveColumnInfo = True
    .ListObject.DisplayName = TblNm_LoNm(A.Name)
    .Refresh BackgroundQuery:=False
End With
End Sub
Function WcAddWs(A As WorkbookConnection) As Worksheet
Dim Wb As Workbook, Ws As Worksheet, Lo As ListObject, Qt As QueryTable
Set Wb = A.Parent
Set Ws = WbAddWs(Wb, A.Name)
Ws.Name = A.Name
WcAt A, WsA1(Ws)
Set WcAddWs = Ws
End Function
Function WFbWb_zExpOupTbl() As Workbook
Set WFbWb_zExpOupTbl = FbWb_zExpOupTbl(WFb)
End Function

Function FbWb_zExpOupTbl(A) As Workbook
Dim O As Workbook
Set O = NewWb
AyRunABX FbOupTny(A), "WbAddWc", O, A
ItrDo O.Connections, "WcAddWs"
WbRfh O
Set FbWb_zExpOupTbl = O
End Function
Sub PushObj_zNonNothing(OY, Obj)
If IsNothing(Obj) Then Exit Sub
PushObj OY, Obj
End Sub
Function ItrpAyInto(A, P, OInto)
Dim X, O
O = OInto
Erase O
For Each X In A
    Push O, ObjPrp(X, P)
Next
ItrpAyInto = 0
End Function
Function WbWcAy_zOle(A As Workbook) As OLEDBConnection()
Dim O() As OLEDBConnection
WbWcAy_zOle = AyRmvEmp(ItrpAyInto(A.Connections, "OLEDBConnection", O))
End Function
Function WbWcSy_zOle(A As Workbook) As String()
WbWcSy_zOle = OyPrpSy(WbWcAy_zOle(A), "Connection")
End Function
Sub ZZ_WbWcSy()
D WbWcSy_zOle(FxWb(TpFx))
End Sub
Function WbWcNy(A As Workbook) As String()
WbWcNy = ItrNy(A.Connections)
End Function
Function WbAddWc(A As Workbook, Fb$, Nm$) As WorkbookConnection
Set WbAddWc = A.Connections.Add2(Nm, Nm, FbWcStr(Fb), Nm, XlCmdType.xlCmdTable)
End Function
Function SplitSC(A) As String()
SplitSC = Split(A, ";")
End Function
Function SplitCrLf(A) As String()
SplitCrLf = Split(A, vbCrLf)
End Function
Function AyKeepLasN(A, N)
Dim O, J&, I&, U&, Fm&, NewU&
U = UB(A)
If U < N Then AyKeepLasN = A: Exit Function
O = A
Fm = U - N + 1
NewU = N - 1
For J = Fm To U
    Asg O(J), O(I)
    I = I + 1
Next
ReDim Preserve O(NewU)
AyKeepLasN = O
End Function
Sub ZZ_LinesKeepLasN()
Dim Ay$(), A$, J%
For J = 0 To 9
Push Ay, "Line " & J
Next
A = Join(Ay, vbCrLf)
'Debug.Print fLasN(A, 3)
End Sub
Function LinesKeepLasN$(A$, N%)
Dim Ay$()
Ay = SplitCrLf(A)
LinesKeepLasN = JnCrLf(AyKeepLasN(Ay, N))
End Function
Function FbDaoCn(A) As DAO.Connection
Set FbDaoCn = DBEngine.OpenConnection(A)
End Function
Function CvCtl(A) As Access.Control
Set CvCtl = A
End Function
Function CvBtn(A) As Access.CommandButton
Set CvBtn = A
End Function
Function IsBtn(A) As Boolean
IsBtn = TypeName(A) = "CommandButton"
End Function
Function IsTgl(A) As Boolean
IsTgl = TypeName(A) = "ToggleButton"
End Function
Function CvTgl(A) As Access.ToggleButton
Set CvTgl = A
End Function
Sub CmdTurnOffTabStop(AcsCtl)
Dim A As Access.Control
Set A = AcsCtl
If Not HasPfx(A.Name, "Cmd") Then Exit Sub
Select Case True
Case IsBtn(A): CvBtn(A).TabStop = False
Case IsTgl(A): CvTgl(A).TabStop = False
End Select
End Sub
Sub FrmSetCmdNotTabStop(A As Access.Form)
ItrDo A.Controls, "CmdTurnOffTabStop"
End Sub
Function FxAdoCnStr$(A)
FxAdoCnStr = FmtQQ("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=?;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""", A)
End Function
Function FxOleCnStr$(A)
FxOleCnStr = "OLEDb;" & FxAdoCnStr(A)
End Function
Function FbOleCnStr$(A)
FbOleCnStr = "OLEDb;" & FbAdoCnStr(A)
End Function

Function FbAdoCnStr$(A)
'Const C$ = "Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False"
Const C$ = "Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False"
FbAdoCnStr = FmtQQ(C, A)
End Function
Function AdoCnStr_Cn(A) As adodb.Connection
Dim O As New adodb.Connection
O.Open A
Set AdoCnStr_Cn = O
End Function
Function FxCn(A) As adodb.Connection
Set FxCn = AdoCnStr_Cn(FxAdoCnStr(A))
End Function
Function FbCn(A) As adodb.Connection
Set FbCn = AdoCnStr_Cn(FbAdoCnStr(A))
End Function

Function FxCat(A) As Catalog
Set FxCat = CnCat(FxCn(A))
End Function

Function CnCat(A As adodb.Connection) As Catalog
Dim O As New Catalog
Set O.ActiveConnection = A
Set CnCat = O
End Function

Function FbTny(A) As String()
Dim Db As Database
Set Db = FbDb(A)
FbTny = DbTny(Db)
Db.Close
'FbTny = CvSy(AyWhPredXPNot(CatTny(FbCat(A)), "HasPfx", "MSys"))
End Function

Function AyCln(A)
Dim O
O = A
Erase O
AyCln = O
End Function

Function LinIsInT1Ay(A, T1Ay$()) As Boolean
LinIsInT1Ay = AyHas(T1Ay, LinT1(A))
End Function

Function AyWhPredXPNot(A, PredXP$, P)
If Sz(A) = 0 Then AyWhPredXPNot = AyCln(A): Exit Function
Dim O, X
O = AyCln(A)
For Each X In A
    If Not Run(PredXP, X, P) Then
        Push O, X
    End If
Next
AyWhPredXPNot = O
End Function

Function AyWhPredXP(A, XP$, P)
If Sz(A) = 0 Then AyWhPredXP = A: Exit Function
Dim O, X
O = AyCln(A)
For Each X In A
    If Run(XP, X, P) Then
        Push O, X
    End If
Next
AyWhPredXP = O
End Function
Function FbCat(A) As Catalog
Set FbCat = CnCat(FbCn(A))
End Function
Function CatTny(A As Catalog) As String()
CatTny = ItrNy(A.Tables)
End Function
Function AyTakBefOrAll(A, Sep$) As String()
Dim O$(), I
If Sz(A) = 0 Then Exit Function
For Each I In A
    Push O, TakBefOrAll(CStr(I), Sep)
Next
AyTakBefOrAll = O
End Function
Function AyDist(A)
Dim O, I
If Sz(A) = 0 Then AyDist = A: Exit Function
O = AyCln(A)
For Each I In A
    PushNoDup O, I
Next
AyDist = O
End Function
Sub PushNoDup(A, M)
If AyHas(A, M) Then Exit Sub
Push A, M
End Sub
Sub ZZ_FxWsNy()
Const Fx$ = "Users\user\Desktop\Invoices 2018-02.xlsx"
D FxWsNy(Fx)
End Sub
Function IsSngQuoted(A) As Boolean
IsSngQuoted = IsQuoted(A, "'")
End Function
Function IsSqBktQuoted(A) As Boolean
IsSqBktQuoted = IsQuoted(A, "[", "]")
End Function
Function IsQuoted(A, Q1$, Optional ByVal Q2$) As Boolean
If Q2 = "" Then Q2 = Q1
If FstChr(A) <> Q1 Then Exit Function
If LasChr(A) <> Q2 Then Exit Function
IsQuoted = True
End Function
Function RmvSngQuote$(A)
If Not IsSngQuoted(A) Then RmvSngQuote = A: Exit Function
RmvSngQuote = RmvFstLasChr(A)
End Function
Function AyRmvSngQuote(A) As String()
AyRmvSngQuote = AyMapSy(A, "RmvSngQuote")
End Function
Function FxWsNy(A) As String()
Dim T$()
T = CatTny(FxCat(A))
FxWsNy = AyDist(AyTakBefOrAll(AyRmvSngQuote(T), "$"))
End Function

Sub DbtImpTbl(A As Database, TT)
Dim Tny$(), J%, S$
Tny = CvNy(TT)
For J = 0 To UB(Tny)
    DbDrpTbl A, "#I" & Tny(J)
    S = FmtQQ("Select * into [#I?] from [?]", Tny(J), Tny(J))
    A.Execute S
Next
End Sub
Function LnkColStr_Ly(A) As String()
Dim A1$(), A2$(), Ay() As LnkCol
Ay = LnkColStr_LnkColAy(A)
A1 = LnkColAy_Ny(Ay)
A2 = AyAlignL(AyQuoteSqBkt(LnkColAy_ExtNy(Ay)))
Dim J%, O$()
For J = 0 To UB(A1)
    Push O, A2(J) & "  " & A1(J)
Next
LnkColStr_Ly = O
End Function
Function AyLasEle(A)
Asg A(UB(A)), AyLasEle
End Function

Function AscIsDig(A%) As Boolean
AscIsDig = &H30 <= A And A <= &H39
End Function

Property Get LnkCol(Nm$, Ty As DAO.DataTypeEnum, Extnm$) As LnkCol
Dim O As New LnkCol
Set LnkCol = O.Init(Nm, Ty, Extnm)
End Property

Function LnkColStr_LnkColAy(A) As LnkCol()
Dim Emp() As LnkCol, Ay$()
Ay = SplitVBar(A): If Sz(Ay) = 0 Then Stop: Exit Function
LnkColStr_LnkColAy = AyMapInto(Ay, "LinLnkCol", Emp)
End Function

Function SplitVBar(A) As String()
SplitVBar = Split(A, "|")
End Function

Sub ZZ_LinLnkCol()
Dim A$, Act As LnkCol, Exp As LnkCol
A = "AA Txt XX"
Exp = LnkCol("AA", dbText, "AA")
GoSub Tst
Exit Sub
Tst:
Act = LinLnkCol(A)
Debug.Assert LnkColIsEq(Act, Exp)
Return
End Sub
Function LnkColIsEq(A As LnkCol, B As LnkCol) As Boolean
With A
    If .Extnm <> B.Extnm Then Exit Function
    If .Ty <> B.Ty Then Exit Function
    If .Nm <> B.Nm Then Exit Function
End With
LnkColIsEq = True
End Function
Function LinLnkCol(A) As LnkCol
Dim Nm$, ShtTy$, Extnm$, Ty As DAO.DataTypeEnum
LinTTRstAsg A, Nm, ShtTy, Extnm
Extnm = RmvOptSqBkt(Extnm)
Ty = DaoShtTy_Ty(ShtTy)
Set LinLnkCol = LnkCol(Nm, Ty, IIf(Extnm = "", Nm, Extnm))
End Function
Function RmvFstLasChr$(A)
RmvFstLasChr = RmvFstChr(RmvLasChr(A))
End Function
Function DbtCnStr$(A As Database, T)
DbtCnStr = A.TableDefs(T).Connect
End Function
Sub DbtImpMap(A As Database, T$, LnkColStr$, Optional WhBExpr$)
If FstChr(T) <> ">" Then
    Debug.Print "FstChr of T must be >"
    Stop
End If
'Assume [>?] T exist
'Create [#I?] T
Dim S$
S = LnkColStr_ImpSql(LnkColStr, T, WhBExpr)
DbDrpTbl A, "#I" & Mid(T, 2)
A.Execute S
End Sub

Function LnkColStr_ImpSql$(A$, T$, Optional WhBExpr$)
Dim Ay() As LnkCol
Ay = LnkColStr_LnkColAy(A)
LnkColStr_ImpSql = LnkColAy_ImpSql(Ay, T, WhBExpr)
End Function


Function FstChr$(A)
FstChr = Left(A, 1)
End Function

Function LasChr$(A)
LasChr = Right(A, 1)
End Function

Property Get Drs(Fny0, Dry()) As Drs
Dim O As New Drs
Set Drs = O.Init(CvNy(Fny0), Dry)
End Property
Function ApSy(ParamArray Ap()) As String()
Dim Av(): Av = Ap
Dim O$(), J%, U&
U = UB(Av)
For J = 0 To UB(Av)
    PushNonEmpty O, Av(J)
Next
ApSy = O
End Function
Function DbtHasFld(A As Database, T$, F$) As Boolean
DbtHasFld = ItrHasNm(A.TableDefs(T).Fields, F)
End Function
Sub ZZ_SampleLo()
LoVis SampleLo
End Sub
Function SampleLo() As ListObject
Set SampleLo = DrsLo(SampleDrs, NewA1, "T_Sample")
End Function
Function DrsLo(A As Drs, At As Range, Optional LoNm$) As ListObject
Set DrsLo = RgLo(SqRg(DrsSq(A), At), LoNm)
End Function
Function SqRg(A, At As Range) As Range
Dim O As Range
Set O = RgReSz(At, A)
O.Value = A
Set SqRg = O
End Function

Function SampleDrs() As Drs
Set SampleDrs = Drs("A B C D E F", SampleDry)
End Function
Function SampleDry() As Variant()
Dim O(), Dr(), I%, J%
For J = 0 To 9
    ReDim Dr(5)
    For I = 0 To 5
        Dr(I) = J * 100 + I
    Next
    Push O, Dr
Next
SampleDry = O
End Function
Function AyIdx&(A, Itm)
AyIdx = AyIdxFm(A, Itm, 0)
End Function
Function AyIdxFm&(A, Itm, Fm&)
Dim O&
For O = Fm To UB(A)
    If A(O) = Itm Then AyIdxFm = O: Exit Function
Next
AyIdxFm = -1
End Function
Sub ZZ_AyHasAyInSeq()
Dim A, B
A = Array(1, 2, 3, 4, 5, 6, 7, 8)
B = Array(2, 4, 6)
Debug.Assert AyHasAyInSeq(A, B) = True

End Sub
Function AyHasAyInSeq(A, B) As Boolean
Dim BItm, Ix&
If Sz(B) = 0 Then Stop
For Each BItm In B
    Ix = AyIdxFm(A, BItm, Ix)
    If Ix = -1 Then Exit Function
    Ix = Ix + 1
Next
AyHasAyInSeq = True
End Function
Sub LoSetOutLin_zX(A As ListObject, ReSeqSpec)
'XReSeq
Dim L1Ny$(), LFny$()
LFny = LoFny(A)
L1Ny = CvNy(ReSeqSpec)
If AyHasAyInSeq(LFny, L1Ny) Then Stop
Dim C As ListColumn
For Each C In A.ListColumns
    If Not AyHas(L1Ny, C.Name) Then
        C.Range.EntireColumn.OutlineLevel = 2
    End If
Next
End Sub

Function LikAy_HasNm(A$(), Nm) As Boolean
If Sz(A) = 0 Then Exit Function
'Debug.Print "LikAy_HasNm: " & Nm
Dim Lik
For Each Lik In A
    If Nm Like Lik Then LikAy_HasNm = True: Exit Function
Next
End Function
Sub LcSetAlignC(A As ListColumn)
A.DataBodyRange.HorizontalAlignment = XlHAlign.xlHAlignCenter
End Sub
Sub LoSetAlignC_zX(A As ListObject, XAlignC$)
Dim LikAy$(), C As ListColumn
LikAy = SslSy(XAlignC)
For Each C In A.ListColumns
    If StrInLikAy(C.Name, LikAy) Then
        LcSetAlignC C
    End If
Next
End Sub

Sub LoSetWdt_zX(A As ListObject, XWdt$())
'XWdt is {Wdt} {FldLikSsl..}
'        ..
If Sz(XWdt) = 0 Then Exit Sub
Dim C As ListColumn
For Each C In A.ListColumns
    LcSetWdt C, LcnmWdt(C.Name, XWdt)
Next
End Sub

Sub LoSetTot_zX(A As ListObject, XTot$, TotCalc As XlTotalsCalculation)
'XTot is {FldLikSsl..}
'        ..
Dim IsSet As Boolean
If Trim(XTot) = "" Then Exit Sub
Dim LikAy$(), C As ListColumn
LikAy = SslSy(XTot)
For Each C In A.ListColumns
    If StrInLikAy(C.Name, LikAy) Then
        C.TotalsCalculation = TotCalc
        IsSet = True
    End If
Next
If IsSet Then
    A.ShowTotals = True
End If
End Sub

Sub AyDoABX(Ay, ABX$, A, B)
If Sz(Ay) = 0 Then Exit Sub
Dim X
For Each X In Ay
    Run ABX, A, B, X
Next
End Sub
Sub AyDoPX(Ay, PX$, P)
If Sz(Ay) = 0 Then Exit Sub
Dim X
For Each X In Ay
    Run PX, P, X
Next
End Sub
Function SqRplLo(A, Lo As ListObject) As ListObject
Dim LoNm$, At As Range
LoNm = Lo.Name
Set At = Lo.Range
Lo.Delete
Set SqRplLo = RgLo(SqRg(A, At), LoNm)
End Function
Function SqLo(A, At As Range, Optional LoNm$) As ListObject
Set SqLo = RgLo(SqRg(A, At), LoNm)
End Function
Function WbLoAy(A As Workbook) As ListObject()
Dim Ws As Worksheet, O() As ListObject, I
For Each Ws In A.Sheets
    OyPushItr O, Ws.ListObjects
Next
WbLoAy = O
End Function
Sub ZZ_WbTLoAy()
D OyNy(WbTLoAy(TpWb))
End Sub
Sub ZZ_WbLoAy()
D OyNy(WbLoAy(TpWb))
End Sub
Function OyNy(A) As String()
If Sz(A) = 0 Then Exit Function
OyNy = ItrNy(A)
End Function
Function WbTLoAy(A As Workbook) As ListObject()
WbTLoAy = OyWhNmHasPfx(WbLoAy(A), "T_")
End Function
Sub OyPushItr(OY, Itr)
Dim I
For Each I In Itr
    PushObj OY, I
Next
End Sub
Function AscIsUCase(A%) As Boolean
AscIsUCase = 65 <= A And A <= 90
End Function
Function AscIsLCase(A%) As Boolean
AscIsLCase = 97 <= A And A <= 122
End Function
Function AscIsLetter(A%) As Boolean
AscIsLetter = True
If AscIsUCase(A) Then Exit Function
If AscIsLCase(A) Then Exit Function
AscIsLetter = False
End Function
Function RmvFstNonLetter$(A)
If AscIsLetter(Asc(A)) Then
    RmvFstNonLetter = A
Else
    RmvFstNonLetter = RmvFstChr(A)
End If
End Function
Function AyRmvFstNonLetter(A) As String()
AyRmvFstNonLetter = AyMapSy(A, "RmvFstNonLetter")
End Function
Function DbtNewWb(A As Database, TT) As Workbook

End Function

Function DbtRplLo(A As Database, T$, Lo As ListObject, Optional ReSeqSpec$) As ListObject
Set DbtRplLo = SqRplLo(DbtSq(A, T, ReSeqSpec), Lo)
End Function
Sub ZZ_LoKeepFstCol()
LoKeepFstCol LoVis(SampleLo)
End Sub
Sub LoKeepFstCol(A As ListObject)
Dim J%
For J = A.ListColumns.Count To 2 Step -1
    A.ListColumns(J).Delete
Next
End Sub
Function WbLo(A As Workbook, LoNm$) As ListObject
Dim Ws As Worksheet
For Each Ws In A.Sheets
    If WsHasLo(Ws, LoNm) Then Set WbLo = Ws.ListObjects(LoNm): Exit Function
Next
End Function

Function WsHasLo(A As Worksheet, LoNm$) As Boolean
WsHasLo = ItrHasNm(A.ListObjects, LoNm)
End Function
Sub LoKeepFstRow(A As ListObject)
Dim J%
For J = A.ListRows.Count To 2 Step -1
    A.ListRows(J).Delete
Next
End Sub
Sub DbDrpTT(A As Database, TT)
AyDoPX CvNy(TT), "DbDrpTbl", A
End Sub
Sub DbDrpTbl(A As Database, T)
If DbHasTbl(A, T) Then A.Execute FmtQQ("Drop Table [?]", T)
End Sub
Sub SavRec()
DoCmd.RunCommand acCmdSaveRecord
End Sub

Function DbqSy(A As Database, Sql) As String()
DbqSy = RsSy(A.OpenRecordset(Sql))
End Function
Function DbqRs(A As Database, Sql) As DAO.Recordset
Set DbqRs = A.OpenRecordset(Sql)
End Function

Function Acs() As Access.Application
Static X As Boolean, Y As Access.Application
On Error GoTo X
If X Then
    Set Y = New Access.Application
    Y.Visible = True
    X = True
End If
If Y.Application.Name = "Microsoft Access" Then
    Set Acs = Y
    Exit Function
End If
X:
    Set Y = New Access.Application
    Y.Visible = True
    Debug.Print "Acs: New Acs instance is crreated."
Set Acs = Y
End Function

Sub AcsVis(A As Access.Application)
If Not A.Visible Then A.Visible = True
End Sub

Function IsNothing(A) As Boolean
IsNothing = TypeName(A) = "Nothing"
End Function

Sub ZZ_AyAddPfx()
Dim A, Act$(), Pfx$, Exp$()
A = Array(1, 2, 3, 4)
Pfx = "* "
Exp = ApSy("* 1", "* 2", "* 3", "* 4")
GoSub Tst
Exit Sub
Tst:
Act = AyAddPfx(A, Pfx)
Debug.Assert AyIsEq(Act, Exp)
Return
End Sub

Sub ZZ_AyAddSfx()
Dim A, Act$(), Sfx$, Exp$()
A = Array(1, 2, 3, 4)
Sfx = "#"
Exp = ApSy("1#", "2#", "3#", "4#")
GoSub Tst
Exit Sub
Tst:
Act = AyAddSfx(A, Sfx)
Debug.Assert AyIsEq(Act, Exp)
Return
End Sub

Sub ZZ_AyAddPfxSfx()
Dim A, Act$(), Sfx$, Pfx$, Exp$()
A = Array(1, 2, 3, 4)
Pfx = "* "
Sfx = "#"
Exp = ApSy("* 1#", "* 2#", "* 3#", "* 4#")
GoSub Tst
Exit Sub
Tst:
Act = AyAddPfxSfx(A, Pfx, Sfx)
Debug.Assert AyIsEq(Act, Exp)
Return
End Sub

Function AyAddPfx(A, Pfx$) As String()
AyAddPfx = AyMapXPSy(A, "AddPfx", Pfx)
End Function

Function AyAddSfx(A, Sfx$) As String()
AyAddSfx = AyMapXPSy(A, "AddSfx", Sfx)
End Function

Function AddPfx$(A$, Pfx$)
AddPfx = Pfx & A
End Function

Function AddSfx$(A$, Sfx$)
AddSfx = A & Sfx
End Function

Function AyAddPfxSfx(A, Pfx$, Sfx$) As String()
AyAddPfxSfx = AyMapXABSy(A, "AddPfxSfx", Pfx, Sfx)
End Function

Function AyMapXABSy(Ay, XAB$, A, B) As String()
AyMapXABSy = AyMapXABInto(Ay, XAB, A, B, EmpSy)
End Function
Function AddPfxSfx$(A$, Pfx$, Sfx$)
AddPfxSfx = Pfx & A & Sfx
End Function
Function AyMapXABInto(Ay, XAB$, A, B, OInto)
Dim O, X, J&, U&
O = OInto
Erase O
If U = -1 Then AyMapXABInto = O: Exit Function
For Each X In Ay
    Asg Run(XAB, X, A, B), O(J)
    J = J + 1
Next
AyMapXABInto = O
End Function
Function IsObjAy(A) As Boolean
IsObjAy = VarType(A) = vbArray + vbObject
End Function
Function AyRmvEle(A, Ele)
Dim O: O = AyCln(A)
Dim X
If Sz(A) = 0 Then AyRmvEle = O: Exit Function
For Each X In A
    If X <> Ele Then Push O, X
Next
AyRmvEle = O
End Function
Function AyRmvEleAt(A, Optional At&)
Dim O, J&, U&
U = UB(A)
O = A
Select Case True
Case U = 0
    Erase O
    AyRmvEleAt = O
    Exit Function
Case IsObjAy(A)
    For J = At To U - 1
        Set O(J) = O(J + 1)
    Next
Case Else
    For J = At To U - 1
        O(J) = O(J + 1)
    Next
End Select
ReDim Preserve O(U - 1)
AyRmvEleAt = O
End Function
Function AbIsEq(A, B) As Boolean
If VarType(A) <> VarType(B) Then Exit Function
Select Case True
Case IsObject(A): AbIsEq = ObjPtr(A) = ObjPtr(B)
Case IsArray(A): AbIsEq = AyIsEq(A, B)
Case Else: AbIsEq = A = B
End Select
End Function
Private Sub ZZZ_AyShift()
Dim Ay(), Exp, Act, ExpAyAft()
Ay = Array(1, 2, 3, 4)
Exp = 1
ExpAyAft = Array(2, 3, 4)
GoSub Tst
Exit Sub
Tst:
Act = AyShift(Ay)
Debug.Assert AbIsEq(Exp, Act)
Debug.Assert AyIsEq(Ay, ExpAyAft)
Return
End Sub
Function AyShift(Ay)
AyShift = Ay(0)
Ay = AyRmvEleAt(Ay)
End Function
Private Sub ZZZ_PfxSsl_Sy()
Dim A$, Exp$()
A = "A B C D"
Exp = SslSy("AB AC AD")
GoSub Tst
Exit Sub
Tst:
Dim Act$()
Act = PfxSsl_Sy(A)
Debug.Assert AyIsEq(Act, Exp)
Return
End Sub
Function ItrFstPrpEq(A, PrpNm$, V)
Dim I, OP
For Each I In A
    OP = ObjPrp(I, PrpNm)
    If OP = V Then
        Asg I, ItrFstPrpEq
        Exit Function
    End If
Next
Exit Function ' If not found return Empty
'Impossible
Debug.Print PrpNm, V
For Each I In A
    Debug.Print ObjPrp(I, PrpNm)
Next
Stop
End Function
Function ObjPrp(A, PrpNm)
On Error GoTo X
Dim V
V = CallByName(A, PrpNm, VbGet)
Asg V, ObjPrp
Exit Function
X:
Debug.Print "ObjPrp: " & Err.Description
End Function
Function ItrPrpSy(A, PrpNm$) As String()
ItrPrpSy = ItrPrpInto(A, PrpNm, EmpSy)
End Function
Function ItrPrpInto(A, PrpNm$, OInto)
Dim O, I
O = OInto
Erase O
For Each I In A
    Push O, ObjPrp(I, PrpNm)
Next
ItrPrpInto = O
End Function
Function WbWsCdNy(A As Workbook) As String()
WbWsCdNy = ItrPrpSy(A.Sheets, "CodeName")
End Function
Function FxWsCdNy(A) As String()
Dim Wb As Workbook
Set Wb = FxWb(A)
FxWsCdNy = WbWsCdNy(Wb)
Wb.Close False
End Function
Function PfxSsl_Sy(A) As String()
Dim Ay$(), Pfx$
Ay = SslSy(A)
Pfx = AyShift(Ay)
PfxSsl_Sy = AyAddPfx(Ay, Pfx)
End Function
Function ApnWAcs(A)
Dim O As Access.Application
AcsOpn O, ApnWFb(A)
Set ApnWAcs = O
End Function
Function ApnAcs(A) As Access.Application
AcsOpn Acs, ApnWFb(A)
Set ApnAcs = Acs
End Function
Sub AcsOpn(A As Access.Application, Fb$)
Select Case True
Case IsNothing(A.CurrentDb)
    A.OpenCurrentDatabase Fb
Case A.CurrentDb.Name = Fb
Case Else
    A.CurrentDb.Close
    A.OpenCurrentDatabase Fb
End Select
End Sub
Sub ApnBrwWDb(A)
Dim Fb$
Fb = ApnWFb(A)
AcsOpn Acs, Fb
AcsVis Acs
End Sub
Sub FbEns(A)
If FfnIsExist(A) Then Exit Sub
FbCrt A
End Sub
Function FbCrt(A) As Database
Set FbCrt = DBEngine.CreateDatabase(A, dbLangGeneral)
End Function
Sub FxRfhFbCnStr(A, Fb$)
WbRfhFbCnStr(FxWb(A), Fb).Close True
End Sub
Function WbRfhFbCnStr(A As Workbook, Fb$) As Workbook
ItrDoXP A.Connections, "WcRfhCnStr", FbWcStr(Fb)
Set WbRfhFbCnStr = A
End Function
Sub FbOpn(A)
Acs.OpenCurrentDatabase A
AcsVis Acs
End Sub
Function FbDb(A) As Database
Set FbDb = DBEngine.OpenDatabase(A)
End Function
Function PthFnIr(A, Optional Spec$ = "*") As VBA.Collection
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
Function PthUpOne$(A)
PthUpOne = TakBefOrAllRev(RmvSfx(A, "\"), "\") & "\"
End Function

Sub PthMovFilUp(A)
Dim I, Tar$
Tar$ = PthUp(A)
For Each I In PthFnIr(A)
    FfnMov I, Tar
Next
End Sub

Function ApnWFb$(A)
ApnWFb = ApnWPth(A) & "Wrk.accdb"
End Function
Sub WPthBrw()
PthBrw WPth
End Sub
Function WPth$()
WPth = ApnWPth(Apn)
End Function
Function ApnWPth$(A)
Dim P$
P = TmpHom & A & "\"
PthEns P
ApnWPth = P
End Function
Function DbIsOk(A As Database) As Boolean
On Error GoTo X
DbIsOk = IsStr(A.Name)
Exit Function
X:
End Function
Function WsC(A As Worksheet, C) As Range
Dim R As Range
Set R = A.Columns(C)
Set WsC = R.EntireColumn
End Function

Function ApnWDb(A) As Database
Static X As Boolean, Y As Database
If Not X Then
    X = True
    FbEns ApnWFb(A)
    Set Y = FbDb(ApnWFb(A))
End If
If Not DbIsOk(Y) Then Set Y = FbDb(ApnWFb(A))
Set ApnWDb = Y
End Function
Function DbqAny(A As Database, Sql) As Boolean
DbqAny = RsAny(DbqRs(A, Sql))
End Function
Function DbHasTbl(A As Database, T) As Boolean
DbHasTbl = DbqAny(A, FmtQQ("Select * from MSysObjects where Name='?' and Type in (1,6)", T))
End Function
Function AyWdt%(A)
Dim O%, J&
For J = 0 To UB(A)
    O = Max(O, Len(A(J)))
Next
AyWdt = O
End Function
Function TTStru$(TT)
TTStru = DbttStru(CurrentDb, TT)
End Function
Function TblStru$(T$)
TblStru = DbtStru(CurrentDb, T)
End Function

Function QTbl$(T$, Optional WhBExpr$)
QTbl = "Select *" & PFm(T) & PWh(WhBExpr)
End Function
Function WtPrpLoFmlVbl$(T)
WtPrpLoFmlVbl = FbtPrpLoFmlVbl(WFb, T)
End Function
Property Let FbtPrpLoFmlVbl(A, T, LoFmlVbl$)
DbtPrpLoFmlVbl(FbDb(A), T) = LoFmlVbl
End Property

Property Get FbtPrpLoFmlVbl$(A, T)
If A = "" And T = "" Then Exit Property
FbtPrpLoFmlVbl = DbtPrpLoFmlVbl(FbDb(A), T)
End Property

Function FbtFny(A$, T$) As String()
FbtFny = RsFny(DbqRs(FbDb(A), QTbl(T)))
End Function
Function Max(A, B)
If A > B Then
    Max = A
Else
    Max = B
End If
End Function
Function Min(A, B)
If A > B Then
    Min = B
Else
    Min = A
End If
End Function

Function CvNy(Ny0) As String()
Select Case True
Case IsMissing(Ny0)
Case IsStr(Ny0): CvNy = SslSy(Ny0)
Case IsSy(Ny0): CvNy = Ny0
Case IsArray(Ny0): CvNy = AySy(Ny0)
Case Else: Stop
End Select
End Function
Function AySy(A) As String()
If Sz(A) = 0 Then Exit Function
AySy = ItrAy(A, EmpSy)
End Function
Function EmpSy() As String()
End Function
Function EmpLngAy() As Long()
End Function
Function EmpAy() As Variant()
End Function
Sub TpMinLo()
Dim O As Workbook
Set O = TpWb
WbMinLo O
O.Save
WbVis O
End Sub

Function TpIdxWs() As Worksheet
Set TpIdxWs = WbWsCd(TpWb, "WsIdx")
End Function
Function TpWsCdNy() As String()
TpWsCdNy = FxWsCdNy(TpFx)
End Function
Function TpWb() As Workbook
Set TpWb = FxWb(TpFx)
End Function
Function TpWcSy() As String()
Dim W As Workbook, X As Excel.Application
Set X = New Excel.Application
Set W = X.Workbooks.Open(TpFx)
TpWcSy = WbWcSy_zOle(W)
W.Close False
Set W = Nothing
X.Quit
Set X = Nothing
End Function
Property Get TpFxm$()
TpFxm = PgmObjPth & TpFnn & ".xlsm"
End Property
Property Get TpFx$()
TpFx = PgmObjPth & TpFnn & ".xlsx"
End Property
Property Get TpFnn$()
TpFnn = Apn & "(Template)"
End Property

Sub TpOpn()
FxOpn TpFx
End Sub
Function WsPtAy(A As Worksheet) As PivotTable()
Dim O() As PivotTable, Pt As PivotTable
For Each Pt In A.PivotTables
    PushObj O, Pt
Next
WsPtAy = O
End Function

Function WbPtAy(A As Workbook) As PivotTable()
Dim O() As PivotTable, Ws As Worksheet
For Each Ws In A.Sheets
    PushObjAy O, WsPtAy(Ws)
Next
WbPtAy = O
End Function
Function ItrAy(A, OInto)
Dim O, I
O = OInto
Erase O
For Each I In A
    Push O, I
Next
ItrAy = O
End Function

Function OupPth$()
Dim A$
A = CDbPth & "Output\"
PthEns A
OupPth = A
End Function
Function OupPth_zPm$()
OupPth_zPm = PnmVal("OupPth")
End Function
Function YYYYMMDD_IsVdt(A) As Boolean
On Error Resume Next
YYYYMMDD_IsVdt = Format(CDate(A), "YYYY-MM-DD") = A
End Function
Function PgmObjPth$()
PgmObjPth = PthEns(CDbPth & "PgmObj\")
End Function
Function FfnPth$(A)
Dim P%: P = InStrRev(A, "\")
If P = 0 Then Exit Function
FfnPth = Left(A, P)
End Function
Private Function ErzFws__2(Fx$, WsNm$, ColNy$()) As String()

End Function
Private Function ErzFws__3(Fx$, WsNm$, ColNy$(), DtaTyAy() As DAO.DataTypeEnum) As String()

End Function
Sub ZZ_ErAyzFxWsMissingCol()
'" [Material]             As Sku," & _
'" [Plant]                As Whs," & _
'" [Storage Location]     As Loc," & _
'" [Batch]                As BchNo," & _
'" [Unrestricted]         As OH " & _

End Sub
Function TblF_Ty(T, F) As DAO.DataTypeEnum

End Function
Function TblErAyzCol(T$, ColNy$(), DtaTyAy() As DAO.DataTypeEnum, Optional AddTblLinMsg As Boolean) As String()
Dim Fny$(), F, Fny1$(), Fny2$()
Fny = TblFny(T)
For Each F In ColNy
    If AyHas(Fny, F) Then
        Push F, Fny1
    Else
        Push F, Fny2
    End If
Next
Dim O$()
If Sz(Fny2) > 0 Then
    Dim J%
    For J = 0 To UB(ColNy)
        If AyHas(Fny2, ColNy(J)) Then
            If TblF_Ty(T, ColNy(J)) <> DtaTyAy(J) Then
                Push O, "Column [?] has unexpected DataType[?].  It is expected to be [?]"
            End If
        End If
    Next
End If
If AddTblLinMsg Then
    Push O, ""
    
End If
End Function
Function FfnNotFndChk(A) As String()
If FfnIsExist(A) Then Exit Function
FfnNotFndChk = MsgAp_Ly("[File] not exist", A)
End Function
Function ChkFst(ChkSsl$) As String()
Dim O$(), I
For Each I In SslSy(ChkSsl)
    O = Run(I)
    If Sz(O) > 0 Then
        ChkFst = O
        Exit Function
    End If
Next
End Function
Function ChkAll(ChkSsl$) As String()
Dim O$(), I
For Each I In SslSy(ChkSsl)
    PushAy O, Run(I)
Next
ChkAll = O
End Function
Function UnderLin$(A)
UnderLin = String(Len(A), "-")
End Function
Function UnderLinDbl$(A)
UnderLinDbl = String(Len(A), "=")
End Function
Property Get PnmVal(Pnm$)
PnmVal = CurrentDb.TableDefs("Prm").OpenRecordset.Fields(Pnm).Value
End Property
Function DteLasDayOfMth(A As Date) As Date
DteLasDayOfMth = DtePrvDay(DteFstDteOfMth(DteNxtMth(A)))
End Function
Function DteFstDteOfMth(A As Date) As Date
DteFstDteOfMth = DateSerial(Year(A), Month(A), 1)
End Function
Function DtePrvDay(A As Date) As Date
DtePrvDay = DateAdd("D", -1, A)
End Function

Sub ZZ_AyMax()
Dim A()
Dim Act
Act = AyMax(A)
Stop
End Sub
Function AyWhLik(A, Lik) As String()
AyWhLik = AyWhPredXP(A, "Lik", Lik)
End Function
Function AyMax(A)
Dim O, I
If Sz(A) = 0 Then Exit Function
For Each I In A
    O = Max(O, I)
Next
AyMax = O
End Function
Function FldsFny(A As DAO.Fields) As String()
FldsFny = ItrNy(A)
End Function
Sub PthBrw(A)
Shell FmtQQ("Explorer ""?""", A), vbMaximizedFocus
End Sub
Function PthEnsSfx$(A)
PthEnsSfx = EnsSfx(A, "\")
End Function
Function ItrNy(A) As String()
Dim O$(), I
For Each I In A
    Push O, I.Name
Next
ItrNy = O
End Function
Sub Push(O, M)
Dim N&
N = Sz(O)
ReDim Preserve O(N)
If IsObject(M) Then
    Set O(N) = M
Else
    O(N) = M
End If
End Sub
Sub PushObj(O, M)
Dim N&
N = Sz(O)
ReDim Preserve O(N)
Set O(N) = M
End Sub
Sub PushObjAy(O, M)
If Sz(M) = 0 Then Exit Sub
Dim I
For Each I In M
    PushObj O, I
Next
End Sub
Private Sub ZZ_PthFxAy()
Dim A$()
A = PthFxAy(CurDir)
AyDmp A
End Sub

Function DteIsVdt(A) As Boolean
On Error Resume Next
DteIsVdt = Format(CDate(A), "YYYY-MM-DD") = A
End Function
Private Sub ZZ_TblFny()
AyDmp TblFny(">KE24")
End Sub
Function RsAy(A As DAO.Recordset, Optional FldNm$) As Variant()
RsAy = RsAyInto(A, FldNm, EmpAy)
End Function
Function RsAyInto(A As DAO.Recordset, FldNm$, OInto)
Dim O: O = OInto: Erase O
Dim Ix
Ix = IIf(FldNm = "", 0, FldNm)
With A
    If .EOF Then RsAyInto = O: Exit Function
    .MoveFirst
    While Not .EOF
        Push O, .Fields(Ix).Value
        .MoveNext
    Wend
    .Close
End With
RsAyInto = O
End Function
Function RsSy(A As DAO.Recordset, Optional FldNm$) As String()
RsSy = RsAyInto(A, FldNm, EmpSy)
End Function
Function RsLngAy(A As DAO.Recordset, Optional FldNm$) As Long()
RsLngAy = RsAyInto(A, FldNm, EmpLngAy)
End Function
Sub ZZ_SqlFny()
Const S$ = "SELECT qSku.*" & _
" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
AyDmp SqlFny(S)
End Sub
Function SqlFny(A) As String()
SqlFny = RsFny(SqlRs(A))
End Function
Sub ZZ_SqlRs()
Const S$ = "SELECT qSku.*" & _
" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
AyBrw RsCsvLy(SqlRs(S))
End Sub

Function SqlRs(A) As DAO.Recordset
Set SqlRs = CurrentDb.OpenRecordset(A)
End Function
Private Sub ZZ_SqlSy()
AyDmp SqlSy("Select Distinct UOR from [>Imp]")
End Sub
Function SqpzInBExpr$(Ay, FldNm$, Optional WithQuote As Boolean)
Const C$ = "[?] in (?)"
Dim B$
    If WithQuote Then
        B = JnComma(AyQuoteSng(Ay))
    Else
        B = JnComma(Ay)
    End If
SqpzInBExpr = FmtQQ(C, FldNm, B)
End Function
Function SqlSy(A) As String()
SqlSy = DbqSy(CurrentDb, A)
End Function
Function SqlLngAy(A) As Long()
SqlLngAy = DbqLngAy(CurrentDb, A)
End Function
Function FxChkWs(A, Optional FxKind$ = "Excel file", Optional WsNy0$ = "Sheet1") As String()
If Not FfnIsExist(A) Then
    Dim M$
    M = FmtQQ("[?] not found in [folder]", FxKind)
    FxChkWs = MsgLy(M, FfnFn(A), FfnPth(A))
    Exit Function
End If
If FxHasWs(A, WsNy0) Then Exit Function
M = FmtQQ("[?] in [folder] does not have [expected worksheets], but [these worksheets].", FxKind)
FxChkWs = MsgAp_Ly(M, FfnFn(A), FfnPth(A), CvNy(WsNy0), FxWsNy(A))
End Function
Sub ABAsg(AB$, OA$, OB$)
BrkAsg AB, " ", OA, OB
End Sub
Function Either(A As Boolean, ABFun$)
Dim Fst$, Snd$
ABAsg ABFun, Fst, Snd
If A Then
    Either = Run(Fst)
Else
    Either = Run(Snd)
End If
End Function
Function AyAdd(A, B)
Dim O
O = A
PushAy O, B
AyAdd = O
End Function
Sub ZZ_DbtWhDupKey()
TTDrp "#A #B"
DoCmd.RunSQL "Select Distinct Sku,BchNo,CLng(Rate) as RateRnd into [#A] from ZZ_DbtUpdSeq"
DbtWhDupKey CurrentDb, "#A", "Sku BchNo", "#B"
TTBrw "#B"
Stop
TTDrp "#B"
End Sub
Sub TTWbBrw(TT, Optional UseWc As Boolean)
WbVis TTWb(TT, UseWc)
End Sub
Sub TblBrw(T)
DoCmd.OpenTable T
End Sub
Function CvTT(A) As String()
CvTT = CvNy(A)
End Function

Sub TTBrw(TT)
'OFunAyDo DoCmd, "OpenTable", CvTT(TT)
End Sub

Sub DbtWhDupKey(A As Database, T$, KK, TarTbl$)
Dim Ky$(), K$, Jn$, Tmp$, J%
Ky = SslSy(KK)
Tmp = "##" & TmpNm
K = JnComma(Ky)
For J = 0 To UB(Ky)
    Ky(J) = FmtQQ("x.?=a.?", Ky(J), Ky(J))
Next
Jn = Join(Ky, " and ")
A.Execute FmtQQ("Select Distinct ?,Count(*) as Cnt into [?] from [?] group by ? having Count(*)>1", K, Tmp, T, K)
A.Execute FmtQQ("Select x.* into [?] from [?] x inner join [?] a on ?", TarTbl, T, Tmp, Jn)
DbDrpTbl A, Tmp
End Sub
Sub C()
Debug.Assert VarType(Actual) = VarType(Expect)
If IsArray(Actual) Then
    Debug.Assert AyIsEq(Actual, Expect)
Else
    Debug.Assert Actual = Expect
End If
End Sub
Sub D(A)
AyDmp VarLy(A)
End Sub
Sub AyDmp(A)
Dim I
If Sz(A) = 0 Then Exit Sub
For Each I In A
    Debug.Print I
Next
End Sub
Function TblFny(A) As String()
TblFny = DbtFny(CurrentDb, A)
End Function

Function DbtFny(A As Database, T) As String()
DbtFny = RsFny(DbtRs(A, T))
End Function
Function DbtIsXls(A As Database, T) As Boolean
On Error Resume Next
DbtIsXls = HasPfx(A.TableDefs(T).Connect, "Excel")
End Function
Function SplitSpc(A) As String()
SplitSpc = Split(A, " ")
End Function
Function SqlAny(A) As Boolean
SqlAny = DbqAny(CurrentDb, A)
End Function
Function RsAny(A As DAO.Recordset) As Boolean
RsAny = Not A.EOF
End Function
Function TblIsExist(T$) As Boolean
TblIsExist = DbHasTbl(CurrentDb, T)
End Function
Sub TblOpn(TblSsl$)
AyDo SslSy(TblSsl), "TblOpn_1"
End Sub
Sub AyDo(A, FunNm$)
If Sz(A) = 0 Then Exit Sub
Dim I
For Each I In A
    Run FunNm, I
Next
End Sub
Sub TblOpn_1(T)
DoCmd.OpenTable T
End Sub
Function RplDblSpc$(A)
Dim P%, O$, J%
O = A
While InStr(O, "  ") > 0
    J = J + 1
    If J > 50000 Then Stop
    O = Replace(O, "  ", " ")
Wend
RplDblSpc = O
End Function

Function SslSy(A) As String()
SslSy = SplitSpc(RplDblSpc(Trim(A)))
End Function
Sub ItrNmDo(A, DoFun$)
Dim I
For Each I In A
    Run DoFun, I.Name
Next
End Sub
Sub AcsClsTbl(A As Access.Application)
Dim T As AccessObject
For Each T In A.CodeData.AllTables
    A.DoCmd.Close acTable, T.Name
Next
End Sub
Sub AcstCls(A As Access.Application, T$)
A.DoCmd.Close acTable, T, acSaveYes
End Sub
Sub AcsttCls(A As Access.Application, TT)
AyDoPX CvNy(TT), "AcstCls", A
End Sub

Sub ClsTbl()
AcsClsTbl Application
End Sub

Sub TTCls(TT)
AyDo CvNy(TT), "TblCls"
End Sub


Sub TblCls(T)
DoCmd.Close acTable, T
End Sub
Sub TblDrp(T$)
DbDrpTbl CurrentDb, T
End Sub
Sub TTDrp(TT)
DbDrpTT CurrentDb, TT
End Sub

Function DbHasQry(A As Database, Q) As Boolean
DbHasQry = DbqAny(A, FmtQQ("Select * from MSysObjects where Name='?' and Type=5", Q))
End Function

Sub DbDrpQry(A As Database, Q)
If DbHasQry(A, Q) Then A.QueryDefs.Delete Q
End Sub

Sub DbCrtQry(A As Database, Q, Sql)
If Not DbHasQry(A, Q) Then
    Dim QQ As New QueryDef
    QQ.Sql = Sql
    QQ.Name = Q
    A.QueryDefs.Append QQ
Else
    A.QueryDefs(Q).Sql = Sql
End If
End Sub
Function LinShiftTerm$(O$)
Dim A$, P%
A = LTrim(O)
P = InStr(A, " ")
If P = 0 Then
    LinShiftTerm = A
    O = ""
    Exit Function
End If
LinShiftTerm = Left(A, P - 1)
O = LTrim(Mid(A, P + 1))
End Function

Sub LinTTRstAsg(A, OT1$, OT2$, ORst$)
Dim Ay$()
Ay = LinTTRst(A)
OT1 = Ay(0)
OT2 = Ay(1)
ORst = RTrim(Ay(2))
End Sub

Function LinTTRst(A) As String()
Dim O$(2), L$
L = A
O(0) = LinShiftTerm(L)
O(1) = LinShiftTerm(L)
O(2) = L
LinTTRst = O
End Function
Function AyMinus(A, B)
If Sz(B) = 0 Then AyMinus = A: Exit Function
If Sz(A) = 0 Then AyMinus = A: Exit Function
Dim O, I
O = A
Erase O
For Each I In A
    If Not AyHas(B, I) Then Push O, I
Next
AyMinus = O
End Function

Sub DbtRen(A As Database, Fm, ToTbl, Optional ReOpnFst As Boolean)
If ReOpnFst Then DbReOpn A
A.TableDefs(Fm).Name = ToTbl
End Sub

Function DbtChkCol(A As Database, T$, LnkColStr$) As String()
Dim Ay() As LnkCol, O$(), Fny$(), J%, Ty As DAO.DataTypeEnum, F$
Ay = LnkColStr_LnkColAy(LnkColStr)
Fny = LnkColAy_ExtNy(Ay)
O = DbtChkFny(A, T, Fny)
If Sz(O) > 0 Then DbtChkCol = O: Exit Function
For J = 0 To UB(Ay)
    F = Ay(J).Extnm
    Ty = Ay(J).Ty
    PushNonEmpty O, DbtChkFldType(A, T, F, Ty)
Next
If Sz(0) > 0 Then
    PushMsgUnderLin O, "Some field has unexpected type"
    DbtChkCol = O
End If
End Function
Function TakAft$(A, S)
TakAft = TakAftAt(A, InStr(A, S), S)
End Function
Function TakAftAt$(A, At&, S)
If At = 0 Then Exit Function
TakAftAt = Mid(A, At + Len(S))
End Function
Function TakAftRev$(A, S)
TakAftRev = TakAftAt(A, InStrRev(A, S), S)
End Function
Function TakBefOrAll$(A, S)
TakBefOrAll = StrDft(TakBef(A, S), A)
End Function
Function StrDft$(A, B)
StrDft = IIf(A = "", B, A)
End Function
Function TakAftOrAll$(A, S)
TakAftOrAll = StrDft(TakAft(A, S), A)
End Function
Function TakAftOrAllRev$(A, S)
TakAftOrAllRev = StrDft(TakAftRev(A, S), A)
End Function
Function TakBefOrAllRev$(A, S)
TakBefOrAllRev = StrDft(TakBefRev(A, S), A)
End Function
Function TakBefAt(A, At&)
If At = 0 Then Exit Function
TakBefAt = Left(A, At - 1)
End Function
Function TakBef$(A, S)
TakBef = TakBefAt(A, InStr(A, S))
End Function
Function TakBefRev$(A, S)
TakBefRev = TakBefAt(A, InStrRev(A, S))
End Function

Function DbtXlsLnkInf(A As Database, T) As XlsLnkInf
Dim Cn$
Cn = DbtCnStr(A, T)
If Not IsPfx(Cn, "Excel") Then Exit Function
With DbtXlsLnkInf
    .IsXlsLnk = True
    .Fx = TakBefOrAll(TakAft(Cn, "DATABASE="), ";")
    .WsNm = A.TableDefs(T).SourceTableName
    If LasChr(.WsNm) <> "$" Then Stop
    .WsNm = RmvLasChr(.WsNm)
End With
End Function

Function AyOfAy_Ay(A)
If Sz(A) = 0 Then AyOfAy_Ay = A: Exit Function
Dim O, J&
O = A(0)
For J = 1 To UB(A)
    PushAy O, A(J)
Next
AyOfAy_Ay = O
End Function
Function ISpecINm$(A)
ISpecINm = LinT1(A)
End Function
Sub LSpecDmp(A)
Debug.Print RplVBar(A)
End Sub
Function CurY() As Byte
CurY = CurYY - 2000
End Function
Function CurYY%()
CurYY = Year(Now)
End Function
Function CurM() As Byte
CurM = Month(Now)
End Function
Function LSpecLy(A) As String()
Const L2Spec$ = ">GLAnp |" & _
    "Whs    Txt Plant |" & _
    "Loc    Txt [Storage Location]|" & _
    "Sku    Txt Material |" & _
    "PstDte Txt [Posting Date] |" & _
    "MovTy  Txt [Movement Type]|" & _
    "Qty    Txt Quantity|" & _
    "BchNo  Txt Batch |" & _
    "Where Plant='8601' and [Storage Location]='0002' and [Movement Type] like '6*'"
End Function
Function HasPfx(A, Pfx) As Boolean
HasPfx = Left(A, Len(Pfx)) = Pfx
End Function
Function HasSfx(A, Sfx) As Boolean
HasSfx = Right(A, Len(Sfx)) = Sfx
End Function
Sub LSpecAsg(A, Optional OTblNm$, Optional OLnkColStr$, Optional OWhBExpr$)
Dim Ay$()
Ay = AyTrim(SplitVBar(A))
OTblNm = AyShift(Ay)
If LinT1(AyLasEle(Ay)) = "Where" Then
    OWhBExpr = LinRmvTerm(Pop(Ay))
Else
    OWhBExpr = ""
End If
OLnkColStr = JnVBar(Ay)
End Sub
Function Pop(A)
Pop = AyLasEle(A)
AyRmvLasEle A
End Function
Sub AyRmvLasEle(A)
If Sz(A) = 1 Then
    Erase A
    Exit Sub
End If
ReDim Preserve A(UB(A) - 1)
End Sub
Sub LSpecAy_Asg(A$(), OTny$(), OLnkColStrAy$(), OWhBExprAy$())
Dim U%, J%
U = UB(A)
ReDim OTny(U)
ReDim OLnkColStrAy(U)
ReDim OWhBExprAy(U)
For J = 0 To U
    LSpecAsg A(J), OTny(J), OLnkColStrAy(J), OWhBExprAy(J)
Next
End Sub

Function DbImp(A As Database, LSpec$()) As String()
Dim O$(), J%, T$(), L$(), W$(), U%
LSpecAy_Asg LSpec, T, L, W
U = UB(LSpec)
For J = 0 To U
    PushAy O, DbtChkCol(A, T(J), L(J))
Next
If Sz(O) > 0 Then DbImp = O: Exit Function
For J = 0 To U
    DbtImpMap A, T(J), L(J), W(J)
Next
DbImp = O
End Function

Function DbtMissFny_Er(A As Database, T$, MissFny$(), ExistingFny$()) As String()
Dim X As XlsLnkInf, O$(), I
If Sz(MissFny) = 0 Then Exit Function
X = DbtXlsLnkInf(A, T)
If X.IsXlsLnk Then
    Push O, "Excel File       : " & X.Fx
    Push O, "Worksheet        : " & X.WsNm
    PushUnderLin O
    For Each I In ExistingFny
        Push O, "Worksheet Column : " & QuoteSqBkt(CStr(I))
    Next
    PushUnderLin O
    For Each I In MissFny
        Push O, "Missing Column   : " & QuoteSqBkt(CStr(I))
    Next
    PushMsgUnderLinDbl O, "Columns are missing"
Else
    Push O, "Database : " & A.Name
    Push O, "Table    : " & T
    For Each I In MissFny
        Push O, "Field    : " & QuoteSqBkt(CStr(I))
    Next
    PushMsgUnderLinDbl O, "Above Fields are missing"
End If
DbtMissFny_Er = O
End Function

Function DbtChkFny(A As Database, T$, ExpFny$()) As String()
Dim Miss$(), TFny$(), O$(), I
TFny = DbtFny(A, T)
Miss = AyMinus(ExpFny, TFny)
DbtChkFny = DbtMissFny_Er(A, T, Miss, TFny)
End Function
Function QuoteSqBkt$(A)
QuoteSqBkt = "[" & A & "]"
End Function
Function PushMsgUnderLin(O$(), M$)
Push O, M
Push O, UnderLin(M)
End Function
Function PushUnderLin(O$())
Push O, UnderLin(AyLasEle(O))
End Function
Function PushUnderLinDbl(O$())
Push O, UnderLinDbl(AyLasEle(O))
End Function
Function PushMsgUnderLinDbl(O$(), M$)
Push O, M
Push O, UnderLinDbl(M)
End Function
Function DaoTy_ShtTy$(A As DAO.DataTypeEnum)
Dim O$
Select Case A
Case DAO.DataTypeEnum.dbByte: O = "Byt"
Case DAO.DataTypeEnum.dbLong: O = "Lng"
Case DAO.DataTypeEnum.dbInteger: O = "Int"
Case DAO.DataTypeEnum.dbDate: O = "Dte"
Case DAO.DataTypeEnum.dbText: O = "Txt"
Case DAO.DataTypeEnum.dbBoolean: O = "Yes"
Case DAO.DataTypeEnum.dbDouble: O = "Dbl"
Case DAO.DataTypeEnum.dbCurrency: O = "Cur"
Case DAO.DataTypeEnum.dbMemo: O = "Mem"
Case Else: O = "?" & A & "?"
End Select
DaoTy_ShtTy = O
End Function

Function DaoShtTy_Ty(A) As DAO.DataTypeEnum
Dim O As DAO.DataTypeEnum
Select Case A
Case "Byt": O = DAO.DataTypeEnum.dbByte
Case "Mem": O = DAO.DataTypeEnum.dbMemo
Case "Lng": O = DAO.DataTypeEnum.dbLong
Case "Int": O = DAO.DataTypeEnum.dbInteger
Case "Dte": O = DAO.DataTypeEnum.dbDate
Case "Txt": O = DAO.DataTypeEnum.dbText
Case "Yes": O = DAO.DataTypeEnum.dbBoolean
Case "Dbl": O = DAO.DataTypeEnum.dbDouble
Case "Sng": O = DAO.DataTypeEnum.dbSingle
Case "Cur": O = DAO.DataTypeEnum.dbCurrency
Case Else
    If HasPfx(A, "T") Then
        Dim S As Byte
        S = Val(Mid(A, 2))
        If 1 <= S And S <= 99 Then
            Debug.Print "DaoShtTy_Ty: invalid[" & A & "]        "
        End If
    End If
End Select
DaoShtTy_Ty = O
End Function
Function TimSz_TSz$(A As Date, Sz&)
TimSz_TSz = DteDTim(A) & "." & Sz
End Function
Function DftFfnAy(FfnAy0) As String()
Select Case True
Case IsStr(FfnAy0): DftFfnAy = ApSy(FfnAy0)
Case IsSy(FfnAy0): DftFfnAy = FfnAy0
Case IsArray(FfnAy0): DftFfnAy = AySy(FfnAy0)
End Select
End Function
Property Get FfnCpyToPthIfDif(FfnAy0, Pth$) As String()
Const M_Sam$ = "File is same the one in Path."
Const M_Copied$ = "File is copied to Path."
Const M_NotFnd$ = "File not found, cannot copy to Path."
PthEns Pth
Dim B$, Ay$(), I, O$(), M$(), Msg$
Ay = DftFfnAy(FfnAy0): If Sz(Ay) = 0 Then Exit Property
For Each I In Ay
    Select Case True
    Case FfnIsExist(CStr(I))
        B = Pth & FfnFn(I)
        Select Case True
        Case FfnIsSam(B, CStr(I))
            Msg = M_Sam: GoSub Prt
        Case Else
            Fso.CopyFile I, B, True
            Msg = M_Copied: GoSub Prt
        End Select
    Case Else
        Msg = M_NotFnd: GoSub Prt
        Push O, "File : " & I
    End Select
Next
If Sz(O) > 0 Then
    PushMsgUnderLinDbl O, "Above files not found"
    FfnCpyToPthIfDif = O
End If
Exit Property
Prt:
    Debug.Print FmtQQ("FfnCpyToPthIfDif: ? Path=[?] File=[?]", Msg, Pth, I)
    Return
End Property
Function FfnIsSamMsg(A$, B$, Sz&, Tim$, Optional Msg$) As String()
Dim O$()
Push O, "File 1   : " & A
Push O, "File 2   : " & B
Push O, "File Size: " & Sz
Push O, "File Time: " & Tim
Push O, "File 1 and 2 have same size and time"
If Msg <> "" Then Push O, Msg
FfnIsSamMsg = O
End Function
Function FfnIsSam(A$, B$) As Boolean
If FfnTim(A) <> FfnTim(B) Then Exit Function
If FfnSz(A) <> FfnSz(B) Then Exit Function
FfnIsSam = True
End Function
Function FfnSz&(A)
If FfnIsExist(A) Then
    FfnSz = FileLen(A)
Else
    FfnSz = -1
End If
End Function
Function FfnTim(A) As Date
If FfnIsExist(A) Then FfnTim = FileDateTime(A)
End Function
Function FfnDTim$(A)
If FfnIsExist(A) Then
    FfnDTim = DteDTim(FileDateTime(A))
End If
End Function
Function AyTrim(A) As String()
If Sz(A) = 0 Then Exit Function
Dim O$(), J&, U&
U = UB(A)
ReDim O(U)
For J = 0 To U
    O(J) = Trim(A(J))
Next
AyTrim = O
End Function
Function DbtChkFldType$(A As Database, T$, F, Ty As DAO.DataTypeEnum)
Dim ActTy As DAO.DataTypeEnum
ActTy = A.TableDefs(T).Fields(F).Type
If ActTy <> Ty Then
    DbtChkFldType = FmtQQ("Table[?] field[?] should have type[?], but now it has type[?]", T, F, DaoTy_ShtTy(Ty), DaoTy_ShtTy(ActTy))
End If
End Function
Function OyPrpSy(A, PrpNm$) As String()
If Sz(A) = 0 Then Exit Function
OyPrpSy = ItrPrpSy(A, PrpNm)
End Function
Function OyPrpInto(A, PrpNm$, OInto)
If Sz(A) = 0 Then Exit Function
OyPrpInto = ItrPrpInto(A, PrpNm, OInto)
End Function
Function LnkColAy_ExtNy(A() As LnkCol) As String()
LnkColAy_ExtNy = OyPrpSy(A, "Extnm")
End Function
Function LnkColAy_Ny(A() As LnkCol) As String()
LnkColAy_Ny = OyPrpSy(A, "Nm")
End Function
Sub WbVdtOupNy(A As Workbook, OupNy$())
Dim O$(), N$, B$(), WsCdNy$()
WsCdNy = WbWsCdNy(A)
O = AyMinus(AyAddPfx(OupNy, "WsO"), WsCdNy)
If Sz(O) > 0 Then
    N = "OupNy":  B = OupNy:  GoSub Dmp
    N = "WbCdNy": B = WsCdNy: GoSub Dmp
    N = "Mssing": B = O:      GoSub Dmp
    Stop
    Exit Sub
End If
Exit Sub
Dmp:
Debug.Print UnderLin(N)
Debug.Print N
Debug.Print UnderLin(N)
AyDmp B
Return
End Sub
Function RsDrs(A As DAO.Recordset) As Drs
Dim Fny$(), Dry()
Fny = RsFny(A)
Dry = RsDry(A)
Set RsDrs = Drs(Fny, Dry)
End Function
Function RsDr(A As DAO.Recordset) As Variant()
RsDr = FldsDr(A.Fields)
End Function
Function RsDry(A As DAO.Recordset) As Variant()
Dim O()
Push O, RsFny(A)
With A
    While Not .EOF
        Push O, RsDr(A)
        .MoveNext
    Wend
End With
RsDry = O
End Function
Function LoHasFny(A As ListObject, Fny$()) As Boolean
Dim Miss$(), FnyzLo$()
FnyzLo = LoFny(A)
Miss = AyMinus(Fny, FnyzLo)
If Sz(Miss) > 0 Then Exit Function
LoHasFny = True
End Function
Function WsFstLo(A As Worksheet) As ListObject
Set WsFstLo = ItrFstItm(A.ListObjects)
End Function
Function ItrFstItm(A)
Dim I
For Each I In A
    Asg I, ItrFstItm
Next
End Function
Function DrsNRow&(A As Drs)
DrsNRow = Sz(A.Dry)
End Function
Function SqAddSngQuote(A)
Dim NC%, C%, R&, O
O = A
NC = UBound(A, 2)
For R = 1 To UBound(A, 1)
    For C = 1 To NC
        If IsStr(O(R, C)) Then
            O(R, C) = "'" & O(R, C)
        End If
    Next
Next
SqAddSngQuote = O
End Function
Sub FldsPutSq(A As DAO.Fields, Sq, R&)
Dim C%, F As DAO.Field
C = 1
For Each F In A
    Sq(R, C) = F.Value
    C = C + 1
Next
End Sub
Function RsSq(A As DAO.Recordset) As Variant()
RsSq = DrySq(RsDry(A))
End Function
Sub DbtPutLo(A As Database, T$, Lo As ListObject)
Dim Sq(), Drs As Drs, Rs As DAO.Recordset
Set Rs = DbtRs(A, T)
If Not AyIsEq(RsFny(Rs), LoFny(Lo)) Then
    Debug.Print "--"
    Debug.Print "Rs"
    Debug.Print "--"
    AyDmp RsFny(Rs)
    Debug.Print "--"
    Debug.Print "Lo"
    Debug.Print "--"
    AyDmp LoFny(Lo)
    Stop
End If
Sq = SqAddSngQuote(RsSq(Rs))
LoMin Lo
SqPutAt Sq, Lo.DataBodyRange
End Sub
Sub LoEnsNRow(A As ListObject, NRow&)
LoMin A
Exit Sub
If NRow > 1 Then
    Debug.Print A.InsertRowRange.Address
    Stop
End If
End Sub
Function DrsCol(A As Drs, F) As Variant()
DrsCol = DrsColInto(A, F, EmpAy)
End Function
Function AyIx&(A, M)
Dim J&
For J = 0 To UB(A)
    If A(J) = M Then AyIx = J: Exit Function
Next
AyIx = -1
End Function
Function LoSy(A As ListObject, ColNm$) As String()
Dim Sq()
Sq = A.ListColumns(ColNm).DataBodyRange.Value
LoSy = SqColSy(Sq, 1)
End Function
Function LoFny(A As ListObject) As String()
LoFny = ItrNy(A.ListColumns)
End Function
Sub AyPutLoCol(A, Lo As ListObject, ColNm$)
Dim At As Range, C As ListColumn, R As Range
'AyDmp LoFny(Lo)
'Stop
Set C = Lo.ListColumns(ColNm)
Set R = C.DataBodyRange
Set At = R.Cells(1, 1)
AyPutCol A, At
End Sub
Function AySqH(A) As Variant()
Dim O(), N&, J&
N = Sz(A)
If N = 0 Then Exit Function
ReDim Sq(1 To 1, 1 To N)
For J = 1 To N
    O(1, J) = A(J - 1)
Next
AySqH = O
End Function
Function AySqV(A) As Variant()
Dim O(), N&, J&
N = Sz(A)
If N = 0 Then Exit Function
ReDim O(1 To N, 1 To 1)
For J = 1 To N
    O(J, 1) = A(J - 1)
Next
AySqV = O
End Function
Sub AyPutCol(A, At As Range)
Dim Sq()
Sq = AySqV(A)
RgReSz(At, Sq).Value = Sq
End Sub
Sub AyPutRow(A, At As Range)
Dim Sq()
Sq = AySqH(A)
RgReSz(At, Sq).Value = Sq
End Sub
Function DrsColInto(A As Drs, F, OInto)
Dim O, Ix%, Dry(), Dr
Ix = AyIx(A.Fny, F): If Ix = -1 Then Stop
O = OInto
Erase O
Dry = A.Dry
If Sz(Dry) = 0 Then DrsColInto = O: Exit Function
For Each Dr In Dry
    Push O, Dr(Ix)
Next
DrsColInto = O
End Function
Sub RsBrw(A As DAO.Recordset)
RsBrw_zSingleRec A
End Sub
Sub PrmBrw()
RsBrw TblRs("Prm")
End Sub
Sub RsBrw_zSingleRec(A As DAO.Recordset)
AyBrw RsLy_zSingleRec(A)
End Sub
Function RsNy(A As DAO.Recordset) As String()
RsNy = ItrNy(A.Fields)
End Function
Function RsLy_zSingleRec(A As DAO.Recordset)
RsLy_zSingleRec = NyAv_Ly(RsNy(A), RsVy(A), 0)
End Function
Function DrsColSy(A As Drs, F) As String()
DrsColSy = DrsColInto(A, F, EmpSy)
End Function
Function ObjNm$(A)
If IsNothing(A) Then ObjNm = "#nothing#": Exit Function
On Error GoTo X
ObjNm = A.Name
Exit Function
X:
ObjNm = Err.Description
End Function
Function DbNm$(A As Database)
DbNm = ObjNm(A)
End Function

Function DbtHasLnk(A As Database, T, S$, Cn$)
Dim I As DAO.TableDef
For Each I In A.TableDefs
    If I.Name = T Then
        If I.SourceTableName <> S Then Exit Function
        If EnsSfxSC(I.Connect) <> EnsSfxSC(Cn) Then Exit Function
        DbtHasLnk = True
        Exit Function
    End If
Next
End Function
Sub CrtDtaFb()
If IsDev Then Exit Sub
If FfnIsExist(DtaFb) Then Exit Sub
FbCrt DtaFb
Dim Src, Tar$, TarFb$
TarFb = DtaFb
For Each Src In CcmTny
    Tar = Mid(Src, 2)
    Application.DoCmd.CopyObject TarFb, Tar, acTable, Src
    Debug.Print MsgLin("CrtDtaFb: Cpy [Src] to [Tar]", Src, Tar)
Next
End Sub
Function MsgLin$(A$, ParamArray Ap())
Dim Av(): Av = Ap
MsgLin = MsgAv_Lin(A, Av)
End Function
Sub DbtLnk(A As Database, T, S$, Cn$)
On Error GoTo X
Dim TT As New DAO.TableDef
If DbtHasLnk(A, T, S, Cn) Then
    'Debug.Print MsgLin("DbtLnk: [Tbl] has same [Src] & [Cn] in [Db]", T, S, Cn, DbNm(A))
    Exit Sub
End If
DbDrpTbl A, T
With TT
    .Connect = Cn
    .Name = T
    .SourceTableName = S
    A.TableDefs.Append TT
    'Debug.Print MsgLin("DbtLnk: [Tbl] has linked to [Src] in [Db] with [Cn]", T, S, DbNm(A), Cn)
End With
Exit Sub
X:
Dim M$
M = Err.Description
Er "Cannot create [Table] from [Source] using [CnStr] in [Database].  It gets [error].", _
    T, S, Cn, DbNm(A), M
End Sub
Sub TblLnk(T$, S$, Cn$)
DbtLnk CurrentDb, T, S, Cn
End Sub
Sub TblLnkFb(T, Fb$, Optional FbTbl$)
DbttLnkFb CurrentDb, T, Fb, FbTbl
End Sub
Function CvNothing(A)
If IsEmpty(A) Then Set CvNothing = Nothing: Exit Function
Set CvNothing = A
End Function
Function WbWsCd(A As Workbook, WsCdNm$) As Worksheet
Set WbWsCd = CvNothing(ItrFstPrpEq(A.Sheets, "CodeName", WsCdNm))
End Function
Function WbLasWs(A As Workbook) As Worksheet
Set WbLasWs = A.Sheets(A.Sheets.Count)
End Function
Function WbWs(A As Workbook, WsNm) As Worksheet
Set WbWs = A.Sheets(WsNm)
End Function
Function FxWb(A) As Workbook
Set FxWb = NewXls.Workbooks.Open(A)
End Function
Function WsLo(A As Worksheet, LoNm$) As ListObject
Dim Lo As ListObject
For Each Lo In A.ListObjects
    If Lo.Name = LoNm Then
        Set WsLo = Lo
        Exit Function
    End If
Next
End Function
Function TblPk(T) As String()
TblPk = DbtPk(CurrentDb, T)
End Function
Function TblRg(A$, At As Range) As Range
Set TblRg = DbtRg(CurrentDb, A, At)
End Function
Function DbtRg(A As Database, T, At As Range) As Range
Set DbtRg = SqRg(DbtSq(A, T), At)
End Function
Function AyAddAp(ParamArray Ap())
Dim Av(): Av = Ap
Dim O, J%
O = Ap(0)
For J = 1 To UB(Av)
    PushAy O, Av(J)
Next
AyAddAp = O
End Function
Function AlignL$(A, W%)
AlignL = A & Space(W - Len(A))
End Function

Function AyMapXPSy(A, MapXPFunNm$, P) As String()
AyMapXPSy = AyMapXPInto(A, MapXPFunNm, P, EmpSy)
End Function

Function AyMapXPInto(A, MapXPFunNm$, P, OInto)
Dim O, J&
O = OInto
Erase O
If Sz(A) = 0 Then AyMapXPInto = O: Exit Function
ReDim O(UB(A))
For J = 0 To UB(A)
    Asg Run(MapXPFunNm, A(J), P), O(J)
Next
AyMapXPInto = O
End Function

Function AyAlignL(A) As String()
AyAlignL = AyMapXPSy(A, "AlignL", AyWdt(A))
End Function
Function LSpecLnkColStr$(A)
Dim L$
LSpecAsg A, , L
LSpecLnkColStr = L
End Function
Function LnkColAy_ImpSql$(A() As LnkCol, T$, Optional WhBExpr$)
If FstChr(T) <> ">" Then
    Debug.Print "T must have first char = '>'"
    Stop
End If
Dim Ny$(), ExtNy$(), J%, O$(), S$, N$(), E$()
Ny = LnkColAy_Ny(A)
ExtNy = LnkColAy_ExtNy(A)
N = AyAlignL(Ny)
E = AyAlignL(AyQuoteSqBkt(ExtNy))
Erase O
For J = 0 To UB(Ny)
    If ExtNy(J) = Ny(J) Then
        Push O, FmtQQ("     ?    ?", Space(Len(E(J))), N(J))
    Else
        Push O, FmtQQ("     ? As ?", E(J), N(J))
    End If
Next
S = Join(O, "," & vbCrLf)
LnkColAy_ImpSql = FmtQQ("Select |?| Into [#I?]| From [?] |?", S, RmvFstChr(T), T, PWh(WhBExpr))
End Function
Sub WbMinLo(A As Workbook)
ItrDo A.Sheets, "WsMinLo"
End Sub
Sub WsMinLo(A As Worksheet)
If A.CodeName = "WsIdx" Then Exit Sub
ItrDo A.ListObjects, "LoMin"
End Sub
Sub LoMin(A As ListObject)
Dim R1 As Range, R2 As Range
Set R1 = A.DataBodyRange
If R1.Rows.Count >= 2 Then
    Set R2 = RgRR(R1, 2, R1.Rows.Count)
    R2.Delete
End If
End Sub
Function RgRR(A As Range, R1, R2) As Range
Set RgRR = RgRCRC(A, R1, 1, R2, A.Columns.Count)
End Function
Sub FxMinLo(A)
Dim Wb As Workbook
Set Wb = FxWb(A)
WbMinLo Wb
Wb.Save
Wb.Close
End Sub
Sub PcRfh(A As PivotCache)
A.MissingItemsLimit = xlMissingItemsNone
A.Refresh
End Sub
Sub ItrDo(A, DoNm$)
Dim I
For Each I In A
    Run DoNm, I
Next
End Sub
Sub ItrDoXP(A, DoXPNm$, P)
Dim I
For Each I In A
    Run DoXPNm, I, P
Next
End Sub
Function IsProd() As Boolean
IsProd = Not IsDev
End Function
Function IsDev() As Boolean
Static X As Boolean, Y As Boolean
If Not X Then
    X = True
    Y = Not PthIsExist("N:\SAPAccessReports\")
End If
IsDev = Y
End Function

Sub FunPAy_Do(A, P)
Dim FunP
For Each FunP In A
    Run CStr(FunP), P
Next
End Sub
Function OupFx_Crt$(A)
OupFx_Crt = AttExp("Tp", A)
End Function
Sub TpRfh()
WbVis WbRfh(TpWb)
End Sub
Sub OupFx_Gen(A$, Fb$, ParamArray WbFmtrAp())
Dim Av(): Av = WbFmtrAp
TpWrtFfn A
WbFmt FxRfh(A, Fb), Av
End Sub
Function FxRfh(A, Fb$) As Workbook
Set FxRfh = WbRfh(FxWb(A), Fb)
End Function
Sub WbFmt(A As Workbook, WbFmtrAv())
If True Then
    FunPAy_Do WbFmtrAv, A
Else
    Dim J%
    For J = 0 To UB(WbFmtrAv)
        Run WbFmtrAv(J), A
    Next
End If
WbMax(WbVis(A)).Save
End Sub
Sub TpGenFx(TpFx$, OupFx$, Fb$, ParamArray WbFmtrAp())
Dim Av(): Av = WbFmtrAp
FfnCpy TpFx, OupFx
WbFmt FxRfh(OupFx, Fb), Av
End Sub

Function WbVis(A As Workbook) As Workbook
XlsVis A.Application
Set WbVis = A
End Function
Function CvLo(A) As ListObject
Set CvLo = A
End Function
Function DbOupTny(A As Database) As String()
DbOupTny = DbqSy(A, "Select Name from MSysObjects where Name like '@*' and Type =1")
End Function

Function ObjHasNmPfx(O, NmPfx$) As Boolean
ObjHasNmPfx = HasPfx(ObjNm(O), NmPfx)
End Function

Function OyWhNmHasPfx(A, Pfx$)
OyWhNmHasPfx = OyWhPredXP(A, "ObjHasNmPfx", Pfx)
End Function

Function OyWhPredXP(A, XP$, P)
Dim O, X
O = A
Erase O
For Each X In A
    If Run(XP, X, P) Then
        PushObj O, X
    End If
Next
OyWhPredXP = O
End Function

Function WbOupLoAy(A As Workbook) As ListObject()
WbOupLoAy = OyWhNmHasPfx(WbLoAy(A), "T_")
End Function

Sub FbRplWbLo(Fb$, A As Workbook)
Dim I, Lo As ListObject, Db As Database
Set Db = FbDb(Fb)
For Each I In WbOupLoAy(A)
    Set Lo = I
    DbtRplLo Db, "@" & Mid(Lo.Name, 3), Lo
Next
Db.Close
Set Db = Nothing
End Sub

Function WbRfh(A As Workbook, Optional Fb$) As Workbook
ItrDoXP A.Connections, "WcRfh", Fb
ItrDo A.PivotCaches, "PcRfh"
ItrDo A.Sheets, "WsRfh"
ItrDo WbLoAy(A), "LoRfhAllFmt"
Set WbRfh = A
End Function
Sub WbDltWc(A As Workbook)
ItrDo A.Connections, "WcDlt"
End Sub
Sub ZZ_RplBet()
Dim A$, Exp$, By$, S1$, S2$
S1 = "Data Source="
S2 = ";"
A = "aa;Data Source=???;klsdf"
By = "xx"
Exp = "aa;Data Source=xx;klsdf"
GoSub Tst
Exit Sub
Tst:
Dim Act$
Act = RplBet(A, By, S1, S2)
Debug.Assert Exp = Act
Return
End Sub
Function RplBet$(A$, By$, S1$, S2$)
Dim P1%, P2%, B$, C$

P1 = InStr(A, S1)
If P1 = 0 Then Stop
P2 = InStr(P1 + Len(S1), CStr(A), S2)
If P2 = 0 Then Stop
B = Left(A, P1 + Len(S1) - 1)
C = Mid(A, P2 + Len(S2) - 1)
RplBet = B & By & C
End Function
Function FbWcStr$(A)
FbWcStr = FbOleCnStr(A)
'FbWcStr = FmtQQ("OLEDB;Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False", A)
End Function
Sub WcRfhCnStr(A As WorkbookConnection, Optional Fb$)
If IsNothing(A.OLEDBConnection) Then Exit Sub
If Fb = "" Then Exit Sub
Dim Cn$
Const Ver$ = "0.0.1"
Select Case Ver
Case "0.0.1"
    Dim S$
    S = A.OLEDBConnection.Connection
    Cn = RplBet(S, CStr(Fb), "Data Source=", ";")
Case "0.0.2"
    Cn = FbWcStr(Fb)
End Select
A.OLEDBConnection.Connection = Cn
End Sub

Sub WcRfh(A As WorkbookConnection, Optional Fb$)
If IsNothing(A.OLEDBConnection) Then Exit Sub
WcRfhCnStr A, Fb
A.OLEDBConnection.BackgroundQuery = False
A.OLEDBConnection.Refresh
End Sub

Sub WcDlt(A As WorkbookConnection)
A.Delete
End Sub

Function QtPrpLoFmlVbl$(A As QueryTable)
QtPrpLoFmlVbl = FbtStr_PrpLoFmlVbl(QtFbtStr(A))
End Function

Function CnStr_DtaSrc$(A)
CnStr_DtaSrc = TakBet(A, "Data Source=", ";")
End Function

Sub ZZZ_TakBet()
Dim A$, FmStr, ToStr, Exp$
A = "lkjsdf;dkfjl;Data Source=Johnson;lsdfjldf"
FmStr = "Data Source="
ToStr = ";"
Exp = "Johnson"
GoSub Tst
Exit Sub
Tst:
    Dim Act$
    Act = TakBet(A, FmStr, ToStr)
    Debug.Assert Act = Exp
    Return
End Sub
Function TakBet$(A, FmStr, ToStr)
Dim P1&, P2&
P1 = InStr(A, FmStr): If P1 = 0 Then Exit Function
P2 = InStr(P1, A, ToStr)
Dim FmIx&, L&
FmIx = P1 + Len(FmStr)
L = P2 - FmIx
TakBet = Mid(A, FmIx, L)
End Function
Function TpMainQt() As QueryTable
Set TpMainQt = WbMainQt(TpWb)
End Function
Property Get NowDTim$()
NowDTim = DteDTim(Now)
End Property
Function DteDTim$(Dte)
If Not IsDte(Dte) Then Exit Function
DteDTim = Format(Dte, "YYYY-MM-DD HH:MM:SS")
End Function
Sub ZZ_WbRfhFml()
Dim Wb As Workbook
Set Wb = WbVis(TpWb)
WbRfhLoFml Wb
Stop
End Sub
Sub ZZ_WbRfhLoAlignC()
Dim Wb As Workbook
Set Wb = WbVis(TpWb)
WbRfhLoAlignC Wb
Stop
End Sub

Sub ZZ_WbRfhLoFmt()
Dim Wb As Workbook
Set Wb = WbVis(TpWb)
WbRfhLoFml Wb
WbRfhLoFmt Wb
Stop
End Sub

Sub WbRfhLoFmt(A As Workbook)
AyDo WbLoAy(A), "LoRfhFmt"
End Sub
Function RgNMoreTop(A As Range, Optional N% = 1)
Dim O As Range
Set O = RgRR(A, 1 - N, A.Rows.Count)
Set RgNMoreTop = O
End Function
Sub Z_RgNMoreBelow()
Dim R As Range, Act As Range, Ws As Worksheet
Set Ws = NewWs
Set R = Ws.Range("A3:B5")
Set Act = RgNMoreTop(R, 1)
Debug.Print Act.Address
Stop
Debug.Print RgRR(R, 1, 2).Address
Stop
End Sub
Function RgNMoreBelow(A As Range, Optional N% = 1)
Set RgNMoreBelow = RgRR(A, 1, A.Rows.Count + N)
End Function
Sub RgAsgRCRC(A As Range, OR1, OC1, OR2, OC2)
OR1 = A.Row
OR2 = OR1 + A.Rows.Count - 1
OC1 = A.Column
OC2 = OC1 + A.Columns.Count - 1
End Sub

Sub LoDlt(A As ListObject)
Dim R As Range, R1, C1, R2, C2, Ws As Worksheet
Set Ws = LoWs(A)
Set R = RgNMoreBelow(RgNMoreTop(A.DataBodyRange))
RgAsgRCRC R, R1, C1, R2, C2
A.QueryTable.Delete
WsRCRC(Ws, R1, C1, R2, C2).ClearContents
End Sub

Sub Z_LoReset()
Dim Wb As Workbook, LoAy() As ListObject
Set Wb = FxWb("C:\users\user\desktop\a.xlsx")
WbVis Wb
LoAy = WbLoAy(Wb)
LoReset LoAy(0)
End Sub
Sub LoReset(A As ListObject)
'When LoRfh, if the fields of Db's table has been reorder, the Lo will not follow the order
'Delete the Lo, add back the Wc then WcAt to reset the Lo
'TblNm : from Lo.Name = T_XXX is the key to get the table name.
'Fb    : use WFb
Dim LoNm$, T$, At As Range, Wb As Workbook
Set Wb = LoWb(A)
Set At = RgRC(A.DataBodyRange, 0, 1)
LoNm = A.Name
T = LoNm_TblNm(LoNm)
LoDlt A
WcAt WbAddWc(Wb, WFb, T), At
End Sub
Sub WbRfhLoAlignC(A As Workbook)
AyDo WbLoAy(A), "LoRfhAlignC"
End Sub

Sub WbRfhLoFml(A As Workbook)
AyDo WbTLoAy(A), "LoRfhFml"
End Sub
Function FfnTSz$(A)
If Not FfnIsExist(A) Then Exit Function
FfnTSz = FfnDTim(A) & "." & FfnSz(A)
End Function
Function FfnAsgTSz(A, OTim As Date, OSz&)
If Not FfnIsExist(A) Then
    OTim = 0
    OSz = 0
    Exit Function
End If
OTim = FfnTim(A)
OSz = FfnSz(A)
End Function
Function TSzTim(A) As Date
TSzTim = TakBef(A, ".")
End Function
Function TpMainLo() As ListObject
Set TpMainLo = WbMainLo(TpWb)
End Function
Function TpMainPrpLoFmlVbl$()
TpMainPrpLoFmlVbl = LoPrpLoFmlVbl(TpMainLo)
End Function
Function QtPrpLoFmtVbl$(A As QueryTable)
If IsNothing(A) Then Exit Function
QtPrpLoFmtVbl = FbtStr_PrpLoFmlVbl(QtFbtStr(A))
End Function
Function WbMainLo(A As Workbook) As ListObject
Dim O As Worksheet, Lo As ListObject
Set O = WbMainWs(A):              If IsNothing(O) Then Exit Function
Set WbMainLo = WsLo(O, "T_Main")
End Function
Function WbMainQt(A As Workbook) As QueryTable
Dim Lo As ListObject
Set Lo = WbMainLo(A): If IsNothing(A) Then Exit Function
Set WbMainQt = Lo.QueryTable
End Function
Function WbMainWs(A As Workbook) As Worksheet
Set WbMainWs = WbWsCd(A, "WsOMain")
End Function
Function LoFbtStr$(A As ListObject)
LoFbtStr = QtFbtStr(A.QueryTable)
End Function
Function FbtStr_PrpLoFmlVbl$(A)
Dim Fb$, T$
FbtStr_Asg A, Fb, T
FbtStr_PrpLoFmlVbl = FbtPrpLoFmlVbl(Fb, T)
End Function
Sub FbtStr_Asg(A, OFb$, OT$)
If A = "" Then
    OFb = ""
    OT = ""
    Exit Sub
End If
BrkAsg A, "].[", OFb, OT
If FstChr(OFb) <> "[" Then Stop
If LasChr(OT) <> "]" Then Stop

OFb = RmvFstChr(OFb)
OT = RmvLasChr(OT)
End Sub
Function HasSqBkt(A) As Boolean
HasSqBkt = FstChr(A) = "[" And LasChr(A) = "]"
End Function
Function RmvSqBkt$(A)
If Not HasSqBkt(A) Then Stop
RmvSqBkt = RmvFstLasChr(A)
End Function
Function RmvOptSqBkt$(A)
If Not HasSqBkt(A) Then RmvOptSqBkt = A: Exit Function
RmvOptSqBkt = RmvFstLasChr(A)
End Function
Function LoPrpLoFmlVbl$(A As ListObject)
LoPrpLoFmlVbl = QtPrpLoFmlVbl(LoQt(A))
End Function
Function LoQt(A As ListObject) As QueryTable
On Error Resume Next
Set LoQt = A.QueryTable
End Function
Function TpMainFbtStr$()
Dim Wb As Workbook, Qt As QueryTable
Set Wb = TpWb
Set Qt = WbMainQt(Wb)
TpMainFbtStr = QtFbtStr(Qt)
WbQuit Wb
End Function
Sub WbQuit(A As Workbook)
XlsQuit A.Application
End Sub
Sub XlsQuit(A As Excel.Application)
ItrDo A.Workbooks, "WbClsNoSav"
A.Quit
Set A = Nothing
End Sub
Sub WbClsNoSav(A As Workbook)
A.Close False
End Sub
Function QtFbtStr$(A As QueryTable)
If IsNothing(A) Then Exit Function
Dim Ty As XlCmdType, Tbl$, CnStr$
With A
    Ty = .CommandType
    If Ty <> xlCmdTable Then Exit Function
    Tbl = .CommandText
    CnStr = .Connection
End With
QtFbtStr = FmtQQ("[?].[?]", CnStr_DtaSrc(CnStr), Tbl)
End Function
Sub WsRfh(A As Worksheet)
ItrDo A.QueryTables, "QtRfh"
ItrDo A.PivotTables, "PtRfh"
ItrDo A.ListObjects, "LoRfh"
End Sub

Sub LoRfh(A As Excel.ListObject)
LoReset A
Exit Sub
Dim Qt As QueryTable
Set Qt = LoQt(A)
If IsNothing(Qt) Then Exit Sub
QtRfh Qt
End Sub
Sub QtRfh(A As Excel.QueryTable)
A.BackgroundQuery = False
A.Refresh
End Sub
Sub PtRfh(A As Excel.PivotTable)
A.Update
End Sub
Function WsWb(A As Worksheet) As Workbook
Set WsWb = A.Parent
End Function
Function LoVis(A As ListObject) As ListObject
XlsVis A.Application
Set LoVis = A
End Function
Function WsVis(A As Worksheet)
XlsVis A.Application
Set WsVis = A
End Function
Sub XlsVis(A As Excel.Application)
If Not A.Visible Then A.Visible = True
End Sub
Function SqPutAt(A, At As Range) As Range
Dim O As Range
Set O = RgReSz(At, A)
O.Value = A
Set SqPutAt = O
End Function
Function RgWs(A As Range) As Worksheet
Set RgWs = A.Parent
End Function
Function RgRC(A As Range, R, C) As Range
Set RgRC = A.Cells(R, C)
End Function
Function RgRCRC(A As Range, R1, C1, R2, C2) As Range
Set RgRCRC = RgWs(A).Range(RgRC(A, R1, C1), RgRC(A, R2, C2))
End Function
Function RgReSz(A As Range, Sq) As Range
Set RgReSz = RgRCRC(A, 1, 1, UBound(Sq, 1), UBound(Sq, 2))
End Function
Sub ZZ_TblSq()
Dim A()
A = TblSq("@Oup")
Stop
End Sub
Function NewWb(Optional WsNm$ = "Sheet1") As Workbook
Dim O As Workbook, Ws As Worksheet
Set O = NewXls.Workbooks.Add
Set Ws = WbFstWs(O)
If Ws.Name <> WsNm Then Ws.Name = WsNm
Set NewWb = O
End Function
Function WbFstWs(A As Workbook) As Worksheet
Set WbFstWs = A.Sheets(1)
End Function
Function NewWs(Optional WsNm$ = "Sheet") As Worksheet
Set NewWs = WbFstWs(NewWb(WsNm))
End Function
Function NewA1(Optional WsNm$ = "Sheet1") As Range
Set NewA1 = WsA1(NewWs(WsNm))
End Function
Function SqNewA1(A, Optional WsNm$ = "Data") As Range
Dim A1 As Range
Set A1 = NewA1(WsNm)
Set SqNewA1 = SqPutAt(A, A1)
End Function
Function WsRC(A As Worksheet, R, C) As Range
Set WsRC = A.Cells(R, C)
End Function
Function WsRCRC(A As Worksheet, R1, C1, R2, C2) As Range
Set WsRCRC = A.Range(WsRC(A, R1, C1), WsRC(A, R2, C2))
End Function
Function RgA1LasCell(A As Range) As Range
Dim L As Range, R, C
Set L = A.SpecialCells(xlCellTypeLastCell)
R = L.Row
C = L.Column
Set RgA1LasCell = WsRCRC(RgWs(A), A.Row, A.Column, R, C)
End Function
Function RgLo(A As Range, Optional LoNm$) As ListObject
Dim O As ListObject
Set O = RgWs(A).ListObjects.Add(xlSrcRange, A, , XlYesNoGuess.xlYes)
'LoAutoFit O
If LoNm <> "" Then O.Name = LoNm
Set RgLo = O
End Function
Function RgVis(A As Range) As Range
XlsVis A.Application
Set RgVis = A
End Function
Sub DbttWrtFx(A As Database, TT, Fx$)
DbttWb(A, TT).SaveAs Fx
End Sub
Sub WsClrLo(A As Worksheet)
Dim Ay() As ListObject, J%
Ay = ItrAy(A.ListObjects, Ay)
For J = 0 To UB(Ay)
    Ay(J).Delete
Next
End Sub
Sub TTWrtFx(TT, Fx$)
DbttWrtFx CurrentDb, TT, Fx
End Sub
Function WbAddWs(A As Workbook, Optional WsNm, Optional BefWsNm$, Optional AftWsNm$) As Worksheet
Dim O As Worksheet, Bef As Worksheet, Aft As Worksheet
WbDltWs A, WsNm
Select Case True
Case BefWsNm <> ""
    Set Bef = A.Sheets(BefWsNm)
    Set O = A.Sheets.Add(Bef)
Case AftWsNm <> ""
    Set Aft = A.Sheets(AftWsNm)
    Set O = A.Sheets.Add(, Aft)
Case Else
    Set O = A.Sheets.Add
End Select
O.Name = WsNm
Set WbAddWs = O
End Function
Sub WbDltWs(A As Workbook, WsNm)
If WbHasWs(A, WsNm) Then
    A.Application.DisplayAlerts = False
    WbWs(A, WsNm).Delete
    A.Application.DisplayAlerts = True
End If
End Sub
Function ItrHasNm(A, Nm) As Boolean
Dim I
For Each I In A
    If I.Name = Nm Then ItrHasNm = True: Exit Function
Next
End Function

Function WbHasWs(A As Workbook, WsNm) As Boolean
WbHasWs = ItrHasNm(A.Sheets, WsNm)
End Function

Sub FfnCpy(A, ToFfn$, Optional OvrWrt As Boolean)
If OvrWrt Then FfnDlt ToFfn
FileSystem.FileCopy A, ToFfn
End Sub

Sub FfnDlt(A)
If FfnIsExist(A) Then Kill A
End Sub

Function PthIsExist(A) As Boolean
If A = "" Then Exit Function
On Error Resume Next
PthIsExist = Dir(A, vbDirectory) <> ""
End Function
Function FfnIsExist(A) As Boolean
If A = "" Then Exit Function
On Error Resume Next
FfnIsExist = Dir(A) <> ""
End Function
Function TTWb(TT, Optional UseWc As Boolean) As Workbook
Set TTWb = DbttWb(CurrentDb, TT, UseWc)
End Function
Function DbttWb(A As Database, TT, Optional UseWc As Boolean) As Workbook
Dim O As Workbook
Set O = NewWb
Set DbttWb = WbAddDbtt(O, A, TT, UseWc)
WbWs(O, "Sheet1").Delete
End Function
Function WbA1(A As Workbook, Optional WsNm) As Range
Set WbA1 = WsA1(WbAddWs(A, WsNm))
End Function
Sub DbtRenCol(A As Database, T, Fm, NewCol)
A.TableDefs(T).Fields(Fm).Name = NewCol
End Sub
Function DbDesy(A As Database) As String()
Dim T$(), D$()
T = DbTny(A)
DbDesy = AyRmvEmp(AyMapPXSy(T, "DbtTblDes", A))
End Function
Function AyRmvEmp(A)
Dim O: O = AyCln(A)
If Sz(A) > 0 Then
    Dim X
    For Each X In A
        PushNonEmpty O, X
    Next
End If
AyRmvEmp = O
End Function
Function DbtTblDes$(A As Database, T)
Dim D$
D = DbtDes(A, T)
If D = "" Then Exit Function
DbtTblDes = T & " " & D
End Function
Function DbtAt_Lo(A As Database, T$, At As Range, Optional UseWc As Boolean) As ListObject
Dim N$, Q As QueryTable
N = TblNm_LoNm(T)
If UseWc Then
    Set Q = RgWs(At).ListObjects.Add(SourceType:=0, Source:=FbAdoCnStr(A.Name), Destination:=At).QueryTable
    With Q
        .CommandType = xlCmdTable
        .CommandText = T
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = T
        .Refresh BackgroundQuery:=False
    End With
    Exit Function
End If
Set DbtAt_Lo = RgLo(DbtRg(A, T, At), N)
End Function
Function LoWb(A As ListObject) As Workbook
Set LoWb = LoWs(A).Parent
End Function
Function WbAddDbt(A As Workbook, Db As Database, T$, Optional UseWc As Boolean) As Workbook
Set WbAddDbt = LoWb(DbtAt_Lo(Db, T, WbA1(A, T), UseWc))
End Function
Function TblNm_LoNm$(TblNm)
TblNm_LoNm = "T_" & RmvFstNonLetter(TblNm)
End Function
Function LoNm_TblNm$(LoNm)
If Not HasPfx(LoNm, "T_") Then Stop
LoNm_TblNm = "@" & RmvPfx(LoNm, "T_")
End Function
Sub AyDoPPXP(A, PPXP$, P1, P2, P3)
Dim X
For Each X In A
    Run PPXP, P1, P2, X, P3
Next
End Sub

Function WbAddDbtt(A As Workbook, Db As Database, TT, Optional UseWc As Boolean) As Workbook
AyDoPPXP CvTT(TT), "WbAddDbt", A, Db, UseWc
Set WbAddDbtt = A
End Function

Sub ZZ_RsAsg()
Dim Y As Byte, M As Byte
RsAsg TblRs("YM"), Y, M
Stop
End Sub
Sub RsAsg(A As DAO.Recordset, ParamArray OAp())
Dim F As DAO.Field, J%, U%
Dim Av(): Av = OAp
U = UB(Av)
For Each F In A.Fields
    OAp(J) = F.Value
    If J = U Then Exit Sub
    J = J + 1
Next
End Sub
Function DbqLngAy(A As Database, Sql) As Long()
DbqLngAy = RsLngAy(A.OpenRecordset(Sql))
End Function
Function LinesSrt$(A)
LinesSrt = JnCrLf(AySrt(LinesSplit(A)))
End Function
Function LinesSplit(A) As String()
LinesSplit = SplitCrLf(A)
End Function
Function AySrt(A)
If Sz(A) = 0 Then Exit Function
Dim O: O = A
AyQSrt O, 0, UB(A)
AySrt = O
End Function
Sub ZZ_AySrt()
Dim A, Exp
A = Array(9, 2, 4, 3, 4)
Exp = Array(2, 3, 4, 4, 9)
GoSub Tst
Exit Sub
Tst:
Dim Act
Act = AySrt(A)
Debug.Assert AyIsEq(Act, Exp)
Return
End Sub
Sub AyQSrt(A, L&, H&)
If L >= H Then Exit Sub
Dim P&
P = AyPartition(A, L, H)
AyQSrt A, L, P
AyQSrt A, P + 1, H
End Sub
Function AyReverse(A)
Dim O: O = A
Dim J&, U&
U = UB(O)
For J = 0 To U
    O(J) = A(U - J)
Next
AyReverse = O
End Function
Function AyPartition&(A, L&, H&)
Dim V, I&, J&, X
V = A(L)
I = L - 1
J = H + 1
Dim Z&
Do
    Z = Z + 1
    If Z > 1000 Then Stop
    Do
        I = I + 1
    Loop Until A(I) >= V
    
    Do
        J = J - 1
    Loop Until A(J) <= V

    If I >= J Then
        AyPartition = J
        Exit Function
    End If

     X = A(I)
     A(I) = A(J)
     A(J) = X
Loop
End Function
Function DbStru$(A As Database)
DbStru = DbttStru(A, DbTny(A))
End Function
Function DbTny(A As Database) As String()
DbTny = DbqSy(A, "Select Name from MSysObjects where Type in (1,6) and Name not Like 'MSys*' and Name not Like 'f_????????????????????????????????_*' and Name not like '~TMP*'")
End Function
Function IsPfx(A$, Pfx$) As Boolean
IsPfx = Left(A, Len(Pfx)) = Pfx
End Function
Function DbtNRec&(A As Database, T)
DbtNRec = DbqV(A, FmtQQ("Select Count(*) from [?]", T))
End Function
Function DbtCsv(A As Database, T) As String()
DbtCsv = RsCsvLy(DbtRs(A, T))
End Function
Function DbtLo(A As Database, T$, At As Range) As ListObject
Set DbtLo = SqLo(DbtSq(A, T), At, TblNm_LoNm(T))
End Function
Function DSpecNm$(A)
DSpecNm = TakAftDotOrAll(LinT1(A))
End Function
Function TakAftDotOrAll$(A)
TakAftDotOrAll = TakAftOrAll(A, ".")
End Function
Function LoWs(A As ListObject) As Worksheet
Set LoWs = A.Parent
End Function
Function JnSC$(A)
JnSC = Join(A, ";")
End Function
Function DbtRs(A As Database, T) As DAO.Recordset
Set DbtRs = A.OpenRecordset(T)
End Function
Function TblRs(T) As DAO.Recordset
Set TblRs = DbtRs(CurrentDb, T)
End Function
Sub TimFn(FnNm$)
Dim A!, B!
A = Timer
Run FnNm
B = Timer
Debug.Print FnNm, B - A
End Sub
Function RsCsvLyByFny0(A As DAO.Recordset, Fny0) As String()
Dim Fny$(), Flds As Fields, F
Dim O$(), J&, I%, UFld%, Dr()
Fny = CvNy(Fny0)
UFld = A.Fields.Count - 1
While Not A.EOF
    J = J + 1
    If J Mod 5000 = 0 Then Debug.Print "RsCsvLy: " & J
    If J > 100000 Then Stop
    ReDim Dr(UFld)
    I = 0
    Set Flds = A.Fields
    For Each F In Fny
        Dr(I) = VarCsv(Flds(F).Value)
        I = I + 1
    Next
    Push O, Join(Dr, ",")
    A.MoveNext
Wend
RsCsvLyByFny0 = O
End Function
Function RsCsvLy(A As DAO.Recordset) As String()
Dim O$(), J&, I%, UFld%, Dr(), F As DAO.Field
UFld = A.Fields.Count - 1
While Not A.EOF
    J = J + 1
    If J Mod 5000 = 0 Then Debug.Print "RsCsvLy: " & J
    If J > 100000 Then Stop
    ReDim Dr(UFld)
    I = 0
    For Each F In A.Fields
        Dr(I) = VarCsv(F.Value)
        I = I + 1
    Next
    Push O, Join(Dr, ",")
    A.MoveNext
Wend
RsCsvLy = O
End Function

Function TblNRow&(T$, Optional WhBExpr$)
TblNRow = DbtNRow(CurrentDb, T, WhBExpr)
End Function
Function PWh$(WhBExpr$)
If WhBExpr = "" Then Exit Function
PWh = PSep & "Where" & PSep1 & WhBExpr
End Function
Function DbtNRow&(A As Database, T$, Optional WhBExpr$)
Dim S$
S = "Select Count(*)" & PFm(T) & PWh(WhBExpr)
DbtNRow = DbqLng(A, S)
End Function
Function TblNCol&(T)
TblNCol = DbtNCol(CurrentDb, T)
End Function
Function DbtNCol&(A As Database, T)
DbtNCol = A.OpenRecordset(T).Fields.Count
End Function
Function TblSq(A) As Variant()
TblSq = DbtSq(CurrentDb, A)
End Function
Function DbtSq(A As Database, T, Optional ReSeqSpec$) As Variant()
Dim Q$
Q = QSel(T, ReSeqSpec_Fny(ReSeqSpec))
DbtSq = RsSq(DbqRs(A, Q))
End Function
Sub ZZ_QSel()
Debug.Print QSel("A")
End Sub
Function QSel$(T, Optional Fny0, Optional FldExprDic As Dictionary)
QSel = PSel(Fny0, FldExprDic) & PFm(T)
End Function
Function PFm$(T)
PFm = PSep & "From [" & T & "]"
End Function
Function PFmAlias$(T$, Alias$)
PFmAlias = PFm(T) & " " & Alias
End Function
Function PSel$(Fny0, Optional FldExprDic As Dictionary)
Dim Fny$()
Fny = CvNy(Fny0)
If Sz(Fny) = 0 Then
    PSel = "Select *"
    Exit Function
End If
PSel = "Select " & JnComma(CvNy(Fny0))
End Function
Function PAddCol$(Fny0, FldDfnDic As Dictionary)
Dim Fny$(), O$(), J%
Fny = CvNy(Fny0)
ReDim O(UB(Fny))
For J = 0 To UB(Fny)
    O(J) = Fny(J) & " " & FldDfnDic(Fny(J))
Next
PAddCol = PSep & "Add Column " & JnComma(O)
End Function
Function FxWs(A, Optional WsNm$ = "Data") As Worksheet
Set FxWs = WbWs(FxWb(A), WsNm)
End Function
Sub FldsPutSq1(A As DAO.Fields, Sq, Optional R& = 1, Optional NoTxtSngQ As Boolean)
DrPutSq FldsDr(A), Sq, R, NoTxtSngQ
End Sub
Sub DrPutSq(A, Sq, Optional R& = 1, Optional NoTxtSngQ As Boolean)
Dim J%, I
If NoTxtSngQ Then
    For Each I In A
        J = J + 1
        Sq(R, J) = I
    Next
    Exit Sub
End If
For Each I In A
    J = J + 1
    If IsStr(I) Then
        Sq(R, J) = "'" & I
    Else
        Sq(R, J) = I
    End If
Next
End Sub
Sub RsPutSq(A As DAO.Recordset, Sq, R&, Optional NoTxtSngQ As Boolean)
FldsPutSq1 A.Fields, Sq, R, NoTxtSngQ
End Sub
Function WsRCC(A As Worksheet, R, C1, C2) As Range
Set WsRCC = WsRCRC(A, R, C1, R, C2)
End Function
Function WsCC(A As Worksheet, C1, C2) As Range
Set WsCC = WsRCC(A, 1, C1, C2).EntireColumn
End Function
Function WsRR(A As Worksheet, R1&, R2&) As Range
Set WsRR = A.Rows(R1 & ":" & R2)
End Function
Function WsA1(A As Worksheet) As Range
Set WsA1 = A.Cells(1, 1)
End Function
Function FxLo(A$, Optional WsNm$ = "Data", Optional LoNm$ = "Data") As ListObject
Set FxLo = WsLo(WbWs(FxWb(A), WsNm), LoNm)
End Function
Function TblCnStr$(T$)
TblCnStr = CurrentDb.TableDefs(T).Connect
End Function
Function AutoExec()
'D "AutoExec:"
'D "-Before LnkCcm: CnSy--------------------------"
'D CnSy
'D "-Before LnkCcm: Srcy--------------------------"
'D Srcy
'
LnkCcm
'D "-After LnkCcm: CnSy--------------------------"
'D CnSy
'D "-After LnkCcm: Srcy--------------------------"
'D Srcy
End Function
Function TblSrc$(T$)
TblSrc = CurrentDb.TableDefs(T).SourceTableName
End Function
Property Get Srcy() As String()
Srcy = DbSrcy(CurrentDb)
End Property
Function DbSrcy(A As Database) As String()
Dim S()
Dim T$()
T = AyQuoteSqBkt(DbTny(A))
S = AyMapPX(T, "DbtSrc", A)
DbSrcy = AyabNonEmpBLy(T, S)
End Function
Function DbqLng&(A As Database, Sql)
DbqLng = DbqV(A, Sql)
End Function
Function SqlLng&(A)
SqlLng = DbqLng(CurrentDb, A)
End Function
Function QQSqlV(A, ParamArray Ap())
Dim Av(): Av = Ap
QQSqlV = SqlV(FmtQQAv(A, Av))
End Function
Function SqlV(A)
SqlV = DbqV(CurrentDb, A)
End Function
Sub ZZ_DbqV()
Dim A
A = SqlV("Select Fx from OHYM")
Stop
End Sub
Function DbqTim(A As Database, Sql) As Date
DbqTim = DbqV(A, Sql)
End Function
Function DbqV(A As Database, Sql)
With A.OpenRecordset(Sql)
    If .EOF Then Exit Function
    DbqV = .Fields(0).Value
End With
End Function
Function TblNRec&(A)
TblNRec = SqlLng(FmtQQ("Select Count(*) from [?]", A))
End Function
Function ErzFileNotFound(FfnAy0) As String()
Dim Ay$(), I, O$()
Ay = DftFfnAy(FfnAy0)
If Sz(Ay) = 0 Then Exit Function
For Each I In Ay
    If Not FfnIsExist(CStr(I)) Then
        Push O, I
    End If
Next
If Sz(O) = 0 Then Exit Function
ErzFileNotFound = MsgAp_Ly("[File(s)] not found", O)
End Function
Function LinIsPrpLin(ByVal A) As Boolean
Const C$ = "Property Get "
A = RmvPfx(A, "Private ")
A = RmvPfx(A, "Public ")
If Not HasPfx(A, C) Then Exit Function
A = RmvPfx(A, C)
A = RmvNm(A)
A = RmvTyChr(A)
If Left(A, 2) <> "()" Then Exit Function
LinIsPrpLin = True
End Function
Function LinPrpNm$(ByVal A)
Const C$ = "Property Get "
A = RmvPfx(A, "Private ")
If Not HasPfx(A, C) Then Exit Function
A = RmvPfx(A, C)
LinPrpNm = LinNm(A)
End Function
Sub FfnAssExist(A)
If AyBrwEr(ErzFileNotFound(A)) Then Stop
End Sub
Sub DbtLnkFx(A As Database, T, Fx, Optional WsNm$ = "Sheet1")
Dim Cn$: Cn = FxDaoCnStr(Fx)
Dim Src$: Src = WsNm & "$"
DbtLnk A, T, Src, Cn
End Sub
Function AyIsSrt(A) As Boolean
AyIsSrt = True
If Sz(A) = 0 Then Exit Function
Dim J&, Las
Las = A(0)
For J = 1 To UB(A)
    If Las >= A(J) Then AyIsSrt = True: Exit Function
Next
End Function
Function RmvPfx$(A, Pfx)
If HasPfx(A, Pfx) Then RmvPfx = Mid(A, Len(Pfx) + 1) Else RmvPfx = A
End Function
Function RmvSfx$(A, Sfx)
If HasSfx(A, Sfx) Then RmvSfx = Left(A, Len(A) - Len(Sfx))
End Function
Sub TTLnkFb(TT, Fb$, Optional Fbtt)
DbttLnkFb CurrentDb, TT, Fb, Fbtt
End Sub
Sub DbttLnkFb(A As Database, TT, Fb$, Optional Fbtt)
Dim Tny$(), FbTny$()
Tny = CvNy(TT)
FbTny = CvNy(Fbtt)
    Select Case True
    Case Sz(FbTny) = Sz(Tny)
    Case Sz(FbTny) = 0:  FbTny = Tny
    Case Else:           Er "[TT]-[Sz1] and [Fbtt]-[Sz2] are diff.  (@DbttLnkFb)", TT, Sz(Tny), Fbtt, Sz(FbTny)
    End Select
Dim Cn$: Cn = FbDaoCnStr(Fb)
Dim J%
For J = 0 To UB(Tny)
    DbtLnk A, Tny(J), FbTny(J), Cn
Next
End Sub
Sub TblLnkFx(T$, Fx$, Optional WsNm$ = "Sheet1")
DbtLnkFx CurrentDb, T, Fx, WsNm
End Sub
Function FbDaoCnStr$(A)
FbDaoCnStr = ";DATABASE=" & A & ";"
End Function

Function AyHasT1(A, T1) As Boolean

End Function
Function AyHas(A, M) As Boolean
Dim I
If Sz(A) = 0 Then Exit Function
For Each I In A
    If I = M Then
        AyHas = True
        Exit Function
    End If
Next
End Function

Function AyQuoteSqBkt(A) As String()
AyQuoteSqBkt = AyQuote(A, "[]")
End Function
Function ItrWhPrpIsTrueInto(A, P, OInto)
Dim O: O = OInto: Erase O
Dim X
For Each X In A
    If ObjPrp(A, P) Then
        Push O, X
    End If
Next
ItrWhPrpIsTrueInto = O
End Function

Function ItrWhPrpIsTrue(A, P)
ItrWhPrpIsTrue = ItrWhPrpIsTrueInto(A, P, EmpAy)
End Function

Function DbtPk(A As Database, T) As String()
Dim I As DAO.Index
Set I = DbtPIdx(A, T): If IsNothing(I) Then Exit Function
DbtPk = ItrNy(I.Fields)
End Function
Function DbtPIdx(A As Database, T) As DAO.Index
Dim O As DAO.Index
For Each O In A.TableDefs(T).Indexes
    If O.Primary Then Set DbtPIdx = O: Exit Function
Next
End Function
Function AyQuoteSng(A) As String()
AyQuoteSng = AyQuote(A, "'")
End Function

Function DbtStru$(A As Database, T$)
Dim F$(), X$(), Y$(), XX$, YY$
F = DbtFny(A, T)
If DbtIsXls(A, T) Then
    F = AyQuoteSqBkt(F)
    DbtStru = T & ": " & JnSpc(F)
    Exit Function
End If
X = DbtPk(A, T)
Y = AyMinus(F, X)
If Sz(X) > 0 Then
    XX = JnSpc(X) & " | "
End If
YY = JnSpc(Y)
DbtStru = T & ": " & XX & YY
End Function
Function DbttStru$(A As Database, TT)
Dim Tny$(), O$(), J%
Tny = AySrt(CvNy(TT))
For J = 0 To UB(Tny)
    Push O, DbtStru(A, Tny(J))
Next
DbttStru = JnCrLf(O)
End Function
Sub DbtfChgDteToTxt(A As Database, T$, F)
A.Execute FmtQQ("Alter Table [?] add column [###] text(12)", T)
A.Execute FmtQQ("Update [?] set [###] = Format([?],'YYYY-MM-DD')", T, F)
A.Execute FmtQQ("Alter Table [?] Drop Column [?]", T, F)
A.Execute FmtQQ("Alter Table [?] Add Column [?] text(12)", T, F)
A.Execute FmtQQ("Update [?] set [?] = [###]", T, F)
A.Execute FmtQQ("Alter Table [?] Drop Column [###]", T)
End Sub
Function JnComma$(A)
JnComma = Join(A, ",")
End Function
Function JnSpc$(A)
JnSpc = Join(A, " ")
End Function
Function UB&(A)
UB = Sz(A) - 1
End Function

Sub PushNonEmpty(O, A)
If A = "" Then Exit Sub
Push O, A
End Sub
Function DaoTy_Str$(T As DAO.DataTypeEnum)
Dim O$
Select Case T
Case DAO.DataTypeEnum.dbBoolean: O = "Boolean"
Case DAO.DataTypeEnum.dbDouble: O = "Double"
Case DAO.DataTypeEnum.dbText: O = "Text"
Case DAO.DataTypeEnum.dbDate: O = "Date"
Case DAO.DataTypeEnum.dbByte: O = "Byte"
Case DAO.DataTypeEnum.dbInteger: O = "Int"
Case DAO.DataTypeEnum.dbLong: O = "Long"
Case DAO.DataTypeEnum.dbDouble: O = "Doubld"
Case DAO.DataTypeEnum.dbDate: O = "Date"
Case DAO.DataTypeEnum.dbDecimal: O = "Decimal"
Case DAO.DataTypeEnum.dbCurrency: O = "Currency"
Case DAO.DataTypeEnum.dbSingle: O = "Single"
Case DAO.DataTypeEnum.dbAttachment: O = "Attachment"
Case DAO.DataTypeEnum.dbMemo: O = "Memo"
Case DAO.DataTypeEnum.dbLongBinary: O = "LongBinary"
Case DAO.DataTypeEnum.dbBinary: O = "Binary"
Case DAO.DataTypeEnum.dbGUID: O = "GUID"
Case Else: Stop
End Select
DaoTy_Str = O
End Function
Function DbqryRs(A As Database, Q) As DAO.Recordset
Set DbqryRs = A.QueryDefs(Q).OpenRecordset
End Function
Function RplVBar$(A)
RplVBar = Replace(A, "|", vbCrLf)
End Function
Function Sz&(A)
On Error Resume Next
Sz = UBound(A) + 1
End Function
Function AyBrwEr(A) As Boolean
If Sz(A) = 0 Then Exit Function
AyBrwEr = True
AyBrw A
End Function
Sub AyBrw(A)
If Sz(A) = 0 Then Exit Sub
StrBrw Join(A, vbCrLf)
End Sub
Function TFTy(T$, F$) As DAO.DataTypeEnum
TFTy = DbtfTy(CurrentDb, T, F)
End Function
Function DbtfTy(A As Database, T$, F$) As DAO.DataTypeEnum
DbtfTy = A.TableDefs(T).Fields(F).Type
End Function
Function DbtfTyStr$(A As Database, T$, F$)
DbtfTyStr = DaoTy_Str(DbtfTy(A, T, F))
End Function
Function StrWrt$(A, Ft$, Optional IsNotOvrWrt As Boolean)
Fso.CreateTextFile(Ft, Overwrite:=Not IsNotOvrWrt).Write A
StrWrt = Ft
End Function
Sub FtBrw(A)
If FfnPing(A) Then Exit Sub
'Shell "code.cmd """ & A & """", vbHide
Shell "notepad.exe """ & A & """", vbMaximizedFocus
End Sub
Function JnCrLf$(A)
JnCrLf = Join(A, vbCrLf)
End Function
Function AyWrt$(A, Ft$)
AyWrt = StrWrt(JnCrLf(A), Ft)
End Function

Sub StrBrw(A)
Dim T$
T = TmpFt
StrWrt A, T
FtBrw T
End Sub
Function TmpFxm$(Optional Fdr$, Optional Fnn0$)
TmpFxm = TmpFfn(".xlsm", Fdr, Fnn0)
End Function

Function TmpFfn$(Optional Ext$, Optional Fdr$, Optional Fnn0$)
Dim Fnn$
If Fnn0 = "" Then
    Fnn = TmpNm
Else
    Fnn = Fnn0
End If
TmpFfn = TmpPth(Fdr) & Fnn & Ext
End Function

Sub FbBrw(A)
Acs.OpenCurrentDatabase A
Acs.Visible = True
End Sub
Function TmpFb$(Optional Fdr$, Optional Fnn$)
TmpFb = TmpFfn(".accdb", Fdr, Fnn)
End Function

Function TmpFt$(Optional Fdr$, Optional Fnn$)
TmpFt = TmpFfn(".txt", Fdr, Fnn)
End Function

Function TmpCmd$(Optional Fdr$, Optional Fnn$)
TmpCmd = TmpFfn(".cmd", Fdr, Fnn)
End Function

Function TmpFx$(Optional Fdr$, Optional Fnn$)
TmpFx = TmpFfn(".xlsx", Fdr, Fnn)
End Function

Function TmpNm$()
Static X&
TmpNm = "T" & Format(Now(), "YYYYMMDD_HHMMSS") & "_" & X
X = X + 1
End Function

Function TmpPth$(Optional Fdr$)
Dim O$
    O = WPth
    If Fdr <> "" Then
        O = WPth & Fdr & "\"
        PthEns O
    End If
    O = O & TmpNm & "\"
    PthEns O
TmpPth = O
End Function
Function DbtUpdToDteFld__1(A As Database, T$, KeyFld$, FmDteFld$) As Date()
Dim K$(), FmDte() As Date, ToDte() As Date, J&, CurKey$, NxtKey$, NxtFmDte As Date
With DbtRs(A, T)
    While Not .EOF
        Push FmDte, .Fields(FmDteFld).Value
        Push K, .Fields(KeyFld).Value
        .MoveNext
    Wend
End With
Dim U&
U = UB(K)
ReDim ToDte(U)
For J = 0 To U - 1
    CurKey = K(J)
    NxtKey = K(J + 1)
    NxtFmDte = FmDte(J + 1)
    If CurKey = NxtKey Then
        ToDte(J) = DateAdd("D", -1, NxtFmDte)
    Else
        ToDte(J) = DateSerial(2099, 12, 31)
    End If
Next
ToDte(U) = DateSerial(2099, 12, 31)
DbtUpdToDteFld__1 = ToDte
End Function
Sub ZZ_DbtUpdToDteFld()
DoCmd.RunSQL "Select * into [#A] from ZZ_DbtUpdToDteFld order by Sku,PermitDate"
DbtUpdToDteFld CurrentDb, "#A", "PermitDateEnd", "Sku", "PermitDate"
Stop
TTDrp "#A"
End Sub
Sub DbtUpdToDteFld(A As Database, T$, ToDteFld$, KeyFld$, FmDteFld$)
Dim ToDte() As Date, J&
ToDte = DbtUpdToDteFld__1(A, T, KeyFld, FmDteFld)
With DbtRs(A, T)
    While Not .EOF
        .Edit
        .Fields(ToDteFld).Value = ToDte(J): J = J + 1
        .Update
        .MoveNext
    Wend
    .Close
End With
End Sub

Function LinT1$(ByVal A)
LinT1 = LinShiftTerm(CStr(A))
End Function

Property Get TblImpSpec(T$, LnkSpec$, Optional WhBExpr$) As TblImpSpec
Dim O As New TblImpSpec
Set TblImpSpec = O.Init(T, LnkSpec$, WhBExpr)
End Property

Function TmpHom$()
Static X$
If X = "" Then
    X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
End If
TmpHom = X
End Function

Function FmtQQ$(QQVbl$, ParamArray Ap())
Dim Av(): Av = Ap
FmtQQ = FmtQQAv(QQVbl, Av)
End Function

Function SqlDry(A) As Variant()
SqlDry = DbqDry(CurrentDb, A)
End Function
Function DbqDry(A As Database, Sql) As Variant()
Dim O()
With DbqRs(A, Sql)
    While Not .EOF
        Push O, FldsDr(.Fields)
        .MoveNext
    Wend
    .Close
End With
DbqDry = O
End Function
Function Xls(Optional Vis As Boolean) As Excel.Application
Static X As Boolean, Y As Excel.Application
Dim J%
Beg:
    J = J + 1
    If J > 10 Then Stop
If Not X Then
    X = True
    Set Y = New Excel.Application
End If
On Error GoTo XX
Dim A$
A = Y.Name
Set Xls = Y
If Vis Then XlsVis Y
Exit Function
XX:
    X = True
    GoTo Beg
End Function
Sub AcsQuit()
Dim A As Access.Application
Set A = Acs
A.Quit
Set A = Nothing
End Sub
Function DbtPutAtByCn(A As Database, T$, At As Range, Optional LoNm0$) As ListObject
If FstChr(T) <> "@" Then Stop
Dim LoNm$, Lo As ListObject
If LoNm0 = "" Then
    LoNm = "Tbl" & RmvFstChr(T)
Else
    LoNm = LoNm0
End If
Dim AtA1 As Range, CnStr, Ws As Worksheet
Set AtA1 = RgRC(At, 1, 1)
Set Ws = RgWs(At)
With Ws.ListObjects.Add(SourceType:=0, Source:=Array( _
        FmtQQ("OLEDB;Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share D", A.Name) _
        , _
        "eny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Databa" _
        , _
        "se Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Je" _
        , _
        "t OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Com" _
        , _
        "pact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=" _
        , _
        "False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False" _
        ), Destination:=AtA1).QueryTable '<---- At
        .CommandType = xlCmdTable
        .CommandText = Array(T) '<-----  T
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = LoNm '<------------ LoNm
        .Refresh BackgroundQuery:=False
    End With

End Function
Function NewXls(Optional Vis As Boolean) As Excel.Application
Dim O As New Excel.Application
If Vis Then O.Visible = True
Set NewXls = O
End Function
Function SqlStrCol(A) As String()
SqlStrCol = RsStrCol(CurrentDb.OpenRecordset(A))
End Function
Sub DicDmp(A As Dictionary)
Dim K
For Each K In A
    Debug.Print K, A(K)
Next
End Sub

Sub SqlAy_Run(SqlAy$())
Dim I
For Each I In SqlAy
    DoCmd.RunSQL I
Next
End Sub

Function RsStrCol(A As DAO.Recordset) As String()
Dim O$()
With A
    While Not .EOF
        Push O, .Fields(0).Value
        .MoveNext
    Wend
End With
RsStrCol = O
End Function
Function SqColInto(A, C%, OInto) As String()
Dim O
O = OInto
Erase O
Dim NR&, J&
NR = UBound(A, 1)
ReDim O(NR - 1)
For J = 1 To NR
    O(J - 1) = A(J, C%)
Next
SqColInto = O
End Function
Function SqColSy(A, C%) As String()
SqColSy = SqColInto(A, C, EmpSy)
End Function
Function AtVBar(A As Range) As Range
If IsEmpty(A.Value) Then Stop
If IsEmpty(RgRC(A, 2, 1).Value) Then
    Set AtVBar = RgRC(A, 1, 1)
    Exit Function
End If
Set AtVBar = RgCRR(A, 1, 1, A.End(xlDown).Row - A.Row + 1)
End Function
Function RgCRR(A As Range, C, R1, R2) As Range
Set RgCRR = RgRCRC(A, R1, C, R2, C)
End Function
Function SqSyV(A) As String()
SqSyV = SqColSy(A, 1)
End Function
Sub RgFillCol(A As Range)
Dim Rg As Range
Dim Sq()
Sq = SqzVBar(A.Rows.Count)
RgReSz(A, Sq).Value = Sq
End Sub
Sub RgFillRow(A As Range)
Dim Rg As Range
Dim Sq()
Sq = SqzHBar(A.Rows.Count)
RgReSz(A, Sq).Value = Sq
End Sub
Function SqzVBar(N%) As Variant()
Dim O(), J%
ReDim O(1 To N, 1 To 1)
For J = 1 To N
    O(J, 1) = J
Next
SqzVBar = O
End Function
Function SqzHBar(N%) As Variant()
Dim O(), J%
ReDim O(1 To 1, 1 To N)
For J = 1 To N
    O(1, J) = J
Next
SqzHBar = O
End Function
Sub FxOpn(A)
If Not FfnIsExist(A) Then
    MsgBox "File not found: " & vbCrLf & vbCrLf & A
    Exit Sub
End If
Dim C$
C = FmtQQ("Excel ""?""", A)
Debug.Print C
Shell C, vbMaximizedFocus
'Xls(Vis:=True).Workbooks.Open A
End Sub
Function AyQuote(A, Q$) As String()
If Sz(A) = 0 Then Exit Function
Dim Q1$, Q2$
Select Case True
Case Len(Q) = 1: Q1 = Q: Q2 = Q
Case Len(Q) = 2: Q1 = Left(Q, 1): Q2 = Right(Q, 1)
Case Else: Stop
End Select

Dim I, O$()
For Each I In A
    Push O, Q1 & I & Q2
Next
AyQuote = O
End Function
Function FldsDr(A As DAO.Fields) As Variant()
FldsDr = ItrVy(A)
End Function
Function SubStrCnt%(A, SubStr$)
Dim J&, O%, P%, L%
L = Len(SubStr)
P = InStr(A, SubStr)
While P > 0
    O = O + 1
    J = J + 1: If J > 100000 Then Stop
    P = InStr(P + L, A, SubStr)
Wend
SubStrCnt = O
End Function
Function RgCC(A As Range, C1, C2) As Range
Set RgCC = RgRCRC(A, 1, C1, A.Rows.Count, C2)
End Function

Sub ZZ_FmtQQAv()
Debug.Print FmtQQ("klsdf?sdf?dsklf", 2, 1)
End Sub
Function FmtQQAv$(QQVbl, Av())
Dim O$, I, Cnt
O = Replace(QQVbl, "|", vbCrLf)
Cnt = SubStrCnt(QQVbl, "?")
If Cnt <> Sz(Av) Then
    MsgBrw "[QQVal] has [N-?], but not match with [Av]-[Sz]", QQVbl, Cnt, Av, Sz(Av)
    Stop
    Exit Function
End If
For Each I In Av
    O = Replace(O, "?", I, Count:=1)
Next
FmtQQAv = O
End Function
Sub PushAy(O, A)
If Sz(A) = 0 Then Exit Sub
Dim I
For Each I In A
    Push O, I
Next
End Sub

Function AyIsEmpty(A) As Boolean
AyIsEmpty = Sz(A) = 0
End Function
Function AyIsAllEq(A) As Boolean
If Sz(A) <= 1 Then AyIsAllEq = True: Exit Function
Dim A0, J&
A0 = A(0)
For J = 2 To UB(A)
    If A0 <> A(0) Then Exit Function
Next
AyIsAllEq = True
End Function
Function FfnNxt$(A)
If Not FfnIsExist(A) Then FfnNxt = A: Exit Function
Dim J%, O$
For J = 1 To 99
    O = FfnNxtN(A, J)
    If Not FfnIsExist(O) Then FfnNxt = O: Exit Function
Next
Stop
End Function

Function FfnAddFnSfx$(A, Sfx$)
FfnAddFnSfx = FfnPth(A) & FfnFnn(A) & Sfx & FfnExt(A)
End Function

Function FfnNxtN$(A, N%)
If 1 > N Or N > 99 Then Stop
Dim Sfx$
Sfx = "(" & Format(N, "00") & ")"
FfnNxtN = FfnAddFnSfx(A, Sfx)
End Function

Function PthSel$(A, Optional Tit$ = "Select a Path", Optional BtnNm$ = "Use this path")
With FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
    .InitialFileName = Nz(A, "")
    .Show
    If .SelectedItems.Count = 1 Then
        PthSel = PthEnsSfx(.SelectedItems(1))
    End If
End With
End Function
Sub ZZ_PthSel()
MsgBox FfnSel("C:\")
End Sub
Function FfnSel$(A, Optional FSpec$ = "*", Optional Tit$ = "Select a file", Optional BtnNm$ = "Use the File Name")
With FileDialog(msoFileDialogFilePicker)
    .Filters.Clear
    .Title = Tit
    .AllowMultiSelect = False
    .Filters.Add "", FSpec
    .InitialFileName = A
    .ButtonName = BtnNm
    .Show
    If .SelectedItems.Count = 1 Then
        FfnSel = .SelectedItems(1)
    End If
End With
End Function
Sub TxtbSelPth(A As Access.TextBox)
Dim R$
R = PthSel(A.Value)
If R = "" Then Exit Sub
A.Value = R
End Sub
Function FfnFn$(A)
Dim P%: P = InStrRev(A, "\")
If P = 0 Then FfnFn = A: Exit Function
FfnFn = Mid(A, P + 1)
End Function

Function FfnFnn$(A)
FfnFnn = FfnCutExt(FfnFn(A))
End Function
Function FfnCutExt$(A)
Dim B$, C$, P%
B = FfnFn(A)
P = InStrRev(B, ".")
If P = 0 Then
    C = B
Else
    C = Left(B, P - 1)
End If
FfnCutExt = FfnPth(A) & C
End Function
Function PthEns$(A)
If Dir(A, VbFileAttribute.vbDirectory) = "" Then MkDir A
PthEns = A
End Function
Function AyWhOPred(A, Obj, Pred$)
If Sz(A) = 0 Then AyWhOPred = A: Exit Function
Dim I, O, X
O = AyCln(A)
For Each I In A
    X = CallByName(Obj, Pred, VbMethod, I)
    If X Then
        Push O, I
    End If
Next
AyWhOPred = O
End Function
Function PthFfnAy(A, Optional Spec$ = "*") As String()
PthFfnAy = AyAddPfx(PthFnAy(A, Spec), PthEnsSfx(A))
End Function
Function Lik(A, K) As Boolean
Lik = A Like K
End Function
Function ItrNyWhLik(A, Lik) As String()
ItrNyWhLik = AyWhLik(ItrNy(A), Lik)
End Function
Function PthFdrAy(A, Optional Spec$ = "*", Optional Atr% = 0) As String()
PthFdrAy = ItrNyWhLik(Fso.GetFolder(A).SubFolders, Spec)
End Function
Function PthUp$(A, Optional Up% = 1)
Dim O$, J%
O = A
For J = 1 To Up
    O = PthUpOne(O)
Next
PthUp = O
End Function
Function Cd$(Optional A)
If IsEmp(A) Then
    Cd = PthEnsSfx(CurDir)
    Exit Function
End If
ChDir A
Cd = PthEnsSfx(A)
End Function
Function TmpFdrAy(Optional Spec$ = "*") As String()

End Function
Function CurFdrAy(Optional Spec$ = "*") As String()
CurFdrAy = PthFdrAy(CurDir)
End Function
Function CurFnAy(Optional Spec$ = "*") As String()
CurFnAy = PthFnAy(CurDir, Spec)
End Function
Function PthFnAy(A, Optional Spec$ = "*") As String()
Dim O$(), B$, P$
P = PthEnsSfx(A)
B = Dir(P & Spec)
Dim J%
While B <> ""
    J = J + 1
    If J > 1000 Then Stop
    Push O, B
    B = Dir
Wend
PthFnAy = O
End Function

Function FfnExt$(Ffn)
Dim P%: P = InStrRev(Ffn, ".")
If P = 0 Then Exit Function
FfnExt = Mid(Ffn, P)
End Function

Function PthFxAy(A) As String()
Dim O$(), B$
If Right(A, 1) <> "\" Then Stop
B = Dir(A & "*.xls")
Dim J%
While B <> ""
    J = J + 1
    If J > 1000 Then Stop
    If FfnExt(B) = ".xls" Then
        Push O, A & B
    End If
    B = Dir
Wend
PthFxAy = O
End Function

Function RmvLasChr$(A)
RmvLasChr = Left(A, Len(A) - 1)
End Function
Function RmvFstChr$(A)
RmvFstChr = Mid(A, 2)
End Function

Function AyIsEq(A, B) As Boolean
Dim U&, J&
U = UB(A)
If UB(B) <> U Then Exit Function
For J = 0 To U
    If A(J) <> B(J) Then Exit Function
Next
AyIsEq = True
End Function
Function RsIsBrk(A As DAO.Recordset, GpKy$(), LasVy()) As Boolean
RsIsBrk = Not AyIsEq(RsVy(A, GpKy), LasVy)
End Function
Function RsVy(A As DAO.Recordset, Optional Ky0) As Variant()
RsVy = FldsVy(A.Fields, Ky0)
End Function
Function FldsVyByKy(A As DAO.Fields, Ky$()) As Variant()
Dim O(), J%, K
If Sz(Ky) = 0 Then
    FldsVyByKy = ItrVy(A)
    Exit Function
End If
ReDim O(UB(Ky))
For Each K In Ky
    O(J) = A(K).Value
    J = J + 1
Next
FldsVyByKy = O
End Function
Sub ZZ_FldsVy()
Dim Rs As DAO.Recordset, Vy()
Set Rs = CurrentDb.OpenRecordset("Select * from SkuB")
With Rs
    While Not .EOF
        Vy = RsVy(Rs)
        Debug.Print JnComma(Vy)
        .MoveNext
    Wend
    .Close
End With
End Sub
Function ItrPrpAy(A, PrpNm$) As Variant()
Dim O(), I
For Each I In A
    Push O, CallByName(I, PrpNm, VbGet)
Next
ItrPrpAy = O
End Function
Function ItrVy(A) As Variant()
ItrVy = ItrPrpAy(A, "Value")
End Function
Function IsDte(A) As Boolean
IsDte = VarType(A) = vbDate
End Function
Function IsBool(A) As Boolean
IsBool = VarType(A) = vbBoolean
End Function
Function IsStr(A) As Boolean
IsStr = VarType(A) = vbString
End Function
Function IsEmp(A) As Boolean
IsEmp = True
Select Case True
Case IsStr(A)
    IsEmp = Trim(A) = ""
Case IsArray(A)
    IsEmp = Sz(A) = 0
Case IsEmpty(A), IsNothing(A)
    IsEmp = False
End Select
End Function
Function IsSy(A) As Boolean
IsSy = VarType(A) = vbString + vbArray
End Function
Function CvSy(A) As String()
Select Case True
Case IsSy(A): CvSy = A
Case IsArray(A): CvSy = AySy(A)
Case Else: CvSy = ApSy(CStr(A))
End Select
End Function
Function FldsVy(A As DAO.Fields, Optional Ky0) As Variant()
Select Case True
Case IsMissing(Ky0)
    FldsVy = ItrVy(A)
Case IsStr(Ky0)
    FldsVy = FldsVyByKy(A, SslSy(Ky0))
Case IsSy(Ky0)
    FldsVy = FldsVyByKy(A, CvSy(Ky0))
Case Else
    Stop
End Select
End Function
Private Sub ZZ_SslSqBktCsv()
Debug.Print SslSqBktCsv("a b c")
End Sub
Function SslSqBktCsv$(A)
Dim B$(), C$()
B = SslSy(A)
C = AyQuoteSqBkt(B)
SslSqBktCsv = JnComma(C)
End Function
Function Ny0SqBktCsv$(A)
Dim B$(), C$()
B = CvNy(A)
C = AyQuoteSqBkt(B)
Ny0SqBktCsv = JnComma(C)
End Function
Function RsFny(A As DAO.Recordset) As String()
RsFny = FldsFny(A.Fields)
End Function

Function AyHasAy(A, Ay) As Boolean
Dim I
For Each I In Ay
    If Not AyHas(A, I) Then Exit Function
Next
AyHasAy = True
End Function

Function SqlQQStr_Sy(Sql, QQStr$) As String()
Dim Dry: Dry = SqlDry(Sql)
If AyIsEmpty(Dry) Then Exit Function
Dim O$()
Dim Dr
For Each Dr In Dry
    Push O, FmtQQAv(QQStr, CvAv(Dr))
Next
SqlQQStr_Sy = O
End Function
Function NewTd(T, FdAy() As DAO.Field) As DAO.TableDef
Dim O As New DAO.TableDef
O.Name = T
AyDoPX FdAy, "TdAddFd", O
Set NewTd = O
End Function

Function FldsCsv$(A As DAO.Fields)
FldsCsv = AyCsv(ItrVy(A))
End Function
Function VarCsv$(A)
Select Case True
Case IsStr(A): VarCsv = """" & A & """"
Case IsDte(A): VarCsv = Format(A, "YYYY-MM-DD HH:MM:SS")
Case Else: VarCsv = Nz(A, "")
End Select
End Function
Function AyMapInto(A, MapFunNm$, OInto)
Dim J&, O, I, U&
O = OInto
Erase O
U = UB(A)
If U = -1 Then
    AyMapInto = O
    Exit Function
End If
ReDim O(U)
For Each I In A
    Asg Run(MapFunNm, I), O(J)
    J = J + 1
Next
AyMapInto = O
End Function
Sub Asg(Fm, OTo)
If IsObject(Fm) Then
    Set OTo = Fm
Else
    OTo = Nz(Fm, "")
End If
End Sub
Function ItrMap(A, Map$) As Variant()
ItrMap = ItrMapInto(A, Map, EmpAy)
End Function
Function AyMapSy(A, MapFunNm$) As String()
AyMapSy = AyMapInto(A, MapFunNm, EmpSy)
End Function
Function AyCsv$(A)
AyCsv = Join(A, ",")
Exit Function
Dim J%
For J = 0 To UB(A)
    A(J) = VarCsv(A(J))
Next
AyCsv = Join(A, ",")
End Function
Sub ZZ_DbtUpdSeq()
DoCmd.SetWarnings False
DoCmd.RunSQL "Select * into [#A] from ZZ_DbtUpdSeq order by Sku,PermitDate"
DoCmd.RunSQL "Update [#A] set BchRateSeq=0, Rate=Round(Rate,0)"
DbtUpdSeq CurrentDb, "#A", "BchRateSeq", "Sku", "Sku Rate"
TblOpn "#A"
Stop
DoCmd.RunSQL "Drop Table [#A]"
End Sub

Sub DbtUpdSeq(A As Database, T$, SeqFldNm$, Optional RestFny0, Optional IncFny0)
'Assume T is sorted
'
'Update A->T->SeqFldNm using RestFny0,IncFny0, assume the table has been sorted
'Update A->T->SeqFldNm using OrdFny0, RestFny0,IncFny0
Dim RestFny$(), IncFny$(), Sql$
Dim LasRestVy(), LasIncVy(), Seq&, OrdS$, Rs As DAO.Recordset
'OrdFny RestAy IncAy Sql
RestFny = CvNy(RestFny0)
IncFny = CvNy(IncFny0)
If Sz(RestFny) = 0 And Sz(IncFny) = 0 Then
    With A.OpenRecordset(T)
        Seq = 1
        While Not .EOF
            .Edit
            .Fields(SeqFldNm) = Seq
            Seq = Seq + 1
            .Update
            .MoveNext
        Wend
        .Close
    End With
    Exit Sub
End If
'--
Set Rs = A.OpenRecordset(T) ', RecordOpenOptionsEnum.dbOpenForwardOnly, dbForwardOnly)
With Rs
    While Not .EOF
        If RsIsBrk(Rs, RestFny, LasRestVy) Then
            Seq = 1
            LasRestVy = RsVy(Rs, RestFny)
            LasIncVy = RsVy(Rs, IncFny)
        Else
            If RsIsBrk(Rs, IncFny, LasIncVy) Then
                Seq = Seq + 1
                LasIncVy = RsVy(Rs, IncFny)
            End If
        End If
        .Edit
        .Fields(SeqFldNm).Value = Seq
        .Update
        .MoveNext
    Wend
End With
End Sub

Function CvRg(A) As Range
Set CvRg = A
End Function

Function PnmFn$(A)
PnmFn = PnmVal(A & "Fn")
End Function

Function RsCsv$(A As DAO.Recordset)
RsCsv = FldsCsv(A.Fields)
End Function

Function AyQuoteSqBktCsv$(A)
AyQuoteSqBktCsv = JnComma(AyQuoteSqBkt(A))
End Function

Function LinRmvTerm$(ByVal A$)
LinShiftTerm A
LinRmvTerm = A
End Function
Sub AppExp()
PthClr SrcPth
AppExpMd
AppExpFrm
AppExpStru
End Sub
Sub AppExpFrm()
Dim Nm$, P$, I
P = SrcPth
For Each I In AppFrmNy
    Nm = I
    SaveAsText acForm, Nm, P & Nm & ".Frm.Txt"
Next
End Sub
Function AppFrmNy() As String()
AppFrmNy = ItrNy(CodeProject.AllForms)
End Function
Function AppMdNy() As String()
AppMdNy = ItrNy(CodeProject.AllModules)
End Function
Function Stru$()
Stru = DbStru(CurrentDb)
End Function
Sub AppExpStru()
StrWrt Stru, SrcPth & "Stru.txt"
End Sub
Sub FfnDltIfExist(A)
On Error GoTo X
If FfnIsExist(A) Then Kill A
Exit Sub
X:
Debug.Print "FfnDltIfExist: Unable to delete file [" & A & "].  Er[" & Err.Description & "]"
End Sub
Sub FfnAy_DltIfExist(A)
AyDo A, "FfnDltIfExist"
End Sub

Sub PthClr(A)
FfnAy_DltIfExist PthFfnAy(A)
End Sub
Sub SrcPthBrw()
PthBrw SrcPth
End Sub
Function SrcPth$()
Dim X As Boolean, Y$
If Not X Then
    X = True
    Y = CDbPth & "Src\"
    PthEns Y
End If
SrcPth = Y
End Function
Sub AppExpMd()
Dim MdNm$, I, P$
P = SrcPth
For Each I In AppMdNy
    MdNm = I
    SaveAsText acModule, MdNm, P & MdNm & ".bas"
Next
End Sub
Sub ZZ_DbtReSeq()
DbtReSeq CurrentDb, "ZZ_DbtUpdSeq", "Permit PermitD"
End Sub

Sub DbtReSeq(A As Database, T, ReSeqSpec$)
DbtReSeq_zFny A, T, ReSeqSpec_Fny(ReSeqSpec)
End Sub
Function DiczT1RestLy(T1RestLy$()) As Dictionary
Dim I, L$, K$, O As New Dictionary
If Sz(T1RestLy) > 0 Then
    For Each I In T1RestLy
        L = I
        K = LinT1(L)
        O.Add K, L
    Next
End If
Set DiczT1RestLy = O
End Function

Sub Z_ReSeqSpec_OutLinL1Ay()
D ReSeqSpec_OutLinL1Ay(Z_ReSeqSpec)
End Sub
Function AyT1Ay(A) As String()
AyT1Ay = AyMapSy(A, "LinT1")
End Function
Function AyT2Ay(A) As String()
AyT2Ay = AyMapSy(A, "LinT2")
End Function
Sub ZZ_ReSeqSpec_OutLinL1Ay()
AyBrw ReSeqSpec_OutLinL1Ay(Z_ReSeqSpec)
End Sub

Function ReSeqSpec_OutLinL1Ay(A) As String()
Dim B$()
B = SplitVBar(A)
AyShift B
ReSeqSpec_OutLinL1Ay = AyT1Ay(B)
End Function

Sub ZZ_ReSeqSpec_Fny()
AyBrw ReSeqSpec_Fny(Z_ReSeqSpec)
End Sub

Function ReSeqSpec_Fny(A) As String()
Dim Ay$(), D As Dictionary, O$(), L1$, L
Ay = SplitVBar(A)
If Sz(Ay) = 0 Then Exit Function
L1 = AyShift(Ay)
Set D = DiczT1RestLy(Ay)
For Each L In SslSy(L1)
    If D.Exists(L) Then
        Push O, D(L)
    Else
        Push O, L
    End If
Next
ReSeqSpec_Fny = SslSy(JnSpc(O))
End Function
Sub DbReOpn(A As Database)
Dim Nm$
Nm = A.Name
A.Close
Set A = DAO.DBEngine.OpenDatabase(Nm)
End Sub
Sub DbtReSeq_zFny(A As Database, T, Fny$())
Dim TFny$(), F$(), J%, FF
TFny = DbtFny(A, T)
If Sz(TFny) = Sz(Fny) Then
    F = Fny
Else
    F = AyAdd(Fny, AyMinus(TFny, Fny))
End If
For Each FF In F
    J = J + 1
    A.TableDefs(T).Fields(FF).OrdinalPosition = J
Next
End Sub
Function OyDrs(A, PrpNy0) As Drs
Dim Fny$(), Dry()
Fny = CvNy(PrpNy0)
Dry = OyDry(A, Fny)
Set OyDrs = Drs(Fny, Dry)
End Function
Function ObjDr(A, PrpNy0) As Variant()
Dim PrpNy$(), U%, O(), J%
PrpNy = CvNy(PrpNy0)
U = UB(PrpNy)
ReDim O(U)
For J = 0 To U
    Asg ObjPrp(A, PrpNy(J)), O(J)
Next
ObjDr = O
End Function

Function OyDry(A, PrpNy0) As Variant()
Dim O(), U%, I
Dim PrpNy$()
PrpNy = CvNy(PrpNy0)
For Each I In A
    Push O, ObjDr(I, PrpNy)
Next
OyDry = O
End Function

Sub ZZ_OyDrs()
WsVis DrsWs(OyDrs(CurrentDb.TableDefs("ZZ_DbtUpdSeq").Fields, "Name Type OrdinalPosition"))
End Sub

Function DrsWs(A As Drs) As Worksheet
Set DrsWs = SqWs(DrsSq(A))
End Function

Function DrsPutAt(A As Drs, At As Range) As Range
Set DrsPutAt = SqPutAt(DrsSq(A), At)
End Function

Function DryWs(A) As Worksheet
Set DryWs = SqWs(DrySq(A))
End Function

Function DryNCol%(A)
Dim O%, Dr
For Each Dr In A
    O = Max(O, Sz(Dr))
Next
DryNCol = O
End Function

Function DrySq(A) As Variant()
Dim O(), C%, R&, Dr
Dim NC%, NR&
NC = DryNCol(A)
NR = Sz(A)
ReDim O(1 To NR, 1 To NC)
For R = 1 To NR
    Dr = A(R - 1)
    For C = 1 To Min(Sz(Dr), NC)
        O(R, C) = Dr(C - 1)
    Next
Next
DrySq = O
End Function

Function DbPth$(A As Database)
DbPth = FfnPth(A.Name)
End Function

Function DrsNCol%(A As Drs)
DrsNCol = Max(Sz(A.Fny), DryNCol(A.Dry))
End Function

Sub TpWrtFfn(Ffn$)
AttExp "Tp", Ffn
End Sub

Sub TpExp()
AttExp "Tp", TpFx
End Sub

Sub TpImp()
Dim A$
A = TpFx
If Not FfnIsExist(A) Then
    FunMsgBrw "TpImp", "[Tp] not exist, no TpImp.", A
    Exit Sub
End If
If AttIsOld("Tp", A) Then AttImp "Tp", A
End Sub

Function AttIsOld(A, Ffn) As Boolean
Dim T1 As Date, T2 As Date
T1 = AttTim(A)
T2 = FfnTim(Ffn)
Dim Msg$
Msg = FmtQQ("[Att] is ? in comparing with [file] using [Att-Tim] & [file-Tim].  Is Att [Older]?  (@AttIsOld)", IIf(T1 < T2, "older", "newer"), "?")
FunMsgDmp "AttIsOld", Msg, A, Ffn, T1, T2, T1 < T2
AttIsOld = T1 < T2
End Function

Function DbtPkIxNm$(A As Database, T)
Dim I As DAO.Index
For Each I In A.TableDefs(T).Indexes
    If I.Primary Then DbtPkIxNm = I.Name
Next
End Function

Function AttTim(A) As Date
AttTim = TfkV("Att", "FilTim", A)
End Function

Function AttSz(A) As Date
AttSz = TfkV("Att", "FilSz", A)
End Function
Function DrsSq(A As Drs) As Variant()
Dim O(), C%, R&, Dr(), Dry()
Dim Fny$(), NC%, NR&
Dry = A.Dry
Fny = A.Fny

NR = Sz(Dry)
NC = DrsNCol(A)
If Sz(Fny) <> NC Then Stop
ReDim O(1 To NR + 1, 1 To NC)
For C = 1 To NC
    O(1, C) = Fny(C - 1)
Next
For R = 1 To NR
    Dr = Dry(R - 1)
    For C = 1 To Min(Sz(Dr), NC)
        O(R + 1, C) = Dr(C - 1)
    Next
Next
DrsSq = O
End Function
Function SqWs(A) As Worksheet
Set SqWs = LoWs(SqLo(A, NewA1))
End Function
Sub SqlRun(A)
CurrentDb.Execute A
End Sub
Sub QQ(QQSql, ParamArray Ap())
Dim Av(): Av = Ap
CurrentDb.Execute FmtQQAv(QQSql, Av)
End Sub
Sub WImpTbl(TT)
DbtImpTbl W, TT
End Sub

Function WbMax(A As Workbook) As Workbook
A.Application.WindowState = xlMaximized
Set WbMax = A
End Function

Sub Done()
MsgBox "Done"
End Sub

Function WtChkCol(T$, LnkColStr$) As String()
WtChkCol = DbtChkCol(W, T, LnkColStr)
End Function
'
'Function WQQRs(QQSql, ParamArray Ap()) As Recordset
'Dim Av(): Av = Ap
'Set WQQRs = DbqRs(W, FmtQQAv(QQSql, Av))
'End Function
'Function W.OpenRecordset(Sql)
'Set W.OpenRecordset = DbqRs(W, Sql)
'End Function
'Function WQV(Sql)
'WQV = DbqV(W, Sql)
'End Function
Sub WttLnkFb(TT, Fb$, Optional Fbtt)
DbttLnkFb W, TT, Fb, Fbtt
End Sub

Sub WQuit()
WCls
Quit
End Sub

Sub WtLnkFx(T, Fx, Optional WsNm$ = "Sheet1")
DbtLnkFx W, T, Fx, WsNm
End Sub

Sub WOpn()
Set X_W = FbDb(WFb)
End Sub

Sub WKill()
WCls
FfnDltIfExist WFb
End Sub
Sub WCrt()
FbCrt WFb
End Sub
Property Get WExist() As Boolean
WExist = FfnIsExist(WFb)
End Property
Sub WEns()
If Not WExist Then WCrt
End Sub
Function QV(A)
QV = SqlV(A)
End Function
Function WQQV(A, ParamArray Ap())
Dim Av(): Av = Ap
WQQV = DbqV(W, FmtQQAv(A, Av))
End Function
Sub WtRenCol(T, Fm, NewCol)
DbtRenCol W, T, Fm, NewCol
End Sub
Sub WRun(A)
On Error GoTo X
W.Execute A
Exit Sub
X:
Debug.Print Err.Description
Debug.Print A
Debug.Print "?WStru("""")"
On Error Resume Next
DbCrtQry W, "Query1", A
Stop
End Sub
Function FfnAlreadyLdMsgLy(A, FilKind$, LdTim$) As String()
Dim Sz&, Tim$, Ld$, Msg$
Sz = FfnSz(A)
Tim = FfnDTim(A)
Msg = FmtQQ("[?] file of [time] and [size] is already loaded [at].", FilKind)
FfnAlreadyLdMsgLy = MsgLy(Msg, A, Tim, Sz, LdTim)
End Function
Function QQDTim$(QQSql$, ParamArray Ap())
Dim Av(): Av = Ap
QQDTim = DbqDTim(CurrentDb, FmtQQAv(QQSql, Av))
End Function
Function QQTim(QQSql$, ParamArray Ap()) As Date
Dim Av(): Av = Ap
QQTim = DbqTim(CurrentDb, FmtQQAv(QQSql, Av))
End Function
Function DbqDTim$(A As Database, Sql)
DbqDTim = DteDTim(DbqV(A, Sql))
End Function
Sub ZZZ_AyIns()
Dim A, M, At&, Exp
A = Array(1, 2, 3)
M = "X"
Exp = Array("X", 1, 2, 3)
GoSub Tst
Exit Sub
Tst:
Dim Act
Act = AyIns(A, M, At)
Debug.Assert AyIsEq(Act, Exp)
Return
End Sub
Function AyIns(A, M, Optional At&)
Dim O, N&, J&
N = Sz(A)
O = A
ReDim Preserve O(N)
If IsObject(M) Then
    For J = N To At + 1 Step -1
        Set O(J) = O(J - 1)
    Next
Else
    For J = N To At + 1 Step -1
        O(J) = O(J - 1)
    Next
End If
Asg M, O(At)
AyIns = O
End Function
Function TdHasCnStr(A As DAO.TableDef) As Boolean
TdHasCnStr = A.Connect <> ""
End Function

Property Get LnkTny() As String()
LnkTny = CDbLnkTny
End Property
Function ItrWhPredPrpAyInto(A, Pred$, P, OInto)
Dim O: O = OInto
Erase O
Dim X
For Each X In A
    If Run(Pred, X) Then
        Push O, ObjPrp(X, P)
    End If
Next
ItrWhPredPrpAyInto = O
End Function
Function ItrWhPredPrpAy(A, Pred$, P)
ItrWhPredPrpAy = ItrWhPredPrpAyInto(A, Pred, P, EmpAy)
End Function
Function ItrWhPredPrpSy(A, Pred$, P) As String()
ItrWhPredPrpSy = ItrWhPredPrpAyInto(A, Pred, P, EmpSy)
End Function
Function DbLnkTny(A As Database) As String()
DbLnkTny = ItrWhPredPrpSy(A.TableDefs, "TdHasCnStr", "Name")
End Function
Sub DrpLnkTbl()
CDbDrpLnkTbl
End Sub

Sub DbDrpLnkTbl(A As Database)
DbDrpTT A, DbLnkTny(A)
End Sub
Sub RsUpd(A As DAO.Recordset, ParamArray Ap())
Dim Av(): Av = Ap
RsUpdDr A, Av
End Sub
Function FfnStamp(A) As Variant()
FfnStamp = Array(A, FfnSz(A), FfnTim(A), Now)
End Function

Sub RsUpdDr(A As DAO.Recordset, Dr)
Dim J%, V
With A
    .Edit
    For Each V In Dr
        .Fields(J).Value = V
        J = J + 1
    Next
    .Update
End With
End Sub
Sub AyAsg(A, ParamArray OAp())
Dim Av(): Av = OAp
Dim J%
For J = 0 To Min(UB(Av), UB(A))
    Asg A(J), OAp(J)
Next
End Sub
Function NewAcs(Optional Hid As Boolean) As Access.Application
Dim O As New Access.Application
If Not Hid Then O.Visible = True
Set NewAcs = O
End Function
Sub BrwDtaFb()
Acs.OpenCurrentDatabase DtaFb
End Sub
Property Get DtaFn$()
DtaFn = Apn & "_Data.accdb"
End Property
Property Get DtaFb$()
DtaFb = AppHom & DtaFn
End Property
Function M_PrvM(M As Byte) As Byte
M_PrvM = IIf(M = 1, 12, M - 1)
End Function
Function M_NxtM(M As Byte) As Byte
M_NxtM = IIf(M = 12, 1, M + 1)
End Function
Function YM_YofPrvM(Y As Byte, M As Byte) As Byte
YM_YofPrvM = IIf(M = 1, Y - 1, Y)
End Function
Function YM_YofNxtM(Y As Byte, M As Byte) As Byte
YM_YofNxtM = IIf(M = 12, Y + 1, Y)
End Function

Sub AyAsgT1AyRestAy(A, OT1Ay$(), ORestAy$())
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
Function AyFstT1$(A, T1)
AyFstT1 = AyFstPredXPYes(A, "LinHasT1", T1)
End Function
Function AyFstPredXPYes(A, XP$, P)
If Sz(A) = 0 Then Debug.Print "AyFstPredXP: No element in Ay": Exit Function
Dim X
For Each X In A
    If Run(XP, X, P) Then Asg X, AyFstPredXPYes: Exit Function
Next
Debug.Print FmtQQ("AyFstPredXP: No element in Ay of NEle[?] having Pred[?] P[?] being Yes", Sz(A), XP, P)
End Function
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
    AyAsgT1AyRestAy A, T1, Rest
T1 = AyAlignL(T1)
AyAlignT1 = AyabAddWSpc(T1, Rest)
End Function

Sub MsgDmp(A$, ParamArray Ap())
Dim Av(): Av = Ap
AyDmp MsgAv_Ly(A, Av)
End Sub

Sub FunMsgDmpAv(A, Msg$, Av())
AyDmp FunMsgLy(A, Msg, Av)
End Sub

Sub FunMsgDmp(A, Msg$, ParamArray Ap())
Dim Av(): Av = Ap
FunMsgDmpAv A, Msg, Av
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
DbtPrp(A, T, C_Des) = Des
End Property
Property Let TblPrp(T, P, V)
DbtPrp(CurrentDb, T, P) = V
End Property
Function DbtHasPrp(A As Database, T, P) As Boolean
DbtHasPrp = ItrHasNm(A.TableDefs(T).Properties, P)
End Function

Property Let DbtPrp(A As Database, T, P, V)
If DbtHasPrp(A, T, P) Then
    A.TableDefs(T).Properties(P).Value = V
Else
    A.TableDefs(T).Properties.Append DbtCrtPrp(A, T, P, V)
End If
End Property
Function DbtCrtPrp(A As Database, T, P, V) As DAO.Property
Set DbtCrtPrp = A.TableDefs(T).CreateProperty(P, VarDaoTy(V), V)
End Function
Property Get DbtDes$(A As Database, T)
DbtDes = DbtPrp(A, T, C_Des)
End Property
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

Function AyMap(A, Map$)
Dim O
O = A
Erase O
AyMap = AyMapInto(A, Map, O)
End Function

Sub MsgAp_Brw(Msg$, ParamArray Ap())
Dim Av(): Av = Ap
MsgAv_Brw Msg, Av
End Sub
Sub FunMsgBrw(Fun$, Msg$, ParamArray Ap())
Dim Av(): Av = Ap
AyBrw FunMsgAv_Ly(Fun, Msg, Av)
End Sub

Sub TdAddId(A As DAO.TableDef)
A.Fields.Append NewFd_zId(A.Name)
End Sub

Sub TdAddStamp(A As DAO.TableDef, F)
A.Fields.Append NewFd(F, DAO.dbDate, Dft:="Now")
End Sub

Function CvFF(FF) As String()
CvFF = CvNy(FF)
End Function

Sub TdAddLngFld(A As DAO.TableDef, FF)
Dim F
For Each F In CvFF(FF)
    A.Fields.Append NewFd(F, dbLong)
Next
End Sub
Sub TdAddTxtFld(A As DAO.TableDef, FF, Optional Sz% = 255)
Dim F
For Each F In CvFF(FF)
    A.Fields.Append NewFd(F, dbText, Sz)
Next
End Sub
Sub TdAddLngTxt(A As DAO.TableDef, FF)
Dim F
For Each F In CvFF(FF)
    A.Fields.Append NewFd(F, dbMemo)
Next
End Sub

Sub DbtCrtPk(A As Database, T)
Q = FmtQQ("Create Index PrimaryKey on ? (?) with Primary", T, T): A.Execute Q
End Sub

Sub DbttCrtPk(A As Database, TT)
AyDoPX CvTT(TT), "DbtCrtPk", A
End Sub

Function DbtSk(A As Database, T) As String()
'Sk is Fny of a table with same name as Table name and with Unique
Dim I As DAO.Index:
Set I = DbtSkIdx(A, T): If IsNothing(I) Then Exit Function
DbtSk = ItrNy(I.Fields)
End Function

Function ItrXPPredFst(A, XP$, P)
Dim X
For Each X In A
    If Run(XP, X, P) Then Asg X, ItrXPPredFst: Exit Function
Next
End Function

Function IdxIsSk(A As DAO.Index, T) As Boolean
If A.Name <> T Then Exit Function
IdxIsSk = A.Unique
End Function

Function DbtSkIdx(A As Database, T) As DAO.Index
Dim O
Asg ItrXPPredFst(A.TableDefs(T).Indexes, "IdxIsSk", T), O
If Not IsEmpty(O) Then
    Set DbtSkIdx = O
End If
End Function

Sub DbtCrtSk(A As Database, T, SKey, FF)
Q = FmtQQ("Create Unique Index ? on ? (?)", SKey, T, JnComma(CvFF(FF))): A.Execute Q
End Sub

Function FtLines$(A)
FtLines = Fso.GetFile(A).OpenAsTextStream.ReadAll
End Function
Function AyRmvT1(A) As String()
AyRmvT1 = AyMapSy(A, "LinRmvT1")
End Function
Function FfnPing(A) As Boolean
If Not FfnIsExist(A) Then Debug.Print "[" & A & "] not found": FfnPing = True
End Function

Sub DbRun(A As Database, Sql)
A.Execute Sql
End Sub
Sub DaoShtTySz_BrkAsg(A, OTy As DAO.DatabaseTypeEnum, OSz%)
OSz = Val(Mid(A, 4))
OTy = DaoShtTy_Ty(Left(A, 3))
End Sub

Sub DbAppTd(A As Database, Td As DAO.TableDef)
A.TableDefs.Append Td
End Sub

Function RsLin$(A As DAO.Recordset, Optional Sep$ = " ")
RsLin = Join(RsDr(A), Sep)
End Function

Sub DbCrtSchm(A As Database, SchmLy$())
Dim M As New Schm
M.DbCrtSchm A, SchmLy
End Sub

Function AppDtaPth$()
AppDtaPth = PthEns(AppDtaHom & Apn & "\")
End Function

Function AppDtaHom$()
AppDtaHom = PthUp(TmpHom)
End Function

Function AscIsFstNmChr(A%) As Boolean
AscIsFstNmChr = AscIsLetter(A)
End Function

Function AscIsNmChr(A%) As Boolean
AscIsNmChr = True
If AscIsLetter(A) Then Exit Function
If AscIsDig(A) Then Exit Function
AscIsNmChr = A = 95 '_
End Function
Sub Z_RmvNm()
Dim Nm$
Nm = "lksdjfsd f"
Expect = " f"
GoSub Tst
Exit Sub
Tst:
    Actual = RmvNm(Nm)
    C
    Return
End Sub
Function IsTyChr(A) As Boolean
If Len(A) <> 1 Then Stop
IsTyChr = InStr("!@#$%^&", A) > 0
End Function

Function RmvTyChr$(A)
If IsTyChr(FstChr(A)) Then RmvTyChr = RmvFstChr(A) Else RmvTyChr = A
End Function
Function LinNm$(A)
Dim O%
If Not AscIsFstNmChr(Asc(FstChr(A))) Then Exit Function
For O = 1 To Len(A)
    If Not AscIsNmChr(Asc(Mid(A, O, 1))) Then GoTo X
Next
X:
    LinNm = Left(A, O - 1)
End Function
Function RmvNm$(A)
Dim O%
If Not AscIsFstNmChr(Asc(FstChr(A))) Then GoTo X
For O = 1 To Len(A)
    If Not AscIsNmChr(Asc(Mid(A, O, 1))) Then GoTo X
Next
X:
    If O > 0 Then RmvNm = Mid(A, O): Exit Function
    RmvNm = A
End Function
Function AySng(A, Optional Msg$ = "AySng")
If Sz(A) <> 1 Then Debug.Print "AySng.Er: [" & Msg & "]": Exit Function
Asg A(0), AySng
End Function

Function AyT1Chd(A, T1) As String()
AyT1Chd = AyRmvT1(AyWhT1EqV(A, T1))
End Function

Property Get CurCdPne() As VBIDE.CodePane
Set CurCdPne = CurVbe.ActiveCodePane
End Property

Property Get CurMd() As VBIDE.CodeModule
Set CurMd = CurCdPne.CodeModule
End Property

Property Get CurVbe() As VBE
Set CurVbe = Application.VBE
End Property
Sub Z_T1LikSslAy_T1()
Dim A$(), Nm$
A = SplitVBar("a bb* *dd | c x y")
Nm = "x"
Expect = "c"
GoSub Tst
Exit Sub
Tst:
    Actual = T1LikSslAy_T1(A, Nm)
    C
    Return
End Sub
Function T1LikSslAy_T1$(T1LikSslAy$(), Nm)
Dim L, T1$
If Sz(T1LikSslAy) = 0 Then Exit Function
For Each L In T1LikSslAy
    T1 = LinShiftT1(L)
    If StrInLikSsl(Nm, L) Then
        T1LikSslAy_T1 = T1
        Exit Function
    End If
Next
End Function
Function ItrFstPrpEqV(A, P, V)
Dim X
For Each X In A
    If ObjPrp(X, P) = V Then Asg X, ItrFstPrpEqV: Exit Function
Next
End Function
Function OyHas(A, Obj) As Boolean
Dim X, OP&
OP = ObjPtr(Obj)
For Each X In A
    If ObjPtr(X) = OP Then OyHas = True: Exit Function
Next
End Function
Sub WinClsAllExcept(A() As VBIDE.Window)
Dim I, W As VBIDE.Window, V As VBIDE.Window
For Each W In CurVbe.Windows
    If Not OyHas(A, W) Then
        W.Visible = False
    End If
Next
For Each I In A
    Set W = I
    If Not W.Visible Then W.Visible = True
Next
End Sub
Property Get WinLcl() As VBIDE.Window
Set WinLcl = ItrFstPrpEqV(CurVbe.Windows, "Type", vbext_wt_Locals)
End Property
Property Get WinImm() As VBIDE.Window
Set WinImm = ItrFstPrpEqV(CurVbe.Windows, "Type", vbext_wt_Immediate)
End Property
Property Get CurWin() As VBIDE.Window
Set CurWin = CurCdPne.Window
End Property
Function ApInto(O, ParamArray Ap())
Dim Av(): Av = Ap
Dim X
Erase O
For Each X In Av
    Push O, X
Next
ApInto = O
End Function
Sub WinSetDbg()
Dim A() As VBIDE.Window
Dim W() As VBIDE.Window
W = ApInto(A, WinLcl, CurWin, WinImm)
WinClsAllExcept W
WinAlignV
End Sub
Property Get OCBmain() As Office.CommandBar
Set OCBmain = CurVbe.CommandBars("Menu Bar")
End Property
Property Get OCCwinVert() As Office.CommandBarButton
Set OCCwinVert = ItrFstPrpEqV(OCPwin.Controls, "Caption", "Tile &Vertically")
End Property
Property Get OCCedtSelAll() As Office.CommandBarButton
Set OCCedtSelAll = ItrFstPrpEqV(OCPedt.Controls, "Caption", "Select &All")
End Property
Property Get OCCedtClr() As Office.CommandBarButton
Set OCCedtClr = ItrFstPrpEqV(OCPedt.Controls, "Caption", "C&lear")
End Property
Property Get OCPwin() As Office.CommandBarPopup
Set OCPwin = ItrFstPrpEqV(OCBmain.Controls, "Caption", "&Window")
End Property
Property Get OCPedt() As Office.CommandBarPopup
Set OCPedt = ItrFstPrpEqV(OCBmain.Controls, "Caption", "&Edit")
End Property
Sub WinImmClr()
WinImm.SetFocus
WinClr
End Sub
Sub WinClr()
OCCedtSelAll.Execute
'OCCedtClr.Execute
End Sub
Sub WinAlignV()
OCCwinVert.Execute
End Sub
Function FdLin$(A As DAO.Field)

End Function

Sub AA()
Dim A As New Schm
A.Z
'D A.Tny
D A.Ly
Stop
End Sub