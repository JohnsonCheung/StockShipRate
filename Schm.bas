Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Const C_FEle$ = "FEle"
Const C_Ele$ = "Ele"
Const C_TFld$ = "TFld"
Const C_TDes$ = "TDes"
Const C_FDes$ = "FDes"
Private X_Ly$()
Public T, F, L
Property Get E$()
Select Case True
Case IsId: E = "*Id"
Case IsFk: E = "*Fk"
Case Else: E = LinT1(LinFEle)
End Select
End Property

Property Get LinFEle$()
Dim A$()
A = LyFEle
If Sz(A) = 0 Then Exit Function
For Each L In A
    If StrInLikSsl(F, LinRmvT1(L)) Then
        LinFEle = L
        Exit Property
    End If
Next
End Property

Property Get Ly()
Ly = X_Ly
End Property

Property Let Ly(V)
X_Ly = V
End Property

Public Property Get Z_Ly() As String()
Dim O$()
Push O, "dfd"
Push O, "Ele Mem   Mem"
Push O, "Ele Amt   Cur;Dft=0"
Push O, "Ele Txt   Txt;Req;AlwZLen;Dft=Johnson;VRul=VRul;VTxt=VTxt"
Push O, "Ele Nm    T20;Req;NonEmp"
Push O, "Ele Crt   Dte;Req;Dft=Now;"
Push O, "Ele Dte   Dte"
Push O, "Ele Des   Txt"
Push O, "Ele Sc    Dbl"
Push O, "FEle Amt *Amt"
Push O, "FEle Crt CrtDte"
Push O, "FEle Dte *Dte"
Push O, "FEle Txt Fun *Txt"
Push O, "FEle Mem Lines"
Push O, "TFld Sess * CrtDte"
Push O, "TFld Msg  * Fun *Txt | CrtDte"
Push O, "TFld Lg   * Sess Msg CrtDte"
Push O, "TFld LgV  * Lg Lines"
Push O, "FDes Fun Function name that call the log"
Push O, "FDes Fun Function name that call the log"
Push O, "TDes Msg it will a new record when Lg-function is first time using the Fun+MsgTxt"
Push O, "TDes Msg ..."
Z_Ly = O
End Property

Property Get TFELy() As String()
Dim O$()
For Each T In Tny
    For Each F In Fny
        Push O, ApLin(T, F, E)
    Next
Next
TFELy = O
End Property

Property Get TFEFdLy() As String()
Dim O$()
For Each T In Tny
    Debug.Print T
    For Each F In Fny
        Push O, ApLin(T, F, E, FdStr)
    Next
Next
TFEFdLy = O
End Property

Function ItmLy(A) As String()
ItmLy = AyRmvT1(AyWhT1EqV(Ly, A))
End Function

Property Get Ly_Er() As String()
Ly_Er = AyWhPredXPNot(Ly, "LinInT1Ay", Sy(C_Ele, C_FDes, C_TDes, C_TFld))
End Property

Property Get EleLy() As String():  EleLy = ItmLy(C_Ele):    End Property
Property Get LyFEle() As String(): LyFEle = ItmLy(C_FEle):  End Property
Property Get LyTFld() As String(): LyTFld = ItmLy(C_TFld):  End Property
Property Get LyFDes() As String(): LyFDes = ItmLy(C_FDes):  End Property
Property Get LyTDes() As String(): LyTDes = ItmLy(C_TDes):  End Property
Property Get PkTny() As String(): PkTny = AyT1Ay(PkTFLy): End Property

Sub Z()
Z_Ini
Z_Tny
Z_DbCrtSchm
End Sub
Sub Z_Ini()
If Sz(X_Ly) = 0 Then X_Ly = Z_Ly
End Sub

Sub Z_Tny()
Z_Ini
Expect = SslSy("Sess Msg Lg LgV")
Actual = Tny
C
End Sub

Sub ZZ_Tny()
Z_Ini
GoSub Sep
D "Tny"
D "---"
D Tny
GoSub Sep
For Each T In Tny
    GoSub Prt
Next
D SkSqy
D PkSqy
Exit Sub
Prt:
    D T
    D UnderLin(T)
    D Fny
    GoSub Sep
    Return
Sep:
    D "--------------------"
    Return
End Sub

Property Get EleLin$()
EleLin = AyFstT1(EleLy, E)
End Property

Property Get EleSpecStr$()
EleSpecStr = LinRmvT1(EleLin)
End Property

Property Get FdDr() As Variant()
With FdSpec
FdDr = Array(.F, .Ty, .Sz, .Req, .AlwZLen, .Dft, .VRul, .VTxt)
End With
End Property

Property Get FdStr$()
FdStr = FdM.FdStr(Fd)
End Property

Property Get Fd() As dao.Field
Select Case True
Case IsId: Set Fd = NewFd_zId(F)
Case IsFk: Set Fd = NewFd_zFk(F)
Case Else: Set Fd = NewFd_zSpec(FdSpec)
End Select
End Property

Property Get FdSpec() As FdSpec
FdSpec = EleSpecStr_FdSpec(EleSpecStr, F)
End Property

Property Get Td() As dao.TableDef
Set Td = NewTd(T, FdAy)
End Property

Property Get Tny() As String()
Tny = AyMapSy(LyTFld, "LinT1")
End Property

Property Get TdAy() As dao.TableDef()
Dim O() As dao.TableDef
For Each T In Tny
    PushObj O, Td
Next
TdAy = O
End Property

Property Get PkSqy() As String()
PkSqy = AyMapSy(PkTny, "TnPkSql")
End Property

Property Get SkSslAy() As String()
Dim A$(), O$()
A = LyTFld
If Sz(A) = 0 Then Exit Property
For Each L In A
    PushNonEmpty O, SkSsl
Next
SkSslAy = O
End Property

Property Get SkSsl$()
Dim A$, B$
A = SkP1: If A = "" Then Exit Property
B = Replace(A, " * ", "")
SkSsl = Replace(B, "*", LinT1(B))
End Property

Property Get SkP1$()
SkP1 = Trim(TakBef(L, "|"))
End Property
Property Get PkTFLy() As String()
PkTFLy = AyWhPred(LyTFld, "TFLinHasPk")
End Property

Property Get SkSqy() As String()
Dim O$(), A$(), B$(), J%, U%
A = SkSslAy
U = UB(A)
If UB(A) = -1 Then Exit Function
ReDim O(U)
For J = 0 To U
    T = LinShiftT1(A(J))
    O(J) = TnSkSql(T, A(J))
Next
SkSqy = O
End Property

Sub Z_DbCrtSchm()
Dim Fb$
Fb = TmpFb
FbCrt Fb
DbCrtSchm FbDb(Fb), Z_Ly
FbBrw Fb
End Sub

Sub DbCrtSchm(A As Database, SchmLy$())
Ly = SchmLy
AyDoPX TdAy, "DbAppTd", A
AyDoPX PkSqy, "DbRun", A
AyDoPX SkSqy, "DbRun", A
End Sub

Property Get TFLin$()
TFLin = AySng(AyWhT1EqV(LyTFld, T), "Schm.TFLin.PrpEr")
End Property

Property Get Fny() As String()
Dim A$, B$
A = TFLin
If LinShiftT1(A) <> T Then Debug.Print "Schm.Fny PrpEr": Exit Property
B = Replace(A, "*", T)
Fny = AyRmvEle(SslSy(B), "|")
End Property

Property Get FdAy() As dao.Field()
Dim O() As dao.Field
For Each F In Fny
    PushObj O, Fd
Next
FdAy = O
End Property

Property Get IsFk() As Boolean
IsFk = AyHas(Tny, F)
End Property

Property Get IsId() As Boolean
IsId = T = F
End Property