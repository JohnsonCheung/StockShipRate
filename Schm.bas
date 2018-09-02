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
Property Get Ele$()
On Error GoTo X
Select Case True
Case IsId: Ele = "*Id"
Case IsFk: Ele = "*Fk"
Case Else: Ele = LinT1(LinFEle)
End Select
Exit Property
X: Debug.Print "Schm.E: PrpEr.."
End Property

Property Get LinFEle$()
On Error GoTo X
LinFEle = T1LikSslAy_T1(LyFEle, F)
Exit Property
X: Debug.Print "Schm.LinFEle: PrpEr.."
End Property

Property Get Ly()
On Error GoTo X
Ly = X_Ly
Exit Property
X: Debug.Print "Schm.Ly: PrpEr.."
End Property

Sub SetLy(Ly$())
X_Ly = Ly
End Sub

Property Get Z_Ly() As String()
On Error GoTo X
Dim O$()
Push O, "dfd"
Push O, "Ele Mem   Mem"
Push O, "Ele Txt   Txt"
Push O, "Ele Crt   Dte;Req;Dft=Now;"
Push O, "FEle Amt *Amt"
Push O, "FEle Crt CrtTim"
Push O, "FEle Dte *Dte"
Push O, "FEle Txt Fun *Txt"
Push O, "FEle Mem Lines"
Push O, "TFld Sess * CrtTim"
Push O, "TFld Msg  * Fun *Txt | CrtTim"
Push O, "TFld Lg   * Sess Msg CrtTim"
Push O, "TFld LgV  * Lg Lines"
Push O, "FDes Fun Function name that call the log"
Push O, "FDes Fun Function name that call the log"
Push O, "TDes Msg it will a new record when Lg-function is first time using the Fun+MsgTxt"
Push O, "TDes Msg ..."
Z_Ly = O
Exit Property
X: Debug.Print "Schm.Z_Ly: PrpEr.."
End Property

Function TFELy() As String()
Dim O$()
For Each T In Tny
    For Each F In Fny
        Push O, ApLin(T, F, Ele)
    Next
Next
TFELy = O
End Function

Function QTFEF1Ly() As String()
Dim O$()
For Each T In Tny
    For Each F In Fny
        Push O, ApLin(T, F, Ele, EleSpec)
    Next
Next
QTFEF1Ly = O
End Function

Function ItmLy(A) As String()
ItmLy = AyT1Chd(Ly, A)
End Function
Property Get ErLy() As String()
ErLy = AyWhPredXPNot(Ly, "LinInT1Ay", Sy(C_Ele, C_FDes, C_FEle, C_TDes, C_TFld))
End Property
Property Get ErNoTFld() As String()
If Sz(LyTFld) = 0 Then ErNoTFld = Sy("No TFld lines")
End Property

Property Get ErDupT() As String()
Dim Dup$()
Dup = AyWhDup(Tny)
ErDupT = Sy(FmtQQ("These [?] is duplicated in TFld-lines", JnSpc(Dup)))
End Property
Property Get EleAy() As String()
EleAy = AyT1Ay(EleLy)
End Property
Function ErDupE() As String()
Dim Dup$()
Dup = AyWhDup(EleAy)
If Sz(Dup) > 0 Then Push ErDupE, FmtQQ("These E[?] are duplicated in Ele-lines", JnSpc(Dup))
End Function
Function ErDupF() As String()
Dim Tny1$(): Tny1 = Tny
If Sz(Tny1) = 0 Then Exit Function
Dim Dup$()
For Each T In Tny1
    Dup = AyWhDup(Fny)
    If Sz(Dup) > 0 Then Push ErDupF, FmtQQ("These F[?] are duplicated in T[?]", JnSpc(Dup), T)
Next
End Function
Function ErEleLin() As String()
Dim Dup$()
Dup = AyWhDup(Tny)
If Sz(Dup) > 0 Then Push ErDupE, FmtQQ("These T[?] are duplicated in TFld-lines", JnSpc(Dup))
End Function

Function ErFldHasNoEle() As String()
Dim Tny1$(), Fny1$()
For Each T In Tny1
    Fny1 = Fny
    For Each F In Fny1
        E = Ele
        If E = "" Then Push Ly_Er, FmtQQ("T[?] F[?] has no TEle", T, F)
    Next
Next
End Function

Property Get Ly_Er() As String()
On Error GoTo X
Ly_Er = AyAddAp(ErLy, ErNoTFld, ErDupT, ErDupF, ErDupE, ErEle, ErFldHasNoEle)
Exit Property
X: Debug.Print "Schm.Ly_Er: PrpEr.."
End Property

Property Get EleLy() As String():  EleLy = ItmLy(C_Ele):   End Property
Property Get LyFEle() As String(): LyFEle = ItmLy(C_FEle): End Property
Property Get LyTFld() As String(): LyTFld = ItmLy(C_TFld): End Property
Property Get LyFDes() As String(): LyFDes = ItmLy(C_FDes): End Property
Property Get LyTDes() As String(): LyTDes = ItmLy(C_TDes): End Property
Property Get PkTny() As String(): PkTny = AyT1Ay(PkTFLy):  End Property

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
On Error GoTo X
Dim A$
A = Ele
Select Case A
Case "*Id", "*Fk": EleLin = A
Case Else:         EleLin = AyFstT1(EleLy, A)
End Select
Exit Property
X: Debug.Print "Schm.EleLin: PrpEr.."
End Property

Property Get EleSpec$()
On Error GoTo X
Dim A$, B$
A = EleLin
B = LinShiftT1(A)
Select Case B
Case "*Id": EleSpec = "Lng;Req"
Case "*Fk": EleSpec = "Lng"
Case Else:  EleSpec = A
End Select
Exit Property
X: Debug.Print "Schm.EleSpec: PrpEr.."
End Property

Property Get FdScl$()
On Error GoTo X
FdScl = F & ";" & EleSpec
Exit Property
X: Debug.Print "Schm.FdScl: PrpEr.."
End Property

Property Get No_F() As Boolean
No_F = F = ""
End Property
Property Get No_T() As Boolean
No_T = T = ""
End Property

Property Get Fd() As DAO.Field
If No_F Then Exit Property
Select Case True
Case IsId: Set Fd = NewFd_zId(F)
Case IsFk: Set Fd = NewFd_zFk(F)
Case Else: Set Fd = NewFd_zFdScl(FdScl)
End Select
Exit Property
X: Debug.Print "Schm.Fd1: PrpEr.."
End Property

Function Td() As DAO.TableDef
If No_T Then Exit Function
Set Td = NewTd(T, FdAy)
End Function

Property Get Tny() As String()
On Error GoTo X
Tny = AyMapSy(LyTFld, "LinT1")
Exit Property
X: Debug.Print "Schm.Tny: PrpEr.."
End Property

Function TdAy() As DAO.TableDef()
Dim O() As DAO.TableDef
For Each T In Tny
    PushObj O, Td
Next
TdAy = O
End Function

Property Get PkSqy() As String()
On Error GoTo X
PkSqy = AyMapSy(PkTny, "TnPkSql")
Exit Property
X: Debug.Print "Schm.PkSqy: PrpEr.."
End Property

Property Get SkSslAy() As String()
'On Error GoTo X
Dim A$(), O$()
A = LyTFld
If Sz(A) = 0 Then Exit Property
For Each L In A
    PushNonEmpty O, SkSsl
Next
SkSslAy = O
Exit Property
X: Debug.Print "Schm.SkSslAy: PrpEr.."
End Property

Property Get SkSsl$()
On Error GoTo X
Dim A$, B$
A = SkP1: If A = "" Then Exit Property
B = Replace(A, " * ", "")
SkSsl = Replace(B, "*", LinT1(B))
Exit Property
X: Debug.Print "Schm.SkSsl: PrpEr.."
End Property

Property Get SkP1$()
On Error GoTo X
SkP1 = Trim(TakBef(L, "|"))
Exit Property
X: Debug.Print "Schm.SkP1: PrpEr.."
End Property

Property Get PkTFLy() As String()
On Error GoTo X
PkTFLy = AyWhPred(LyTFld, "TFLinHasPk")
Exit Property
X: Debug.Print "Schm.PkTFLy: PrpEr.."
End Property

Function SkSqy() As String()
On Error GoTo X
Dim O$(), A$(), B$(), J%, U%, T$
A = SkSslAy
U = UB(A)
If UB(A) = -1 Then Exit Function
ReDim O(U)
For J = 0 To U
    T = LinShiftT1(A(J))
    O(J) = TnSkSql(T, A(J))
Next
SkSqy = O
Exit Function
X: Debug.Print "Schm.SkSqy: PrpEr.."
End Function

Sub Z_DbCrtSchm()
Dim Fb$
Fb = TmpFb
FbCrt Fb
DbCrtSchm FbDb(Fb)
Kill Fb
End Sub

Sub DbCrtSchm(A As Database)
If AyBrwEr(Ly_Er) Then Exit Sub
AyDoPX TdAy, "DbAppTd", A
AyDoPX PkSqy, "DbRun", A
AyDoPX SkSqy, "DbRun", A
End Sub

Property Get TFLin$()
On Error GoTo X
If No_T Then Exit Property
TFLin = AySng(AyWhT1EqV(LyTFld, T), "Schm.TFLin.PrpEr")
Exit Property
X: Debug.Print "Schm.TFLin: PrpEr.."
End Property

Property Get Fny() As String()
On Error GoTo X
Dim A$, B$
A = TFLin
If LinShiftT1(A) <> T Then Debug.Print "Schm.Fny PrpEr": Exit Property
B = Replace(A, "*", T)
Fny = AyRmvEle(SslSy(B), "|")
Exit Property
X: Debug.Print "Schm.Fny: PrpEr.."
End Property

Function FdAy() As DAO.Field()
Dim O() As DAO.Field
For Each F In Fny
    PushObj O, Fd
Next
FdAy = O
End Function

Property Get IsFk() As Boolean
On Error GoTo X
IsFk = AyHas(Tny, F)
Exit Property
X: Debug.Print "Schm.IsFk: PrpEr.."
End Property

Property Get IsId() As Boolean
On Error GoTo X
IsId = T = F
Exit Property
X: Debug.Print "Schm.IsId: PrpEr.."
End Property

Sub A()
Z_Ini
T = "Sess"
F = "Sess"
WinSetDbg
Stop
T = "LgV"
Stop
End Sub