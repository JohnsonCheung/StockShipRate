Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Const C_E$ = "E"
Const C_TF$ = "TF"
Const C_EF$ = "EF"
Const C_D$ = "D"
Private X_Schmy$()
Public T, F, L
Property Get E$()
On Error GoTo X
E = LinT1(EFLin)
Exit Property
X: Debug.Print "Schm.E: PrpEr.."
End Property

Property Get EFLin$()
On Error GoTo X
EFLin = T1LikSslAy_T1(EFLy, F)
Exit Property
X: Debug.Print "Schm.EFLin: PrpEr.."
End Property

Property Get Ly()
On Error GoTo X
Ly = X_Schmy
Exit Property
X: Debug.Print "Schm.Ly: PrpEr.."
End Property

Sub SetLy(Ly$())
X_Schmy = Ly
End Sub

Function TFELy() As String()
Dim O$()
For Each T In Tny
    For Each F In Fny
        Push O, ApLin(T, F, E)
    Next
Next
TFELy = O
End Function

Function QTFEF1Ly() As String()
Dim O$()
For Each T In Tny
    For Each F In Fny
        Push O, ApLin(T, F, E, EleSpec)
    Next
Next
QTFEF1Ly = O
End Function

Private Function ItmLy(A) As String()
ItmLy = AyT1Chd(Ly, A)
End Function
Function ErLy() As String()
If Sz(X_Schmy) = 0 Then
    ErLy = Sy("no Ly is given")
    Exit Function
End If
If Sz(ErLy) > 0 Then Exit Function
ErLy = AyWhPredXPNot(Ly, "LinInT1Ay", Sy(C_E, C_D, C_EF, C_TF))
End Function

Function ErNoTFld() As String()
If Sz(TFLy) = 0 Then ErNoTFld = Sy("No TFld lines")
End Function

Property Get ErDupT() As String()
On Error GoTo X
ErDupT = AyDupChk(Tny, "These T[?] is duplicated in TFld-lines")
Exit Property
X: Debug.Print "Schm.ErDupT: PrpEr.."
End Property
Property Get EAy() As String()
On Error GoTo X
EAy = AyT1Ay(ELy)
Exit Property
X: Debug.Print "Schm.EAy: PrpEr.."
End Property
Private Sub Z_ErDupE()
Dim Ly$()
Ly = Sy("Ele AA", "Ele BB", "Ele AA")
Expect = Sy("These Ele[AA] are duplicated in Ele-lines")
GoSub Tst
Exit Sub
Tst:
    SetLy Ly
    Actual = ErDupE
    C
    Return
End Sub
Function ErDupE() As String()
ErDupE = AyDupChk(EAy, "These Ele[?] are duplicated in Ele-lines")
End Function
Private Sub Z_ErDupF()
Dim Ly$()
Ly = Sy("TFld AA BB BB")
Expect = Sy("These F[BB] are duplicated in T[AA]")
GoSub Tst
Exit Sub
Tst:
    SetLy Ly
    Actual = ErDupF
    C
    Return
End Sub
Private Sub Z_ErDupT()
Dim Ly$()
Ly = Sy("TFld AA BB BB", "TFld AA DD")
Expect = Sy("These T[AA] is duplicated in TFld-lines")
GoSub Tst
Exit Sub
Tst:
    SetLy Ly
    Actual = ErDupT
    C
    Return
End Sub

Function ErDupF() As String()
For Each T In AyNz(Tny)
    Push ErDupF, AyDupChk(Fny, FmtQQ("These F[?] are duplicated in T[?]", "?", T))
Next
End Function

Function ErEle() As String()
ErEle = AyDupChk(EAy, "These Ele[?] are duplicated in Ele-lines")
End Function

Function ErFldHasNoEle() As String()
For Each T In AyNz(Tny)
    For Each F In AyNz(Fny)
       PushNonEmp ErFldHasNoEle, StrEmpChkMsg(E, FmtQQ("T[?] F[?] has no TEle", T, F))
    Next
Next
End Function

Property Get Er() As String()
On Error GoTo X
Er = AyAddAp(ErLy, ErNoTFld, ErDupT, ErDupF, ErDupE, ErEle, ErFldHasNoEle)
Exit Property
X: Debug.Print "Schm.Er_Ly: PrpEr.."
End Property

Property Get EFLy() As String():  EFLy = ItmLy(C_EF): End Property
Property Get TFLy() As String():  TFLy = ItmLy(C_TF): End Property
Property Get ELy() As String():   ELy = ItmLy(C_E):   End Property
Property Get DLy() As String():   DLy = ItmLy(C_D): End Property
Property Get PkTny() As String(): PkTny = AyT1Ay(PkTFLy):  End Property

Sub Z()
Z_ErDupT
Z_ErDupF
Z_ErDupE
Z_Tny
Exit Sub
Z_DbCrtSchm
End Sub
Sub Z_Ini()
X_Schmy = Z_Schmy
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

Property Get ELin$()
On Error GoTo X
ELin = AyFstT1(ELy, E)
Exit Property
X: Debug.Print "Schm.ELin: PrpEr.."
End Property

Property Get EleSpec$()
On Error GoTo X
Select Case True
Case IsId: EleSpec = "*Id"
Case IsFk: EleSpec = "*Fk"
Case Else: EleSpec = LinRmvT1(ELin)
End Select
Exit Property
X: Debug.Print "Schm.EleSpec: PrpEr.."
End Property

Property Get FdScl$()
On Error GoTo X
FdScl = ApScl(T, F, EleSpec)
Exit Property
X: Debug.Print "Schm.FdScl: PrpEr.."
End Property

Property Get No_F() As Boolean
On Error GoTo X
No_F = F = ""
Exit Property
X: Debug.Print "Schm.No_F: PrpEr.."
End Property
Property Get No_T() As Boolean
On Error GoTo X
No_T = T = ""
Exit Property
X: Debug.Print "Schm.No_T: PrpEr.."
End Property

Property Get Fd() As DAO.Field
On Error GoTo X
If No_F Then Exit Property
Select Case True
Case IsId: Set Fd = NewFd_zId(F)
Case IsFk: Set Fd = NewFd_zFk(F)
Case Else: Set Fd = NewFd_zFdScl(FdScl)
End Select
Exit Property
X: Debug.Print "Schm.Fd: PrpEr.."
End Property

Function Td() As DAO.TableDef
If No_T Then Exit Function
Set Td = NewTd(T, FdAy)
End Function

Property Get Tny() As String()
On Error GoTo X
Tny = AyMapSy(TFLy, "LinT1")
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
On Error GoTo X
'On Error GoTo X
Dim A$(), O$()
A = TFLy
If Sz(A) = 0 Then Exit Property
For Each L In A
    PushNonEmp O, SkSsl
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
PkTFLy = AyWhPred(TFLy, "TFLinHasPk")
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
If AyBrwEr(Er) Then Exit Sub
AyDoPX TdAy, "DbAppTd", A
AyDoPX PkSqy, "DbRun", A
AyDoPX SkSqy, "DbRun", A
End Sub

Property Get TFLin$()
On Error GoTo X
If No_T Then Exit Property
TFLin = AySng(AyWhT1EqV(TFLy, T), "Schm.TFLin.PrpEr")
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
Sub Z_SchmScly()
Z_Ini
D SchmScly
End Sub
Function SchmScly() As String()
Dim O$()
For Each T In AyNz(Tny)
    PushAy O, TdScly
Next
SchmScly = O
End Function
Function TdScly() As String()
TdScly = AyIns(FdScly, TdScl)
End Function
Function TdScl$()
TdScl = ApScl(T, TDes)
End Function
Function TDes$()
TDes = AddLbl("Des", "Des")
End Function
Function FDes$()
FDes = AddLbl("Des", "Des")
End Function
Function FdScly() As String()
Dim O$()
For Each F In Fny
    Push O, FdScl
Next
FdScly = O
End Function
Sub Z_ErLy()
Dim Ly$()
Expect = Sy("No Ly is given")
GoSub Tst
Exit Sub
Tst:
    SetLy Ly
    Actual = ErLy
    C
    Return
End Sub