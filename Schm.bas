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
On Error GoTo X
Select Case True
Case IsId: E = "*Id"
Case IsFk: E = "*Fk"
Case Else: E = LinT1(LinFEle)
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

Property Get TFELy() As String()
On Error GoTo X
Dim O$()
For Each T In Tny
    For Each F In Fny
        Push O, ApLin(T, F, E)
    Next
Next
TFELy = O
Exit Property
X: Debug.Print "Schm.TFELy: PrpEr.."
End Property

Property Get TFEFdLy() As String()
On Error GoTo X
Dim O$()
For Each T In Tny
    For Each F In Fny
        Push O, ApLin(T, F, E, FdStr)
    Next
Next
TFEFdLy = O
Exit Property
X: Debug.Print "Schm.TFEFdLy: PrpEr.."
End Property

Function ItmLy(A) As String()
ItmLy = AyT1Chd(Ly, A)
End Function

Property Get Ly_Er() As String()
On Error GoTo X
Ly_Er = AyWhPredXPNot(Ly, "LinInT1Ay", Sy(C_Ele, C_FDes, C_FEle, C_TDes, C_TFld))
Exit Property
X: Debug.Print "Schm.Ly_Er: PrpEr.."
End Property

Property Get EleLy() As String():  EleLy = ItmLy(C_Ele):    End Property
Property Get LyFEle() As String(): LyFEle = ItmLy(C_FEle):  End Property
Property Get LyTFld() As String(): LyTFld = ItmLy(C_TFld):  End Property
Property Get LyFDes() As String(): LyFDes = ItmLy(C_FDes):  End Property
Property Get LyTDes() As String(): LyTDes = ItmLy(C_TDes):  End Property
Property Get PkTny() As String(): PkTny = AyT1Ay(PkTFLy): End Property

Sub Z()
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
A = E
If A = "*Id" Then Exit Property
If A = "*Fk" Then Exit Property
EleLin = AyFstT1(EleLy, E)
Exit Property
X: Debug.Print "Schm.EleLin: PrpEr.."
End Property

Property Get EleSpecStr$()
On Error GoTo X
EleSpecStr = LinRmvT1(EleLin)
Exit Property
X: Debug.Print "Schm.EleSpecStr: PrpEr.."
End Property


Property Get FdStr$()
On Error GoTo X
FdStr = EleSpecStr
Exit Property
X: Debug.Print "Schm.FdStr: PrpEr.."
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
Case Else: Set Fd = NewFd_zFdStr(F, FdStr)
End Select
Exit Property
X: Debug.Print "Schm.Fd1: PrpEr.."
End Property

Property Get Td() As DAO.TableDef
On Error GoTo X
If No_T Then Exit Property
Set Td = NewTd(T, FdAy)
Exit Property
X: Debug.Print "Schm.Td: PrpEr.."
End Property

Property Get Tny() As String()
On Error GoTo X
Tny = AyMapSy(LyTFld, "LinT1")
Exit Property
X: Debug.Print "Schm.Tny: PrpEr.."
End Property

Property Get TdAy() As DAO.TableDef()
On Error GoTo X
Dim O() As DAO.TableDef
For Each T In Tny
    PushObj O, Td
Next
TdAy = O
Exit Property
X: Debug.Print "Schm.TdAy: PrpEr.."
End Property

Property Get PkSqy() As String()
On Error GoTo X
PkSqy = AyMapSy(PkTny, "TnPkSql")
Exit Property
X: Debug.Print "Schm.PkSqy: PrpEr.."
End Property

Property Get SkSslAy() As String()
On Error GoTo X
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

Property Get SkSqy() As String()
On Error GoTo X
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
Exit Property
X: Debug.Print "Schm.SkSqy: PrpEr.."
End Property

Sub Z_DbCrtSchm()
Dim Fb$
Fb = TmpFb
FbCrt Fb
DbCrtSchm FbDb(Fb), Z_Ly
Kill Fb
End Sub

Sub DbCrtSchm(A As Database, SchmLy$())
SetLy SchmLy
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

Property Get FdAy() As DAO.Field()
On Error GoTo X
Dim O() As DAO.Field
For Each F In Fny
    PushObj O, Fd
Next
FdAy = O
Exit Property
X: Debug.Print "Schm.FdAy: PrpEr.."
End Property

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
End Sub