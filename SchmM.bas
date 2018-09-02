Option Compare Database
Option Explicit
Const C_FEle$ = "FEle"
Const C_Ele$ = "Ele"
Const C_TFld$ = "TFld"
Const C_TDes$ = "TDes"
Const C_FDes$ = "FDes"
Private X_Ly$()
Function E$(T, F)
Select Case True
Case IsId(T, F): E = "*Id"
Case IsFk(F): E = "*Fk"
Case Else: E = LinT1(LinFEle(F))
End Select
End Function

Function LinFEle$(F)
Dim A$(), L
A = LyFEle
If Sz(A) = 0 Then Exit Function
For Each L In A
    If StrInLikSsl(F, LinRmvT1(L)) Then
        LinFEle = L
        Exit Function
    End If
Next
End Function

Function Ly()
Ly = X_Ly
End Function
Sub SetLy(Ly$())
X_Ly = Ly
End Sub

Public Function Z_Ly() As String()
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
End Function

Function TFELy() As String()
Dim O$(), T, F
For Each T In Tny
    For Each F In Fny(T)
        Push O, ApLin(T, F, E(T, F))
    Next
Next
TFELy = O
End Function

Function TFEFdLy() As String()
Dim O$(), T, F, E1
For Each T In Tny
    For Each F In Fny(T)
        E1 = E(T, F)
        Push O, ApLin(T, F, E1, FdStr(E1))
    Next
Next
TFEFdLy = O
End Function

Function ItmLy(A) As String()
ItmLy = AyT1Chd(Ly, A)
End Function

Function Ly_Er() As String()
Ly_Er = AyWhPredXPNot(Ly, "LinInT1Ay", Sy(C_Ele, C_FDes, C_TDes, C_TFld))
End Function

Function EleLy() As String():  EleLy = ItmLy(C_Ele):    End Function
Function LyFEle() As String(): LyFEle = ItmLy(C_FEle):  End Function
Function LyTFld() As String(): LyTFld = ItmLy(C_TFld):  End Function
Function LyFDes() As String(): LyFDes = ItmLy(C_FDes):  End Function
Function LyTDes() As String(): LyTDes = ItmLy(C_TDes):  End Function
Function PkTny() As String(): PkTny = AyT1Ay(PkTFLy):   End Function

Sub Z()
Z_Ini
Z_Tny
Z_DbCrtSchm
End Sub
Private Sub Z_Ini()
If Sz(X_Ly) = 0 Then X_Ly = Z_Ly
End Sub

Sub Z_Tny()
Z_Ini
Expect = SslSy("Sess Msg Lg LgV")
Actual = Tny
C
End Sub

Sub ZZ_Tny()
Dim T
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
    D Fny(T)
    GoSub Sep
    Return
Sep:
    D "--------------------"
    Return
End Sub

Function EleLin$(E)
EleLin = AyFstT1(EleLy, E)
End Function

Function EleSpecStr$(E)
EleSpecStr = LinRmvT1(EleLin(E))
End Function
Function FdStr$(E)
FdStr = EleSpecStr(E)
End Function

Function Fd(T, F, EleSpecStr$) As DAO.Field
Select Case True
Case IsId(T, F): Set Fd = NewFd_zId(F)
Case IsFk(F): Set Fd = NewFd_zFk(F)
Case Else: Set Fd = NewFd_zFdStr(F, EleSpecStr)
End Select
End Function

Function Td(T) As DAO.TableDef
Set Td = NewTd(T, FdAy(T))
End Function

Function Tny() As String()
Tny = AyMapSy(LyTFld, "LinT1")
End Function

Function TdAy() As DAO.TableDef()
Dim O() As DAO.TableDef, T
For Each T In Tny
    PushObj O, Td(T)
Next
TdAy = O
End Function

Function PkSqy() As String()
PkSqy = AyMapSy(PkTny, "TnPkSql")
End Function

Function SkSslAy() As String()
Dim A$(), O$(), L
A = LyTFld
If Sz(A) = 0 Then Exit Function
For Each L In A
    PushNonEmpty O, SkSsl(L)
Next
SkSslAy = O
End Function

Function SkSsl$(L)
Dim A$, B$
A = SkP1(L): If A = "" Then Exit Function
B = Replace(A, " * ", "")
SkSsl = Replace(B, "*", LinT1(B))
End Function

Function SkP1$(L)
SkP1 = Trim(TakBef(L, "|"))
End Function
Function PkTFLy() As String()
PkTFLy = AyWhPred(LyTFld, "TFLinHasPk")
End Function

Function SkSqy() As String()
Dim O$(), A$(), B$(), J%, U%, T
A = SkSslAy
U = UB(A)
If UB(A) = -1 Then Exit Function
ReDim O(U)
For J = 0 To U
    T = LinShiftT1(A(J))
    O(J) = TnSkSql(T, A(J))
Next
SkSqy = O
End Function

Sub Z_DbCrtSchm()
Dim Fb$
Fb = TmpFb
FbCrt Fb
DbCrtSchm1 FbDb(Fb), Z_Ly
FbBrw Fb
End Sub

Sub DbCrtSchm1(A As Database, SchmLy$())
SetLy SchmLy
AyDoPX TdAy, "DbAppTd", A
AyDoPX PkSqy, "DbRun", A
AyDoPX SkSqy, "DbRun", A
End Sub

Function TFLin$(T)
TFLin = AySng(AyWhT1EqV(LyTFld, T), "Schm.TFLin.PrpEr")
End Function

Function Fny(T) As String()
Dim A$, B$
A = TFLin(T)
If LinShiftT1(A) <> T Then Debug.Print "Schm.Fny PrpEr": Exit Function
B = Replace(A, "*", T)
Fny = AyRmvEle(SslSy(B), "|")
End Function

Function FdAy(T) As DAO.Field()
Dim O() As DAO.Field, F
For Each F In Fny(T)
    PushObj O, Fd(T, F, E(T, F))
Next
FdAy = O
End Function

Function IsFk(F) As Boolean
IsFk = AyHas(Tny, F)
End Function

Function IsId(T, F) As Boolean
IsId = T = F
End Function