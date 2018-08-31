Option Compare Database
Option Explicit
Const C_Ele$ = "Ele"
Const C_TFld$ = "TFld"
Const C_TDes$ = "TDes"
Const C_FDes$ = "FDes"
Const C_Req$ = "Req"

Public Lines$
Type SchmLinesBrk
    Ty_TF() As String
    Ty_Fld() As String
    Ty_Sfx() As String
    TFld() As String
    Dft() As String
    Req() As String
    FDes() As String
    TDes() As String
    XRmkDic As Dictionary 'Each given pfx-line will have its remark here.  [Rmk] is any [']-line above the lien.  The [Key] is Key+Ix
    XErLy() As String    'Any line other than given Pfx
End Type
Public T, F, L
Private X1$(), X2$

Private Sub X(A)
Push X1, A
End Sub

Public Property Get Z_Lines$()
If X2 <> "" Then Z_Lines = X2: Exit Property
X "dfd"
X "Ele Lines Mem"
X "Ele Amt   Cur;Dft=0"
X "Ele Txt   Txt"
X "Ele Nm    T20;Req;NonEmp"
X "Ele Crt   Dte;Req;Dft=Now;"
X "Ele Des   Txt"
X "Ele Sc    Dbl"
X "FEle Amt . *Amt"
X "FEle Dte . *Dte"
X "FEle Txt . Fun *Txt"
X "FEle Crt CrtDte"
X "Req  * Lines Fun MsgTxt .."
X "TFld Sess * CrtDte"
X "TFld Msg  * Fun *Txt | CrtDte"
X "TFld Lg   * Sess Msg CrtDte"
X "TFld LgV  * Lg Lines"
X "FDes Fun Function name that call the log"
X "FDes Fun Function name that call the log"
X "TDes Msg it will a new record when Lg-function is first time using the Fun+MsgTxt"
X "TDes Msg ..."
X2 = JnCrLf(X1)
Z_Lines = X2
End Property
Function ItmLy(A) As String()
ItmLy = AyWhT1EqV(Ly, A)
End Function
Property Get Ly() As String()
Ly = SplitCrLf(Lines)
End Property

Property Get Ly_Er() As String()
On Error GoTo X
Ly_Er = AyWhPredXPNot(Ly, "LinInT1Ay", Sy(C_Ele, C_FDes, C_Req, C_TDes, C_TFld))
Exit Property
X:
Debug.Print "Ly_Er"
End Property

Property Get Lyz_Ele() As String():  Lyz_Ele = ItmLy(C_Ele):   End Property
Property Get Lyz_Req() As String():  Lyz_Req = ItmLy(C_Req):   End Property
Property Get Lyz_TFld() As String(): Lyz_TFld = ItmLy(C_TFld): End Property
Property Get Lyz_FDes() As String(): Lyz_FDes = ItmLy(C_FDes): End Property
Property Get Lyz_TDes() As String(): Lyz_TDes = ItmLy(C_TDes): End Property
Property Get PkTny() As String(): PkTny = AyT2Ay(AyWhPred(Lyz_TFld, "TFLinHasPk")): End Property
Property Get SkTny() As String(): SkTny = AyT2Ay(AyWhPred(Lyz_TFld, "TFLinHasSk")): End Property

Sub Z_TySz()
Lines = Z_Lines
T = "Sess"
F = "CrtDte"
Expect = "Dte"
GoSub Tst
Exit Sub
Tst:
'    Actual = TySz
    C
    Return
End Sub
Property Get Req() As Boolean

End Property
Sub Z_Req()
Lines = Z_Lines
F = "Lines":  Debug.Assert Req = True
F = "Fun":    Debug.Assert Req = True
F = "MsgTxt": Debug.Assert Req = True
F = "XX":     Debug.Assert Req = False
End Sub

Sub Z_Dft()
Z_Ini
F = "CrtDte":  Debug.Assert Dft = "Now"
F = "Fun":     Debug.Assert Dft = ""
End Sub
Sub Z_Ini()
Lines = Z_Lines
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
Property Get E$()

End Property
Property Get EleLin$()
If No_F Then Exit Property
End Property

Function FdDr() As Variant()
Dim A$
A = EleLin
FdDr = Array(F, Ty)
End Function

Property Get Fd() As DAO.Field
If T = F Then
    Set Fd = NewFd_zId(F)
Else
    Set Fd = NewFd_zDr(FdDr)
End If
End Property
Property Get FReq() As Boolean
Exit Property
For Each L In Lyz_Ele
    If AyHas(SslSy(L), F) Then
        FReq = True
        Exit Property
    End If
Next
End Property

Function Ty() As DAO.DataTypeEnum
If Sz(Lyz_TFld) = 0 Then Exit Function
Dim A$, L
A = T & " " & F
For Each L In Lyz_TFld
'    If HasPfx(L, A) Then TySz_zTF = RmvPfx(L, A): Exit Function
Next
End Function

Function TySz_zF$()
'If Sz(Lyz_Fld) = 0 Then Exit Function
'Dim L, O$
'For Each L In X.Ty_Fld
'    O = LinShiftT1(L)
'    If AyHas(SslSy(L), F) Then TySz_zF = O: Exit Function
'Next
End Function
Function EFd() As DAO.Field2
'Ty = FIdTy:  If Ty <> 0 Then Exit Function
'Ty = TFTy:   If Ty <> 0 Then Exit Function
'Ty = FTy:    If Ty <> 0 Then Exit Function
'Ty = FSfxTy: If Ty <> 0 Then Exit Function
Stop
End Function

Property Get Dft$()
Exit Property
Dim O$, Ssl$
'For Each L In LDft
'    BrkAsg L, "|", O, Ssl
'    If AyHas(SslSy(Ssl), F) Then
'        Dft = O
'        Exit Property
'    End If
'Next
End Property

Property Get Td() As DAO.TableDef
Set Td = NewTd(T, FdAy)
End Property

Property Get Tny() As String()
Tny = AyMapSy(Lyz_TFld, "LinT2")
End Property

Property Get TdAy() As DAO.TableDef()
Dim O() As DAO.TableDef
For Each T In Tny
    PushObj O, Td
Next
TdAy = O
End Property

Property Get PkSqy() As String()
PkSqy = AyMapSy(PkTny, "TnPkSql")
End Property

Function TFldSkSsl$(A)
Dim Rest$
Rest = LinRmvT1(A)
TFldSkSsl = Replace(RmvPfx(Rest, "*"), "*", T)
End Function

Property Get SkSslAy() As String()
SkSslAy = AyMapSy(AyWhPred(Lyz_TFld, "TFLinHasSk"), "TFldSkSsl")
End Property

Property Get SkSqy() As String()
Dim O$(), A$(), B$(), J%, U%
A = SkTny
U = UB(A)
If Sz(A) = 0 Then Stop
B = SkSslAy
If U <> UB(B) Then Stop
ReDim O(U)
For J = 0 To U
    O(J) = TnSkSql(A(J), B(J))
Next
SkSqy = O
End Property

Sub ZZ_DbCrtSchm()
Dim Fb$
Fb = TmpFb
FbCrt Fb
DbCrtSchm FbDb(Fb), Z_Lines
FbBrw Fb
End Sub

Sub DbCrtSchm(A As Database, SchmLines$)
Lines = SchmLines
AyDoPX TdAy, "DbAppTd", A
AyDoPX PkSqy, "DbRun", A
AyDoPX SkSqy, "DbRun", A
End Sub
Property Get No_T() As Boolean
No_T = T = ""
End Property
Function AySng(A)
If Sz(A) <> 1 Then Stop
Asg A(0), AySng
End Function
Property Get TFLin$()
If No_T Then Exit Property
TFLin = AySng(AyWhT2EqV(Lyz_TFld, T))
End Property

Property Get Fny() As String()
Dim A$, B$
A = TFLin
If LinShiftT1(A) <> "TFld" Then Debug.Print "Schm.Fny.PrpEr2": Exit Property
If LinShiftT1(A) <> T Then Debug.Print "Schm.Fny.PrpEr3": Exit Property
B = Replace(A, "*", T)
Fny = AyRmvEle(SslSy(B), "|")
End Property

Property Get No_F() As Boolean
No_F = F = ""
End Property

Property Get FdAy() As DAO.Field()
Dim O() As DAO.Field
For Each F In Fny
    PushObj O, Fd
Next
FdAy = O
End Property

Sub A()
Lines = Z_Lines
T = "Sess"
Stop
End Sub