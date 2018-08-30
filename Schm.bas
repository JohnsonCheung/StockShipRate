Option Compare Database
Option Explicit
Private Type SchmLinesBrk
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
Private X As SchmLinesBrk
Private T, F
Public Const SampleSchmLines$ = _
"Ty_Fld Mem Lines ..          " & vbCrLf & _
"Ty_Fld Txt Fun ..            " & vbCrLf & _
"Ty_Sfx Dte Dte ..            " & vbCrLf & _
"Ty_Sfx Txt Txt ..            " & vbCrLf & _
"Dft Now | CrtDte ..          " & vbCrLf & _
"Req Lines Fun MsgTxt ..      " & vbCrLf & _
"TFld Sess * CrtDte           " & vbCrLf & _
"TFld Msg  * Fun *Txt | CrtDte" & vbCrLf & _
"TFld Lg   * Sess Msg CrtDte  " & vbCrLf & _
"TFld LgV  * Lg Lines         " & vbCrLf & _
"FDes Fun Function name that call the log" & vbCrLf & _
"FDes Fun Function name that call the log" & vbCrLf & _
"TDes Msg it will a new record when Lg-function is first time using the Fun+MsgTxt" & _
"TDes Msg ..."
Sub AA()
Stop
End Sub
Private Property Get ZZX() As SchmLinesBrk
ZZX = SchmLines_Brk(SampleSchmLines)
End Property
Sub ZZZ_TySz()
X = ZZX
T = "Sess"
F = "CrtDte"
Expect = "Dte"
GoSub Tst
Exit Sub
Tst:
    Actual = TySz
    C
    Return
End Sub

Private Sub ZZZ_Req()
X = ZZX
F = "Lines":  Debug.Assert Req = True
F = "Fun":    Debug.Assert Req = True
F = "MsgTxt": Debug.Assert Req = True
F = "XX":     Debug.Assert Req = False
End Sub

Sub ZZ_Dft()
X = ZZX
F = "CrtDte":  Debug.Assert Dft = "Now"
F = "Fun":    Debug.Assert Dft = ""
End Sub

Sub AAA()
End Sub
Private Sub ZZZ_Tny()
X = ZZX
Expect = SslSy("Sess Msg Lg LgV")
Actual = Tny
C
End Sub

Sub ZZ_Tny()
X = ZZX
GoSub Sep
D "Tny"
D "---"
D Tny
GoSub Sep
For Each T In Tny
    GoSub Prt
Next
D SkSql
D PkSql
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

Private Sub ZZ_SchmLines_BrkAsg()
Dim Td() As dao.TableDef, Pk$(), Sk$()
SchmLines_BrkAsg SampleSchmLines, Td, Pk, Sk
Stop
End Sub
Private Function SchmLines_Brk(SchmLines) As SchmLinesBrk
With SchmLines_Brk
    LinesBrkAsg1 SchmLines, .XErLy, .XRmkDic, _
        "FDes   Dft   Req   Ty_Fld   Ty_Sfx   Ty_TF   TDes   TFld", _
        .FDes, .Dft, .Req, .Ty_Fld, .Ty_Sfx, .Ty_TF, .TDes, .TFld
    .Req = SslAy_Sy(.Req)
End With
End Function
Private Sub SchmLines_BrkAsg(A, OTdAy() As dao.TableDef, OPkSqlAy$(), OSkSqlAy$())
X = SchmLines_Brk(A)
OPkSqlAy = PkSql
OSkSqlAy = SkSql
OTdAy = TdAy
End Sub

Private Function Fd() As dao.Field
Dim O As dao.Field, IsId As Boolean, Sz%, Ty As dao.DataTypeEnum
IsId = T = F
If IsId Then
    Set Fd = DaoFld(F, IsId:=IsId)
Else
    DaoShtTySz_BrkAsg TySz, Ty, Sz
    Set Fd = DaoFld(F, Ty, Sz, , IsId, Dft, Req)
End If
End Function

Private Function Req() As Boolean
Dim L
For Each L In X.Req
    If AyHas(SslSy(L), F) Then
        Req = True
        Exit Function
    End If
Next
End Function

Private Property Get TF$()
TF = T & " " & F
End Property

Private Function TySz_TF$()
If Sz(X.Ty_TF) = 0 Then Exit Function
Dim A$, L
A = TF
For Each L In X.Ty_TF
    If HasPfx(L, A) Then TySz_TF = RmvPfx(L, A): Exit Function
Next
End Function

Private Function TySz_F$()
If Sz(X.Ty_Fld) = 0 Then Exit Function
Dim L, O$
For Each L In X.Ty_Fld
    O = LinShiftT1(L)
    If AyHas(SslSy(L), F) Then TySz_F = O: Exit Function
Next
End Function

Private Function TySz_Sfx$()
If Sz(X.Ty_Sfx) = 0 Then Exit Function
Dim L, O$
For Each L In X.Ty_Sfx
    O = LinShiftT1(L)
    If StrInSfxAy(F, SslSy(L)) Then TySz_Sfx = O: Exit Function
Next
End Function
Private Function TySz_Id$()
If AyHas(Tny, F) Then TySz_Id = "Lng"
End Function
Private Function TySz$()
TySz = TySz_Id:  If TySz <> "" Then Exit Function
TySz = TySz_TF:  If TySz <> "" Then Exit Function
TySz = TySz_F:   If TySz <> "" Then Exit Function
TySz = TySz_Sfx: If TySz <> "" Then Exit Function
Stop
End Function

Private Function Dft$()
If Sz(X.Dft) = 0 Then Exit Function
Dim L, O$, Ssl$
For Each L In X.Dft
    BrkAsg L, "|", O, Ssl
    If AyHas(SslSy(Ssl), F) Then
        Dft = O
        Exit Function
    End If
Next
End Function

Private Function Td() As dao.TableDef
Set Td = NewTd(T, FdAy)
End Function

Private Function Tny() As String()
Tny = AyMapSy(X.TFld, "LinT1")
End Function

Private Function TdAy() As dao.TableDef()
Dim O() As dao.TableDef
For Each T In Tny
    PushObj O, Td
Next
TdAy = O
End Function

Private Function PkSql() As String()
Dim PkTny$()
    Dim B$()
    B = AyWhPred(X.TFld, "TFLinHasPk")
    PkTny = AyMapSy(B, "LinT1")
PkSql = AyMapSy(PkTny, "TnPkSql")
End Function
Private Function SkSql() As String()
Dim J%, U%, B$(), O$()
Dim T$, SkSsl$, Lin$
B = AyWhPred(X.TFld, "TFLinHasSk")
B = AyMapXPSy(B, "TakBef", "|")
U = UB(B)
If U = -1 Then Exit Function
ReDim O(U)
For J = 0 To U
    Lin = B(J): GoSub X
    O(J) = TnSkSql(T, SkSsl)
Next
SkSql = O
Exit Function
X:
    Dim A$
    BrkAsg Lin, " ", T, A
    SkSsl = Replace(RmvPfx(A, "*"), "*", T)
    Return
End Function

Sub ZZ_DbCrtSchm()
Dim Fb$
Fb = TmpFb
FbCrt Fb
DbCrtSchm FbDb(Fb), SampleSchmLines
FbBrw Fb
End Sub

Sub DbCrtSchm(A As Database, SchmLines$)
Dim Td() As dao.TableDef, Pk$(), Sk$()
SchmLines_BrkAsg SchmLines, Td, Pk, Sk
AyDoPX Td, "DbAppTd", A
AyDoPX Pk, "DbRun", A
AyDoPX Sk, "DbRun", A
End Sub

Private Function Fny() As String()
Dim A$(), B$, C$, Tbl$
A = LyWhT1EqV(X.TFld, T)
If Sz(A) <> 1 Then Stop
B = A(0)
Tbl = LinShiftT1(B)
If T <> Tbl Then Stop
C = Replace(B, "*", Tbl)
Fny = AyRmvEle(SslSy(C), "|")
End Function

Private Function FdAy() As dao.Field()
Dim O() As dao.Field
For Each F In Fny
    PushObj O, Fd
Next
FdAy = O
End Function