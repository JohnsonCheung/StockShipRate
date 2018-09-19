Option Compare Database
Option Explicit
Private X_L As Database
Private X_Sess&
Private X_Msg&
Private X_Lg&

Public Const LgSchmNm$ = "LgSchm" ' The LgSchm-Spnm
Function BB1()
BB1 = 1
End Function

Property Get LgSchmLines$()
LgSchmLines = SpnmLines(LgSchmNm)
End Property

Property Get LgSchmLy() As String()
LgSchmLy = SpnmLy(LgSchmNm)
End Property

Sub LgSchmImp()
SpnmImp LgSchmNm
End Sub

Sub LgSchmExpIfNotExist()
SpnmExpIfNotExist LgSchmNm
End Sub

Sub LgSchmExp()
SpnmExp LgSchmNm
End Sub

Property Get LgSchmFt$()
LgSchmFt = SpnmFt(LgSchmNm)
End Property

Sub LgSchmBrw()
LgSchmExpIfNotExist
SpnmBrw LgSchmNm
End Sub

Sub LgSchmIni()
SpnmFtIni LgSchmNm
End Sub

Private Function L() As Database
On Error GoTo X
If IsNothing(X_L) Then
    LgOpn
End If
Set L = X_L
Exit Function
X:
Dim Er$, ErNo%
ErNo = Err.Number
Er = Err.Description
If ErNo = 3024 Then
    LgSchmImp
    LgCrt_v1
    LgOpn
    Set L = X_L
    Exit Function
End If
NyLyDmp "Err Er#", Er, ErNo
Stop
End Function

Sub LgBeg()
Lg ".", "Beg"
End Sub

Sub LgEnd()
Lg ".", "End"
End Sub

Private Sub LgOpn()
Set X_L = FbDb(LgFb)
End Sub

Sub LgCrt_v1()
Dim Fb$
Fb = LgFb
If FfnIsExist(Fb) Then Exit Sub
SchmM.DbCrtSchm FbCrt(Fb), LgSchmLy
End Sub

Sub LgCrt()
FbCrt LgFb
Dim Db As Database, T As DAO.TableDef
Set Db = FbDb(LgFb)
'
Set T = New DAO.TableDef
T.Name = "Sess"
TdAddId T
TdAddStamp T, "Dte"
Db.TableDefs.Append T
'
Set T = New DAO.TableDef
T.Name = "Msg"
TdAddId T
TdAddTxtFld T, "Fun"
TdAddTxtFld T, "MsgTxt"
TdAddStamp T, "Dte"
Db.TableDefs.Append T
'
Set T = New DAO.TableDef
T.Name = "Lg"
TdAddId T
TdAddLngFld T, "Sess"
TdAddLngFld T, "Msg"
TdAddStamp T, "Dte"
Db.TableDefs.Append T
'
Set T = New DAO.TableDef
T.Name = "LgV"
TdAddId T
TdAddLngFld T, "Lg"
TdAddLngTxt T, "Val"
Db.TableDefs.Append T

DbttCrtPk Db, "Sess Msg Lg LgV"
DbtCrtSk Db, "Msg", "Fun MsgTxt"
End Sub

Private Sub EnsSess()
If X_Sess > 0 Then Exit Sub
With L.TableDefs("Sess").OpenRecordset
    .AddNew
    X_Sess = !Sess
    .Update
    .Close
End With
End Sub

Private Sub EnsMsg(Fun$, MsgTxt$)
With L.TableDefs("Msg").OpenRecordset
    .Index = "Msg"
    .Seek "=", Fun, MsgTxt
    If .NoMatch Then
        .AddNew
        !Fun = Fun
        !MsgTxt = MsgTxt
        X_Msg = !Msg
        .Update
    Else
        X_Msg = !Msg
    End If
End With
End Sub

Property Get LgDb() As Database
Set LgDb = L
End Property

Property Get LgPth$()
LgPth = AppDtaPth
End Property

Private Sub WrtLg(Fun$, MsgTxt$)
With L.TableDefs("Lg").OpenRecordset
    .AddNew
    !Sess = X_Sess
    !Msg = X_Msg
    X_Lg = !Lg
    .Update
End With
End Sub

Sub Lg(Fun$, MsgTxt$, ParamArray Ap())
EnsSess
EnsMsg Fun, MsgTxt
WrtLg Fun, MsgTxt
Dim Av(): Av = Ap
If Sz(Av) = 0 Then Exit Sub
Dim J%, V
With L.TableDefs("LgV").OpenRecordset
    For Each V In Av
        .AddNew
        !Lines = VarLines(V)
        .Update
    Next
    .Close
End With
End Sub

Sub LgDbBrw()
Acs.OpenCurrentDatabase LgFb
AcsVis Acs
End Sub
Sub LgSchmKill()
FfnDltIfExist LgSchmFt
End Sub
Sub LgKill()
LgCls
If FfnIsExist(LgFb) Then Kill LgFb: Exit Sub
Debug.Print "LgFb-[" & LgFb & "] not exist"
End Sub

Sub LgCls()
On Error GoTo Er
X_L.Close
Er:
Set X_L = Nothing
End Sub

Property Get LgFb$()
LgFb = LgPth & LgFn
End Property
Property Get LgFn$()
LgFn = "Lg.accdb"
End Property

Sub SessBrw(Optional A&)
AyBrw SessLy(CvSess(A))
End Sub

Private Function CvSess&(A&)
If A > 0 Then CvSess = A: Exit Function
CvSess = DbqV(L, "select Max(Sess) from Sess")
End Function

Function SessLgAy(A&) As Long()
Q = FmtQQ("select Lg from Lg where Sess=? order by Lg", A)
SessLgAy = DbqLngAy(L, Q)
End Function

Function SessLy(Optional A&) As String()
Dim LgAy&()
LgAy = SessLgAy(A)
SessLy = AyOfAy_Ay(AyMap(LgAy, "LgLy"))
End Function

Sub LgAsg(A&, OSess&, ODTim$, OFun$, OMsgTxt$)
Q = FmtQQ("select Fun,MsgTxt,Sess,x.CrtTim from Lg x inner join Msg a on x.Msg=a.Msg where Lg=?", A)
Dim D As Date
RsAsg L.OpenRecordset(Q), OFun, OMsgTxt, OSess, D
ODTim = DteDTim(D)
End Sub

Function LgLy(A&) As String()
Dim Fun$, MsgTxt$, DTim$, Sess&, SFx$
LgAsg A, Sess, DTim, Fun, MsgTxt
SFx = FmtQQ(" @? Sess(?) Lg(?)", DTim, Sess, A)
LgLy = FunMsgLy(Fun & SFx, MsgTxt, LgLinesAy(A))
End Function

Function LgLinesAy(A&) As Variant()
Q = FmtQQ("Select Lines from LgV where Lg = ? order by LgV", A)
LgLinesAy = RsAy(L.OpenRecordset(Q))
End Function

Function CurLgRs(Optional Top% = 50) As DAO.Recordset
Set CurLgRs = L.OpenRecordset(FmtQQ("Select Top ? x.*,Fun,MsgTxt from Lg x left join Msg a on x.Msg=a.Msg order by Sess desc,Lg", Top))
End Function

Function CurLgLy(Optional Sep$ = " ", Optional Top% = 50) As String()
CurLgLy = RsLy(CurLgRs(Top), Sep)
End Function

Sub LgLis(Optional Sep$ = " ", Optional Top% = 50)
CurLgLis Sep, Top
End Sub

Sub CurLgLis(Optional Sep$ = " ", Optional Top% = 50)
D CurLgLy(Sep, Top)
End Sub

Sub SessLis(Optional Sep$ = " ", Optional Top% = 50)
CurSessLis Sep, Top
End Sub

Sub CurSessLis(Optional Sep$ = " ", Optional Top% = 50)
D CurSessLy(Sep, Top)
End Sub

Function CurSessLy(Optional Sep$, Optional Top% = 50) As String()
CurSessLy = RsLy(CurSessRs(Top), Sep)
End Function

Function CurSessRs(Optional Top% = 50) As DAO.Recordset
Set CurSessRs = L.OpenRecordset(FmtQQ("Select Top ? * from sess order by Sess desc", Top))
End Function

Function SessNLg%(A&)
SessNLg = DbqV(L, "Select Count(*) from Lg where Sess=" & A)
End Function

Private Sub Z_Lg()
LgKill
Debug.Assert Dir(LgFb) = ""
LgBeg
Debug.Assert Dir(LgFb) = LgFn
End Sub


Sub Z()
Z_Lg
End Sub