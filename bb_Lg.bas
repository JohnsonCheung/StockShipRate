Option Compare Database
Option Explicit
Private X_L As Database
Private X_Sess&
Private X_Msg&
Private X_Lg&
Private Function L() As Database
If IsNothing(X_L) Then
    Set X_L = FbDb(LgFb)
End If
Set L = X_L
End Function
Sub LgEns()
If Not FfnIsExist(LgFb) Then LgCrt
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
DbtCrtSk Db, "Msg", "Msg", "Fun MsgTxt"
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
    If .EOF Then
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
Sub WrtLg(Fun$, MsgTxt$)
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
        !Val = VarLines(V)
        .Update
    Next
    .Close
End With
End Sub
Sub LgBrw()
Acs.OpenCurrentDatabase LgFb
AcsVis Acs
End Sub

Sub LgKill()
LgCls
FfnDltIfExist LgFb
End Sub

Sub LgCls()
On Error GoTo Er
X_L.Close
Er:
Set X_L = Nothing
End Sub

Property Get LgFb$()
LgFb = WPth & LgFn
End Property
Property Get LgFn$()
LgFn = "Lg.accdb"
End Property

Sub SessBrw(Optional A&)
AyBrw SessLy(CvSess(A))
End Sub

Function CvSess&(A&)
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

Function LgLy(A&) As String()
Dim Fun$, MsgTxt$
Q = FmtQQ("select Fun,MsgTxt from Msg where Msg in (Select Msg from Lg where Lg=?)", A)
RsAsg L.OpenRecordset(Q), Fun, MsgTxt
LgLy = FunMsgAv_Ly(Fun, MsgTxt, LgValAy(A))
End Function

Function LgValAy(A&) As Variant()
Q = FmtQQ("Select Val from LgV where Lg = ? order by LgV", A)
LgValAy = RsAy(L.OpenRecordset(Q))
End Function

Sub SessLis(Optional Top% = 50)
With L.OpenRecordset(FmtQQ("Select Top ? * from sess order by Sess desc", Top))
    While Not .EOF
        D !Sess & " " & DteDTim(!Dte) & " NLg-" & SessNLg(CLng(!Sess))
        .MoveNext
    Wend
    .Close
End With
End Sub

Function SessNLg%(A&)
SessNLg = DbqV(L, "Select Count(*) from Lg where Sess=" & A)
End Function