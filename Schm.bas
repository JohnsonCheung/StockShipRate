Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Const C_E$ = "E"
Const C_T$ = "T"
Const C_F$ = "F"
Const C_D$ = "D"
Private X_Schmy$()

Friend Function Init(SchmLines$) As Schm
X_Schmy = SplitCrLf(SchmLines)
Set Init = Me
End Function

Property Get Ly()
On Error GoTo X
Ly = X_Schmy
Exit Property
X: Debug.Print "Schm.Ly.PrpEr...["; Err.Description; "]"
End Property

Sub SetLy(Ly$())
X_Schmy = Ly
End Sub

Function Tz(T) As SchmT
Dim X   As New SchmT
Set Tz = X.Init(Me, T)
End Function

Property Get TFELy() As String()
On Error GoTo X
Dim O$(), T, F
For Each T In AyNz(Tny)
    For Each F In AyNz(Tz(T).Fny)
        Push O, ApLin(T, F, Tz(T).Fz(F).E)
    Next
Next
TFELy = O
Exit Property
X: Debug.Print "Schm.TFELy.PrpEr...["; Err.Description; "]"
End Property
Property Get TFEF1Ly() As String()
On Error GoTo X
Dim O$(), T, I
For Each T In AyNz(Tny)
    For Each I In AyNz(Tz(T).Fzy)
        With CvFz(I)
           Push O, ApLin(T, .F, .E, .ESpec)
        End With
    Next
Next
TFEF1Ly = O
Exit Property
X: Debug.Print "Schm.TFEF1Ly.PrpEr...["; Err.Description; "]"
End Property

Private Function ItmLy(A) As String()
ItmLy = AyT1Chd(Ly, A)
End Function

Property Get ErPfx() As String()
On Error GoTo X
If Sz(X_Schmy) = 0 Then
    ErPfx = Sy("no Ly is given")
    Exit Property
End If
ErPfx = AyWhPredXPNot(Ly, "LinIsInT1Ay", Sy(C_E, C_D, C_F, C_T))
Exit Property
X: Debug.Print "Schm.ErPfx.PrpEr...["; Err.Description; "]"
End Property

Property Get ErNoTFld() As String()
On Error GoTo X
If Sz(TLy) = 0 Then ErNoTFld = Sy("No TFld lines")
Exit Property
X: Debug.Print "Schm.ErNoTFld.PrpEr...["; Err.Description; "]"
End Property

Property Get ErDupT() As String()
On Error GoTo X
ErDupT = AyDupChk(Tny, "These T[?] is duplicated in TFld-lines")
Exit Property
X: Debug.Print "Schm.ErDupT.PrpEr...["; Err.Description; "]"
End Property

Property Get EAy() As String()
On Error GoTo X
EAy = AyT1Ay(ELy)
Exit Property
X: Debug.Print "Schm.EAy.PrpEr...["; Err.Description; "]"
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

Property Get ErDupE() As String()
On Error GoTo X
ErDupE = AyDupChk(EAy, "These Ele[?] are duplicated in Ele-lines")
Exit Property
X: Debug.Print "Schm.ErDupE.PrpEr...["; Err.Description; "]"
End Property

Private Sub Z_ErDupF()
Dim Ly$()
Ly = Sy("T AA BB BB")
Expect = Sy("These F[BB] are duplicated in T[AA]")
GoSub Tst
Exit Sub
Tst:
    SetLy Ly
    Actual = ErDupF
    C
    Stop
    Return
End Sub

Private Sub Z_ErDupT()
Dim Ly$()
Ly = Sy("T AA BB BB", "T AA DD")
Expect = Sy("These T[AA] is duplicated in TFld-lines")
GoSub Tst
Exit Sub
Tst:
    SetLy Ly
    Actual = ErDupT
    C
    Return
End Sub

Property Get ErDupF() As String()
On Error GoTo X
Dim T, Fny$(), O$(), M$
For Each T In AyNz(Tzy)
    With CvTz(T)
        M = FmtQQ("These F[?] are duplicated in T[?]", "?", .T)
        Stop
        PushAy O, AyDupChk(.Fny, M)
    End With
Next
ErDupF = O
Exit Property
X: Debug.Print "Schm.ErDupF.PrpEr...["; Err.Description; "]"
End Property

Property Get ErEle() As String()
On Error GoTo X
ErEle = AyDupChk(EAy, "These Ele[?] are duplicated in Ele-lines")
Exit Property
X: Debug.Print "Schm.ErEle.PrpEr...["; Err.Description; "]"
End Property

Property Get ErFldHasNoEle() As String()
On Error GoTo X
Dim T, F
For Each T In AyNz(Tny)
    For Each F In Tz(T).Fzy
        With CvFz(F)
            If .E = "" Then
                Push ErFldHasNoEle, FmtQQ("T[?] F[?] cannot be found in any EF-lines", T, .F)
            End If
        End With
    Next
Next
Exit Property
X: Debug.Print "Schm.ErFldHasNoEle.PrpEr...["; Err.Description; "]"
End Property

Property Get Er() As String()
On Error GoTo X
Er = AyAddAp(ErPfx, ErNoTFld, ErDupT, ErDupF, ErDupE, ErEle, ErFldHasNoEle)
Exit Property
X: Debug.Print "Schm.Er.PrpEr...["; Err.Description; "]"
End Property

Property Get FLy() As String():     FLy = ItmLy(C_F): End Property
Property Get TLy() As String():     TLy = ItmLy(C_T): End Property
Property Get ELy() As String():     ELy = ItmLy(C_E):   End Property
Property Get DLy() As String():     DLy = ItmLy(C_D): End Property

Sub Z()
Z_ErDupT
Z_ErDupF
Z_ErDupE
Z_Tny
Exit Sub
Z_DbCrtSchm
End Sub

Sub Z_Ini()
X_Schmy = LgIniSchmy
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
    D Tz(T).Fny
    GoSub Sep
    Return
Sep:
    D "--------------------"
    Return
End Sub

Property Get Tny() As String()
On Error GoTo X
Tny = AyMapSy(TLy, "LinT1")
Exit Property
X: Debug.Print "Schm.Tny.PrpEr...["; Err.Description; "]"
End Property

Property Get TdAy() As DAO.TableDef()
On Error GoTo X
TdAy = OyPrpInto(Tzy, "Td", TdAy)
Exit Property
X: Debug.Print "Schm.TdAy.PrpEr...["; Err.Description; "]"
End Property

Property Get PkSqy() As String()
On Error GoTo X
PkSqy = AyRmvEmp(OyPrpSy(Tzy, "PkSql"))
Exit Property
X: Debug.Print "Schm.PkSqy.PrpEr...["; Err.Description; "]"
End Property

Property Get Tzy() As SchmT()
On Error GoTo X
Tzy = AyMapObjFunXInto(Tny, Me, "Tz", Tzy)
Exit Property
X: Debug.Print "Schm.Tzy.PrpEr...["; Err.Description; "]"
End Property

Property Get SkSqy() As String()
On Error GoTo X
SkSqy = AyRmvEmp(OyPrpSy(Tzy, "SkSql"))
Exit Property
X: Debug.Print "Schm.SkSqy.PrpEr...["; Err.Description; "]"
End Property

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

Property Get Scly() As String()
On Error GoTo X
Dim O$(), I
For Each I In AyNz(Tzy)
    PushAy O, CvTz(I).Scly
Next
Scly = O
Exit Property
X: Debug.Print "Schm.Scly.PrpEr...["; Err.Description; "]"
End Property

Sub Z_ErPfx()
Dim Ly$()
Expect = Sy("No Ly is given")
GoSub Tst
Exit Sub
Tst:
    SetLy Ly
    Actual = ErPfx
    C
    Return
End Sub