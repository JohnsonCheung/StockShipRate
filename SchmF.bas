Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Schm As Schm, T$, F$
Friend Function Init(Schm, T, F) As SchmF
Set Me.Schm = Schm
Me.T = T
Me.F = F
Set Init = Me
End Function

Property Get ELin$()
On Error GoTo X
ELin = AyFstT1(Schm.ELy, E)
Exit Property
X: Debug.Print "SchmF.ELin.PrpEr...["; Err.Description; "]"
End Property

Property Get ESpec$()
On Error GoTo X
Select Case True
Case IsId: ESpec = "*Id"
Case IsFk: ESpec = "*Fk"
Case Else: ESpec = LinRmvTT(ELin)
End Select
Exit Property
X: Debug.Print "SchmF.ESpec.PrpEr...["; Err.Description; "]"
End Property

Property Get IsFk() As Boolean
On Error GoTo X
IsFk = AyHas(Schm.Tny, F)
Exit Property
X: Debug.Print "SchmF.IsFk.PrpEr...["; Err.Description; "]"
End Property

Property Get IsId() As Boolean
On Error GoTo X
IsId = T = F
Exit Property
X: Debug.Print "SchmF.IsId.PrpEr...["; Err.Description; "]"
End Property
Property Get E$()
On Error GoTo X
E = LinT1(FLin)
Exit Property
X: Debug.Print "SchmF.E.PrpEr...["; Err.Description; "]"
End Property
Property Get FLin$()
On Error GoTo X
FLin = T1LikLikSslAy_T1(Schm.FLy, T, F)
Exit Property
X: Debug.Print "SchmF.FLin.PrpEr...["; Err.Description; "]"
End Property
Property Get ESpecScl$()
ESpecScl = JnSC(LinTermAy(ESpec))
End Property
Property Get Scl$()
On Error GoTo X
Scl = ApScl(F, ESpec)
Exit Property
X: Debug.Print "SchmF.Scl.PrpEr...["; Err.Description; "]"
End Property

Property Get Fd() As DAO.Field
On Error GoTo X
If F = "" Then Exit Property
Select Case True
Case IsId: Set Fd = NewFd_zId(F)
Case IsFk: Set Fd = NewFd_zFk(F)
Case Else: Set Fd = NewFd_zFdScl(Scl)
End Select
Exit Property
X: Debug.Print "SchmF.Fd.PrpEr...["; Err.Description; "]"
End Property

Friend Sub Z()
Init NewSchm(LgSchmLines), "LgV", "Lines"
Stop
End Sub