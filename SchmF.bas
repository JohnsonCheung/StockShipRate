Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Public Schm As Schm, T$, F$
Friend Property Get Init(Schm, T, F) As SchmF
Set Me.Schm = Schm
Me.T = T
Me.F = F
End Property
Property Get ELy() As String()
On Error GoTo X
ELy = Schm.ELy
Exit Property
X: Debug.Print "SchmF.ELy.PrpEr...["; Err.Description; "]"
End Property

Property Get ELin$()
On Error GoTo X
ELin = AyFstT1(ELy, E)
Exit Property
X: Debug.Print "SchmF.ELin.PrpEr...["; Err.Description; "]"
End Property

Property Get ESpec$()
On Error GoTo X
Select Case True
Case IsId: ESpec = "*Id"
Case IsFk: ESpec = "*Fk"
Case Else: ESpec = LinRmvT1(ELin)
End Select
Exit Property
X: Debug.Print "SchmF.ESpec.PrpEr...["; Err.Description; "]"
End Property
Property Get Tny() As String()
On Error GoTo X
Tny = Schm.Tny
Exit Property
X: Debug.Print "SchmF.Tny.PrpEr...["; Err.Description; "]"
End Property
Property Get IsFk() As Boolean
On Error GoTo X
IsFk = AyHas(Tny, F)
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

Property Get Scl$()
On Error GoTo X
Scl = ApScl(F, EleSpec)
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