Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Public Schm As Schm, T$
Friend Function Init(A As Schm, T) As SchmT
Set Schm = A
Me.T = T
Set Init = Me
End Function

Property Get TLy() As String()
On Error GoTo X
TLy = Schm.TLy
Exit Property
X: Debug.Print "SchmT.TLy.PrpEr...["; Err.Description; "]"
End Property
Property Get TLin$()
On Error GoTo X
TLin = AySng(AyWhT1(TLy, T), "Schmy.TLin.PrpEr")
Exit Property
X: Debug.Print "SchmT.TLin.PrpEr...["; Err.Description; "]"
End Property

Property Get Fny() As String()
On Error GoTo X
Dim A$, B$
A = TLin
If LinShiftT1(A) <> T Then Debug.Print "SchmT.Fny PrpEr": Exit Property
B = Replace(A, "*", T)
Fny = AyRmvEle(SslSy(B), "|")
Exit Property
X: Debug.Print "SchmT.Fny.PrpEr...["; Err.Description; "]"
End Property

Property Get FdAy() As DAO.Field()
On Error GoTo X
Dim O() As DAO.Field
Dim F, I
For Each I In Fzy
    Set F = I
    PushObj O, F.Fd
Next
FdAy = O
Exit Property
X: Debug.Print "SchmT.FdAy.PrpEr...["; Err.Description; "]"
End Property

Function Fz(F) As SchmF
Dim O As New SchmF
Set Fz = O.Init(Schm, T, F)
End Function

Property Get Fzy() As SchmF()
On Error GoTo X
Dim O() As SchmF, F
For Each F In AyNz(Fny)
    PushObj O, Fz(F)
Next
Fzy = O
Exit Property
X: Debug.Print "SchmT.Fzy.PrpEr...["; Err.Description; "]"
End Property
Property Get Td() As DAO.TableDef
On Error GoTo X
Set Td = NewTd(T, FdAy)
Exit Property
X: Debug.Print "SchmT.Td.PrpEr...["; Err.Description; "]"
End Property
Property Get SkSql$()
Dim A$
A = SkSsl: If A = "" Then Exit Function
SkSql = SqlzCrtSk(T, SslSy(A))
End Property
Property Get SkSsl$()
On Error GoTo X
SkSsl = TLinSkSsl(TLin)
Exit Property
X: Debug.Print "SchmT.SkSsl.PrpEr...["; Err.Description; "]"
End Property
Property Get Scly() As String()
On Error GoTo X
Scly = AyIns(FdScly, Scl)
Exit Property
X: Debug.Print "SchmT.Scly.PrpEr...["; Err.Description; "]"
End Property

Property Get PkSql$()
If AyHas(Fny, T) Then PkSql = SqlzCrtPk(T)
End Property

Property Get FdScly() As String()
On Error GoTo X
Dim O$(), F, Ay
Ay = AyNz(Fzy)
For Each F In AyNz(Fzy)
    Push O, CvFz(F).Scl
Next
FdScly = O
Exit Property
X: Debug.Print "SchmT.FdScly.PrpEr...["; Err.Description; "]"
End Property

Property Get Scl$()
On Error GoTo X
Scl = ApScl(T, Des)
Exit Property
X: Debug.Print "SchmT.Scl.PrpEr...["; Err.Description; "]"
End Property

Property Get Des$()
On Error GoTo X
Des = DLyDes_zT(Schm.DLy, T)
Exit Property
X: Debug.Print "SchmT.Des.PrpEr...["; Err.Description; "]"
End Property