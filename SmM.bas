Option Compare Database
Option Explicit
Public Const SmLines$ = "E Mem | Mem Req AlwZLen" & _
vbCrLf & "E Txt | Txt Req" & _
vbCrLf & "E Crt | Dte Req Dft=Now" & _
vbCrLf & "E Crt | Dte Req Dft=Now" & _
vbCrLf & "E Dte | Dte" & _
vbCrLf & "F Amt * | *Amt" & _
vbCrLf & "F Crt * | CrtDte" & _
vbCrLf & "F Dte * | *Dte" & _
vbCrLf & "F Txt * | Fun * Txt" & _
vbCrLf & "F Mem * | Lines" & _
vbCrLf & "T Sess | * CrtDte" & _
vbCrLf & "T Msg  | * Fun *Txt | CrtDte" & _
vbCrLf & "T Lg   | * Sess Msg CrtDte" & _
vbCrLf & "T LgV  | * Lg Lines" & _
vbCrLf & "D . Fun | Function name that call the log" & _
vbCrLf & "D . Fun | Function name that call the log" & _
vbCrLf & "D . Msg | it will a new record when Lg-function is first time using the Fun+MsgTxt" & _
vbCrLf & "D . Msg | ..."
Private Type E
    E As String
    Req As Boolean
    Ty As DAO.DataTypeEnum
    AlwZ As Boolean
    TxtSz As Byte
    VRul As String
    VTxt As String
    Dft As String
End Type
Private Type F
    E As String
    LikT As String
    LnkFny() As String
End Type
Private Type D
    T As String
    F As String
    Des As String
End Type
Private Type T
    T As String
    Fny() As String
End Type
Private Type Rslt
    Er() As String
    SkSqy() As String
    PkSqy() As String
    Td() As DAO.TableDef
    FDes() As FDes
    TDes() As TDes
End Type
Private Type Dta
    E() As E
    F() As F
    T() As T
    D() As D
    Tny() As String
End Type
Private Type Brk
    Er() As String
    Dta As Dta
End Type

Sub DbCrtSchm(A As Database, SmLines$)
With Rslt(SmLines)
    AyBrwThw .Er
    AyDoPX .Td, "DbAddTd", A
    AyDoPX .PkSqy, "DbRun", A
    AyDoPX .SkSqy, "DbRun", A
    AyDoPX .FDes, "DbSetFDes", A
    AyDoPX .TDes, "DbSetTDes", A
End With
End Sub

Private Function Brk(SmLines$) As Brk
Stop '
End Function
Private Function ErNoTny(Tny$()) As String()
Stop '
End Function
Private Function ErDupT(Tny$()) As String()
Stop '
End Function

Private Function ErDupF(A() As T) As String()
Stop '
End Function

Private Function ErDupE(A() As E) As String()
Stop '
End Function

Private Function ErFldEleIsInVdt() As String()
Stop '
End Function

Private Function Er(A As Brk) As String()
Dim D As Dta
D = A.Dta
Er = AyAddAp( _
     A.Er _
   , ErNoTny(D.Tny) _
   , ErDupT(D.Tny) _
   , ErDupE(D.E) _
   , ErFldEleIsInVdt() _
   , ErDupF(D.T))
End Function

Private Function PkSqy(A As Dta) As String()

End Function

Private Function IsSk(TLin$) As Boolean

End Function

Private Function SkSsl$(TLin$)

End Function

Private Function TItm(T, A() As T) As T
Dim J%
For J = 0 To UBound(A)
    With A(J)
        If .T = T Then TItm = A(J): Exit Function
    End With
Next
End Function

Private Function SkSql$(T, TBrk() As T)
Dim Fm$, Into$, Ny$(), M As T
Dim Ny$()
M = TItm(T, TBrk)
If Sz(M.Sk) = 0 Then Exit Function
Ny = M.Fny
Fm = ">" & T
Into = "#I" & T
SkSql = SqlzCrtSk(T, M.Sk)
TnPkSql
End Function
Private Function SkSqy(A As Dta) As String()
Dim T, O$()
For Each T In A.Tny
    PushNonEmp O, SkSql(T, A.T)
Next
SkSqy = O
End Function
Private Function Td(A As Dta) As DAO.TableDef()
Dim O() As DAO.TableDef, I
For Each I In A.Tny
    PushObj O, NewTd(I, FdAy(I, A))
Next
Td = O
End Function
Private Function FdAy(T, A As Dta) As DAO.Field()

End Function
Private Function TDes(A As Dta) As TDes()

End Function
Private Function FDes(A As Dta) As FDes()

End Function

Private Function Rslt(SmLines$) As Rslt
Dim B As Brk
    B = Brk(SmLines)
Dim E$()
    E = Er(B)
    If Sz(E) > 0 Then Rslt.Er = E: Exit Function
With Rslt
    Dim D As Dta
    D = B.Dta
    .PkSqy = PkSqy(D)
    .SkSqy = SkSqy(D)
    .Td = Td(D)
    .FDes = FDes(D)
    .TDes = TDes(D)
End With
End Function