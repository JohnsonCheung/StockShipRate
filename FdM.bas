Option Compare Database
Option Explicit
Function FdStr_Fd(A$, F) As DAO.Field2
Dim J%, L$, T$, Ay$(), Sz%, Rq As Boolean, Ty As DAO.DataTypeEnum, AlwZLen As Boolean, Dft$, VRul$, VTxt$
If A = "" Then Exit Function
Ay = AyRmvEmp(AyTrim(SplitSC(A)))
T = Ay(0)
Sz = DaoShtTy_Sz(T)
Ty = DaoShtTy_Ty(T)
For J = 1 To UB(Ay)
    L = Ay(J)
    Select Case True
    Case L = "Req": Rq = True
    Case L = "AlwZLen": AlwZLen = True
    Case HasPfx(L, "Dft="): Dft = RmvPfx(L, "Dft=")
    Case HasPfx(L, "VRul="): VRul = RmvPfx(L, "VRul=")
    Case HasPfx(L, "VTxt="): VTxt = RmvPfx(L, "VTxt=")
    Case Else: Debug.Print "FdSpec: there is itm[" & L & "] in EleLin[" & A & "] unexpected."
    End Select
Next
Dim O As New DAO.Field
With O
    .AllowZeroLength = AlwZLen
    .Name = F
    .DefaultValue = Dft
    .Required = Rq
    .Size = Sz
    .Type = Ty
    .ValidationRule = VRul
    .ValidationText = VTxt
End With
Set O = FdStr_Fd
End Function
Function NewFd_zFdStr(F, FdStr$) As DAO.Field2
Set NewFd_zFdStr = FdStr_Fd(FdStr, F)
End Function
Function FdStr$(A As DAO.Field)
Dim Ty$, Rq$, AlwZLen$
AlwZLen = IIf(A.AllowZeroLength, "*AlwZLen", "")
Ty = DaoTy_ShtTy(A.Type) & "." & A.Size
Rq = IIf(A.Required, "Req", "")
FdStr = ApLin(A.Name, Ty, Rq, AlwZLen, A.ValidationRule, A.ValidationText)
End Function