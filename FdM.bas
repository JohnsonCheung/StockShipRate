Option Compare Database
Option Explicit
Type FdSpec
    F As String
    Ty As dao.DataTypeEnum
    Sz As Integer
    AlwZLen As Boolean
    Dft As String
    VRul As String
    VTxt As String
    Req As Boolean
End Type
Private A As dao.Field
Function EleSpecStr_FdSpec(A$, F) As FdSpec
Dim J%, L$, T$, Ay$(), Sz%, Rq As Boolean, Ty As dao.DataTypeEnum, AlwZLen As Boolean, Dft$, VRul$, VTxt$
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
With EleSpecStr_FdSpec
    .AlwZLen = AlwZLen
    .F = F
    .Dft = Dft
    .Req = Rq
    .Sz = Sz
    .Ty = Ty
    .VRul = VRul
    .VTxt = VTxt
End With
End Function
Function FdStr$(Fd As dao.Field)
Set A = Fd
FdStr = ApLin(Nm, Ty, Rq, AlwZLen, VRul, VTxt)
End Function
Private Function Nm$()
Nm = A.Name
End Function
Private Function VTxt$()
VTxt = A.ValidationText
End Function
Private Function VRul$()
VRul = A.ValidationRule
End Function
Private Function AlwZLen$()
AlwZLen = IIf(A.AllowZeroLength, "*AlwZLen", "")
End Function
Private Function Ty$()
Ty = DaoTy_ShtTy(A.Type) & "." & A.Size
End Function
Private Function Rq$()
Rq = IIf(A.Required, "Req", "")
End Function