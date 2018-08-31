Option Compare Database
Option Explicit
Private A As DAO.Field
Function FdLin$(Fd As DAO.Field)
Set A = Fd
FdLin = ApLin(Nm, Ty, Rq, AlwZer, Rul, VdtTxt)
End Function
Private Function Nm$()
Nm = A.Name
End Function
Private Function VdtTxt$()

End Function
Private Function Rul$()

End Function
Private Function AlwZer$()

End Function
Private Function Ty$()
Ty = A.Type
End Function
Private Function Rq$()

End Function