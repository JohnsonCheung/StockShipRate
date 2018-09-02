Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Private X_Md As CodeModule, Lno
Sub Ens()
If IsNothing(X_Md) Then Exit Sub
If X_Md.Parent.Type <> vbext_ct_ClassModule Then Exit Sub

Dim J%, L%(), Nm$
L = PrpLnoAy
If Not AyIsSrt(L) Then Stop
For J = UB(L) To 0 Step -1
    Lno = L(J)
    EnsOnePrp
Next
End Sub
Property Get PrpNm$()
PrpNm = LinPrpNm(Lin)
End Property
Private Sub EnsOnePrp()
Const C$ = "On Error Goto X"
Dim Nm$
Nm = PrpNm
If Nm = "" Then Stop
If HasSubStr(X_Md.Lines(Lno, 1), "End Property") Then
    Exit Sub
End If
If X_Md.Lines(Lno + 1, 1) <> C Then X_Md.InsertLines Lno + 1, C

'Ensure LblX
    Dim E$, L%, A$
    E = LblX_Expected_Lin
    L = LnoOf_LblX
    If L <> 0 Then A = X_Md.Lines(L, 1)
    If E <> A Then
        If L = 0 Then
            LblX_Ins
        Else
            LblX_Rpl L
        End If
    End If
'Ensure Exit Property
    L = LnoOf_InsExitPrp
    If L <> 0 Then X_Md.InsertLines L, "Exit Property"
End Sub
Property Get LblX_Expected_Lin$()
If IsNothing(X_Md) Then Exit Property
LblX_Expected_Lin = FmtQQ("X: Debug.Print ""?.?: PrpEr..""", MdNm(X_Md), PrpNm)
End Property
    
Property Get LnoOf_InsExitPrp%()
If LnoOf_ExitPrp <> 0 Then Exit Property
Dim L%
L = LnoOf_LblX
If L = 0 Then Stop
LnoOf_InsExitPrp = L
End Property
Sub LblX_Ins()
Dim L%
L = LnoOf_EndPrp
If L = 0 Then Stop
X_Md.InsertLines L - 1, LblX_Expected_Lin
End Sub
Sub LblX_Rpl(L%)
If L = 0 Then Stop
X_Md.ReplaceLine L, LblX_Expected_Lin
End Sub
Property Get LnoOf_EndPrp%()
LnoOf_EndPrp = MdLno_zEndPrp(X_Md, Lno)
End Property
Property Get LnoOf_ExitPrp%()
LnoOf_ExitPrp = MdLno_zExitPrp(X_Md, Lno)
End Property
Property Get LnoOf_LblX%()
Dim L%
L = LnoOf_EndPrp
If L = 0 Then Exit Property
If HasPfx(X_Md.Lines(L - 1, 1), "X: Debug.Print") Then LnoOf_LblX = L - 1
End Property


Property Get PrpLy() As String()
Dim O$(), A%(), J%, L$
A = PrpLnoAy
For J = 0 To UB(A)
    Lno = A(J)
    L = Lin
    Push O, Lin
Next
PrpLy = O
End Property
Property Get PrpNy() As String()
On Error GoTo X
Dim O$(), A%(), J%
A = PrpLnoAy
For J = 0 To UB(A)
    Lno = A(J)
    Push O, LinPrpNm(Lin)
Next
PrpNy = O
Exit Property
X: Debug.Print "PrpNy.."
End Property
Property Get PrpLnoAy() As Integer()
On Error GoTo X
Dim O%()
For Lno = 1 To X_Md.CountOfLines
    If IsPrpLin Then
        Push O, Lno
    End If
Next
PrpLnoAy = O
Exit Property
X: Debug.Print "PrpLnoAy.."
End Property
Property Get Lin$()
On Error GoTo X
If 1 <= Lno And Lno <= X_Md.CountOfLines Then
    Lin = X_Md.Lines(Lno, 1)
End If
Exit Property
X: Debug.Print "Lin.."
End Property
Property Get IsPrpLin() As Boolean
On Error GoTo X
Dim L$
L = Lin
If Not LinIsPrpLin(L) Then Exit Property
IsPrpLin = Not HasSfx(L, "End Property")
Exit Property
X: Debug.Print "IsPrpLin"
End Property
Sub A()
Stop
End Sub
Sub Z_Ini()
Set X_Md = CurVbe.VBProjects("StockShipRate").VBComponents("Schm").CodeModule
End Sub
Sub Z()
Z_Ini
'D PrpNy
'D PrpLnoAy 'EnsOnEr
Ens
End Sub