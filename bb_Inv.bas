Option Compare Database
Option Explicit
Private PrcTim As Date
Const ILnkColStr_zInvH$ = ""
Const ILnkColStr_zInvD$ = ""

Sub InvImp()
Dim Fx, FxAy$(), O$()
PrcTim = Now
FxAy = IFxAy
If Sz(FxAy) = 0 Then MsgBox "No file in import folder", vbCritical: Exit Sub
For Each Fx In FxAy
    PushAy O, IFxImp(CStr(Fx))
Next
If AyBrwEr(O) Then InvPthBrw
End Sub
Function IFxAy() As String()
IFxAy = PthFfnAy(InvPth, "*.xlsx")
End Function
Function IFxImp(A$) As String()
WIni
Dim B$(), C$(), O$()
B = WtLnkFx(">InvH", A, "InvH")
C = WtLnkFx(">InvD", A, "InvD")
O = AyAdd(B, C)
If Sz(O) > 0 Then IFxImp = O: Exit Function
B = WtChkCol(">InvH", ILnkColStr_zInvH)
C = WtChkCol(">InvD", ILnkColStr_zInvD)
O = AyAdd(B, C)
If Sz(O) > 0 Then IFxImp = O: Exit Function

End Function
Sub InvExp()
Dim O$
O = CurDbPth & "Import Shipment Invoice\"
PthEns O
End Sub
Function InvPth$()

End Function
Sub InvPthBrw()
PthBrw InvPth
End Sub