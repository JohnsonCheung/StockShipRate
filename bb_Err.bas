Option Compare Database
Option Explicit
Private StsM$()
Private ErrM$()
Private Sub BrwSts()
AyBrwEr StsM
Erase StsM
End Sub
Private Sub BrwBrw()
AyBrwEr ErrM
Erase ErrM
End Sub
Private Sub PushErr(Msg$, ParamArray Ap())
Dim Av(): Av = Ap
AyPushMsgAv ErrM, Msg, Av
End Sub
Private Sub PushSts(Msg$, ParamArray Ap())
Dim Av(): Av = Ap
AyPushMsgAv StsM, Msg, Av
End Sub