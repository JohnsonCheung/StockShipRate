Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
Dim X_PmFx() As PmFil
Dim X_PmFb() As PmFil
Dim X_PmSw() As PmSw
Dim X_Inp() As String
Dim X_IpSw() As IpSw
Dim X_IpFx() As IpFil
Dim X_IpFb() As IpFil
Dim X_IpS1() As String
Dim X_IpWs() As IpWs
Dim X_IpWh() As IpWh
Dim X_StInp() As StInp
Dim X_StEle() As StEle
Dim X_StExt() As StExt
Dim X_StFld() As StFld

Friend Function Init(PmFx() As PmFil, PmFb() As PmFil, PmSw() As PmSw, _
Inp$(), IpSw() As IpSw, IpFx() As IpFil, IpFb() As IpFil, IpS1$(), IpWs() As IpWs, IpWh() As IpWh, _
StInp() As StInp, StEle() As StEle, StExt() As StExt, StFld() As StFld)
X_PmFx = PmFx
X_PmFb = PmFb
X_PmSw = PmSw

X_Inp = Inp
X_IpSw = IpSw
X_IpFx = IpFx
X_IpFb = IpFb
X_IpS1 = IpS1
X_IpWs = IpWs
X_IpWh = IpWh

X_StInp = StInp
X_StEle = StEle
X_StExt = StExt
X_StFld = StFld
Set Init = Me
End Function

Function PmFx() As PmFil():   PmFx = X_PmFx:  End Function
Function PmFb() As PmFil():   PmFb = X_PmFb:  End Function
Function PmSw() As PmSw():    PmSw = X_PmSw:  End Function
Function Inp() As String():    Inp = X_Inp:   End Function
Function IpSw() As IpSw():    IpSw = X_IpSw:  End Function
Function IpFx() As IpFil():   IpFx = X_IpFx:  End Function
Function IpFb() As IpFil():   IpFb = X_IpFb:  End Function
Function IpS1() As String():  IpS1 = X_IpS1:  End Function
Function IpWs() As IpWs():    IpWs = X_IpWs:  End Function
Function IpWh() As IpWh():    IpWh = X_IpWh:  End Function
Function StInp() As StInp(): StInp = X_StInp: End Function
Function StEle() As StEle(): StEle = X_StEle: End Function
Function StExt() As StExt(): StExt = X_StExt: End Function
Function StFld() As StFld(): StFld = X_StFld: End Function