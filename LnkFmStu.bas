Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Public Stu$
Private X_Inp$()
Property Get Inp() As String()
Inp = X_Inp
End Property
Property Let Inp(V$())
X_Inp = V
End Property