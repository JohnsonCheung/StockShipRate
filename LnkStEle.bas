Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Public Ele$, Stu$
Private X_Fny$()
Property Get Fny() As String()
Fny = X_Fny
End Property
Property Let Fny(V$())
X_Fny = V
End Property