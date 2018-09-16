Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Public Fil$
Private X_Inpy$()
Property Get Inpy() As String()
Inpy = X_Inpy
End Property
Property Let Inpy(V$())
X_Inpy = V
End Property