Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Public Fm$, Into$
Private X_Ny$(), X_Ey$()
Public Wh$
Property Get Ny() As String()
Ny = X_Ny
End Property
Property Get Ey() As String()
Ey = X_Ey
End Property
Property Let Ny(V$())
X_Ny = V
End Property
Property Let Ey(V$())
X_Ey = V
End Property