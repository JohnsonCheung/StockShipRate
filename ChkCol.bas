Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Public FilKind$, Ffn$, T$
Private X_F() As EptFldTy
Property Get F() As EptFldTy()
F = X_F
End Property
Property Let F(V() As EptFldTy)
X_F = V
End Property