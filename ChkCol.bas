Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Public FilKind$, Ffn$, T$
Private X_F() As ExpectFldTy
Property Get F() As ExpectFldTy()
F = X_F
End Property
Property Let F(V() As ExpectFldTy)
X_F = V
End Property