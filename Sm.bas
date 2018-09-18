Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Private X_E() As SmE, X_T() As SmT, X_F() As SmF, X_D() As SmD, X_Er$()
Property Get E() As SmE(): E = X_E: End Property
Property Get T() As SmT(): T = X_T: End Property
Property Get D() As SmD(): D = X_D: End Property
Property Get F() As SmF(): F = X_F: End Property
Property Get Er() As String(): X_Er = Er: End Property
Property Let Er(V$()): X_Er = V: End Property
Property Let E(V() As SmE): X_E = V: End Property
Property Let F(V() As SmF): X_F = V: End Property
Property Let T(V() As SmT): X_T = V: End Property
Property Let D(V() As SmD): X_D = V: End Property