Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Private X_F() As DaoSmF, X_Er() As String
Property Get F() As DaoSmF(): F = X_F: End Property
Property Get Er() As String(): Er = X_Er: End Property
Property Let Er(V$()): X_Er = V: End Property
Property Let F(V() As DaoSmF): X_F = V: End Property