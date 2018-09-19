Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Private X_D() As DaoSmD, X_Er() As String
Property Get Er() As String(): Er = X_Er: End Property
Property Get D() As DaoSmD(): D = X_D: End Property
Property Let Er(V$()): X_Er = V: End Property
Property Let D(V() As DaoSmD): E = V: End Property