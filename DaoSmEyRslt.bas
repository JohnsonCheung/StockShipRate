Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Private X_E() As DaoSmE, X_Er() As String
Property Get Er() As String(): Er = X_Er: End Property
Property Get E() As DaoSmE(): E = X_E: End Property
Property Let Er(V$()): X_Er = V: End Property
Property Let E(V() As DaoSmE): X_E = V: End Property