Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Public F As DaoSmF
Private X_Er() As String
Property Get Er() As String(): Er = X_Er: End Property
Property Let Er(V$()): X_Er = V: End Property