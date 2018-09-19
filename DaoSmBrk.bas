Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Private X_Er() As String
Public Dta As DaoSmDta
Property Get Er() As String(): X_Er = Er: End Property
Property Let Er(V$()): X_Er = V: End Property