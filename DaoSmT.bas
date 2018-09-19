Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Public T$
Private X_Fny$()
Private X_Sk$()
Property Get Fny() As String(): Fny = X_Fny: End Property
Property Let Fny(V$()): X_Fny = V: End Property
Property Get Sk() As String(): Sk = X_Sk: End Property
Property Let Sk(V$()): X_Sk = V: End Property