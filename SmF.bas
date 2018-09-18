Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Public E$, LikT$
Private X_LikFny$()
Property Get LikFny() As String(): LikFny = X_LikFny: End Property
Property Let LikFny(V$()): X_LikFny = V: End Property