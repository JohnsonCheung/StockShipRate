Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Public Ffn$
Private X_TblNy$()
Property Get TblNy() As String()
TblNy = X_TblNy
End Property
Property Let TblNy(V$())
X_TblNy = V
End Property