Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Public F$
Private Ty() As ADODB.DataTypeEnum
Property Get TyAy() As ADODB.DataTypeEnum()
TyAy = Ty
End Property
Property Let TyAy(V() As ADODB.DataTypeEnum)
Ty = V
End Property