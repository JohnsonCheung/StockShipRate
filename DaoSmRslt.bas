Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Private _
  X_Er() As String _
, X_SkSqy() As String _
, X_PkSqy() As String _
, X_Td() As DAO.TableDef _
, X_FDes() As FDes _
, X_TDes() As TDes

Property Get Er() As String(): Er = X_Er: End Property
Property Get Td() As DAO.TableDef(): Td = X_Td: End Property
Property Get SkSqy() As String(): SkSqy = X_SkSqy: End Property
Property Get PkSqy() As String(): PkSqy = X_PkSqy: End Property
Property Get TDes() As TDes(): TDes = X_TDes: End Property
Property Get FDes() As FDes(): FDes = X_FDes: End Property

Property Let Er(V$()): X_Er = V: End Property
Property Let SkSqy(V$()): X_SkSqy = V: End Property
Property Let PkSqy(V$()): X_PkSqy = V: End Property
Property Let Td(V() As DAO.TableDef): X_Td = V: End Property
Property Let TDes(V() As TDes): X_TDes = V: End Property
Property Let FDes(V() As FDes): X_FDes = V: End Property