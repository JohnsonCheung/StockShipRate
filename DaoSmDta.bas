Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Private _
  X_E() As DaoSmE _
, X_F() As DaoSmF _
, X_T() As DaoSmT _
, X_D() As DaoSmD _
, X_Eny() As String _
, X_Tny() As String
Property Get E() As DaoSmE(): E = X_E: End Property
Property Get T() As DaoSmT(): T = X_T: End Property
Property Get F() As DaoSmF(): F = X_F: End Property
Property Get D() As DaoSmD(): D = X_D: End Property
Property Get Eny() As String(): Eny = X_Eny: End Property
Property Get Tny() As String(): Tny = X_Tny: End Property

Property Let E(V() As DaoSmE): X_E = V: End Property
Property Let F(V() As DaoSmF): X_F = V: End Property
Property Let T(V() As DaoSmT): X_T = V: End Property
Property Let D(V() As DaoSmD): X_D = V: End Property
Property Let Eny(V$()): X_Eny = V: End Property
Property Let Tny(V$()): X_Tny = V: End Property