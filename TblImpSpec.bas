Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Public Tbl$, LnkColStr$, WhBExpr$
Friend Property Get Init(Tbl$, LnkColStr$, Optional WhBExpr$)
Me.Tbl = Tbl
Me.LnkColStr = LnkColStr
Me.WhBExpr = WhBExpr
Set Init = Me
End Property