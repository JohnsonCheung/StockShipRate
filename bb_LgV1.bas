Option Compare Database
Option Explicit
Sub LgCrt()
Dim SsnLines$
SsnLines = SchemaLines
Const SsnLines$ = _
"Tbl Sess " & vbCrLf & _
"Fld ll " & vbCrLf & _
"df" & vbCrLf & _
"jksdf"

Const SfxFsnLines$ = _
"Txt Memo" & vbCrLf & _
"Dte Date" & vbCrLf & _
""
Const TsnLines$ = ""
Const SsnLines$ = TsnLines & vbCrLf & NmFsnLines & vbCrLf & SfxFsnLines
FbCrt LgFb
DbCrtSchema FbDb(LgFb), SsnLines
Exit Sub
End Sub