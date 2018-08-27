Option Compare Database
Option Explicit
Sub LgCrt()
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
'
Set T = New DAO.TableDef
T.Name = "Sess"
TdAddId T
TdAddStamp T, "Dte"
Db.TableDefs.Append T
'
Set T = New DAO.TableDef
T.Name = "Msg"
TdAddId T
TdAddTxtFld T, "Fun"
TdAddTxtFld T, "MsgTxt"
TdAddStamp T, "Dte"
Db.TableDefs.Append T
'
Set T = New DAO.TableDef
T.Name = "Lg"
TdAddId T
TdAddLngFld T, "Sess"
TdAddLngFld T, "Msg"
TdAddStamp T, "Dte"
Db.TableDefs.Append T
'
Set T = New DAO.TableDef
T.Name = "LgV"
TdAddId T
TdAddLngFld T, "Lg"
TdAddLngTxt T, "Val"
Db.TableDefs.Append T

DbttCrtPk Db, "Sess Msg Lg LgV"
DbtCrtSk Db, "Msg", "Msg", "Fun MsgTxt"
End Sub