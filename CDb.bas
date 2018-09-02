Option Compare Database
Option Explicit
Property Get CDbTny() As String()
CDbTny = DbTny(CurrentDb)
End Property
Property Get Tny() As String()
Tny = CDbTny
End Property

Property Get CDbAttNy() As String()
CDbAttNy = DbAttNy(CurrentDb)
End Property

Property Get CDbCnSy() As String()
CDbCnSy = DbCnSy(CurrentDb)
End Property

Property Get CDbLnkTny() As String()
CDbLnkTny = DbLnkTny(CurrentDb)
End Property

Sub CDbDrpLnkTbl()
DbDrpLnkTbl CurrentDb
End Sub
Function CDbScly() As String()
CDbScly = DbScly(CurrentDb)
End Function
Function CDbPth$()
CDbPth = FfnPth(CurrentDb.Name)
End Function