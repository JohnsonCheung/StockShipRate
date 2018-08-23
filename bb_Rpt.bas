Option Compare Database
Option Explicit
Public Y As Byte
Public M As Byte
Public LnkColStr As New LnkColStr
Function FmDte() As Date
FmDte = DateSerial(Y, M, 1)
End Function
Function ToDte() As Date
ToDte = DteLasDayOfMth(FmDte)
End Function
Function FmYYYYxMMxDD$()
FmYYYYxMMxDD = Format(FmDte, "YYYY-MM-DD")
End Function
Function ToYYYYxMMxDD$()
ToYYYYxMMxDD = Format(ToDte, "YYYY-MM-DD")
End Function
Sub AssYM()
If Y = 0 Or M = 0 Then Stop
End Sub
Function BegY() As Byte
AssYM
BegY = IIf(M = 1, Y - 1, Y)
End Function
Function BegM() As Byte
AssYM
BegM = IIf(M = 1, 12, M - 1)
End Function
Function NxtY() As Byte
AssYM
NxtY = IIf(M = 12, Y + 1, Y)
End Function
Function NxtM() As Byte
AssYM
NxtM = IIf(M = 12, 1, M + 1)
End Function
Function YYYYxMM$()
YYYYxMM = YYYY & "-" & MM
End Function
Function MM$()
MM = Format(M, "00")
End Function
Function YYYY$()
YYYY = Format(2000 + Y)
End Function

Function QQRun(QQ$)
Dim IQ, Q$
For Each IQ In CvNy(QQ)
    Q = IQ
    Select Case True
    Case HasPfx(Q, "O"): Q = "@" & Mid(IQ, 2)
    Case HasPfx(Q, "Tmp")
    Case Else: Stop
    End Select
    
    MsgRunQry Q
    Run IQ
Next
End Function
Sub MsgRunQry(A$)
MsgSet "Running query (" & A & ") ..."
End Sub

Function OupPth$()
Dim A$
A = CurDbPth & "Output\"
PthEns A
OupPth = A
End Function
Function IFbStkShpRate$()
If IsDev Then
    IFbStkShpRate = CurrentDb.Name
Else
    IFbStkShpRate = "N:\SAPAcessReports\StockShipRate\StockShipRate_Data.accdb"
End If
End Function
Function OupFx$()
OupFx = OupPth & FmtQQ("? ?.xlsx", Apn, YYYYxMM)
End Function
Private Sub MsgSet(A$)
Form_Main.MsgSet A
End Sub
Private Sub MsgClr()
Form_Main.MsgClr
End Sub

Function IFxUOM$()
IFxUOM = PnmFfn("UOM")
End Function



Sub DocUOM()
'InpX: [>UOM]     Material [Base Unit of Measure] [Material Description] [Unit per case]
'Oup : UOM        Sku      SkuUOM                 Des                    Sc_U

'Note on [Sales text.xls]
'Col  Xls Title            FldName     Means
'F    Base Unit of Measure SkuUOM      either COL (bottle) or PCE (set)
'J    Unit per case        Sc_U        how many unit per AC
'K    SC                   SC_U        how many unit per SC   ('no need)
'L    COL per case         AC_B        how many bottle per AC
'-----
'Letter meaning
'B = Bottle
'AC = act case
'SC = standard case
'U = Unit  (Bottle(COL) or Set (PCE))

' "SC              as SC_U," & _  no need
' "[COL per case]  as AC_B," & _ no need
End Sub
Sub TblYM_NxtYM(OY As Byte, OM As Byte)
Dim YM%
YM = SqlV("Select Max(Y*100+M) from YM")
OY = YM \ 100
OM = YM Mod 100
If OM = 12 Then
    OM = 1
    OY = OY + 1
Else
    OM = OM + 1
End If
End Sub
Sub TblYM_Ins()
TblYM_NxtYM Y, M
YM_Ins Y, M
End Sub
Sub YM_Ins(Y As Byte, M As Byte)
Dim J%, I%
For J = Y To CurY - 1
    For I = 1 To 12
        If Not SqlAny(FmtQQ("Select Y from [YM] where Y=? and M=?", J, I)) Then
            DoCmd.RunSQL FmtQQ("Insert into [YM] (Y,M) values (?,?)", J, I)
        End If
    Next
Next
For I = 1 To CurM
    If Not SqlAny(FmtQQ("Select Y from [YM] where Y=? and M=?", CurY, I)) Then
        DoCmd.RunSQL FmtQQ("Insert into [YM] (Y,M) values (?,?)", CurY, I)
    End If
Next
End Sub
Sub TblYM_Ini()
If Y = 0 Or M = 0 Then Stop
Dim NRec%
NRec = SqlV("Select Count(*) from YM where Y<" & Y)
If NRec > 0 Then
    If MsgBox(FmtQQ("There are [?] months of data before year[?] month[?].   Delete them", NRec, Y, M) & "?", vbYesNo) <> vbYes Then Exit Sub
End If
DoCmd.RunSQL FmtQQ("Delete * from YM where Y<? or (Y=? and M<?)", Y, Y, M)
YM_Ins Y, M
End Sub

Sub ZZ_YM_Ini()
Y = 18
M = 1
TblYM_Ini
End Sub