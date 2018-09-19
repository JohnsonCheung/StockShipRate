'Option Compare Database
'Option Explicit
'Const C_F$ = "F"
'Const C_E$ = "E"
'Const C_T$ = "T"
'Const C_D$ = "D"
'Private ZZF$
'Private ZZT$
'Private Type Des
'    F() As FDes
'    T() As TDes
'End Type
'Private Type Rslt
'    Er() As String
'    FDes() As FDes
'    TDes() As TDes
'    Sk() As String
'    Pk() As String
'    Td() As DAO.TableDef
'End Type
'Private Type Dta
'    Eny() As String
'    Tny() As String
'    T() As String
'    F() As String
'    E() As String
'    D() As String
'End Type
'Private Type Brk
'    Dta As Dta
'    Er() As String
'End Type
'
'Private Function FLyTF_E$(A$(), T, F)
'FLyTF_E = LinT1(FLyTF_FLin(A, T, F))
'End Function
'
'Private Sub Z_FLyTF_FLin()
'ZZT = "Msg"
'ZZF = "MsgTxt"
'Ept = "Txt * | Fun * Txt"
'Act = ZZFLin
'C
'End Sub
'
'Private Property Get ZZFLin$()
'ZZFLin = FLyTF_FLin(ZZFLy, ZZT, ZZF)
'End Property
'
'Private Function FLyTF_FLin$(A$(), T, F)
'FLyTF_FLin = T1LikLikSslAy_FstT2T3Eq(A, T, F)
'End Function
'
'Private Property Get ZZELy() As String()
'ZZELy = ClnELy(ZZCln)
'End Property
'
'Private Property Get ZZSchmy() As String()
'ZZSchmy = LgIniSchmy
'End Property
'
'Private Property Get ZZQTFELy() As String()
'ZZQTFELy = SchmyQTFELy(ZZSchmy)
'End Property
'
'Function SchmyQTFELy(A$()) As String()
'Dim Cln$()
'Cln = LyCln(A)
'SchmyQTFELy = QTFELy(ClnTLy(Cln), ClnFLy(Cln))
'End Function
'
'Function QTFELy(TLy$(), FLy$()) As String()
'Dim O$(), T, F, E$
'For Each T In AyNz(AyT1Ay(TLy))
'    For Each F In AyNz(TLyT_Fny(TLy, T))
'        E = FLyTF_E(FLy, T, F)
'        Push O, ApLin(T, F, E)
'    Next
'Next
'QTFELy = O
'End Function
'
'Function DtaTFEFdLy(A As Dta) As String()
'Dim O$(), T, F, E$
'For Each T In AyNz(A.Tny)
'    For Each F In TLyT_Fny(A.F, T)
'        E = FLyTF_E(A.F, T, F)
'        Push O, ApLin(T, F, E, ELyFE_FdScl(A.E, F, E))
'    Next
'Next
'DtaTFEFdLy = O
'End Function
'
'Function ClnFLy(A$()) As String(): ClnFLy = AyT1Chd(A, C_F): End Function
'Function ClnTLy(A$()) As String(): ClnTLy = AyT1Chd(A, C_T): End Function
'Function ClnELy(A$()) As String(): ClnELy = AyT1Chd(A, C_E): End Function
'Function ClnDLy(A$()) As String(): ClnDLy = AyT1Chd(A, C_D): End Function
'
'Function TLyPkTny(A$()) As String()
'TLyPkTny = AyT1Ay(AyWhPred(A, "TLinHasPk"))
'End Function
'
'
'Private Sub ZZ_FdAy()
'ZZT = "Sess"
'Act = ZZFdAy
'Stop
'End Sub
'
'Private Property Get ZZFLy() As String()
'ZZFLy = ClnFLy(ZZSchmy)
'End Property
'Private Function ZZDta() As Dta
'
'End Function
'Private Function ZZBrk() As Brk
'
'End Function
'
'Private Function ZZFdAy() As DAO.Field2()
'ZZFdAy = DtaT_FdAy(ZZDta, ZZT)
'End Function
'
'Private Property Get ZZTny() As String()
'ZZTny = AyT1Ay(ZZTLy)
'End Property
'
'Private Sub Z_Tny()
'ChkEq ZZTny, SslSy("Sess Msg Lg LgV")
'End Sub
'
'Private Sub ZZ_Tny()
'Dim T, Tny$(), TLy$()
'TLy = ZZTLy
'Tny = ZZTny
'GoSub Sep
'D "Tny"
'D "---"
'D ZZTny
'GoSub Sep
'For Each T In Tny
'    GoSub Prt
'Next
'D TLySkSqy(TLy)
'D TLyPkSqy(TLy)
'Exit Sub
'Prt:
'    D T
'    D UnderLin(T)
'    D TLyT_Fny(TLy, T)
'    GoSub Sep
'    Return
'Sep:
'    D "--------------------"
'    Return
'End Sub
'
'Function ELyE_ELin$(A$(), E)
'ELyE_ELin = AyFstT1(A, E)
'End Function
'
'Function ELyE_EScl$(A$(), E)
'ELyE_EScl = LinRmvT1(ELyE_ELin(A, E))
'End Function
'
'Function ELyFE_FdScl$(A$(), F, E)
'ELyFE_FdScl = F & ";" & ELyE_EScl(A, E)
'End Function
'
'Private Function Fd(F, T, Tny$(), FLy$(), ELy$()) As DAO.Field
'Select Case True
'Case T = F: Set Fd = NewFd_zId(F)
'Case AyHas(Tny, F): Set Fd = NewFd_zFk(F)
'Case Else:
'    Dim E$, FdScl$
'    E = FLyTF_E(FLy, T, F)
'    FdScl = ELyFE_FdScl(ELy, F, E)
'    Set Fd = NewFd_zFdScl(FdScl)
'End Select
'End Function
'
'Private Function DtaT_Td(A As Dta, T) As DAO.TableDef
'Set DtaT_Td = NewTd(T, DtaT_FdAy(A, T))
'End Function
'
'Private Function DtaTdAy(A As Dta) As DAO.TableDef()
'Dim O() As DAO.TableDef, T
'For Each T In A.Tny
'    PushObj O, DtaT_Td(A, T)
'Next
'DtaTdAy = O
'End Function
'
'Function TLyPkSqy(A$()) As String()
'TLyPkSqy = AyMapSy(TLyPkTny(A), "TnPkSql")
'End Function
'Function TLySkSslAy(A$()) As String()
'Dim O$(), L
'If Sz(A) = 0 Then Exit Function
'For Each L In A
'    PushNonEmp O, TLinSkSsl(L)
'Next
'TLySkSslAy = O
'End Function
'
'Function TLinSkSsl$(A)
'Dim B$, C$
'B = Trim(TakBef(A, "|")): If B = "" Then Exit Function
'C = Replace(B, " * ", " ")
'TLinSkSsl = Replace(C, "*", LinT1(A))
'End Function
'
'Private Property Get ZZSkSqy() As String()
'ZZSkSqy = TLySkSqy(ZZTLy)
'End Property
'
'Private Property Get ZZTLy() As String()
'ZZTLy = ClnTLy(ZZSchmy)
'End Property
'
'Function TLySkSqy(A$()) As String()
'TLySkSqy = AyMapSy(AyRmvEmp(TLySkSslAy(A)), "TnSkSsl_SkSql")
'End Function
'
'Private Sub Z_DbCrtSchm()
'Dim Fb$
'Fb = TmpFb
'DbCrtSchm FbCrt(Fb), ZZSchmy
'Kill Fb
'End Sub
'
'Private Function XEr(A As Brk) As String()
'Dim EAy$(), A1$(), A2$(), A3$(), A4$(), A5$(), A6$()
'With A.Dta
'    EAy = AyT1Ay(.E)
'    A1 = ErDupE(EAy)
'    A2 = ErDupF(.Tny, .T)
'    A3 = ErDupT(.Tny)
'    A4 = ErE(EAy)
'    A5 = ErFldHasNoEle(.Tny, .T, .F)
'    A6 = ErFEle_NotIn_EAy(.F, EAy)
'End With
'XEr = AyAddAp(A1, A2, A3, A4, A5, A6, A.Er)
'End Function
'Function ErDupT(Tny$()) As String()
'ErDupT = AyDupChk(Tny, "These T[?] is duplicated in TFld-lines")
'End Function
'
'Function ErDupE(EAy$()) As String()
'ErDupE = AyDupChk(EAy, "These Ele[?] are duplicated in Ele-lines")
'End Function
'
'Function ErDupF(Tny$(), TF$()) As String()
'Dim T
'For Each T In AyNz(Tny)
'    PushAy ErDupF, AyDupChk(TLyT_Fny(TF, T), FmtQQ("These F[?] are duplicated in T[?]", "?", T))
'Next
'End Function
'
'Function ELinChk(ByVal A$) As String()
'ELinChk = SclChk(TakAft(A, ";"), VdtEleSclNmSsl)
'End Function
'
'Function ErE(ELy$()) As String()
'ErE = AyOfAy_Ay(AyMap(ELy, "ELinChk"))
'End Function
'
'Function ErFldHasNoEle(Tny$(), TLy$(), FLy$()) As String()
'Dim T, F, E$
'For Each T In AyNz(Tny)
'    For Each F In AyNz(TLyT_Fny(TLy, T))
'        If T = F Then GoTo Nxt
'        If AyHas(Tny, F) Then GoTo Nxt
'        E = FLyTF_E(FLy, T, F)
'        If E = "" Then
'            Push ErFldHasNoEle, FmtQQ("T[?] F[?] cannot be found in any EF-lines", T, F)
'        End If
'Nxt:
'    Next
'Next
'End Function
'
'Function FLyEAy(FLy$()) As String()
'FLyEAy = AyT1Ay(FLy)
'End Function
'
'Function Er_F_of_ELy_NotIn_EAy(FLy$(), EAy$()) As String()
'Dim E$(), Er$()
'E = FLyEAy(FLy)
'Er = AyMinus(E, EAy)
'If Sz(Er) = 0 Then Exit Function
'ErFEle_NotIn_EAy = MsgLy("These [Ele] in F-lines are not found in the elements of E-lines", JnSpc(Er))
'End Function
'
'Function ErNoTF(T$()) As String()
'ErNoTF = AyEmpChk(T, "No TFld lines")
'End Function
'
'Sub DbCrtSchm(A As Database, Schmy$())
'With XRslt(Schmy)
'    AyBrwThw .Er
'    AyDoPX .Td, "DbAppTd", A
'    AyDoPX .Pk, "DbRun", A
'    AyDoPX .Sk, "DbRun", A
'    AyDoPX .TDes, "DbSetTDes", A
'    AyDoPX .FDes, "DbSetFDes", A
'End With
'End Sub
'
'Private Function XBrk(Schmy$()) As Brk
'Dim Ny$()
'Ny = Sy(C_T, C_F, C_E, C_D)
'With XBrk.Dta
'    AyAsg ClnBrk1(LyCln(Schmy), Ny), .T, .F, .E, .D, XBrk.Er
'    .Tny = AyT1Ay(.T)
'End With
'End Function
'
'Private Function ZZCln() As String()
'ZZCln = LyCln(ZZSchmy)
'End Function
'
'Private Function ZZDLy() As String()
'ZZDLy = AyWhRmvT1(ZZCln, "D")
'End Function
'
'Private Function XRslt(Schmy$()) As Rslt
'Dim Brk As Brk
'    Brk = XBrk(Schmy)
'    Dim Er$()
'    Er = XEr(Brk)
'    If Sz(Er) > 0 Then
'        Dim UL$()
'        UL = ApSy("--------------------------------")
'        XRslt.Er = AyAddAp(Schmy, UL, Er)
'        Exit Function
'    End If
'Dim O As Rslt, XDes As Des
'With Brk.Dta
'    XDes = DLyDes(.D)
'    O.Td = DtaTdAy(Brk.Dta)
'    O.Sk = TLySkSqy(.T)
'    O.Pk = TLyPkSqy(.T)
'    O.TDes = XDes.T
'    O.FDes = XDes.F
'End With
'XRslt = O
'End Function
'Sub Z_DLyDes()
'Dim X As Des
'X = DLyDes(ZZDLy)
'Stop
'End Sub
'Function SchmyEr(A$()) As String()
'SchmyEr = XEr(XBrk(A))
'End Function
'
'Function CvTDes(A) As TDes
'Set CvTDes = A
'End Function
'
'Function CvFDes(A) As FDes
'Set CvFDes = A
'End Function
'
'Sub PushTDes(O() As TDes, T, Des$)
'Dim I
'If Sz(O) > 0 Then
'    For Each I In O
'        With CvTDes(I)
'            If .T = T Then
'                .Des = Des & vbCrLf & Des
'                Exit Sub
'            End If
'        End With
'    Next
'End If
'
'Dim M As New TDes
'M.Des = Des
'M.T = T
'Push O, M
'End Sub
'
'Sub PushFDes(O() As FDes, T, F, Des$)
'Dim I
'If Sz(O) > 0 Then
'    For Each I In O
'        With CvFDes(I)
'            If .T = T Then
'                If .F = F Then
'                    .Des = Des & vbCrLf & Des
'                    Exit Sub
'                End If
'            End If
'        End With
'    Next
'End If
'
'Dim M As New FDes
'M.Des = Des
'M.T = T
'M.F = F
'Push O, M
'End Sub
'
'Function DLyDes(A$()) As Des
'Dim OT() As TDes
'Dim OF() As FDes
'Dim DLin, T$, Des$, F$, V$
'If Sz(A) = 0 Then Exit Function
'For Each DLin In A
'    AyAsg Lin3TAy(DLin), T, F, V, Des
'    If V <> "|" Then Stop
'    If F = "." Then
'        PushTDes OT, T, Des
'    Else
'        PushFDes OF, T, F, Des
'    End If
'Next
'DLyDes.T = OT
'DLyDes.F = OF
'End Function
'
'Function TLyT_Fny(A$(), T) As String()
'Dim B$
'B = AyFstT1(A, T)
'If LinShiftT1(B) <> T Then Stop
'B = Replace(B, "*", T)
'TLyT_Fny = AyRmvEle(SslSy(B), "|")
'End Function
'
'Private Function DtaT_FdAy(A As Dta, T) As DAO.Field2()
'Dim O() As DAO.Field2, F
'With A
'    For Each F In TLyT_Fny(.T, T)
'        PushObj O, Fd(F, T, .Tny, .F, .E)
'    Next
'End With
'DtaT_FdAy = O
'End Function
'
'Sub Z()
'Z_DbCrtSchm
'Z_DLyDes
'Z_FLyTF_FLin
'Z_Tny
'End Sub