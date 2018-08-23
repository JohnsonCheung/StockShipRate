Option Compare Database

Sub TmpRate_Upd_YM(A$)
Dim Avg#, Max#, Min#, NRec%
With WqRs(FmtQQ("Select Avg(RateSc) as [Avg], Max(RateSc) as [Max], Min(RateSc) as [Min], Count(*) as NRec from [?]", A))
    Avg = !Avg
    Max = !Max
    Min = !Min
    NRec = !NRec
    .Close
End With
With QQSqlRs("Select RateSc_Avg,RateSc_Max,RateSc_Min,RateSc_NRec,RateSc_LoadDte from [YM] where Y=? and M=?", Y, M)
    .Edit
    !RateSc_Avg = Avg
    !RateSc_Max = Max
    !RateSc_Min = Min
    !RateSc_NRec = NRec
    !RateSc_LoadDte = Now
    .Update
    .Close
End With
End Sub