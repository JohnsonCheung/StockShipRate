Option Compare Database
Option Explicit
Private Sub Main()
SetColr_ToDo
PtFmt
HowToEnsFirstTime_FmtSpec
Addition_API_forMsgLy
End Sub

Private Sub PtFmt()

End Sub

Private Sub Addition_API_forMsgLy()
'eg.  call FmtSpec_Imp, it will show detail of message, but there is one line of [xx] .. [xx] only.
'     try to not display this line.
End Sub

Private Sub HowToEnsFirstTime_FmtSpec()
'No table-Spec
'No rec-Fmt
End Sub

Private Sub SetColr_ToDo()
'TstStep
'   Call Gen
'   Call FmtSpec_Brw 'Edt
'       Edit and Save, then Call Gen will auto import
'where to add autoImp?
'   ?
'AutoImp will show msg if import/noImport
'ColrLy
'   what is the common color name in DotNet Library
'       Use Enums: System.Drawing.KnownColor is no good, because the EnmNm is in seq, it is not return
'       Use VBA.ColorConstants-module is good, but there is few constant
'       Answer: Use *KnownColor to feed in struct-*Color, there is *Color.ToArgb & *KnownColor has name
'               Run the FSharp program.
'               Put the generated file
'                   in
'                       C:\Users\user\Source\Repos\EnumLines\EnumLines\bin\Debug\ColorLines.Const.Txt
'                   Into
'                       C:\Users\user\Desktop\MHD\SAPAccessReports\StockShipRate\StockShipRate\Spec
'               Run ConstGen: It will addd the Public Const ColorLines = ".... at end
'               Put Fct-Module
'To find some common values to feed into ColrLines
'
'Colr* 4-functions
'    ColrStr_MayColr
'    ColrStr
'    ColrLy
'    ColrLines
End Sub

Sub FSharpBuildKnownColor()
'// Learn more about F# at http://fsharp.org
'// See the 'F# Tutorial' project for more help.
'open System.Drawing
'open System
'open System.IO
'open System.Windows.Forms
'
'type slis = String list
'type sy = String[]
'type sseq = String seq
'let slis_lines(a:slis) = String.Join("\r\n",a)
'let sy_lines(a:sy) = String.Join("\r\n",a)
'let str_wrt ft a = File.WriteAllText(ft,a)
'let sseq_wrt ft (a:sseq) = File.WriteAllLines(ft,a)
'let slis_wrt ft a = a|> sseq_wrt ft
'let mayStr_wrt a ft = match a with | Some a -> str_wrt a ft | _ -> ()
'Let colorConstFt = "ColorLines.Txt"
'//let knownColor_lin a = a.ToString() + " " + Color.FromKnownColor(a).ToArgb().ToString()
'let knownColor_lin a = "Public Const " + a.ToString() + "& = " + Color.FromKnownColor(a).ToArgb().ToString()
'let sy_wrt a ft = a |> sseq_wrt ft
'let arr_seq<'a>(a:Array) = seq { for i in a -> unbox i }
'let arr_ay<'a>(a:Array) = [|for i in a -> unbox i|]
'let arr_lis<'a>(a:Array) = [for i in a -> unbox i]
'let knownColorArr = Enum.GetValues(KnownColor.ActiveBorder.GetType())
'let knownColorLis = knownColorArr |> arr_lis<KnownColor>
'let colorConstLis = knownColorLis |> List.map knownColor_lin |> List.sort
'let wrt_colorConstFt() = slis_wrt colorConstFt colorConstLis
'[<EntryPoint>]
'let main argv =
'    printfn "%A" argv
'//    MessageBox.Show System.Environment.CurrentDirectory |> ignore
'    do wrt_colorConstFt()
'    0 // return an integer exit code
End Sub