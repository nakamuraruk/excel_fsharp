#r "Microsoft.Office.Interop.Excel"
(* Reference: https://msdn.microsoft.com/EN-US/library/microsoft.office.interop.excel.aspx *)
open System
open System.IO
open System.Reflection
open Microsoft.Office.Interop.Excel

let wbs :Workbooks = ApplicationClass(Visible=true).Workbooks
let wb: Workbook = wbs.Open("foobar.xlsx")
let sheet : Worksheet = wb.Worksheets.["summary"]  :?> Worksheet

let charts : ChartObjects = sheet.ChartObjects() :?> ChartObjects
let chartobject: ChartObject = charts.Add(400.0, 20.0, 550.0, 350.0)
let chart = chartobject.Chart :> _Chart
let a:Series  = (chart.SeriesCollection() :?> SeriesCollection).NewSeries()
a.XValues <- "=tl14sv003_tab10k_thr1!$B$8:$B$10007":> obj
a.Values <- "=tl14sv003_tab10k_thr1!$E$8:$E$10007":> obj

let b: LegendEntry  = chart.Legend.LegendEntries(1) :?> LegendEntry
b.Delete()
//a.ChartType <- XlChartType.xlXYScatterLines

 //   println "%A" bb
// wb.SaveAs("abc.xlsx")
