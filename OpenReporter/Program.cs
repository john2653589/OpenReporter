using Rugal.Net.OpenReporter.Extention;
using Rugal.Net.OpenReporter.Interface;
using Rugal.Net.OpenReporter.Ods.Core;

var FullFileName = @"c:\Crims\test.ods";

var Sheet = new OdsReport()
    .ReadFile(FullFileName)
    .FindSheet("工作表1")
    .AsOpenSheet();

Sheet
    .AsCell("D1", Cell => Cell.SetValue(DateTime.Now.ToString("yyyy-MM-dd")))
    .For(() =>
    {

        Sheet["A1"].SetValue("標題");
        Sheet["A3"].SetValue("跨Row");
        Sheet["A5"].SetValue("跨Row資料");
    })
    .Reporter
    .Save()
    ;

Sheet.Reporter.Save();
