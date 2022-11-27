using Rugal.Net.OpenReporter.Extention;
using Rugal.Net.OpenReporter.Ods.Core;

var TestData = new List<TestModel>
{
    new TestModel()
    {
        Key = "第一區處",
        Value = 111
    },
    new TestModel()
    {
        Key = "第二區處",
        Value = 222
    },
    new TestModel()
    {
        Key = "第三區處",
        Value = 33
    },
    new TestModel()
    {
        Key = "第四區處",
        Value = 7777
    },
};

var ExportPath = @"C:\crims2";
var TemplateFileName = @"c:\Crims\OdsTemplate_ReportConstruction_Area_Day.ods";
var ExportFileName = "save";

var Sheet = new OdsReport()
    .ReadFile(TemplateFileName)
    .AsSheet("台灣自來水公司各單位每日報工統計表");

Sheet["L2"].SetValue("哈哈哈");
var Buffer = Sheet
    .PositionRow(4)
    //設定標題
    .ForEachFrom_AutoInsert(TestData, (Item, Row) =>
    {
        var Cell = Row["A"];
        Row["A"].SetValue(Item.Key);
        Row["B"].SetValue(Item.Value);
        Row["C"].SetValue(Item.Value);
        Row["D"].SetValue(Item.Value);
        Row["E"].SetValue(Item.Value);
        Row["F"].SetValue(Item.Value);
        Row["G"].SetValue(Item.Value);
        Row["H"].SetValue(Item.Value);
        Row["I"].SetValue(Item.Value);
        Row["J"].SetValue(Item.Value);
        Row["K"].SetValue(Item.Value);
        Row["L"].SetValue(Item.Value);
    })
    .Report
    //.SaveAsAndReadDelete(ExportFileName, ExportPath)
    .SaveAsClose(ExportFileName, ExportPath)
    ;

var ss = 1;

class TestModel
{
    public string Key { get; set; }
    public int Value { get; set; }
}