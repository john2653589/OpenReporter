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
var TemplateFileName = @"c:\Crims\test.ods";
var ExportFileName = "save";

var Sheet = new OdsReport()
    .ReadFile(TemplateFileName)
    .AsSheet("工作表2");

Sheet
    .PositionRow(2)
    //設定標題
    .ForEachFrom(TestData, (Item, Row) =>
    {
        var Cell = Row["A"];
        Row["A"].SetValue(Item.Key);
        Row["B"].SetValue(Item.Value);
        Sheet.InsertRowAfterFromClear();
    })
    .Report
    .SaveAs(ExportFileName, ExportPath)
    ;

class TestModel
{
    public string Key { get; set; }
    public int Value { get; set; }
}