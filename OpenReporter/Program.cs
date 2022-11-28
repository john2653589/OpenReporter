using Rugal.Net.OpenReporter.Extention;
using Rugal.Net.OpenReporter.Ods.Core;
using Rugal.Net.OpenReporter.Ods.Extention;

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
    .AsSheet("工作表1");

var ImagePath = @"C:\Crims\DOG.jpg";
Sheet.AsCell("A1", Item =>
{
    var ImageBuffer = File.ReadAllBytes(ImagePath);
    var Base64 = Convert.ToBase64String(ImageBuffer);

    var Cell = Item as OdsCell;
    Cell.CellNode.SetImage(Base64);

});
var Buffer = Sheet
    .Report
    .SaveAsClose(ExportFileName, ExportPath)
    ;

var ss = 1;

class TestModel
{
    public string Key { get; set; }
    public int Value { get; set; }
}