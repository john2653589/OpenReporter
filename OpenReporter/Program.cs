using DocumentFormat.OpenXml.Packaging;
using Rugal.Net.OpenReporter.Extention;
using Rugal.Net.OpenReporter.Ods.Core;
using Rugal.Net.OpenReporter.Ods.Extention;
using Rugal.OpenExcel.Core;
using Utilities;

var ExportPath = @"C:\crims2";
//var TemplateFileName = @"c:\Crims\test.ods";
var TemplateFileName = @"c:\Crims\test.xlsx";
var ExportFileName = "save";
var ImagePath = @"C:\Crims\test2.jpg";

//var Excel = new OpenExcelModel(TemplateFileName);
//var Sheet = Excel["工作表1"];
//var Reader = Sheet.Reader(Excel);

//var ImagePart = Sheet.WorksheetPart.AddImagePart(ImagePartType.Png);

//var Buffer = File.ReadAllBytes(ImagePath);
//var Memory = new MemoryStream(Buffer);

//ExcelTools.AddImage(Sheet.WorksheetPart, Buffer, "aa", 3, 3);

//Excel.SaveAsXlsxAndExit(ExportFileName, ExportPath);

var Sheet = new OdsReport()
    .ReadFile(TemplateFileName)
    .AsSheet("工作表1");

Sheet.AsCell("A1", Item =>
{
    var ImageBuffer = File.ReadAllBytes(ImagePath);
    var Base64 = Convert.ToBase64String(ImageBuffer);

    var Cell = Item as OdsCell;
    Cell.CellNode.SetImage(Base64, "60pt", "60pt");

});
var Buffer = Sheet
    .Report
    .SaveAsClose(ExportFileName, ExportPath)
    ;
