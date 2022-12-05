using System;
using System.Data;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Jpeg;

namespace Rugal.OpenExcel.Core
{
    public class ExcelTools
    {
        public static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
        {
            IEnumerable<Sheet> sheets =
               document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
               Elements<Sheet>().Where(s => s.Name == sheetName);

            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist
                return null;
            }

            string relationshipId = sheets.First().Id.Value;
            return (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
        }

        //public static void AddImage(bool createFile, string excelFile, string sheetName,
        //                            string imageFileName, string imgDesc,
        //                            int colNumber, int rowNumber)
        //{
        //    using (var imageStream = new FileStream(imageFileName, FileMode.Open))
        //    {
        //        AddImage(createFile, excelFile, sheetName, imageStream, imgDesc, colNumber, rowNumber);
        //    }
        //}

        //public static void AddImage(WorksheetPart worksheetPart,
        //                            string imageFileName, string imgDesc,
        //                            int colNumber, int rowNumber)
        //{
        //    using (var imageStream = new FileStream(imageFileName, FileMode.Open))
        //    {
        //        AddImage(worksheetPart, imageStream, imgDesc, colNumber, rowNumber);
        //    }
        //}

        //public static void AddImage(bool createFile, string excelFile, string sheetName,
        //                            Stream imageStream, string imgDesc,
        //                            int colNumber, int rowNumber)
        //{
        //    SpreadsheetDocument spreadsheetDocument = null;
        //    WorksheetPart worksheetPart = null;
        //    if (createFile)
        //    {
        //        // Create a spreadsheet document by supplying the filepath
        //        spreadsheetDocument = SpreadsheetDocument.Create(excelFile, SpreadsheetDocumentType.Workbook);

        //        // Add a WorkbookPart to the document
        //        WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
        //        workbookpart.Workbook = new Workbook();

        //        // Add a WorksheetPart to the WorkbookPart
        //        worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
        //        worksheetPart.Worksheet = new Worksheet(new SheetData());

        //        // Add Sheets to the Workbook
        //        Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
        //            AppendChild<Sheets>(new Sheets());

        //        // Append a new worksheet and associate it with the workbook
        //        Sheet sheet = new Sheet()
        //        {
        //            Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
        //            SheetId = 1,
        //            Name = sheetName
        //        };
        //        sheets.Append(sheet);
        //    }
        //    else
        //    {
        //        // Open spreadsheet
        //        spreadsheetDocument = SpreadsheetDocument.Open(excelFile, true);

        //        // Get WorksheetPart
        //        worksheetPart = GetWorksheetPartByName(spreadsheetDocument, sheetName);
        //    }

        //    AddImage(worksheetPart, imageStream, imgDesc, colNumber, rowNumber);

        //    worksheetPart.Worksheet.Save();

        //    spreadsheetDocument.Close();
        //}

        public static void AddImage(WorksheetPart worksheetPart, byte[] Buffer,
            int RowIndex, int ColumnIndex, int SetWidth = 150, int SetHeight = 100, string ImageDesc = null)
        {
            // We need the image stream more than once, thus we create a memory copy


            var drawingsPart = worksheetPart.DrawingsPart;
            if (drawingsPart == null)
                drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

            if (!worksheetPart.Worksheet.ChildElements.OfType<Drawing>().Any())
            {
                worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            if (drawingsPart.WorksheetDrawing == null)
            {
                drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
            }

            var worksheetDrawing = drawingsPart.WorksheetDrawing;

            var imagePart = drawingsPart.AddImagePart(ImagePartType.Jpeg);
            imagePart.FeedData(new MemoryStream(Buffer));

            A.Extents extents = new A.Extents();
            var bm = Image.Load(Buffer, out var Format);
            var extentsCx = (long)(SetWidth * 914400 / 96);
            var extentsCy = (long)(SetHeight * 914400 / 96);

            var colOffset = 0;
            var rowOffset = 0;

            var nvps = worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>();
            var nvpId = nvps.Count() > 0
                ? (UInt32Value)worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>().Max(p => p.Id.Value) + 1
                : 1U;

            var oneCellAnchor = new Xdr.OneCellAnchor(
                new Xdr.FromMarker
                {
                    ColumnId = new Xdr.ColumnId((ColumnIndex - 1).ToString()),
                    RowId = new Xdr.RowId((RowIndex - 1).ToString()),
                    ColumnOffset = new Xdr.ColumnOffset(colOffset.ToString()),
                    RowOffset = new Xdr.RowOffset(rowOffset.ToString()),
                },
                new Xdr.Extent { Cx = extentsCx, Cy = extentsCy },
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = "Picture " + nvpId, Description = ImageDesc },
                        new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })
                    ),
                    new Xdr.BlipFill(
                        new A.Blip { Embed = drawingsPart.GetIdOfPart(imagePart), CompressionState = A.BlipCompressionValues.Print },
                        new A.Stretch(new A.FillRectangle())
                    ),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0, Y = 0 },
                            new A.Extents { Cx = extentsCx, Cy = extentsCy }
                        ),
                        new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }
                    )
                ),
                new Xdr.ClientData()
            );

            worksheetDrawing.Append(oneCellAnchor);
        }
    }
}