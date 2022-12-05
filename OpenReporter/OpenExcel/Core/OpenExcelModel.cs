using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Rugal.OpenExcel.Core
{
    /// <summary>
    /// OpenExcel
    /// Version 1.0.0
    /// From Rugal Tu
    /// </summary>
    public class OpenExcelModel : IDisposable
    {
        private string CurrentPath => Environment.CurrentDirectory;
        public string FullFileName { get; set; }
        public bool IsFromTemplate { get; set; }
        public bool IsFromStream { get; set; }
        public Stream MemoryStream { get; set; }
        public SpreadsheetDocument Document { get; set; }

        #region Lambda Property
        public WorkbookPart WorkbookPart => Document.WorkbookPart;
        public Workbook Workbook => WorkbookPart.Workbook;
        public IEnumerable<Sheet> SheetList => Workbook.Descendants<Sheet>();
        #endregion

        #region Lambda Function
        public Sheet GetSheetInfo(string SheetName) =>
            SheetList.FirstOrDefault(Item => Item.Id == GetSheetId(SheetName));
        public string GetSheetId(string SheetName) =>
            SheetList.FirstOrDefault(Item => Item.Name == SheetName)?.Id;
        public WorksheetPart GetSheetPart(string Id) =>
            WorkbookPart.GetPartById(Id) as WorksheetPart;
        public OpenExcelReader Reader(string SheetName, (int Row, int Col) EndIdx, (int Row, int Col) StartIdx = default) => this[SheetName].Reader(this, EndIdx, StartIdx);
        #endregion
        public Worksheet this[string SheetName] => GetSheetPart(GetSheetId(SheetName)).Worksheet;
        public OpenExcelModel(string FileName, string AssignPath, bool _IsFromTemplate = true)
        {
            AssignPath = string.IsNullOrWhiteSpace(AssignPath) ? CurrentPath : AssignPath;
            FullFileName = Path.Combine(AssignPath, FileName);
            IsFromTemplate = _IsFromTemplate;
            Document = Open();
        }
        public OpenExcelModel(string _FullFileName, bool _IsFromTemplate = true)
        {
            FullFileName = _FullFileName;
            IsFromTemplate = _IsFromTemplate;
            Document = Open();
        }
        public OpenExcelModel(Stream _MemoryStream)
        {
            MemoryStream = _MemoryStream;
            IsFromStream = true;
            Document = OpenWithStream();
        }
        public SpreadsheetDocument Open()
        {
            try
            {
                var Ret = IsFromTemplate ?
                    SpreadsheetDocument.CreateFromTemplate(FullFileName) :
                    SpreadsheetDocument.Open(FullFileName, true);
                return Ret;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }
        public SpreadsheetDocument OpenWithStream()
        {
            try
            {
                var Ret = SpreadsheetDocument.Open(MemoryStream, false);
                return Ret;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }
        public string SaveAsXlsx(string FileName, string AssignPath = null)
        {
            if (!FileName.ToLower().Contains(".xls"))
                FileName += ".xlsx";

            AssignPath = string.IsNullOrWhiteSpace(AssignPath) ? CurrentPath : AssignPath;
            var SaveFileName = Path.Combine(AssignPath, FileName);
            if (!Directory.Exists(AssignPath))
                Directory.CreateDirectory(AssignPath);

            var SaveDoc = Document.SaveAs(SaveFileName);
            SaveDoc.Close();
            SaveDoc.Dispose();
            return SaveFileName;
        }
        public string SaveAsXlsxAndExit(string FileName, string AssignPath = null)
        {
            var SaveFileName = SaveAsXlsx(FileName, AssignPath);
            ExitExcel();
            return SaveFileName;
        }
        public void ExitExcel()
        {
            Document?.Close();
            Document?.Dispose();
            Document = null;
            MemoryStream?.Flush();
            MemoryStream?.Close();
            MemoryStream?.Dispose();
            MemoryStream = null;
        }
        public void DeleteFile(string DeleteFileName)
        {
            if (File.Exists(DeleteFileName))
                File.Delete(DeleteFileName);
        }
        public Worksheet AddSheet(string SheetName)
        {
            var NewSheetPart = Document.WorkbookPart.AddNewPart<WorksheetPart>();
            NewSheetPart.Worksheet = new Worksheet(new SheetData());

            var SheetId = (uint)1;
            if (Workbook.Sheets.Any())
                SheetId = Workbook.Sheets
                    .Elements<Sheet>()
                    .Select(Item => Item.SheetId.Value)
                    .Max() + 1;

            var Id = WorkbookPart.GetIdOfPart(NewSheetPart);

            var NewSheet = new Sheet()
            {
                Name = SheetName,
                SheetId = SheetId,
                Id = Id
            };
            Workbook.Sheets.Append(NewSheet);

            return GetSheetPart(Id).Worksheet;
        }
        public void DeleteSheet(string SheetName)
        {
            var Id = GetSheetId(SheetName);
            var Sheet = SheetList.FirstOrDefault(Item => Item.Id == Id);
            WorkbookPart.DeletePart(Id);
            Sheet.Remove();
        }
        public void SetCalc(bool IsHasCalcCell)
        {
            var CalcProperty = Workbook.CalculationProperties;
            CalcProperty.ForceFullCalculation = true;
            CalcProperty.FullCalculationOnLoad = true;
        }

        #region Dispose
        private bool IsDispose;
        public void Dispose()
        {
            ExitExcel();
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool Disposing)
        {
            if (IsDispose)
                return;
            IsDispose = true;
        }
        #endregion  
    }
}
