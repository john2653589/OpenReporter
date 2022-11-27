using Ionic.Zip;
using Rugal.Net.OpenReporter.Interface;
using Rugal.Net.OpenReporter.Ods.Extention;
using System.Data;
using System.Globalization;
using System.Xml;

namespace Rugal.Net.OpenReporter.Ods.Core
{
    /// <summary>
    /// OdsReporter
    /// Version 1.0.0
    /// From Rugal Tu
    /// </summary>
    public class OdsReport : IOpenReport
    {
        #region Node Property
        internal XmlNodeList SheetNodes => ContentXml.NodeList_Sheet(NamespaceManager);
        #endregion

        #region Internal Property
        internal ZipFile OdsZip { get; set; }
        internal XmlDocument ContentXml { get; set; }
        internal XmlNamespaceManager NamespaceManager { get; set; }
        #endregion

        #region Public Property

        #endregion

        #region Interface Method
        public IOpenReport ReadFile(string FullFileName)
        {
            InitZipFile(FullFileName);
            InitContentXml();
            InitNamespaceManager();
            return this;
        }
        public IOpenSheet FindSheet(string SheetName)
        {
            var SheetNode = SheetNodes
                .First(Item => Item.Attr_TableName().Value == SheetName);

            var Sheet = new OdsSheet(this, SheetNode);
            return Sheet;
        }
        public IOpenReport Save()
        {
            using var WriteMemory = new MemoryStream();
            ContentXml.Save(WriteMemory);
            WriteMemory.Seek(0, SeekOrigin.Begin);
            OdsZip.UpdateEntry("content.xml", WriteMemory);
            OdsZip.Save();
            return this;
        }
        #endregion

        #region As Method
        public IOpenSheet AsSheet(string SheetName) => FindSheet(SheetName);
        #endregion

        #region Get Set Process
        private XmlNode GetSheetRootAndRemoveChildren()
        {
            var RootNode = SheetNodes.Item(0).ParentNode;
            foreach (XmlNode Node in SheetNodes)
                RootNode.RemoveChild(Node);

            return RootNode;
        }
        #endregion

        #region Init Process
        private void InitZipFile(Stream Stream)
        {
            OdsZip = ZipFile.Read(Stream);
        }
        private void InitZipFile(string FullFileName)
        {
            OdsZip = ZipFile.Read(FullFileName);
        }
        private void InitNamespaceManager()
        {
            NamespaceManager = new XmlNamespaceManager(ContentXml.NameTable);
            foreach (var Item in OdsProperty.OdsNamespaces)
                NamespaceManager.AddNamespace(Item.Key, Item.Value);
        }
        private void InitContentXml()
        {
            var ContentZip = OdsZip["content.xml"];
            using var ContentStream = new MemoryStream();
            ContentZip.Extract(ContentStream);
            ContentStream.Seek(0, SeekOrigin.Begin);

            ContentXml = new XmlDocument();
            ContentXml.Load(ContentStream);
        }
        #endregion

        #region Save Process
        private void SaveSheet(DataTable Sheet, XmlNode RootNode)
        {
            var OwnerDocument = RootNode.OwnerDocument;

            var CreateSheetName = OwnerDocument.CreateAttribute("table:name", OdsProperty.TableUri);
            CreateSheetName.Value = Sheet.TableName;

            var CreateSheet = OwnerDocument.CreateElement("table:table", OdsProperty.TableUri);
            CreateSheet.Attributes.Append(CreateSheetName);

            SaveColumnDefinition(Sheet, CreateSheet, OwnerDocument);
            SaveRows(Sheet, CreateSheet, OwnerDocument);

            RootNode.AppendChild(CreateSheet);
        }
        private void SaveColumnDefinition(DataTable sheet, XmlNode SheetNode, XmlDocument OwnerDocument)
        {
            var CreateCellRepeated = OwnerDocument.CreateAttribute(OdsProperty.PATH_CellRepeated, OdsProperty.TableUri);
            CreateCellRepeated.Value = sheet.Columns.Count.ToString(CultureInfo.InvariantCulture);

            var CreateColumnDefinition = OwnerDocument.CreateElement("table:table-column", OdsProperty.TableUri);
            CreateColumnDefinition.Attributes.Append(CreateCellRepeated);

            SheetNode.AppendChild(CreateColumnDefinition);
        }
        private void SaveRows(DataTable Sheet, XmlNode SheetNode, XmlDocument OwnerDocument)
        {
            foreach (DataRow Row in Sheet.Rows)
            {
                var RowNode = OwnerDocument.CreateElement("table:table-row", OdsProperty.TableUri);

                SaveCell(Row, RowNode, OwnerDocument);

                SheetNode.AppendChild(RowNode);
            }
        }
        private void SaveCell(DataRow Row, XmlNode RowNode, XmlDocument OwnerDocument)
        {
            foreach (var Cells in Row.ItemArray)
            {
                var CreateCellNode = OwnerDocument.CreateElement("table:table-cell", OdsProperty.TableUri);

                if (Cells != DBNull.Value)
                {
                    var CreateValueType = OwnerDocument.CreateAttribute("office:value-type", OdsProperty.TableUri);
                    CreateValueType.Value = "string";
                    CreateCellNode.Attributes.Append(CreateValueType);

                    var CreateCellValue = OwnerDocument.CreateElement("text:p", OdsProperty.TextUri);
                    CreateCellValue.InnerText = Cells.ToString();
                    CreateCellNode.AppendChild(CreateCellValue);
                }
                RowNode.AppendChild(CreateCellNode);
            }
        }

      
        #endregion

        private DataTable GetSheet(XmlNode tableNode)
        {
            DataTable sheet = new DataTable(tableNode.Attributes["table:name"].Value);

            XmlNodeList rowNodes = tableNode.NodeList_Row(NamespaceManager);

            int rowIndex = 0;
            foreach (XmlNode rowNode in rowNodes)
                GetRow(rowNode, sheet, NamespaceManager, ref rowIndex);

            return sheet;
        }
        private void GetRow(XmlNode RowNode, DataTable sheet, XmlNamespaceManager nmsManager, ref int rowIndex)
        {
            var RowRepeated = RowNode.Attr_RowRepeated();
            if (RowRepeated == null || Convert.ToInt32(RowRepeated.Value, CultureInfo.InvariantCulture) == 1)
            {
                while (sheet.Rows.Count < rowIndex)
                    sheet.Rows.Add(sheet.NewRow());

                DataRow row = sheet.NewRow();

                var CellNodes = RowNode.NodeList_Cell(NamespaceManager);

                int cellIndex = 0;
                foreach (XmlNode cellNode in CellNodes)
                    GetCell(cellNode, row, ref cellIndex);

                sheet.Rows.Add(row);

                rowIndex++;
            }
            else
            {
                rowIndex += Convert.ToInt32(RowRepeated.Value, CultureInfo.InvariantCulture);
            }

            // sheet must have at least one cell
            if (sheet.Rows.Count == 0)
            {
                sheet.Rows.Add(sheet.NewRow());
                sheet.Columns.Add();
            }
        }

        private void GetCell(XmlNode CellNode, DataRow row, ref int cellIndex)
        {
            var CellRepeated = CellNode.Attr_CellRepeated();

            if (CellRepeated == null)
            {
                DataTable sheet = row.Table;

                while (sheet.Columns.Count <= cellIndex)
                    sheet.Columns.Add();

                row[cellIndex] = CellNode.GetCellValue();

                cellIndex++;
            }
            else
            {
                cellIndex += Convert.ToInt32(CellRepeated.Value, CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        /// Writes DataSet as .ods file.
        /// </summary>
        /// <param name="odsFile">DataSet that represent .ods file.</param>
        /// <param name="outputFilePath">The name of the file to save to.</param>
        public void WriteOdsFile(DataSet odsFile, string outputFilePath)
        {
            //ZipFile templateFile = this.GetZipFile(Assembly.GetExecutingAssembly().GetManifestResourceStream("OdsReadWrite.template.ods"));
            //ZipFile templateFile = this.GetZipFile(File.OpenRead(@"C:\CODE\ODSExport\ODSExport\bin\template.ods"));
            //ZipFile templateFile = this.GetZipFile(File.OpenRead(Server.MapPath("/") + "template.ods");
            //ZipFile templateFile = GetZipFile(File.OpenRead(HttpContext.Current.Server.MapPath("/template.ods")));

            //var SheetRoot = GetSheetRootAndRemoveChildren();
            //foreach (DataTable sheet in odsFile.Tables)
            //    SaveSheet(sheet, SheetRoot);

            //SaveContentXml(templateFile, contentXml);

            //templateFile.Save(outputFilePath);
        }
    }
}