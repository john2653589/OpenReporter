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
        public IOpenReport SaveAs(string ExportFileName, string ExportPath)
        {
            if (!ExportFileName.ToLower().Contains(".ods"))
                ExportFileName += ".ods";

            var FullFileName = Path.Combine(ExportPath, ExportFileName);
            SaveAs(FullFileName);
            return this;
        }
        public IOpenReport SaveAs(string SaveFullFileName)
        {
            var Info = new FileInfo(SaveFullFileName);
            if (!Info.Directory.Exists)
                Info.Directory.Create();

            using var WriteMemory = new MemoryStream();
            ContentXml.Save(WriteMemory);
            WriteMemory.Seek(0, SeekOrigin.Begin);
            OdsZip.UpdateEntry("content.xml", WriteMemory);

            if (!SaveFullFileName.ToLower().Contains(".ods"))
                SaveFullFileName += ".ods";
            OdsZip.Save(SaveFullFileName);
            return this;
        }
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
    }
}