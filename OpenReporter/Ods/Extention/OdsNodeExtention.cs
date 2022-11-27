using Rugal.Net.OpenReporter.Ods.Core;
using System.Xml;

namespace Rugal.Net.OpenReporter.Ods.Extention
{
    public static class XmlDocumentExtention
    {
        public static XmlNodeList NodeList_Sheet(this XmlDocument ContentXml, XmlNamespaceManager NamespaceManager)
        {
            var SheetNodes = ContentXml.SelectNodes(OdsProperty.PATH_SheetNodes, NamespaceManager);
            return SheetNodes;
        }
    }

    public static class XmlNodeExtention
    {
        #region Row Extenion
        public static XmlNodeList NodeList_Row(this XmlNode Sheet, XmlNamespaceManager NamespaceManager)
        {
            var RowNodes = Sheet.SelectNodes(OdsProperty.PATH_RowNodes, NamespaceManager);
            return RowNodes;
        }
        public static XmlAttribute Attr_RowRepeated(this XmlNode RowNode)
        {
            var RowAttr = RowNode.Attributes[OdsProperty.PATH_RowRepeated];
            return RowAttr;
        }
        public static bool IsRepeatedRow(this XmlNode RowNode, out int RepeatedNumber)
        {
            RepeatedNumber = 0;
            var GetAttr = RowNode.Attr_RowRepeated();
            if (GetAttr is null)
                return false;

            RepeatedNumber = int.Parse(GetAttr.Value);
            return true;
        }
        #endregion

        #region Cell Extention
        public static XmlNodeList NodeList_Cell(this XmlNode RowNode, XmlNamespaceManager NamespaceManager)
        {
            var CellNodes = RowNode.SelectNodes(OdsProperty.PATH_CellNodes, NamespaceManager);
            return CellNodes;
        }
        public static XmlAttribute Attr_Value(this XmlNode CellNode)
        {
            var CellValueArrt = CellNode.Attributes[OdsProperty.PATH_Value];
            return CellValueArrt;
        }
        public static XmlAttribute Attr_Value_Create(this XmlNode CellNode)
        {
            var CellValueArrt = CellNode.Attributes[OdsProperty.PATH_Value];
            if (CellValueArrt is null)
            {
                CellValueArrt = CellNode.CreateAttr_Value();
                CellNode.Attributes.Append(CellValueArrt);
            }
            return CellValueArrt;
        }
        public static XmlAttribute Attr_ValueType(this XmlNode CellNode)
        {
            var ValueTypeArrt = CellNode.Attributes[OdsProperty.PATH_ValueType];
            if (ValueTypeArrt is null)
            {
                ValueTypeArrt = CellNode.CreateAttr_ValueType();
                CellNode.Attributes.Append(ValueTypeArrt);
            }
            return ValueTypeArrt;
        }
        public static XmlAttribute Attr_TableName(this XmlNode SheetNode)
        {
            var CellValueArrt = SheetNode.Attributes[OdsProperty.PATH_TableName];
            return CellValueArrt;
        }
        public static XmlAttribute Attr_CellRepeated(this XmlNode CellNode)
        {
            var CellRepeatedAttr = CellNode.Attributes[OdsProperty.PATH_CellRepeated];
            return CellRepeatedAttr;
        }
        public static XmlAttribute Attr_CellSpanned(this XmlNode CellNode)
        {
            var CellSpannedAttr = CellNode.Attributes[OdsProperty.PATH_CellSpanned];
            return CellSpannedAttr;
        }

        public static bool IsRepeatedCell(this XmlNode CellNode, out int RepeatedNumber)
        {
            RepeatedNumber = 0;
            var GetAttr = CellNode.Attr_CellRepeated();
            if (GetAttr is null)
                return false;

            RepeatedNumber = int.Parse(GetAttr.Value);
            return true;
        }
        public static bool IsSpannedCell(this XmlNode CellNode, out int SpannedNumber)
        {
            SpannedNumber = 0;
            var GetAttr = CellNode.Attr_CellSpanned();
            if (GetAttr is null)
                return false;

            SpannedNumber = int.Parse(GetAttr.Value);
            return true;
        }

        public static XmlAttribute CreateAttr_ValueType(this XmlNode CellNode)
        {
            var Document = CellNode.OwnerDocument;
            var Attribute = Document.CreateAttribute(OdsProperty.PATH_ValueType, OdsProperty.OfficeUri);
            return Attribute;
        }
        public static XmlAttribute CreateAttr_Value(this XmlNode CellNode)
        {
            var Document = CellNode.OwnerDocument;
            var Attribute = Document.CreateAttribute(OdsProperty.PATH_Value, OdsProperty.OfficeUri);
            return Attribute;
        }
        public static XmlElement CreateElement_Text(this XmlNode CellNode)
        {
            var Document = CellNode.OwnerDocument;
            var Element = Document.CreateElement(OdsProperty.PATH_Text, OdsProperty.TextUri);
            return Element;
        }

        public static string GetCellValue(this XmlNode CellNode)
        {
            var CellValue = CellNode.Attr_Value();
            if (CellValue is not null)
                return CellValue.Value;
            else
                return string.IsNullOrEmpty(CellNode.InnerText) ? null : CellNode.InnerText;
        }
        public static XmlNode SetValue(this XmlNode CellNode, string Value)
        {
            var Attr = CellNode.Attr_ValueType();
            Attr.Value = "string";
            CellNode.SetInnerText(Value);
            return CellNode;
        }
        public static XmlNode SetValue(this XmlNode CellNode, int Value) => CellNode.SetValue((decimal)Value);
        public static XmlNode SetValue(this XmlNode CellNode, double Value) => CellNode.SetValue((decimal)Value);
        public static XmlNode SetValue(this XmlNode CellNode, decimal Value)
        {
            var ValueTypeAttr = CellNode.Attr_ValueType();
            ValueTypeAttr.Value = "float";

            var OfficeValueAttr = CellNode.Attr_Value_Create();
            OfficeValueAttr.Value = Value.ToString();

            CellNode.SetInnerText(Value.ToString());
            return CellNode;
        }
        public static XmlNode SetInnerText(this XmlNode CellNode, string Value)
        {
            var CreateCellValue = CellNode.CreateElement_Text();
            CreateCellValue.InnerText = Value.ToString();

            CellNode.Select()
                .ForEach(Item => CellNode.RemoveChild(Item));

            CellNode.AppendChild(CreateCellValue);
            return CellNode;
        }

        public static XmlNode ClearValue(this XmlNode CellNode)
        {
            foreach (XmlNode Data in CellNode.ChildNodes)
                CellNode.RemoveChild(Data);

            CellNode.Attributes.Remove(CellNode.Attr_ValueType());
            CellNode.Attributes.Remove(CellNode.Attr_Value());

            return CellNode;
        }

        #endregion
    }
}