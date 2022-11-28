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
            var CellNodes = RowNode.SelectNodes(OdsProperty.PATH_Table_CellNodes, NamespaceManager);
            return CellNodes;
        }
        public static XmlAttribute Attr_Value(this XmlNode CellNode)
        {
            var CellValueArrt = CellNode.Attributes[OdsProperty.PATH_Office_Value];
            return CellValueArrt;
        }
        public static XmlAttribute Attr_Value_Create(this XmlNode CellNode)
        {
            var CellValueArrt = CellNode.Attributes[OdsProperty.PATH_Office_Value];
            if (CellValueArrt is null)
            {
                CellValueArrt = CellNode.CreateAttr_Value();
                CellNode.Attributes.Append(CellValueArrt);
            }
            return CellValueArrt;
        }
        public static XmlAttribute Attr_ValueType(this XmlNode CellNode)
        {
            var ValueTypeArrt = CellNode.Attributes[OdsProperty.PATH_Office_ValueType];
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
            var CellRepeatedAttr = CellNode.Attributes[OdsProperty.PATH_Table_CellRepeated];
            return CellRepeatedAttr;
        }
        public static XmlAttribute Attr_CellSpanned(this XmlNode CellNode)
        {
            var CellSpannedAttr = CellNode.Attributes[OdsProperty.PATH_Table_CellSpanned];
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
            var Attribute = Document.CreateAttribute(OdsProperty.PATH_Office_ValueType, OdsProperty.OfficeUri);
            return Attribute;
        }
        public static XmlAttribute CreateAttr_Value(this XmlNode CellNode)
        {
            var Document = CellNode.OwnerDocument;
            var Attribute = Document.CreateAttribute(OdsProperty.PATH_Office_Value, OdsProperty.OfficeUri);
            return Attribute;
        }

        public static XmlAttribute CreateAttr_Xlink_Show(this XmlNode CellNode)
        {
            var Document = CellNode.OwnerDocument;
            var Attribute = Document.CreateAttribute(OdsProperty.PATH_Xlink_Show, OdsProperty.XlinkUri);
            return Attribute;
        }
        public static XmlAttribute CreateAttr_Xlink_Actuate(this XmlNode CellNode)
        {
            var Document = CellNode.OwnerDocument;
            var Attribute = Document.CreateAttribute(OdsProperty.PATH_Xlink_Actuate, OdsProperty.XlinkUri);
            return Attribute;
        }
        public static XmlAttribute CreateAttr(this XmlNode CellNode, string AttrPath, string AttrUri)
        {
            var Document = CellNode.OwnerDocument;
            var Attribute = Document.CreateAttribute(AttrPath, AttrUri);
            return Attribute;
        }
        public static XmlElement CreateElement_Text(this XmlNode CellNode)
        {
            var Document = CellNode.OwnerDocument;
            var Element = Document.CreateElement(OdsProperty.PATH_Text, OdsProperty.TextUri);
            return Element;
        }
        public static XmlElement CreateElement_DrawFrame(this XmlNode CellNode)
        {
            var Document = CellNode.OwnerDocument;
            var Element = Document.CreateElement(OdsProperty.PATH_DrawFrame, OdsProperty.DrawUri);
            return Element;
        }
        public static XmlElement CreateElement_DrawImage(this XmlNode CellNode)
        {
            var Document = CellNode.OwnerDocument;
            var Element = Document.CreateElement(OdsProperty.PATH_DrawImage, OdsProperty.DrawUri);
            return Element;
        }
        public static XmlElement CreateElement_BinaryData(this XmlNode CellNode)
        {
            var Document = CellNode.OwnerDocument;
            var Element = Document.CreateElement(OdsProperty.PATH_Office_BinaryData, OdsProperty.OfficeUri);
            return Element;
        }
        public static XmlNode AddAttr(this XmlNode Node, string AttrPath, string AttrUri, string Value)
        {
            var Attribute = Node.CreateAttr(AttrPath, AttrUri);
            Attribute.Value = Value;
            Node.Attributes.Append(Attribute);
            return Node;
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
        public static XmlNode SetImage(this XmlNode CellNode, string Base64Value, string Width, string Height)
        {
            CellNode.ClearValue();

            var DrawFrame = CellNode
                .CreateElement_DrawFrame()
                .AddAttr("draw:name", "draw", "img1")
                .AddAttr("svg:width", "svg", Width)
                .AddAttr("svg:height", "svg", Height)
                ;

            var DrawImage = CellNode.CreateElement_DrawImage();

            var Attr_XlinkShoe = DrawImage.CreateAttr_Xlink_Show();
            Attr_XlinkShoe.Value = "embed";
            DrawImage.Attributes.Append(Attr_XlinkShoe);

            var Attr_Xlink_Actuate = DrawImage.CreateAttr_Xlink_Actuate();
            Attr_Xlink_Actuate.Value = "onLoad";
            DrawImage.Attributes.Append(Attr_Xlink_Actuate);


            var BinaryData = CellNode.CreateElement_BinaryData();
            BinaryData.InnerText = Base64Value;

            DrawImage.AppendChild(BinaryData);
            DrawFrame.AppendChild(DrawImage);

            CellNode.AppendChild(DrawFrame);
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