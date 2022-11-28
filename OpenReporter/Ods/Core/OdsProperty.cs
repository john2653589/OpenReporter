using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rugal.Net.OpenReporter.Ods.Core
{
    public static class OdsProperty
    {
        #region ReadOnly Filed
        public static readonly ImmutableDictionary<string, string> OdsNamespaces = new Dictionary<string, string>
        {
            {"table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0"},
            {"office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"},
            {"style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0"},
            {"text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0"},
            {"draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"},
            {"fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"},
            {"dc", "http://purl.org/dc/elements/1.1/"},
            {"meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0"},
            {"number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"},
            {"presentation", "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"},
            {"svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"},
            {"chart", "urn:oasis:names:tc:opendocument:xmlns:chart:1.0"},
            {"dr3d", "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"},
            {"math", "http://www.w3.org/1998/Math/MathML"},
            {"form", "urn:oasis:names:tc:opendocument:xmlns:form:1.0"},
            {"script", "urn:oasis:names:tc:opendocument:xmlns:script:1.0"},
            {"ooo", "https://openoffice.org/2004/office"},
            {"ooow", "https://openoffice.org/2004/writer"},
            {"oooc", "https://openoffice.org/2004/calc"},
            {"dom", "http://www.w3.org/2001/xml-events"},
            {"xforms", "http://www.w3.org/2002/xforms"},
            {"xsd", "http://www.w3.org/2001/XMLSchema"},
            {"xsi", "http://www.w3.org/2001/XMLSchema-instance"},
            {"rpt", "https://openoffice.org/2005/report"},
            {"of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2"},
            {"rdfa", "https://docs.oasis-open.org/opendocument/meta/rdfa#"},
            {"config", "urn:oasis:names:tc:opendocument:xmlns:config:1.0"},
            { "xlink", "http://www.w3.org/1999/xlink" },
        }.ToImmutableDictionary();

        public static readonly string PATH_SheetNodes = "/office:document-content/office:body/office:spreadsheet/table:table";

        public static readonly string PATH_TableName = "table:name";

        public static readonly string PATH_RowNodes = "table:table-row";
        public static readonly string PATH_RowRepeated = "table:number-rows-repeated";

        public static readonly string PATH_Table_CellNodes = "table:table-cell";
        public static readonly string PATH_Table_CellRepeated = "table:number-columns-repeated";
        public static readonly string PATH_Table_CellSpanned = "table:number-columns-spanned";

        public static readonly string PATH_Office_Value = "office:value";
        public static readonly string PATH_Office_ValueType = "office:value-type";

        public static readonly string PATH_Text = "text:p";
        public static readonly string PATH_DrawFrame = "draw:frame";
        public static readonly string PATH_DrawImage = "draw:image";

        public static readonly string PATH_Office_BinaryData = "office:binary-data";

        public static readonly string PATH_Xlink_Show = "xlink:show";
        public static readonly string PATH_Xlink_Actuate = "xlink:actuate";

        #endregion

        #region Url Property
        public static string TableUri => GetNamespaceUri("table");
        public static string TextUri => GetNamespaceUri("text");
        public static string OfficeUri => GetNamespaceUri("office");
        public static string DrawUri => GetNamespaceUri("draw");
        public static string XlinkUri => GetNamespaceUri("xlink");
        #endregion

        #region Method
        public static string GetNamespaceUri(string UriName)
        {
            if (OdsNamespaces.TryGetValue(UriName, out var Value))
                return Value;
            throw new InvalidOperationException("Can't find that namespace URI");
        }
        #endregion
    }
}
