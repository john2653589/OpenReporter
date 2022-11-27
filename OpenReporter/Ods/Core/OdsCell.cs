using Rugal.Net.OpenReporter.Interface;
using Rugal.Net.OpenReporter.Model;
using Rugal.Net.OpenReporter.Ods.Extention;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Rugal.Net.OpenReporter.Ods.Core
{
    public class OdsCell : IOpenCell
    {
        #region Interface Property
        public IOpenRow Row => OdsRow;
        #endregion

        #region Ods Property
        public int AbsColumnIndex { get; private set; }
        public OdsRow OdsRow { get; }
        public XmlNode CellNode { get; }
        public CellPosition Position { get; set; }
        #endregion

        public OdsCell(int _AbsColumnIndex, XmlNode _CellNode, OdsRow _OdsRow)
        {
            AbsColumnIndex = _AbsColumnIndex;
            OdsRow = _OdsRow;
            CellNode = _CellNode;

            InitPosition();
        }

        #region Interface Method
        public object GetValue()
        {
            var GetValue = CellNode.GetCellValue();
            return GetValue;
        }
        public TValue ConvertValue<TValue>(object Value)
        {
            return default;
        }

        public void SetValue(object Value)
        {
            if (Value is null)
                CellNode.SetValue("");
            else if (Value is string StringValue)
                CellNode.SetValue(StringValue);
            else if (Value is int IntValue)
                CellNode.SetValue(IntValue);
            else if (Value is double DoubleValue)
                CellNode.SetValue(DoubleValue);
            else if (Value is decimal DecimalValue)
                CellNode.SetValue(DecimalValue);
            else
                CellNode.SetValue(Value.ToString());
        }
        #endregion

        #region Internal Method
        internal void InitPosition()
        {
            var RowIndex = OdsRow.AbsRowIndex;
            Position = new CellPosition(RowIndex, AbsColumnIndex);
        }

        #endregion

    }
}
