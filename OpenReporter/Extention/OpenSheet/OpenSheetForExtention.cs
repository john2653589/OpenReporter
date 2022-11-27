using Rugal.Net.OpenReporter.Interface;
using Rugal.Net.OpenReporter.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rugal.Net.OpenReporter.Extention
{
    public static class OpenSheetForExtention
    {
        public static IOpenSheet ForRow(this IOpenSheet Sheet, Action<IOpenRow> RowAction)
        {
            RowAction.Invoke(Sheet.CurrentRow());
            return Sheet;
        }
        public static IOpenSheet ForRow(this IOpenSheet Sheet, Func<IOpenRow, bool> RowAction)
        {
            var Row = Sheet.CurrentRow();
            while (RowAction.Invoke(Row))
            {
                Row = Sheet.NextRow();
            }
            return Sheet;
        }
        public static IOpenSheet For(this IOpenSheet Sheet, Action SheetAction)
        {
            SheetAction.Invoke();
            return Sheet;
        }
    }
}