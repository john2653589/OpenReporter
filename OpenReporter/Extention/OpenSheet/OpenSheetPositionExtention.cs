using Rugal.Net.OpenReporter.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rugal.Net.OpenReporter.Extention
{
    public static class OpenSheetPositionExtention
    {
        public static IOpenSheet Position_Row(this IOpenSheet Sheet, int RowIndex)
        {
            Sheet.InitPosition()
                .TrySet_RowIndex(RowIndex);

            return Sheet;
        }

        public static IOpenRow CurrentRow(this IOpenSheet Sheet)
        {
            var CurrentRow = Sheet.FindRows(Sheet.CurrentRowIndex);
            return CurrentRow;
        }

        public static IOpenRow NextRow(this IOpenSheet Sheet)
        {
            Sheet.ForPosition.AddRow();
            var NextRow = Sheet.CurrentRow();
            return NextRow;
        }

    }
}
