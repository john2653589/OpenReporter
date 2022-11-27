using Rugal.Net.OpenReporter.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rugal.Net.OpenReporter.Extention
{
    public static class IOpenReportAsExtention
    {
        public static IOpenSheet AsOpenSheet(this IOpenSheet Sheet)
        {
            return Sheet;
        }
    }
}
