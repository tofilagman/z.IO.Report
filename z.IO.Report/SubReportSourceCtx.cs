using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace z.IO.Report
{
    public class SubReportSourceCtx
    {
        public string Name { get; set; }
        public DataTable DataSource { get; set; }
    }
}
