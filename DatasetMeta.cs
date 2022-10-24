using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PaginatedReportGenerator
{
    internal class DatasetMeta
    {
        public string name;
        public string fetchxml;

        public DatasetMeta(string name, string fetchxml)
        {
            this.name = name;
            this.fetchxml = fetchxml;
        }
    }
}
