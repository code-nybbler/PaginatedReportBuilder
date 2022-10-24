﻿using System.Collections.Generic;
using System.Xml.Linq;

namespace PaginatedReportGenerator
{
    internal class ViewMeta
    {
        public List<string> fields;
        public XDocument fetchxml;

        public ViewMeta(List<string> fields, XDocument fetchxml)
        {
            this.fields = fields;
            this.fetchxml = fetchxml;
        }
    }
}
