/****************************************************************************************
 FILE VERSION: 1 (2026-10-25)

 Changelog:
   v1 (2026-10-25) - New class. Hebrew table headings of the specialRequest report, read from
                      the last result set of GetDataForspecialRequestReport and printed by
                      the template as {{Labels_...}}.
****************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.DTOs
{
    public class SpecialRequestLabels
    {
        public string ActivityType { get; set; }

        public string Requester { get; set; }

        public string Authority { get; set; }

        public string DecisionDate { get; set; }

        public string CommitTime { get; set; }

        public string SampleReturn { get; set; }

        public string Remark { get; set; }

        public string PerformedBy { get; set; }

        public string CurrentInventory { get; set; }

        public string RequiredInventory { get; set; }

        public string RowNumber { get; set; }

        public string Description { get; set; }

        public string CustomsItem { get; set; }

        public string Quantity { get; set; }

        public string Weight { get; set; }
    }
}
