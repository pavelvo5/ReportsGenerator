/****************************************************************************************
 FILE VERSION: 1 (2026-10-25)

 Changelog:
   v1 (2026-10-25) - New model for table SpecialActDtl, used by the specialRequest report.
                      The same class is also used for the "required inventory" rows
                      (built from SpecialActRepackNew by the stored procedure).
****************************************************************************************/
using Dapper.Contrib.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Models
{
    [Table("SpecialActDtl")]
    public class SpecialActDtl
    {
        [Key]
        public int RecID { get; set; }

        public int? SpecialAct_RecID { get; set; }

        public string ManifestID { get; set; }

        public string ConsignmentID { get; set; }

        public int? RequestID { get; set; }

        public short? ActivityType { get; set; }

        public short? RowNumber { get; set; }

        public string ReqRowDesc { get; set; }

        public int? Quantity { get; set; }

        public double? Weight { get; set; }

        public long? CustomsItem { get; set; }

        public double? SampleValue { get; set; }

        public string CurrencyTypeCode { get; set; }

        public DateTime? SampleReturnDate { get; set; }

        public string FormattedQuantity => Quantity?.ToString("N0") ?? string.Empty;

        public string FormattedWeight => Weight?.ToString("0.##") ?? string.Empty;
    }
}
