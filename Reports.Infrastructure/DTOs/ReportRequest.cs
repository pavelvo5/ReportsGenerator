using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Reports.Infrastructure.Models.Enums;

namespace Reports.Infrastructure.DTOs
{
    public class ReportRequest
    {
        [Required]
        public string ReportID { get; set; }
        [Required]
        public string ManifestID { get; set; }
        [Required]
        public Dictionary<string, object> Parameters { get; set; }
        public bool IsPrint { get; set; }
        public string PrinterName { get; set; }
        [Required]
        public string User { get; set; }
        [Required]
        public string ConnectionString { get; set; }

        /// <summary>
        /// Optional. Requested output byte-format for the response: "PDF" or "Excel".
        /// Only relevant when IsPrint is false (download/email flows). Has no effect on printing.
        /// Left null/empty (the default) => 100% legacy behavior, including the existing
        /// PrinterName == "PDF" convention used when printing to a virtual PDF printer.
        /// Only meaningful for reports whose native ReportsDtl.ReportFormat is "Excel"
        /// (i.e. handled by ExcelReportGenerator); ignored by PDFReportGenerator.
        /// </summary>
        public string OutputFormat { get; set; }



        public ReportRequest()
        {
            Parameters = new Dictionary<string, object>();
        }
    }
}
