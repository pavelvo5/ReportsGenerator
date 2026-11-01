/****************************************************************************************
 FILE VERSION: 2 (2026-10-25)

 Changelog:
   v1 - Baseline as provided.
   v2 (2026-10-25) - Added DecisionAppName (existing column of table SpecialAct) and three
                      properties filled only when the stored procedure returns them
                      (GetDataForspecialRequestReport): RequesterTypeDesc,
                      LatestSampleReturnDate; plus formatted date strings for the new report.
                      Other reports using this model are not affected (the properties stay null).
****************************************************************************************/
using Dapper.Contrib.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Models
{
    [Table("SpecialAct")]
    public class SpecialAct
    {
        [Key]
        public int RecID { get; set; }

        public string ManifestID { get; set; }

        public string ConsignmentID { get; set; }

        public int? RequestID { get; set; }

        public DateTime? DecisionDate { get; set; }

        public short? ActivityType { get; set; }

        public string ActivityTypeDesc { get; set; }

        public string AdditionalDesc { get; set; }

        public short? RequesterType { get; set; }

        public short? AuthorityCode { get; set; }

        public short? ActivityStat { get; set; }

        public DateTime? CommitTime { get; set; }

        public string PerformedBy { get; set; }

        public string Remark { get; set; }

        public short? SpecialActivityTypeEssence { get; set; }

        public string ResponseStat60 { get; set; }

        public string ResponseID60 { get; set; }

        public string ResponseErr60 { get; set; }

        public DateTime? TransmissionTime60 { get; set; }

        public DateTime? LastSuccessTrans60 { get; set; }

        public string AuthorityDesc { get; set; }

        public string DecisionAppName { get; set; }

        // Filled only by GetDataForspecialRequestReport
        public string RequesterTypeDesc { get; set; }

        // Filled only by GetDataForspecialRequestReport (MAX of SpecialActDtl.SampleReturnDate)
        public DateTime? LatestSampleReturnDate { get; set; }

        public string FormattedDecisionDate => DecisionDate?.ToString("dd/MM/yy") ?? string.Empty;

        // Decision date followed by the approving application/user in brackets, when there is one
        public string FormattedDecision => FormattedDecisionDate + (string.IsNullOrEmpty(DecisionAppName) ? string.Empty : $" ({DecisionAppName})");

        public string FormattedCommitTime => CommitTime?.ToString("dd/MM/yy HH:mm") ?? string.Empty;

        public string FormattedLatestSampleReturnDate => LatestSampleReturnDate?.ToString("dd/MM/yy") ?? string.Empty;

    }
}
