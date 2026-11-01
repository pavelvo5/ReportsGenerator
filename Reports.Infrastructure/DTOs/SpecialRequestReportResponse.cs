/****************************************************************************************
 FILE VERSION: 1 (2026-10-25)

 Changelog:
   v1 (2026-10-25) - New response model for the PDF report "specialRequest"
                      (בקשה לפעולה מיוחדת). Filled by GetDataForspecialRequestReport
                      in ReportRepositoryDapper. The Is* flags choose which bottom section
                      of the template is printed (5 = repack, 7 = sample, otherwise other).
                      Labels holds the Hebrew table headings, returned by the stored procedure.
                      Uses the shared Title1.html: VarSequence (= RequestID) and DateOpen
                      (print time, as in R60ExOut) feed its title line and date.
****************************************************************************************/
using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.DTOs
{
    public class SpecialRequestReportResponse
    {
        public Manifest Manifest { get; set; }

        public Consignment Consignment { get; set; }

        public SpecialAct SpecialAct { get; set; }

        public List<SpecialActDtl> CurrentLineList { get; set; }

        public List<SpecialActDtl> RequiredLineList { get; set; }

        public SpecialRequestLabels Labels { get; set; }

        public ReportDtl ReportDtl { get; set; }

        // Number printed by Title1.html after the report name (= RequestID)
        public int VarSequence { get; set; }

        // Date printed by Title1.html (print time, as in R60ExOut)
        public string DateOpen => DateTime.Now.ToString("dd/MM/yyyy HH:mm");

        public bool IsRepack => SpecialAct?.ActivityType == 5;

        public bool IsSample => SpecialAct?.ActivityType == 7;

        public bool IsOther => !IsRepack && !IsSample;

        // ActivityTypeDesc + space + AdditionalDesc (shown for "other" activity types)
        public string OtherDescription => $"{SpecialAct?.ActivityTypeDesc} {SpecialAct?.AdditionalDesc}".Trim();

        public SpecialRequestReportResponse()
        {
            Manifest = new Manifest();
            Consignment = new Consignment();
            SpecialAct = new SpecialAct();
            CurrentLineList = new List<SpecialActDtl>();
            RequiredLineList = new List<SpecialActDtl>();
            Labels = new SpecialRequestLabels();
            ReportDtl = new ReportDtl();
        }
    }
}
