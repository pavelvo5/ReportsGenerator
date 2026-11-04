/****************************************************************************************
 FILE VERSION: 2 (2026-11-03)

 Changelog:
   v1 - Baseline as provided.
   v2 (2026-11-03) - DateOpen (the date printed by Title1.html) is now the delivery date
                      and time, SpecialAct.CommitTime, instead of the print time. Falls
                      back to the print time when CommitTime is empty.
****************************************************************************************/
using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.DTOs
{
    public class R60ExOutReportResponse
    {
        public Manifest Manifest { get; set; }

        public Consignment Consignment { get; set; }

        public List<EntryLineMoveView> EntryLineMoveList { get; set; }

        // Delivery date and time; print time when CommitTime is empty
        public string DateOpen => (SpecialAct?.CommitTime ?? DateTime.Now).ToString("dd/MM/yyyy HH:mm");

        public ReportDtl ReportDtl { get; set; }

        public int VarSequence { get; set; }

        public SpecialAct SpecialAct { get; set; }

        public string DriverName { get; set; }

        public string DriverID { get; set; }

        public string LicensePlate { get; set; }

        public string SumExDelivered => EntryLineMoveList.Sum(item => item.exDelivered)?.ToString("N0") ?? string.Empty;



        public R60ExOutReportResponse()
        {
            Manifest = new Manifest();
            Consignment = new Consignment();
            EntryLineMoveList = new List<EntryLineMoveView>();
            ReportDtl = new ReportDtl();
            SpecialAct = new SpecialAct();

        }
    }
}
