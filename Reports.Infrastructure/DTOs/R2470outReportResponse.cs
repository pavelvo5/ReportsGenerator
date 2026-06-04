using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.DTOs
{
    public class R2470outReportResponse
    {
        public Manifest Manifest { get; set; }

        public Consignment Consignment { get; set; }

        public ConsignmentRelease ConsignmentRelease { get; set; }

        public List<EntryLinesMovesViewCalc> EntryLinesMovesViewCalcList { get; set; }

        public string DateOpen => MoveDate ?? DateTime.Now.ToString("dd/MM/yyyy HH:mm");
        public string MoveDate { get; set; }

        public int? Balance { get; set; }

        public string FormattedBalance => Balance?.ToString("N0") ?? string.Empty;

        public string TotalLineQuantityDeclared => EntryLinesMovesViewCalcList.Sum(item => item.LineQuantityDeclared)?.ToString("N0") ?? string.Empty;

        public string TotalQuantityMoveBefore => EntryLinesMovesViewCalcList.Sum(item => item.TotalQuantityMoveBefore)?.ToString("N0") ?? string.Empty;

        public string TotalBalanceDelivery => EntryLinesMovesViewCalcList.Sum(item => item.BalanceDelivery)?.ToString("N0") ?? string.Empty;

        public string TotalQuantityMoveNow => EntryLinesMovesViewCalcList.Sum(item => item.TotalQuantityMoveNow)?.ToString("N0") ?? string.Empty;

        public string TotalNewBalanceDelivery => EntryLinesMovesViewCalcList.Sum(item => item.NewBalanceDelivery)?.ToString("N0") ?? string.Empty;

        public string TotalNotArrived => EntryLinesMovesViewCalcList.Sum(item => item.NotArrived)?.ToString("N0") ?? string.Empty;

        public ReportDtl ReportDtl { get; set; }

        public int VarSequence { get; set; }

        public string Driver { get; set; }

        public string Remarks { get; set; }

        public string User { get; set; }

        public string TruckID { get; set; }

        public GoodList GoodList { get; set; }



        public R2470outReportResponse()
        {
            Manifest = new Manifest();
            Consignment = new Consignment();
            ConsignmentRelease = new ConsignmentRelease();
            EntryLinesMovesViewCalcList = new List<EntryLinesMovesViewCalc>();
            ReportDtl = new ReportDtl();
            GoodList = new GoodList();

        }
    }
}
