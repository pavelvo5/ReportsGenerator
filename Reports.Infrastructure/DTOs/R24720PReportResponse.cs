using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.DTOs
{
    public class R24720PReportResponse
    {
        public Manifest Manifest { get; set; }

        public Consignment Consignment { get; set; }

        public List<EntryLinesMovesViewCalc> EntryLinesMovesViewCalcList { get; set; }

        public string SumLineQuantityDeclared => EntryLinesMovesViewCalcList.Sum(item => item.LineQuantityDeclared)?.ToString("N0") ?? string.Empty;

        public string SumTotalQuantityMoveBefore => EntryLinesMovesViewCalcList.Sum(item => item.TotalQuantityMoveBefore)?.ToString("N0") ?? string.Empty;

        public string SumBalanceDelivery => EntryLinesMovesViewCalcList.Sum(item => item.BalanceDelivery)?.ToString("N0") ?? string.Empty;

        public string SumTotalQuantityMoveNow => EntryLinesMovesViewCalcList.Sum(item => item.TotalQuantityMoveNow)?.ToString("N0") ?? string.Empty;

        public string SumNewBalanceDelivery => EntryLinesMovesViewCalcList.Sum(item => item.NewBalanceDelivery)?.ToString("N0") ?? string.Empty;

        public string MoveDate { get; set; }
        public string DateOpen => MoveDate ?? DateTime.Now.ToString("dd/MM/yyyy HH:mm");

        public ReportDtl ReportDtl { get; set; }

        public string VarSequence => $"{Consignment.Gush}-1";



        public R24720PReportResponse()
        {
            Manifest = new Manifest();
            Consignment = new Consignment();
            EntryLinesMovesViewCalcList = new List<EntryLinesMovesViewCalc>();
            ReportDtl = new ReportDtl();

        }
    }
}

