using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.DTOs
{
    public class EntryLinesMovesViewCalc
    {
        public int? LineNum { get; set; }

        public string LineID { get; set; }

        public string LinePackType { get; set; }

        public int? LineQuantityDeclared { get; set; }

        public int? TotalQuantityMoveBefore { get; set; }

        public int? BalanceDelivery { get; set; }

        public int? TotalQuantityMoveNow { get; set; }

        public int? NewBalanceDelivery { get; set; }

        public int? NotArrived { get; set; }

        public DateTime? LastUpdated { get; set; }

        public string FormattedLineQuantityDeclared => LineQuantityDeclared?.ToString("N0") ?? string.Empty;

        public string FormattedTotalQuantityMoveBefore => TotalQuantityMoveBefore?.ToString("N0") ?? string.Empty;

        public string FormattedBalanceDelivery => BalanceDelivery?.ToString("N0") ?? string.Empty;

        public string FormattedTotalQuantityMoveNow => TotalQuantityMoveNow?.ToString("N0") ?? string.Empty;

        public string FormattedNewBalanceDelivery => NewBalanceDelivery?.ToString("N0") ?? string.Empty;

        public string FormattedNotArrived => NotArrived?.ToString("N0") ?? string.Empty;

        public DateTime? ExpiryDate { get; set; }


    }
}
