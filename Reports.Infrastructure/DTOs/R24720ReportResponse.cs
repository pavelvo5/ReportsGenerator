/****************************************************************************************
 FILE VERSION: 1 (2026-08-13)

 Changelog:
   v1 (2026-08-13) - Added HasContainers (computed bool) to fix the container table
                      appearing twice, back-to-back, in R24720Report.html. The table's
                      outer wrapper had been gated on Control1050List itself, but
                      Mustache's {{#list}} always iterates - it can't mean "show once
                      if non-empty" - so the whole table (including its header)
                      repeated once per container, with the inner row-loop re-
                      rendering inside each repeat. HasContainers is a genuine boolean
                      instead, used only to gate the table; the inner per-row loop
                      still uses Control1050List unchanged.

 Base: this file had not been modified before this change - edited directly from the
 original ReportsGenerator-main.zip (uploaded 2026-07-19), so there is no prior
 version to be out of sync with.
****************************************************************************************/
using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.DTOs
{
    public class R24720ReportResponse
    {
        public Manifest Manifest { get; set; }

        public Consignment Consignment { get; set; }

        public List<Item> ItemsList { get; set; }

        public List<Control1050> Control1050List { get; set; }

        public List<EntryLineMoveView> EntryLineMoveList { get; set; }

        public string SumWeight => ItemsList.Sum(item => item.ItemWeight)?.ToString("N0") ?? string.Empty;

        public string SumQuantityMove => EntryLineMoveList.Sum(item => item.LineQuantityMove)?.ToString("N0") ?? string.Empty;

        public string DateOpen => Consignment.OpeningDate?.ToString("dd/MM/yyyy HH:mm");

        // Used to show/hide the container table as a simple yes/no check, separate from
        // Control1050List itself (which is also used to loop the table's rows). Mustache
        // has no "if list is non-empty, but don't iterate" construct - {{#name}} on a list
        // always iterates - so gating the table on the list itself caused it to repeat once
        // per container, with the inner row-loop re-rendering inside each repeat. This plain
        // bool is a genuine conditional instead.
        public bool HasContainers => Control1050List != null && Control1050List.Any();

        public ReportDtl ReportDtl { get; set; }

        public int VarSequence { get; set; }



        public R24720ReportResponse()
        {
            Manifest = new Manifest();
            Consignment = new Consignment();
            ItemsList = new List<Item>();
            Control1050List = new List<Control1050>();
            EntryLineMoveList = new List<EntryLineMoveView>();
        }
    }
}
