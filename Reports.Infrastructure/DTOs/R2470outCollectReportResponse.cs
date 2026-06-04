using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.DTOs
{
    public class R2470outCollectReportResponse
    {
        public Manifest Manifest { get; set; }

        public CollectReleaseMaster CollectReleaseMaster { get; set; }

        public CustomersList CustomersList { get; set; }

        public List<EntryLineMoveView> EntryLineMoveList { get; set; }

        public string SumQuantityMove => EntryLineMoveList.Sum(item => item.LineQuantityMove)?.ToString("N0") ?? string.Empty;

        public string MoveDate { get; set; }
        public string DateOpen => MoveDate ?? DateTime.Now.ToString("dd/MM/yyyy HH:mm");

        public ReportDtl ReportDtl { get; set; }

        public int UnitedMovRef { get; set; }



        public R2470outCollectReportResponse()
        {
            Manifest = new Manifest();
            CollectReleaseMaster = new CollectReleaseMaster();
            CustomersList = new CustomersList();
            EntryLineMoveList = new List<EntryLineMoveView>();
            ReportDtl = new ReportDtl();
        }
    }
}
