/****************************************************************************************
 FILE VERSION: 2 (2026-11-24)

 Changelog:
   v1 - Baseline as provided.
   v2 (2026-11-24) - ContainersNumber: when no container arrived yet in a 1050 message
                      (Control1050List is empty), the container numbers are taken from
                      Consignment.ConsignmentDescription (mapped from message 91), where
                      they appear as "מספר מכולה <number>". Containers from 1050 still win
                      when they exist.
****************************************************************************************/
using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace Reports.Infrastructure.DTOs
{
    public class R912470ReportResponse
    {
        public Manifest Manifest { get; set; }

        public Consignment Consignment { get; set; }

        public List<Item> ItemsList { get; set; }

        public List<Control1050> Control1050List { get; set; }

        public string TotalSum => ItemsList.Sum(item => item.Quantity)?.ToString("N0") ?? string.Empty;

        // Containers from 1050; before a 1050 arrives, the numbers written in the 91 description
        public string ContainersNumber => Control1050List.Any()
            ? string.Join(", ", Control1050List.Select(item => item.ContainerNumber))
            : string.Join(", ", ContainersFromDescription());

        // Matches "מספר מכולה XXXX1234567" (4 letters + 7 digits) in the consignment description
        private static readonly Regex ContainerInDescription =
            new Regex(@"מספר מכולה\s+([A-Za-z]{4}\d{7})", RegexOptions.Compiled);

        private IEnumerable<string> ContainersFromDescription()
        {
            var description = Consignment?.ConsignmentDescription;
            if (string.IsNullOrEmpty(description))
                return Enumerable.Empty<string>();

            return ContainerInDescription.Matches(description)
                .Cast<Match>()
                .Select(m => m.Groups[1].Value.ToUpper())
                .Distinct();
        }



        public R912470ReportResponse()
        {
            Manifest = new Manifest();
            Consignment = new Consignment();
            ItemsList = new List<Item>();
            Control1050List = new List<Control1050>();
        }
    }
}
