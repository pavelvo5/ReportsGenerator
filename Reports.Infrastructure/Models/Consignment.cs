using Dapper.Contrib.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reports.Infrastructure.Models
{
    [Table("Consignments")]
    public class Consignment
    {
        [Key]
        public int RecID { get; set; }

        public int? ManifestRecID { get; set; }

        public string ManifestID { get; set; }

        public string ConsignmentID { get; set; }

        public string CargoIdentifierType { get; set; }

        public string CargoIdentifierKey1 { get; set; }

        public string CargoIdentifierKey2 { get; set; }

        public string CargoIdentifierKey3 { get; set; }

        public string ConsignmentStatus { get; set; }

        public DateTime? ReceiveDate { get; set; }

        public DateTime? LastUpdatedDate { get; set; }

        public string FatherConsignment { get; set; }

        public string DeliverySiteNumber { get; set; }

        public string StoringSiteNumber { get; set; }

        public string LoadingSiteNumber { get; set; }

        public string UnloadingSiteNumber { get; set; }

        public string ShipVisitID { get; set; }

        public DateTime? ExpectedArrival { get; set; }

        public string ShipAgentID { get; set; }

        public string AutonomyRegionType { get; set; }

        public string ImporterID { get; set; }

        public long? DeclarationID { get; set; }

        public string ConsignmentDescription { get; set; }

        public short? ReleaseMessageCode { get; set; }

        public DateTime? ReleaseDate { get; set; }

        public string GeneralDataType { get; set; }

        public string CustomsAgent { get; set; }

        public string CustomsAgentID { get; set; }

        public short? BilledType { get; set; }

        public string BilledCA { get; set; }

        public string BilledCAID { get; set; }

        public string BilledImporter { get; set; }

        public string BilledImporterID { get; set; }

        public string BOL { get; set; }

        public string CustomsAgentFile { get; set; }

        public string ImporterName { get; set; }

        public double? CifValueNis { get; set; }

        public double? DealValueNIS { get; set; }

        public double? TotalTax { get; set; }

        public double? FobValueNis { get; set; }

        public double? FOBvaluePC { get; set; }

        public double? ExchangeRate { get; set; }

        public string CurrencyType { get; set; }

        public DateTime? TaxationDate { get; set; }

        public double? GoodsPriceAmountInPrimaryCurrency { get; set; }

        public string PrimaryCurrency { get; set; }

        public int? GovernmentProcedureType { get; set; }

        public int? DeclarationMessageType { get; set; }

        public int? ApplicationID { get; set; }

        public double? DeclarationVer { get; set; }

        public string PackageID { get; set; }

        public string FileName { get; set; }

        public string ShipName { get; set; }

        public string Port { get; set; }

        public short? F20units { get; set; }

        public short? F40units { get; set; }

        public int? TEU { get; set; }

        public int? ContainerLength { get; set; }

        public int? Volume { get; set; }

        public double? StorageArea { get; set; }

        public int? PalletsAllocations { get; set; }

        public int? NumOfPallets { get; set; }

        public int? Gush { get; set; }

        public short? LockInd { get; set; }

        public string InternalRemarks { get; set; }

        public string CustomerRemarks { get; set; }

        public string TransportCompanyName { get; set; }

        public string MSG70RejectReason { get; set; }

        public string MSG70ResponseStat { get; set; }

        public string MSG70ResponseID { get; set; }

        public string MSG70ResponseErr { get; set; }

        public DateTime? MSG70TransmissionTime { get; set; }

        public DateTime? MSG70LastSuccessTransmission { get; set; }

        public string ResponseStat110 { get; set; }

        public string ResponseID110 { get; set; }

        public string ResponseErr110 { get; set; }

        public DateTime? TransmissionTime110 { get; set; }

        public DateTime? LastSuccessTrans110 { get; set; }

        public string NDIReason { get; set; }

        public int? NDIstatus { get; set; }

        public string NDI_IDmsg1050 { get; set; }

        public int? NDI_ReleaseID { get; set; }

        public string ResponseStat3240 { get; set; }

        public string ResponseID3240 { get; set; }

        public string ResponseErr3240 { get; set; }

        public DateTime? TransmissionTime3240 { get; set; }

        public DateTime? LastSuccessTrans3240 { get; set; }

        public short? Mortgaged { get; set; }

        public string Mortgagee { get; set; }

        public DateTime? MortgagedDate { get; set; }

        public DateTime? UnMortgagedDate { get; set; }

        public DateTime? OpeningDate { get; set; }

        public int? GeneralGoodsCode { get; set; }

        public string MortgagedRemark { get; set; }

        public string FormattedDealValueNIS => DealValueNIS?.ToString("N0") ?? string.Empty;

        public string FormattedFobValueNis => FobValueNis?.ToString("N0") ?? string.Empty;

        public string FormattedGush => Gush.ToString().Length > 2 ? $"{Gush.ToString().Substring(0, 2)}/{Gush.ToString().Substring(2)}" : Gush.ToString();

        public string FormattedReleaseDate => ReleaseDate?.ToString("dd/MM/yyyy HH:mm");

        public string FormattedReleaseDateShort => ReleaseDate?.ToString("dd/MM/yy");

        // --- Added for R912470 storage-request header/footer note ---
        // Populated only when the SP returns these columns (e.g. GetDataForR912470Report);
        // left null/default for every other report using this shared model.
        public string ReportTitle { get; set; }

        public bool ShowStorageRequestNote { get; set; }

        public string StorageRequestStatusLabel { get; set; }

        public string FormattedMSG70LastSuccessTransmission => MSG70LastSuccessTransmission?.ToString("dd/MM/yyyy HH:mm") ?? string.Empty;


    }
}
