using Dapper;
using Reports.Infrastructure.DTOs;
using Reports.Infrastructure.Exceptions;
using Reports.Infrastructure.Logger;
using Reports.Infrastructure.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Reports.Infrastructure.Models.Enums;

namespace Reports.Infrastructure.Repositories
{
    public class ReportRepositoryDapper : IReportRepositoryDapper
    {
        private readonly string connectionString;
        private readonly ILogger logger;


        public ReportRepositoryDapper(string connectionString, ILogger logger)
        {
            this.connectionString = connectionString;
            this.logger = logger;
        }

        public async Task<dynamic> GetDataAsync(ReportRequest request, ReportDtl reportDtl)
        {
            try
            {
                Manifest manifest = await GetManifestByManifestIDAsync(request.ManifestID);

                if(manifest == null)
                    throw new CustomException((int)ErrorMessages.ErrorCodes.NoDataFound, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.NoDataFound]);

                //For SP parameters
                request.Parameters["ManifestID"] = request.ManifestID;                

                if (Enum.TryParse(reportDtl.FunctionName, true, out StoredProcedure storedProcedure))
                {
                    switch (storedProcedure)
                    {
                        case StoredProcedure.GetDataForR912470Report:
                            R912470ReportResponse R912470ReportResponse = await GetDataForR912470Report(request.Parameters, manifest);
                            if(R912470ReportResponse.Consignment == null)
                                throw new CustomException((int)ErrorMessages.ErrorCodes.NoDataFound, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.NoDataFound]);
                            return R912470ReportResponse;
                        case StoredProcedure.GetDataForR2470Report:
                            R2470ReportResponse R2470ReportResponse = await GetDataForR2470Report(request.Parameters, manifest);
                            if (R2470ReportResponse.ConsignmentRelease == null)
                                throw new CustomException((int)ErrorMessages.ErrorCodes.NoDataFound, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.NoDataFound]);
                            return R2470ReportResponse;
                        case StoredProcedure.GetDataForR24720Report:
                            R24720ReportResponse R24720ReportResponse = await GetDataForR24720Report(request.Parameters, manifest, reportDtl);
                            if (R24720ReportResponse.Consignment == null)
                                throw new CustomException((int)ErrorMessages.ErrorCodes.NoDataFound, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.NoDataFound]);
                            return R24720ReportResponse;
                        case StoredProcedure.GetDataForR1050MTReport:
                            R1050MTReportResponse R1050MTReportResponse = await GetDataForR1050MTReport(request.Parameters, manifest);
                            if (R1050MTReportResponse.Consignment == null)
                                throw new CustomException((int)ErrorMessages.ErrorCodes.NoDataFound, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.NoDataFound]);
                            return R1050MTReportResponse;
                        case StoredProcedure.GetDataForR2470outReport:
                            R2470outReportResponse R2470outReportResponse = await GetDataForR2470outReport(request.Parameters, manifest, reportDtl, request.User);
                            if (R2470outReportResponse.Consignment == null)
                                throw new CustomException((int)ErrorMessages.ErrorCodes.NoDataFound, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.NoDataFound]);
                            return R2470outReportResponse;
                        case StoredProcedure.GetDataForMrtg24720Report:
                            Mrtg24720ReportResponse Mrtg24720ReportResponse = await GetDataForMrtg24720Report(request.Parameters, manifest, reportDtl);
                            if (Mrtg24720ReportResponse.Consignment == null)
                                throw new CustomException((int)ErrorMessages.ErrorCodes.NoDataFound, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.NoDataFound]);
                            return Mrtg24720ReportResponse;
                        case StoredProcedure.GetDataForR24720PReport:
                            R24720PReportResponse R24720PReportResponse = await GetDataForR24720PReport(request.Parameters, manifest, reportDtl);
                            if (R24720PReportResponse.Consignment == null)
                                throw new CustomException((int)ErrorMessages.ErrorCodes.NoDataFound, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.NoDataFound]);
                            return R24720PReportResponse;
                        case StoredProcedure.GetDataForR60ExOutReport:
                            R60ExOutReportResponse R60ExOutReportResponse = await GetDataForR60ExOutReport(request.Parameters, manifest, reportDtl);
                            if (R60ExOutReportResponse.Consignment == null)
                                throw new CustomException((int)ErrorMessages.ErrorCodes.NoDataFound, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.NoDataFound]);
                            return R60ExOutReportResponse;
                        case StoredProcedure.GetDataForR60ExInReport:
                            R60ExInReportResponse R60ExInReportResponse = await GetDataForR60ExInReport(request.Parameters, manifest, reportDtl);
                            if (R60ExInReportResponse.Consignment == null)
                                throw new CustomException((int)ErrorMessages.ErrorCodes.NoDataFound, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.NoDataFound]);
                            return R60ExInReportResponse;
                        case StoredProcedure.GetDataForR60splitReport:
                            R60splitReportResponse R60splitReportResponse = await GetDataForR60splitReport(request.Parameters, manifest, reportDtl);
                            if (R60splitReportResponse.Consignment == null)
                                throw new CustomException((int)ErrorMessages.ErrorCodes.NoDataFound, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.NoDataFound]);
                            return R60splitReportResponse;
                        case StoredProcedure.GetDataForR2470outCollectReport:
                            R2470outCollectReportResponse R2470outCollectReportResponse = await GetDataForR2470outCollectReport(request.Parameters, manifest, reportDtl);
                            if (R2470outCollectReportResponse.CollectReleaseMaster == null)
                                throw new CustomException((int)ErrorMessages.ErrorCodes.NoDataFound, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.NoDataFound]);
                            return R2470outCollectReportResponse;

                        case StoredProcedure.GetDataForRCollectPlanningReport:
                            RCollectPlanningReportResponse RCollectPlanningReportResponse = await GetDataForRCollectPlanningReport(request.Parameters, manifest, reportDtl);
                            if (RCollectPlanningReportResponse.CollectReleaseMaster == null)
                                throw new CustomException((int)ErrorMessages.ErrorCodes.NoDataFound, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.NoDataFound]);
                            return RCollectPlanningReportResponse;

                        case StoredProcedure.GetDataForInvRepForCustomsReport:
                            InvRepForCustomsReportResponse InvRepForCustomsReportResponse = await GetDataForInvRepForCustomsReport(request.Parameters, manifest, reportDtl);
                            if (InvRepForCustomsReportResponse.Consignment == null)
                                throw new CustomException((int)ErrorMessages.ErrorCodes.NoDataFound, ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.NoDataFound]);
                            return InvRepForCustomsReportResponse;

                        default:
                            break;
                    }

                }
            }
            catch (CustomException ex)
            {
                logger.WriteLog($"Error to Get Data Async: {ex}");
                throw;
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Data Async: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
            return null;
        }

        public async Task<Manifest> GetManifestByManifestIDAsync(string manifestID)
        {
            try
            {               
                using (var conn = new SqlConnection(connectionString))
                {
                    var parameters = new { ManifestID = manifestID };

                    return await conn.QueryFirstOrDefaultAsync<Manifest>(
                        StoredProcedure.GetManifestByManifestID.ToString(),
                        parameters,
                        commandType: CommandType.StoredProcedure);                    
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Manifest By ManifestID Async: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.DBAccessFailure, $"{ ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.DBAccessFailure] } : {ex.Message}");
            }
        }
        public async Task<ReportDtl> GetReportsDtlByReportIDAsync(string reportID)
        {
            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    var parameters = new { ReportID = reportID };

                    return await conn.QueryFirstOrDefaultAsync<ReportDtl>(
                        StoredProcedure.GetReportsDtl.ToString(),
                        parameters,
                        commandType: CommandType.StoredProcedure);
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get ReportsDtl By ReportID Async: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.DBAccessFailure, $"{ ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.DBAccessFailure] } : {ex.Message}");
            }
        }
        public async Task<R912470ReportResponse> GetDataForR912470Report(Dictionary<string, object> parameters, Manifest manifest)
        {
            try
            {
                using (var connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();

                    using (var multi = await connection.QueryMultipleAsync(StoredProcedure.GetDataForR912470Report.ToString(), new DynamicParameters(parameters), commandType: CommandType.StoredProcedure))
                    {                        
                        var response = new R912470ReportResponse
                        {
                            Consignment = await multi.ReadFirstOrDefaultAsync<Consignment>(),
                            ItemsList = (await multi.ReadAsync<Item>()).ToList(),
                            Control1050List = (await multi.ReadAsync<Control1050>()).ToList(),
                            Manifest = manifest
                        };
                        
                        return response;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Data For R912470 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.DBAccessFailure, $"{ ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.DBAccessFailure] } : {ex.Message}");
            }
        }

        public async Task<R2470ReportResponse> GetDataForR2470Report(Dictionary<string, object> parameters, Manifest manifest)
        {
            try
            {
                using (var connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();

                    using (var multi = await connection.QueryMultipleAsync(StoredProcedure.GetDataForR2470Report.ToString(), new DynamicParameters(parameters), commandType: CommandType.StoredProcedure))
                    {
                        var response = new R2470ReportResponse
                        {
                            ConsignmentRelease = await multi.ReadFirstOrDefaultAsync<ConsignmentRelease>(),
                            Consignment = await multi.ReadFirstOrDefaultAsync<Consignment>(),
                            ConsignmentReleaseItemList = (await multi.ReadAsync<ConsignmentReleaseItem>()).ToList(),
                            Manifest = manifest
                        };
                        return response;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Data For R2470 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.DBAccessFailure, $"{ ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.DBAccessFailure] } : {ex.Message}");
            }
        }

        public async Task<R24720ReportResponse> GetDataForR24720Report(Dictionary<string, object> parameters, Manifest manifest, ReportDtl reportDtl)
        {
            try
            {
                var parametersToExclude = new HashSet<string> { "InvMovRecID" };

                var dynamicParams = new DynamicParameters();

                foreach (var item in parameters)
                {
                    if (!parametersToExclude.Contains(item.Key))
                    {
                        dynamicParams.Add(item.Key, item.Value);
                    }
                }
                using (var connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();

                    using (var multi = await connection.QueryMultipleAsync(StoredProcedure.GetDataForR24720Report.ToString(), dynamicParams, commandType: CommandType.StoredProcedure))
                    {
                        var response = new R24720ReportResponse
                        {
                            Consignment = await multi.ReadFirstOrDefaultAsync<Consignment>(),
                            EntryLineMoveList = (await multi.ReadAsync<EntryLineMoveView>()).ToList(),
                            ItemsList = (await multi.ReadAsync<Item>()).ToList(),
                            Control1050List = (await multi.ReadAsync<Control1050>()).ToList(),
                            Manifest = manifest,
                            ReportDtl = reportDtl,
                            VarSequence = Convert.ToInt32(parameters["InvMovRecID"])
                    };
                        
                        return response;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Data For R24720 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.DBAccessFailure, $"{ ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.DBAccessFailure] } : {ex.Message}");
            }
        }

        public async Task<R1050MTReportResponse> GetDataForR1050MTReport(Dictionary<string, object> parameters, Manifest manifest)
        {
            try
            {
                using (var connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();

                    using (var multi = await connection.QueryMultipleAsync(StoredProcedure.GetDataForR1050MTReport.ToString(), new DynamicParameters(parameters), commandType: CommandType.StoredProcedure))
                    {
                        var response = new R1050MTReportResponse
                        {
                            Consignment = await multi.ReadFirstOrDefaultAsync<Consignment>(),
                            Control1050 = await multi.ReadFirstOrDefaultAsync<Control1050>(),
                            Manifest = manifest,
                        };
                        return response;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Data For R1050MT Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.DBAccessFailure, $"{ ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.DBAccessFailure] } : {ex.Message}");
            }
        }

        public async Task<R2470outReportResponse> GetDataForR2470outReport(Dictionary<string, object> parameters, Manifest manifest, ReportDtl reportDtl, string user)
        {
            try
            {
                var parametersToExclude = new HashSet<string> { "Driver", "Remarks", "TruckID" };

                var dynamicParams = new DynamicParameters();

                foreach (var item in parameters)
                {
                    if (!parametersToExclude.Contains(item.Key))
                    {
                        dynamicParams.Add(item.Key, item.Value);
                    }
                }
                using (var connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();

                    using (var multi = await connection.QueryMultipleAsync(StoredProcedure.GetDataForR2470outReport.ToString(), dynamicParams, commandType: CommandType.StoredProcedure))
                    {
                        var response = new R2470outReportResponse
                        {
                            Consignment = await multi.ReadFirstOrDefaultAsync<Consignment>(),
                            ConsignmentRelease = await multi.ReadFirstOrDefaultAsync<ConsignmentRelease>(),
                            EntryLinesMovesViewCalcList = (await multi.ReadAsync<EntryLinesMovesViewCalc>()).ToList(),
                            Balance = await multi.ReadFirstOrDefaultAsync<int?>(),
                            MoveDate = await multi.ReadFirstOrDefaultAsync<string>(),
                            User = await multi.ReadFirstOrDefaultAsync<string>(),
                            Remarks = await multi.ReadFirstOrDefaultAsync<string>(),                            
                            GoodList = await multi.ReadFirstOrDefaultAsync<GoodList>(),
                            VarSequence = await multi.ReadFirstOrDefaultAsync<int>(),
                            Manifest = manifest,
                            ReportDtl = reportDtl,
                            //VarSequence = Convert.ToInt32(parameters["InvMovRecID"]),
                            Driver = parameters["Driver"].ToString(),
                            //Remarks = parameters["Remarks"].ToString(),
                            TruckID = parameters["TruckID"].ToString(),

                        };
                        return response;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Data For R2470out Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.DBAccessFailure, $"{ ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.DBAccessFailure] } : {ex.Message}");
            }
        }

        public async Task<Mrtg24720ReportResponse> GetDataForMrtg24720Report(Dictionary<string, object> parameters, Manifest manifest, ReportDtl reportDtl)
        {
            try
            {
                var parametersToExclude = new HashSet<string> { "InvMovRecID" };

                var dynamicParams = new DynamicParameters();

                foreach (var item in parameters)
                {
                    if (!parametersToExclude.Contains(item.Key))
                    {
                        dynamicParams.Add(item.Key, item.Value);
                    }
                }
                using (var connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();

                    using (var multi = await connection.QueryMultipleAsync(StoredProcedure.GetDataForMrtg24720Report.ToString(), dynamicParams, commandType: CommandType.StoredProcedure))
                    {
                        var response = new Mrtg24720ReportResponse
                        {
                            Consignment = await multi.ReadFirstOrDefaultAsync<Consignment>(),
                            EntryLineMoveList = (await multi.ReadAsync<EntryLineMoveView>()).ToList(),
                            ItemsList = (await multi.ReadAsync<Item>()).ToList(),
                            Manifest = manifest,
                            ReportDtl = reportDtl,
                            VarSequence = Convert.ToInt32(parameters["InvMovRecID"])
                        };
                        return response;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Data For Mrtg24720 Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.DBAccessFailure, $"{ ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.DBAccessFailure] } : {ex.Message}");
            }
        }

        public async Task<R24720PReportResponse> GetDataForR24720PReport(Dictionary<string, object> parameters, Manifest manifest, ReportDtl reportDtl)
        {
            try
            {                
                using (var connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();

                    using (var multi = await connection.QueryMultipleAsync(StoredProcedure.GetDataForR24720PReport.ToString(), new DynamicParameters(parameters), commandType: CommandType.StoredProcedure))
                    {
                        var response = new R24720PReportResponse
                        {
                            Consignment = await multi.ReadFirstOrDefaultAsync<Consignment>(),
                            EntryLinesMovesViewCalcList = (await multi.ReadAsync<EntryLinesMovesViewCalc>()).ToList(),
                            MoveDate = await multi.ReadFirstOrDefaultAsync<string>(),
                            Manifest = manifest,
                            ReportDtl = reportDtl
                        };

                        return response;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Data For R24720P Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.DBAccessFailure, $"{ ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.DBAccessFailure] } : {ex.Message}");
            }
        }

        public async Task<R60ExOutReportResponse> GetDataForR60ExOutReport(Dictionary<string, object> parameters, Manifest manifest, ReportDtl reportDtl)
        {
            try
            {
                var parametersToExclude = new HashSet<string> { "DriverName", "DriverID", "LicensePlate" };

                var dynamicParams = new DynamicParameters();

                foreach (var item in parameters)
                {
                    if (!parametersToExclude.Contains(item.Key))
                    {
                        dynamicParams.Add(item.Key, item.Value);
                    }
                }
                using (var connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();

                    using (var multi = await connection.QueryMultipleAsync(StoredProcedure.GetDataForR60ExOutReport.ToString(), dynamicParams, commandType: CommandType.StoredProcedure))
                    {
                        var response = new R60ExOutReportResponse
                        {
                            Consignment = await multi.ReadFirstOrDefaultAsync<Consignment>(),
                            SpecialAct = await multi.ReadFirstOrDefaultAsync<SpecialAct>(),
                            EntryLineMoveList = (await multi.ReadAsync<EntryLineMoveView>()).ToList(),
                            Manifest = manifest,
                            ReportDtl = reportDtl,
                            VarSequence = Convert.ToInt32(parameters["RequestID"]),
                            DriverName = parameters["DriverName"].ToString(),
                            DriverID = parameters["DriverID"].ToString(),
                            LicensePlate = parameters["LicensePlate"].ToString()
                        };                        

                        if (response.EntryLineMoveList != null)
                        {
                            for (int i = 0; i < response.EntryLineMoveList.Count; i++)
                            {
                                response.EntryLineMoveList[i].SerialNumber = i + 1;
                            }
                        }

                        return response;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Data For R60ExOut Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.DBAccessFailure, $"{ ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.DBAccessFailure] } : {ex.Message}");
            }
        }

        public async Task<R60ExInReportResponse> GetDataForR60ExInReport(Dictionary<string, object> parameters, Manifest manifest, ReportDtl reportDtl)
        {
            try
            {                
                using (var connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();

                    using (var multi = await connection.QueryMultipleAsync(StoredProcedure.GetDataForR60ExInReport.ToString(), new DynamicParameters(parameters), commandType: CommandType.StoredProcedure))
                    {
                        var response = new R60ExInReportResponse
                        {
                            Consignment = await multi.ReadFirstOrDefaultAsync<Consignment>(),
                            SpecialAct = await multi.ReadFirstOrDefaultAsync<SpecialAct>(),
                            EntryLineMoveList = (await multi.ReadAsync<EntryLineMoveView>()).ToList(),
                            Manifest = manifest,
                            ReportDtl = reportDtl,
                            VarSequence = Convert.ToInt32(parameters["RequestID"])
                        };

                        if (response.EntryLineMoveList != null)
                        {
                            for (int i = 0; i < response.EntryLineMoveList.Count; i++)
                            {
                                response.EntryLineMoveList[i].SerialNumber = i + 1;
                            }
                        }

                        return response;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Data For R60ExIn Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.DBAccessFailure, $"{ ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.DBAccessFailure] } : {ex.Message}");
            }
        }

        public async Task<R60splitReportResponse> GetDataForR60splitReport(Dictionary<string, object> parameters, Manifest manifest, ReportDtl reportDtl)
        {
            try
            {
                using (var connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();

                    using (var multi = await connection.QueryMultipleAsync(StoredProcedure.GetDataForR60splitReport.ToString(), new DynamicParameters(parameters), commandType: CommandType.StoredProcedure))
                    {
                        var response = new R60splitReportResponse
                        {
                            Consignment = await multi.ReadFirstOrDefaultAsync<Consignment>(),
                            EntryLineList = (await multi.ReadAsync<EntryLine>()).ToList(),
                            Manifest = manifest,
                            ReportDtl = reportDtl,
                            VarSequence = Convert.ToInt32(parameters["RequestID"])
                        };

                        return response;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Data For R60split Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.DBAccessFailure, $"{ ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.DBAccessFailure] } : {ex.Message}");
            }
        }

        public async Task<R2470outCollectReportResponse> GetDataForR2470outCollectReport(Dictionary<string, object> parameters, Manifest manifest, ReportDtl reportDtl)
        {
            try
            {
                using (var connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();

                    using (var multi = await connection.QueryMultipleAsync(StoredProcedure.GetDataForR2470outCollectReport.ToString(), new DynamicParameters(parameters), commandType: CommandType.StoredProcedure))
                    {
                        var response = new R2470outCollectReportResponse
                        {
                            CollectReleaseMaster = await multi.ReadFirstOrDefaultAsync<CollectReleaseMaster>(),
                            CustomersList = await multi.ReadFirstOrDefaultAsync<CustomersList>(),
                            EntryLineMoveList = (await multi.ReadAsync<EntryLineMoveView>()).ToList(),
                            MoveDate = await multi.ReadFirstOrDefaultAsync<string>(),
                            Manifest = manifest,
                            ReportDtl = reportDtl,
                            UnitedMovRef = Convert.ToInt32(parameters["UnitedMovRef"])
                        };

                        if (response.EntryLineMoveList != null)
                        {
                            for (int i = 0; i < response.EntryLineMoveList.Count; i++)
                            {
                                response.EntryLineMoveList[i].SerialNumber = i + 1;
                            }
                        }

                        return response;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Data For R2470outCollect Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.DBAccessFailure, $"{ ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.DBAccessFailure] } : {ex.Message}");
            }
        }


        public async Task<RCollectPlanningReportResponse> GetDataForRCollectPlanningReport(Dictionary<string, object> parameters, Manifest manifest, ReportDtl reportDtl)
        {
            try
            {
                using (var connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();

                    using (var multi = await connection.QueryMultipleAsync(StoredProcedure.GetDataForRCollectPlanningReport.ToString(), new DynamicParameters(parameters), commandType: CommandType.StoredProcedure))
                    {
                        var response = new RCollectPlanningReportResponse
                        {
                            CollectReleaseMaster = await multi.ReadFirstOrDefaultAsync<CollectReleaseMaster>(),
                            CustomersList = await multi.ReadFirstOrDefaultAsync<CustomersList>(),
                            EntryLineMoveList = (await multi.ReadAsync<EntryLineMoveView>()).ToList(),
                            Manifest = manifest,
                            ReportDtl = reportDtl,
                            UnitedMovRef = Convert.ToInt32(parameters["UnitedMovRef"])
                        };

                        if (response.EntryLineMoveList != null)
                        {
                            var result = new List<EntryLineMoveView>();

                            var grouped = response.EntryLineMoveList
                                .GroupBy(x => x.CatalogID)
                                .OrderBy(g => g.Key);

                            int serial = 1;

                            foreach (var group in grouped)
                            {
                                var list = group.ToList();

                                
                                foreach (var item in list)
                                {
                                    item.SerialNumber = serial++;
                                    item.MissingQuantity = null;
                                    if(item.ReleaseQuantity < 0)
                                    {
                                        item.ReleaseQuantity = 0;
                                    }

                                    result.Add(item);
                                }

                                // Summary row
                                var sumRelease = list
                                    .Where(x => x.ReleaseQuantity > 0)
                                    .Sum(x => x.ReleaseQuantity);
                                var quantity = list.First().Quantity;

                                var summaryRow = new EntryLineMoveView
                                {
                                    SerialNumber = null,
                                    Quantity = null,
                                    CatalogID = "",

                                    MissingQuantity = ((sumRelease - quantity) >= 0 ? (sumRelease - quantity): 0), 
                                    ReleaseQuantity = (sumRelease >= 0 ? sumRelease : 0),
                                    Comment = (sumRelease < quantity) ? "*" : "",


                                    FormattedGush = "",
                                    SumLineQuantityMove = null,
                                    DeclarationID = null
                                };

                                result.Add(summaryRow);
                            }

                            response.EntryLineMoveList = result;
                        }

                        return response;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Data For RCollectPlanningRepor Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.DBAccessFailure, $"{ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.DBAccessFailure]} : {ex.Message}");
            }
        }

        public async Task<InvRepForCustomsReportResponse> GetDataForInvRepForCustomsReport(Dictionary<string, object> parameters, Manifest manifest, ReportDtl reportDtl)
        {
            try
            {
                using (var connection = new SqlConnection(connectionString))
                {
                    await connection.OpenAsync();

                    using (var multi = await connection.QueryMultipleAsync(StoredProcedure.GetDataForInvRepForCustomsReport.ToString(), new DynamicParameters(parameters), commandType: CommandType.StoredProcedure))
                    {
                        var response = new InvRepForCustomsReportResponse
                        {
                            Consignment = await multi.ReadFirstOrDefaultAsync<Consignment>(),
                            Item = await multi.ReadFirstOrDefaultAsync<Item>(),
                            InventoryMoveEntry = await multi.ReadFirstOrDefaultAsync<InventoryMove>(),
                            InventoryMoveAdd = await multi.ReadFirstOrDefaultAsync<InventoryMove>(),
                            ConsignmentReleaseList = (await multi.ReadAsync<ConsignmentRelease>()).ToList(),
                            Manifest = manifest,
                            ReportDtl = reportDtl,
                        };
                        

                        int remainingAfterDelivery = 0;
                        var qtyEntry = response.InventoryMoveEntry.QuantityOnMove ?? 0;
                        var qtyAdd = response.InventoryMoveAdd.QuantityOnMove ?? 0;

                        if (response.ConsignmentReleaseList != null)
                        {
                            for (int i = 0; i < response.ConsignmentReleaseList.Count; i++)
                            {
                                remainingAfterDelivery += response.ConsignmentReleaseList[i].LineQuantityMove ?? 0 ;
                                response.ConsignmentReleaseList[i].RemainingAfterDelivery = qtyEntry + qtyAdd - remainingAfterDelivery;
                            }
                        }

                        return response;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Data For InvRepForCustoms Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.DBAccessFailure, $"{ ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.DBAccessFailure] } : {ex.Message}");
            }
        }

    }
}
