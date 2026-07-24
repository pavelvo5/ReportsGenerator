/****************************************************************************************
 FILE VERSION: 1 (2026-07-22)

 Changelog:
   v1 (2026-07-22) - Added an InvBck case to GetFilteredParameters, injecting IsPrint
                      and OutputFormat from the request into the SQL parameters passed
                      to GetDataForInvBckReport (scoped only to that one report - no
                      other report's stored procedure call is affected). Required for
                      GetDataForInvBckReport.sql's compact/portrait layout to trigger.
****************************************************************************************/
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
    public class ReportRepositoryAdoNet : IReportRepositoryAdoNet
    {
        private readonly string connectionString;
        private readonly ILogger logger;


        public ReportRepositoryAdoNet(string connectionString, ILogger logger)
        {
            this.connectionString = connectionString;
            this.logger = logger;
        }


        public Dictionary<string, object> GetFilteredParameters(ReportRequest request, ReportDtl reportDtl)
        {
            try
            {
                HashSet<string> parametersToRemove = new HashSet<string>();
                Dictionary<string, object> parametersToAdd = new Dictionary<string, object>();

                if (Enum.TryParse(reportDtl.FunctionName, true, out StoredProcedure storedProcedure))
                {
                    switch (storedProcedure)
                    {
                        case StoredProcedure.GetDataForStorageCalcReport:
                            parametersToRemove.Add("VAT");
                            break;

                        case StoredProcedure.GetDataForInvBckReport:
                            // The procedure decides its compact/portrait print layout itself
                            // (column set, truncation, header text) based on these two -
                            // IsPrint is the active trigger today; OutputFormat is passed
                            // through too so the trigger can move to it later without any
                            // C# change (see notes in the stored procedure).
                            parametersToAdd["IsPrint"] = request.IsPrint;
                            parametersToAdd["OutputFormat"] = request.OutputFormat;
                            break;

                        default:
                            break;
                    }
                }

                return request.Parameters
                    .Where(p => !parametersToRemove.Contains(p.Key) && !parametersToAdd.ContainsKey(p.Key))
                    .Concat(parametersToAdd)
                    .ToDictionary(p => p.Key, p => p.Value);
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Filtered Parameters For {reportDtl.ReportID} Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.GlobalError, ex.Message);
            }
        }
        public DataSet GetData(ReportRequest request, ReportDtl reportDtl)
        {
            try
            {
                Dictionary<string, object> parameters = GetFilteredParameters(request, reportDtl);

                using (SqlConnection conn = new SqlConnection(connectionString))
                using (SqlCommand cmd = new SqlCommand(reportDtl.FunctionName, conn))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    if (parameters != null)
                    {
                        foreach (var param in parameters)
                        {
                            cmd.Parameters.AddWithValue(param.Key, param.Value ?? DBNull.Value);
                        }
                    }

                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        DataSet ds = new DataSet();
                        adapter.Fill(ds);
                        return ds;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.WriteLog($"Error to Get Data For {reportDtl.ReportID} Report: {ex}");
                throw new CustomException((int)ErrorMessages.ErrorCodes.DBAccessFailure, $"{ErrorMessages.Messages[(int)ErrorMessages.ErrorCodes.DBAccessFailure]} : {ex.Message}");
            }
        }
    }
}
